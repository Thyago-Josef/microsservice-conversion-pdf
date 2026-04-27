import os
import uuid
import subprocess
import re
import fitz
from flask import Flask, request, jsonify, send_file
from docling.document_converter import DocumentConverter
from pdf2docx import Converter
from docx import Document

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "output"
PANDOC_PATH = r"C:\pandoc-3.9.0.2\pandoc.exe"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

converter = DocumentConverter()


def clean_oriental_chars(docx_path):
    """Remove apenas caracteres chineses/japoneses/coreanos específicos"""
    try:
        doc = Document(docx_path)
        for para in doc.paragraphs:
            text = para.text
            if not text:
                continue
            # Remove apenas caracteres CJK específicos
            cleaned = []
            for c in text:
                cp = ord(c)
                # Mantém: latin, pontuação, números, acentos
                # Remove: CJK Unificado (4E00-9FFF), Hiragana (3040-30FF), Katakana, Hangul
                if not (0x4E00 <= cp <= 0x9FFF or 0x3040 <= cp <= 0x30FF or 0xAC00 <= cp <= 0xD7AF):
                    cleaned.append(c)
            new_text = ''.join(cleaned)
            if new_text != text:
                para.text = new_text
        doc.save(docx_path)
    except Exception as e:
        print(f"⚠️ Erro ao limpar caracteres: {e}")


def convert_pdf_pdf2docx(pdf_path, docx_path):
    """Converte PDF para DOCX mantendo formatação"""
    print(f"🎯 Convertendo via pdf2docx: {pdf_path}")
    cv = Converter(pdf_path)
    cv.convert(docx_path, start=0, end=None)
    cv.close()

    print("🧹 Limpando caracteres orientais...")
    clean_oriental_chars(docx_path)

    print(f"✅ DOCX gerado: {docx_path}")


def clean_markdown(md_path):
    """Limpa caracteres orientais do Markdown"""
    with open(md_path, "r", encoding="utf-8") as f:
        content = f.read()

    cleaned = []
    for char in content:
        cp = ord(char)
        if (
            cp <= 0x024F
            or 0x0300 <= cp <= 0x036F
            or char in '\n\r\t'
            or 0x1F000 <= cp <= 0x1FFFF
            or 0x2600 <= cp <= 0x27BF
        ):
            cleaned.append(char)

    content = "".join(cleaned)

    lines = content.split('\n')
    filtered_lines = []
    for line in lines:
        stripped = line.strip()
        has_real_text = any(c.isalpha() for c in stripped)

        if stripped == '':
            filtered_lines.append('')
        elif stripped.startswith('#'):
            filtered_lines.append(line)
        elif stripped == '---':
            filtered_lines.append(line)
        elif has_real_text:
            filtered_lines.append(line)

    content = '\n'.join(filtered_lines)
    content = re.sub(r'\n{3,}', '\n\n', content)

    with open(md_path, "w", encoding="utf-8") as f:
        f.write(content.strip())

    print("✅ Markdown limpo")


# ==================== ROTAS ====================

@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "UP", "service": "docling-converter"})


@app.route("/convert/docx", methods=["POST"])
def convert_to_docx():
    if "file" not in request.files:
        return jsonify({"error": "Nenhum arquivo enviado"}), 400

    strategy = request.form.get("strategy", "pdf2docx")

    file = request.files["file"]
    file_id = str(uuid.uuid4())
    input_path = os.path.join(UPLOAD_FOLDER, f"{file_id}_{file.filename}")
    file.save(input_path)

    try:
        docx_filename = f"{file_id}.docx"
        docx_path = os.path.join(OUTPUT_FOLDER, docx_filename)

        if strategy == "pdf2docx":
            convert_pdf_pdf2docx(input_path, docx_path)

        elif strategy == "docling":
            from analyzer import analyze_pdf, extract_doc_info
            from reference_builder import build_reference_docx

            print(f"🎯 Convertendo via Docling + Pandoc: {input_path}")

            styles, icons = analyze_pdf(input_path)
            doc_info = extract_doc_info(styles)

            result = converter.convert(input_path)
            markdown_content = result.document.export_to_markdown()

            md_filename = f"{file_id}.md"
            md_path = os.path.join(OUTPUT_FOLDER, md_filename)
            with open(md_path, "w", encoding="utf-8") as f:
                f.write(markdown_content)

            clean_markdown(md_path)

            ref_path = build_reference_docx(doc_info, OUTPUT_FOLDER)
            lang = doc_info.get("language", "pt-BR")

            result_pandoc = subprocess.run([
                PANDOC_PATH,
                md_path,
                "-o", docx_path,
                "--from", "markdown",
                "--to", "docx",
                "--standalone",
                f"--reference-doc={ref_path}",
                "--metadata", f"lang={lang}"
            ], capture_output=True, text=True)

            if result_pandoc.returncode != 0:
                raise RuntimeError(f"Pandoc falhou: {result_pandoc.stderr}")

            os.remove(md_path)
            os.remove(ref_path)

            print(f"✅ DOCX gerado: {docx_path}")

        else:
            return jsonify({"error": "Estratégia inválida. Use: pdf2docx ou docling"}), 400

        return jsonify({
            "status": "success",
            "file_id": file_id,
            "strategy": strategy,
            "download_url": f"/download/{docx_filename}"
        })

    except Exception as e:
        print(f"❌ Erro: {str(e)}")
        return jsonify({"error": str(e)}), 500

    finally:
        if os.path.exists(input_path):
            os.remove(input_path)


@app.route("/convert/markdown", methods=["POST"])
def convert_to_markdown():
    if "file" not in request.files:
        return jsonify({"error": "Nenhum arquivo enviado"}), 400

    from analyzer import analyze_pdf, extract_doc_info

    file = request.files["file"]
    file_id = str(uuid.uuid4())
    input_path = os.path.join(UPLOAD_FOLDER, f"{file_id}_{file.filename}")
    file.save(input_path)

    try:
        styles, icons = analyze_pdf(input_path)
        doc_info = extract_doc_info(styles)

        result = converter.convert(input_path)
        markdown_content = result.document.export_to_markdown()

        md_filename = f"{file_id}.md"
        md_path = os.path.join(OUTPUT_FOLDER, md_filename)
        with open(md_path, "w", encoding="utf-8") as f:
            f.write(markdown_content)

        clean_markdown(md_path)

        return jsonify({
            "status": "success",
            "file_id": file_id,
            "download_url": f"/download/{md_filename}",
            "doc_info": doc_info
        })

    except Exception as e:
        print(f"❌ Erro: {str(e)}")
        return jsonify({"error": str(e)}), 500

    finally:
        if os.path.exists(input_path):
            os.remove(input_path)


@app.route("/download/<path:filename>", methods=["GET"])
def download(filename):
    path = os.path.join(OUTPUT_FOLDER, filename)
    if not os.path.exists(path):
        return jsonify({"error": "Arquivo não encontrado"}), 404
    return send_file(path, as_attachment=True)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)