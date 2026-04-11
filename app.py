# import os
# import re
# import uuid
# import subprocess
# from flask import Flask, request, jsonify, send_file
# from docling.document_converter import DocumentConverter
# from analyzer import analyze_pdf, extract_doc_info
# from reference_builder import build_reference_docx

# app = Flask(__name__)

# UPLOAD_FOLDER = "uploads"
# OUTPUT_FOLDER = "output"
# PANDOC_PATH = r"C:\pandoc-3.9.0.2\pandoc.exe"

# os.makedirs(UPLOAD_FOLDER, exist_ok=True)
# os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# converter = DocumentConverter()


# def clean_markdown(md_path):
#     with open(md_path, "r", encoding="utf-8") as f:
#         content = f.read()

#     # Limpa caractere por caractere preservando estrutura
#     cleaned = []
#     for char in content:
#         cp = ord(char)
#         if (
#             cp <= 0x024F
#             or 0x0300 <= cp <= 0x036F
#             or char in '\n\r\t'
#             or 0x1F000 <= cp <= 0x1FFFF
#             or 0x2600 <= cp <= 0x27BF
#         ):
#             cleaned.append(char)

#     content = "".join(cleaned)

#     # Remove linhas que ficaram com 1-2 caracteres (resíduos de ícones)
#     # MAS preserva linhas que começam com # (títulos Markdown)
#     lines = content.split('\n')
#     filtered_lines = []
#     for line in lines:
#         stripped = line.strip()
#         # Mantém a linha se: começa com #, tem mais de 3 chars, ou está vazia (espaçamento)
#         if stripped.startswith('#') or len(stripped) > 3 or stripped == '':
#             filtered_lines.append(line)

#     content = '\n'.join(filtered_lines)

#     # Remove mais de 2 linhas em branco consecutivas
#     content = re.sub(r'\n{3,}', '\n\n', content)

#     with open(md_path, "w", encoding="utf-8") as f:
#         f.write(content.strip())

#     print("✅ Markdown limpo")


# @app.route("/health", methods=["GET"])
# def health():
#     return jsonify({"status": "UP", "service": "docling-converter"})


# @app.route("/convert/markdown", methods=["POST"])
# def convert_to_markdown():
#     if "file" not in request.files:
#         return jsonify({"error": "Nenhum arquivo enviado"}), 400

#     file = request.files["file"]
#     file_id = str(uuid.uuid4())
#     input_path = os.path.join(UPLOAD_FOLDER, f"{file_id}_{file.filename}")
#     file.save(input_path)

#     try:
#         # PASSO 1: Analisa o PDF com pymupdf
#         print(f"🔍 Analisando PDF: {input_path}")
#         styles, icons = analyze_pdf(input_path)
#         doc_info = extract_doc_info(styles)
#         doc_info["has_icons"] = len(icons) > 0
#         print(f"✅ Análise: fonte={doc_info.get('main_font')}, "
#               f"tamanho={doc_info.get('body_size')}, "
#               f"ícones={len(icons)}")

#         # PASSO 2: Converte com Docling
#         print("📄 Convertendo com Docling...")
#         result = converter.convert(input_path)
#         markdown_content = result.document.export_to_markdown()

#         # PASSO 3: Salva o Markdown
#         md_filename = f"{file_id}.md"
#         md_path = os.path.join(OUTPUT_FOLDER, md_filename)
#         with open(md_path, "w", encoding="utf-8") as f:
#             f.write(markdown_content)

#         # PASSO 4: Limpa caracteres problemáticos
#         print("🧹 Limpando Markdown...")
#         clean_markdown(md_path)

#         # PASSO 5: Cria reference.docx dinâmico com as fontes do PDF
#         print(f"🎨 Criando referência com fonte: {doc_info.get('main_font')}")
#         ref_path = build_reference_docx(doc_info, OUTPUT_FOLDER)

#         # PASSO 6: Markdown -> DOCX via Pandoc
#         print("📝 Gerando DOCX com Pandoc...")
#         docx_filename = f"{file_id}.docx"
#         docx_path = os.path.join(OUTPUT_FOLDER, docx_filename)

#         # Pega o idioma detectado
#         lang = doc_info.get("language", "pt-BR")
#         print(f"🌐 Usando idioma: {lang}")


#         result_pandoc = subprocess.run([
#             PANDOC_PATH,
#             md_path,
#             "-o", docx_path,
#             "--from", "markdown",
#             "--to", "docx",
#             "--standalone",
#             f"--reference-doc={ref_path}",
#              "--metadata", f"lang={lang}"
#         ], capture_output=True, text=True)

#         if result_pandoc.returncode != 0:
#             raise RuntimeError(f"Pandoc falhou: {result_pandoc.stderr}")

#         # Limpa arquivos intermediários
#         os.remove(md_path)
#         os.remove(ref_path)

#         print(f"✅ DOCX gerado: {docx_path}")
#         return jsonify({
#             "status": "success",
#             "file_id": file_id,
#             "download_url": f"/download/{docx_filename}",
#             "doc_info": doc_info
#         })

#     except Exception as e:
#         print(f"❌ Erro: {str(e)}")
#         return jsonify({"error": str(e)}), 500

#     finally:
#         if os.path.exists(input_path):
#             os.remove(input_path)


# @app.route("/download/<filename>", methods=["GET"])
# def download(filename):
#     path = os.path.join(OUTPUT_FOLDER, filename)
#     if not os.path.exists(path):
#         return jsonify({"error": "Arquivo não encontrado"}), 404
#     return send_file(path, as_attachment=True)


# if __name__ == "__main__":
#     app.run(host="0.0.0.0", port=5000, debug=False)
import os
import re
import uuid
import subprocess
import fitz
import io
import pytesseract
from PIL import Image
from flask import Flask, request, jsonify, send_file
from docling.document_converter import DocumentConverter
from analyzer import analyze_pdf, extract_doc_info
from reference_builder import build_reference_docx

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "output"
PANDOC_PATH = r"C:\pandoc-3.9.0.2\pandoc.exe"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

converter = DocumentConverter()


# ==================== FUNÇÕES AUXILIARES ====================

def is_form_document(styles):
    short_spans = sum(1 for s in styles if len(s["text"].strip()) < 3)
    total_spans = len(styles)

    if total_spans == 0:
        return False

    ratio = short_spans / total_spans
    print(f"📊 Ratio de spans curtos: {ratio:.2f}")

    is_sparse = total_spans < 50
    print(f"📊 Total spans: {total_spans}, esparso: {is_sparse}")

    return ratio > 0.25 or is_sparse


def extract_text_pymupdf(pdf_path):
    doc = fitz.open(pdf_path)
    markdown_lines = []

    for page_num, page in enumerate(doc):
        if page_num > 0:
            markdown_lines.append("\n---\n")

        words = page.get_text("words")
        words.sort(key=lambda w: (round(w[1] / 8) * 8, w[0]))

        current_y = None
        current_line = []

        for word in words:
            x0, y0, x1, y1, text, *_ = word
            y_rounded = round(y0 / 8) * 8

            if current_y is None:
                current_y = y_rounded

            if abs(y_rounded - current_y) > 8:
                if current_line:
                    markdown_lines.append(" ".join(current_line))
                current_line = [text]
                current_y = y_rounded
            else:
                current_line.append(text)

        if current_line:
            markdown_lines.append(" ".join(current_line))

        markdown_lines.append("")

    doc.close()
    return '\n'.join(markdown_lines)


def extract_text_ocr(pdf_path):
    doc = fitz.open(pdf_path)
    markdown_lines = []

    for page_num, page in enumerate(doc):
        if page_num > 0:
            markdown_lines.append("\n---\n")

        # Renderiza página como imagem
        mat = fitz.Matrix(2.0, 2.0)
        pix = page.get_pixmap(matrix=mat)
        img_data = pix.tobytes("png")

        # OCR na imagem
        image = Image.open(io.BytesIO(img_data))
        text = pytesseract.image_to_string(image, lang="por")

        markdown_lines.append(text)

    doc.close()
    return '\n'.join(markdown_lines)


def clean_markdown(md_path):
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


def debug_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    page = doc[0]
    blocks = page.get_text("dict")["blocks"]

    print("=== DEBUG SPANS PRIMEIRA PÁGINA ===")
    for block in blocks:
        if block["type"] != 0:
            continue
        for line in block["lines"]:
            for span in line["spans"]:
                text = span["text"].strip()
                if text:
                    print(f"  Y={span['origin'][1]:.1f} X={span['origin'][0]:.1f} | '{text}'")
    print("===================================")
    doc.close()


# ==================== ROTAS ====================

@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "UP", "service": "docling-converter"})


@app.route("/convert/markdown", methods=["POST"])
def convert_to_markdown():
    if "file" not in request.files:
        return jsonify({"error": "Nenhum arquivo enviado"}), 400

    file = request.files["file"]
    file_id = str(uuid.uuid4())
    input_path = os.path.join(UPLOAD_FOLDER, f"{file_id}_{file.filename}")
    file.save(input_path)

    try:
        # PASSO 1: Analisa o PDF
        print(f"🔍 Analisando PDF: {input_path}")
        styles, icons = analyze_pdf(input_path)
        doc_info = extract_doc_info(styles)
        doc_info["has_icons"] = len(icons) > 0
        print(f"✅ Análise: fonte={doc_info.get('main_font')}, "
              f"tamanho={doc_info.get('body_size')}, "
              f"idioma={doc_info.get('language')}, "
              f"ícones={len(icons)}")

        # PASSO 2: Detecta tipo e escolhe estratégia
        debug_pdf(input_path)
        is_form = is_form_document(styles)
        has_real_text = len(styles) > 20

        print(f"📄 Tipo: {'Formulário' if is_form else 'Documento comum'}, "
              f"texto real: {has_real_text}")

        if not has_real_text:
            print("🔎 Usando OCR completo (PDF sem texto real)...")
            markdown_content = extract_text_ocr(input_path)
        elif is_form:
            print("📋 Usando pymupdf (formulário com campos)...")
            markdown_content = extract_text_pymupdf(input_path)
        else:
            print("📄 Usando Docling (documento comum)...")
            result = converter.convert(input_path)
            markdown_content = result.document.export_to_markdown()

        # PASSO 3: Salva o Markdown
        md_filename = f"{file_id}.md"
        md_path = os.path.join(OUTPUT_FOLDER, md_filename)
        with open(md_path, "w", encoding="utf-8") as f:
            f.write(markdown_content)

        # PASSO 4: Limpa caracteres problemáticos
        print("🧹 Limpando Markdown...")
        clean_markdown(md_path)

        # PASSO 5: Cria reference.docx dinâmico
        print(f"🎨 Criando referência com fonte: {doc_info.get('main_font')}")
        ref_path = build_reference_docx(doc_info, OUTPUT_FOLDER)

        # PASSO 6: Markdown -> DOCX via Pandoc
        lang = doc_info.get("language", "pt-BR")
        print(f"📝 Gerando DOCX com Pandoc... (idioma: {lang})")
        docx_filename = f"{file_id}.docx"
        docx_path = os.path.join(OUTPUT_FOLDER, docx_filename)

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
        return jsonify({
            "status": "success",
            "file_id": file_id,
            "download_url": f"/download/{docx_filename}",
            "doc_info": doc_info
        })

    except Exception as e:
        print(f"❌ Erro: {str(e)}")
        return jsonify({"error": str(e)}), 500

    finally:
        if os.path.exists(input_path):
            os.remove(input_path)


@app.route("/download/<filename>", methods=["GET"])
def download(filename):
    path = os.path.join(OUTPUT_FOLDER, filename)
    if not os.path.exists(path):
        return jsonify({"error": "Arquivo não encontrado"}), 404
    return send_file(path, as_attachment=True)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)