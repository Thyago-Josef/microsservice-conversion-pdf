import os
import uuid
from flask import Flask, request, jsonify, send_file
from docling.document_converter import DocumentConverter

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "output"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

converter = DocumentConverter()

@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "UP"})

@app.route("/convert/markdown", methods=["POST"])
def convert_to_markdown():
    if "file" not in request.files:
        return jsonify({"error": "Nenhum arquivo enviado"}), 400

    file = request.files["file"]
    file_id = str(uuid.uuid4())
    input_path = os.path.join(UPLOAD_FOLDER, f"{file_id}_{file.filename}")
    file.save(input_path)

    try:
        # Docling converte PDF -> Markdown
        result = converter.convert(input_path)
        markdown_content = result.document.export_to_markdown()

        # Salva o Markdown
        md_filename = f"{file_id}.md"
        md_path = os.path.join(OUTPUT_FOLDER, md_filename)
        with open(md_path, "w", encoding="utf-8") as f:
            f.write(markdown_content)

        return jsonify({
            "status": "success",
            "file_id": file_id,
            "download_url": f"/download/{md_filename}"
        })

    except Exception as e:
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