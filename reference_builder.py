import os
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH


def build_reference_docx(doc_info, output_dir):
    doc = Document()

    main_font = doc_info.get("main_font", "Calibri")
    bold_font = doc_info.get("bold_font", main_font)
    body_size = doc_info.get("body_size", 11)
    title_sizes = doc_info.get("title_sizes", [16, 14, 12])
    color_hex = doc_info.get("main_color_hex", "#000000")

    # Cor principal do texto
    r, g, b = hex_to_rgb(color_hex)

    # --- Estilo Normal (corpo) ---
    normal = doc.styles["Normal"]
    normal.font.name = main_font
    normal.font.size = Pt(body_size)
    normal.font.color.rgb = RGBColor(r, g, b)

    # --- Heading 1 ---
    h1_size = title_sizes[0] if len(title_sizes) > 0 else body_size + 6
    h1 = doc.styles["Heading 1"]
    h1.font.name = bold_font
    h1.font.size = Pt(h1_size)
    h1.font.bold = True
    h1.font.color.rgb = RGBColor(r, g, b)
    h1.font.underline = False
    h1.paragraph_format.space_before = Pt(12)
    h1.paragraph_format.space_after = Pt(6)

    # --- Heading 2 ---
    h2_size = title_sizes[1] if len(title_sizes) > 1 else body_size + 4
    h2 = doc.styles["Heading 2"]
    h2.font.name = bold_font
    h2.font.size = Pt(h2_size)
    h2.font.bold = True
    h2.font.color.rgb = RGBColor(r, g, b)
    h2.font.underline = False
    h2.paragraph_format.space_before = Pt(8)
    h2.paragraph_format.space_after = Pt(4)

    # --- Heading 3 ---
    h3_size = title_sizes[2] if len(title_sizes) > 2 else body_size + 2
    h3 = doc.styles["Heading 3"]
    h3.font.name = bold_font
    h3.font.size = Pt(h3_size)
    h3.font.bold = True
    h3.font.color.rgb = RGBColor(r, g, b)
    h3.font.underline = False

    # Salva o reference.docx
    ref_path = os.path.join(output_dir, "reference_temp.docx")
    doc.save(ref_path)
    return ref_path


def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip("#")
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))