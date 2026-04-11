# import fitz
# from collections import Counter


# def analyze_pdf(pdf_path):
#     doc = fitz.open(pdf_path)
#     styles = []
#     icons = []
#     colors = []

#     for page_num, page in enumerate(doc):
#         # Extrai texto com estilo completo
#         blocks = page.get_text("dict")["blocks"]
#         for block in blocks:
#             if block["type"] == 0:  # bloco de texto
#                 for line in block["lines"]:
#                     for span in line["spans"]:
#                         text = span["text"].strip()
#                         if not text:
#                             continue

#                         styles.append({
#                             "text": text,
#                             "font": span["font"],
#                             "size": round(span["size"], 1),
#                             "bold": is_bold(span["font"], span["flags"]),
#                             "italic": is_italic(span["font"], span["flags"]),
#                             "color": span["color"],
#                             "x": round(span["origin"][0], 1),
#                             "y": round(span["origin"][1], 1),
#                             "page": page_num
#                         })

#                         colors.append(span["color"])

#         # Extrai imagens (ícones e logos)
#         for img in page.get_images(full=True):
#             bbox = page.get_image_bbox(img)
#             if bbox:
#                 icons.append({
#                     "page": page_num,
#                     "x": round(bbox.x0, 1),
#                     "y": round(bbox.y0, 1),
#                     "width": round(bbox.width, 1),
#                     "height": round(bbox.height, 1),
#                     "is_icon": bbox.width < 30 and bbox.height < 30
#                 })

#     doc.close()
#     return styles, icons


# def is_bold(font_name, flags):
#     font_lower = font_name.lower()
#     return "bold" in font_lower or "black" in font_lower or (flags & 2**4 != 0)


# def is_italic(font_name, flags):
#     font_lower = font_name.lower()
#     return "italic" in font_lower or "oblique" in font_lower or (flags & 2**1 != 0)


# def extract_doc_info(styles):
#     if not styles:
#         return {}

#     # Fonte mais usada = fonte principal
#     fonts = Counter(s["font"] for s in styles)
#     main_font = fonts.most_common(1)[0][0]

#     # Tamanho mais comum = tamanho do corpo
#     sizes = Counter(round(s["size"]) for s in styles)
#     body_size = sizes.most_common(1)[0][0]

#     # Tamanhos maiores = títulos (ordenados do maior para menor)
#     title_sizes = sorted(
#         set(round(s["size"]) for s in styles if round(s["size"]) > body_size),
#         reverse=True
#     )

#     # Cor principal do texto
#     text_colors = Counter(s["color"] for s in styles)
#     main_color = text_colors.most_common(1)[0][0]

#     # Converte cor de int para hex
#     main_color_hex = int_to_hex(main_color)

#     # Detecta se tem fonte bold separada
#     bold_fonts = [s["font"] for s in styles if s["bold"]]
#     bold_font = Counter(bold_fonts).most_common(1)[0][0] if bold_fonts else main_font

#     # Limpa nome da fonte para uso no Pandoc
#     clean_font = clean_font_name(main_font)
#     clean_bold_font = clean_font_name(bold_font)

#     return {
#         "main_font": clean_font,
#         "bold_font": clean_bold_font,
#         "body_size": body_size,
#         "title_sizes": title_sizes[:3],  # máximo 3 níveis de título
#         "main_color_hex": main_color_hex,
#         "has_icons": False  # atualizado pelo enricher
#     }


# def int_to_hex(color_int):
#     try:
#         r = (color_int >> 16) & 0xFF
#         g = (color_int >> 8) & 0xFF
#         b = color_int & 0xFF
#         return f"#{r:02X}{g:02X}{b:02X}"
#     except Exception:
#         return "#000000"


# def clean_font_name(font_name):
#     # Remove sufixos técnicos como ",Bold" ou "-BoldItalic"
#     for suffix in [",Bold", ",Italic", ",BoldItalic", "-Bold",
#                    "-Italic", "-BoldItalic", "-Regular", ",Regular"]:
#         font_name = font_name.replace(suffix, "")
#     return font_name.strip()


import fitz
from collections import Counter
from langdetect import detect


def analyze_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    styles = []
    icons = []

    for page_num, page in enumerate(doc):
        blocks = page.get_text("dict")["blocks"]
        for block in blocks:
            if block["type"] == 0:
                for line in block["lines"]:
                    for span in line["spans"]:
                        text = span["text"].strip()
                        if not text:
                            continue

                        styles.append({
                            "text": text,
                            "font": span["font"],
                            "size": round(span["size"], 1),
                            "bold": is_bold(span["font"], span["flags"]),
                            "italic": is_italic(span["font"], span["flags"]),
                            "color": span["color"],
                            "x": round(span["origin"][0], 1),
                            "y": round(span["origin"][1], 1),
                            "page": page_num
                        })

        for img in page.get_images(full=True):
            bbox = page.get_image_bbox(img)
            if bbox:
                icons.append({
                    "page": page_num,
                    "x": round(bbox.x0, 1),
                    "y": round(bbox.y0, 1),
                    "width": round(bbox.width, 1),
                    "height": round(bbox.height, 1),
                    "is_icon": bbox.width < 30 and bbox.height < 30
                })

    doc.close()
    return styles, icons


def is_bold(font_name, flags):
    font_lower = font_name.lower()
    return "bold" in font_lower or "black" in font_lower or (flags & 2**4 != 0)


def is_italic(font_name, flags):
    font_lower = font_name.lower()
    return "italic" in font_lower or "oblique" in font_lower or (flags & 2**1 != 0)


def detect_language(styles):
    # Junta todo o texto extraído
    full_text = " ".join(s["text"] for s in styles if s["text"].strip())

    try:
        lang_code = detect(full_text)
        # Mapeia para formato BCP-47 que o Pandoc entende
        lang_map = {
            "pt": "pt-BR",
            "en": "en-US",
            "es": "es",
            "fr": "fr",
            "de": "de",
            "it": "it",
            "nl": "nl",
            "ru": "ru",
            "zh-cn": "zh-CN",
            "ja": "ja",
            "ko": "ko"
        }
        detected = lang_map.get(lang_code, "pt-BR")
        print(f"🌐 Idioma detectado: {detected} (código bruto: {lang_code})")
        return detected
    except Exception as e:
        print(f"⚠️ Não foi possível detectar idioma: {e}. Usando pt-BR como padrão.")
        return "pt-BR"


def extract_doc_info(styles):
    if not styles:
        return {}

    # Fonte mais usada = fonte principal
    fonts = Counter(s["font"] for s in styles)
    main_font = fonts.most_common(1)[0][0]

    # Tamanho mais comum = tamanho do corpo
    sizes = Counter(round(s["size"]) for s in styles)
    body_size = sizes.most_common(1)[0][0]

    # Tamanhos maiores = títulos
    title_sizes = sorted(
        set(round(s["size"]) for s in styles if round(s["size"]) > body_size),
        reverse=True
    )

    # Cor principal do texto
    text_colors = Counter(s["color"] for s in styles)
    main_color = text_colors.most_common(1)[0][0]
    main_color_hex = int_to_hex(main_color)

    # Fonte bold
    bold_fonts = [s["font"] for s in styles if s["bold"]]
    bold_font = Counter(bold_fonts).most_common(1)[0][0] if bold_fonts else main_font

    # Limpa nomes das fontes
    clean_font = clean_font_name(main_font)
    clean_bold_font = clean_font_name(bold_font)

    # Detecta idioma
    language = detect_language(styles)

    return {
        "main_font": clean_font,
        "bold_font": clean_bold_font,
        "body_size": body_size,
        "title_sizes": title_sizes[:3],
        "main_color_hex": main_color_hex,
        "has_icons": False,
        "language": language
    }


def int_to_hex(color_int):
    try:
        r = (color_int >> 16) & 0xFF
        g = (color_int >> 8) & 0xFF
        b = color_int & 0xFF
        return f"#{r:02X}{g:02X}{b:02X}"
    except Exception:
        return "#000000"


def clean_font_name(font_name):
    for suffix in [",Bold", ",Italic", ",BoldItalic", "-Bold",
                   "-Italic", "-BoldItalic", "-Regular", ",Regular"]:
        font_name = font_name.replace(suffix, "")
    return font_name.strip()