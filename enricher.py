import re

ICON_MAP = {
    "\uf095": "📞 ",
    "\uf10b": "📞 ",
    "\uf0e0": "✉ ",
    "\uf003": "✉ ",
    "\uf041": "📍 ",
    "\uf3c5": "📍 ",
    "\uf08c": "🔗 LinkedIn: ",
    "\uf0e1": "🔗 LinkedIn: ",
    "\uf09b": "💻 GitHub: ",
    "\uf113": "💻 GitHub: ",
    "\uf0ac": "🌐 ",
    "\uf57d": "🌐 ",
}


def enrich_markdown(md_path, styles):
    with open(md_path, "r", encoding="utf-8") as f:
        content = f.read()

    # 1. Substitui ícones mapeados
    for icon_char, replacement in ICON_MAP.items():
        content = content.replace(icon_char, replacement)

    # 2. Remove TODOS os caracteres fora do range Latin + português
    content = remove_non_latin(content)

    # 3. Remove linhas que ficaram vazias ou só com símbolos
    content = clean_empty_lines(content)

    with open(md_path, "w", encoding="utf-8") as f:
        f.write(content)

    return content


def remove_non_latin(content):
    # Mantém apenas:
    # - ASCII básico (letras, números, pontuação)
    # - Latin Extended (acentos do português: á, é, ã, ç etc)
    # - Espaços, quebras de linha, tabs
    # - Símbolos Markdown: # * _ [ ] ( ) - > `
    # - Emojis mapeados acima (range \U0001F000+)
    result = []
    for char in content:
        cp = ord(char)
        if (
            cp <= 0x024F          # ASCII + Latin Extended (cobre português)
            or 0x0300 <= cp <= 0x036F   # Combining diacritics
            or char in '\n\r\t'
            or 0x1F000 <= cp <= 0x1FFFF # Emojis
            or 0x2600 <= cp <= 0x27BF   # Símbolos comuns (✉, 📍 etc)
        ):
            result.append(char)
        # tudo fora disso (coreano, japonês, ícones não mapeados) é descartado

    return "".join(result)


def clean_empty_lines(content):
    # Remove linhas com só 1-2 caracteres (resíduos de ícones)
    content = re.sub(r'(?m)^.{0,2}$\n', '', content)
    # Remove múltiplas linhas em branco consecutivas
    content = re.sub(r'\n{3,}', '\n\n', content)
    return content.strip()