"""
Parser do briefing (.docx) do Jornal Mural — lê parágrafo a parágrafo via
python-docx (mais confiável que markdown/pandoc pra esse tipo de estrutura).

Regra do cliente: a primeira matéria do briefing é sempre a CAPA.
"""
import re
from docx import Document


def _is_bold_paragraph(p):
    runs_with_text = [r for r in p.runs if r.text.strip()]
    return bool(runs_with_text) and all(r.bold for r in runs_with_text)


def parse_briefing(docx_path):
    doc = Document(docx_path)
    paragraphs = [p.text.strip() for p in doc.paragraphs]
    bold_flags = [_is_bold_paragraph(p) for p in doc.paragraphs]

    paginas = []
    cartazes = []
    pagina_atual = None
    materia_atual = None
    estado = None  # 'aguardando_titulo' | 'lendo_corpo'
    em_cartazes = False

    for text, is_bold in zip(paragraphs, bold_flags):
        if not text:
            continue

        if text.lower().startswith("cartazes"):
            em_cartazes = True
            continue

        # formato de cartaz: "Nome do cartaz: ARQUIVO.jpg" — detectado pelo
        # padrão em si, já que nem todo briefing tem um cabeçalho "Cartazes:"
        m_cartaz = re.match(r"[^:]{2,40}:\s*(.+\.(?:jpg|jpeg|png))\s*$", text, re.IGNORECASE)
        if m_cartaz and not text.lower().startswith(("editoria", "foto")):
            cartazes.append(m_cartaz.group(1).strip())
            continue

        if em_cartazes:
            continue

        m_pagina = re.match(r"p[aá]gina\s+(única|dupla)", text, re.IGNORECASE)
        if m_pagina:
            if pagina_atual and pagina_atual["materias"]:
                paginas.append(pagina_atual)
            tipo = "unica" if "nica" in m_pagina.group(1).lower() else "dupla"
            pagina_atual = {"tipo": tipo, "materias": []}
            materia_atual = None
            estado = None
            continue

        if re.match(r"mat[eé]ria\s+\d+", text, re.IGNORECASE):
            resto = re.sub(r"^mat[eé]ria\s+\d+\s*", "", text, flags=re.IGNORECASE).strip()
            if not resto:
                continue
            text = resto  # a linha continha "Matéria N" + conteúdo real juntos; processa o resto normalmente

        m_editoria = re.match(r"editoria:\s*(.+)", text, re.IGNORECASE)
        if m_editoria:
            materia_atual = {
                "editoria": m_editoria.group(1).strip(),
                "titulo": None,
                "corpo": "",
                "foto_arquivo": None,
            }
            if pagina_atual is not None:
                pagina_atual["materias"].append(materia_atual)
            estado = "aguardando_titulo"
            continue

        m_foto = re.match(r"foto:\s*(.+)", text, re.IGNORECASE)
        if m_foto and materia_atual is not None:
            conteudo = m_foto.group(1).strip()
            m_unir = re.match(r"unir\s+essa:\s*(.+)", conteudo, re.IGNORECASE)
            if m_unir:
                partes = re.split(r"\s+e\s+essa:\s*", m_unir.group(1), flags=re.IGNORECASE)
                materia_atual["foto_arquivo"] = [p.strip() for p in partes if p.strip()]
            else:
                materia_atual["foto_arquivo"] = conteudo
            estado = None
            continue

        if text.startswith("=") or text.startswith("/"):
            continue

        if materia_atual is None:
            continue

        if estado == "aguardando_titulo" and is_bold:
            materia_atual["titulo"] = text
            estado = "lendo_corpo"
        elif estado == "lendo_corpo":
            materia_atual["corpo"] = (materia_atual["corpo"] + " " + text).strip()

    if pagina_atual and pagina_atual["materias"]:
        paginas.append(pagina_atual)

    if paginas and paginas[0]["materias"]:
        paginas[0]["materias"][0]["is_capa"] = True

    return paginas, cartazes


if __name__ == "__main__":
    import json
    paginas, cartazes = parse_briefing("/mnt/user-data/uploads/JORNAL_MURAL_BRIEFING_-_07_07_26.docx")
    print(json.dumps(paginas, indent=2, ensure_ascii=False))
    print("CARTAZES:", cartazes)
