"""
Parser do briefing (.docx) do Jornal Mural — lê parágrafo a parágrafo via
python-docx (mais confiável que markdown/pandoc pra esse tipo de estrutura).

Regra do cliente: a primeira matéria do briefing é sempre a CAPA.

Aguenta dois jeitos de escrever o briefing:
  - Título e corpo em parágrafos separados (o título costuma vir em negrito).
  - Título e corpo GRUDADOS no mesmo parágrafo, separados por quebra de linha
    (e às vezes até com a "Editoria:" junto). Por isso o parser trabalha por
    LINHA lógica (quebra os parágrafos nos '\\n') e não exige que o título
    esteja em negrito.
"""
import re
from docx import Document


def _is_bold_paragraph(p):
    runs_with_text = [r for r in p.runs if r.text.strip()]
    return bool(runs_with_text) and all(r.bold for r in runs_with_text)


def _linhas_logicas(doc):
    """Achata os parágrafos em linhas lógicas, quebrando também nos '\\n'
    internos — assim título e corpo grudados no mesmo parágrafo viram linhas
    separadas, e o parser funciona igual nos dois formatos de briefing."""
    linhas = []
    for p in doc.paragraphs:
        bold = _is_bold_paragraph(p)
        for pedaco in p.text.split("\n"):
            t = pedaco.strip()
            if t:
                linhas.append((t, bold))
    return linhas


def _extrai_cartaz(text):
    """Pega o nome do arquivo de uma linha de cartaz tipo
    'Descrição: ARQUIVO.jpg' ou 'Descrição – ARQUIVO.jpg' — aceita ':', '–',
    '—' e ' - ' como separador entre a descrição e o arquivo. Retorna None se
    a linha não terminar num arquivo de imagem."""
    if not re.search(r"\.(?:jpg|jpeg|png)\s*$", text, re.IGNORECASE):
        return None
    partes = re.split(r"\s*[:–—]\s*|\s+-\s+", text, maxsplit=1)
    nome = partes[-1].strip() if len(partes) > 1 else text.strip()
    m = re.search(r"(.+\.(?:jpg|jpeg|png))\s*$", nome, re.IGNORECASE)
    return m.group(1).strip() if m else None


def parse_briefing(docx_path):
    doc = Document(docx_path)
    linhas = _linhas_logicas(doc)

    paginas = []
    cartazes = []
    pagina_atual = None
    materia_atual = None
    estado = None  # 'aguardando_titulo' | 'lendo_corpo'
    em_cartazes = False

    for text, is_bold in linhas:
        low = text.lower()

        if low.startswith("cartazes"):
            em_cartazes = True
            continue

        # depois do cabeçalho "Cartazes:", toda linha com arquivo é um cartaz
        if em_cartazes:
            nome = _extrai_cartaz(text)
            if nome:
                cartazes.append(nome)
            continue

        # cartaz solto (briefing sem cabeçalho "Cartazes:"): "Nome: ARQUIVO.jpg"
        if not low.startswith(("editoria", "foto")):
            m_cartaz = re.match(r"[^:]{2,40}:\s*(.+\.(?:jpg|jpeg|png))\s*$", text, re.IGNORECASE)
            if m_cartaz:
                cartazes.append(m_cartaz.group(1).strip())
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
            text = resto  # "Matéria N" + conteúdo real juntos: processa o resto
            low = text.lower()

        m_editoria = re.match(r"editoria:\s*(.+)", text, re.IGNORECASE)
        if m_editoria:
            materia_atual = {
                "editoria": m_editoria.group(1).strip(),
                "titulo": None,
                "corpo": "",
                "foto_arquivo": None,
                "foto_inteira": False,
                "foco": None,      # override manual do enquadramento (FOCO: direita, etc.)
                "boxes": [],       # blocos "ABRIR BOX ...": cada um {label, texto}
                "_em_boxes": False,
            }
            if pagina_atual is not None:
                pagina_atual["materias"].append(materia_atual)
            estado = "aguardando_titulo"
            continue

        m_foto = re.match(r"foto:\s*(.+)", text, re.IGNORECASE)
        if m_foto and materia_atual is not None:
            conteudo = m_foto.group(1).strip()
            # marcação opcional de acabamento (fundo desfocado / foto inteira /
            # sem cortar / banner): a foto entra inteira, com fundo desfocado
            marca = re.search(
                r"[\(\[][^)\]]*(desfoc\w*|inteira|sem\s+cortar|n[ãa]o\s+cortar|banner)[^)\]]*[\)\]]",
                conteudo, re.IGNORECASE)
            if marca:
                materia_atual["foto_inteira"] = True
                conteudo = (conteudo[:marca.start()] + conteudo[marca.end():]).strip()
            # override de enquadramento embutido na própria linha da foto:
            # "Foto: ARQUIVO.jpg (foco: direita)" ou com porcentagem opcional
            # "(foco: direita 20%)" — aceita ':' ou espaço, e '%' opcional. Usa o
            # PRIMEIRO marcador como valor e remove TODOS (o usuário pode ter
            # deixado vários testando — não podem sobrar grudados no nome).
            FOCO_MARK = re.compile(
                r"[\(\[][^)\]]*foco[:\s]+"
                r"(esquerda|direita|centro|meio|cima|topo|baixo|base|fundo)"
                r"\s*(\d+(?:[.,]\d+)?)?\s*%?"
                r"[^)\]]*[\)\]]", re.IGNORECASE)
            m_foco_inline = FOCO_MARK.search(conteudo)
            if m_foco_inline:
                direcao = m_foco_inline.group(1).lower()
                pct = m_foco_inline.group(2)
                materia_atual["foco"] = (direcao + " " + pct) if pct else direcao
                conteudo = FOCO_MARK.sub("", conteudo)
                conteudo = re.sub(r"\s{2,}", " ", conteudo).strip()
            m_unir = re.match(r"unir\s+essa:\s*(.+)", conteudo, re.IGNORECASE)
            if m_unir:
                # separador entre as fotos pode ser " e ESSA:" ou " + ESSA:"
                partes = re.split(r"\s*(?:\+|\be\b)\s*essa:\s*", m_unir.group(1),
                                  flags=re.IGNORECASE)
                materia_atual["foto_arquivo"] = [p.strip() for p in partes if p.strip()]
            else:
                materia_atual["foto_arquivo"] = conteudo
            estado = None
            continue

        # override de enquadramento em linha própria: "Foco: direita"
        m_foco = re.match(r"foco:\s*(.+)", text, re.IGNORECASE)
        if m_foco and materia_atual is not None:
            materia_atual["foco"] = m_foco.group(1).strip().lower()
            continue

        if text.startswith("=") or text.startswith("/"):
            continue

        if materia_atual is None:
            continue

        # primeira linha "normal" depois da editoria = título; o resto = corpo.
        # Não exige negrito, porque no formato grudado o título não é negrito.
        if estado == "aguardando_titulo":
            materia_atual["titulo"] = text
            estado = "lendo_corpo"
        elif estado == "lendo_corpo":
            # diretiva de box: "ABRIR BOX ...: <conteúdo do 1º box>". A partir
            # dela, cada linha lógica seguinte é um box separado. O texto que
            # vier ANTES da diretiva (na mesma linha) ainda é corpo.
            m_abre_box = re.search(
                r"abrir\s+(?:um\s+)?box\b[^:]*:\s*", text, re.IGNORECASE)
            if m_abre_box and not materia_atual["_em_boxes"]:
                antes = text[:m_abre_box.start()].strip()
                if antes:
                    materia_atual["corpo"] = (materia_atual["corpo"] + " " + antes).strip()
                materia_atual["_em_boxes"] = True
                depois = text[m_abre_box.end():].strip()
                if depois:
                    materia_atual["boxes"].append(depois)
            elif materia_atual["_em_boxes"]:
                materia_atual["boxes"].append(text.strip())
            else:
                materia_atual["corpo"] = (materia_atual["corpo"] + " " + text).strip()

    if pagina_atual and pagina_atual["materias"]:
        paginas.append(pagina_atual)

    # descarta matérias sem título (ex.: uma "Editoria:" duplicada/vazia que o
    # briefing às vezes tem) e páginas que ficaram sem nenhuma matéria
    for p in paginas:
        p["materias"] = [m for m in p["materias"] if m.get("titulo")]
    paginas = [p for p in paginas if p["materias"]]

    if paginas and paginas[0]["materias"]:
        paginas[0]["materias"][0]["is_capa"] = True

    # converte cada box (string "Rótulo: conteúdo") em {label, texto} e remove
    # a flag interna de controle
    for p in paginas:
        for m in p["materias"]:
            m.pop("_em_boxes", None)
            boxes_convertidos = []
            for box in m.get("boxes", []):
                partes = box.split(":", 1)
                if len(partes) == 2 and 0 < len(partes[0].strip()) <= 40:
                    boxes_convertidos.append({
                        "label": partes[0].strip(),
                        "texto": partes[1].strip(),
                    })
                else:
                    boxes_convertidos.append({"label": None, "texto": box.strip()})
            m["boxes"] = boxes_convertidos

    return paginas, cartazes


if __name__ == "__main__":
    import json
    import sys
    caminho = sys.argv[1] if len(sys.argv) > 1 else "briefing.docx"
    paginas, cartazes = parse_briefing(caminho)
    print(json.dumps(paginas, indent=2, ensure_ascii=False))
    print("CARTAZES:", cartazes)
