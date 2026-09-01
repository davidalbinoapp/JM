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


# Siglas das unidades da AGA que podem aparecer como tags no briefing.
# MG (Minas Gerais) é o destino PADRÃO — matéria/página sem tag nenhuma vai
# pra MG. As outras: QZ (Queiroz), LM (Lamego), CDS (Córrego do Sítio),
# GO (Goiás). Ver "tags de unidade" no CEREBRO_DO_JM.md.
UNIDADES_VALIDAS = ("MG", "QZ", "LM", "CDS", "GO")


def _extrai_unidades(text):
    """Lê as tags de unidade de uma linha do briefing — '(MG) (QZ) (LM) (CDS)'
    numa linha 'Unidades:' própria, ou embutidas na linha da página
    ('Página única >> ... (MG) (QZ) (LM)'). Devolve a lista de siglas VÁLIDAS
    na ordem em que aparecem, sem repetir; ou None se a linha não traz nenhuma
    tag conhecida.

    Filtra pelo conjunto UNIDADES_VALIDAS de propósito: assim um '(002)' de
    nome de foto ou um '(2)' qualquer não é confundido com unidade."""
    achados = []
    for m in re.findall(r"\(\s*([A-Za-z]{2,3})\s*\)", text):
        sig = m.upper()
        if sig in UNIDADES_VALIDAS and sig not in achados:
            achados.append(sig)
    return achados or None


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


def _extrai_cartaz(text, exigir_extensao=True):
    """Pega o nome do arquivo de uma linha de cartaz tipo
    'Descrição: ARQUIVO.jpg' ou 'Descrição – ARQUIVO.jpg' — aceita ':', '–',
    '—' e ' - ' como separador entre a descrição e o arquivo.

    `exigir_extensao=True` (padrão, usado FORA da seção 'Cartazes:', onde a linha
    pode ser texto comum): a linha PRECISA terminar num arquivo de imagem, senão
    retorna None — evita pescar parágrafo qualquer como cartaz.

    `exigir_extensao=False` (dentro da seção 'Cartazes:', onde toda linha já é
    sabidamente um cartaz): aceita o nome mesmo SEM a extensão .jpg/.jpeg/.png no
    fim, desde que a linha tenha o formato 'Descrição: nome' (um separador) e o
    nome tenha cara de arquivo. O casamento por similaridade (match_photo_file)
    acha o arquivo real na pasta mesmo sem a extensão. Corrige o caso do briefing
    em que o nome do cartaz foi digitado sem o '.jpg' no fim.

    O split é `maxsplit=1` no PRIMEIRO separador, então um ' - ' interno do nome
    (ex.: 'Cuiabá - Campanha') não quebra o filename."""
    tem_ext = bool(re.search(r"\.(?:jpg|jpeg|png)\s*$", text, re.IGNORECASE))
    if exigir_extensao and not tem_ext:
        return None
    partes = re.split(r"\s*[:–—]\s*|\s+-\s+", text, maxsplit=1)
    nome = partes[-1].strip() if len(partes) > 1 else text.strip()
    if tem_ext:
        m = re.search(r"(.+\.(?:jpg|jpeg|png))\s*$", nome, re.IGNORECASE)
        return m.group(1).strip() if m else None
    # Sem extensão: só aceita se veio de uma linha 'Descrição: nome' (houve
    # separador, len(partes) > 1) e o nome tem cara de arquivo (mín. 4 chars com
    # algum alfanumérico) — descarta traços decorativos ('=====') e sobras.
    if len(partes) > 1 and len(nome) >= 4 and re.search(r"[A-Za-z0-9]", nome):
        return nome
    return None


def parse_briefing(docx_path):
    doc = Document(docx_path)
    linhas = _linhas_logicas(doc)

    paginas = []
    cartazes = []
    pagina_atual = None
    materia_atual = None
    estado = None  # 'aguardando_titulo' | 'lendo_corpo'
    em_cartazes = False
    unidades_pendentes = None  # tags '(MG)(QZ)...' lidas numa linha 'Unidades:'
                               # à espera da próxima 'Página' pra colar nela

    for text, is_bold in linhas:
        low = text.lower()

        if low.startswith("cartazes"):
            em_cartazes = True
            continue

        # linha 'Unidades: (MG) (QZ) (LM) (CDS)' — guarda as siglas pra colar na
        # PRÓXIMA página. É a forma canônica de indicar destino no briefing
        # multi-unidade; sem essa linha (e sem tag inline na página) a página
        # cai no padrão MG lá na normalização (build_deck).
        if low.startswith("unidades"):
            unidades_pendentes = _extrai_unidades(text)
            continue

        # depois do cabeçalho "Cartazes:", toda linha com arquivo é um cartaz —
        # aqui aceitamos até nomes SEM extensão (exigir_extensao=False), porque a
        # seção só tem cartazes e o match_photo_file acha o arquivo por similaridade.
        if em_cartazes:
            nome = _extrai_cartaz(text, exigir_extensao=False)
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
            # unidade da página: prioriza tag INLINE na própria linha
            # ('Página única >> ... (MG) (QZ) (LM)'); senão usa o que veio da
            # linha 'Unidades:' logo acima; None = herda o padrão (MG) depois.
            unidades_inline = _extrai_unidades(text)
            pagina_atual = {
                "tipo": tipo,
                "materias": [],
                "unidades": unidades_inline or unidades_pendentes,
            }
            unidades_pendentes = None
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
                "focos_unir": None,  # foco POR foto numa união (lista alinhada a foto_arquivo)
                "mosaico": False,  # (mosaico): compõe a união em mosaico, não em fatias iguais
                "boxes": [],       # blocos "ABRIR BOX ...": cada um {label, texto}
                "boxes_intro": [],  # linha(s)-título da seção de boxes (terminam em ":")
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
            # marcação opcional "(mosaico)": compõe a união num mosaico de células
            # variadas (pra foto larga que não cabe num corte vertical) em vez de
            # fatias iguais. É retirada do texto antes de ler os nomes das fotos.
            m_mosaico = re.search(r"[\(\[]\s*mosaico\s*[\)\]]", conteudo, re.IGNORECASE)
            if m_mosaico:
                materia_atual["mosaico"] = True
                conteudo = (conteudo[:m_mosaico.start()] + conteudo[m_mosaico.end():]).strip()
            # marcação opcional de ZOOM num marcador SEPARADO: "(zoom: 15%)",
            # "(zoom 1.2)" ou "(zoom)" solto — além de aceitar o zoom escrito
            # DENTRO do foco ("(foco: esquerda 40% zoom 15%)"). Aqui a gente extrai
            # o marcador solto e injeta no foco mais abaixo, pra o motor (que lê o
            # zoom junto do foco) enxergar das duas formas. Sem isso, um "(zoom:)"
            # separado ficaria grudado no NOME do arquivo e sujaria o casamento.
            ZOOM_MARK = re.compile(r"[\(\[]\s*zoom\b[:\s]*([^)\]]*)[\)\]]", re.IGNORECASE)
            zoom_solto = None
            m_zoom = ZOOM_MARK.search(conteudo)
            if m_zoom:
                zoom_solto = ("zoom " + m_zoom.group(1).strip()).strip()
                conteudo = re.sub(r"\s{2,}", " ",
                                  conteudo[:m_zoom.start()] + conteudo[m_zoom.end():]).strip()
            # override de enquadramento embutido na linha da foto. FOCO_MARK
            # captura o CONTEÚDO inteiro do marcador "(foco: ...)". Numa foto
            # ÚNICA, junta todos os marcadores num foco só (permite 2D:
            # "(foco: direita 30% baixo 20%)"). Numa UNIÃO, o foco é lido POR
            # FOTO — cada parte carrega o seu, pra enquadrar cada foto da união
            # de forma independente. O "(2)" de nomes tipo "... (2).jpg" NÃO
            # casa (o marcador exige a palavra "foco:" dentro dos parênteses).
            FOCO_MARK = re.compile(
                r"[\(\[]\s*foco[:\s]+([^)\]]*)[\)\]]", re.IGNORECASE)
            m_unir = re.match(r"unir\s+essa:\s*(.+)", conteudo, re.IGNORECASE)
            if m_unir:
                # separador entre as fotos pode ser " e ESSA:", " + ESSA:",
                # ", ESSA:" (vírgula) ou ", e ESSA:" — tudo opcional, então
                # uma união de 3+ fotos com vírgula (ex.: "A, ESSA: B E ESSA: C")
                # separa certinho as 3, e não engole a do meio.
                partes = re.split(r"\s*(?:,\s*)?(?:\+|\be\b)?\s*essa:\s*", m_unir.group(1),
                                  flags=re.IGNORECASE)
                nomes, focos_partes = [], []
                for parte in partes:
                    parte = parte.strip()
                    if not parte:
                        continue
                    fs = FOCO_MARK.findall(parte)
                    focos_partes.append(
                        " ".join(f.strip() for f in fs if f.strip()).lower()
                        if fs else None)
                    parte = re.sub(r"\s{2,}", " ", FOCO_MARK.sub("", parte)).strip()
                    nomes.append(parte)
                # zoom solto vale pra união inteira: injeta em cada parte
                if zoom_solto:
                    focos_partes = [((f + " " + zoom_solto).strip() if f else zoom_solto)
                                    for f in focos_partes]
                materia_atual["foto_arquivo"] = nomes
                materia_atual["focos_unir"] = focos_partes
            else:
                focos_inline = FOCO_MARK.findall(conteudo)
                if focos_inline:
                    materia_atual["foco"] = " ".join(
                        f.strip() for f in focos_inline if f.strip()).lower()
                    conteudo = re.sub(r"\s{2,}", " ", FOCO_MARK.sub("", conteudo)).strip()
                # zoom solto entra no foco (cria o foco se não havia direção)
                if zoom_solto:
                    base = materia_atual.get("foco") or ""
                    materia_atual["foco"] = (base + " " + zoom_solto).strip()
                materia_atual["foto_arquivo"] = conteudo
            estado = None
            continue

        # override de enquadramento em linha própria: "Foco: direita"
        m_foco = re.match(r"foco:\s*(.+)", text, re.IGNORECASE)
        if m_foco and materia_atual is not None:
            materia_atual["foco"] = m_foco.group(1).strip().lower()
            continue

        # marcador "/ FECHA COM LOGO DA AGA /": a página atual encerra o JM
        # daquela(s) unidade(s) e recebe a logo de fechamento. Pode haver MAIS
        # de um no briefing multi-unidade (uma última página por unidade).
        if re.search(r"fecha\s+com\s+logo", low):
            if pagina_atual is not None:
                pagina_atual["fecha_com_logo"] = True
            continue

        if text.startswith("=") or text.startswith("/"):
            continue

        if materia_atual is None:
            continue

        # diretiva de box: tratada AQUI, fora da máquina de estado, porque ela
        # pode vir tanto ANTES quanto DEPOIS da linha "Foto:" — e a foto zera o
        # estado (estado=None), o que antes fazia o box vir depois da foto ser
        # descartado calado. Aceita os dois verbos que o briefing usa: "ABRIR
        # box ...:" e "CRIAR [um] box ...:" (ex.: "CRIAR UM BOX COM ESSAS
        # INFORMAÇÕES E INCLUIR NESTE SLIDE:"). O que vier ANTES da diretiva na
        # mesma linha ainda é corpo; uma vez aberta, cada linha lógica seguinte
        # é um box, até um novo bloco (editoria/matéria/página/separador)
        # reiniciar materia_atual. Só vale com título já lido (senão uma linha
        # de título que por acaso citasse "box" seria confundida).
        if materia_atual["titulo"]:
            m_abre_box = re.search(
                r"(?:abrir|criar)\s+(?:um\s+)?box\b[^:]*:\s*", text, re.IGNORECASE)
            if m_abre_box and not materia_atual["_em_boxes"]:
                antes = text[:m_abre_box.start()].strip()
                if antes:
                    materia_atual["corpo"] = (materia_atual["corpo"] + " " + antes).strip()
                materia_atual["_em_boxes"] = True
                depois = text[m_abre_box.end():].strip()
                if depois:
                    materia_atual["boxes"].append(depois)
                continue
            if materia_atual["_em_boxes"]:
                linha = text.strip()
                # uma linha que TERMINA em ":" (sem valor depois) é um título/
                # intro da seção de boxes — ex.: "Período de reconfiguração dos
                # rádios comunicadores:" — e NÃO uma caixa. Vai pro campo
                # `boxes_intro`, que o motor renderiza como HEADER acima das
                # caixas (nem no corpo, nem virando uma caixa). As caixas de
                # verdade são "Rótulo: valor" (Mina Cuiabá: ..., Mina Lamego: ...).
                if linha.endswith(":"):
                    materia_atual["boxes_intro"].append(linha)
                else:
                    materia_atual["boxes"].append(linha)
                continue

        # primeira linha "normal" depois da editoria = título; o resto = corpo.
        # Não exige negrito, porque no formato grudado o título não é negrito.
        if estado == "aguardando_titulo":
            materia_atual["titulo"] = text
            estado = "lendo_corpo"
        elif estado == "lendo_corpo":
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
