"""
Motor de geração do Jornal Mural (AGA).

Respeita a estrutura fixa do PPT:
  1. CAPA           -> slide 0, editado IN PLACE (nunca duplicado/removido)
  2. MATÉRIAS       -> bloco dinâmico, gerado a partir do briefing (dupla/única)
  3. CARTAZES       -> slides fixos, NUNCA tocados
  4. PÁGINA FINAL   -> slide fixo, NUNCA tocado nesta versão (thumbnail
                       reconstruído fica pra uma próxima etapa)
"""
import copy
import io
import os
import re
import math
import unicodedata
import zipfile
import tempfile
from PIL import Image
from pptx import Presentation
from pptx.util import Emu


def extract_media_from_template(template_path):
    """Extrai a pasta ppt/media do próprio template pra uma pasta temporária,
    assim o motor não depende de uma pasta de mídia preparada à parte —
    o .pptx do template já carrega a biblioteca de tags de editoria."""
    tmp_dir = tempfile.mkdtemp(prefix="jm_media_")
    with zipfile.ZipFile(template_path) as z:
        for name in z.namelist():
            if name.startswith("ppt/media/"):
                z.extract(name, tmp_dir)
    return os.path.join(tmp_dir, "ppt", "media")

import os as _os

TAGS_DIR = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), "tags_editorias")

EDITORIA_TAGS = {
    "SEGURANÇA": "seguranca.png",
    "NOSSA EMPRESA": "nossa_empresa.png",
    "LEGADO": "legado.png",
    "EVENTOS": "eventos.png",
    "MEIO AMBIENTE": "meio_ambiente.png",
    "DIVERSIDADE": "diversidade.png",
    "ESSÊNCIA AGA": "essencia_aga.png",
    "SAÚDE E BEM-ESTAR": "saude_e_bem_estar.png",
    "INOVAÇÃO E TECNOLOGIA": "inovacao_e_tecnologia.png",
    "PESSOAS": "pessoas.png",
    "EXCELÊNCIA OPERACIONAL": "excelencia_operacional.png",
}

FONTS_DIR = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), "fonts")
MODELS_DIR = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), "models")
YUNET_PATH = _os.path.join(MODELS_DIR, "face_detection_yunet.onnx")

# a fonte real do template é Arial. Ela não pode ser redistribuída aqui
# (é uma fonte comercial da Monotype), mas Mac e Windows normalmente já
# vêm com ela instalada — então tentamos usar a Arial de verdade primeiro,
# e só caímos pra Liberation Sans (compatível em métrica com a Arial,
# feita de propósito pra isso) se a Arial não for encontrada no sistema.
_CANDIDATOS_ARIAL_REGULAR = [
    "/Library/Fonts/Arial.ttf",
    "/System/Library/Fonts/Supplemental/Arial.ttf",
    "C:\\Windows\\Fonts\\arial.ttf",
    "/usr/share/fonts/truetype/msttcorefonts/Arial.ttf",
]
_CANDIDATOS_ARIAL_BOLD = [
    "/Library/Fonts/Arial Bold.ttf",
    "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
    "C:\\Windows\\Fonts\\arialbd.ttf",
    "/usr/share/fonts/truetype/msttcorefonts/Arial_Bold.ttf",
]


def _fonte_abre(caminho):
    """Só aceita a fonte se o PIL conseguir abrir de fato. Existir o arquivo
    não basta: pode ser um .ttf quebrado, um atalho, ou um arquivo baixado
    pela metade — e aí o erro só apareceria lá na frente, no meio da
    montagem, como um 'cannot open resource' sem explicação."""
    from PIL import ImageFont
    try:
        ImageFont.truetype(caminho, 40)
        return True
    except Exception:
        return False


def _resolver_caminho_fonte(negrito):
    reserva = _os.path.join(FONTS_DIR, "LiberationSans-Bold.ttf" if negrito
                            else "LiberationSans-Regular.ttf")
    for caminho in (_CANDIDATOS_ARIAL_BOLD if negrito else _CANDIDATOS_ARIAL_REGULAR):
        if _os.path.isfile(caminho) and _fonte_abre(caminho):
            return caminho
    if _os.path.isfile(reserva) and _fonte_abre(reserva):
        return reserva
    raise RuntimeError(
        "Nenhuma fonte utilizável foi encontrada. O sistema procura a Arial "
        "instalada no computador e, se não achar, usa a Liberation Sans que "
        "vem junto do projeto. Confira se a pasta 'fonts/' está do lado do "
        "jm_engine.py com os arquivos LiberationSans-Regular.ttf e "
        "LiberationSans-Bold.ttf (cerca de 400 KB cada)."
    )


FONT_REGULAR = _resolver_caminho_fonte(negrito=False)
FONT_BOLD = _resolver_caminho_fonte(negrito=True)


def _contar_linhas_texto(texto, largura_emu, tamanho_pt, negrito=False):
    """Estima quantas linhas o texto vai ocupar dentro de uma caixa de
    largura `largura_emu`, pra permitir reposicionar elementos abaixo dela
    de forma proporcional ao conteúdo real (em vez de uma posição fixa)."""
    import math
    from PIL import ImageFont
    fonte_path = FONT_BOLD if negrito else FONT_REGULAR
    px_size = max(int(tamanho_pt * 4), 8)
    font = ImageFont.truetype(fonte_path, px_size)
    # desconta as margens internas padrão da caixa de texto do PowerPoint
    # (0,1" de cada lado). Sem isso o cálculo usa a largura cheia e acha que
    # cabe mais texto por linha do que realmente cabe — contando uma linha a
    # menos em títulos que ficam no limiar, e deixando o corpo subir demais.
    MARGEM_INTERNA_EMU = 91440  # 0,1 polegada
    largura_util = max(largura_emu - 2 * MARGEM_INTERNA_EMU, 12700)
    largura_px = (largura_util / 12700) * 4

    def _largura(s):
        b = font.getbbox(s)
        return b[2] - b[0]

    def _linhas_do_token(s):
        # Um token indivisível (uma palavra longa, ou um par grudado pelo
        # espaço inseparável da prevenção de viúva) que seja MAIS LARGO que a
        # caixa é quebrado pelo PowerPoint mesmo sem ponto de quebra. Aqui a
        # gente conta quantas linhas ele realmente ocupa. Sem isso o contador
        # trata esse token como 1 linha só, subestima a altura do título e o
        # corpo é posicionado alto demais, invadindo o título.
        return max(1, math.ceil(_largura(s) / largura_px))

    total_linhas = 0
    for paragrafo in texto.split("\n"):
        # split só no espaço normal: o espaço inseparável ( ) usado pra
        # evitar viúva NÃO é ponto de quebra. Se usarmos o .split() padrão do
        # Python (que trata   como espaço), o contador acha que cabe uma
        # linha a menos do que o PowerPoint realmente quebra — e aí o corpo é
        # posicionado alto demais e invade o título.
        palavras = [p for p in paragrafo.split(" ") if p]
        if not palavras:
            total_linhas += 1  # linha em branco também ocupa altura
            continue
        linha_atual = palavras[0]
        linhas_neste_paragrafo = _linhas_do_token(palavras[0])
        for palavra in palavras[1:]:
            teste = linha_atual + " " + palavra
            if _largura(teste) > largura_px:
                # quebra: a palavra desce pra próxima linha (e, se ela mesma
                # for mais larga que a caixa, ocupa mais de uma linha)
                linhas_neste_paragrafo += _linhas_do_token(palavra)
                linha_atual = palavra
            else:
                linha_atual = teste
        total_linhas += linhas_neste_paragrafo
    return max(total_linhas, 1)


def _fonte_pt_do_titulo(title_box, default=40):
    """Tamanho (pt) da fonte atual do título, pra medir quebra de linha."""
    runs = title_box.text_frame.paragraphs[0].runs
    if runs and runs[0].font.size:
        return runs[0].font.size / 12700
    return default


def _evitar_viuva(texto, largura_emu=None, tamanho_pt=None):
    """Troca o último espaço por espaço inseparável, pra a última palavra
    do título nunca ficar sozinha numa linha (viúva tipográfica).

    Se a largura da caixa e o tamanho da fonte forem informados, só gruda
    as duas últimas palavras quando elas CABEM juntas numa linha. Grudar um
    par que não cabe é contraproducente: o PowerPoint quebra o bloco
    inseparável mesmo assim, sobrando um caco (ex.: um "?" sozinho) — pior
    que a viúva que a gente queria evitar. Nesse caso, deixa quebrar natural."""
    palavras = texto.split(" ")
    if len(palavras) < 2:
        return texto
    if largura_emu and tamanho_pt:
        from PIL import ImageFont
        MARGEM_INTERNA_EMU = 91440  # 0,1" de cada lado, como no contador de linhas
        largura_util = max(largura_emu - 2 * MARGEM_INTERNA_EMU, 12700)
        largura_px = (largura_util / 12700) * 4
        font = ImageFont.truetype(FONT_BOLD, max(int(tamanho_pt * 4), 8))
        par = palavras[-2] + " " + palavras[-1]
        b = font.getbbox(par)
        if (b[2] - b[0]) > largura_px:
            return texto  # nao cabe grudado: deixa quebrar natural
    return " ".join(palavras[:-1]) + " " + palavras[-1]

# Abaixo desta fração da largura da caixa, uma linha do título que NÃO é a
# última é considerada "órfã" (curta demais no meio do título) e dispara o
# rebalanceamento. Acima disso, a quebra gulosa é tida como boa e mantida.
LIMIAR_ORFAO_TITULO = 0.55


def _quebrar_titulo_balanceado(texto, largura_emu, tamanho_pt, negrito=True):
    """Reequilibra as quebras do título: mantém o MESMO número de linhas que a
    quebra natural (gulosa) do PowerPoint, mas distribui as palavras pra
    minimizar a largura da linha mais cheia — evita que uma palavra curta
    ('CMG') fique órfã sozinha enquanto outra linha fica lotada. Como cada linha
    fica MENOS cheia que na gulosa, nunca transborda mais do que já transbordava
    (a maior linha do balanceado ≤ a maior linha da gulosa, que já cabia).
    Devolve o título com '\\n' nos pontos de quebra; inalterado se cabe em 1
    linha ou faltar largura/fonte. Preserva o espaço inseparável da viúva (o
    par grudado continua um token só)."""
    if not (largura_emu and tamanho_pt):
        return texto
    try:
        from PIL import ImageFont
        MARG = 91440  # 0,1" de cada lado, como no contador de linhas
        largura_px = (max(largura_emu - 2 * MARG, 12700) / 12700) * 4
        fonte = FONT_BOLD if negrito else FONT_REGULAR
        font = ImageFont.truetype(fonte, max(int(tamanho_pt * 4), 8))

        def W(s):
            b = font.getbbox(s)
            return b[2] - b[0]

        tokens = [t for t in texto.split(" ") if t]  # \xa0 (viúva) fica no token
        if len(tokens) < 2:
            return texto
        # quebra gulosa (o que o PowerPoint faria sozinho): guarda as linhas
        greedy = []
        cur = tokens[0]
        for t in tokens[1:]:
            if W(cur + " " + t) > largura_px:
                greedy.append(cur)
                cur = t
            else:
                cur = cur + " " + t
        greedy.append(cur)
        n = len(greedy)
        if n < 2:
            return texto
        # GATE: só rebalanceia se a quebra gulosa deixou uma linha curta NO MEIO
        # (não a última) — o órfão tipo "CMG". Se o guloso já está bem
        # distribuído, não mexe (evita regressão em título que já quebrava bem).
        if not any(W(l) < LIMIAR_ORFAO_TITULO * largura_px for l in greedy[:-1]):
            return texto
        m = len(tokens)
        # dp[k][i] = menor "largura da maior linha" ao quebrar tokens[i:] em k linhas
        INF = float("inf")
        dp = [[INF] * (m + 1) for _ in range(n + 1)]
        nxt = [[m] * (m + 1) for _ in range(n + 1)]
        dp[0][m] = 0.0
        for k in range(1, n + 1):
            for i in range(m - 1, -1, -1):
                # a 1ª linha leva tokens[i:j]; o resto (tokens[j:]) em k-1 linhas.
                # j vai até deixar pelo menos (k-1) tokens pro resto.
                for j in range(i + 1, m - (k - 1) + 1):
                    cand = max(W(" ".join(tokens[i:j])), dp[k - 1][j])
                    if cand < dp[k][i]:
                        dp[k][i] = cand
                        nxt[k][i] = j
        if dp[n][0] == INF:
            return texto
        linhas = []
        i = 0
        for k in range(n, 0, -1):
            j = nxt[k][i]
            linhas.append(" ".join(tokens[i:j]))
            i = j
        return "\n".join(linhas)
    except Exception:
        return texto


def _reposicionar_corpo_apos_titulo(title_box, body_box, texto_titulo=None):
    """Ajusta o topo da caixa de corpo com base na altura REAL que o título
    vai ocupar (em vez de confiar numa posição fixa do template, que só
    funciona se o título tiver exatamente o mesmo número de linhas do
    título original usado no design).

    Conta o texto que ESTÁ na caixa (com o espaço inseparável da viúva já
    aplicado e com a largura/fonte já ajustadas por causa do QR), e não o
    título original — senão a contagem quebra em um número de linhas
    diferente do que o PowerPoint realmente mostra, e o corpo acaba
    posicionado alto demais, invadindo o título."""
    run = title_box.text_frame.paragraphs[0].runs[0]
    tamanho_pt = (run.font.size / 12700) if run.font.size else 28
    texto_real = title_box.text_frame.text or texto_titulo or ""
    linhas = _contar_linhas_texto(texto_real, title_box.width, tamanho_pt, negrito=True)
    altura_linha_emu = int(tamanho_pt * 12700 * 1.25)
    gap_padrao_emu = int(0.5 * 360000)
    novo_top = title_box.top + linhas * altura_linha_emu + gap_padrao_emu
    body_box.top = novo_top


def _ajustar_corpo_sem_transbordar(body_box, texto_corpo, limite_inferior_emu, tamanho_min_pt=11):
    """Depois de reposicionar o corpo, garante que o texto não vaza pro
    espaço do próximo elemento abaixo — reduz o tamanho da fonte
    gradualmente até caber na altura disponível. É isso que impede
    sobreposição quando o título E o corpo são grandes ao mesmo tempo."""
    from pptx.util import Pt
    if not body_box.text_frame.paragraphs[0].runs:
        return
    run = body_box.text_frame.paragraphs[0].runs[0]
    tamanho_pt = (run.font.size / 12700) if run.font.size else 19
    disponivel = limite_inferior_emu - body_box.top - int(0.3 * 360000)
    if disponivel <= 0:
        return

    while tamanho_pt > tamanho_min_pt:
        linhas = _contar_linhas_texto(texto_corpo, body_box.width, tamanho_pt, negrito=False)
        altura_necessaria = linhas * tamanho_pt * 12700 * 1.25
        if altura_necessaria <= disponivel:
            break
        tamanho_pt -= 1

    for paragraph in body_box.text_frame.paragraphs:
        for r in paragraph.runs:
            r.font.size = Pt(tamanho_pt)


def _ajustar_titulo_largura_reduzida(title_box, texto_titulo, tamanho_min_pt=24):
    """Quando a caixa de título fica mais estreita (ex.: por causa do QR
    code), reduz a fonte até nenhuma palavra ultrapassar a largura —
    evita quebra no meio da palavra e mantém o texto legível."""
    from PIL import ImageFont
    from pptx.util import Pt
    if not title_box.text_frame.paragraphs[0].runs:
        return
    run = title_box.text_frame.paragraphs[0].runs[0]
    tamanho_pt = (run.font.size / 12700) if run.font.size else 40
    tamanho_original = tamanho_pt
    largura_px = (title_box.width / 12700) * 4
    # o espaço inseparável (viúva) gruda duas palavras num "pedaço" só —
    # é esse pedaço que não pode quebrar, então é ele que precisa caber,
    # não a palavra isolada
    pedacos = texto_titulo.split(" ")

    while tamanho_pt > tamanho_min_pt:
        font = ImageFont.truetype(FONT_BOLD, max(int(tamanho_pt * 4), 8))
        maior_largura = max(font.getbbox(p)[2] - font.getbbox(p)[0] for p in pedacos)
        if maior_largura <= largura_px * 0.92:
            break
        tamanho_pt -= 1

    if tamanho_pt != tamanho_original:
        for r in title_box.text_frame.paragraphs[0].runs:
            r.font.size = Pt(tamanho_pt)


# ---------------------------------------------------------------------------
# Utilidades de shape
# ---------------------------------------------------------------------------

def duplicate_slide(prs, template_slide):
    """Cria um slide novo (sempre no fim da apresentação) copiando o XML de
    todos os shapes de `template_slide`, E as relações de imagem — assim
    fotos/tags que SERÃO substituídas funcionam via _replace_picture/_replace_tag,
    e imagens que NÃO são substituídas (logo, estoque de tags não usadas)
    continuam válidas em vez de ficarem quebradas."""
    new_slide = prs.slides.add_slide(template_slide.slide_layout)
    for shape in list(new_slide.shapes):
        new_slide.shapes._spTree.remove(shape._element)
    for shape in template_slide.shapes:
        new_el = copy.deepcopy(shape._element)
        new_slide.shapes._spTree.append(new_el)
    for rId, rel in template_slide.part.rels.items():
        if rel.reltype == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image":
            new_slide.part.rels._rels[rId] = rel
    return new_slide


def _set_run_text_keep_format(text_frame, new_text):
    para = text_frame.paragraphs[0]
    runs = para.runs
    if not runs:
        return
    runs[0].text = new_text
    for extra in runs[1:]:
        extra._r.getparent().remove(extra._r)
    for p in text_frame.paragraphs[1:]:
        p._p.getparent().remove(p._p)


LARANJA_BOX = (0xFE, 0xA4, 0x2F)  # laranja do rótulo dos boxes (vem do template)


def _preencher_box_textbox(box_shape, label, texto):
    """Preenche uma caixa de box do template: o rótulo mantém o negrito laranja
    que já vem do template e o conteúdo entra em preto regular (igual ao
    modelo do cliente: 'CEA:' laranja + horários em preto)."""
    from pptx.dml.color import RGBColor
    tf = box_shape.text_frame
    para = tf.paragraphs[0]
    runs = para.runs
    if not runs:
        return
    base = runs[0]  # herda formatação (negrito + laranja) do template
    base.text = (label + ": ") if label else ""
    for extra in runs[1:]:
        extra._r.getparent().remove(extra._r)
    for pextra in tf.paragraphs[1:]:
        pextra._p.getparent().remove(pextra._p)
    run_txt = para.add_run()
    run_txt.text = texto
    run_txt.font.bold = False
    run_txt.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
    if base.font.size:
        run_txt.font.size = base.font.size
    if base.font.name:
        run_txt.font.name = base.font.name


def _boxes_no_corpo(body_box, boxes_data):
    """Fallback pra quando o slide NÃO tem placeholders de box (ex.: um layout
    que não seja o do cliente): anexa os boxes ao próprio corpo como parágrafos
    (rótulo laranja em negrito + texto preto). Sem as linhas divisórias, mas o
    conteúdo aparece legível e estruturado em qualquer layout."""
    from pptx.dml.color import RGBColor
    tf = body_box.text_frame
    base_run = tf.paragraphs[0].runs[0] if tf.paragraphs[0].runs else None
    tam = base_run.font.size if (base_run and base_run.font.size) else None
    for box in boxes_data:
        p = tf.add_paragraph()
        if box["label"]:
            rl = p.add_run()
            rl.text = box["label"] + ": "
            rl.font.bold = True
            rl.font.color.rgb = RGBColor(*LARANJA_BOX)
            if tam:
                rl.font.size = tam
        rt = p.add_run()
        rt.text = box["texto"]
        rt.font.bold = False
        if tam:
            rt.font.size = tam


def _render_boxes(slide, body_box, placeholders, linhas, boxes_data, warnings, contexto=""):
    """Renderiza os boxes da matéria. Estratégia (independente de número de
    slide): se o layout já traz placeholders de box (caixas extras + linhas
    divisórias), PREENCHE em vez de apagar; se não traz, cria os boxes dentro
    do corpo (fallback). Ajusta a quantidade de placeholders/linhas ao número
    real de boxes do briefing."""
    n = len(boxes_data)
    if not placeholders:
        _boxes_no_corpo(body_box, boxes_data)
        return

    # preenche os primeiros placeholders com os boxes disponíveis
    for i in range(min(n, len(placeholders))):
        b = boxes_data[i]
        _preencher_box_textbox(placeholders[i], b["label"], b["texto"])

    # menos boxes que placeholders: remove os que sobraram (e as linhas de baixo)
    sobrando = placeholders[n:]
    for extra in sobrando:
        extra._element.getparent().remove(extra._element)
    if sobrando and linhas:
        for ln in sorted(linhas, key=lambda s: s.top or 0, reverse=True)[:len(sobrando)]:
            ln._element.getparent().remove(ln._element)

    # mais boxes que placeholders: junta o excedente no último box, pra não
    # perder informação, e avisa
    if n > len(placeholders):
        resto = boxes_data[len(placeholders):]
        extra_txt = " ".join(
            ((b["label"] + ": ") if b["label"] else "") + b["texto"] for b in resto)
        p = placeholders[-1].text_frame.add_paragraph()
        r = p.add_run()
        r.text = extra_txt
        warnings.append(
            f"{contexto}: {n} boxes no briefing, mas o layout tem "
            f"{len(placeholders)} — os extras foram juntados no último box.")


def _place_qr_code(slide, qr_path, tag_box, slide_width_emu):
    """Posiciona o QR code (sempre quadrado) encostado na borda direita,
    alinhado verticalmente com a tag da matéria à qual ele pertence."""
    from pptx.util import Emu
    lado = tag_box.height  # QR é quadrado, mesma altura de referência da tag (6.09cm ~ constante)
    lado = Emu(int(6.09 * 360000))
    margem_direita = Emu(int(2.0 * 360000))
    left = slide_width_emu - lado - margem_direita
    top = tag_box.top
    slide.shapes.add_picture(qr_path, left, top, lado, lado)


CORTE_DIAGONAL_PROPORCAO = 0.092  # ~9.2% da altura, medido nas fotos do template

# Viés do corte vertical usado como FALLBACK (quando não há rosto detectado):
# quanto do excedente sai do TOPO (o resto sai da base). 0.5 = centralizado
# (decepa a cabeça de quem está em pé); 0.2 = puxa pro topo, preservando
# cabeça e sacrificando mais o rodapé. Com rosto, quem manda é
# _offset_corte_vertical (dá folga acima da cabeça).
VIES_CORTE_VERTICAL = 0.2


# ---------------------------------------------------------------------------
# Enquadramento inteligente: quando a foto é mais larga que a máscara e precisa
# ser cortada nas laterais, em vez de cortar SEMPRE pelo centro (que decepa
# quem está fora do meio), o motor detecta os rostos e desloca o corte pra
# manter as pessoas enquadradas. Se o OpenCV/modelo não existirem, ou se a foto
# não tiver rosto (prédio, gráfico), cai no corte central de sempre — nunca
# piora e nunca quebra.
# ---------------------------------------------------------------------------
_YUNET = None
_YUNET_TENTOU = False
_ROSTOS_CACHE = {}       # path -> lista de rostos (evita redetecção na mesma build)
_SALIENCIA_CACHE = {}    # path -> (mapa, W0, H0) da saliência


# Reforço dos rostos SOBRE o mapa de saliência. O rosto é assunto forte, mas
# não pode deixar uma MULTIDÃO sequestrar o corte: o peso é pela CONFIANÇA² do
# rosto (o palestrante nítido/frontal, score alto, pesa muito mais que as nucas
# da plateia) vezes a raiz da área (a área conta, mas sem explodir). Peso baixo
# de propósito — a saliência é quem manda; o rosto só ajusta.
FATOR_REFORCO_ROSTO = 0.004
# Acima de tantos rostos a cena é "multidão": o snap de borda é desligado (senão
# trocaria o assunto por um rosto qualquer encostado na borda da janela).
MAX_ROSTOS_PARA_SNAP = 3


def _get_yunet():
    """Carrega o detector de rosto YuNet uma vez só (cacheado). Devolve None se
    o OpenCV ou o modelo não estiverem disponíveis."""
    global _YUNET, _YUNET_TENTOU
    if _YUNET_TENTOU:
        return _YUNET
    _YUNET_TENTOU = True
    try:
        import cv2
        if _os.path.isfile(YUNET_PATH):
            _YUNET = cv2.FaceDetectorYN.create(YUNET_PATH, "", (320, 320),
                                               score_threshold=0.6)
    except Exception:
        _YUNET = None
    return _YUNET


def _detectar_rostos(image_path):
    """Rostos (x, y, w, h, score) em coordenadas da imagem original, cacheado
    por caminho (a mesma foto é consultada várias vezes numa build)."""
    if image_path not in _ROSTOS_CACHE:
        _ROSTOS_CACHE[image_path] = _detectar_rostos_raw(image_path)
    return _ROSTOS_CACHE[image_path]


def _detectar_rostos_raw(image_path):
    """Rostos (x, y, w, h, score) em coordenadas da imagem original. Roda numa
    versão reduzida (~1024px) pra velocidade. Devolve [] se não houver
    detector/modelo/rosto — aí o chamador usa o enquadramento padrão."""
    det = _get_yunet()
    if det is None:
        return []
    try:
        import cv2
        import numpy as np
        # carrega via PIL (como o resto do motor) e converte pro formato do
        # OpenCV — evita o "libpng warning: iCCP" que o cv2.imread cospe no
        # terminal ao ler PNGs com perfil de cor levemente torto (é só ruído,
        # mas suja a saída). PIL lê essas PNGs sem reclamar.
        img = cv2.cvtColor(np.array(Image.open(image_path).convert("RGB")),
                           cv2.COLOR_RGB2BGR)
        H, W = img.shape[:2]
        escala = 1024.0 / W if W > 1024 else 1.0
        sw, sh = max(int(W * escala), 1), max(int(H * escala), 1)
        det.setInputSize((sw, sh))
        entrada = cv2.resize(img, (sw, sh)) if escala != 1.0 else img
        _, faces = det.detect(entrada)
        if faces is None:
            return []
        return [(float(f[0] / escala), float(f[1] / escala),
                 float(f[2] / escala), float(f[3] / escala), float(f[-1]))
                for f in faces]
    except Exception:
        return []


# ---------------------------------------------------------------------------
# Saliência visual (numpy puro, SEM dependência nova): responde "onde está o
# assunto?" quando o rosto sozinho não basta. Método do resíduo espectral
# (Hou & Zhang): realça o que é SINGULAR na imagem (o palestrante isolado, o
# globo brilhante) e SUPRIME padrões repetitivos (a textura de cadeiras/nucas
# de uma plateia). É isso que impede a multidão de sequestrar o corte e o que
# reposiciona uma foto sem rosto cujo assunto está fora do centro.
# ---------------------------------------------------------------------------
def _box_blur(a, k):
    """Média de janela k x k via soma acumulada (numpy puro, sem scipy)."""
    import numpy as np
    if k < 2:
        return a
    pad = k // 2
    ap = np.pad(a, pad, mode="edge")
    cs = np.cumsum(np.cumsum(ap, axis=0), axis=1)
    cs = np.pad(cs, ((1, 0), (1, 0)), mode="constant")
    H, W = a.shape
    y1 = np.arange(H)[:, None]; x1 = np.arange(W)[None, :]
    y2 = y1 + k; x2 = x1 + k
    tot = cs[y2, x2] - cs[y1, x2] - cs[y2, x1] + cs[y1, x1]
    return tot / (k * k)


def _mapa_saliencia(image_path, largura=64):
    """Mapa de saliência (0..1) por resíduo espectral, numa versão minúscula da
    foto. Devolve (mapa, W0, H0) ou None se numpy/Pillow falharem — aí o
    chamador cai no corte de sempre (nunca quebra, nunca piora)."""
    if image_path in _SALIENCIA_CACHE:
        return _SALIENCIA_CACHE[image_path]
    dados = None
    try:
        import numpy as np
        im = Image.open(image_path).convert("L")
        W0, H0 = im.size
        h = max(int(round(largura * H0 / W0)), 16)
        g = np.asarray(im.resize((largura, h), Image.BILINEAR), dtype=np.float64)
        F = np.fft.fft2(g)
        logamp = np.log(np.abs(F) + 1e-8)
        fase = np.angle(F)
        residuo = logamp - _box_blur(logamp, 3)
        S = np.abs(np.fft.ifft2(np.exp(residuo + 1j * fase))) ** 2
        S = _box_blur(S, 5)
        S -= S.min()
        if S.max() > 0:
            S /= S.max()
        dados = (S, W0, H0)
    except Exception:
        dados = None
    _SALIENCIA_CACHE[image_path] = dados
    return dados


def _centro_importancia(image_path, rostos, eixo):
    """Centro do 'assunto' ao longo de um eixo (0 = x/horizontal, 1 = y/
    vertical): energia da saliência + reforço dos rostos. Devolve px na imagem
    original, ou None se não houver sinal (aí volta pro corte padrão)."""
    dados = _mapa_saliencia(image_path)
    if dados is None:
        return None
    try:
        import numpy as np
        S, W0, H0 = dados
        S = S.copy()
        hs, ws = S.shape
        esc_x = W0 / ws
        esc_y = H0 / hs
        for (x, y, w, h, s) in rostos:
            gx0 = max(0, int(x / esc_x)); gx1 = min(ws, max(int((x + w) / esc_x), gx0 + 1))
            gy0 = max(0, int(y / esc_y)); gy1 = min(hs, max(int((y + h) / esc_y), gy0 + 1))
            S[gy0:gy1, gx0:gx1] += (s ** 2) * np.sqrt(w * h) * FATOR_REFORCO_ROSTO
        if eixo == 0:
            perfil = S.sum(axis=0); esc = esc_x
        else:
            perfil = S.sum(axis=1); esc = esc_y
        total = float(perfil.sum())
        if total <= 0:
            return None
        idx = np.arange(len(perfil))
        centro = float((idx * perfil).sum() / total)
        return (centro + 0.5) * esc
    except Exception:
        return None


# O corte central só é abandonado se mover pro assunto capturar mais que este
# tanto da energia de saliência. Protege logo/gráfico já bem centrado de ser
# descentralizado por uma saliência fraca e levemente assimétrica (ex.: o peso
# de tinta de um texto puxando o centro-de-massa pro lado).
LIMIAR_ENERGIA_DESCENTRAR = 0.95


def _offset_horizontal_saliencia_sem_rosto(image_path, largura_total, new_w):
    """x0 do corte horizontal para foto SEM rosto: usa a saliência pra achar o
    assunto (globo brilhante, objeto fora do centro), mas só sai do centro se
    a janela no assunto capturar materialmente mais energia que a janela
    central — senão devolve o centro (não mexe no que já estava bem centrado).
    Devolve None se a saliência não estiver disponível."""
    dados = _mapa_saliencia(image_path)
    if dados is None:
        return None
    try:
        import numpy as np
        S, W0, H0 = dados
        col = S.sum(axis=0)
        ws = len(col)
        esc = W0 / ws
        total = float(col.sum())
        if total <= 0:
            return None
        lo, hi = 0, largura_total - new_w
        x0_mid = (largura_total - new_w) // 2
        centro = (float((np.arange(ws) * col).sum() / total) + 0.5) * esc
        x0_cent = max(lo, min(hi, int(round(centro - new_w / 2.0))))

        def energia(x0px):
            g0 = x0px / esc
            g1 = (x0px + new_w) / esc
            i0 = max(0, int(np.floor(g0)))
            i1 = min(ws, int(np.ceil(g1)))
            return float(col[i0:i1].sum())

        e_assunto = energia(x0_cent)
        if e_assunto <= 0:
            return x0_mid
        if energia(x0_mid) >= LIMIAR_ENERGIA_DESCENTRAR * e_assunto:
            return x0_mid                 # o centro já pega quase tudo: não descentraliza
        return x0_cent
    except Exception:
        return None


def _parse_foco(foco):
    """Interpreta o 'FOCO:' do briefing -> (direcao, fracao).
    - direcao: 'esquerda'/'direita'/'centro'/'cima'/'baixo' (ou None).
    - fracao: 0.0..1.0 = o quanto empurrar A PARTIR do enquadramento automático
      rumo àquela borda. None quando não veio número -> vale o comportamento
      antigo (empurra até a borda = 100%). Ex.: 'direita 20' -> ('direita', 0.2);
      'direita' -> ('direita', None); 'centro' -> ('centro', None)."""
    if not foco:
        return None, None
    s = unicodedata.normalize("NFKD", str(foco)).encode("ascii", "ignore").decode().lower()
    direcao = None
    for chave, canon in (
        ("esquerda", "esquerda"), ("esq", "esquerda"), ("left", "esquerda"),
        ("direita", "direita"), ("dir", "direita"), ("right", "direita"),
        ("centro", "centro"), ("center", "centro"), ("meio", "centro"),
        ("cima", "cima"), ("topo", "cima"), ("top", "cima"),
        ("baixo", "baixo"), ("base", "baixo"), ("bottom", "baixo"), ("fundo", "baixo"),
    ):
        if chave in s:
            direcao = canon
            break
    frac = None
    m = re.search(r"(\d+(?:[.,]\d+)?)", s)
    if m:
        frac = max(0.0, min(1.0, float(m.group(1).replace(",", ".")) / 100.0))
    return direcao, frac


def _foco_horizontal(foco):
    """(direcao, fracao) do FOCO só quando é um eixo HORIZONTAL."""
    d, frac = _parse_foco(foco)
    return (d, frac) if d in ("esquerda", "direita", "centro") else (None, None)


def _foco_vertical(foco):
    """(direcao, fracao) do FOCO só quando é um eixo VERTICAL."""
    d, frac = _parse_foco(foco)
    return (d, frac) if d in ("cima", "baixo", "centro") else (None, None)


def _empurrar_para_borda(auto, lo, hi, direcao_baixa, direcao_alta, direcao, frac):
    """Aplica o override direcional sobre o offset automático `auto`: empurra
    `frac` do caminho entre `auto` e a borda (lo ou hi). frac None = borda (100%,
    o comportamento antigo do FOCO sem número)."""
    if direcao == direcao_baixa:
        f = 1.0 if frac is None else frac
        return int(round(max(lo, min(hi, auto - f * (auto - lo)))))
    if direcao == direcao_alta:
        f = 1.0 if frac is None else frac
        return int(round(max(lo, min(hi, auto + f * (hi - auto)))))
    return auto  # centro / None -> mantém o automático


def _centro_horizontal_rostos(rostos):
    """Centro x ponderado pela ÁREA de cada rosto: os rostos maiores/mais
    próximos (os assuntos principais) pesam mais, e um eventual falso-positivo
    pequeno e distante quase não desloca o corte."""
    soma = peso_total = 0.0
    for (x, y, w, h, s) in rostos:
        peso = w * h
        soma += (x + w / 2.0) * peso
        peso_total += peso
    return (soma / peso_total) if peso_total > 0 else None


def _offset_corte_horizontal(new_image_path, largura_total, new_w, foco=None):
    """x0 do corte horizontal. Sem FOCO (ou 'FOCO: centro') usa o enquadramento
    automático (assunto). 'FOCO: esquerda/direita' empurra a PARTIR do
    automático rumo à borda; com porcentagem ('direita 20') empurra só essa
    fração do caminho, sem número vai até a borda (o comportamento antigo)."""
    lo, hi = 0, largura_total - new_w
    auto = _offset_horizontal_auto(new_image_path, largura_total, new_w)
    direcao, frac = _foco_horizontal(foco)
    # convenção: o ASSUNTO se move na direção da seta. "direita" (assunto vai pra
    # direita do quadro) = janela pra ESQUERDA = x0 menor (lo). "esquerda" =
    # janela pra direita = x0 maior (hi).
    return _empurrar_para_borda(auto, lo, hi, "direita", "esquerda", direcao, frac)


def _offset_horizontal_auto(new_image_path, largura_total, new_w):
    """Enquadramento horizontal AUTOMÁTICO (sem override), pelo ASSUNTO:
    1) centro de importância = saliência (suprime multidão, acha o singular) +
       reforço dos rostos;
    2) fallback ao centro ponderado dos rostos / centro geométrico se a
       saliência não estiver disponível.
    Quando TODOS os rostos cabem na janela, ainda garante contê-los (mantém o
    comportamento de sempre pras fotos de 1–3 pessoas — baixa regressão)."""
    lo, hi = 0, largura_total - new_w
    x0_centro = (largura_total - new_w) // 2

    rostos = _detectar_rostos(new_image_path)
    if not rostos:
        # sem rosto: saliência com guarda de energia (não descentraliza gráfico
        # já centrado); cai no centro se a saliência não estiver disponível
        x0 = _offset_horizontal_saliencia_sem_rosto(new_image_path, largura_total, new_w)
        return x0 if x0 is not None else x0_centro

    rx0 = min(r[0] for r in rostos)
    rx1 = max(r[0] + r[2] for r in rostos)
    todos_cabem = (rx1 - rx0) <= new_w
    if todos_cabem:
        # as pessoas cabem inteiras na janela: o assunto são ELAS → centraliza
        # no centro dos rostos. A saliência NÃO entra aqui de propósito: num
        # trio contra um fundo laranja/luminoso, a saliência do fundo puxava o
        # corte e jogava as pessoas pro lado. Com rosto que cabe, quem manda é
        # o rosto (comportamento estável de sempre).
        centro = _centro_horizontal_rostos(rostos)
        if centro is None:
            centro = x0_centro + new_w / 2.0
    else:
        # multidão (os rostos não cabem todos): aí sim a saliência acha o
        # assunto singular e suprime a repetição da plateia.
        centro = _centro_importancia(new_image_path, rostos, eixo=0)
        if centro is None:
            c = _centro_horizontal_rostos(rostos)
            centro = c if c is not None else (x0_centro + new_w / 2.0)
    x0 = int(round(centro - new_w / 2.0))
    x0 = max(lo, min(hi, x0))
    if todos_cabem:                                # garante conter todos os rostos
        x0 = min(max(x0, int(rx1 - new_w)), int(rx0))
        x0 = max(lo, min(hi, x0))
    # snap só com POUCOS rostos (sujeito claro). Em multidão, o snap trocaria
    # o assunto por um rosto qualquer de borda — então é desligado.
    if len(rostos) <= MAX_ROSTOS_PARA_SNAP:
        x0 = _snap_fora_dos_rostos(x0, new_w, rostos, lo, hi)
    return x0


def _snap_fora_dos_rostos(x0, new_w, rostos, lo, hi):
    def fatia(pos):
        return any(x < pos < x + w or x < pos + new_w < x + w
                   for (x, y, w, h, s) in rostos)
    if not fatia(x0):
        return x0
    candidatos = [lo, hi]
    for (x, y, w, h, s) in rostos:
        # posições que encostam a borda ESQ ou DIR da janela na borda de um rosto
        candidatos += [x, x + w, x - new_w, x + w - new_w]
    validos = [c for c in candidatos if lo <= c <= hi and not fatia(c)]
    return int(min(validos, key=lambda c: abs(c - x0))) if validos else x0


def _offset_corte_vertical(new_image_path, altura_total, new_h, foco=None):
    """y0 do corte vertical. Sem FOCO (ou 'FOCO: centro') usa o automático.
    'FOCO: cima/baixo' empurra a PARTIR do automático rumo à borda; com
    porcentagem ('baixo 20') empurra só essa fração, sem número vai até a
    borda."""
    lo, hi = 0, altura_total - new_h
    auto = _offset_vertical_auto(new_image_path, altura_total, new_h)
    direcao, frac = _foco_vertical(foco)
    # mesma convenção do horizontal: o assunto se move na direção da seta.
    # "baixo" (assunto pra baixo) = janela pra CIMA = y0 menor (lo); "cima" =
    # janela pra baixo = y0 maior (hi).
    return _empurrar_para_borda(auto, lo, hi, "baixo", "cima", direcao, frac)


def _offset_vertical_auto(new_image_path, altura_total, new_h):
    """Corte vertical AUTOMÁTICO: com rostos, dá FOLGA acima da cabeça do rosto
    mais alto (não decepa cabeça/cabelo/boné) e joga o excedente pra BASE; sem
    rosto, usa o viés fixo pro topo (VIES_CORTE_VERTICAL), o comportamento
    antigo — o vertical automático fica intocado."""
    lo, hi = 0, altura_total - new_h
    y0_padrao = int((altura_total - new_h) * VIES_CORTE_VERTICAL)
    rostos = _detectar_rostos(new_image_path)
    if not rostos:
        return y0_padrao
    fy0 = min(y for (x, y, w, h, s) in rostos)                 # topo do rosto mais alto
    fh = next(h for (x, y, w, h, s) in rostos if y == fy0)     # altura desse rosto
    folga = int(0.7 * fh)                                       # cabelo/cabeça/boné acima da caixa
    y0 = max(lo, min(hi, fy0 - folga))
    # se a BASE da janela cortar um rosto ao meio, SOBE só o necessário pra
    # deixá-lo inteiro fora (nunca desce, pra não voltar a cortar as cabeças)
    for (x, y, w, h, s) in sorted(rostos, key=lambda r: r[1]):
        if y < y0 + new_h < y + h:
            y0 = max(lo, y - new_h)
            break
    return y0


def _aplicar_corte_diagonal(im, proporcao=CORTE_DIAGONAL_PROPORCAO, cor_fundo=(255, 255, 255)):
    """Pinta o triângulo no canto superior-esquerdo, reproduzindo o corte
    diagonal que já vem queimado nas fotos do Jornal Mural."""
    im = im.copy()
    w, h = im.size
    perna = int(round(h * proporcao))
    if perna <= 0:
        return im
    px = im.load()
    for y in range(perna):
        limite_x = perna - y  # reta de (perna,0) a (0,perna)
        for x in range(min(limite_x, w)):
            px[x, y] = cor_fundo
    return im


def _compose_blurred_fit(image_path, largura_alvo_px, altura_alvo_px, escurecer=0.7):
    """Encaixa a foto INTEIRA na caixa (sem cortar nada), preenchendo o vazio
    com uma versão ampliada e desfocada da própria foto. É o acabamento que
    se faz à mão pra fotos que não podem ser cortadas — banner, telão, arte
    com texto nas bordas: em vez de perder as laterais, mostra tudo e o fundo
    borrado dá o preenchimento, sem barra preta dura."""
    from PIL import ImageFilter, ImageEnhance
    im = Image.open(image_path).convert("RGB")
    alvo_ratio = largura_alvo_px / altura_alvo_px
    w, h = im.size
    src_ratio = w / h

    # fundo: a foto preenchendo toda a caixa (cover), bem desfocada e um
    # pouco escurecida, virando uma textura neutra atrás da peça
    if src_ratio > alvo_ratio:
        nb = int(h * alvo_ratio)
        x0 = (w - nb) // 2
        fundo = im.crop((x0, 0, x0 + nb, h))
    else:
        nb = int(w / alvo_ratio)
        y0 = (h - nb) // 2
        fundo = im.crop((0, y0, w, y0 + nb))
    fundo = fundo.resize((largura_alvo_px, altura_alvo_px))
    fundo = fundo.filter(ImageFilter.GaussianBlur(
        radius=max(largura_alvo_px, altura_alvo_px) // 30))
    fundo = ImageEnhance.Brightness(fundo).enhance(escurecer)

    # frente: a foto inteira (contain), centralizada sobre o fundo
    escala = min(largura_alvo_px / w, altura_alvo_px / h)
    fw, fh = int(w * escala), int(h * escala)
    frente = im.resize((fw, fh))
    canvas = fundo.copy()
    canvas.paste(frente, ((largura_alvo_px - fw) // 2, (altura_alvo_px - fh) // 2))
    return canvas


def _replace_picture(picture_shape, new_image_path, corte_diagonal=True,
                     foto_inteira=False, foco=None, warnings=None, contexto=""):
    left, top, width, height = (picture_shape.left, picture_shape.top,
                                 picture_shape.width, picture_shape.height)
    target_ratio = width / height

    if foto_inteira:
        # acabamento "foto inteira": mostra tudo, com fundo desfocado. Renderiza
        # o composto na proporção exata da caixa, a ~150 dpi.
        dpi = 150
        tw = max(int(width / 914400 * dpi), 300)
        th = max(int(height / 914400 * dpi), 300)
        im = _compose_blurred_fit(new_image_path, tw, th)
    else:
        im = Image.open(new_image_path).convert("RGB")
        w, h = im.size
        src_ratio = w / h
        # avisa se o FOCO pedido é pro eixo que essa foto NÃO corta (ex.: pediu
        # cima/baixo mas a foto é cortada nas laterais) — senão o override fica
        # "sem efeito" calado, o que confunde.
        if foco and warnings is not None:
            _dir, _ = _parse_foco(foco)
            horizontal = src_ratio > target_ratio
            if _dir in ("esquerda", "direita") and not horizontal:
                warnings.append(f"{contexto}: FOCO '{foco}' (lateral) sem efeito — essa "
                                f"foto é cortada em cima/baixo; use cima/baixo.")
            elif _dir in ("cima", "baixo") and horizontal:
                warnings.append(f"{contexto}: FOCO '{foco}' (cima/baixo) sem efeito — essa "
                                f"foto é cortada nas laterais; use esquerda/direita.")
        if src_ratio > target_ratio:
            new_w = int(h * target_ratio)
            # enquadra pelo ASSUNTO (saliência + rostos); respeita o override
            # FOCO: do briefing; cai no centro se não houver sinal
            x0 = _offset_corte_horizontal(new_image_path, w, new_w, foco=foco)
            im = im.crop((x0, 0, x0 + new_w, h))
        else:
            new_h = int(w / target_ratio)
            # dá folga acima das cabeças e joga o excedente pra base (cai no
            # viés fixo pro topo se não houver rosto); respeita o FOCO: vertical
            y0 = _offset_corte_vertical(new_image_path, h, new_h, foco=foco)
            im = im.crop((0, y0, w, y0 + new_h))

    if corte_diagonal:
        im = _aplicar_corte_diagonal(im)

    buf = io.BytesIO()
    im.save(buf, format="JPEG", quality=92)
    buf.seek(0)

    slide = picture_shape.part.slide
    old_el = picture_shape._element
    sp_parent = old_el.getparent()
    idx = list(sp_parent).index(old_el)
    sp_parent.remove(old_el)
    new_pic = slide.shapes.add_picture(buf, left, top, width, height)
    new_el = new_pic._element
    new_el.getparent().remove(new_el)
    sp_parent.insert(idx, new_el)


def _resolver_editoria(editoria_name):
    key = editoria_name.strip().upper()
    if key in EDITORIA_TAGS:
        return key
    # tenta casar por prefixo/substring (ex.: "EXCELÊNCIA" -> "EXCELÊNCIA OPERACIONAL")
    candidatos = [k for k in EDITORIA_TAGS if k.startswith(key) or key in k]
    if len(candidatos) == 1:
        return candidatos[0]
    return None


def _replace_tag(tag_shape, editoria_name, limite_direito_emu=None):
    """Troca a tag de editoria pela arte certa. A tag mantém a altura do
    template e a largura sai da proporção do PNG — então editoria de nome
    comprido ('INOVAÇÃO E TECNOLOGIA') gera uma tag mais larga. Se essa
    largura fizer a tag passar do `limite_direito_emu` (ex.: a borda esquerda
    da foto na capa), a tag é reduzida proporcionalmente (largura E altura)
    até caber, em vez de invadir a foto."""
    from pptx.util import Emu
    key = _resolver_editoria(editoria_name)
    fname = EDITORIA_TAGS.get(key) if key else None
    if not fname:
        raise ValueError(f"Editoria '{editoria_name}' não mapeada em EDITORIA_TAGS. "
                          f"Adicione o par nome->arquivo antes de gerar.")
    path = _os.path.join(TAGS_DIR, fname)
    left, top, height = tag_shape.left, tag_shape.top, tag_shape.height
    im = Image.open(path)
    w, h = im.size
    new_width = int(height * (w / h))

    slide = tag_shape.part.slide

    # nunca deixa a tag passar da borda direita do slide; e, se veio um limite
    # mais apertado (a foto, na capa), respeita esse limite.
    pres = slide.part.package.presentation_part.presentation
    limite = limite_direito_emu if limite_direito_emu is not None \
        else pres.slide_width - Emu(int(0.5 * 360000))
    if left + new_width > limite > left:
        escala = (limite - left) / new_width
        new_width = int(new_width * escala)
        height = int(height * escala)

    old_el = tag_shape._element
    sp_parent = old_el.getparent()
    idx = list(sp_parent).index(old_el)
    sp_parent.remove(old_el)
    new_pic = slide.shapes.add_picture(path, left, top, new_width, height)
    new_el = new_pic._element
    new_el.getparent().remove(new_el)
    sp_parent.insert(idx, new_el)
    return new_pic


def _top_cm(shape):
    return shape.top / 360000


# ---------------------------------------------------------------------------
# Casamento de arquivo de foto (nome do briefing -> arquivo real na pasta)
# ---------------------------------------------------------------------------

def _normalize(s):
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]", "", s.lower())


def _compose_merged_photo(photo_paths, target_ratio, foco=None):
    """Junta 2 ou mais fotos lado a lado (com um filete branco de separação),
    cada uma enquadrada pelas PESSOAS (mesmo corte por rosto das fotos normais:
    detecta os rostos e evita cortar alguém ao meio; cai no centro se não houver
    rosto), formando UMA imagem só na proporção da caixa. O corte diagonal do
    canto é aplicado depois, sobre o conjunto inteiro, como em qualquer foto."""
    n = len(photo_paths)
    divisor_frac = 0.008
    altura_ref = 1000
    largura_total = int(altura_ref * target_ratio)
    divisor_px = max(2, int(largura_total * divisor_frac))
    largura_util = largura_total - divisor_px * (n - 1)
    largura_cada = largura_util // n

    canvas = Image.new("RGB", (largura_total, altura_ref), (255, 255, 255))
    x = 0
    sub_ratio = largura_cada / altura_ref
    for path in photo_paths:
        im = Image.open(path).convert("RGB")
        w, h = im.size
        src_ratio = w / h
        if src_ratio > sub_ratio:
            new_w = int(h * sub_ratio)
            # enquadra pelo assunto (saliência + rostos); mesmo FOCO: aplicado a
            # cada foto da união
            x0 = _offset_corte_horizontal(path, w, new_w, foco=foco)
            im = im.crop((x0, 0, x0 + new_w, h))
        else:
            new_h = int(w / sub_ratio)
            y0 = _offset_corte_vertical(path, h, new_h, foco=foco)  # folga acima das cabeças
            im = im.crop((0, y0, w, y0 + new_h))
        im = im.resize((largura_cada, altura_ref))
        canvas.paste(im, (x, 0))
        x += largura_cada + divisor_px

    return canvas


LIMIAR_CASAMENTO_FOTO = 0.80  # abaixo disso, considera "não encontrado"


def match_photo_file(foto_arquivo, fotos_dir):
    """Casa o nome de foto do briefing com o arquivo real na pasta, tolerando
    diferenças de acento/espaço/maiúscula (via _normalize). Usa similaridade
    real (difflib) em vez de só prefixo em comum: nomes com o mesmo prefixo e
    o mesmo sufixo mas miolo diferente — típico dos cartazes, tipo
    'AGA_CARTAZ JM_Show OOP Nirvana_220626-KN' vs '...Orgulho LGBTQIA+...' —
    NÃO podem casar. Quando nada passa do limiar, retorna None (o chamador
    avisa 'não encontrado', em vez de trazer o arquivo errado calado)."""
    from difflib import SequenceMatcher
    if not foto_arquivo:
        return None
    alvo = _normalize(foto_arquivo)
    if not alvo:
        return None
    melhor, melhor_score = None, 0.0
    for cand in os.listdir(fotos_dir):
        cnorm = _normalize(cand)
        if not cnorm:
            continue
        score = SequenceMatcher(None, alvo, cnorm).ratio()
        # um nome curto contido inteiro no outro é um casamento forte
        if len(min(alvo, cnorm, key=len)) >= 6 and (alvo in cnorm or cnorm in alvo):
            score = max(score, 0.9)
        if score > melhor_score:
            melhor, melhor_score = cand, score
    if melhor and melhor_score >= LIMIAR_CASAMENTO_FOTO:
        return os.path.join(fotos_dir, melhor)
    return None


def resolve_foto(foto_arquivo, fotos_dir, target_ratio, warnings, contexto="", foco=None):
    """Resolve o campo 'foto_arquivo' (string única ou lista pra unir) num
    caminho de arquivo pronto pra usar em _replace_picture."""
    if isinstance(foto_arquivo, list):
        paths = []
        for nome in foto_arquivo:
            p = match_photo_file(nome, fotos_dir)
            if not p:
                warnings.append(f"{contexto}: foto '{nome}' (parte de uma união) não encontrada.")
                return None
            paths.append(p)
        composta = _compose_merged_photo(paths, target_ratio, foco=foco)
        # nome determinístico (não depende do hash aleatório do Python), então
        # re-rodar sobrescreve em vez de acumular; a faxina em build_deck apaga
        # de vez depois do .pptx salvo.
        import hashlib
        chave = hashlib.md5("||".join(foto_arquivo).encode("utf-8")).hexdigest()[:12]
        tmp_path = os.path.join(fotos_dir, f"_merged_{chave}.jpg")
        composta.save(tmp_path, quality=95)
        return tmp_path
    else:
        p = match_photo_file(foto_arquivo, fotos_dir)
        if not p:
            warnings.append(f"{contexto}: foto '{foto_arquivo}' não encontrada na pasta de fotos.")
        return p


# ---------------------------------------------------------------------------
# Populamento de slides
# ---------------------------------------------------------------------------

# Respiro (cm) entre a tag de editoria e a foto na capa, quando a editoria é
# comprida e a tag encosta na foto. Ver populate_capa.
FOLGA_TAG_FOTO_CM = 0.6


def populate_capa(slide, materia, fotos_dir, media_dir, warnings):
    textboxes = [s for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip()]
    pictures = [s for s in slide.shapes if s.shape_type == 13]

    title_box, body_box = None, None
    for tb in textboxes:
        runs = tb.text_frame.paragraphs[0].runs
        if not runs:
            continue
        if runs[0].font.bold and runs[0].font.size and runs[0].font.size >= 400000:
            title_box = tb
        elif not runs[0].font.bold:
            body_box = tb

    tag_box = None
    photo_box = None
    for p in pictures:
        try:
            fmt = p.image.ext
        except Exception:
            continue
        h_cm = p.height / 360000
        if fmt == "png" and 1.5 < h_cm < 2.2:
            tag_box = p
        elif fmt in ("jpg", "jpeg") and h_cm > 15:
            if photo_box is None or p.width * p.height > photo_box.width * photo_box.height:
                photo_box = p

    if title_box:
        pt_titulo = _fonte_pt_do_titulo(title_box)
        titulo_final = _evitar_viuva(materia["titulo"], title_box.width, pt_titulo)
        # rebalanceia as quebras (evita palavra órfã sozinha numa linha, tipo o
        # "CMG" da capa de Queiroz); mantém o mesmo nº de linhas
        titulo_final = _quebrar_titulo_balanceado(titulo_final, title_box.width, pt_titulo)
        _set_run_text_keep_format(title_box.text_frame, titulo_final)
    else:
        warnings.append("CAPA: não encontrei a caixa de título pra substituir.")
    if body_box:
        _set_run_text_keep_format(body_box.text_frame, materia["corpo"])
    else:
        warnings.append("CAPA: não encontrei a caixa de corpo pra substituir.")
    if title_box and body_box:
        _reposicionar_corpo_apos_titulo(title_box, body_box, materia["titulo"])
        limite_inferior = slide.part.package.presentation_part.presentation.slide_height - Emu(int(1.0 * 360000))
        _ajustar_corpo_sem_transbordar(body_box, materia["corpo"], limite_inferior)
    if tag_box:
        # a tag não pode invadir a foto (que na capa fica à direita): passa a
        # borda esquerda da foto, menos uma folga de respiro, como limite pra
        # tag encolher. 0.3cm deixava a tag "quase encostando" na foto quando a
        # editoria é comprida (ex.: "INOVAÇÃO E TECNOLOGIA"); 0.6cm dá um respiro
        # visível sem encolher a tag além do necessário.
        limite_tag = None
        if photo_box is not None and photo_box.left > tag_box.left:
            limite_tag = photo_box.left - Emu(int(FOLGA_TAG_FOTO_CM * 360000))
        _replace_tag(tag_box, materia["editoria"], limite_direito_emu=limite_tag)
    else:
        warnings.append("CAPA: não encontrei a tag de editoria pra substituir.")

    if photo_box:
        ratio = photo_box.width / photo_box.height
        foto_path = resolve_foto(materia.get("foto_arquivo"), fotos_dir, ratio, warnings, "CAPA",
                                 foco=materia.get("foco"))
        if foto_path:
            _replace_picture(photo_box, foto_path,
                             foto_inteira=materia.get("foto_inteira", False),
                             foco=materia.get("foco"), warnings=warnings, contexto="CAPA")
    else:
        warnings.append("CAPA: não encontrei a foto principal pra substituir.")


def populate_double_slide(slide, story_top, story_bottom, fotos_dir, media_dir, warnings):
    pictures = [s for s in slide.shapes if s.shape_type == 13]
    textboxes = [s for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip()]

    pictures.sort(key=_top_cm)
    textboxes.sort(key=_top_cm)

    photo_top, tag_top = sorted(pictures[:2], key=lambda s: s.width, reverse=True)
    photo_bottom, tag_bottom = sorted(pictures[2:], key=lambda s: s.width, reverse=True)

    title_top = body_top = title_bottom = body_bottom = None
    for tb in textboxes:
        runs = tb.text_frame.paragraphs[0].runs
        if not runs:
            continue
        bold = runs[0].font.bold
        if _top_cm(tb) < 16:
            if bold:
                title_top = tb
            else:
                body_top = tb
        else:
            if bold:
                title_bottom = tb
            else:
                body_bottom = tb

    slide_height = slide.part.package.presentation_part.presentation.slide_height
    limite_top_story = photo_top.top + photo_top.height  # corpo não passa da base da própria foto
    limite_bottom_story = min(photo_bottom.top + photo_bottom.height,
                               slide_height - Emu(int(1.0 * 360000)))

    for story, title_box, body_box, photo_box, tag_box, limite_inferior in [
        (story_top, title_top, body_top, photo_top, tag_top, limite_top_story),
        (story_bottom, title_bottom, body_bottom, photo_bottom, tag_bottom, limite_bottom_story),
    ]:
        qr_bottom = None
        if story.get("qr_path"):
            lado_qr = Emu(int(6.09 * 360000))
            margem_qr = Emu(int(2.0 * 360000))
            qr_bottom = title_box.top + lado_qr  # QR fica alinhado ao topo da tag/título
            limite_direito = slide.part.package.presentation_part.presentation.slide_width - lado_qr - margem_qr - Emu(int(0.3 * 360000))
            nova_largura = limite_direito - title_box.left
            if nova_largura < title_box.width:
                title_box.width = nova_largura
                # o corpo NÃO encolhe por causa do QR — ele fica só na linha
                # do título/tag, e o corpo (texto corrido, alinhado à esquerda)
                # não chega a alcançar essa área mesmo usando a largura cheia

        titulo_final = _evitar_viuva(story["titulo"], title_box.width,
                                     _fonte_pt_do_titulo(title_box))
        _set_run_text_keep_format(title_box.text_frame, titulo_final)
        _set_run_text_keep_format(body_box.text_frame, story["corpo"])
        if story.get("qr_path"):
            _ajustar_titulo_largura_reduzida(title_box, titulo_final)
        _reposicionar_corpo_apos_titulo(title_box, body_box, story["titulo"])
        _ajustar_corpo_sem_transbordar(body_box, story["corpo"], limite_inferior)
        _replace_tag(tag_box, story["editoria"])
        foto_path = resolve_foto(story.get("foto_arquivo"), fotos_dir,
                                  photo_box.width / photo_box.height, warnings,
                                  f"Matéria '{story['titulo']}'", foco=story.get("foco"))
        if foto_path:
            _replace_picture(photo_box, foto_path,
                             foto_inteira=story.get("foto_inteira", False),
                             foco=story.get("foco"), warnings=warnings,
                             contexto=f"Matéria '{story['titulo']}'")
        if story.get("qr_path"):
            _place_qr_code(slide, story["qr_path"], tag_box, slide.part.package.presentation_part.presentation.slide_width)


def populate_single_slide(slide, materia, fotos_dir, media_dir, warnings):
    """Matéria de página única (não-capa). Baseado no slide-modelo enviado
    pelo cliente (CEA/Centro de Memória): tag + título + corpo + foto.
    Caixas extras de informação (ex.: horários) não são geradas
    automaticamente — ficam como um ajuste manual pontual, se existirem."""
    pictures = [s for s in slide.shapes if s.shape_type == 13]
    textboxes = [s for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip()]
    textboxes.sort(key=_top_cm)

    on_slide_pics = [p for p in pictures if p.left >= 0]
    on_slide_pics_sorted = sorted(on_slide_pics, key=lambda p: p.width * p.height, reverse=True)
    photo_box = on_slide_pics_sorted[0] if on_slide_pics_sorted else None
    tag_box = None
    for p in on_slide_pics_sorted[1:]:
        h_cm = p.height / 360000
        if 1.5 < h_cm < 2.2:
            tag_box = p
            break

    title_box = textboxes[0] if textboxes else None
    body_box = textboxes[1] if len(textboxes) > 1 else None
    GAP_PADRAO = Emu(int(0.5 * 360000))

    if tag_box:
        novo_tag = _replace_tag(tag_box, materia["editoria"])
        # a tag sempre alinha com o topo da foto — hoje o motor sempre
        # remove as caixas extras de horário/info (não há dado do briefing
        # pra preenchê-las), então o slide final é sempre o caso "limpo".
        if photo_box:
            novo_tag.top = photo_box.top
            if title_box:
                title_box.top = novo_tag.top + novo_tag.height + GAP_PADRAO
    else:
        warnings.append(f"Página única '{materia['titulo']}': não encontrei a tag de editoria.")

    if title_box:
        titulo_final = _evitar_viuva(materia["titulo"], title_box.width,
                                     _fonte_pt_do_titulo(title_box))
        _set_run_text_keep_format(title_box.text_frame, titulo_final)
    else:
        warnings.append(f"Página única '{materia['titulo']}': não encontrei a caixa de título.")
    if body_box:
        _set_run_text_keep_format(body_box.text_frame, materia["corpo"])
    else:
        warnings.append(f"Página única '{materia['titulo']}': não encontrei a caixa de corpo.")
    boxes_data = materia.get("boxes") or []
    linhas_slide = [s for s in slide.shapes if s.shape_type == 9]
    if title_box and body_box:
        _reposicionar_corpo_apos_titulo(title_box, body_box, materia["titulo"])
        limite_inferior = slide.part.package.presentation_part.presentation.slide_height - Emu(int(1.0 * 360000))
        # com boxes, o corpo (intro) não pode crescer até a área dos boxes:
        # o limite passa a ser o topo da 1ª linha divisória
        if boxes_data and linhas_slide:
            topo_boxes = min(s.top for s in linhas_slide if s.top is not None)
            limite_inferior = min(limite_inferior, topo_boxes - Emu(int(0.2 * 360000)))
        _ajustar_corpo_sem_transbordar(body_box, materia["corpo"], limite_inferior)
    if photo_box:
        foto_path = resolve_foto(materia.get("foto_arquivo"), fotos_dir,
                                  photo_box.width / photo_box.height, warnings,
                                  f"Página única '{materia['titulo']}'", foco=materia.get("foco"))
        if foto_path:
            _replace_picture(photo_box, foto_path,
                             foto_inteira=materia.get("foto_inteira", False),
                             foco=materia.get("foco"), warnings=warnings,
                             contexto=f"Página única '{materia['titulo']}'")
    else:
        warnings.append(f"Página única '{materia['titulo']}': não encontrei a foto principal.")

    # boxes do briefing ("ABRIR BOX ...:"): se a matéria tem boxes, PREENCHE os
    # placeholders do layout (mantendo as linhas divisórias); se não tem,
    # remove as caixas extras e as linhas (comportamento padrão do layout limpo)
    placeholders = textboxes[2:]
    if boxes_data:
        _render_boxes(slide, body_box, placeholders, linhas_slide, boxes_data,
                      warnings, contexto=f"Página única '{materia['titulo']}'")
    else:
        for extra_tb in placeholders:
            extra_tb._element.getparent().remove(extra_tb._element)
        for line_shape in linhas_slide:
            line_shape._element.getparent().remove(line_shape._element)

    # a logo de encerramento só deve aparecer na ÚLTIMA matéria do JM inteiro —
    # como esse template une o slide-modelo com a logo já embutida, ela é
    # removida aqui sempre, e adicionada de volta (uma única vez) pelo
    # orquestrador (build_deck) depois de saber qual slide é realmente o último.
    for p in list(slide.shapes):
        if p.shape_type == 13:
            h_cm = p.height / 360000
            w_cm = p.width / 360000
            if 2.5 < w_cm < 6 and 2 < h_cm < 4.5 and p.left > 0:
                p._element.getparent().remove(p._element)


def _remover_assets_fora_da_prancheta(slide):
    """Remove o que está estacionado fora da área visível do slide e não deve
    ir pro arquivo final: o estoque de tags não usadas que o template carrega,
    e enfeites soltos na 'mesa' ao lado da prancheta (ex.: um quadrado laranja
    parado num left negativo). Trata qualquer tipo de shape — não só imagem —
    porque esses enfeites costumam ser auto-shapes (retângulos), que a versão
    antiga deixava passar. Só remove o que está 100% fora; um elemento que
    sangra pra fora mas ainda aparece em parte é preservado."""
    pres = slide.part.package.presentation_part.presentation
    sw, sh = pres.slide_width, pres.slide_height
    for shape in list(slide.shapes):
        l, t, w, h = shape.left, shape.top, shape.width, shape.height
        if None in (l, t, w, h):
            continue
        totalmente_fora = (l + w <= 0) or (t + h <= 0) or (l >= sw) or (t >= sh)
        # comportamento antigo (mantido): imagem de estoque parada em left/top
        # negativo é removida mesmo que sangre um pouco pra dentro
        imagem_parada_fora = (shape.shape_type == 13 and (l < 0 or t < 0))
        if totalmente_fora or imagem_parada_fora:
            shape._element.getparent().remove(shape._element)


LOGO_FECHAMENTO_PATH = _os.path.join(TAGS_DIR, "logo_aga_fechamento.png")


def _adicionar_logo_final(slide):
    left = Emu(int(36.2 * 360000))
    top = Emu(int(25.22 * 360000))
    width = Emu(int(4.65 * 360000))
    height = Emu(int(3.1 * 360000))
    margem = Emu(int(0.6 * 360000))
    slide.shapes.add_picture(LOGO_FECHAMENTO_PATH, left, top, width, height)

    # dá "respiro": qualquer caixa de texto que invada a área horizontal/vertical
    # da logo tem a largura reduzida pra parar antes dela
    narrowed_title, narrowed_body = None, None
    for shape in slide.shapes:
        if shape.has_text_frame and shape.text_frame.text.strip():
            buffer_seguranca = Emu(int(0.5 * 360000))
            invade_vertical = (shape.top < top + height) and (shape.top + shape.height > top - buffer_seguranca)
            invade_horizontal = (shape.left + shape.width) > (left - margem)
            if invade_vertical and invade_horizontal and shape.left < left:
                shape.width = (left - margem) - shape.left
                runs = shape.text_frame.paragraphs[0].runs
                if runs and runs[0].font.bold:
                    narrowed_title = shape
                else:
                    narrowed_body = shape

    if narrowed_title is not None and narrowed_body is not None:
        texto_titulo = narrowed_title.text_frame.text
        texto_corpo = narrowed_body.text_frame.text
        _reposicionar_corpo_apos_titulo(narrowed_title, narrowed_body, texto_titulo)
        limite_inferior = slide.part.package.presentation_part.presentation.slide_height - Emu(int(0.3 * 360000))
        _ajustar_corpo_sem_transbordar(narrowed_body, texto_corpo, limite_inferior)
    elif narrowed_body is not None:
        texto_corpo = narrowed_body.text_frame.text
        limite_inferior = slide.part.package.presentation_part.presentation.slide_height - Emu(int(0.3 * 360000))
        _ajustar_corpo_sem_transbordar(narrowed_body, texto_corpo, limite_inferior)


# ---------------------------------------------------------------------------
# Orquestrador principal
# ---------------------------------------------------------------------------

_PADRAO_MERGED = re.compile(r"^_merged_[0-9a-f]+\.jpg$", re.IGNORECASE)


def _limpar_merged(fotos_dir):
    """Apaga os JPEGs intermediários de fotos unidas (`_merged_*.jpg`) que o
    motor grava na pasta da semana ao compor uma união. Depois do .pptx salvo,
    os bytes já ficam embutidos no arquivo — esses intermediários não servem
    mais. Rodado no início do build (limpa restos de execuções anteriores) e no
    fim (limpa os desta execução), pra não acumular na pasta do cliente."""
    try:
        for nome in os.listdir(fotos_dir):
            if _PADRAO_MERGED.match(nome):
                try:
                    os.remove(os.path.join(fotos_dir, nome))
                except OSError:
                    pass
    except OSError:
        pass


def build_deck(template_path, paginas, fotos_dir, output_path, media_dir=None,
               capa_slide_index=0, materia_dupla_template_index=1,
               materia_unica_template_index=4,
               old_materia_slide_indices=(1, 2, 3, 4),
               cartazes=None, cartaz_slide_indices=(5, 6, 7, 8)):
    warnings = []
    _limpar_merged(fotos_dir)   # tira restos de fotos unidas de execuções anteriores
    if media_dir is None:
        media_dir = extract_media_from_template(template_path)
    prs = Presentation(template_path)
    slides_originais = list(prs.slides)

    capa_materia = paginas[0]["materias"][0]
    populate_capa(slides_originais[capa_slide_index], capa_materia, fotos_dir, media_dir, warnings)
    _remover_assets_fora_da_prancheta(slides_originais[capa_slide_index])

    template_dupla = slides_originais[materia_dupla_template_index]
    template_unica = slides_originais[materia_unica_template_index]
    novas_materia_slides = []

    for pagina in paginas[1:]:
        if pagina["tipo"] == "dupla":
            m1, m2 = (pagina["materias"] + [None])[:2]
            new_slide = duplicate_slide(prs, template_dupla)
            populate_double_slide(new_slide, m1, m2 or {"editoria": m1["editoria"], "titulo": "", "corpo": "", "foto_arquivo": None},
                                   fotos_dir, media_dir, warnings)
            _remover_assets_fora_da_prancheta(new_slide)
            novas_materia_slides.append(new_slide)
        else:
            new_slide = duplicate_slide(prs, template_unica)
            populate_single_slide(new_slide, pagina["materias"][0], fotos_dir, media_dir, warnings)
            _remover_assets_fora_da_prancheta(new_slide)
            novas_materia_slides.append(new_slide)

    novos_cartaz_slides = []
    if novas_materia_slides:
        _adicionar_logo_final(novas_materia_slides[-1])
    else:
        _adicionar_logo_final(slides_originais[capa_slide_index])

    if cartazes:
        template_cartaz = slides_originais[cartaz_slide_indices[0]]
        for img_path in cartazes:
            new_slide = duplicate_slide(prs, template_cartaz)
            pic = [s for s in new_slide.shapes if s.shape_type == 13][0]
            left, top, w, h = pic.left, pic.top, pic.width, pic.height
            old_el = pic._element
            parent = old_el.getparent()
            pos = list(parent).index(old_el)
            parent.remove(old_el)
            new_pic = new_slide.shapes.add_picture(img_path, left, top, w, h)
            new_el = new_pic._element
            new_el.getparent().remove(new_el)
            parent.insert(pos, new_el)
            novos_cartaz_slides.append(new_slide)

    # --- reordenar: capa, novas matérias, novos cartazes, página final; remove antigos ---
    sld_id_lst = prs.slides._sldIdLst
    all_ids = list(sld_id_lst)

    capa_id = all_ids[capa_slide_index]
    indices_a_remover = set(old_materia_slide_indices)
    if cartazes:
        indices_a_remover |= set(cartaz_slide_indices)
    old_ids_a_remover = [all_ids[i] for i in indices_a_remover]
    resto_ids = [sid for i, sid in enumerate(all_ids)
                 if i != capa_slide_index and i not in indices_a_remover]

    n_novos = len(novas_materia_slides) + len(novos_cartaz_slides)
    novos_ids_todos = all_ids[len(all_ids) - n_novos:] if n_novos else []
    novos_materia_ids = novos_ids_todos[:len(novas_materia_slides)]
    novos_cartaz_ids = novos_ids_todos[len(novas_materia_slides):]
    resto_ids = [sid for sid in resto_ids if sid not in novos_ids_todos]



    for sid in all_ids:
        sld_id_lst.remove(sid)

    if cartazes:
        ordem_final = [capa_id] + novos_materia_ids + novos_cartaz_ids + resto_ids
    else:
        ordem_final = [capa_id] + novos_materia_ids + resto_ids
    for sid in ordem_final:
        sld_id_lst.append(sid)

    for sid in old_ids_a_remover:
        # remove definitivamente os slides antigos (matérias e cartazes) não usados nesta edição
        rId = sid.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        prs.part.drop_rel(rId)

    prs.save(output_path)
    _limpar_merged(fotos_dir)   # os intermediários já estão embutidos no .pptx
    return warnings


def _encontrar_soffice():
    """Procura o LibreOffice tanto no PATH quanto nos lugares onde ele é
    instalado no Mac e no Windows, porque o app quase nunca coloca o
    comando `soffice` no PATH sozinho — quem instala pelo instalador
    normal (ou pelo Homebrew cask) fica com o executável só dentro do
    pacote do aplicativo."""
    import shutil
    cmd = shutil.which("soffice") or shutil.which("libreoffice")
    if cmd:
        return cmd
    candidatos = [
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        _os.path.expanduser("~/Applications/LibreOffice.app/Contents/MacOS/soffice"),
        "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
        "C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe",
    ]
    for caminho in candidatos:
        if _os.path.isfile(caminho):
            return caminho
    return None


def _pdf_para_imagens(pdf_path, tmp_dir, dpi=100):
    """Converte cada página do PDF numa imagem, sem depender de nenhum
    programa externo: usa o PyMuPDF (pacote `pymupdf`, instalado via pip e
    que já traz o motor de renderização embutido). Se por algum motivo o
    PyMuPDF não estiver disponível, cai pro Poppler (`pdftoppm`) caso ele
    exista na máquina."""
    import subprocess
    import shutil

    try:
        import fitz  # PyMuPDF
        saidas = []
        with fitz.open(pdf_path) as doc:
            zoom = dpi / 72.0
            matriz = fitz.Matrix(zoom, zoom)
            for i, pagina in enumerate(doc):
                pix = pagina.get_pixmap(matrix=matriz)
                destino = _os.path.join(tmp_dir, f"pg-{i + 1:03d}.png")
                pix.save(destino)
                saidas.append(destino)
        return saidas

    except ImportError:
        pdftoppm_cmd = shutil.which("pdftoppm")
        if not pdftoppm_cmd:
            raise RuntimeError(
                "Falta uma forma de transformar o PDF em imagens. Instale o "
                "PyMuPDF com: pip3 install pymupdf"
            )
        subprocess.run(
            [pdftoppm_cmd, "-png", "-r", str(dpi), pdf_path,
             _os.path.join(tmp_dir, "pg")],
            check=True, capture_output=True,
        )
        return sorted(
            _os.path.join(tmp_dir, f) for f in _os.listdir(tmp_dir)
            if f.startswith("pg")
        )


def regenerate_final_page(pptx_path, soffice_script_path=None):
    """Renderiza todos os slides já gerados (exceto a própria página final) e
    monta uma grade de miniaturas na última página, substituindo o que
    estava lá antes (que pertencia a outra edição).

    Precisa do LibreOffice instalado (para renderizar o .pptx tal como ele
    aparece). A conversão de PDF pra imagem é feita pelo PyMuPDF, que vem
    junto via pip — não precisa de Poppler. Se o LibreOffice não for
    encontrado, a função avisa e não faz nada: o resto do arquivo (capa,
    matérias, cartazes) já foi gerado normalmente antes desse passo."""
    import subprocess
    import tempfile

    soffice_cmd = _encontrar_soffice()
    if not soffice_cmd:
        print("Aviso: página final não regenerada — o LibreOffice não foi "
              "encontrado. Instale com: brew install --cask libreoffice "
              "(Mac) e rode de novo. O restante do arquivo (capa, matérias, "
              "cartazes) foi gerado normalmente.")
        return

    prs = Presentation(pptx_path)
    slides = list(prs.slides)
    final_idx = len(slides) - 1
    n_thumbs = final_idx

    tmp_dir = tempfile.mkdtemp(prefix="jm_final_")
    # perfil próprio evita conflito com uma instância do LibreOffice já
    # aberta pelo usuário (aí o --headless não trava esperando a outra).
    perfil = _os.path.join(tmp_dir, "lo_profile")
    resultado = subprocess.run(
        [soffice_cmd, "--headless", "--norestore",
         f"-env:UserInstallation=file://{perfil}",
         "--convert-to", "pdf", "--outdir", tmp_dir, pptx_path],
        capture_output=True,
    )
    generated_pdf = os.path.join(tmp_dir, os.path.splitext(os.path.basename(pptx_path))[0] + ".pdf")
    if resultado.returncode != 0 or not os.path.isfile(generated_pdf):
        print("Aviso: página final não regenerada — o LibreOffice não "
              "conseguiu converter o arquivo. O restante do JM foi gerado "
              "normalmente.")
        return
    thumb_paths = _pdf_para_imagens(generated_pdf, tmp_dir, dpi=100)[:n_thumbs]

    final_slide = slides[final_idx]
    for shape in list(final_slide.shapes):
        if shape.shape_type == 13 or shape.shape_type == 6:  # picture ou group
            shape._element.getparent().remove(shape._element)

    slide_w = prs.slide_width
    slide_h = prs.slide_height
    margem = Emu(int(0.5 * 360000))
    area_w = slide_w - 2 * margem
    area_h = slide_h - 2 * margem

    n = len(thumb_paths)
    if n == 0:
        prs.save(pptx_path)
        return
    cols = math.ceil(math.sqrt(n * (area_w / area_h)))
    cols = max(1, min(cols, n))
    rows = math.ceil(n / cols)

    gap = Emu(int(0.3 * 360000))
    cell_w = (area_w - gap * (cols - 1)) // cols
    thumb_ratio = slide_w / slide_h
    cell_h = int(cell_w / thumb_ratio)
    if rows * cell_h + gap * (rows - 1) > area_h:
        cell_h = (area_h - gap * (rows - 1)) // rows
        cell_w = int(cell_h * thumb_ratio)

    grid_w = cell_w * cols + gap * (cols - 1)
    grid_h = cell_h * rows + gap * (rows - 1)
    start_x = (slide_w - grid_w) // 2
    start_y = (slide_h - grid_h) // 2

    for i, thumb in enumerate(thumb_paths):
        r, c = divmod(i, cols)
        left = start_x + c * (cell_w + gap)
        top = start_y + r * (cell_h + gap)
        final_slide.shapes.add_picture(thumb, left, top, cell_w, cell_h)

    prs.save(pptx_path)


def limpar_coautoria(pptx_path):
    """Remove do arquivo final as partes de rastreamento de coautoria/
    alterações do PowerPoint (authors.xml, revisionInfo.xml, changesInfos/).

    Por que isso é necessário: essas partes guardam o histórico de quem
    mexeu em cada slide, identificando os slides por um ID interno. Como o
    motor remove os slides originais do template e cria os novos, esses IDs
    deixam de existir — e aí o PowerPoint abre o arquivo, vê referências
    apontando pra slides que sumiram e acusa o arquivo como corrompido
    ('o PowerPoint encontrou um problema... deseja reparar?'). Num arquivo
    gerado do zero essas partes não têm utilidade nenhuma, então a saída
    mais limpa é simplesmente removê-las. Retorna quantas partes tirou."""
    import posixpath

    def eh_coautoria(nome):
        return (nome in ("ppt/authors.xml", "ppt/revisionInfo.xml")
                or nome.startswith("ppt/changesInfos/"))

    with zipfile.ZipFile(pptx_path) as z:
        nomes = z.namelist()
        conteudos = {n: z.read(n) for n in nomes}

    dropar = set(n for n in nomes if eh_coautoria(n))
    # inclui também os _rels dessas partes, se existirem
    for n in nomes:
        if n.endswith(".rels"):
            base = n.replace("_rels/", "").rsplit(".rels", 1)[0]
            if eh_coautoria(base):
                dropar.add(n)
    if not dropar:
        return 0

    # tira as <Relationship> que apontam pras partes removidas
    rels_name = "ppt/_rels/presentation.xml.rels"
    if rels_name in conteudos:
        rels = conteudos[rels_name].decode("utf-8")

        def rel_dropada(m):
            alvo = re.search(r'Target="([^"]+)"', m.group(0))
            if not alvo:
                return False
            full = posixpath.normpath(posixpath.join("ppt", alvo.group(1)))
            return full in dropar

        rels = re.sub(r"<Relationship\b[^>]*/>",
                      lambda m: "" if rel_dropada(m) else m.group(0), rels)
        conteudos[rels_name] = rels.encode("utf-8")

    # tira os <Override> do Content_Types das partes removidas
    ct_name = "[Content_Types].xml"
    if ct_name in conteudos:
        ct = conteudos[ct_name].decode("utf-8")

        def override_dropado(m):
            pn = re.search(r'PartName="([^"]+)"', m.group(0))
            return bool(pn) and pn.group(1).lstrip("/") in dropar

        ct = re.sub(r"<Override\b[^>]*/>",
                    lambda m: "" if override_dropado(m) else m.group(0), ct)
        conteudos[ct_name] = ct.encode("utf-8")

    tmp = pptx_path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for n in nomes:
            if n not in dropar:
                out.writestr(n, conteudos[n])
    os.replace(tmp, pptx_path)
    return len(dropar)
