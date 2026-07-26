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
        linhas_neste_paragrafo = 1
        for palavra in palavras[1:]:
            teste = linha_atual + " " + palavra
            bbox = font.getbbox(teste)
            if (bbox[2] - bbox[0]) > largura_px:
                linhas_neste_paragrafo += 1
                linha_atual = palavra
            else:
                linha_atual = teste
        total_linhas += linhas_neste_paragrafo
    return max(total_linhas, 1)


def _evitar_viuva(texto):
    """Troca o último espaço por espaço inseparável, pra a última palavra do
    título nunca ficar sozinha numa linha (viúva tipográfica)."""
    partes = texto.rsplit(" ", 1)
    if len(partes) == 2:
        return partes[0] + "\u00a0" + partes[1]
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

# Quando a foto é mais alta que a caixa e precisa ser cortada em cima/embaixo,
# esse viés diz quanto do excedente sai do TOPO (o resto sai da base). 0.5 =
# corte centralizado (decepa a cabeça de quem está em pé); 0.2 = puxa pro
# topo, preservando rosto/cabeça e sacrificando mais o rodapé da foto.
VIES_CORTE_VERTICAL = 0.2


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
                     foto_inteira=False):
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
        if src_ratio > target_ratio:
            new_w = int(h * target_ratio)
            x0 = (w - new_w) // 2
            im = im.crop((x0, 0, x0 + new_w, h))
        else:
            new_h = int(w / target_ratio)
            y0 = int((h - new_h) * VIES_CORTE_VERTICAL)  # puxa o corte pro topo
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


def _compose_merged_photo(photo_paths, target_ratio):
    """Junta 2 ou mais fotos lado a lado (com um filete branco de separação),
    cada uma enquadrada priorizando seu próprio assunto (center-crop por
    enquanto — ajuste fino manual quando o assunto não estiver centralizado),
    formando UMA imagem só na proporção da caixa. O corte diagonal do canto
    é aplicado depois, sobre o conjunto inteiro, como em qualquer foto normal."""
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
            x0 = (w - new_w) // 2
            im = im.crop((x0, 0, x0 + new_w, h))
        else:
            new_h = int(w / sub_ratio)
            y0 = int((h - new_h) * VIES_CORTE_VERTICAL)  # puxa o corte pro topo
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


def resolve_foto(foto_arquivo, fotos_dir, target_ratio, warnings, contexto=""):
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
        composta = _compose_merged_photo(paths, target_ratio)
        tmp_path = os.path.join(fotos_dir, f"_merged_{abs(hash(tuple(foto_arquivo)))}.jpg")
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
        titulo_final = _evitar_viuva(materia["titulo"])
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
        # borda esquerda da foto, menos uma folga, como limite pra tag encolher
        limite_tag = None
        if photo_box is not None and photo_box.left > tag_box.left:
            limite_tag = photo_box.left - Emu(int(0.3 * 360000))
        _replace_tag(tag_box, materia["editoria"], limite_direito_emu=limite_tag)
    else:
        warnings.append("CAPA: não encontrei a tag de editoria pra substituir.")

    if photo_box:
        ratio = photo_box.width / photo_box.height
        foto_path = resolve_foto(materia.get("foto_arquivo"), fotos_dir, ratio, warnings, "CAPA")
        if foto_path:
            _replace_picture(photo_box, foto_path,
                             foto_inteira=materia.get("foto_inteira", False))
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

        titulo_final = _evitar_viuva(story["titulo"])
        _set_run_text_keep_format(title_box.text_frame, titulo_final)
        _set_run_text_keep_format(body_box.text_frame, story["corpo"])
        if story.get("qr_path"):
            _ajustar_titulo_largura_reduzida(title_box, titulo_final)
        _reposicionar_corpo_apos_titulo(title_box, body_box, story["titulo"])
        _ajustar_corpo_sem_transbordar(body_box, story["corpo"], limite_inferior)
        _replace_tag(tag_box, story["editoria"])
        foto_path = resolve_foto(story.get("foto_arquivo"), fotos_dir,
                                  photo_box.width / photo_box.height, warnings,
                                  f"Matéria '{story['titulo']}'")
        if foto_path:
            _replace_picture(photo_box, foto_path,
                             foto_inteira=story.get("foto_inteira", False))
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
    tem_boxes_extras = len(textboxes) > 2
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
        titulo_final = _evitar_viuva(materia["titulo"])
        _set_run_text_keep_format(title_box.text_frame, titulo_final)
    else:
        warnings.append(f"Página única '{materia['titulo']}': não encontrei a caixa de título.")
    if body_box:
        _set_run_text_keep_format(body_box.text_frame, materia["corpo"])
    else:
        warnings.append(f"Página única '{materia['titulo']}': não encontrei a caixa de corpo.")
    if title_box and body_box:
        _reposicionar_corpo_apos_titulo(title_box, body_box, materia["titulo"])
        limite_inferior = slide.part.package.presentation_part.presentation.slide_height - Emu(int(1.0 * 360000))
        _ajustar_corpo_sem_transbordar(body_box, materia["corpo"], limite_inferior)
    if photo_box:
        foto_path = resolve_foto(materia.get("foto_arquivo"), fotos_dir,
                                  photo_box.width / photo_box.height, warnings,
                                  f"Página única '{materia['titulo']}'")
        if foto_path:
            _replace_picture(photo_box, foto_path,
                             foto_inteira=materia.get("foto_inteira", False))
    else:
        warnings.append(f"Página única '{materia['titulo']}': não encontrei a foto principal.")

    if len(textboxes) > 2:
        for extra_tb in textboxes[2:]:
            extra_tb._element.getparent().remove(extra_tb._element)
        for line_shape in [s for s in slide.shapes if s.shape_type == 9]:
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

def build_deck(template_path, paginas, fotos_dir, output_path, media_dir=None,
               capa_slide_index=0, materia_dupla_template_index=1,
               materia_unica_template_index=4,
               old_materia_slide_indices=(1, 2, 3, 4),
               cartazes=None, cartaz_slide_indices=(5, 6, 7, 8)):
    warnings = []
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
