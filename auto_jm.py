"""
auto_jm.py — atalho pra montar o Jornal Mural passando SÓ a pasta da semana.

Em vez de escrever o comando inteiro à mão (caminho do briefing + um --qr por
QR, com o título exato de cada matéria), este script descobre tudo sozinho:

  - Acha o briefing (.docx) dentro da pasta — prefere o que tem "BRIEFING" no
    nome; se só houver um .docx, usa ele.
  - Acha os arquivos de QR (nome começa com "qr") e casa cada um com a matéria
    certa pelo TÍTULO que está no próprio nome do arquivo
    (ex.: "qr code Centro de Memória comemora 32 anos.png"), usando a mesma
    similaridade (difflib) que o motor usa pra casar foto.
  - Mostra o que detectou (briefing + QRs casados) e JÁ GERA o .pptx.

Uso:
    cd ~/Documents/Artes/AGA/2026/JM\\ SYSTEM/JM-main
    python3 auto_jm.py "../30-06"

Opcional:
    python3 auto_jm.py "../30-06" --dry-run     # só mostra, não gera
    python3 auto_jm.py "../30-06" saida.pptx    # força o nome do arquivo

Pegadinhas continuam valendo:
  - Toda matéria precisa da linha "Editoria:" no briefing (senão o motor pula).
  - Coloque TODAS as fotos e QRs na pasta ANTES de rodar.
  - Marcação de foto inteira / fundo desfocado continua vindo do briefing
    (escreva "(fundo desfocado)" na linha da Foto).
"""
import os
import sys
import glob
from datetime import date
from difflib import SequenceMatcher

from briefing_parser import parse_briefing
from run_jm import montar_jornal_mural


LIMIAR_QR = 0.45  # abaixo disso, não casa o QR (evita casar no título errado)


def _achar_briefing(pasta):
    """Retorna o caminho do .docx do briefing dentro da pasta, ou None."""
    docs = [d for d in glob.glob(os.path.join(pasta, "*.docx"))
            if not os.path.basename(d).startswith("~$")]  # ignora temporários do Word
    if not docs:
        return None
    if len(docs) == 1:
        return docs[0]
    # mais de um: prefere o que tem "briefing" no nome
    com_briefing = [d for d in docs if "briefing" in os.path.basename(d).lower()]
    if len(com_briefing) == 1:
        return com_briefing[0]
    return (com_briefing or docs)[0]


def _titulo_do_arquivo_qr(caminho):
    """Extrai o título da matéria do nome do arquivo de QR, tirando o prefixo
    'qr code' / 'qr' e a extensão. Ex.: 'qr code Centro de Memória.png' ->
    'Centro de Memória'."""
    nome = os.path.splitext(os.path.basename(caminho))[0]
    # tira prefixos comuns: "qr code", "qrcode", "qr-code", "qr_", "qr "
    baixo = nome.lower()
    for prefixo in ("qr code", "qrcode", "qr-code", "qr_code", "qr code_", "qr"):
        if baixo.startswith(prefixo):
            nome = nome[len(prefixo):]
            break
    return nome.strip(" -_")


def _achar_qrs(pasta):
    """Todos os arquivos de imagem cujo nome começa com 'qr'."""
    achados = []
    for caminho in glob.glob(os.path.join(pasta, "*")):
        base = os.path.basename(caminho).lower()
        if base.startswith("qr") and base.endswith((".png", ".jpg", ".jpeg")):
            achados.append(caminho)
    return sorted(achados)


def _ratio(a, b):
    return SequenceMatcher(None, a.lower().strip(), b.lower().strip()).ratio()


def _casar_qrs(pasta, titulos):
    """Casa cada arquivo de QR ao melhor título de matéria. Retorna
    (qr_codes dict {titulo: caminho}, avisos list)."""
    qr_codes = {}
    avisos = []
    usados = set()
    for caminho in _achar_qrs(pasta):
        alvo = _titulo_do_arquivo_qr(caminho)
        candidatos = [(t, _ratio(alvo, t)) for t in titulos if t not in usados]
        if not candidatos:
            avisos.append(f"QR '{os.path.basename(caminho)}': sem matéria livre pra casar.")
            continue
        melhor_titulo, melhor_score = max(candidatos, key=lambda x: x[1])
        if melhor_score < LIMIAR_QR:
            avisos.append(
                f"QR '{os.path.basename(caminho)}' não casou com nenhuma matéria "
                f"(melhor palpite: '{melhor_titulo}', {melhor_score:.0%}). "
                f"Passe manualmente com --qr se for necessário."
            )
            continue
        qr_codes[melhor_titulo] = caminho
        usados.add(melhor_titulo)
    return qr_codes, avisos


def main():
    argv = [a for a in sys.argv[1:]]
    dry_run = "--dry-run" in argv
    argv = [a for a in argv if a != "--dry-run"]

    if not argv:
        print("Uso: python3 auto_jm.py \"../PASTA_DA_SEMANA\" [saida.pptx] [--dry-run]")
        sys.exit(1)

    pasta = os.path.normpath(argv[0])
    output_forcado = argv[1] if len(argv) > 1 else None

    if not os.path.isdir(pasta):
        print(f"Erro: pasta não encontrada: {pasta}")
        sys.exit(1)

    briefing = _achar_briefing(pasta)
    if not briefing:
        print(f"Erro: nenhum .docx de briefing encontrado em {pasta}")
        sys.exit(1)

    paginas, _cartazes = parse_briefing(briefing)
    titulos = [m["titulo"] for p in paginas for m in p["materias"] if m.get("titulo")]

    qr_codes, avisos_qr = _casar_qrs(pasta, titulos)

    print("=" * 60)
    print("AUTO JM — o que foi detectado")
    print("=" * 60)
    print(f"Pasta:    {pasta}")
    print(f"Briefing: {os.path.basename(briefing)}")
    print(f"Matérias: {len(titulos)}")
    for t in titulos:
        marca = "  [QR]" if t in qr_codes else ""
        print(f"   • {t}{marca}")
    if qr_codes:
        print(f"\nQR codes casados ({len(qr_codes)}):")
        for titulo, caminho in qr_codes.items():
            print(f"   • {titulo}")
            print(f"       ↳ {os.path.basename(caminho)}")
    else:
        print("\nNenhum QR code detectado na pasta.")
    for a in avisos_qr:
        print(f"   ! {a}")

    # nome de saída (mesma regra do run_jm)
    if output_forcado:
        output_path = output_forcado
    else:
        base = os.path.basename(os.path.normpath(pasta))
        data_hoje = date.today().strftime("%d%m%y")
        output_path = os.path.join(pasta, f"AGA_{base}-{data_hoje}.pptx")
    print(f"\nSaída:    {output_path}")
    print("=" * 60)

    if dry_run:
        print("\n[--dry-run] Nada foi gerado. Rode sem --dry-run pra montar o JM.")
        return

    print("\nMontando o Jornal Mural...\n")
    avisos = montar_jornal_mural(
        briefing_docx=briefing,
        fotos_dir=pasta,
        output_path=output_path,
        qr_codes=qr_codes,
    )

    print(f"\nPronto! Arquivo gerado em: {output_path}")
    todos = avisos_qr + avisos
    if todos:
        print(f"\n{len(todos)} aviso(s):")
        for a in todos:
            print(" -", a)
    else:
        print("Nenhum aviso — tudo casou certinho.")


if __name__ == "__main__":
    main()
