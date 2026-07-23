"""
run_jm.py — roda o Jornal Mural inteiro de uma vez: lê o briefing, casa as
fotos, monta capa + matérias + cartazes, e regenera a página final.

Uso básico:
    from run_jm import montar_jornal_mural
    avisos = montar_jornal_mural(
        briefing_docx="caminho/do/briefing.docx",
        fotos_dir="pasta/com/as/fotos/da/semana",
        output_path="JM_pronto.pptx",
        qr_codes={"Título exato da matéria": "caminho/do/qr.png"},  # opcional
    )
    for a in avisos:
        print(a)

O que esse script faz sozinho, sem precisar reescrever nada a cada semana:
  - Lê o Word e identifica capa, matérias (única/dupla) e cartazes.
  - Casa cada nome de foto do Word com o arquivo real na pasta de fotos
    (tolera pequenas diferenças de acentuação/espaço no nome do arquivo).
  - Reconhece "UNIR ESSA: ... e ESSA: ..." e junta as fotos automaticamente.
  - Identifica quantos cartazes existem e ajusta o número de slides.
  - Aplica toda a diagramação automática: alinhamento de tag, prevenção de
    viúva, ajuste de fonte pra não sobrepor texto, corte diagonal da foto,
    espaço pro QR code, logo de fechamento só na última matéria.
  - Regenera a página final com as miniaturas certas dessa edição.

O que ainda precisa de intervenção humana antes de rodar:
  - Baixar as fotos da semana (do SharePoint) pra uma pasta local.
  - Se alguma matéria tiver QR code, informar no parâmetro `qr_codes`
    (dict: título exato da matéria -> caminho do arquivo do QR).
  - Conferir o resultado antes de publicar — o motor erra menos, mas a
    palavra final é sempre de quem revisa.
"""
import os
from briefing_parser import parse_briefing
from jm_engine import build_deck, regenerate_final_page, match_photo_file


def montar_jornal_mural(briefing_docx, fotos_dir, output_path,
                         template_path=None, qr_codes=None):
    if template_path is None:
        template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                      "template_compactado.pptx")
    qr_codes = qr_codes or {}

    paginas, cartazes_nomes = parse_briefing(briefing_docx)

    avisos = []
    for pagina in paginas:
        for materia in pagina["materias"]:
            if materia["titulo"] in qr_codes:
                materia["qr_path"] = qr_codes[materia["titulo"]]

    cartazes_paths = []
    for nome in cartazes_nomes:
        caminho = match_photo_file(nome, fotos_dir)
        if caminho:
            cartazes_paths.append(caminho)
        else:
            avisos.append(f"Cartaz '{nome}' não encontrado na pasta de fotos.")

    avisos_build = build_deck(
        template_path=template_path,
        paginas=paginas,
        fotos_dir=fotos_dir,
        output_path=output_path,
        cartazes=cartazes_paths if cartazes_paths else None,
    )
    avisos.extend(avisos_build)

    regenerate_final_page(output_path)

    titulos_com_qr_pendente = [
        m["titulo"] for p in paginas for m in p["materias"]
        if "qr" in (m.get("corpo") or "").lower() and m["titulo"] not in qr_codes
    ]
    for titulo in titulos_com_qr_pendente:
        avisos.append(f"'{titulo}' menciona QR code no texto, mas nenhum QR foi informado pra essa matéria.")

    return avisos


if __name__ == "__main__":
    import argparse
    import json as _json

    parser = argparse.ArgumentParser(
        description="Monta o Jornal Mural a partir do briefing e das fotos da semana."
    )
    parser.add_argument("briefing_docx", help="Caminho do arquivo .docx do briefing")
    parser.add_argument("fotos_dir", help="Pasta com as fotos da semana (e os cartazes, se estiverem juntos)")
    parser.add_argument("output_path", help="Nome do arquivo .pptx que será gerado")
    parser.add_argument(
        "--qr", action="append", default=[], metavar="TITULO=CAMINHO",
        help="QR code de uma matéria específica. Repita --qr pra cada QR. "
             'Exemplo: --qr "Relatórios do ano de 2025=./fotos/qr_relatorio.png"'
    )
    args = parser.parse_args()

    qr_codes = {}
    for item in args.qr:
        if "=" not in item:
            print(f"Aviso: --qr '{item}' ignorado (formato esperado: TITULO=CAMINHO)")
            continue
        titulo, caminho = item.split("=", 1)
        qr_codes[titulo.strip()] = caminho.strip()

    avisos = montar_jornal_mural(
        briefing_docx=args.briefing_docx,
        fotos_dir=args.fotos_dir,
        output_path=args.output_path,
        qr_codes=qr_codes,
    )

    print(f"\nPronto! Arquivo gerado em: {args.output_path}")
    if avisos:
        print(f"\n{len(avisos)} aviso(s):")
        for a in avisos:
            print(" -", a)
    else:
        print("Nenhum aviso — tudo casou certinho.")
