#!/bin/bash
# ---------------------------------------------------------------------------
# Montar JM — roda o Jornal Mural com 2 cliques, SEM digitar caminho nenhum.
#
# Como usar:
#   1) Coloque TODAS as fotos e QR da semana dentro da pasta da semana.
#   2) Dê dois cliques neste arquivo ("Montar JM.command").
#   3) Na janela que abrir, clique na PASTA DA SEMANA e confirme.
#   O motor monta o .pptx dentro da própria pasta e mostra o resultado.
#
# (Na primeira vez, o Mac pode avisar "desenvolvedor não identificado":
#  clique com o botão direito no arquivo -> Abrir -> Abrir.)
# ---------------------------------------------------------------------------

# pasta onde este arquivo mora = a JM-main (o motor). Assim o comando acha
# o auto_jm.py, o template, as tags e as fontes, não importa de onde é aberto.
DIR="$(cd "$(dirname "$0")" && pwd)"
# um nível acima = onde costumam ficar as pastas de semana (ponto de partida
# da janela; você pode navegar pra qualquer lugar a partir dela).
BASE="$(cd "$DIR/.." && pwd)"

# janela nativa do Finder pra escolher a pasta da semana (zero digitação).
PASTA=$(osascript <<EOF 2>/dev/null
try
    set p to choose folder with prompt "Escolha a PASTA DA SEMANA do JM (com as fotos e QR dentro):" default location POSIX file "$BASE"
    return POSIX path of p
end try
EOF
)

if [ -z "$PASTA" ]; then
    echo "Cancelado — nenhuma pasta foi escolhida. Nada foi gerado."
    echo
    read -n 1 -s -r -p "Aperte qualquer tecla pra fechar..."
    exit 0
fi

cd "$DIR" || { echo "Erro: não achei a pasta do motor."; read -n 1 -s -r; exit 1; }

echo "============================================================"
echo "Montando o Jornal Mural da pasta:"
echo "   $PASTA"
echo "============================================================"
echo

python3 auto_jm.py "$PASTA"

echo
read -n 1 -s -r -p "Pronto. Aperte qualquer tecla pra fechar esta janela..."
echo
