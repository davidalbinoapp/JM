#!/bin/bash
# ---------------------------------------------------------------------------
#  Criar atalho (Mac) — gera o "Núcleo Inteligente JM.app": um aplicativo
#  bonito, com ícone da marca, pra rodar o Jornal Mural com um clique.
#
#  Rode UMA VEZ por Mac (dois cliques neste arquivo). Ele cria o "Núcleo
#  Inteligente JM.app" na pasta "JM SYSTEM" (ao lado da JM-main), com cara de app.
#  O ícone fica DENTRO do app, então viaja junto se você copiar o app pra
#  outra máquina. Em um Mac novo (baixado do GitHub), é só rodar isto de novo.
#
#  (Na 1ª vez o Mac pode pedir botão direito -> Abrir -> Abrir.)
# ---------------------------------------------------------------------------
set -e
# --auto = chamado pelo instalador (silencioso, sem pausa no fim).
AUTO=0; [ "$1" = "--auto" ] && AUTO=1
DIR="$(cd "$(dirname "$0")" && pwd)"      # a JM-main (onde este arquivo mora)
BASE="$(cd "$DIR/.." && pwd)"             # a JM SYSTEM (um nível acima)
APP="$BASE/Núcleo Inteligente JM.app"
ICON_PNG="$DIR/icone_montar_jm.png"
WORK="$(mktemp -d)"

if [ ! -f "$ICON_PNG" ]; then
    echo "Erro: não achei o ícone ($ICON_PNG). Ele deveria estar na JM-main."
    [ "$AUTO" = "1" ] || read -n 1 -s -r -p "Aperte qualquer tecla pra fechar..."
    exit 1
fi

echo "Criando o Núcleo Inteligente JM.app em: $BASE"

# 1) o miolo do app (AppleScript) -> compila num .app de verdade
cat > "$WORK/montar.applescript" <<'APPLESCRIPT'
on run
    set cmdPath to my acharLauncher()
    if cmdPath is "" then
        display dialog "Não encontrei a pasta JM-main (o motor). Aponte-a na próxima janela." buttons {"OK"} default button 1 with icon note
        set jm to POSIX path of (choose folder with prompt "Onde está a pasta JM-main?")
        set cmdPath to jm & "Núcleo Inteligente JM.command"
    end if
    do shell script "open " & quoted form of cmdPath
end run

on acharLauncher()
    set appPath to POSIX path of (path to me)
    set appParent to do shell script "dirname " & quoted form of appPath
    set lar to POSIX path of (path to home folder)
    set cands to {appParent & "/JM-main/Núcleo Inteligente JM.command", appParent & "/JM SYSTEM/JM-main/Núcleo Inteligente JM.command", lar & "Documents/Artes/AGA/2026/JM SYSTEM/JM-main/Núcleo Inteligente JM.command"}
    repeat with c in cands
        set c to c as text
        try
            do shell script "test -f " & quoted form of c
            return c
        end try
    end repeat
    return ""
end acharLauncher
APPLESCRIPT

rm -rf "$APP"
osacompile -o "$APP" "$WORK/montar.applescript"

# 2) ícone da marca: PNG -> iconset -> applet.icns (troca o ícone padrão)
ICONSET="$WORK/icon.iconset"; mkdir -p "$ICONSET"
for s in 16 32 128 256 512; do
    sips -z "$s" "$s" "$ICON_PNG" --out "$ICONSET/icon_${s}x${s}.png" >/dev/null
    d=$((s*2))
    sips -z "$d" "$d" "$ICON_PNG" --out "$ICONSET/icon_${s}x${s}@2x.png" >/dev/null
done
iconutil -c icns "$ICONSET" -o "$APP/Contents/Resources/applet.icns"

# 3) faz o Finder atualizar o ícone na hora
touch "$APP"
/System/Library/Frameworks/CoreServices.framework/Frameworks/LaunchServices.framework/Support/lsregister -f "$APP" 2>/dev/null || true

rm -rf "$WORK"
echo
echo "Pronto! Criado: $APP"
echo "Arraste o 'Núcleo Inteligente JM.app' pro Dock ou pra Área de Trabalho e use com 1 clique."
echo
[ "$AUTO" = "1" ] || read -n 1 -s -r -p "Aperte qualquer tecla pra fechar..."
[ "$AUTO" = "1" ] || echo
