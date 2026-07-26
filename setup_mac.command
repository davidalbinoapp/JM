#!/bin/bash
#
# Instalador do Jornal Mural (AGA) — Mac
# --------------------------------------
# Dois cliques neste arquivo e pronto. Ele confere e instala tudo que o
# sistema precisa: as bibliotecas do Python e o LibreOffice (usado só pra
# montar a página final de miniaturas). No fim, faz um autoteste e diz se
# está tudo certo.
#
# Se o Mac disser que "não pode abrir porque é de um desenvolvedor não
# identificado": clique com o botão direito no arquivo -> Abrir -> Abrir.

cd "$(dirname "$0")" || exit 1

# deixa o próprio arquivo executável (caso tenha vindo do GitHub sem isso)
chmod +x "$0" 2>/dev/null

VERDE="\033[0;32m"; VERMELHO="\033[0;31m"; AMARELO="\033[1;33m"; FIM="\033[0m"
ok()   { printf "${VERDE}✅ %s${FIM}\n" "$1"; }
erro() { printf "${VERMELHO}❌ %s${FIM}\n" "$1"; }
info() { printf "${AMARELO}→  %s${FIM}\n" "$1"; }
pausa() { echo ""; read -r -p "Pressione Enter para fechar esta janela." _; }

clear
echo "======================================================"
echo "     Instalador do Jornal Mural (AGA)  —  Mac"
echo "======================================================"
echo ""

# ---------------------------------------------------------------
# 1) Python
# ---------------------------------------------------------------
if ! command -v python3 >/dev/null 2>&1; then
  # tenta resolver sozinho pelo Homebrew, se ele existir (sem senha)
  if command -v brew >/dev/null 2>&1; then
    info "Python 3 não encontrado. Instalando pelo Homebrew..."
    brew install python && hash -r
  fi
fi
if ! command -v python3 >/dev/null 2>&1; then
  erro "Python 3 não encontrado."
  echo "   Vou abrir o site oficial. Baixe o instalador (botão amarelo),"
  echo "   instale, feche esta janela e rode este instalador de novo."
  open "https://www.python.org/downloads/"
  pausa
  exit 1
fi
ok "Python encontrado: $(python3 --version 2>&1)"

# ---------------------------------------------------------------
# 2) Bibliotecas do Python
# ---------------------------------------------------------------
echo ""
info "Instalando as bibliotecas do Python (pode demorar um pouco)..."

# garante que o pip existe (algumas instalações vêm sem ele)
python3 -m pip --version >/dev/null 2>&1 || python3 -m ensurepip --upgrade >/dev/null 2>&1
python3 -m pip install --upgrade pip >/dev/null 2>&1

# Instala tentando várias formas, da mais limpa pra mais tolerante — é assim
# que o instalador se REPARA sozinho nos erros mais comuns em Mac:
#  - "externally-managed-environment" (PEP 668) do Python novo / Homebrew
#  - falta de permissão no site-packages do sistema (resolve com --user)
# A 1ª tentativa mostra o erro na tela (pra ficar registrado); as demais
# são silenciosas e só entram se a anterior falhar.
instalar_libs() {
  python3 -m pip install -r requirements.txt && return 0
  info "Não deu na primeira — tentando reparar (modo do usuário)..."
  python3 -m pip install --user -r requirements.txt >/dev/null 2>&1 && return 0
  info "Tentando contornar o ambiente gerenciado do Python..."
  python3 -m pip install --break-system-packages -r requirements.txt >/dev/null 2>&1 && return 0
  python3 -m pip install --user --break-system-packages -r requirements.txt >/dev/null 2>&1 && return 0
  return 1
}

if instalar_libs; then
  ok "Bibliotecas do Python instaladas."
else
  echo ""
  erro "Não consegui instalar as bibliotecas do Python."
  echo "   Pra ver o erro detalhado, rode no Terminal:"
  echo "     python3 -m pip install -r requirements.txt"
  pausa
  exit 1
fi

# ---------------------------------------------------------------
# 2b) Enquadramento inteligente de fotos (OpenCV) — OPCIONAL
#     Detecta pessoas nas fotos e ajusta o corte pra não cortar
#     ninguém. É opcional: se não instalar, o JM usa o corte central.
# ---------------------------------------------------------------
echo ""
if python3 -c "import cv2" >/dev/null 2>&1; then
  ok "Enquadramento inteligente de fotos (OpenCV): já disponível."
else
  info "Instalando o enquadramento inteligente de fotos (opcional)..."
  if python3 -m pip install opencv-python-headless >/dev/null 2>&1 \
     || python3 -m pip install --user opencv-python-headless >/dev/null 2>&1 \
     || python3 -m pip install --break-system-packages opencv-python-headless >/dev/null 2>&1; then
    ok "Enquadramento inteligente instalado."
  else
    info "OpenCV não instalou — sem problema: as fotos usam o corte central."
  fi
fi

# ---------------------------------------------------------------
# 3) LibreOffice  (só pra montar a página final de miniaturas)
# ---------------------------------------------------------------
echo ""
if [ -d "/Applications/LibreOffice.app" ] || command -v soffice >/dev/null 2>&1; then
  ok "LibreOffice já está instalado."
else
  info "LibreOffice não encontrado. Vou instalar..."
  if command -v brew >/dev/null 2>&1; then
    echo "   (usando o Homebrew — sem precisar de senha)"
    if brew install --cask libreoffice; then
      ok "LibreOffice instalado."
    else
      erro "O Homebrew não conseguiu instalar. Vou tentar pelo site."
      open "https://www.libreoffice.org/download/download-libreoffice/"
    fi
  else
    echo "   Vou abrir a página oficial de download do LibreOffice."
    echo "   Passo a passo (é o jeito normal de instalar app no Mac):"
    echo "     1. Baixe o arquivo .dmg da página que vai abrir."
    echo "     2. Abra o .dmg baixado (dois cliques)."
    echo "     3. Arraste o ícone do LibreOffice pra dentro da pasta"
    echo "        Aplicativos (a seta aponta pra ela)."
    echo "     4. Volte aqui e rode este instalador de novo pra confirmar."
    open "https://www.libreoffice.org/download/download-libreoffice/"
  fi
fi

# ---------------------------------------------------------------
# 4) Autoteste final
# ---------------------------------------------------------------
echo ""
echo "------------------------------------------------------"
info "Conferindo se está tudo pronto..."
echo ""

TUDO_OK=1

python3 - <<'PY'
import importlib, sys
faltando = []
for mod, nome in [("pptx","python-pptx"), ("docx","python-docx"),
                  ("PIL","Pillow"), ("fitz","pymupdf")]:
    try:
        importlib.import_module(mod)
    except Exception:
        faltando.append(nome)
if faltando:
    print("MODS_FALTANDO:" + ",".join(faltando))
    sys.exit(1)
sys.exit(0)
PY
if [ $? -eq 0 ]; then
  ok "Bibliotecas do Python: OK"
else
  erro "Ainda faltam bibliotecas do Python."
  TUDO_OK=0
fi

if [ -d "/Applications/LibreOffice.app" ] || command -v soffice >/dev/null 2>&1; then
  ok "LibreOffice: OK (a página final de miniaturas vai sair sozinha)"
else
  erro "LibreOffice ainda não está instalado."
  echo "   Sem ele o JM é gerado normalmente, mas a última página (as"
  echo "   miniaturas) fica com a versão genérica do template."
  TUDO_OK=0
fi

# linha informativa (não trava o TUDO PRONTO — é um extra)
if python3 -c "import cv2" >/dev/null 2>&1; then
  ok "Enquadramento inteligente de fotos: ativo"
else
  info "Enquadramento inteligente de fotos: indisponível (usa corte central — ok)"
fi

echo ""
echo "======================================================"
if [ "$TUDO_OK" -eq 1 ]; then
  ok "TUDO PRONTO! Pode gerar o Jornal Mural."
else
  erro "Faltou alguma coisa (veja acima). Resolva e rode de novo."
fi
echo "======================================================"
pausa
