@echo off
REM ============================================================
REM  Instalador do Jornal Mural (AGA) - Windows
REM  ------------------------------------------------------------
REM  Dois cliques neste arquivo. Ele confere e instala tudo que o
REM  sistema precisa: as bibliotecas do Python e o LibreOffice
REM  (usado so pra montar a pagina final de miniaturas). No fim,
REM  faz um autoteste e diz se esta tudo certo.
REM ============================================================
chcp 65001 >nul
setlocal enabledelayedexpansion
cd /d "%~dp0"

cls
echo ======================================================
echo      Instalador do Jornal Mural (AGA)  -  Windows
echo ======================================================
echo.

REM ----------------------------------------------------------------
REM  1) Python
REM ----------------------------------------------------------------
set "PYCMD="
where py    >nul 2>&1 && set "PYCMD=py"
if not defined PYCMD (
  where python >nul 2>&1 && set "PYCMD=python"
)

if not defined PYCMD (
  echo [X] Python nao encontrado. Vou tentar instalar...
  where winget >nul 2>&1
  if !errorlevel!==0 (
    winget install -e --id Python.Python.3.12 --accept-source-agreements --accept-package-agreements
    echo.
    echo Feche esta janela e rode o instalador de novo para o Python ser reconhecido.
    pause
    exit /b 1
  ) else (
    echo Abrindo o site oficial do Python. Baixe, instale MARCANDO
    echo a opcao "Add Python to PATH", feche esta janela e rode de novo.
    start "" "https://www.python.org/downloads/"
    pause
    exit /b 1
  )
)
echo [OK] Python encontrado.
%PYCMD% --version

REM ----------------------------------------------------------------
REM  2) Bibliotecas do Python
REM ----------------------------------------------------------------
echo.
echo --^> Instalando as bibliotecas do Python (pode demorar um pouco)...
%PYCMD% -m pip install --upgrade pip >nul 2>&1
%PYCMD% -m pip install -r requirements.txt
if !errorlevel! neq 0 (
  echo.
  echo [X] Nao consegui instalar as bibliotecas do Python. Veja o erro acima.
  pause
  exit /b 1
)
echo [OK] Bibliotecas do Python instaladas.

REM ----------------------------------------------------------------
REM  3) LibreOffice
REM ----------------------------------------------------------------
echo.
set "LO="
if exist "%ProgramFiles%\LibreOffice\program\soffice.exe" set "LO=1"
if exist "%ProgramFiles(x86)%\LibreOffice\program\soffice.exe" set "LO=1"

if defined LO (
  echo [OK] LibreOffice ja esta instalado.
) else (
  echo --^> LibreOffice nao encontrado. Vou instalar...
  where winget >nul 2>&1
  if !errorlevel!==0 (
    winget install -e --id TheDocumentFoundation.LibreOffice --accept-source-agreements --accept-package-agreements
  ) else (
    echo Abrindo a pagina oficial do LibreOffice. Baixe o instalador,
    echo rode ele, e depois rode este instalador de novo para confirmar.
    start "" "https://www.libreoffice.org/download/download-libreoffice/"
  )
)

REM ----------------------------------------------------------------
REM  4) Autoteste final
REM ----------------------------------------------------------------
echo.
echo ------------------------------------------------------
echo --^> Conferindo se esta tudo pronto...
echo.

set "TUDO_OK=1"

%PYCMD% -c "import importlib.util,sys; f=[n for m,n in [('pptx','python-pptx'),('docx','python-docx'),('PIL','Pillow'),('fitz','pymupdf')] if importlib.util.find_spec(m) is None]; sys.exit(1 if f else 0)"
if !errorlevel!==0 (
  echo [OK] Bibliotecas do Python: OK
) else (
  echo [X] Ainda faltam bibliotecas do Python.
  set "TUDO_OK=0"
)

set "LO="
if exist "%ProgramFiles%\LibreOffice\program\soffice.exe" set "LO=1"
if exist "%ProgramFiles(x86)%\LibreOffice\program\soffice.exe" set "LO=1"
if defined LO (
  echo [OK] LibreOffice: OK ^(a pagina final de miniaturas vai sair sozinha^)
) else (
  echo [X] LibreOffice ainda nao esta instalado.
  echo     Sem ele o JM e gerado normalmente, mas a ultima pagina
  echo     ^(as miniaturas^) fica com a versao generica do template.
  set "TUDO_OK=0"
)

echo.
echo ======================================================
if "!TUDO_OK!"=="1" (
  echo [OK] TUDO PRONTO! Pode gerar o Jornal Mural.
) else (
  echo [X] Faltou alguma coisa ^(veja acima^). Resolva e rode de novo.
)
echo ======================================================
echo.
pause
