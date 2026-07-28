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
call :ACHA_PYTHON

if not defined PYCMD (
  echo [X] Python nao encontrado. Vou tentar instalar...
  where winget >nul 2>&1
  if !errorlevel!==0 (
    winget install -e --id Python.Python.3.12 --accept-source-agreements --accept-package-agreements
    REM tenta usar o Python recem-instalado JA NESTA execucao (sem precisar
    REM fechar e abrir de novo): procura o launcher/python no PATH e nas
    REM pastas onde o winget costuma instalar.
    call :ACHA_PYTHON
    if not defined PYCMD (
      echo.
      echo Python instalado. Feche esta janela e rode o instalador de novo
      echo para ele ser reconhecido.
      pause
      exit /b 1
    )
    echo [OK] Python recem-instalado localizado nesta execucao.
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
REM garante que o pip existe (algumas instalacoes vem sem ele)
%PYCMD% -m pip --version >nul 2>&1 || %PYCMD% -m ensurepip --upgrade >nul 2>&1
%PYCMD% -m pip install --upgrade pip >nul 2>&1
%PYCMD% -m pip install -r requirements.txt
if !errorlevel! neq 0 (
  echo    Primeira tentativa falhou - tentando reparar ^(modo do usuario^)...
  %PYCMD% -m pip install --user -r requirements.txt
)
REM o autoteste no fim confirma se as bibliotecas ficaram mesmo instaladas;
REM por isso aqui a gente nao aborta - segue e deixa o autoteste decidir.
echo [OK] Etapa das bibliotecas concluida.

REM ----------------------------------------------------------------
REM  2b) Enquadramento inteligente de fotos (OpenCV) - OPCIONAL
REM      Detecta pessoas nas fotos e ajusta o corte pra nao cortar
REM      ninguem. Se nao instalar, o JM usa o corte central.
REM ----------------------------------------------------------------
echo.
%PYCMD% -c "import cv2" >nul 2>&1
if !errorlevel!==0 (
  echo [OK] Enquadramento inteligente de fotos ^(OpenCV^): ja disponivel.
) else (
  echo --^> Instalando o enquadramento inteligente de fotos ^(opcional^)...
  %PYCMD% -m pip install opencv-python-headless
  if !errorlevel! neq 0 %PYCMD% -m pip install --user opencv-python-headless
)

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

REM linha informativa (nao trava o TUDO PRONTO - e um extra)
%PYCMD% -c "import cv2" >nul 2>&1
if !errorlevel!==0 (
  echo [OK] Enquadramento inteligente de fotos: ativo
) else (
  echo --^> Enquadramento inteligente de fotos: indisponivel ^(usa corte central - ok^)
)

REM ----------------------------------------------------------------
REM  5) Atalho "Montar JM" na Area de Trabalho (icone da marca) -
REM     nasce sozinho aqui, sem nenhum passo extra depois de instalar.
REM     Aponta pro "Montar JM (Windows).bat" e usa o icone .ico.
REM ----------------------------------------------------------------
echo.
powershell -NoProfile -Command "$s=New-Object -ComObject WScript.Shell; $l=$s.CreateShortcut((Join-Path ([Environment]::GetFolderPath('Desktop')) 'Montar JM.lnk')); $l.TargetPath='%~dp0Montar JM (Windows).bat'; $l.WorkingDirectory='%~dp0'; $l.IconLocation='%~dp0icone_montar_jm.ico'; $l.Description='Montar o Jornal Mural'; $l.Save()" >nul 2>&1
if !errorlevel!==0 (
  echo [OK] Atalho "Montar JM" criado na Area de Trabalho.
) else (
  echo --^> Nao consegui criar o atalho na Area de Trabalho ^(da pra criar depois^).
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
exit /b 0

REM ================== subrotinas ==================
:ACHA_PYTHON
REM Define PYCMD com o Python disponivel, sem sair da execucao atual.
REM Ordem: launcher "py", "python" no PATH, e as pastas onde o winget
REM instala (por usuario e para todos os usuarios). Se nada for achado,
REM PYCMD fica indefinido e o chamador decide o que fazer.
if defined PYCMD goto :eof
where py     >nul 2>&1 && ( set "PYCMD=py"     & goto :eof )
where python >nul 2>&1 && ( set "PYCMD=python" & goto :eof )
for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python3*") do (
  if exist "%%D\python.exe" ( set PYCMD="%%D\python.exe" & goto :eof )
)
for /d %%D in ("%ProgramFiles%\Python3*") do (
  if exist "%%D\python.exe" ( set PYCMD="%%D\python.exe" & goto :eof )
)
goto :eof
