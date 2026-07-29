@echo off
chcp 65001 >nul
setlocal enableextensions
rem ===========================================================================
rem  Nucleo Inteligente JM (Windows) — roda o Jornal Mural com 2 cliques, SEM
rem  digitar caminho nenhum. Equivalente ao "Nucleo Inteligente JM.command" do Mac.
rem
rem  Como usar:
rem    1) Coloque TODAS as fotos e QR da semana dentro da pasta da semana.
rem    2) De dois cliques neste arquivo ("Nucleo Inteligente JM (Windows).bat").
rem    3) Na janela que abrir, clique na PASTA DA SEMANA e confirme.
rem    O motor monta o .pptx dentro da propria pasta e mostra o resultado.
rem
rem  (Se o Windows mostrar um aviso azul "O Windows protegeu o computador":
rem   clique em "Mais informacoes" -> "Executar assim mesmo".)
rem ===========================================================================

rem pasta onde este .bat mora = a JM-main (o motor). Assim ele acha o auto_jm.py,
rem o template, as tags e as fontes, nao importa de onde foi aberto.
set "DIR=%~dp0"
if "%DIR:~-1%"=="\" set "DIR=%DIR:~0,-1%"

rem janela nativa do Windows pra escolher a pasta da semana (zero digitacao).
rem O comando PowerShell fica numa LINHA SO de proposito (continuacao com ^ em
rem .bat e fragil). O %DIR% e trocado pelo cmd antes do PowerShell rodar.
set "PASTA="
for /f "usebackq delims=" %%p in (`powershell -NoProfile -STA -Command "Add-Type -AssemblyName System.Windows.Forms | Out-Null; $f = New-Object System.Windows.Forms.FolderBrowserDialog; $f.Description = 'Escolha a PASTA DA SEMANA do JM (com as fotos e QR dentro)'; $f.SelectedPath = (Split-Path -Parent '%DIR%'); if ($f.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { Write-Output $f.SelectedPath }"`) do set "PASTA=%%p"

if not defined PASTA (
  echo Cancelado — nenhuma pasta foi escolhida. Nada foi gerado.
  echo.
  pause
  exit /b 0
)

cd /d "%DIR%"

echo ============================================================
echo Montando o Jornal Mural da pasta:
echo    %PASTA%
echo ============================================================
echo.

rem acha o Python: primeiro o launcher "py" (padrao no Windows), senao "python".
where py >nul 2>nul
if %errorlevel%==0 (
  py auto_jm.py "%PASTA%"
) else (
  python auto_jm.py "%PASTA%"
)

echo.
echo Pronto.
pause
