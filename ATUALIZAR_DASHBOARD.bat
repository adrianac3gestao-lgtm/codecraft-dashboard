@echo off
setlocal

REM --- Caminhos ---
set PASTA=C:\Users\adria\OneDrive\C3 Gestao\EquipeC3\CODECRAFT\7. Dashboard
set LOG=%PASTA%\log_atualizacao.txt

REM --- Entrar na pasta ---
cd /d "%PASTA%"

REM --- Registrar inicio no log ---
echo ========================================= >> "%LOG%"
echo Inicio: %date% %time% >> "%LOG%"

REM --- Rodar Python silenciosamente ---
python gerar_dados_dashboard.py >> "%LOG%" 2>&1
if errorlevel 1 (
    echo ERRO no Python >> "%LOG%"
    exit /b 1
)

REM --- Enviar para GitHub silenciosamente ---
git add . >> "%LOG%" 2>&1
git commit -m "atualizacao automatica %date%" >> "%LOG%" 2>&1
git push >> "%LOG%" 2>&1
if errorlevel 1 (
    echo ERRO no Git push >> "%LOG%"
    exit /b 1
)

REM --- Registrar sucesso ---
echo Sucesso! Dashboard atualizado. >> "%LOG%"
echo ========================================= >> "%LOG%"

endlocal
