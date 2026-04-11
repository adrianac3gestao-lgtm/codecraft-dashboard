@echo off
setlocal

set PASTA=C:\Users\adria\OneDrive\C3 Gestao\EquipeC3\CODECRAFT\7. Dashboard
set LOG=%PASTA%\log_atualizacao.txt

cd /d "%PASTA%"

echo ========================================= >> "%LOG%"
echo Inicio: %date% %time% >> "%LOG%"

REM --- Rodar Python ---
python gerar_dados_dashboard.py < nul >> "%LOG%" 2>&1
if errorlevel 1 (
    echo ERRO no Python >> "%LOG%"
    exit /b 1
)

REM --- Enviar para GitHub com force push para evitar conflitos ---
git add index.html log_atualizacao.txt >> "%LOG%" 2>&1
git commit -m "atualizacao %date%" >> "%LOG%" 2>&1
git push --force origin main >> "%LOG%" 2>&1
if errorlevel 1 (
    echo ERRO no Git push >> "%LOG%"
    exit /b 1
)

echo Sucesso! Dashboard atualizado. >> "%LOG%"
echo ========================================= >> "%LOG%"

endlocal
