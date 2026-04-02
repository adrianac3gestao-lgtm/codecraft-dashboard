@echo off
title Atualizando Dashboard CodeCraft...
color 0A

echo.
echo  ============================================
echo   DASHBOARD CODECRAFT — ATUALIZACAO DIARIA
echo  ============================================
echo.

REM --- Caminho da pasta do dashboard ---
set PASTA=C:\Users\adria\OneDrive\C3 Gestao\EquipeC3\CODECRAFT\7. Dashboard
set EXCEL=C:\Users\adria\OneDrive\C3 Gestao\EquipeC3\CODECRAFT\Relatorio Gerencial_Codecraft2026.xlsx
set PYTHON=python

REM --- Entrar na pasta ---
cd /d "%PASTA%"
if errorlevel 1 (
    echo ERRO: Pasta nao encontrada: %PASTA%
    pause
    exit /b 1
)

echo [1/3] Atualizando dados do Excel...
echo.
%PYTHON% gerar_dados_dashboard.py
if errorlevel 1 (
    echo.
    echo ERRO ao rodar o script Python.
    echo Verifique se o Python esta instalado e o Excel esta fechado.
    pause
    exit /b 1
)

echo.
echo [2/3] Enviando para o GitHub...
echo.
git add .
git commit -m "atualizacao automatica %date% %time%"
git push

if errorlevel 1 (
    echo.
    echo ERRO ao enviar para o GitHub.
    echo Verifique sua conexao com a internet.
    pause
    exit /b 1
)

echo.
echo  ============================================
echo   CONCLUIDO COM SUCESSO!
echo  ============================================
echo.
echo  Link do cliente:
echo  https://adrianac3gestao-lgtm.github.io/codecraft-dashboard
echo.
echo  O dashboard estara atualizado em ~1 minuto.
echo.
echo  Pressione qualquer tecla para fechar...
pause > nul
