@echo off
chcp 65001 >nul
title Assinatura de Lotes DOU - SIFAMA
echo Iniciando automacao de Assinatura de Lotes...
python "%~dp0automacao_assinatura_lotes.py"
if errorlevel 1 (
    echo.
    echo ERRO: Verifique se o Python e o Selenium estao instalados.
    echo Execute: pip install selenium
    pause
)
