@echo off
chcp 65001 >nul
echo ============================================================
echo   BUILD — Assinatura de Lotes SIFAMA
echo ============================================================
echo.

:: Usa sempre o Python 3.10 via modulo para evitar conflito com outras versoes
set PYTHON=python

:: Verifica se Python esta acessivel
%PYTHON% --version >nul 2>&1
if errorlevel 1 (
    echo [ERRO] Python nao encontrado. Instale o Python 3.10+ e tente novamente.
    pause
    exit /b 1
)

:: Instala/atualiza o PyInstaller usando python -m pip (evita conflito entre versoes do Python)
echo [1/3] Instalando/atualizando PyInstaller...
%PYTHON% -m pip install --upgrade pyinstaller
if errorlevel 1 (
    echo [ERRO] Falha ao instalar PyInstaller.
    pause
    exit /b 1
)

:: Instala dependencias do projeto
:: O Selenium Manager (embutido no Selenium 4.6+) baixa o ChromeDriver correto
:: automaticamente — sem chromedriver.exe fixo e sem webdriver-manager.
echo.
echo [2/3] Instalando dependencias (selenium, openpyxl)...
if exist "%~dp0requirements.txt" (
    %PYTHON% -m pip install -r "%~dp0requirements.txt"
) else (
    %PYTHON% -m pip install "selenium>=4.6.0" openpyxl
)
if errorlevel 1 (
    echo [ERRO] Falha ao instalar dependencias.
    pause
    exit /b 1
)

:: Limpa builds anteriores
echo.
echo [3/3] Gerando executavel (pode demorar alguns minutos)...
if exist "%~dp0build" rmdir /s /q "%~dp0build"
if exist "%~dp0dist"  rmdir /s /q "%~dp0dist"
echo.
%PYTHON% -m PyInstaller "%~dp0automacao_assinatura_lotes.spec"
if errorlevel 1 (
    echo.
    echo [ERRO] Falha ao gerar o executavel. Veja o log acima.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo   CONCLUIDO!
echo   Executavel gerado em: dist\AssinaturaLotesSIFAMA.exe
echo ============================================================
echo.
pause
