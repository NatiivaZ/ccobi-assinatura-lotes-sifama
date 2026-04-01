# -*- mode: python ; coding: utf-8 -*-
# Arquivo de configuração do PyInstaller para gerar o .exe
# Execute: pyinstaller automacao_assinatura_lotes.spec

import os
import site

block_cipher = None

# Pasta onde está o projeto
PROJ = os.path.dirname(os.path.abspath(SPEC))

# Localiza o selenium-manager.exe (binário interno do Selenium que detecta e baixa
# o ChromeDriver correto em tempo de execução — obrigatório no bundle PyInstaller)
def _selenium_manager_bin():
    import selenium as _sel
    caminho = os.path.join(
        os.path.dirname(_sel.__file__),
        "webdriver", "common", "windows", "selenium-manager.exe"
    )
    if not os.path.exists(caminho):
        raise FileNotFoundError(
            f"selenium-manager.exe não encontrado em: {caminho}\n"
            "Verifique se o Selenium >= 4.6.0 está instalado."
        )
    return caminho

a = Analysis(
    [os.path.join(PROJ, 'automacao_assinatura_lotes.py')],
    pathex=[PROJ],
    binaries=[
        # selenium-manager.exe: detecta a versão do Chrome e baixa o ChromeDriver correto.
        # Sem ele o .exe não consegue iniciar o navegador.
        (_selenium_manager_bin(), os.path.join("selenium", "webdriver", "common", "windows")),
    ],
    datas=[],
    hiddenimports=[
        # Selenium — core
        'selenium',
        'selenium.webdriver',
        'selenium.webdriver.chrome',
        'selenium.webdriver.chrome.webdriver',
        'selenium.webdriver.chrome.service',
        'selenium.webdriver.chrome.options',
        'selenium.webdriver.chromium',
        'selenium.webdriver.chromium.webdriver',
        'selenium.webdriver.chromium.options',
        'selenium.webdriver.chromium.service',
        'selenium.webdriver.remote',
        'selenium.webdriver.remote.webdriver',
        'selenium.webdriver.remote.command',
        'selenium.webdriver.remote.remote_connection',
        'selenium.webdriver.remote.errorhandler',
        'selenium.webdriver.remote.webelement',
        'selenium.webdriver.common.by',
        'selenium.webdriver.common.action_chains',
        'selenium.webdriver.common.keys',
        'selenium.webdriver.common.options',
        'selenium.webdriver.common.desired_capabilities',
        'selenium.webdriver.common.service',
        'selenium.webdriver.support',
        'selenium.webdriver.support.ui',
        'selenium.webdriver.support.expected_conditions',
        'selenium.webdriver.support.wait',
        'selenium.common',
        'selenium.common.exceptions',
        # Tkinter
        'tkinter',
        'tkinter.ttk',
        'tkinter.messagebox',
        # Exportação XLSX
        'openpyxl',
        'openpyxl.workbook',
        'openpyxl.worksheet',
        'openpyxl.styles',
        'openpyxl.cell',
        # Rede / certificados (usados internamente pelo Selenium)
        'urllib3',
        'urllib3.util',
        'urllib3.util.retry',
        'certifi',
        'ssl',
        'socket',
        'trio',
        'requests',
        'packaging',
        'packaging.version',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='AssinaturaLotesSIFAMA',   # nome do .exe gerado
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,                       # comprime o executável (precisa do UPX instalado; se falhar mude para False)
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,                  # False = sem janela de terminal preta (apenas GUI)
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,                      # coloque o caminho de um .ico aqui se quiser ícone personalizado
)
