# main.spec
# command : pyinstaller main.spec
a = Analysis(
    ['main.py'],
    hiddenimports=[
        'selenium.webdriver.chrome.webdriver',
        'selenium.webdriver.chrome.service',
        'selenium.webdriver.chrome.options',
        'selenium.webdriver.common.by',
        'selenium.webdriver.common.keys',
        'selenium.webdriver.support.ui',
        'selenium.webdriver.support.expected_conditions',
        'selenium.webdriver.remote.webelement',
        'selenium.webdriver.remote.command',
        'webdriver_manager.chrome',
        'webdriver_manager.core.os_manager',
        'webdriver_manager.core.driver_cache',
        'win32com.client',
        'pythoncom',
        'pywintypes',
    ],
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='main',
    console=True,   # keep True so you can see errors
    onefile=True,
    icon='icon.ico',
)