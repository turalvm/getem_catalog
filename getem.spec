# getem.spec
a = Analysis(
    ['getem.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('lang', 'lang'),
        ('sound', 'sound'),
        ('help', 'help'),
        ('nvdaControllerClient64.dll', '.'),
        ('nvdaControllerClient.dll', '.'),
        ('version.txt', '.'),
    ],
    hiddenimports=[
        # CRYPTOGRAPHY (kodunda kullanılıyor)
        'cryptography',
        'cryptography.fernet',
        'cryptography.hazmat',
        'cryptography.hazmat.backends',
        'cryptography.hazmat.primitives',
        'cryptography.hazmat.primitives.ciphers',
        'cryptography.hazmat.primitives.hashes',
        'cryptography.hazmat.primitives.kdf',
        'cryptography.hazmat.primitives.kdf.pbkdf2',
        
        # PANDAS ve bağımlılıkları
        'pandas',
        'numpy',
        'numpy.core',
        'numpy._core',
        'dateutil',
        'pytz',
        
        # SELENIUM
        'selenium',
        'selenium.webdriver',
        'selenium.webdriver.common',
        'selenium.webdriver.common.by',
        'selenium.webdriver.common.keys',
        'selenium.webdriver.common.action_chains',
        'selenium.webdriver.chrome',
        'selenium.webdriver.chrome.options',
        'selenium.webdriver.chrome.service',
        'selenium.webdriver.chrome.webdriver',
        'selenium.webdriver.support',
        'selenium.webdriver.support.ui',
        'selenium.webdriver.support.expected_conditions',
        'selenium.webdriver.remote',
        'selenium.webdriver.remote.webelement',
        'selenium.common',
        'selenium.common.exceptions',
        
        # HTTP ve network
        'urllib3',
        'certifi',
        'charset_normalizer',
        'requests',
        
        # Windows ve sistem
        'ctypes',
        'ctypes.wintypes',
        'win32api',
        'win32con',
        'win32process',
        'win32event',
        'win32file',
        'win32gui',
        'win32com',
        'win32com.client',
        
        # wxPython
        'wx',
        'wx._core',
        'wx._controls',
        
        # Diğer
        'json',
        'pickle',
        'threading',
        'subprocess',
        'zipfile',
        'shutil',
        'webbrowser',
        'winsound',
        're',
        'time',
        'os',
        'sys',
        'io',
        'glob',
        'tempfile',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

# version.txt dosyasından sürüm numarasını oku
version_file = open('version.txt', 'r')
version = version_file.read().strip()
version_file.close()

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='getem_catalog',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)