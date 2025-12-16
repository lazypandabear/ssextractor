# app.spec
block_cipher = None

a = Analysis(
    ['app.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        ('templates', 'templates'),
        # ('service_account.json', '.')  # Uncomment if you bundle it
    ],
    hiddenimports=[],
    hookspath=['.'],  # Ensure your custom hook is picked up
    runtime_hooks=[],
    excludes=[]
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    name='ssextractor',
    debug=False,
    strip=False,
    upx=True,
    console=True
)
