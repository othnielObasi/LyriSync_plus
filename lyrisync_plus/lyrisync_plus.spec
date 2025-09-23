# lyrisync_plus.spec
# PyInstaller spec for LyriSync+
# Generates a single-windowed EXE, bundling splash image and including async deps.

from PyInstaller.utils.hooks import collect_submodules
import os

block_cipher = None

# If you keep the splash.png with your code, this picks it up
datas = []
if os.path.exists('splash.png'):
    datas.append(('splash.png', '.'))

# If you ship any default config file, add it here too (optional)
# if os.path.exists('lyrisync_config.yaml'):
#     datas.append(('lyrisync_config.yaml', '.'))

# Some ttkbootstrap themes or resources are auto-collected by PyInstaller,
# but this makes sure hidden modules are included.
hidden = []
hidden += collect_submodules('aiohttp')
hidden += collect_submodules('websockets')
hidden += collect_submodules('ttkbootstrap')

a = Analysis(
    ['main.py'],
    pathex=['.'],
    binaries=[],
    datas=datas,
    hiddenimports=hidden,
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
    name='LyriSyncPlus',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,   # windowed app
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,       # set to 'app.ico' if you have one and add to datas
)

# Onefile is convenient for distribution
coll = COLLECT(exe, a.binaries, a.zipfiles, a.datas, strip=False, upx=True, upx_exclude=[], name='LyriSyncPlus')
