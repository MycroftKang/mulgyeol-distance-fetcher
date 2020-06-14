# -*- mode: python ; coding: utf-8 -*-
import os
block_cipher = None


a = Analysis(['main.py'],
             pathex=[os.curdir],
             binaries=[],
             datas=[],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          [],
          exclude_binaries=True,
          name='MDF',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          icon='../resources/app/MDF_Icon.ico',
          upx=True,
          version='version_info.txt',
          console=False )
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               upx_exclude=[],
               name='main')
