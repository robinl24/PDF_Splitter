# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['PDF__Splitter.py'],
             pathex=['/Users/i544474/code/QP_410_PDF_Splitter/PDF_Splitter'],
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
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='PDF__Splitter',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False )
app = BUNDLE(exe,
             name='PDF__Splitter.app',
             icon=None,
             bundle_identifier=None)
