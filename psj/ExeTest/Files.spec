# -*- mode: python -*-

block_cipher = None


a = Analysis(['Files', '(x86)\\Windows', 'Kits\\10\\Redist\\10.0.20348.0\\ucrt\\DLLs\\x86', 'exeTest.py'],
             pathex=['C:\\Program', 'D:\\psj\\ExeTest'],
             binaries=[],
             datas=[],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name='Files',
          debug=False,
          strip=False,
          upx=True,
          runtime_tmpdir=None,
          console=True )
