# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['monday.pyw'],
             pathex=['C:\\Users\\PMiller1\\code\\monday'],
             binaries=[],
             datas=[
               ('C:\\Users\\PMiller1\\code\\monday\\getJobBoardGroups.gql', '.\gql'),
               ('C:\\Users\\PMiller1\\code\\monday\\getJobBoardPulses.gql', '.\gql'),
               ('C:\\Users\\PMiller1\\code\\monday\\updateJobBoard.gql', '.\gql'),
             ],
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
          name='updateMondayJobBoard',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False , icon='icon.ico')
