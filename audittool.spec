# -*- mode: python -*-
a = Analysis(['audittool.py'],
             pathex=['D:\\AuditingTool'],
             hiddenimports=[],
             hookspath=None,
             runtime_hooks=None)
pyz = PYZ(a.pure)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          Tree('D:\\AuditingTool', prefix='\\'),
          a.zipfiles,
          a.datas,
          name='audittool.exe',
          debug=False,
          strip=None,
          upx=True,
          console=False,
          icon='D:\\AuditingTool\\appolo.ico' )
