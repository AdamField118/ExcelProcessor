import PyInstaller.__main__

PyInstaller.__main__.run([
    'src/main.py',
    '--onefile',
    '--windowed',
    '--name=ExcelProcessor',
    '--icon=resources/icons/app_icon.png',
    '--add-data=resources;resources',
    '--distpath=dist',
    '--workpath=build',
    '--clean'
])