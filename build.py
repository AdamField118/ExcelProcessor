import PyInstaller.__main__
import sys
import os

PyInstaller.__main__.run([
    'src/main.py',
    '--onefile',
    '--windowed',
    '--name=ExcelProcessor',
    '--icon=resources/icons/pixel_logo.ico',
    '--add-data=resources;resources',
    '--paths=src',
    # Core dependencies
    '--hidden-import=openpyxl',
    '--hidden-import=pandas',
    '--hidden-import=xlrd',
    '--hidden-import=xlsxwriter',
    # PyQt5 dependencies
    '--hidden-import=PyQt5',
    '--hidden-import=PyQt5.QtCore',
    '--hidden-import=PyQt5.QtGui',
    '--hidden-import=PyQt5.QtWidgets',
    # Additional xlrd-related imports
    '--hidden-import=xlrd.biffh',
    '--hidden-import=xlrd.book',
    '--hidden-import=xlrd.sheet',
    '--hidden-import=xlrd.xldate',
    # Pandas engines
    '--hidden-import=pandas.io.excel._xlrd',
    '--hidden-import=pandas.io.excel._openpyxl',
    # Other potential missing imports
    '--hidden-import=logging.handlers',
    '--hidden-import=pkg_resources.py2_warn',
    # Exclude heavy modules
    '--exclude-module=torch',
    '--exclude-module=tensorflow',
    '--exclude-module=matplotlib',
    '--exclude-module=scipy',
    '--exclude-module=IPython',
    '--exclude-module=jupyter',
    '--exclude-module=notebook',
    # Build configuration
    '--distpath=dist',
    '--workpath=build',
    '--clean',
    '--noconfirm'
])