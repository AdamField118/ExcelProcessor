import PyInstaller.__main__
import sys
import os

PyInstaller.__main__.run([
    'src/ets/main.py',
    '--onefile',
    '--windowed',
    '--name=ETSProcessor',
    '--icon=resources/icons/ets_logo.ico',
    '--add-data=resources;resources',
    '--paths=src',
    '--paths=src/ets',
    '--paths=src/gui',
    '--paths=src/core',
    '--paths=src/utils',
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
    # Our modules - with src prefix
    '--hidden-import=src.ets.processor',
    '--hidden-import=src.ets.gui',
    '--hidden-import=src.gui.components',
    '--hidden-import=src.core.base_processor',
    '--hidden-import=src.core.file_handler',
    '--hidden-import=src.utils.validators',
    '--hidden-import=src.utils.logger',
    '--hidden-import=src.utils.resource_manager',
    # Our modules - without src prefix (for frozen imports)
    '--hidden-import=ets.processor',
    '--hidden-import=ets.gui',
    '--hidden-import=gui.components',
    '--hidden-import=core.base_processor',
    '--hidden-import=core.file_handler',
    '--hidden-import=utils.validators',
    '--hidden-import=utils.logger',
    '--hidden-import=utils.resource_manager',
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