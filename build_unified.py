import PyInstaller.__main__
import sys
import os

PyInstaller.__main__.run([
    'src/main.py',
    '--onefile',
    '--windowed',
    '--name=FileProcessor',
    '--icon=resources/icons/pixel_logo.ico',
    '--add-data=resources;resources',
    '--paths=src',
    '--paths=src/gui',
    '--paths=src/core',
    '--paths=src/ets',
    '--paths=src/processors',
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
    '--hidden-import=src.gui.unified_window',
    '--hidden-import=src.gui.components',
    '--hidden-import=src.ets.processor',
    '--hidden-import=src.processors.unified_processor',
    '--hidden-import=src.core.base_processor',
    '--hidden-import=src.core.file_handler',
    '--hidden-import=src.core.excel_processor',
    '--hidden-import=src.utils.validators',
    '--hidden-import=src.utils.logger',
    '--hidden-import=src.utils.resource_manager',
    # Our modules - without src prefix (for frozen imports)
    '--hidden-import=gui.unified_window',
    '--hidden-import=gui.components',
    '--hidden-import=ets.processor',
    '--hidden-import=processors.unified_processor',
    '--hidden-import=core.base_processor',
    '--hidden-import=core.file_handler',
    '--hidden-import=core.excel_processor',
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