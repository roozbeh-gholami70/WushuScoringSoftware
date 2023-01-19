import sys

from cx_Freeze import setup, Executable

# command line python setup.py bdist_msi

company_name = "AminCom"
product_name = "Wushu Scoring Software"
iconFile = "Wushu.ico"

bdist_msi_options = {
"upgrade_code": "{ee0698ac-63dd-4555-8a80-2a49220e9a02}",
"add_to_path": False,
"initial_target_dir": r'[ProgramFiles64Folder]\\{}\\{}'.format(company_name, product_name),
"install_icon":iconFile,
"all_users":True,
}

# GUI applications require a different base on Windows
base = None
if sys.platform == 'win32':
    base = "Win32GUI"

if sys.platform == 'win64':
    base = "Win64GUI"

exe = Executable(script="mainWindowRef.py",
copyright="Copyright (C) 2023 AminCo",
base=base,
targetName = product_name,
shortcut_name=product_name,
shortcut_dir="DesktopFolder",
icon=iconFile,
)

setup(name=product_name,version='1.0.0', 
    description='', executables=[exe],
    options={'bdist_msi': bdist_msi_options, "build_exe": {"zip_include_packages": ["encodings"]}})