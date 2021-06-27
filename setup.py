from distutils.core import setup
import py2exe
from glob import glob
import openpyxl

import sys
#data_files = [("Microsoft.VC90.CRT", glob(r'C:\Windows\WinSxS\x86_microsoft.vc90.crt_1fc8b3b9a1e18e3b_9.0.30729.1_none_e163563597edeada\*.*'))]

sys.argv.append('py2exe')

setup(
    console=['main.py'],
    zipfile="foo/bar.zip",
    options= {'py2exe':{'dll_excludes':['Secur32.dll','SHFOLDER.dll','CRYPT32.dll'],'packages':['openpyxl'], 'skip_archive':True}}
     )