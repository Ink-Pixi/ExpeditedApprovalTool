from cx_Freeze import setup, Executable

exe=Executable(
     script='approvalTool.py',
     base="Win32Gui",
     icon="icon/PS_Icon.ico"
     )
includefiles=['icon/', 'libEGL.dll', 'Expedited Approval.lnk']
includes=['decimal', 'atexit']
excludes=['Tkinter']
packages=[]
setup(

     version = "1.0.1",
     description = "Module for approving expedited orders",
     author = "David Hoy",
     name = "Expedited Approval",
     options = {'build_exe': {'excludes':excludes,'packages':packages,'include_files':includefiles,'includes':includes}},
     executables = [exe]
     )