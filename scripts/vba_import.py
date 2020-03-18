#Import the following library to make use of the DispatchEx to run the macro
import win32com.client as wincl
import os

def runMacro():
    file_path = r'C:\Users\cruff\Documents\Projects\source\Spec-Manager\Spec Manager v1.4.8.xlsm'
    if os.path.exists(file_path):
        # DispatchEx is required in the newest versions of Python
        excel_macro = wincl.DispatchEx("Excel.application")
        excel_path = os.path.expanduser(file_path)
        workbook = excel_macro.Workbooks.Open(Filename = excel_path, ReadOnly =1)
        print('Removing Existing Modules . . .')
        excel_macro.Application.Run\
            ("ThisWorkbook.RemoveAll")
        print('Importing New Modules . . .')
        excel_macro.Application.Run\
            ("ThisWorkbook.VSImport")
        #Save the results in case you have generated data
        workbook.Save()
        excel_macro.Application.Quit()
        del excel_macro
    else:
        print("Import Failed")

if __name__ == "__main__":
    runMacro()