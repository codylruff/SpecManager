#Import the following library to make use of the DispatchEx to run the macro
import win32com.client as wincl
import shutil, os, argparse, json
# This is not really a compiler... It works by calling a set of macros from within a workbook that
# remove old source and add modified source code then save the workbook.
def get_arguments():
    """Parse the commandline arguments from the user"""

    parser = argparse.ArgumentParser(description='Compile Spec Manager `vX.Y.Z`.xlsm')
    parser.add_argument('-t', help='the version number target')

    return parser.parse_args()

def runCompiler(target):
    file_path = target
    if os.path.exists(file_path):
        # DispatchEx is required in the newest versions of Python
        excel_macro = wincl.DispatchEx("Excel.application")
        excel_path = os.path.expanduser(file_path)
        workbook = excel_macro.Workbooks.Open(Filename = excel_path, ReadOnly =1)
        excel_macro.Application.Run\
            ("ThisWorkbook.RemoveAll")
        excel_macro.Application.Run\
            ("ThisWorkbook.VSImport")
        #Save the results in case you have generated data
        workbook.Save()
        excel_macro.Application.Quit()
        del excel_macro
    else:
        print("Compilation Failed")

def main():
    # Get arguments
    args = get_arguments()
    target = args.t
    if target == '':
        target = r'C:\Users\cruff\Documents\Projects\source\Spec-Manager\Spec Manager v2.0.0.xlsm'
    else:
        target = r'C:\Users\cruff\Documents\Projects\source\Spec-Manager\Spec Manager ' + target + '.xlsm'
    
    runCompiler(target)

if __name__ == "__main__":
    main()