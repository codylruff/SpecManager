Attribute VB_Name = "Module1"
Function ChangeActivePrinter() As String
Attribute ChangeActivePrinter.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ChangeActivePrinter Macro
'
    Application.Dialogs(xlDialogPrinterSetup).Show
    ChangeActivePrinter = Application.ActivePrinter
'
End Function
