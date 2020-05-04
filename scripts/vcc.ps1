$Excel = New-Object -comobject Excel.Application
$FilePath = "C:\Users\cruff\Documents\Projects\source\Spec-Manager\Spec Manager v2.0.0.xlsm"
#$Excel.WindowState = -4140
$Excel.visible = $false
$wb = $Excel.Workbooks.Open($FilePath)
$Excel.Run("ThisWorkbook.RemoveAll")
$Excel.Run("ThisWorkbook.VSImport")
$wb.save()
$wb.close()
$Excel.quit()
