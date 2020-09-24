Attribute VB_Name = "LinkBreaker"
' *********************************************************************
'  TITLE:   EXTERNAL LINK UTILITY
'  PURPOSE: Finds all external links in a workbook, including the very
'           hard to find ones. Cleans some links automatically and
'           provides instructions for how to manually remove the others.
'  NOTES:   This can take 2 or 3 minutes to run if a workbook contains
'           a large number of external links.
'  HOW TO:  Open the affected workbook and run this macro.
'  AUTHOR:  jramm
'           https://stackoverflow.com/questions/48337861/removing-external-links
' *********************************************************************
'

' GLOBAL VARIABLES
' ====================
Dim g_ResultBook As Workbook

' MAIN SUB
' ====================
Sub ExternalLinkUtility()
    Excel.Application.ScreenUpdating = False

    ReportExternalLinks ActiveWorkbook

    Excel.Application.ScreenUpdating = True
    If Not g_ResultBook Is Nothing Then
        g_ResultBook.Activate 'bring the result book into view if it's not already.
        Set g_ResultBook = Nothing
    End If
End Sub



'FUNCTION: OutputLinkInfo
'PARAMETERS:
'    wbk - full workbook filepath (Workbook.FullName)
'    wsh - worksheet name (Worksheet.Name)
'    adr - cell address string (A1) or an empty string ("") to omit hyperlink to issue location
'    loc - friendly name we want reader to see (such as "Cell B4" or "My Cool Chart")
'    fml - external link formula that is causing the problem
'    txt - fix instructions (or other notes)
Function OutputLinkInfo(typ As String, wbk As String, wsh As String, loc As String, adr As String, fml As String, txt As String)
    Static resultLn As Long
    'first time called: Create result workbook
    '=========================================
    If g_ResultBook Is Nothing Then
        Set g_ResultBook = Workbooks.Add
        With g_ResultBook.Worksheets.item(1)
            'title row
            .Range("A1").value = "External Link Report"
            .Range("A1").Font.Bold = True
            .Range("A1").Font.Size = 18
            .Range("A1:F1").Interior.Color = RGB(0, 112, 192)
            .Range("A1:F1").Font.Color = RGB(255, 255, 255)
            'column headers row
            .Range("A2").value = "Type"
            .Range("B2").value = "Workbook"
            .Range("C2").value = "Worksheet"
            .Range("D2").value = "Location"
            .Range("E2").value = "Reference"
            .Range("F2").value = "Fix Instructions"
            .Range("A2:F2").Interior.Color = RGB(221, 235, 247)
            .Range("A2:F2").Font.Bold = True
            'set column widths
            .Columns("A").ColumnWidth = 22
            .Columns("B").ColumnWidth = 15
            .Columns("C").ColumnWidth = 28
            .Columns("D").ColumnWidth = 28
            .Columns("E").ColumnWidth = 60
            .Columns("F").ColumnWidth = 60
            'add filter
            .Range("A2:F2").AutoFilter
        End With
        resultLn = 2
    End If

    'every time called: Write single result line using the passed parameters
    '=======================================================================
    resultLn = resultLn + 1

    With g_ResultBook.Worksheets.item(1)
        .Range("A" & resultLn).value = typ
        .Range("B" & resultLn).value = dir(wbk) 'Dir gets us only the filename from the end of the full path
        .Range("C" & resultLn).value = wsh
        .Range("D" & resultLn).value = loc
        If (Len(adr) > 0) And (Len(dir(wbk)) > 0) Then
            .Hyperlinks.Add .Range("D" & resultLn), wbk, "'" & wsh & "'!" & adr, "Jump to this issue", loc
        End If
        .Range("E" & resultLn).value = "'" & fml 'prepend apostrophe to force formula to display as plain text
        .Range("F" & resultLn).value = txt
    End With

End Function



'FUNCTION: OutputLinkInfo
'PARAMETERS:
'    wkbk - workbook to check for external links
Function ReportExternalLinks(wkbk As Excel.Workbook) As String()
    Dim wksht As Excel.Worksheet
    Dim cell As Excel.Range
    Dim numLinks As Integer
    Dim fml As String
    Dim r As Range
    numLinks = 0 'Note that numLinks causes a Runtime error if this macro detects >32,768 external links. The
                 'macro should probably be updated at some point to more gracefully handle this situation, but
                 'I haven't gotten around to it because that scenario is very unlikely.

'``````````````````````````````````````````````````````````
'WORKSHEET-LEVEL CHECKS are performed in the following loop
',,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,

    For Each wksht In wkbk.Worksheets

    ' Search for external links in cell formulas
    ' ==========================================
        For Each cell In wksht.usedRange.Cells
            On Error Resume Next
            fml = cell.Formula
            If err.Number <> 0 Then
                err.Clear
            ElseIf (InStr(fml, "[") <> 0) And (InStr(fml, ".xl") <> 0) Then
                ' the ".xl" check was added to avoid false positives when a user enters brackets in the cell
                ' (for example, if the cell text is "[Test]"). However, this additional check probably causes
                ' this part of the macro to miss external data connections, which won't have .xl in their name
                    On Error GoTo 0
                    numLinks = numLinks + 1
                    OutputLinkInfo "Cell Formula", _
                                   wkbk.FullName, _
                                   wksht.Name, _
                                   "Cell " & cell.Address(False, False), _
                                   cell.Address, _
                                   fml, _
                                   "Delete the formula from this cell."
            End If
            On Error GoTo 0
        Next cell

    ' Search for external links in shapes
    ' ===================================
        Dim shp As Shape
        Dim subshp As Shape
        For Each shp In wksht.shapes
            On Error Resume Next
            fml = shp.DrawingObject.Formula 'will throw an error whenever the shape doesn't have a formula
            If err.Number <> 0 Then
                err.Clear
            ElseIf InStr(fml, "[") <> 0 Then
                On Error GoTo 0
                numLinks = numLinks + 1
                OutputLinkInfo "Shape/Object", _
                               wkbk.FullName, _
                               wksht.Name, _
                               shp.Name, _
                               shp.TopLeftCell.Address & ":" & shp.BottomRightCell.Address, _
                               fml, _
                               "Select the shape. The shape's formula appears in the Excel formula bar. Delete the external reference."
            End If
            On Error GoTo 0

            'iterate subshapes for any groups (supposedly this should catch all no matter how nested they are, but I only tested normal groups 1-level deep)
            If shp.Type = msoGroup Then
                For Each subshp In shp.GroupItems
                    On Error Resume Next
                    fml = subshp.DrawingObject.Formula
                    If err.Number <> 0 Then
                        err.Clear
                    ElseIf InStr(fml, "[") <> 0 Then
                        On Error GoTo 0
                        numLinks = numLinks + 1
                        OutputLinkInfo "Shape/Object", _
                                       wkbk.FullName, _
                                       wksht.Name, _
                                       subshp.Name & " (part of shape group '" & shp.Name & "')", _
                                       subshp.TopLeftCell.Address & ":" & subshp.BottomRightCell.Address, _
                                       fml, _
                                       "Select the shape. The shape's formula appears in the Excel formula bar. Delete the external reference."
                    End If
                    On Error GoTo 0
                Next subshp
            End If
        Next shp

    ' Search for external links in conditional formatting
    ' ===================================================
    ' NOTE: external links in conditional formatting (CF) are some of the weirdest. You can open the CF window
    ' for the cell in Excel, and you won't see any external links in the formula, so there's no way to manually
    ' fix it besides deleting the CF from the cell entirely or copy-and-pasting a valid CF cell over the top of the
    ' affected cell to replace it. I have seen workbooks with hundreds of CF external links, and you can open the
    ' affected cell's CF rule in Excel, and then open a nearby CF rule that does not have an external link, and they
    ' look identical in the CF window in Excel (even though .Formula1 and other .Formula properties are not the
    ' same when accessed from VBA) I have written some code to automatically fix very specific CF rules with
    ' external links, but it would be very difficult to write generic code that could fix any CF rule that has an
    ' external link. There are far too many CF conditions, operators, formulas, and other details and no simple way
    ' to determine how to "fix" them programmatically.
        Dim cForm As Object
        For Each cForm In wksht.Cells().FormatConditions
            On Error Resume Next
            fml = cForm.Formula1
            If err.Number <> 0 Then
                err.Clear
            ElseIf InStr(fml, "[") <> 0 Then
                On Error GoTo 0
                numLinks = numLinks + 1
                OutputLinkInfo "Conditional Formatting", _
                               wkbk.FullName, _
                               wksht.Name, _
                               "Cell " & cForm.AppliesTo.Address(False, False), _
                               cForm.AppliesTo.Address, _
                               fml, _
                               "Select the cell and open the conditional formatting window (Home > Conditional Formatting). " & _
                               "Delete the external link from the conditional formatting formula if you see it. In some cases, " & _
                               "you cannot see external links in the conditional formatting formula. In that scenario, either " & _
                               "delete the conditional formatting from the cell, or copy-and-paste a different cell's valid " & _
                               "conditional formatting over the top of the affected cell in order to fix the issue."
            End If
            On Error GoTo 0
        Next cForm

    ' Search for external links in charts
    ' ===================================
        Dim cht As Excel.ChartObject
        Dim srs As Excel.Series
        Dim chartName As String
        For Each cht In wksht.ChartObjects
            For Each srs In cht.Chart.SeriesCollection
                On Error Resume Next
                fml = srs.Formula
                If err.Number <> 0 Then
                    err.Clear
                ElseIf InStr(fml, "[") <> 0 Then
                    On Error GoTo 0
                    numLinks = numLinks + 1
                    If cht.Chart.HasTitle Then
                        chartName = cht.Chart.ChartTitle.text 'This is the better option when available
                    Else
                        chartName = cht.Chart.Name & " (" & cht.Name & ")"
                    End If
                    OutputLinkInfo "Chart", _
                                   wkbk.FullName, _
                                   wksht.Name, _
                                   chartName, _
                                   cht.TopLeftCell.Address & ":" & cht.BottomRightCell.Address, _
                                   fml, _
                                   "Right-click the chart > Select Data... Click Edit on each series in the Legend Entries " & _
                                   "(Series) list. Remove the external link in the formulas you find there."
                End If
                On Error GoTo 0
            Next srs
        Next cht

    ' Search for external links in pivot tables
    ' =========================================
        Dim pvt As Excel.PivotTable
        For Each pvt In wksht.PivotTables
            On Error Resume Next
            fml = pvt.SourceData
            If err.Number <> 0 Then
                err.Clear
            ElseIf InStr(fml, "[") <> 0 Then
                On Error GoTo 0
                numLinks = numLinks + 1
                OutputLinkInfo "PivotTable", _
                               wkbk.FullName, _
                               wksht.Name, _
                               pvt.Name, _
                               pvt.TableRange1.Address, _
                               fml, _
                               "Click the PivotTable. In the Excel ribbon, go to Analyze > Change Data Source. " & _
                               "Delete the external link from the formula you find there."
            End If
            On Error GoTo 0
        Next pvt

    ' Search for external links in data validation
    ' ============================================
        'NOTE: this section of the code can take a few minutes to run on workbooks where the data validation
        'was applied to an entire column, because it iterates through every cell in the column separately.
        'Probably there's a smarter way to improve the performance of this part of the macro for such scenarios,
        'but I haven't gotten around to trying to improve it.
        Dim dataValExtLinkRanges As Object
        Dim Key As Variant
        Set dataValExtLinkRanges = CreateObject("Scripting.Dictionary")

        'first, iterate over cells with data validation and UNION together the cells associated with each unique external link
        On Error Resume Next
        Set r = wksht.Cells.SpecialCells(xlCellTypeAllValidation)
        If err.Number <> 0 Then
            err.Clear
        Else
            For Each cell In r.Cells
                On Error Resume Next
                fml = cell.Validation.Formula1
                If err.Number <> 0 Then
                    err.Clear
                ElseIf InStr(fml, "[") <> 0 Then
                    On Error GoTo 0
                    'add to dictionary, updating existing range if identical external link was already found
                    Key = fml
                    If dataValExtLinkRanges.Exists(Key) Then
                        Set dataValExtLinkRanges.item(Key) = Application.Union(dataValExtLinkRanges(Key), cell)
                    Else
                        Set dataValExtLinkRanges.item(Key) = cell
                    End If
                End If
            Next cell
        End If
        On Error GoTo 0

        Dim contiguousAddresses() As String
        Dim i As Long
        Dim place As String

        'report the data validation ranges we found that contain external links
        For Each Key In dataValExtLinkRanges.keys()
            contiguousAddresses = VBA.Split(dataValExtLinkRanges(Key).Address, ",") 'split non-contiguous ranges into separate entries
            For i = 0 To UBound(contiguousAddresses)
                numLinks = numLinks + 1
                If Range(contiguousAddresses(i)).CountLarge > 1 Then 'this is just to pluralize "Cells" if there's more than one
                    place = "Cells " & VBA.Replace(contiguousAddresses(i), "$", "")
                Else
                    place = "Cell " & VBA.Replace(contiguousAddresses(i), "$", "")
                End If
                OutputLinkInfo "Data Validation", _
                               wkbk.FullName, _
                               wksht.Name, _
                               place, _
                               contiguousAddresses(i), _
                               VBA.CStr(Key), _
                               "Select the cell and open the data validation window (Data > Data Validation). " & _
                               "Remove the external reference from the data validation formula."
            Next i
        Next Key
        Set dataValExtLinkRanges = Nothing 'clear the dictionary object

' CONTINUE TO NEXT WORKSHEET
' ==========================
    Next wksht


'`````````````````````````````````````````
'WORKBOOK-LEVEL CHECKS are performed below
',,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,

    'reset error handler
    On Error GoTo 0

' Search for external links in named ranges
' =========================================
    'NOTE: This section should be improved by also searching for places where each named range is used
    'in the workbook. It could then delete any named ranges that are unused and leave those that are
    'used, providing the user with more detail about where they're used and how to manually clean them up.
    'For now, it just deletes any ranges with a broken #REF! in the external link or where the external
    'link can't be resolved to a file that actually exists.
    Dim FSO As Object
    Dim startPos As Long
    Dim endPos As Long
    Dim pathPos As Long
    Dim delCt As Long
    delCt = 0
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If wkbk.Names.Count > 0 Then
        For nameCnt = wkbk.Names.Count To 1 Step -1
            If InStr(wkbk.Names(nameCnt).RefersTo, "[") <> 0 Then
                If InStr(wkbk.Names(nameCnt).RefersTo, "#REF!") <> 0 Then 'if it's a broken reference, just delete it
                    wkbk.Names(nameCnt).Delete
                    delCt = delCt + 1
                Else
                    'check the actual filepath to see if it can be resolved.
                    startPos = VBA.InStr(1, wkbk.Names(nameCnt).RefersTo, "='") '+ 2
                    endPos = VBA.InStr(startPos, wkbk.Names(nameCnt).RefersTo, "]") '- 1
                    pathPos = VBA.InStr(1, wkbk.Names(nameCnt).RefersTo, "\") 'verify that this is a filepath (includes filepath folder delimiter)
                    If startPos > 0 And endPos > 0 And pathPos > 0 And FSO.FileExists(VBA.Replace(VBA.Mid(wkbk.Names(nameCnt).RefersTo, startPos + 2, endPos - startPos - 2), "[", "")) = False Then
                        'this is a filepath to a file that does not exist - delete it
                        wkbk.Names(nameCnt).Delete
                        delCt = delCt + 1
                    Else 'external reference does exist - reveal it in Name Manager and tell the user to manually clean it up
                        wkbk.Names(nameCnt).Visible = True
                        numLinks = numLinks + 1
                        OutputLinkInfo "Named Range", _
                                       wkbk.FullName, _
                                       "N/A", _
                                       wkbk.Names(nameCnt).Name, _
                                       "", _
                                       wkbk.Names(nameCnt).RefersTo, _
                                       "Open the name manager (Formulas > Name Manager). This named range has been unhidden and you can now delete it manually."
                    End If
                End If
            End If
        Next nameCnt
    End If

    Set FSO = Nothing

    'report all automatically deleted named ranges as a single entry
    If delCt > 0 Then
        numLinks = numLinks + 1
        OutputLinkInfo "Named Range", _
                       wkbk.FullName, _
                       "N/A", _
                       "(" & delCt & " named ranges)", _
                       "", _
                       "Unrecorded", _
                       "These named ranges included unresolvable external link references and were automatically removed by the utility. " & _
                       "Save the " & dir(wkbk.FullName) & " workbook to preserve the changes."
    End If

' Broadcast message that the utility is finished
' ==============================================
    If numLinks <= 0 Then
        MsgBox ("The utility is finished." & vbNewLine & vbNewLine & "No external links were found in " & dir(wkbk.FullName))
    Else
        If delCt > 0 Then
            MsgBox ("The utility is finished. " & vbNewLine & vbNewLine & (numLinks - 1) & " external links were found that require manual cleanup." _
                    & vbNewLine & delCt & " external links were automatically cleaned up by the utility.")
        Else
            MsgBox ("The utility is finished." & vbNewLine & vbNewLine & numLinks & " external links were found that require manual cleanup.")
        End If
    End If

End Function
