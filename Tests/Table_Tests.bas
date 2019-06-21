Attribute VB_Name = "Table_Tests"

Public Function Tests() As TestSuite
' Contains vba-tests for this module code
    Set Tests = New TestSuite
    Tests.Description = "ListObject Wrapper (Table)"
    
    Dim Reporter As New ImmediateReporter
    Reporter.ListenTo Tests
    
    Dim Suite As New TestSuite
    Dim Test As TestCase
    Dim tbl As SAATI_Data_Manager.Table
    Set tbl = PrepareTableObject

    With Tests.Test("")
        Set Test = Suite.Test("should pass")
        With Test
            .IsEqual shtTestTable.ListObjects("tblTest"), tbl
        End With
        
        .IsEqual Test.Result, TestResultType.Pass
    End With

End Function

Private Function PrepareTableObject() As Table
    Dim tbl As SAATI_Data_Manager.Table
    Set tbl = SAATI_Data_Manager.Factory.CreateTable
    Set tbl.ListObject = shtTestTable.ListObjects("tblTest")
    Set PrepareTableObject = tbl
End Function
