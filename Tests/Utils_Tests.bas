Attribute VB_Name = "Utils_Tests"

Public Function Tests() As TestSuite
' Contains vba-tests for this module code
    Set Tests = New TestSuite
    Tests.Description = "Utils"
    
    Dim Reporter As New ImmediateReporter
    Reporter.ListenTo Tests
    
    Dim Suite As New TestSuite
    Dim Test As TestCase

    With Tests.Test("RemoveWhiteSpace(target As String) As String should remove all whitespace from given string")
        Set Test = Suite.Test("should pass")
        With Test
            .IsEqual SAATI_Data_Manager.Utils.RemoveWhiteSpace(" "), vbNullString
            .IsEqual SAATI_Data_Manager.Utils.RemoveWhiteSpace(" 1 A !"), "1A!"
        End With
        
        .IsEqual Test.Result, TestResultType.Pass
    End With

    With Tests.Test("ConvertToCamelCase(s As String) As String")
        Set Test = Suite.Test("should pass")
        With Test
            .IsEqual SAATI_Data_Manager.Utils.ConvertToCamelCase("camel case"), "CamelCase"
            .IsEqual SAATI_Data_Manager.Utils.ConvertToCamelCase("1Camel Case_Test _with! symbols"), "CamelCaseTestWithSymbols"
            
        End With
        
        .IsEqual Test.Result, TestResultType.Pass
    End With

    With Tests.Test("SplitCamelCase(sString As String, Optional sDelim As String = ' ') As String")
        Set Test = Suite.Test("should pass")
        With Test
            .IsEqual SAATI_Data_Manager.Utils.SplitCamelCase("CamelCase"), "Camel Case" ' Passing
            .IsEqual SAATI_Data_Manager.Utils.SplitCamelCase("1Camel Case_Test _with! symbols"), "1Camel Case_Test _with! symbols" ' Passing
            .IsEqual SAATI_Data_Manager.Utils.SplitCamelCase("SnakeCase", sDelim:="_"), "Snake_Case"
        End With
        
        .IsEqual Test.Result, TestResultType.Pass
    End With
    
    With Tests.Test("Should convert column numbers to letters")
        Set Test = Suite.Test("should pass")
        With Test
            .IsEqual SAATI_Data_Manager.Utils.ConvertNumericToAlpha(27), "AA"
        End With
        
        .IsEqual Test.Result, TestResultType.Pass
    End With
    
    With Tests.Test("Should convert column numbers to letters")
        Set Test = Suite.Test("should pass")
        With Test
            .IsEqual SAATI_Data_Manager.Utils.ConvertNumericToAlpha(27), "AA"
        End With
        
        .IsEqual Test.Result, TestResultType.Pass
    End With
    
    With Tests.Test("RemoveSheet(ws As Worksheet) As Boolean ")
        Set Test = Suite.Test("should pass")
        Dim ws As New Worksheet
        Set ws = SAATI_Data_Manager.Utils.CreateNewSheet("test")
        With Test
            .IsEqual SAATI_Data_Manager.Utils.RemoveSheet(ws), True
        End With
        Set ws = Nothing
        .IsEqual Test.Result, TestResultType.Pass
    End With
    
    
    With Tests.Test("CreateNamedRange(RangeName As String, sht As Worksheet, r As Long, c As Long) As Boolean ")
        Set Test = Suite.Test("should pass")
        Set ws = SAATI_Data_Manager.Utils.CreateNewSheet("test")
        With Test
            .IsEqual SAATI_Data_Manager.Utils.CreateNamedRange("test_range", ws, 1, 1), True
        End With
        SAATI_Data_Manager.Utils.RemoveSheet ws
        Set ws = Nothing
        .IsEqual Test.Result, TestResultType.Pass
    End With

End Function
