Attribute VB_Name = "Specs"
'Public Sub RunTests()
'    Dim Reporter As New WorkbookReporter
'    Reporter.ConnectTo TestRunner
'
'    Reporter.Start NumSuites:=0
'    '                         ^ adjust NumSuites to match number of suites output
'    '                           (used for reporting progress)
'    ' Reporter.Output Suite1
'    ' Reporter.Output Suite2
'
'    Reporter.Done
'End Sub

Public Sub RunTests()
' Called from the TestRunner worksheet with a button.
    Dim Reporter As New WorkbookReporter
    ' Sets the reporter to use a certain sheet
    ' Therefore a sheet named TestRunner is required here.
    Reporter.ConnectTo TestRunner
    ' Sets the number of test suites to run. One for each module with tests in it.
    Reporter.Start NumSuites:=2
    ' Each suite is stored as a function in a module that returns a TestSuite
    ' So essentially tests can be stored each in there respective modules.
    Reporter.Output Utils_Tests.Tests
    Reporter.Output Table_Tests.Tests
    ' Display testing results to the worksheet
    Reporter.Done
End Sub
