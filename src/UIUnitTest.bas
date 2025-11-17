Attribute VB_Name = "UIUnitTest"
'@IgnoreModule FunctionReturnValueDiscarded
'@TestModule
'@Folder "City_Grant_Address_Report.test"


Option Explicit
Option Private Module

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    SheetUtilities.TestSetupCleanup
End Sub

'@TestCleanup
Private Sub TestCleanup()
    SheetUtilities.TestSetupCleanup
End Sub

'@TestMethod
Public Sub TestDecemberPrintRequests()
    On Error GoTo TestFail
    
    MacroEntry AutocorrectAddressesSheet
    
    ' Can only use Fakes.Date once per test
    Fakes.Date.Returns "12/1/2024"
    Autocorrect.printRemainingRequests 8
    Assert.isTrue SheetUtilities.getAutocorrectRequestCharacters.text = "8 / 5000 left until January"
    Fakes.Date.Verify.AtLeastOnce
    
    MacroExit AutocorrectAddressesSheet
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
