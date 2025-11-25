Attribute VB_Name = "RecordsIntegrationTest"
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
    
    getAPIKeyRng.value = Split(getCSV(LibFileTools.GetLocalPath(ThisWorkbook.path) & "\apikeys.csv")(0), ",")(1)
    
    ' ScreenUpdating, Visible result in buggy behavior, don't turn on
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    
    getAPIKeyRng.value = vbNullString
    
    MacroExit InterfaceSheet
End Sub

'@TestInitialize
Private Sub TestInitialize()
    SheetUtilities.TestSetupCleanup
End Sub

'@TestCleanup
Private Sub TestCleanup()
    SheetUtilities.TestSetupCleanup
End Sub

Private Sub PasteTestRecords(ByVal csvPath As String, ByVal pasteFn As String)
    Dim bookToCopy As Workbook
    Set bookToCopy = Workbooks.Open(csvPath)
    Dim rngToCopy As Range
    Set rngToCopy = bookToCopy.Sheets.[_Default](1).UsedRange.Offset(1, 0)
    
    Dim width As Long
    Select Case pasteFn
        Case "InterfaceButtons.PasteInterfaceRecords"
            width = 15
        Case "InterfaceButtons.confirmPasteRxRecordsCalculate"
            width = 21
    End Select
    
    Set rngToCopy = rngToCopy.Resize(rngToCopy.rows.count - 1, width)
    rngToCopy.Copy
    
    ThisWorkbook.Activate
    Application.Run (pasteFn)
    
    setClipboardToBlankLine
    
    bookToCopy.Close
End Sub

Private Sub PasteInterfaceTestRecords(ByVal csvPath As String)
    PasteTestRecords csvPath, "InterfaceButtons.PasteInterfaceRecords"
End Sub

Public Sub PasteRxTestRecordsCalculate(ByVal csvPath As String)
    PasteTestRecords csvPath, "InterfaceButtons.confirmPasteRxRecordsCalculate"
End Sub

'@TestMethod
Public Sub TestMonthNonRxReport()
    On Error GoTo TestFail
    Fakes.MsgBox.Returns vbYes
    
    MacroEntry InterfaceSheet
    PasteInterfaceTestRecords ThisWorkbook.path & "\testdata\testNonRxMonth.csv"
    
    InterfaceButtons.confirmAddRecords
    InterfaceButtons.confirmGenerateNonRxReport

    CompareSheetCSV Assert, NonRxReportSheet.Name, ThisWorkbook.path & "\testdata\testNonRxMonth_nonrxfinalreportoutput.csv", SheetUtilities.getNonRxReportRng(True)
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestMultiDeliveryTypeCount() ' Issue #4
    On Error GoTo TestFail
    
    MacroEntry AddressesSheet
    
    Dim testAddresses As Scripting.Dictionary
    Set testAddresses = New Scripting.Dictionary
    
    Dim deliveryRecord As RecordTuple
    Set deliveryRecord = New RecordTuple
    
    Dim nondeliveryRecord As RecordTuple
    Set nondeliveryRecord = New RecordTuple
    
    Dim deliveryAfterNonDeliveryRecord As RecordTuple
    Set deliveryAfterNonDeliveryRecord = New RecordTuple
    
    Dim deliveryInSameQuarterNonDeliveryRecord As RecordTuple
    Set deliveryInSameQuarterNonDeliveryRecord = New RecordTuple
    
    Dim deliveryBeforeNonDeliveryRecord As RecordTuple
    Set deliveryBeforeNonDeliveryRecord = New RecordTuple
    
    deliveryRecord.SetInCity InCityCode.ValidInCity
    deliveryRecord.guestID = "1"
    deliveryRecord.householdTotal = 1
    nondeliveryRecord.SetInCity InCityCode.ValidInCity
    nondeliveryRecord.guestID = "2"
    nondeliveryRecord.householdTotal = 2
    deliveryAfterNonDeliveryRecord.SetInCity InCityCode.ValidInCity
    deliveryAfterNonDeliveryRecord.guestID = "3"
    deliveryAfterNonDeliveryRecord.householdTotal = 3
    deliveryInSameQuarterNonDeliveryRecord.SetInCity InCityCode.ValidInCity
    deliveryInSameQuarterNonDeliveryRecord.guestID = "4"
    deliveryInSameQuarterNonDeliveryRecord.householdTotal = 4
    deliveryBeforeNonDeliveryRecord.SetInCity InCityCode.ValidInCity
    deliveryBeforeNonDeliveryRecord.guestID = "5"
    deliveryBeforeNonDeliveryRecord.householdTotal = 5
    
    
    deliveryRecord.AddVisit "7/8/2024", "Delivery Service"
    nondeliveryRecord.AddVisit "8/8/2024", "Food Service"
    nondeliveryRecord.AddVisit "11/11/2024", "Food Service"
    deliveryAfterNonDeliveryRecord.AddVisit "7/3/2025", "Food Service"
    deliveryAfterNonDeliveryRecord.AddVisit "8/3/2025", "Food Service"
    deliveryAfterNonDeliveryRecord.AddVisit "10/5/2025", "Delivery Service"
    deliveryInSameQuarterNonDeliveryRecord.AddVisit "11/3/2025", "Delivery Service"
    deliveryInSameQuarterNonDeliveryRecord.AddVisit "12/5/2025", "Food Service"
    deliveryBeforeNonDeliveryRecord.AddVisit "2/3/2025", "Delivery Service"
    deliveryBeforeNonDeliveryRecord.AddVisit "5/5/2025", "Food Service"
    
    testAddresses.Add deliveryRecord.key, deliveryRecord
    testAddresses.Add nondeliveryRecord.key, nondeliveryRecord
    testAddresses.Add deliveryAfterNonDeliveryRecord.key, deliveryAfterNonDeliveryRecord
    testAddresses.Add deliveryInSameQuarterNonDeliveryRecord.key, deliveryInSameQuarterNonDeliveryRecord
    testAddresses.Add deliveryBeforeNonDeliveryRecord.key, deliveryBeforeNonDeliveryRecord
    
    records.writeAddresses AddressesSheet.Name, testAddresses
    
    records.computeInterfaceTotals
    
    CompareSheetCSV Assert, InterfaceSheet.Name, ThisWorkbook.path & "\testdata\testMultiDeliveryTypeCount_multideliverytypetotalsoutput.csv", SheetUtilities.getInterfaceTotalsRng(numDoubleCountedAdditionalDeliveryType)
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestRxAddresses()
    On Error GoTo TestFail
    
    Fakes.MsgBox.Returns vbYes
    
    MacroEntry InterfaceSheet
    PasteInterfaceTestRecords ThisWorkbook.path & "\testdata\testrxaddresses.csv"
    
    InterfaceButtons.confirmAddRecords
    InterfaceButtons.confirmAttemptValidation
    
    MacroEntry RxSheet
    PasteRxTestRecordsCalculate ThisWorkbook.path & "\testdata\testrxrecords.csv"
    
    Assert.isTrue SheetUtilities.getRxMostRecentDateRng.value = "4/21/2025", "Most recent date is incorrect"
    Assert.isTrue SheetUtilities.getRxDiscardedIDsRng.value = "Amazon Rainforest,Needs Autocorrect,Apple Rich", "Discarded IDs are incorrect"
    
    CompareSheetCSV Assert, RxSheet.Name, ThisWorkbook.path & "\testdata\testrx_rxtotalsoutput.csv", SheetUtilities.getRxTotalsRng
    CompareSheetCSV Assert, RxReportSheet.Name, ThisWorkbook.path & "\testdata\testrx_rxfinalreportoutput.csv"
    
    InterfaceButtons.confirmDeleteRxRecords
    
    Assert.isTrue SheetUtilities.getRxMostRecentDateRng.value = "None", "Most recent date is not none"
    Assert.isTrue SheetUtilities.getRxDiscardedIDsRng.value = "None", "Discarded IDs is not none"
    Assert.isTrue RxSheet.UsedRange.rows.count = 10, "Not all data was deleted"
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestAllAddresses()
    On Error GoTo TestFail
       
    MacroEntry InterfaceSheet
    PasteInterfaceTestRecords ThisWorkbook.path & "\testdata\test1addresses.csv"
    
    MacroEntry InterfaceSheet
    addRecords
    
    Assert.isTrue SheetUtilities.getNonDeliveryTotalHeaderRng().value = "Non-Delivery: 18,7,9,Food", "Incorrect non delivery total header name"
    
    CompareSheetCSV Assert, AddressesSheet.Name, ThisWorkbook.path & "\testdata\test1addresses_addressesoutput.csv"
    CompareSheetCSV Assert, InterfaceSheet.Name, ThisWorkbook.path & "\testdata\test1addresses_nondeliverytotalsoutput.csv", getInterfaceTotalsRng(nonDelivery)
    CompareSheetCSV Assert, InterfaceSheet.Name, ThisWorkbook.path & "\testdata\test1addresses_countytotalsoutput.csv", getCountyRng
    CompareSheetCSV Assert, AutocorrectAddressesSheet.Name, ThisWorkbook.path & "\testdata\test1addresses_autocorrectoutput.csv"
    CompareSheetCSV Assert, DiscardsSheet.Name, ThisWorkbook.path & "\testdata\test1addresses_discardsoutput.csv"
    CompareSheetCSV Assert, AutocorrectedAddressesSheet.Name, ThisWorkbook.path & "\testdata\test1addresses_autocorrectedoutput.csv"
    
    PasteInterfaceTestRecords ThisWorkbook.path & "\testdata\test2extraaddresses.csv"

    MacroEntry InterfaceSheet
    addRecords

    CompareSheetCSV Assert, AddressesSheet.Name, ThisWorkbook.path & "\testdata\test2extraaddresses_addressesoutput.csv"
    CompareSheetCSV Assert, InterfaceSheet.Name, ThisWorkbook.path & "\testdata\test2extraaddresses_nondeliverytotalsoutput.csv", getInterfaceTotalsRng(nonDelivery)
    CompareSheetCSV Assert, AutocorrectAddressesSheet.Name, ThisWorkbook.path & "\testdata\test2extraaddresses_autocorrectoutput.csv"
    CompareSheetCSV Assert, DiscardsSheet.Name, ThisWorkbook.path & "\testdata\test2extraaddresses_discardsoutput.csv"
    CompareSheetCSV Assert, AutocorrectedAddressesSheet.Name, ThisWorkbook.path & "\testdata\test2extraaddresses_autocorrectedoutput.csv"

    PasteInterfaceTestRecords ThisWorkbook.path & "\testdata\test3autocorrectaddresses.csv"

    MacroEntry InterfaceSheet
    addRecords

    attemptValidation

    CompareSheetCSV Assert, AddressesSheet.Name, ThisWorkbook.path & "\testdata\test3autocorrectaddresses_addressesoutput.csv"
    CompareSheetCSV Assert, InterfaceSheet.Name, ThisWorkbook.path & "\testdata\test3autocorrectaddresses_nondeliverytotalsoutput.csv", getInterfaceTotalsRng(nonDelivery)
    CompareSheetCSV Assert, AutocorrectAddressesSheet.Name, ThisWorkbook.path & "\testdata\test3autocorrectaddresses_autocorrectoutput.csv"
    CompareSheetCSV Assert, DiscardsSheet.Name, ThisWorkbook.path & "\testdata\test3autocorrectaddresses_discardsoutput.csv"
    CompareSheetCSV Assert, AutocorrectedAddressesSheet.Name, ThisWorkbook.path & "\testdata\test3autocorrectaddresses_autocorrectedoutput.csv"

    Assert.isTrue Autocorrect.getRemainingRequests = 7980

    PasteInterfaceTestRecords ThisWorkbook.path & "\testdata\test4mergeaddresses.csv"

    MacroEntry InterfaceSheet
    addRecords

    CompareSheetCSV Assert, AddressesSheet.Name, ThisWorkbook.path & "\testdata\test4mergeaddresses_addressesoutput.csv"
    CompareSheetCSV Assert, InterfaceSheet.Name, ThisWorkbook.path & "\testdata\test4mergeaddresses_nondeliverytotalsoutput.csv", getInterfaceTotalsRng(nonDelivery)
    CompareSheetCSV Assert, DiscardsSheet.Name, ThisWorkbook.path & "\testdata\test4mergeaddresses_discardsoutput.csv"
    CompareSheetCSV Assert, AutocorrectedAddressesSheet.Name, ThisWorkbook.path & "\testdata\test4mergeaddresses_autocorrectedoutput.csv"

    Fakes.MsgBox.Returns vbYes
    
    MacroExit InterfaceSheet


    InterfaceButtons.confirmDiscardAll

    DiscardsSheet.Select
    Union(DiscardsSheet.Range("A2:A3"), _
          DiscardsSheet.Range("A5:A7"), _
          DiscardsSheet.Range("A10:A14")). _
          Select
    InterfaceButtons.confirmRestoreSelectedDiscard
    
    AutocorrectAddressesSheet.Select
    Union(AutocorrectAddressesSheet.Range("A6:A7"), _
          AutocorrectAddressesSheet.Range("A11")).Select
    InterfaceButtons.confirmDiscardSelected
    
    AddressesSheet.Select
    Union(AddressesSheet.Range("A3"), _
          AddressesSheet.Range("A7"), _
          AddressesSheet.Range("A14")).Select
    InterfaceButtons.confirmMoveAutocorrect

    
    AutocorrectAddressesSheet.Select

    AutocorrectAddressesSheet.Range("C5").value = "13-15 E Deer Park Dr"
    AutocorrectAddressesSheet.Range("D5").value = "Ste 202"
    AutocorrectAddressesSheet.Range("C6").value = "13-15 E Deer Park Dr"
    AutocorrectAddressesSheet.Range("D6").value = "Ste 102"
    AutocorrectAddressesSheet.Range("D2").value = "Unit 102"
    AutocorrectAddressesSheet.Range("D3").value = "Unit 102"
    AutocorrectAddressesSheet.Range("C8").value = "422 Girard St"
    AutocorrectAddressesSheet.Range("D8").value = "Apt 103"
    AutocorrectAddressesSheet.Range("D9").value = "Ste 100"
    AutocorrectAddressesSheet.Range("D11").value = "Apt 1"


    Union(AutocorrectAddressesSheet.Range("A4"), AutocorrectAddressesSheet.Range("A7")).Select
    InterfaceButtons.toggleUserVerified

    AutocorrectedAddressesSheet.Select
    Union(AutocorrectedAddressesSheet.Range("A4"), _
          AutocorrectedAddressesSheet.Range("A6")).Select
    InterfaceButtons.toggleUserVerifiedAutocorrected

    CompareSheetCSV Assert, AddressesSheet.Name, ThisWorkbook.path & "\testdata\test5usereditsaddresses_addressesoutput.csv"
    CompareSheetCSV Assert, InterfaceSheet.Name, ThisWorkbook.path & "\testdata\test5usereditsaddresses_nondeliverytotalsoutput.csv", getInterfaceTotalsRng(nonDelivery)
    CompareSheetCSV Assert, AutocorrectAddressesSheet.Name, ThisWorkbook.path & "\testdata\test5usereditsaddresses_autocorrectoutput.csv"
    CompareSheetCSV Assert, DiscardsSheet.Name, ThisWorkbook.path & "\testdata\test5usereditsaddresses_discardsoutput.csv"
    CompareSheetCSV Assert, AutocorrectedAddressesSheet.Name, ThisWorkbook.path & "\testdata\test5usereditsaddresses_autocorrectedoutput.csv"

    InterfaceButtons.confirmAttemptValidation
    
    InterfaceButtons.confirmGenerateNonRxReport

    CompareSheetCSV Assert, AddressesSheet.Name, ThisWorkbook.path & "\testdata\test6validateduseredits_addressesoutput.csv"
    CompareSheetCSV Assert, InterfaceSheet.Name, ThisWorkbook.path & "\testdata\test6validateduseredits_nondeliverytotalsoutput.csv", getInterfaceTotalsRng(nonDelivery)
    CompareSheetCSV Assert, InterfaceSheet.Name, ThisWorkbook.path & "\testdata\test6validateduseredits_countyoutput.csv", getCountyRng
    CompareSheetCSV Assert, AutocorrectedAddressesSheet.Name, ThisWorkbook.path & "\testdata\test6validateduseredits_autocorrectedoutput.csv"
    CompareSheetCSV Assert, NonRxReportSheet.Name, ThisWorkbook.path & "\testdata\test6validateduseredits_nonrxfinalreportoutput.csv"

    AddressesSheet.Select
    AddressesSheet.Range("A7").Select
    InterfaceButtons.confirmMoveAutocorrect
    
    InterfaceButtons.confirmDeleteAllVisitData

    CompareSheetCSV Assert, AddressesSheet.Name, ThisWorkbook.path & "\testdata\test7deletedata_addressesoutput.csv"
    CompareSheetCSV Assert, InterfaceSheet.Name, ThisWorkbook.path & "\testdata\test7deletedata_nondeliverytotalsoutput.csv", getInterfaceTotalsRng(nonDelivery)
    CompareSheetCSV Assert, InterfaceSheet.Name, ThisWorkbook.path & "\testdata\test7deletedata_countyoutput.csv", getCountyRng
    CompareSheetCSV Assert, AutocorrectAddressesSheet.Name, ThisWorkbook.path & "\testdata\test7deletedata_autocorrectoutput.csv"
    CompareSheetCSV Assert, DiscardsSheet.Name, ThisWorkbook.path & "\testdata\test7deletedata_discardsoutput.csv"
    CompareSheetCSV Assert, AutocorrectedAddressesSheet.Name, ThisWorkbook.path & "\testdata\test7deletedata_autocorrectedoutput.csv"
    CompareSheetCSV Assert, NonRxReportSheet.Name, ThisWorkbook.path & "\testdata\test7deletedata_nonrxfinalreportoutput.csv"

    AddressesSheet.Select
    AddressesSheet.Range("A2").Select
    InterfaceButtons.confirmMoveAutocorrect

    AutocorrectAddressesSheet.Select
    AutocorrectAddressesSheet.Range("A2").Select
    InterfaceButtons.toggleUserVerified
    InterfaceButtons.confirmDiscardSelected
    InterfaceButtons.confirmDiscardAll

    DiscardsSheet.Select
    DiscardsSheet.Range("A2").Select
    InterfaceButtons.confirmRestoreSelectedDiscard

    AutocorrectedAddressesSheet.Select
    AutocorrectedAddressesSheet.Range("A2").Select
    InterfaceButtons.toggleUserVerifiedAutocorrected


    MacroEntry InterfaceSheet
    InterfaceSheet.Select
    PasteInterfaceTestRecords ThisWorkbook.path & "\testdata\test4mergeaddresses.csv"
    MacroExit InterfaceSheet


    InterfaceButtons.confirmAddRecords
    InterfaceButtons.confirmAttemptValidation
    InterfaceButtons.confirmGenerateNonRxReport

    CompareSheetCSV Assert, AddressesSheet.Name, ThisWorkbook.path & "\testdata\test8noservices_addressesoutput.csv"
    CompareSheetCSV Assert, InterfaceSheet.Name, ThisWorkbook.path & "\testdata\test8noservices_nondeliverytotalsoutput.csv", getInterfaceTotalsRng(nonDelivery)
    CompareSheetCSV Assert, InterfaceSheet.Name, ThisWorkbook.path & "\testdata\test8noservices_countyoutput.csv", getCountyRng
    CompareSheetCSV Assert, AutocorrectAddressesSheet.Name, ThisWorkbook.path & "\testdata\test8noservices_autocorrectoutput.csv"
    CompareSheetCSV Assert, DiscardsSheet.Name, ThisWorkbook.path & "\testdata\test8noservices_discardsoutput.csv"
    CompareSheetCSV Assert, AutocorrectedAddressesSheet.Name, ThisWorkbook.path & "\testdata\test8noservices_autocorrectedoutput.csv"
    CompareSheetCSV Assert, NonRxReportSheet.Name, ThisWorkbook.path & "\testdata\test8noservices_nonrxfinalreportoutput.csv"
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestDelivery()
    On Error GoTo TestFail
    
    MacroEntry InterfaceSheet
    PasteInterfaceTestRecords ThisWorkbook.path & "\testdata\testdeliveryaddresses.csv"
    MacroExit InterfaceSheet
    
    Fakes.MsgBox.Returns vbYes
    InterfaceButtons.confirmAddRecords
    InterfaceButtons.confirmAttemptValidation
    InterfaceButtons.confirmGenerateNonRxReport
    
    Assert.isTrue SheetUtilities.getDeliveryTotalHeaderRng.value = "Delivery: Delivery,Food-Delivery", "Delivery service header doesn't match"
    
    CompareSheetCSV Assert, InterfaceSheet.Name, ThisWorkbook.path & "\testdata\testdeliveryaddresses_nondeliverytotalsoutput.csv", getInterfaceTotalsRng(nonDelivery)
    CompareSheetCSV Assert, InterfaceSheet.Name, ThisWorkbook.path & "\testdata\testdeliveryaddresses_deliverytotalsoutput.csv", getInterfaceTotalsRng(delivery)
    CompareSheetCSV Assert, NonRxReportSheet.Name, ThisWorkbook.path & "\testdata\testdeliveryaddresses_nonrxfinalreportoutput.csv"
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub TestNoHouseholdTotal()
    On Error GoTo TestFail
    
    MacroEntry InterfaceSheet
    PasteInterfaceTestRecords ThisWorkbook.path & "\testdata\testnohouseholdtotal.csv"
    MacroExit InterfaceSheet
    
    Fakes.MsgBox.Returns vbYes
    InterfaceButtons.confirmAddRecords
    InterfaceButtons.confirmAttemptValidation
    InterfaceButtons.confirmGenerateNonRxReport
    
    CompareSheetCSV Assert, AddressesSheet.Name, ThisWorkbook.path & "\testdata\testnohouseholdtotal_addressesoutput.csv"
    CompareSheetCSV Assert, InterfaceSheet.Name, ThisWorkbook.path & "\testdata\testnohouseholdtotal_nondeliverytotalsoutput.csv", getInterfaceTotalsRng(nonDelivery)
    CompareSheetCSV Assert, InterfaceSheet.Name, ThisWorkbook.path & "\testdata\testnohouseholdtotal_countyoutput.csv", getCountyRng
    CompareSheetCSV Assert, AutocorrectedAddressesSheet.Name, ThisWorkbook.path & "\testdata\testnohouseholdtotal_autocorrectedoutput.csv"
    CompareSheetCSV Assert, NonRxReportSheet.Name, ThisWorkbook.path & "\testdata\testnohouseholdtotal_nonrxfinalreportoutput.csv"
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestToggleUserVerifiedAutocorrect()
    On Error GoTo TestFail
    
    MacroEntry InterfaceSheet
    PasteInterfaceTestRecords ThisWorkbook.path & "\testdata\testtoggleuserverified.csv"
    
    MacroEntry InterfaceSheet
    addRecords
    attemptValidation
    
    MacroExit InterfaceSheet
    
    
    AutocorrectedAddressesSheet.Select
    AutocorrectedAddressesSheet.Range("A2:A3").Select
    InterfaceButtons.toggleUserVerifiedAutocorrected
    
    AutocorrectedAddressesSheet.Range("A2").Select
    InterfaceButtons.toggleUserVerifiedAutocorrected
    
    CompareSheetCSV Assert, AddressesSheet.Name, ThisWorkbook.path & "\testdata\testtoggleuserverified_addressesoutput.csv"
    CompareSheetCSV Assert, DiscardsSheet.Name, ThisWorkbook.path & "\testdata\testtoggleuserverified_discardsoutput.csv"
    CompareSheetCSV Assert, AutocorrectedAddressesSheet.Name, ThisWorkbook.path & "\testdata\testtoggleuserverified_autocorrectedoutput.csv"
    
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestUserMove()
    On Error GoTo TestFail
    
    MacroEntry InterfaceSheet
    PasteInterfaceTestRecords ThisWorkbook.path & "\testdata\testusermove.csv"
    MacroExit InterfaceSheet
    
    
    Fakes.MsgBox.Returns vbYes

    InterfaceButtons.confirmAddRecords
    
    InterfaceButtons.confirmAttemptValidation
    
    AddressesSheet.Select
    AddressesSheet.Range("A2").Select
    InterfaceButtons.confirmMoveAutocorrect

    AutocorrectAddressesSheet.Select
    AutocorrectAddressesSheet.Range("C2").value = "1 Grantchester Pl"
    AutocorrectAddressesSheet.Range("E2").value = "20878" ' should correct to 20877
    AutocorrectAddressesSheet.Range("A3").Select
    InterfaceButtons.toggleUserVerified
    
    InterfaceButtons.confirmAttemptValidation
    
    CompareSheetCSV Assert, AddressesSheet.Name, ThisWorkbook.path & "\testdata\testusermove_addressesoutput.csv"
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestCountyTotals()
    On Error GoTo TestFail

    MacroEntry InterfaceSheet
    PasteInterfaceTestRecords ThisWorkbook.path & "\testdata\testcounty.csv"

    MacroEntry InterfaceSheet
    addRecords
    
    CompareSheetCSV Assert, InterfaceSheet.Name, ThisWorkbook.path & "\testdata\testcounty_1added_totalsoutput.csv", getCountyRng
    
    attemptValidation
    
    Assert.isTrue SheetUtilities.getCountyTotalServicesRng.value = "Food,Delivery", "Included services are incorrect"
    
    CompareSheetCSV Assert, InterfaceSheet.Name, ThisWorkbook.path & "\testdata\testcounty_2validated_totalsoutput.csv", getCountyRng
    MacroExit InterfaceSheet
    
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub TestHandcorrected()
    On Error GoTo TestFail

    MacroEntry InterfaceSheet
    ' ThisWorkbook.path & "\testdata\testhandcorrected.csv"
    PasteInterfaceTestRecords ThisWorkbook.path & "\testdata\testdifficultaddresses.csv"
    
    MacroEntry InterfaceSheet
    addRecords
    attemptValidation
    MacroExit InterfaceSheet
    
    
'    CompareSheetCSV Assert, AddressesSheet.Name, ThisWorkbook.path & "\testdata\testhandcorrected_addressesoutput.csv"
'    CompareSheetCSV Assert, AutocorrectAddressesSheet.name, ThisWorkbook.path & "\testdata\testhandcorrected_autocorrectoutput.csv"
'    CompareSheetCSV Assert, DiscardsSheet.Name, ThisWorkbook.path & "\testdata\testhandcorrected_discardsoutput.csv"
'    CompareSheetCSV Assert, AutocorrectedAddressesSheet.name, ThisWorkbook.path & "\testdata\testhandcorrected_autocorrectedoutput.csv"
    CompareSheetCSV Assert, AddressesSheet.Name, ThisWorkbook.path & "\testdata\testdifficultaddresses_addressesoutput.csv"
    CompareSheetCSV Assert, AutocorrectAddressesSheet.Name, ThisWorkbook.path & "\testdata\testdifficultaddresses_autocorrectoutput.csv"
    CompareSheetCSV Assert, DiscardsSheet.Name, ThisWorkbook.path & "\testdata\testdifficultaddresses_discardsoutput.csv"
    CompareSheetCSV Assert, AutocorrectedAddressesSheet.Name, ThisWorkbook.path & "\testdata\testdifficultaddresses_autocorrectedoutput.csv"
    
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestOverwrite()
On Error GoTo TestFail
    
    MacroEntry InterfaceSheet
    PasteInterfaceTestRecords ThisWorkbook.path & "\testdata\testoverwrite.csv"
    
    MacroEntry InterfaceSheet
    addRecords
    attemptValidation
    
    Assert.isTrue getInterfaceMostRecentRng.value = "11/5/2023", "Most recent date is not 11/5/2023"
    
    PasteInterfaceTestRecords ThisWorkbook.path & "\testdata\testoverwrite_2.csv"
    
    MacroEntry InterfaceSheet
    addRecords
    
    Assert.isTrue getInterfaceMostRecentRng.value = "12/5/2023", "Most recent date is not 12/5/2023"
    
    MacroExit InterfaceSheet
    
    
    CompareSheetCSV Assert, AddressesSheet.Name, ThisWorkbook.path & "\testdata\testoverwrite_addressesoutput.csv"
    CompareSheetCSV Assert, AutocorrectAddressesSheet.Name, ThisWorkbook.path & "\testdata\testoverwrite_autocorrectoutput.csv"
    CompareSheetCSV Assert, DiscardsSheet.Name, ThisWorkbook.path & "\testdata\testoverwrite_discardsoutput.csv"
    CompareSheetCSV Assert, AutocorrectedAddressesSheet.Name, ThisWorkbook.path & "\testdata\testoverwrite_autocorrectedoutput.csv"
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub testSort()
    On Error GoTo TestFail
    
    MacroEntry AddressesSheet
    
    Dim wbook As Workbook
    Set wbook = Workbooks.Open(fileName:=ThisWorkbook.path & "\testdata\testsortaddresses.csv")
    
    wbook.Worksheets.[_Default](1).UsedRange.Copy
    
    AddressesSheet.Range("A1").PasteSpecial xlPasteValues
    
    SheetUtilities.SortSheet AddressesSheet.Name
    CompareSheetCSV Assert, AddressesSheet.Name, ThisWorkbook.path & "\testdata\testsortaddresses_valid_addressesoutput.csv"
    
    SheetUtilities.ClearSheet AddressesSheet.Name
    
    wbook.Worksheets.[_Default](1).UsedRange.Copy
    DiscardsSheet.Range("A1").PasteSpecial xlPasteValues
    
    setClipboardToBlankLine
    
    wbook.Close
    
    SheetUtilities.SortSheet DiscardsSheet.Name
    MacroExit InterfaceSheet
    
    
    CompareSheetCSV Assert, DiscardsSheet.Name, ThisWorkbook.path & "\testdata\testsortaddresses_raw_discardsoutput.csv"
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


