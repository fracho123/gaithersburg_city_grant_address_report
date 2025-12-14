Attribute VB_Name = "InterfaceButtons"
Option Explicit

'@Folder("City_Grant_Address_Report.src")

' Returns false if user has a filter enabled
Public Function MacroEntry(ByVal wsheetToReturn As Worksheet) As Boolean
    
    Dim i As Long
    For i = 1 To ThisWorkbook.Sheets.count
        Dim wsheet As Worksheet
        Set wsheet = ThisWorkbook.Sheets.[_Default](i)
        
        If wsheet.FilterMode = True Then
            MsgBox "Disable filter on " & wsheet.Name & " and try again"
            MacroEntry = False
            Exit Function
        End If
        
        ThisWorkbook.Sheets.[_Default](i).AutoFilterMode = False
        
        wsheet.Unprotect
    Next
    
    ' NOTE If program encountered an error, status bar won't be reset, so reset it now
    Application.StatusBar = False
    AutocorrectAddressesSheet.macroIsRunning = True
    wsheetToReturn.Activate
    MacroEntry = True
End Function

' NOTE change AutocorrectAddressesSheet when this changes
Public Sub MacroExit(ByVal wsheetToReturn As Worksheet)
    Dim i As Long
    For i = 1 To ThisWorkbook.Sheets.count
        Dim wsheet As Worksheet
        Set wsheet = ThisWorkbook.Sheets.[_Default](i)
        
        wsheet.Unprotect
        wsheet.AutoFilterMode = False
        
        Select Case wsheet.Name
            Case NonRxReportSheet.Name, RxReportSheet.Name
                wsheet.UsedRange.Offset(1, 0).AutoFilter
            ' Rubberduck bug
            '@Ignore UnreachableCase
            Case AddressesSheet.Name, AutocorrectAddressesSheet.Name, AutocorrectedAddressesSheet.Name, DiscardsSheet.Name
                wsheet.UsedRange.AutoFilter
        End Select
        
        wsheet.Protect AllowFormattingColumns:=True, AllowFormattingRows:=True, AllowSorting:=True, AllowFiltering:=True
    Next
    
    Application.StatusBar = False
    wsheetToReturn.Activate
    AutocorrectAddressesSheet.macroIsRunning = False
End Sub

' Returns Nothing if error occurred
Private Function getUniqueSelection(ByVal returnRows As Boolean, ByVal min As Long) As Collection
    Dim uniques As Collection
    Set uniques = New Collection
    
    Dim dict As Scripting.Dictionary
    Set dict = New Scripting.Dictionary
    
    Dim selections As Range
    ' xlCellTypeVisible in case a filter is applied
    If returnRows Then
        Set selections = Selection.SpecialCells(xlCellTypeVisible).rows
    Else
        Set selections = Selection.SpecialCells(xlCellTypeVisible).Columns
    End If
    
    Dim value As Variant
    For Each value In selections
        If returnRows Then
            If value.row < min Then
                MsgBox "Invalid Selection"
                Set getUniqueSelection = Nothing
                Exit Function
            End If
            dict.Item(value.row) = Empty
        Else
            If value.column < min Then
                MsgBox "Invalid Selection"
                Set getUniqueSelection = Nothing
                Exit Function
            End If
            dict.Item(value.column) = Empty
        End If
    Next value
    
    For Each value In dict.Keys()
        uniques.Add value
    Next value
    
    Set getUniqueSelection = uniques
End Function

Private Sub PasteRecords(ByVal sheet As Worksheet)
    ' NOTE not using MacroEntry here, it disables PasteSpecial for some reason
    ' Using Allow Edit Ranges
    
    sheet.Activate
    Application.ScreenUpdating = False
    
    getBlankRow(sheet.Name).Cells.Item(1, 1).PasteSpecial Paste:=xlPasteValues
    
    sheet.Cells.Item(1, 1).Select
    Application.ScreenUpdating = True
End Sub


'@EntryPoint
Public Sub PasteInterfaceRecords()
    InterfaceSheet.Activate
    PasteRecords InterfaceSheet
    MacroExit InterfaceSheet
End Sub

'@EntryPoint
Public Sub confirmPasteRxRecordsCalculate()
    Dim confirmResponse As VbMsgBoxResult
    confirmResponse = MsgBox("Are you sure you wish to paste Rx records and generate the RX report?", vbYesNo + vbQuestion, "Confirmation")
    If confirmResponse = vbNo Then
        Exit Sub
    End If
    
    PasteRecords RxSheet
    
    ' This needs to go AFTER PasteRecords for PasteSpecial to work
    If Not MacroEntry(RxSheet) Then Exit Sub
    
    Dim addresses As Scripting.Dictionary
    Set addresses = records.loadAddresses(AddressesSheet.Name)
    
    Dim discards As Scripting.Dictionary
    Set discards = records.loadAddresses(DiscardsSheet.Name)
    
    Set addresses = MergeDicts(addresses, discards)
    
    Dim out As records.ComputedRx
    out = records.computeRxTotals(addresses)
    out.totals.output
    
    GenerateReport.generateRxReport out.records, addresses
    
    MacroExit RxSheet
End Sub

'@EntryPoint
Public Sub confirmAddRecords()
    Dim confirmResponse As VbMsgBoxResult
    confirmResponse = MsgBox("Are you sure you wish to add records?", vbYesNo + vbQuestion, "Confirmation")
    If confirmResponse = vbNo Then
        Exit Sub
    End If
    
    If Not MacroEntry(ThisWorkbook.ActiveSheet) Then Exit Sub
    
    records.addRecords
    
    MacroExit ThisWorkbook.ActiveSheet
End Sub

'@EntryPoint
Public Sub confirmAttemptValidation()
    ' NOTE this line must be here before calling getRemainingRequests()
    ' unable to test this with Rubberduck due to MsgBox being a Fake
    If Not MacroEntry(ThisWorkbook.ActiveSheet) Then Exit Sub
    
    Dim confirmResponse As VbMsgBoxResult
    confirmResponse = MsgBox("Are you sure you wish to attempt validation? You have " & _
                              CStr(getRemainingRequests()) & " remaining.", _
                              vbYesNo + vbQuestion, "Confirmation")
    If confirmResponse = vbNo Then
        Exit Sub
    End If
    
    Autocorrect.attemptValidation
    
    MacroExit ThisWorkbook.ActiveSheet
End Sub

'@EntryPoint
Public Sub confirmGenerateNonRxReport()
    Dim confirmResponse As VbMsgBoxResult
    confirmResponse = MsgBox("Are you sure you wish to generate the Non-Rx report?", vbYesNo + vbQuestion, "Confirmation")
    If confirmResponse = vbNo Then
        Exit Sub
    End If
    
    If Not MacroEntry(ThisWorkbook.ActiveSheet) Then Exit Sub
    
    GenerateReport.generateNonRxReport
    
    MacroExit ThisWorkbook.ActiveSheet
End Sub

'@EntryPoint
Public Sub confirmDeleteRxRecords()
    Dim confirmResponse As VbMsgBoxResult
    confirmResponse = MsgBox("Are you sure you wish to delete all rx records?", vbYesNo + vbQuestion, "Confirmation")
    If confirmResponse = vbNo Then
        Exit Sub
    End If
    
    If Not MacroEntry(ThisWorkbook.ActiveSheet) Then Exit Sub
    
    SheetUtilities.ClearRxTotals
    
    MacroExit ThisWorkbook.ActiveSheet
End Sub

'@EntryPoint
Public Sub confirmDeleteAllVisitData()
    Dim confirmResponse As VbMsgBoxResult
    confirmResponse = MsgBox("Are you sure you wish to delete all visit data?", vbYesNo + vbQuestion, "Confirmation")
    If confirmResponse = vbNo Then
        Exit Sub
    End If
    
    If Not MacroEntry(ThisWorkbook.ActiveSheet) Then Exit Sub
    
    SheetUtilities.getInterfaceMostRecentRng.value = vbNullString
    SheetUtilities.ClearInterfaceTotals
    SheetUtilities.getCountyRng.value = 0
    SheetUtilities.getNonRxReportRng.Clear
    SheetUtilities.getRxReportRng.Clear
    SheetUtilities.getAddressVisitDataRng(AddressesSheet.Name).Clear
    SheetUtilities.getRng(AddressesSheet.Name, "A2", "A2").Offset(0, SheetUtilities.firstServiceColumn - 2).value = "{}"
    SheetUtilities.getAddressVisitDataRng(AutocorrectAddressesSheet.Name).Clear
    SheetUtilities.getRng(AutocorrectAddressesSheet.Name, "A2", "A2").Offset(0, SheetUtilities.firstServiceColumn - 2).value = "{}"
    SheetUtilities.getAddressVisitDataRng(DiscardsSheet.Name).Clear
    SheetUtilities.getRng(DiscardsSheet.Name, "A2", "A2").Offset(0, SheetUtilities.firstServiceColumn - 2).value = "{}"
    SheetUtilities.getAddressVisitDataRng(AutocorrectedAddressesSheet.Name).Clear
    SheetUtilities.getRng(AutocorrectedAddressesSheet.Name, "A2", "A2").Offset(0, SheetUtilities.firstServiceColumn - 2).value = "{}"

    MacroExit ThisWorkbook.ActiveSheet
End Sub

'Public Sub confirmDeleteService()
'    Dim columns As Collection
'    Set columns = getUniqueSelection(False, SheetUtilities.firstServiceColumn)
'    If columns Is Nothing Then Exit Sub
'
'    Dim confirmResponse As VbMsgBoxResult
'    confirmResponse = MsgBox("Are you sure you wish to delete the selected service(s)?", vbYesNo + vbQuestion, "Confirmation")
'    If confirmResponse = vbNo Then
'        Exit Sub
'    End If
'
'    If Not MacroEntry(ThisWorkbook.ActiveSheet) Then Exit Sub
'
'
'    Dim addressServices() As String
'    addressServices = SheetUtilities.loadServiceNames(AddressesSheet.Name)
'
'    Dim autocorrectedServices() As String
'    autocorrectedServices = SheetUtilities.loadServiceNames(AutocorrectedAddressesSheet.Name)
'
'    Dim addressColsToDelete As Range
'    Dim autocorrectedColsToDelete As Range
'
'    Dim column As Variant
'    For Each column In columns
'        If addressColsToDelete Is Nothing Then
'            Set addressColsToDelete = _
'                AddressesSheet.columns.Item(column)
'        Else
'            Set addressColsToDelete = Union(addressColsToDelete, _
'                AddressesSheet.columns.Item(column))
'        End If
'
'        Dim service As String
'        service = addressServices(column - SheetUtilities.firstServiceColumn)
'
'        Dim i As Long
'        i = 0
'        Do While i <= UBound(autocorrectedServices)
'            If service = autocorrectedServices(i) Then
'                If autocorrectedColsToDelete Is Nothing Then
'                    Set autocorrectedColsToDelete = _
'                        AutocorrectedAddressesSheet _
'                        .columns.Item(i + SheetUtilities.firstServiceColumn)
'                Else
'                    Set autocorrectedColsToDelete = Union(autocorrectedColsToDelete, _
'                            AutocorrectedAddressesSheet _
'                            .columns.Item(i + SheetUtilities.firstServiceColumn))
'                End If
'                Exit Do
'            End If
'            i = i + 1
'        Loop
'    Next column
'
'    addressColsToDelete.EntireColumn.Delete
'
'    If Not autocorrectedColsToDelete Is Nothing Then
'        autocorrectedColsToDelete.EntireColumn.Delete
'    End If
'
'    SheetUtilities.getFinalReportRng.Clear
'    Records.computeInterfaceTotals
'    Records.computeCountyTotals
'
'    MacroExit ThisWorkbook.ActiveSheet
'End Sub

'@EntryPoint
Public Sub confirmDiscardAll()
    Dim confirmResponse As VbMsgBoxResult
    confirmResponse = MsgBox("Are you sure you wish to discard all records?", vbYesNo + vbQuestion, "Confirmation")
    If confirmResponse = vbNo Then
        Exit Sub
    End If
    
    If Not MacroEntry(ThisWorkbook.ActiveSheet) Then Exit Sub
    
    Dim Autocorrect As Scripting.Dictionary
    Set Autocorrect = records.loadAddresses(AutocorrectAddressesSheet.Name)
    
    Dim key As Variant
    For Each key In Autocorrect.Keys()
        records.writeAddress DiscardsSheet.Name, Autocorrect.Item(key)
    Next key
    
    SheetUtilities.ClearSheet AutocorrectAddressesSheet.Name
    SheetUtilities.SortSheet DiscardsSheet.Name
    
    MacroExit ThisWorkbook.ActiveSheet
End Sub

Private Function findRow(ByVal sheetName As String, ByVal key As String) As Range
    Set findRow = ThisWorkbook.Worksheets.[_Default](sheetName).Columns(SheetUtilities.keyColumn). _
                            Find(What:=key, LookIn:=xlValues, LookAt:=xlWhole)
End Function

Private Sub moveSelectedRows(ByVal sourceSheet As String, ByVal destSheet As String, _
                             ByVal removeFromAutocorrected As Boolean)
    Dim rows As Collection
    Set rows = getUniqueSelection(True, 2)
    If rows Is Nothing Then
        Exit Sub
    End If
    
    Dim confirmResponse As VbMsgBoxResult
    confirmResponse = MsgBox("Are you sure you wish to move the selected record(s) from " & _
                             sourceSheet & " to " & destSheet & "?", vbYesNo + vbQuestion, "Confirmation")
    If confirmResponse = vbNo Then
        Exit Sub
    End If
    
    Dim movedRecords As Collection
    Set movedRecords = New Collection
    
    Dim rowsToDelete As Range
    Dim row As Variant
    For Each row In rows
        Dim currentRowRng As Range
        Set currentRowRng = ThisWorkbook.Worksheets.[_Default](sourceSheet).Range("A" & row)
        Dim record As RecordTuple
        Set record = records.loadRecordFromSheet(currentRowRng)
        
        records.writeAddress destSheet, record
        movedRecords.Add record
        
        If rowsToDelete Is Nothing Then
            Set rowsToDelete = currentRowRng
        Else
            Set rowsToDelete = Union(currentRowRng, rowsToDelete)
        End If
    Next row
    
    rowsToDelete.EntireRow.Delete
    SheetUtilities.ClearEmptyServices sourceSheet
    
    ActiveSheet.Cells(1, 1).Select
    SheetUtilities.SortSheet destSheet
    
    If (Not removeFromAutocorrected) Then Exit Sub
       
    Dim movedRecord As Variant
    For Each movedRecord In movedRecords
        Dim foundCell As Range
        Set foundCell = findRow(AutocorrectedAddressesSheet.Name, movedRecord.key)
        If Not foundCell Is Nothing Then
            foundCell.EntireRow.Delete
        End If
    Next movedRecord
End Sub

'@EntryPoint
Public Sub confirmDiscardSelected()
    If Not MacroEntry(ThisWorkbook.ActiveSheet) Then Exit Sub
    
    moveSelectedRows AutocorrectAddressesSheet.Name, DiscardsSheet.Name, False
    
    MacroExit ThisWorkbook.ActiveSheet
End Sub

'@EntryPoint
Public Sub confirmRestoreSelectedDiscard()
    If Not MacroEntry(ThisWorkbook.ActiveSheet) Then Exit Sub
    
    moveSelectedRows DiscardsSheet.Name, AutocorrectAddressesSheet.Name, True
    
    MacroExit ThisWorkbook.ActiveSheet
End Sub

'@EntryPoint
Public Sub confirmMoveAutocorrect()
    If Not MacroEntry(ThisWorkbook.ActiveSheet) Then Exit Sub
    
    moveSelectedRows AddressesSheet.Name, AutocorrectAddressesSheet.Name, True
    SheetUtilities.getRxReportRng.Clear
    SheetUtilities.getNonRxReportRng.Clear
    records.computeInterfaceTotals
    
    MacroExit ThisWorkbook.ActiveSheet
End Sub

'@EntryPoint
Public Sub toggleUserVerified()
    If Not MacroEntry(ThisWorkbook.ActiveSheet) Then Exit Sub
    
    Dim rows As Collection
    Set rows = getUniqueSelection(True, 2)
    
    If rows Is Nothing Then Exit Sub
    
    Dim row As Variant
    For Each row In rows
        AutocorrectAddressesSheet.Cells.Item(row, 2).value = _
            Not AutocorrectAddressesSheet.Cells.Item(row, 2).value
    Next row
    
    MacroExit ThisWorkbook.ActiveSheet
End Sub

'@EntryPoint
Public Sub toggleUserVerifiedAutocorrected()
    If Not MacroEntry(ThisWorkbook.ActiveSheet) Then Exit Sub
    
    Dim rows As Collection
    Set rows = getUniqueSelection(True, 2)
    
    If rows Is Nothing Then Exit Sub
    
    
    Dim row As Variant
    For Each row In rows
        Dim currentRowRng As Range
        Set currentRowRng = AutocorrectedAddressesSheet.Range("A" & row)
        
        currentRowRng.Cells.Item(1, 2).value = Not currentRowRng.Cells.Item(1, 2).value
        
        Dim key As String
        key = currentRowRng.Cells.Item(1, SheetUtilities.keyColumn)
        
        Dim foundCell As Range
        Set foundCell = findRow(AddressesSheet.Name, key)
        
        If Not foundCell Is Nothing Then
            AddressesSheet.rows.Item(foundCell.row).Cells.Item(1, 2) = _
                                        Not AddressesSheet.rows.Item(foundCell.row).Cells.Item(1, 2)
        End If
        
        Set foundCell = DiscardsSheet.Columns.Item(SheetUtilities.keyColumn). _
                            Find(What:=key, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not foundCell Is Nothing Then
            DiscardsSheet.rows.Item(foundCell.row).Cells.Item(1, 2) = _
                                        Not DiscardsSheet.rows.Item(foundCell.row).Cells.Item(1, 2)
        End If
        
    Next row
    
    MacroExit ThisWorkbook.ActiveSheet
End Sub

'@EntryPoint
Public Sub ImportRecords()
    If Not MacroEntry(ThisWorkbook.ActiveSheet) Then Exit Sub
    
    Dim wbook As Workbook
    Set wbook = FileUtilities.getWorkbook()
    
    If wbook Is Nothing Then
        Exit Sub
    End If
    
    Dim versionNum As String
    versionNum = getVersionNum()
    
    SheetUtilities.ClearAll
    
    ' Copy all sheets except for Interface and Final Report
    wbook.Worksheets.[_Default](AddressesSheet.Name).UsedRange.Copy
    AddressesSheet.Range("A1").PasteSpecial xlPasteValues
    
    wbook.Worksheets.[_Default](AutocorrectAddressesSheet.Name).UsedRange.Copy
    AutocorrectAddressesSheet.Range("A1").PasteSpecial xlPasteValues
    
    wbook.Worksheets.[_Default](DiscardsSheet.Name).UsedRange.Copy
    DiscardsSheet.Range("A1").PasteSpecial xlPasteValues
    
    wbook.Worksheets.[_Default](AutocorrectedAddressesSheet.Name).UsedRange.Copy
    AutocorrectedAddressesSheet.Range("A1").PasteSpecial xlPasteValues
    
    setClipboardToBlankLine
    
    records.computeInterfaceTotals
    records.computeCountyTotals
    
    InterfaceSheet.Range("A1").value = versionNum
    ' v3.10 renames Interface to Home, so refer by index instead
    getInterfaceMostRecentRng.value = wbook.Worksheets.[_Default](1).Range(SheetUtilities.mostRecentDateCell).value
    
    wbook.Close
    
    MacroExit ThisWorkbook.ActiveSheet
End Sub


'@EntryPoint
Public Sub OpenGaithersburgStreets()
    ThisWorkbook.FollowHyperlink address:="https://maps.gaithersburgmd.gov/arcgis/rest/services/layers/GaithersburgCityAddresses/MapServer/0/query?where=Core_Address+LIKE+%27%25%27&text=&objectIds=&time=&geometry=&geometryType=esriGeometryEnvelope&inSR=&spatialRel=esriSpatialRelIntersects&relationParam=&outFields=Road_Name%2CRoad_Type&returnGeometry=false&returnTrueCurves=false&maxAllowableOffset=&geometryPrecision=&outSR=&returnIdsOnly=false&returnCountOnly=false&orderByFields=Road_Name&groupByFieldsForStatistics=&outStatistics=&returnZ=false&returnM=false&gdbVersion=&returnDistinctValues=true&resultOffset=&resultRecordCount=&queryByDistance=&returnExtentsOnly=false&datumTransformation=&parameterValues=&rangeValues=&f=html"
End Sub

'@EntryPoint
Public Sub CopyAndOpenCountyTotalsSite()
    Dim values As Range
    Set values = ActiveSheet.rows(Selection.row)
    
    
    ' Use ~ as quote, replace with """" later
    Dim code As Variant
    code = vbNullString
    code = code & "localStorage.setItem('AirtableLocalPersister.formPageElementSavedFormDataByElementId.appSihm9stog1ZEFn.paglEufgcXNk939cP.peljJ5sUBYayWAHwc','{~lastTouchedTime~:~"
    code = code & Format$(Date - 1, "yyyy-mm-dd") & "T21:25:49.955Z" & "~,"
    code = code & "~cellValueByColumnId~:{"
    
    ' ===== Report type (new field + new sel option) =====
    code = code & "~fldT5jmhPTnnuef2s~: ~selXGhScETeagB5Bm~,"
    
    ' ===== Zip Codes (new field + new record IDs) =====
    code = code & "~fldGCWr86vBijnH0L~: ["
    code = code & "{~foreignRowId~:~reckFHeR3xkx9s0WU~,~foreignRowDisplayName~:~20861~},"
    code = code & "{~foreignRowId~:~recaDALUl7ErnXJWk~,~foreignRowDisplayName~:~20906~},"
    code = code & "{~foreignRowId~:~rec6PYak2TL1s9BtO~,~foreignRowDisplayName~:~20916~},"
    code = code & "{~foreignRowId~:~recyWZFWxckLqCWf5~,~foreignRowDisplayName~:~20839~},"
    code = code & "{~foreignRowId~:~recjTVJUZLYq01qjy~,~foreignRowDisplayName~:~20838~},"
    code = code & "{~foreignRowId~:~recPjNimCmfmw5DQR~,~foreignRowDisplayName~:~20813~},"
    code = code & "{~foreignRowId~:~recom6UQpScTzf3YR~,~foreignRowDisplayName~:~20814~},"
    code = code & "{~foreignRowId~:~rec8Gp0LYJfZkdi7s~,~foreignRowDisplayName~:~20815~},"
    code = code & "{~foreignRowId~:~recOITGamcrifX2hT~,~foreignRowDisplayName~:~20816~},"
    code = code & "{~foreignRowId~:~reccMJqmX4WD0DrmL~,~foreignRowDisplayName~:~20817~},"
    code = code & "{~foreignRowId~:~recGhpUZE2AfGsbw4~,~foreignRowDisplayName~:~20824~},"
    code = code & "{~foreignRowId~:~recP8epKliF8VGiX7~,~foreignRowDisplayName~:~20825~},"
    code = code & "{~foreignRowId~:~reccWPZOY9tDt486f~,~foreignRowDisplayName~:~20827~},"
    code = code & "{~foreignRowId~:~recVA5KbCdrV30fz7~,~foreignRowDisplayName~:~20852~},"
    code = code & "{~foreignRowId~:~recwvEx7G2QCyoIuJ~,~foreignRowDisplayName~:~20841~},"
    code = code & "{~foreignRowId~:~recCWObQeUqsB69ma~,~foreignRowDisplayName~:~20862~},"
    code = code & "{~foreignRowId~:~recOVP6McjqisdNJx~,~foreignRowDisplayName~:~20866~},"
    code = code & "{~foreignRowId~:~recARD7UxWbsc5HuB~,~foreignRowDisplayName~:~20818~},"
    code = code & "{~foreignRowId~:~recK2euqytAvgm1tA~,~foreignRowDisplayName~:~20871~},"
    code = code & "{~foreignRowId~:~rec9WVSX0bAk01E7t~,~foreignRowDisplayName~:~20904~},"
    code = code & "{~foreignRowId~:~recR9f8p2TfPwnqMO~,~foreignRowDisplayName~:~20905~},"
    code = code & "{~foreignRowId~:~recwDTWnwWKuU2MRt~,~foreignRowDisplayName~:~20914~},"
    code = code & "{~foreignRowId~:~recAAYzHQsMK4OzRK~,~foreignRowDisplayName~:~20872~},"
    code = code & "{~foreignRowId~:~recG42IfSpAxRR1nk~,~foreignRowDisplayName~:~20874~},"
    code = code & "{~foreignRowId~:~rec6M3rSPnj18UzfV~,~foreignRowDisplayName~:~20878~},"
    code = code & "{~foreignRowId~:~recpyPAzM06jcG5Zn~,~foreignRowDisplayName~:~20855~},"
    code = code & "{~foreignRowId~:~recA4B1fUjDMks5mw~,~foreignRowDisplayName~:~20842~},"
    code = code & "{~foreignRowId~:~rec4MSVC2Bo3Mmkp9~,~foreignRowDisplayName~:~20877~},"
    code = code & "{~foreignRowId~:~recvZ0dWIQGVNUqx5~,~foreignRowDisplayName~:~20879~},"
    code = code & "{~foreignRowId~:~recxtVOehpwplKyWz~,~foreignRowDisplayName~:~20882~},"
    code = code & "{~foreignRowId~:~recUONkgEgwHRGptt~,~foreignRowDisplayName~:~20883~},"
    code = code & "{~foreignRowId~:~recCfMSOTR6uskvf5~,~foreignRowDisplayName~:~20884~},"
    code = code & "{~foreignRowId~:~recWrW5Ut7I9GioA7~,~foreignRowDisplayName~:~20885~},"
    code = code & "{~foreignRowId~:~recLCc2eWVfaHgkDY~,~foreignRowDisplayName~:~20886~},"
    code = code & "{~foreignRowId~:~recpUmnEzxxh6t3ke~,~foreignRowDisplayName~:~20898~},"
    code = code & "{~foreignRowId~:~recsDaO7odalfJWJF~,~foreignRowDisplayName~:~20896~},"
    code = code & "{~foreignRowId~:~rece3FB3HPN8SxqTD~,~foreignRowDisplayName~:~20875~},"
    code = code & "{~foreignRowId~:~reclVkFLkMot3zCRL~,~foreignRowDisplayName~:~20876~},"
    code = code & "{~foreignRowId~:~recEWJcGpBKzl8Ypl~,~foreignRowDisplayName~:~20812~},"
    code = code & "{~foreignRowId~:~recBYvzgcjrcychGv~,~foreignRowDisplayName~:~20891~},"
    code = code & "{~foreignRowId~:~recv1xFKz1DoNSWNk~,~foreignRowDisplayName~:~20895~},"
    code = code & "{~foreignRowId~:~recI5YsdFoc0oU71H~,~foreignRowDisplayName~:~20830~},"
    code = code & "{~foreignRowId~:~recpwgU2teHwhjq9z~,~foreignRowDisplayName~:~20832~},"
    code = code & "{~foreignRowId~:~recBqYUvTjtzt0932~,~foreignRowDisplayName~:~20837~},"
    code = code & "{~foreignRowId~:~recEuaQ8eOGeXc9n6~,~foreignRowDisplayName~:~20854~},"
    code = code & "{~foreignRowId~:~recyDpULBxTpQQbLY~,~foreignRowDisplayName~:~20859~},"
    code = code & "{~foreignRowId~:~recA6NNCdFUa3Qzzt~,~foreignRowDisplayName~:~20847~},"
    code = code & "{~foreignRowId~:~recElDh97gL2blpzV~,~foreignRowDisplayName~:~20848~},"
    code = code & "{~foreignRowId~:~recYzfDIwQE52NW26~,~foreignRowDisplayName~:~20849~},"
    code = code & "{~foreignRowId~:~recN4BytSyXXTcoGF~,~foreignRowDisplayName~:~20850~},"
    code = code & "{~foreignRowId~:~recJZvPGK9NxOpiy5~,~foreignRowDisplayName~:~20851~},"
    code = code & "{~foreignRowId~:~reclV98twecZLR1Fp~,~foreignRowDisplayName~:~20853~},"
    code = code & "{~foreignRowId~:~recxTzSMQAVBSPznH~,~foreignRowDisplayName~:~20860~},"
    code = code & "{~foreignRowId~:~recEQj9TmIMUDngjv~,~foreignRowDisplayName~:~20868~},"
    code = code & "{~foreignRowId~:~recKmi1zMQ546CNZi~,~foreignRowDisplayName~:~20912~},"
    code = code & "{~foreignRowId~:~recVXx8P4AbYpn8wz~,~foreignRowDisplayName~:~20913~},"
    code = code & "{~foreignRowId~:~recx6miNOzrcHrQxu~,~foreignRowDisplayName~:~20901~},"
    code = code & "{~foreignRowId~:~recj0QdHfH80Jn39Q~,~foreignRowDisplayName~:~20902~},"
    code = code & "{~foreignRowId~:~recEu7dTAkdCAegOz~,~foreignRowDisplayName~:~20903~},"
    code = code & "{~foreignRowId~:~reclA4inZtkmJv1Z9~,~foreignRowDisplayName~:~20907~},"
    code = code & "{~foreignRowId~:~recTrFAiJ5ZKhS8gw~,~foreignRowDisplayName~:~20908~},"
    code = code & "{~foreignRowId~:~rec3eEVpND2CHerc0~,~foreignRowDisplayName~:~20910~},"
    code = code & "{~foreignRowId~:~recV1lf0PY2s9jYcM~,~foreignRowDisplayName~:~20911~},"
    code = code & "{~foreignRowId~:~rec4W0bXx1GuK7twl~,~foreignRowDisplayName~:~20915~},"
    code = code & "{~foreignRowId~:~rectKTrh9I6Dx8sVV~,~foreignRowDisplayName~:~20918~},"
    code = code & "{~foreignRowId~:~recMccNHBU1ZorCnf~,~foreignRowDisplayName~:~20880~}"
    code = code & "],"
    
    
    ' ===== Now the numeric fields (same Excel cell math, NEW fld IDs) =====
    code = code & "~fldZN0BOwVtzwg515~:" & values.Cells.Item(1, 2).value & ","
    code = code & "~fldwokr57DoxDn3EX~:" & values.Cells.Item(1, 3).value & ","

    code = code & "~fldVdeLxrqXtv1HwT~:" & values.Cells.Item(1, 9).value & "," ' 20861
    code = code & "~fldOitAqqQmE2zASo~:" & CStr(values.Cells.Item(1, 10).value + values.Cells.Item(1, 84).value) & "," ' 20906
    code = code & "~fld4uG0pREOvMk2hP~:" & CStr(values.Cells.Item(1, 11).value + values.Cells.Item(1, 93).value) & "," ' 20916
    code = code & "~fldhLnXYWyF6rnz0w~:" & values.Cells.Item(1, 12).value & "," ' 20839
    code = code & "~fld0XsG5ezAhm7LPi~:" & values.Cells.Item(1, 13).value & "," ' 20838
    code = code & "~fldoAtLzSc1G6l1e3~:" & values.Cells.Item(1, 14).value & "," ' 20813
    code = code & "~fldn8V1aQ4x3sUhC2~:" & values.Cells.Item(1, 15).value & "," ' 20814
    code = code & "~fldGko07MPC9Cw4Fq~:" & CStr(values.Cells.Item(1, 16).value + values.Cells.Item(1, 27).value) & "," ' 20815
    code = code & "~fldhqlkOxbNLcSLmz~:" & values.Cells.Item(1, 17).value & "," ' 20816
    code = code & "~fldBzVx8fEXHWx86i~:" & values.Cells.Item(1, 18).value & "," ' 20817
    code = code & "~fldasKigxU7L0HaYE~:" & values.Cells.Item(1, 19).value & "," ' 20824
    code = code & "~fldACPmds1luOIWfc~:" & CStr(values.Cells.Item(1, 20).value + values.Cells.Item(1, 28).value) & "," ' 20825
    code = code & "~fld1tyUu8DxLUvwtS~:" & values.Cells.Item(1, 21).value & "," ' 20827
    code = code & "~fldEm2hgmXgcncpqc~:" & CStr(values.Cells.Item(1, 22).value + values.Cells.Item(1, 70).value) & "," ' 20852
    code = code & "~fldBHH15zEuQOzSGc~:" & values.Cells.Item(1, 23).value & "," ' 20841
    code = code & "~fldFQusJ5jh1qDJql~:" & values.Cells.Item(1, 24).value & "," ' 20862
    code = code & "~fldFz3qDfIl8QTMUt~:" & values.Cells.Item(1, 25).value & "," ' 20866
    code = code & "~fldsBFIpe2tMzJxNg~:" & values.Cells.Item(1, 26).value & "," ' 20818
    code = code & "~fldny8LMLt8RrOAZY~:" & values.Cells.Item(1, 29).value & "," ' 20871
    code = code & "~fld5XhCFuO3wJxCur~:" & CStr(values.Cells.Item(1, 30).value + values.Cells.Item(1, 82).value) & "," ' 20904
    code = code & "~fldXYnNjWL2RwuVvO~:" & CStr(values.Cells.Item(1, 31).value + values.Cells.Item(1, 83).value) & "," ' 20905
    code = code & "~fld0JeTuUa91pGPtm~:" & CStr(values.Cells.Item(1, 32).value + values.Cells.Item(1, 91).value) & "," ' 20914
    code = code & "~fldXUOs8ZJYNW5tcB~:" & values.Cells.Item(1, 33).value & "," ' 20872
    code = code & "~fldek98uEt3fRtVFB~:" & CStr(values.Cells.Item(1, 34).value + values.Cells.Item(1, 48).value) & "," ' 20874
    code = code & "~fld44lsv8AiFc4QBS~:" & CStr(values.Cells.Item(1, 35).value + values.Cells.Item(1, 39).value + values.Cells.Item(1, 64).value) & "," ' 20878
    code = code & "~fldwLEP4cGXRkq5fF~:" & CStr(values.Cells.Item(1, 36).value + values.Cells.Item(1, 73).value) & "," ' 20855
    code = code & "~fldJPmTwLQKDqfS43~:" & values.Cells.Item(1, 37).value & "," ' 20842
    code = code & "~fldcH1FBIC1Lkiucj~:" & CStr(values.Cells.Item(1, 38).value + values.Cells.Item(1, 56).value) & "," ' 20877
    code = code & "~fldoUp3DN1OrlNHPc~:" & CStr(values.Cells.Item(1, 40).value + values.Cells.Item(1, 54).value + values.Cells.Item(1, 57).value) & "," ' 20879
    code = code & "~fldAtzhUJvAQJp52C~:" & CStr(values.Cells.Item(1, 41).value + values.Cells.Item(1, 55).value) & "," ' 20882
    code = code & "~fldVrhadLAOZ3fGjl~:" & values.Cells.Item(1, 42).value & "," ' 20883
    code = code & "~fldIxjDIwL3Bzs5Ry~:" & values.Cells.Item(1, 43).value & "," ' 20884
    code = code & "~fldCkjplaeiCBw00x~:" & values.Cells.Item(1, 44).value & "," ' 20885
    code = code & "~fldghnEZa164OZoHz~:" & CStr(values.Cells.Item(1, 45).value + values.Cells.Item(1, 58).value) & "," ' 20886
    code = code & "~fldNKhOatNdhSvPB8~:" & values.Cells.Item(1, 46).value & "," ' 20898
    code = code & "~fldoWTbjdJbdlneKH~:" & values.Cells.Item(1, 47).value & "," ' 20896
    code = code & "~fldFsibkSfEfBGSPF~:" & values.Cells.Item(1, 49).value & "," ' 20875
    code = code & "~fld6fPFVG2q7ZzJIS~:" & values.Cells.Item(1, 50).value & "," ' 20876
    code = code & "~fldQyrc8n6uAFgorC~:" & values.Cells.Item(1, 51).value & "," ' 20812
    code = code & "~fldZDmMsh1tNRaE22~:" & values.Cells.Item(1, 52).value & "," ' 20891
    code = code & "~fldR7VAPIb0lZOY29~:" & values.Cells.Item(1, 53).value & "," ' 20895
    code = code & "~fldaf03OtbtUO2fBu~:" & values.Cells.Item(1, 59).value & "," ' 20830
    code = code & "~fldVoVP2558sITlV7~:" & values.Cells.Item(1, 60).value & "," ' 20832
    code = code & "~fldttgF9jaOLk8ZWS~:" & values.Cells.Item(1, 61).value & "," ' 20837
    code = code & "~flddvRfAeEsub9kkm~:" & CStr(values.Cells.Item(1, 62).value + values.Cells.Item(1, 72).value) & "," ' 20854
    code = code & "~fldY7ncvU7PPRn6dF~:" & CStr(values.Cells.Item(1, 63).value + values.Cells.Item(1, 74).value) & "," ' 20859
    code = code & "~fldSLbAhgEfm4vRjG~:" & values.Cells.Item(1, 65).value & "," ' 20847
    code = code & "~fldRy9UZIAtbbnHZ7~:" & values.Cells.Item(1, 66).value & "," ' 20848
    code = code & "~fldL4GpJtR7LDvQ29~:" & values.Cells.Item(1, 67).value & "," ' 20849
    code = code & "~fldmBpZ0g1YHwjjmX~:" & values.Cells.Item(1, 68).value & "," ' 20850
    code = code & "~fldEmhjUa2Rx3SmJk~:" & values.Cells.Item(1, 69).value & "," ' 20851
    code = code & "~fld8SWLCz37Det9Xx~:" & values.Cells.Item(1, 71).value & "," ' 20853
    code = code & "~fldpbad3Bj3DU5DM6~:" & values.Cells.Item(1, 75).value & "," ' 20860
    code = code & "~fldfofta2nXopQhod~:" & values.Cells.Item(1, 76).value & "," ' 20868
    code = code & "~fldClV99fdJHk3r6U~:" & CStr(values.Cells.Item(1, 77).value + values.Cells.Item(1, 89).value) & "," ' 20912
    code = code & "~fldSrvvy10dyhmBAT~:" & CStr(values.Cells.Item(1, 78).value + values.Cells.Item(1, 90).value) & "," ' 20913
    code = code & "~fldi7OqjtLmPWtLBA~:" & values.Cells.Item(1, 79).value & "," ' 20901
    code = code & "~fldGUwivS5boQYhDx~:" & CStr(values.Cells.Item(1, 80).value + values.Cells.Item(1, 96).value) & "," ' 20902
    code = code & "~fldphb40F2g1zPcvD~:" & values.Cells.Item(1, 81).value & "," ' 20903
    code = code & "~fldW5jHtEnBniaEfs~:" & values.Cells.Item(1, 85).value & "," ' 20907
    code = code & "~fld7P3tA3ZXqR3wyg~:" & values.Cells.Item(1, 86).value & "," ' 20908
    code = code & "~fldK5qGz7cnTvTTNr~:" & values.Cells.Item(1, 87).value & "," ' 20910
    code = code & "~fldBi9W3Y5rFhxgNn~:" & values.Cells.Item(1, 88).value & "," ' 20911
    code = code & "~fldhAm6nYXCUg2J1E~:" & CStr(values.Cells.Item(1, 92).value + values.Cells.Item(1, 97).value) & "," ' 20915
    code = code & "~fldHW4QssP8reBkCv~:" & values.Cells.Item(1, 94).value & "," ' 20918
    code = code & "~fldtnuoFOAbu2Kc0V~:" & values.Cells.Item(1, 95).value ' 20880
    
    code = code & "}}')"
    
    code = code & "; location.reload()"
    
    code = Replace(code, "~", """")

    With CreateObject("htmlfile")
        With .parentWindow.clipboardData
            .setData "text", code
        End With
    End With
    ThisWorkbook.FollowHyperlink address:="https://airtable.com/appSihm9stog1ZEFn/paglEufgcXNk939cP/form"
End Sub

' This macro subroutine may be used to double-check
' street addresses by lookup on the Gaithersburg city address search page in browser window.
'@EntryPoint
'@ExcelHotkey L
Public Sub LookupInCity()
Attribute LookupInCity.VB_ProcData.VB_Invoke_Func = "L\n14"
    Dim currentRowFirstCell As Range
    Set currentRowFirstCell = ThisWorkbook.ActiveSheet.Cells.Item(ActiveCell.row, 1)
    
    Dim record As RecordTuple
    Set record = records.loadRecordFromSheet(currentRowFirstCell)
    
    Dim AddrLookupURL As String
    AddrLookupURL = "https://maps.gaithersburgmd.gov/AddressSearch/index.html?address="
    Dim addr As String
    If (record.GburgFormatValidAddress.Item(addressKey.streetAddress) <> vbNullString) Then
        addr = record.GburgFormatValidAddress.Item(addressKey.streetAddress)
    Else
        addr = record.GburgFormatRawAddress.Item(addressKey.streetAddress)
    End If
    AddrLookupURL = AddrLookupURL & addr
    AddrLookupURL = Replace(AddrLookupURL, " ", "+")
    
    ThisWorkbook.FollowHyperlink address:=AddrLookupURL
End Sub

'@EntryPoint
Public Sub OpenAddressValidationWebsite()
    ThisWorkbook.FollowHyperlink address:="https://developers.google.com/maps/documentation/address-validation/demo"
End Sub

'@EntryPoint
Public Sub OpenUSPSZipcodeWebsite()
    ThisWorkbook.FollowHyperlink address:="https://tools.usps.com/zip-code-lookup.htm?byaddress"
End Sub

'@EntryPoint
Public Sub OpenGoogleMapsWebsite()
    ThisWorkbook.FollowHyperlink address:="https://www.google.com/maps"
End Sub
