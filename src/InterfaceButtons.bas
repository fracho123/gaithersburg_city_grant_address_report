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
            MsgBox "Disable filter on " & wsheet.name & " and try again"
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
        
        Select Case wsheet.name
            Case NonRxReportSheet.name, RxReportSheet.name
                wsheet.UsedRange.Offset(1, 0).AutoFilter
            ' Rubberduck bug
            '@Ignore UnreachableCase
            Case AddressesSheet.name, AutocorrectAddressesSheet.name, AutocorrectedAddressesSheet.name, DiscardsSheet.name
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
    
    getBlankRow(sheet.name).Cells.Item(1, 1).PasteSpecial Paste:=xlPasteValues
    
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
    Set addresses = records.loadAddresses(AddressesSheet.name)
    
    Dim discards As Scripting.Dictionary
    Set discards = records.loadAddresses(DiscardsSheet.name)
    
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
    SheetUtilities.getAddressVisitDataRng(AddressesSheet.name).Clear
    SheetUtilities.getRng(AddressesSheet.name, "A2", "A2").Offset(0, SheetUtilities.firstServiceColumn - 2).value = "{}"
    SheetUtilities.getAddressVisitDataRng(AutocorrectAddressesSheet.name).Clear
    SheetUtilities.getRng(AutocorrectAddressesSheet.name, "A2", "A2").Offset(0, SheetUtilities.firstServiceColumn - 2).value = "{}"
    SheetUtilities.getAddressVisitDataRng(DiscardsSheet.name).Clear
    SheetUtilities.getRng(DiscardsSheet.name, "A2", "A2").Offset(0, SheetUtilities.firstServiceColumn - 2).value = "{}"
    SheetUtilities.getAddressVisitDataRng(AutocorrectedAddressesSheet.name).Clear
    SheetUtilities.getRng(AutocorrectedAddressesSheet.name, "A2", "A2").Offset(0, SheetUtilities.firstServiceColumn - 2).value = "{}"

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
    Set Autocorrect = records.loadAddresses(AutocorrectAddressesSheet.name)
    
    Dim key As Variant
    For Each key In Autocorrect.Keys()
        records.writeAddress DiscardsSheet.name, Autocorrect.Item(key)
    Next key
    
    SheetUtilities.ClearSheet AutocorrectAddressesSheet.name
    SheetUtilities.SortSheet DiscardsSheet.name
    
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
        Set foundCell = findRow(AutocorrectedAddressesSheet.name, movedRecord.key)
        If Not foundCell Is Nothing Then
            foundCell.EntireRow.Delete
        End If
    Next movedRecord
End Sub

'@EntryPoint
Public Sub confirmDiscardSelected()
    If Not MacroEntry(ThisWorkbook.ActiveSheet) Then Exit Sub
    
    moveSelectedRows AutocorrectAddressesSheet.name, DiscardsSheet.name, False
    
    MacroExit ThisWorkbook.ActiveSheet
End Sub

'@EntryPoint
Public Sub confirmRestoreSelectedDiscard()
    If Not MacroEntry(ThisWorkbook.ActiveSheet) Then Exit Sub
    
    moveSelectedRows DiscardsSheet.name, AutocorrectAddressesSheet.name, True
    
    MacroExit ThisWorkbook.ActiveSheet
End Sub

'@EntryPoint
Public Sub confirmMoveAutocorrect()
    If Not MacroEntry(ThisWorkbook.ActiveSheet) Then Exit Sub
    
    moveSelectedRows AddressesSheet.name, AutocorrectAddressesSheet.name, True
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
        Set foundCell = findRow(AddressesSheet.name, key)
        
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
    wbook.Worksheets.[_Default](AddressesSheet.name).UsedRange.Copy
    AddressesSheet.Range("A1").PasteSpecial xlPasteValues
    
    wbook.Worksheets.[_Default](AutocorrectAddressesSheet.name).UsedRange.Copy
    AutocorrectAddressesSheet.Range("A1").PasteSpecial xlPasteValues
    
    wbook.Worksheets.[_Default](DiscardsSheet.name).UsedRange.Copy
    DiscardsSheet.Range("A1").PasteSpecial xlPasteValues
    
    wbook.Worksheets.[_Default](AutocorrectedAddressesSheet.name).UsedRange.Copy
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
    code = code & "localStorage.setItem('AirtableLocalPersister.formPageElementSavedFormDataByElementId.appSbQN8aFnRtJgDl.paghjbBNBGqpEbTLu.pelcNPzsIY7D38G7V','{~lastTouchedTime~:~"
    code = code & Format$(Date - 1, "yyyy-mm-dd") & "T21:25:49.955Z" & "~,"
    code = code & "~cellValueByColumnId~:{"
    code = code & "~fldl0cOUDvUCeDrln~: ~sel4b14BRPbjtmXOJ~,~fldRD1YXowe3Kg54S~: [{~foreignRowId~: ~recWHpCKSzyrrIFL3~,~foreignRowDisplayName~: ~20861~},{~foreignRowId~:~rec6JnzIIR8zbGoPj~,~foreignRowDisplayName~:~20906~},{~foreignRowId~:~rec0UoFQ634w3gzA0~,~foreignRowDisplayName~:~20916~},{~foreignRowId~:~recN7lnXDosJXZgEL~,~foreignRowDisplayName~:~20839~},{~foreignRowId~:~recqaXeoB3b55ZzKF~,~foreignRowDisplayName~:~20838~},{~foreignRowId~:~recYCWccvbmZiR8jg~,~foreignRowDisplayName~:~20813~},{~foreignRowId~:~recED69AIMnaEt9bV~,~foreignRowDisplayName~:~20814~},{~foreignRowId~:~recRNawBdXTAe0o6b~,~foreignRowDisplayName~:~20815~},{~foreignRowId~:~recP2nLMusF7DxOu2~,~foreignRowDisplayName~:~20816~},{~foreignRowId~:~recdQUTmjxgmldw8B~,~foreignRowDisplayName~:~20817~},{~foreignRowId~:~recoNTV14WcI7jlZc~,~foreignRowDisplayName~:~20824~},{~foreignRowId~:~recLYT0mN30ykq29x~,~foreignRowDisplayName~:~20825~},{~foreignRowId~:~recayKEdhMW2JEvFn~,~foreignRowDisplayName~:~20827~},"
    code = code & "{~foreignRowId~:~recoQG3eA8Vg2UZpo~,~foreignRowDisplayName~:~20852~},{~foreignRowId~:~recPoCR4CjPhj9esT~,~foreignRowDisplayName~:~20841~},{~foreignRowId~:~rec0GXUpHUiXKdoIR~,~foreignRowDisplayName~:~20862~},{~foreignRowId~:~recKFsEclsshXcJBj~,~foreignRowDisplayName~:~20866~},{~foreignRowId~:~recJberDjIBO9wk4R~,~foreignRowDisplayName~:~20818~},{~foreignRowId~:~recL9OTHxFea0gseq~,~foreignRowDisplayName~:~20871~},{~foreignRowId~:~rec2gNpYUu3ak9IEd~,~foreignRowDisplayName~:~20904~},{~foreignRowId~:~rec0h0ubwEVbkPuva~,~foreignRowDisplayName~:~20905~},{~foreignRowId~:~recbNubdjBEjkOMwP~,~foreignRowDisplayName~:~20914~},{~foreignRowId~:~rec9d3FCNjKW2tvC3~,~foreignRowDisplayName~:~20872~},{~foreignRowId~:~recmQZOoCAxosd6hO~,~foreignRowDisplayName~:~20874~},{~foreignRowId~:~recQGYrqN1zp6wWKJ~,~foreignRowDisplayName~:~20878~},{~foreignRowId~:~recczFtG9wSAye66t~,~foreignRowDisplayName~:~20855~},{~foreignRowId~:~rech98chfjXO9Ous7~,~foreignRowDisplayName~:~20842~},"
    code = code & "{~foreignRowId~:~recQwMVdsVLsehthI~,~foreignRowDisplayName~:~20877~},{~foreignRowId~:~recI05ELOhwmxTa80~,~foreignRowDisplayName~:~20879~},{~foreignRowId~:~rec5zZWgdLEa634Kx~,~foreignRowDisplayName~:~20882~},{~foreignRowId~:~recV1GOjqlE4dpvni~,~foreignRowDisplayName~:~20883~},{~foreignRowId~:~recbouJuGg7jGjBBb~,~foreignRowDisplayName~:~20884~},{~foreignRowId~:~recfaA9yA2YZTJubT~,~foreignRowDisplayName~:~20885~},{~foreignRowId~:~recuCpfxGyterlcjZ~,~foreignRowDisplayName~:~20886~},{~foreignRowId~:~recO1MSxJrJwBqBd4~,~foreignRowDisplayName~:~20898~},{~foreignRowId~:~recZsp2fjvP7PwXye~,~foreignRowDisplayName~:~20896~},{~foreignRowId~:~rectn5rpd0nWZmJcS~,~foreignRowDisplayName~:~20875~},{~foreignRowId~:~recdKoqjlxg5kxaWt~,~foreignRowDisplayName~:~20876~},{~foreignRowId~:~rec9aQ2bXKeSutK3l~,~foreignRowDisplayName~:~20812~},{~foreignRowId~:~recqCOeKzoQ7aHx44~,~foreignRowDisplayName~:~20891~},{~foreignRowId~:~rec3HskwYXqQgBI64~,~foreignRowDisplayName~:~20895~},"
    code = code & "{~foreignRowId~:~recTfdEPcc84giMfY~,~foreignRowDisplayName~:~20830~},{~foreignRowId~:~recRObHiBYi2ePLsl~,~foreignRowDisplayName~:~20832~},{~foreignRowId~:~recJiwIe86Jot20gg~,~foreignRowDisplayName~:~20837~},{~foreignRowId~:~reckvUVjCOK5hzVhY~,~foreignRowDisplayName~:~20854~},{~foreignRowId~:~recwQgN0onMHWbluK~,~foreignRowDisplayName~:~20859~},{~foreignRowId~:~recdklYkW97iPsHGE~,~foreignRowDisplayName~:~20847~},{~foreignRowId~:~recYA6Xo3rotLCIvY~,~foreignRowDisplayName~:~20848~},{~foreignRowId~:~recJOh8lBL4xTEDT6~,~foreignRowDisplayName~:~20849~},{~foreignRowId~:~recVdg9onP5mN5bPh~,~foreignRowDisplayName~:~20850~},{~foreignRowId~:~recR2jMhZOflDNt8I~,~foreignRowDisplayName~:~20851~},{~foreignRowId~:~recb12lIMIfRuHEk4~,~foreignRowDisplayName~:~20853~},{~foreignRowId~:~rec8Q7SyKhJChivev~,~foreignRowDisplayName~:~20860~},{~foreignRowId~:~recPT00HV2FwpMaQQ~,~foreignRowDisplayName~:~20868~},{~foreignRowId~:~recgcnZCQvJfsqkFT~,~foreignRowDisplayName~:~20912~},"
    code = code & "{~foreignRowId~:~reckodm7oUQiDbU8W~,~foreignRowDisplayName~:~20913~},{~foreignRowId~:~rec4XmPjkaHvgq4cd~,~foreignRowDisplayName~:~20901~},{~foreignRowId~:~recnzLv6tmo6J5m66~,~foreignRowDisplayName~:~20902~},{~foreignRowId~:~recAuxevvctrR7RGd~,~foreignRowDisplayName~:~20903~},{~foreignRowId~:~recr9Snbah0lcWkfp~,~foreignRowDisplayName~:~20907~},{~foreignRowId~:~rec6RgwBZKh1Nk2iV~,~foreignRowDisplayName~:~20908~},{~foreignRowId~:~recqVhZhZ0ZDXsMcf~,~foreignRowDisplayName~:~20910~},{~foreignRowId~:~recBZxXpmzl9cVNGv~,~foreignRowDisplayName~:~20911~},{~foreignRowId~:~recBRWy1CoeZA9aIY~,~foreignRowDisplayName~:~20915~},{~foreignRowId~:~reczeqj8yhSyGcDUE~,~foreignRowDisplayName~:~20918~},{~foreignRowId~:~recHL1TlVwpBqZ9kF~,~foreignRowDisplayName~:~20880~}],~fldAIwqOBPimQ2H18~:[{~foreignRowId~:~recX8z5e6kdwpqX4l~,~foreignRowDisplayName~:~GaithersburgHELP,Inc.~}],"
    
    code = code & "~fldcAkWSVtnA2WSBg~: " & values.Cells.Item(1, 2).value & ","
    code = code & "~fldTgMj5Z4w4jTyR4~:" & values.Cells.Item(1, 3).value & ","
    code = code & "~flds4icAlgYgtKxXa~:" & values.Cells.Item(1, 9).value & ","
    code = code & "~fldAmOauQkZoxmHAP~:" & CStr(values.Cells.Item(1, 10).value + values.Cells.Item(1, 84).value) & ","
    code = code & "~fldUzBeCIr6WrMYxO~:" & CStr(values.Cells.Item(1, 11).value + values.Cells.Item(1, 93).value) & ","
    code = code & "~fld6qKGCjv2lgfnHI~:" & values.Cells.Item(1, 12).value & ","
    code = code & "~fldBk8oK7AZFXjTEG~:" & values.Cells.Item(1, 13).value & ","
    code = code & "~fld6YSs58ktJehdLW~:" & values.Cells.Item(1, 14).value & ","
    code = code & "~fld9QSxMT6PT15f4N~:" & values.Cells.Item(1, 15).value & ","
    code = code & "~fldHtsF4FfGWW4wpp~:" & CStr(values.Cells.Item(1, 16).value + values.Cells.Item(1, 27)) & ","
    code = code & "~fldXq8cJ8Y2NWrJTn~:" & values.Cells.Item(1, 17).value & ","
    code = code & "~fldEc7D11e9vxK8Et~:" & values.Cells.Item(1, 18).value & ","
    code = code & "~fldQSTH5DP03lR4EE~:" & values.Cells.Item(1, 19).value & ","
    code = code & "~fldTMpf9usGCsSyJO~:" & CStr(values.Cells.Item(1, 20).value + values.Cells.Item(1, 28).value) & ","
    code = code & "~fldJRxJhz0lPktGKc~:" & values.Cells.Item(1, 21).value & ","
    code = code & "~fldLPNKB6QcdJ8fDK~:" & CStr(values.Cells.Item(1, 22).value + values.Cells.Item(1, 70).value) & ","
    code = code & "~fldjEMDUhU919URRV~:" & values.Cells.Item(1, 23).value & ","
    code = code & "~fldgVlvfPzxn5eCRq~:" & values.Cells.Item(1, 24).value & ","
    code = code & "~fld67bhZmv77zN9uA~:" & values.Cells.Item(1, 25).value & ","
    code = code & "~fldd0RYD8HnI0rYcG~:" & values.Cells.Item(1, 26).value & ","
    code = code & "~fldP9N2QVs5dxFKT0~:" & values.Cells.Item(1, 29).value & ","
    code = code & "~fldWbPWbKJ3jfq9Fl~:" & CStr(values.Cells.Item(1, 30).value + values.Cells.Item(1, 82).value) & ","
    code = code & "~fldGm2CNYaDfj1WDJ~:" & CStr(values.Cells.Item(1, 31).value + values.Cells.Item(1, 83).value) & ","
    code = code & "~fldSpkgr2gF23ubJB~:" & CStr(values.Cells.Item(1, 32).value + values.Cells.Item(1, 91).value) & ","
    code = code & "~fld0fmGzYhxFtfSMs~:" & values.Cells.Item(1, 33).value & ","
    code = code & "~fldJbb7bJv9fzf81Y~:" & CStr(values.Cells.Item(1, 34).value + values.Cells.Item(1, 48).value) & ","
    code = code & "~fldPRC1T7NrvJ4stH~:" & CStr(values.Cells.Item(1, 35).value + values.Cells.Item(1, 39).value + values.Cells.Item(1, 64).value) & ","
    code = code & "~fldGSQ0dn8s89IHvb~:" & CStr(values.Cells.Item(1, 36).value + values.Cells.Item(1, 73).value) & ","
    code = code & "~fldmjaDRuekAqsa7C~:" & values.Cells.Item(1, 37).value & ","
    code = code & "~fldJMQNMniEiQdbQ0~:" & CStr(values.Cells.Item(1, 38).value + values.Cells.Item(1, 56).value) & ","
    code = code & "~fldGz7xlhjMnRaeop~:" & CStr(values.Cells.Item(1, 40).value + values.Cells.Item(1, 54).value + values.Cells.Item(1, 57).value) & ","
    code = code & "~fld0Hypc9v51vPh24~:" & CStr(values.Cells.Item(1, 41).value + values.Cells.Item(1, 55).value) & ","
    code = code & "~flduaScn4cbua1tzm~:" & values.Cells.Item(1, 42).value & ","
    code = code & "~flduJb5vwcxRZTPr6~:" & values.Cells.Item(1, 43).value & ","
    code = code & "~fldJc3oAaR7WRSUVO~:" & values.Cells.Item(1, 44).value & ","
    code = code & "~fldL4uViATWSYFzUK~:" & CStr(values.Cells.Item(1, 45).value + values.Cells.Item(1, 58).value) & ","
    code = code & "~fldcayWHBXuqwMayz~:" & values.Cells.Item(1, 46).value & ","
    code = code & "~fldp2SINVQlrjeQCk~:" & values.Cells.Item(1, 47).value & ","
    code = code & "~fld9Ego3Gt2GCUNqZ~:" & values.Cells.Item(1, 49).value & ","
    code = code & "~fldEoHWUYrFOzmcix~:" & values.Cells.Item(1, 50).value & ","
    code = code & "~fldcRgzWt69k85pMk~:" & values.Cells.Item(1, 51).value & ","
    code = code & "~fldIrsEhpV0XYRTGr~:" & values.Cells.Item(1, 52).value & ","
    code = code & "~fldkeqSwi6FY35g15~:" & values.Cells.Item(1, 53).value & ","
    code = code & "~fldz2OgXVhm6VYlGb~:" & values.Cells.Item(1, 59).value & ","
    code = code & "~fldLqzEJz4AAyWLjY~:" & values.Cells.Item(1, 60).value & ","
    code = code & "~fldDZYTEXeWwCGebu~:" & values.Cells.Item(1, 61).value & ","
    code = code & "~fldh49H9rAvxFI3Sb~:" & CStr(values.Cells.Item(1, 62).value + values.Cells.Item(1, 72).value) & ","
    code = code & "~fldGXGXeqx2dErAce~:" & CStr(values.Cells.Item(1, 63).value + values.Cells.Item(1, 74).value) & ","
    code = code & "~fld5uhbvzq9QMwoHD~:" & values.Cells.Item(1, 65).value & ","
    code = code & "~fld0jMzTXV1aJkw5S~:" & values.Cells.Item(1, 66).value & ","
    code = code & "~fld5OcDa1nsoLBwfy~:" & values.Cells.Item(1, 67).value & ","
    code = code & "~fldlYAEqyU6JApxio~:" & values.Cells.Item(1, 68).value & ","
    code = code & "~fldJhYaiSG4a6sqHY~:" & values.Cells.Item(1, 69).value & ","
    code = code & "~fldS2dGxCP8LpbJi7~:" & values.Cells.Item(1, 71).value & ","
    code = code & "~fldO7lvG9pd2Y5WzK~:" & values.Cells.Item(1, 75).value & ","
    code = code & "~fldkUX0808Jz0sg59~:" & values.Cells.Item(1, 76).value & ","
    code = code & "~fldK2C3IWuW9rIfVH~:" & CStr(values.Cells.Item(1, 77).value + values.Cells.Item(1, 89).value) & ","
    code = code & "~fldeNgJa1zeSdd7eX~:" & CStr(values.Cells.Item(1, 78).value + values.Cells.Item(1, 90).value) & ","
    code = code & "~fld987t6GKjDJVnPO~:" & values.Cells.Item(1, 79).value & ","
    code = code & "~fldXEfo3haPMVr8Wo~:" & CStr(values.Cells.Item(1, 80).value + values.Cells.Item(1, 96).value) & ","
    code = code & "~fldc2ls2xOBA4wgxz~:" & values.Cells.Item(1, 81).value & ","
    code = code & "~fldRak02MKqgSs6VJ~:" & values.Cells.Item(1, 85).value & ","
    code = code & "~fldLCzskTnmLYHqym~:" & values.Cells.Item(1, 86).value & ","
    code = code & "~fld9dQWSUzcGdTVGI~:" & values.Cells.Item(1, 87).value & ","
    code = code & "~fldjmoMRWUWs5WdXb~:" & values.Cells.Item(1, 88).value & ","
    code = code & "~fldN2QYoc9eBu0vNH~:" & CStr(values.Cells.Item(1, 92).value + values.Cells.Item(1, 97)) & ","
    code = code & "~fldttJzWI1rdLMgtf~:" & values.Cells.Item(1, 94).value & ","
    code = code & "~fldKLR2lQdFHoX9QV~:" & values.Cells.Item(1, 95).value
    
    code = code & "}}')"
    
    code = code & "; location.reload()"
    
    code = Replace(code, "~", """")

    With CreateObject("htmlfile")
        With .parentWindow.clipboardData
            .setData "text", code
        End With
    End With
    ThisWorkbook.FollowHyperlink address:="https://airtable.com/appSbQN8aFnRtJgDl/paghjbBNBGqpEbTLu/form"
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
