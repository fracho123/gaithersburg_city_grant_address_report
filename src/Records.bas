Attribute VB_Name = "Records"
Option Explicit

'@Folder "City_Grant_Address_Report.src"
Public Enum TotalServiceType
    nonDelivery = 1
    delivery = 2
    numDoubleCountedAdditionalDeliveryType = 3
    [_TotalServiceTypeFirst] = nonDelivery
    [_TotalServiceTypeLast] = numDoubleCountedAdditionalDeliveryType
End Enum

Public Enum TotalType
    uniqueGuestID = 1
    uniqueGuestIDHousehold = 2
    nonUniqueGuestID = 3
    nonUniqueHousehold = 4
    rx = 5
    [_TotalTypeFirst] = uniqueGuestID
    [_TotalTypeLast] = rx
End Enum

Public Type ComputedRx
    totals As RxTotals
    records As RxRecords
End Type

Public Function getQuarterNum(ByVal quarter As String) As Long
    Select Case quarter
        Case "Q1"
            getQuarterNum = 1
        Case "Q2"
            getQuarterNum = 2
        Case "Q3"
            getQuarterNum = 3
        Case "Q4"
            getQuarterNum = 4
    End Select
End Function

Private Function loadRecordFromRaw(ByVal recordRowFirstCell As Range) As RecordTuple
    Dim record As RecordTuple
    Set record = New RecordTuple
    
    record.AddVisit recordRowFirstCell.value, recordRowFirstCell.Offset(0, 1).value
    If (CDate(recordRowFirstCell.value) > CDate(getInterfaceMostRecentRng.value)) Then
        getInterfaceMostRecentRng.value = CStr(CDate(recordRowFirstCell.value))
    End If
    
    record.UserVerified = False

    record.guestID = recordRowFirstCell.Offset(0, 2).value
    record.FirstName = recordRowFirstCell.Offset(0, 3).value
    record.LastName = recordRowFirstCell.Offset(0, 4).value
    record.RawAddress = recordRowFirstCell.Offset(0, 5).value
    record.RawUnitWithNum = recordRowFirstCell.Offset(0, 6).value
    record.RawCity = recordRowFirstCell.Offset(0, 7).value
    record.RawState = recordRowFirstCell.Offset(0, 8).value
    record.RawZip = recordRowFirstCell.Offset(0, 9).value
    
    Dim val As String
    val = recordRowFirstCell.Offset(0, 10).value
    ' Count blank totals as 1 for household total and 18+
    record.householdTotal = IIf(IsNumeric(val), val, 1)
    val = recordRowFirstCell.Offset(0, 11).value
    record.zeroToOneTotal = IIf(IsNumeric(val), val, 0)
    val = recordRowFirstCell.Offset(0, 12).value
    record.twoToSeventeenTotal = IIf(IsNumeric(val), val, 0)
    val = recordRowFirstCell.Offset(0, 13).value
    record.eighteenPlusTotal = IIf(IsNumeric(val), val, 1)
    
    Dim rx As Double
    rx = recordRowFirstCell.Offset(0, 14).value
    If IsNumeric(rx) And (rx <> 0) Then record.addRx recordRowFirstCell.value, rx
    
    Set loadRecordFromRaw = record
End Function

Public Function loadRecordFromSheet(ByVal recordRowFirstCell As Range) As RecordTuple
    Dim record As RecordTuple
    Set record = New RecordTuple
    
    Dim services() As String
    services = loadServiceNames(recordRowFirstCell.Worksheet.Name)
    
    record.SetInCity recordRowFirstCell.Offset(0, 0).value
    record.UserVerified = CBool(recordRowFirstCell.Offset(0, 1).value)
    record.validAddress = recordRowFirstCell.Offset(0, 2).value
    record.validUnitWithNum = recordRowFirstCell.Offset(0, 3).value
    record.ValidZipcode = recordRowFirstCell.Offset(0, 4).value
    record.RawAddress = recordRowFirstCell.Offset(0, 5).value
    record.RawUnitWithNum = recordRowFirstCell.Offset(0, 6).value
    record.RawCity = recordRowFirstCell.Offset(0, 7).value
    record.RawState = recordRowFirstCell.Offset(0, 8).value
    record.RawZip = recordRowFirstCell.Offset(0, 9).value
    record.guestID = recordRowFirstCell.Offset(0, 10).value
    record.FirstName = recordRowFirstCell.Offset(0, 11).value
    record.LastName = recordRowFirstCell.Offset(0, 12).value
    record.householdTotal = recordRowFirstCell.Offset(0, 13).value
    record.zeroToOneTotal = recordRowFirstCell.Offset(0, 14).value
    record.twoToSeventeenTotal = recordRowFirstCell.Offset(0, 15).value
    record.eighteenPlusTotal = recordRowFirstCell.Offset(0, 16).value
    
    Set record.RxTotal = JsonConverter.ParseJson(recordRowFirstCell.Offset(0, _
                                                        SheetUtilities.firstServiceColumn - 2).value)
    
    Dim visitData As Scripting.Dictionary
    Set visitData = New Scripting.Dictionary
    
    Dim j As Long
    j = 1
    Do While j <= UBound(services) + 1
        Dim visitJson As String
        visitJson = recordRowFirstCell.Offset(0, SheetUtilities.firstServiceColumn - 2 + j).value
        If visitJson <> vbNullString Then
            visitData.Add services(j - 1), JsonConverter.ParseJson(visitJson)
        End If
        j = j + 1
    Loop
    
    Set record.visitData = visitData
    
    Set loadRecordFromSheet = record
End Function

Public Function loadAddresses(ByVal sheetName As String) As Scripting.Dictionary
    Dim addresses As Scripting.Dictionary
    Set addresses = New Scripting.Dictionary
    
    Dim sheet As Worksheet
    Set sheet = ThisWorkbook.Worksheets.[_Default](sheetName)
       
    If sheet.Range("A2").value = vbNullString Then
        Set loadAddresses = addresses
        Exit Function
    End If
    
    Dim appStatus As Variant
    If Application.StatusBar = False Then appStatus = False Else appStatus = Application.StatusBar
    
    Application.StatusBar = "Loading address 1 of " & getBlankRow(sheetName).row
    
    Dim i As Long
    i = 2
    Do While i < getBlankRow(sheetName).row
        Dim recordRowFirstCell As Range
        Set recordRowFirstCell = sheet.rows.Item(i).Cells.Item(1, 1)
        
        Dim record As RecordTuple
        Set record = loadRecordFromSheet(recordRowFirstCell)
        
        addresses.Add record.key, record
        i = i + 1
        
        Application.StatusBar = "Loading address " & i & " of " & getBlankRow(sheetName).row
    Loop
    
    Application.StatusBar = appStatus

    Set loadAddresses = addresses
End Function

Public Sub writeAddress(ByVal sheetName As String, ByVal record As RecordTuple)
    Dim sheet As Worksheet
    Set sheet = ThisWorkbook.Worksheets.[_Default](sheetName)
    
    ' Saves column numbers per existing service
    Dim serviceCols As Scripting.Dictionary
    Set serviceCols = New Scripting.Dictionary
    
    If sheet.Range("A2").value <> vbNullString Then
        Dim services() As String
        services = loadServiceNames(sheetName)
        Dim i As Long
        i = SheetUtilities.firstServiceColumn
        Dim service As Variant
        For Each service In services
            If service <> vbNullString Then
                serviceCols.Add service, i
                i = i + 1
            End If
        Next
    End If

    Dim recordRow As Range
    Set recordRow = getBlankRow(sheetName)
    
    recordRow.Cells.Item(1, 1).value = record.InCityStr
    recordRow.Cells.Item(1, 2).value = record.UserVerified
    recordRow.Cells.Item(1, 3).value = record.validAddress
    recordRow.Cells.Item(1, 4).value = record.validUnitWithNum
    recordRow.Cells.Item(1, 5).value = record.ValidZipcode
    recordRow.Cells.Item(1, 6).value = record.RawAddress
    recordRow.Cells.Item(1, 7).value = record.RawUnitWithNum
    recordRow.Cells.Item(1, 8).value = record.RawCity
    recordRow.Cells.Item(1, 9).value = record.RawState
    recordRow.Cells.Item(1, 10).value = record.RawZip
    recordRow.Cells.Item(1, 11).value = record.guestID
    recordRow.Cells.Item(1, 12).value = record.FirstName
    recordRow.Cells.Item(1, 13).value = record.LastName
    recordRow.Cells.Item(1, 14).value = record.householdTotal
    recordRow.Cells.Item(1, 15).value = record.zeroToOneTotal
    recordRow.Cells.Item(1, 16).value = record.twoToSeventeenTotal
    recordRow.Cells.Item(1, 17).value = record.eighteenPlusTotal
    
    recordRow.Cells.Item(1, SheetUtilities.firstServiceColumn - 1).value = _
                                                            JsonConverter.ConvertToJson(record.RxTotal)
    
    Dim serviceToAdd As Variant
    For Each serviceToAdd In record.visitData.Keys
        Dim visitDataToAdd As String
        visitDataToAdd = JsonConverter.ConvertToJson(record.visitData.Item(serviceToAdd))
        
        If Not serviceCols.exists(serviceToAdd) Then
            Dim newServiceCol As Long
            newServiceCol = SheetUtilities.firstServiceColumn + UBound(serviceCols.Keys) + 1
            serviceCols.Add serviceToAdd, newServiceCol
            ThisWorkbook.Worksheets.[_Default](sheetName).Cells(1, newServiceCol).value = serviceToAdd
        End If
        
        recordRow.Cells.Item(1, serviceCols.Item(serviceToAdd)).value = visitDataToAdd
    Next serviceToAdd
End Sub

Public Sub writeAddresses(ByVal sheetName As String, ByVal addresses As Scripting.Dictionary)
    ClearSheet sheetName
    
    Dim appStatus As Variant
    If Application.StatusBar = False Then appStatus = False Else appStatus = Application.StatusBar
    
    Dim i As Long
    Application.StatusBar = "Writing address 1 of " & UBound(addresses.Keys) + 1
    
    Dim key As Variant
    For Each key In addresses.Keys
        writeAddress sheetName, addresses.Item(key)
        Application.StatusBar = "Writing address " & i & " of " & UBound(addresses.Keys) + 1
        i = i + 1
    Next key
    
    Application.StatusBar = appStatus
    
End Sub

Private Sub incrementCountyTotal(ByVal record As RecordTuple)
    Dim monthVisited(1 To 12) As Boolean
    Dim uniqueGuestIDTotal(1 To 12) As Long
    Dim uniqueGuestIDHouseholdTotal(1 To 12) As Long
    Dim guestIDTotal(1 To 12) As Long
    Dim householdTotal(1 To 12) As Long
    Dim childrenTotal(1 To 12) As Long
    Dim adultTotal(1 To 12) As Long
    
    Dim monthNum As Long
    
    Dim includedServices As String
    includedServices = SheetUtilities.getCountyTotalServicesRng.value
    
    Dim service As Variant
    For Each service In record.visitData.Keys
        If service = "Rx Asst" Then
            GoTo NextService
        Else
            If includedServices = vbNullString Then
                includedServices = service
            ElseIf InStr(1, includedServices, service, vbTextCompare) = 0 Then
                includedServices = includedServices & "," & service
            End If
        End If
        Dim quarter As Variant
        For Each quarter In record.visitData.Item(service).Keys
            Dim visit As Variant
            For Each visit In record.visitData.Item(service).Item(quarter)
                monthNum = Month(visit)
                guestIDTotal(monthNum) = guestIDTotal(monthNum) + 1
                householdTotal(monthNum) = householdTotal(monthNum) + record.householdTotal
                childrenTotal(monthNum) = childrenTotal(monthNum) + record.zeroToOneTotal + _
                                          record.twoToSeventeenTotal
                adultTotal(monthNum) = adultTotal(monthNum) + record.eighteenPlusTotal
                If Not monthVisited(monthNum) Then
                    uniqueGuestIDTotal(monthNum) = uniqueGuestIDTotal(monthNum) + 1
                    uniqueGuestIDHouseholdTotal(monthNum) = uniqueGuestIDHouseholdTotal(monthNum) + _
                                                            record.householdTotal
                    monthVisited(monthNum) = True
                End If
            Next visit
        Next quarter
NextService:
    Next service
    
    SheetUtilities.getCountyTotalServicesRng.value = includedServices
    
    Dim zip As String
    If record.ValidZipcode <> vbNullString Then
        zip = record.ValidZipcode
    Else
        zip = record.RawZip
    End If
    
    Dim uniqueCols As Scripting.Dictionary
    Set uniqueCols = SheetUtilities.uniqueCountyZipCols()
    
    Dim zipCol As Long
    
    If uniqueCols.exists(zip) Then
        zipCol = uniqueCols.Item(zip)
    Else
        ' BUG assumes poorer city
    Select Case zip
        Case 20906
            If record.RawCity = "Ashton" Or record.RawCity = "Aspen Hill" Then
                zipCol = CountyTotalCols.zip20906AshtonAspenHill
            Else
                zipCol = CountyTotalCols.zip20906SilverSpring
            End If
        Case 20916
            If record.RawCity = "Ashton" Or record.RawCity = "Aspen Hill" Then
                zipCol = CountyTotalCols.zip20916AshtonAspenHill
            Else
                zipCol = CountyTotalCols.zip20916SilverSpring
            End If
        Case 20815
            If record.RawCity = "Chevy Chase" Or record.RawCity = "Clarksburg" Then
                zipCol = CountyTotalCols.zip20815ChevyChaseClarksburg
            Else
                zipCol = CountyTotalCols.zip20815Bethesda
            End If
        Case 20825
            If record.RawCity = "Chevy Chase" Or record.RawCity = "Clarksburg" Then
                zipCol = CountyTotalCols.zip20825ChevyChaseClarksburg
            Else
                zipCol = CountyTotalCols.zip20825Bethesda
            End If
        Case 20852
            If record.RawCity = "Bethesda" Then
                zipCol = CountyTotalCols.zip20852Bethesda
            Else
                zipCol = CountyTotalCols.zip20852Rockville
            End If
        Case 20904
            If record.RawCity = "Colesville" Or record.RawCity = "Damascus" Then
                zipCol = CountyTotalCols.zip20904ColesvilleDamascus
            Else
                zipCol = CountyTotalCols.zip20904SilverSpring
            End If
        Case 20905
            If record.RawCity = "Colesville" Or record.RawCity = "Damascus" Then
                zipCol = CountyTotalCols.zip20905ColesvilleDamascus
            Else
                zipCol = CountyTotalCols.zip20905SilverSpring
            End If
        Case 20914
            If record.RawCity = "Colesville" Or record.RawCity = "Damascus" Then
                zipCol = CountyTotalCols.zip20914ColesvilleDamascus
            Else
                zipCol = CountyTotalCols.zip20914SilverSpring
            End If
        Case 20874
            If record.RawCity = "Darnestown" Or record.RawCity = "Derwood" Or record.RawCity = "Dickerson" Then
                zipCol = CountyTotalCols.zip20874DarnestownDerwoodDickerson
            Else
                zipCol = CountyTotalCols.zip20874GarrettParkGermantownGlenEcho
            End If
        Case 20878
            If record.RawCity = "Darnestown" Or record.RawCity = "Derwood" Or record.RawCity = "Dickerson" Then
                zipCol = CountyTotalCols.zip20878DarnestownDerwoodDickerson
            ElseIf record.RawCity = "Poolesville" Or record.RawCity = "Potomac" Then
                zipCol = CountyTotalCols.zip20878PoolesvillePotomac
            Else
                zipCol = CountyTotalCols.zip20878Gaithersburg
            End If
        Case 20855
            If record.RawCity = "Darnestown" Or record.RawCity = "Derwood" Or record.RawCity = "Dickerson" Then
                zipCol = CountyTotalCols.zip20855DarnestownDerwoodDickerson
            Else
                zipCol = CountyTotalCols.zip20855Rockville
            End If
        Case 20877
            If record.RawCity = "Montgomery Village" Or record.RawCity = "Olney" Then
                zipCol = CountyTotalCols.zip20877MontgomeryVillageOlney
            Else
                zipCol = CountyTotalCols.zip20877Gaithersburg
            End If
        Case 20882
            If record.RawCity = "Kensington" Or record.RawCity = "Laytonsville" Then
                zipCol = CountyTotalCols.zip20882KensingtonLaytonsville
            Else
                zipCol = CountyTotalCols.zip20882Gaithersburg
            End If
        Case 20886
            If record.RawCity = "Montgomery Village" Or record.RawCity = "Olney" Then
                zipCol = CountyTotalCols.zip20886MontgomeryVillageOlney
            Else
                zipCol = CountyTotalCols.zip20886Gaithersburg
            End If
        Case 20879
            If record.RawCity = "Kensington" Or record.RawCity = "Laytonsville" Then
                zipCol = CountyTotalCols.zip20879KensingtonLaytonsville
            ElseIf record.RawCity = "Montgomery Village" Or record.RawCity = "Olney" Then
                zipCol = CountyTotalCols.zip20879MontgomeryVillageOlney
            Else
                zipCol = CountyTotalCols.zip20879Gaithersburg
            End If
        Case 20854
            If record.RawCity = "Poolesville" Or record.RawCity = "Potomac" Then
                zipCol = CountyTotalCols.zip20854PoolesvillePotomac
            Else
                zipCol = CountyTotalCols.zip20854Rockville
            End If
        Case 20859
            If record.RawCity = "Poolesville" Or record.RawCity = "Potomac" Then
                zipCol = CountyTotalCols.zip20859PoolesvillePotomac
            Else
                zipCol = CountyTotalCols.zip20859Rockville
            End If
        Case 20912
            If record.RawCity = "Sandy Spring" Or record.RawCity = "Spencerville" Or record.RawCity = "Takoma Park" Then
                zipCol = CountyTotalCols.zip20912SandySpringSpencervilleTakomaPark
            Else
                zipCol = CountyTotalCols.zip20912SilverSpring
            End If
        Case 20913
            If record.RawCity = "Sandy Spring" Or record.RawCity = "Spencerville" Or record.RawCity = "Takoma Park" Then
                zipCol = CountyTotalCols.zip20913SandySpringSpencervilleTakomaPark
            Else
                zipCol = CountyTotalCols.zip20913SilverSpring
            End If
        Case 20902
            If record.RawCity = "Washington Grove" Or record.RawCity = "Wheaton" Then
                zipCol = CountyTotalCols.zip20902WashingtonGroveWheaton
            Else
                zipCol = CountyTotalCols.zip20902SilverSpring
            End If
        Case 20915
            If record.RawCity = "Washington Grove" Or record.RawCity = "Wheaton" Then
                zipCol = CountyTotalCols.zip20915WashingtonGroveWheaton
            Else
                zipCol = CountyTotalCols.zip20915SilverSpring
            End If
        Case Else
            zipCol = -1
        End Select
    End If
    
    Dim i As Long
    For i = 1 To 12
        ' TODO off-by-one error on column enum
        getCountyRng.Cells.Item(i, CountyTotalCols.householdDuplicate - 1) = _
            getCountyRng.Cells.Item(i, CountyTotalCols.householdDuplicate - 1) + guestIDTotal(i)
        getCountyRng.Cells.Item(i, CountyTotalCols.householdUnduplicate - 1) = _
            getCountyRng.Cells.Item(i, CountyTotalCols.householdUnduplicate - 1) + uniqueGuestIDTotal(i)
        getCountyRng.Cells.Item(i, CountyTotalCols.individualDuplicate - 1) = _
            getCountyRng.Cells.Item(i, CountyTotalCols.individualDuplicate - 1) + householdTotal(i)
        getCountyRng.Cells.Item(i, CountyTotalCols.individualUnduplicate - 1) = _
            getCountyRng.Cells.Item(i, CountyTotalCols.individualUnduplicate - 1) + uniqueGuestIDHouseholdTotal(i)
        getCountyRng.Cells.Item(i, CountyTotalCols.childrenDuplicate - 1) = _
            getCountyRng.Cells.Item(i, CountyTotalCols.childrenDuplicate - 1) + childrenTotal(i)
        getCountyRng.Cells.Item(i, CountyTotalCols.adultDuplicate - 1) = _
            getCountyRng.Cells.Item(i, CountyTotalCols.adultDuplicate - 1) + adultTotal(i)
        getCountyRng.Cells.Item(i, CountyTotalCols.poundsFood - 1) = _
            getCountyRng.Cells.Item(i, CountyTotalCols.poundsFood - 1) + (householdTotal(i) * 8)
        
        If zipCol <> -1 Then
            getCountyRng.Cells.Item(i, zipCol - 1) = getCountyRng.Cells.Item(i, zipCol - 1) + guestIDTotal(i)
        End If
        
    Next i
End Sub

Private Sub loadAddressComputeCountyTotal(ByVal sheetName As String)
    Dim appStatus As Variant
    If Application.StatusBar = False Then appStatus = False Else appStatus = Application.StatusBar
    
    Dim addresses As Scripting.Dictionary
    Set addresses = records.loadAddresses(sheetName)
    
    Dim recordProgress As Long
    recordProgress = 1
    Application.StatusBar = "County totaling " & sheetName & " address 1 of " & UBound(addresses.Keys) + 1
    
    Dim key As Variant
    For Each key In addresses.Keys
        incrementCountyTotal addresses.Item(key)
        recordProgress = recordProgress + 1
        Application.StatusBar = "County totaling " & sheetName & " address " & recordProgress & " of " & UBound(addresses.Keys) + 1
    Next key
    
    Application.StatusBar = appStatus
End Sub

Public Function computeRxTotals(ByVal addresses As Scripting.Dictionary) As ComputedRx
    Dim totals As RxTotals
    Set totals = New RxTotals
    
    RxSheet.Columns.Item("A").NumberFormat = "mm/dd/yyyy"
    With SheetUtilities.getPastedRxRecordsRng
        .Sort key1:=.Columns.Item(1), order1:=xlAscending, Header:=xlNo
    End With
    
    Dim rxReportRecords As RxRecords
    Set rxReportRecords = New RxRecords
    
    Dim notInCityRecords As RxRecords
    Set notInCityRecords = New RxRecords
    
    Dim i As Long
    i = SheetUtilities.getPastedRxRecordsRng.row
    Do While i < getBlankRow(RxSheet.Name).row
        Dim guestID As String
        guestID = RxSheet.Cells.Item(i, 2).value
        
        If Not addresses.exists(guestID) Then GoTo DiscardRecord
        
        Dim addressRecord As RecordTuple
        Set addressRecord = addresses.Item(guestID)
        
        Dim visitDate As Date
        visitDate = RxSheet.Cells.Item(i, 1)
        
        Dim quarter As QuarterNum
        quarter = getQuarterNum(getQuarterStr(CStr(visitDate)))
        
        Dim hasRx As Boolean
        hasRx = False
        
        Dim j As Long
        For j = 1 To 6
            Dim Name As String
            Dim prevName As String
            Dim medication As String
            medication = RxSheet.Cells.Item(i, 1 + (3 * j)).value
            If medication = vbNullString Then GoTo NextName
            
            Name = RxSheet.Cells.Item(i, 3 * j).value
            If (Name = vbNullString) Then
                If (prevName <> vbNullString) Then
                    ' Assume previous name
                    Name = prevName
                Else
                    ' Assume address name if medication with no name
                    Name = addressRecord.cleanName
                End If
            End If
            
            hasRx = True
            
            Name = CleanString(Name)
            
            ' Replace - ; : with space
            ' Get proper cased first word
            medication = Replace(medication, "-", " ", vbTextCompare)
            medication = Replace(medication, ";", " ", vbTextCompare)
            medication = Replace(medication, ":", " ", vbTextCompare)
            medication = LWordTrim(medication)(0)
            medication = StrConv(medication, vbProperCase)
            
            If rxReportRecords.exists(Name) Then
                rxReportRecords.guestRecord(Name).quarter(quarter) = True
                If Not rxReportRecords.guestRecord(Name).medications.exists(medication) Then
                    totals.numUndupRxInGburgServed(quarter) = totals.numUndupRxInGburgServed(quarter) + 1
                    totals.numUndupRxInProgramServed(quarter) = totals.numUndupRxInProgramServed(quarter) + 1
                    rxReportRecords.guestRecord(Name).medications.Item(medication) = 1
                End If
            ElseIf notInCityRecords.exists(Name) Then
                If Not notInCityRecords.guestRecord(Name).medications.exists(medication) Then
                    totals.numUndupRxInProgramServed(quarter) = totals.numUndupRxInProgramServed(quarter) + 1
                    notInCityRecords.guestRecord(Name).medications.Item(medication) = 1
                End If
            Else
                Dim newRecord As RxRecord
                Set newRecord = New RxRecord
                newRecord.guestID = guestID
                newRecord.medications.Item(medication) = 1
                newRecord.quarter(quarter) = True
                
                totals.numUndupRxInProgramServed(quarter) = totals.numUndupRxInProgramServed(quarter) + 1
                If addressRecord.InCity = ValidInCity Then
                    totals.numUndupResidentsServedRx(quarter) = totals.numUndupResidentsServedRx(quarter) + 1
                    totals.numUndupRxInGburgServed(quarter) = totals.numUndupRxInGburgServed(quarter) + 1
                    
                    'Rubberduck bug
                    '@Ignore ValueRequired, ObjectVariableNotSet
                    rxReportRecords.guestRecord(Name) = newRecord
                Else
                    '@Ignore ValueRequired, ObjectVariableNotSet
                    notInCityRecords.guestRecord(Name) = newRecord
                End If
            End If
            
            prevName = Name
NextName:
        Next j
        
        If Not hasRx Then GoTo DiscardRecord
        
        Dim cost As Double
        cost = RxSheet.Cells.Item(i, 21).value
        
        totals.costRxInProgram(quarter) = totals.costRxInProgram(quarter) + cost
        If addressRecord.InCity = ValidInCity Then
            totals.costRxInGburg(quarter) = totals.costRxInGburg(quarter) + cost
        End If
        
        If totals.mostRecentRxDate = vbNullString Then
            totals.mostRecentRxDate = visitDate
        ElseIf visitDate > DateValue(totals.mostRecentRxDate) Then
            totals.mostRecentRxDate = visitDate
        End If
        
        GoTo NextRxRecord
DiscardRecord:
        If totals.discardedIDs = vbNullString Then
            totals.discardedIDs = guestID
        Else
            totals.discardedIDs = totals.discardedIDs & "," & guestID
        End If
NextRxRecord:
        i = i + 1
    Loop
    
    Dim out As ComputedRx
    Set out.totals = totals
    Set out.records = rxReportRecords
    computeRxTotals = out
End Function

Public Sub computeCountyTotals()
    getCountyRng.value = 0
    
    loadAddressComputeCountyTotal AddressesSheet.Name
    loadAddressComputeCountyTotal AutocorrectAddressesSheet.Name
    loadAddressComputeCountyTotal DiscardsSheet.Name
End Sub

Private Function getServiceType(ByVal service As Variant) As TotalServiceType
    If InStr(1, service, "delivery", vbTextCompare) > 0 Then
        getServiceType = delivery
    Else
        getServiceType = nonDelivery
    End If
End Function

Public Sub computeInterfaceTotals()
    SheetUtilities.ClearInterfaceTotals
    
    Dim addresses As Scripting.Dictionary
    Set addresses = records.loadAddresses(AddressesSheet.Name)
    
    ' First key determines total service type, second key determines total type
    Dim totals As Scripting.Dictionary
    Set totals = New Scripting.Dictionary
    Dim totalServicename As TotalServiceType
    Dim totalTypename As TotalType
    For totalServicename = [_TotalServiceTypeFirst] To [_TotalServiceTypeLast]
        Dim totalTypes As Scripting.Dictionary
        Set totalTypes = New Scripting.Dictionary
        For totalTypename = [_TotalTypeFirst] To [_TotalTypeLast]
            ' initialized to 0
            Dim Quarters As Scripting.Dictionary
            Set Quarters = New Scripting.Dictionary
            Quarters.Item(1) = 0
            Quarters.Item(2) = 0
            Quarters.Item(3) = 0
            Quarters.Item(4) = 0
            totalTypes.Add totalTypename, SheetUtilities.cloneDict(Quarters)
        Next totalTypename
        totals.Add totalServicename, totalTypes
    Next totalServicename
    
    Dim uniqueNonDeliveryServices As Scripting.Dictionary
    Set uniqueNonDeliveryServices = New Scripting.Dictionary
    Dim uniqueDeliveryServices As Scripting.Dictionary
    Set uniqueDeliveryServices = New Scripting.Dictionary

    
    Dim appStatus As Variant
    If Application.StatusBar = False Then appStatus = False Else appStatus = Application.StatusBar
    
    Dim recordProgress As Long
    recordProgress = 1
    Application.StatusBar = "Totaling address 1 of " & UBound(addresses.Keys) + 1

    Dim key As Variant
    For Each key In addresses.Keys
        Dim record As RecordTuple
        Set record = addresses.Item(key)
        
        ' Gaithersburg totals
        If record.InCity = InCityCode.ValidInCity Then
            Dim quarter As Variant
            Dim qNum As Long
            For Each quarter In record.RxTotal.Keys
                qNum = getQuarterNum(quarter)
                Dim visit As Variant
                For Each visit In record.RxTotal.Item(quarter).Keys
                    ' Display Rx expenses for all delivery and non-delivery under non-delivery
                    totals.Item(nonDelivery).Item(rx)(qNum) = totals.Item(nonDelivery).Item(rx)(qNum) + _
                                                              record.RxTotal.Item(quarter).Item(visit)
                Next visit
            Next quarter
            
            ' Store whether a service type was present in each quarter
            ' so we can calculate later the number of double counts
            Dim deliveryServiceQuarters(1 To 4) As Boolean
            Dim nondeliveryServiceQuarters(1 To 4) As Boolean
            Dim i As Long
            For i = 1 To 4
                deliveryServiceQuarters(i) = False
                nondeliveryServiceQuarters(i) = False
            Next i
            
            ' Calculate totals for Delivery and Nondelivery
            ' Keep track of whether we've counted unduplicated for Delivery and Non Delivery
            ' using a dict of TotalServiceType : earliest quarter added
            Dim earliestQuarterAdded As Dictionary
            Set earliestQuarterAdded = New Scripting.Dictionary
            Dim service As Variant
            For Each service In record.visitData.Keys
                Dim serviceType As TotalServiceType
                serviceType = getServiceType(service)
                If serviceType = delivery Then
                    uniqueDeliveryServices.Item(service) = 1
                ElseIf InStr(1, service, "Rx Asst", vbTextCompare) > 0 Then
                    GoTo NextIteration
                Else
                    uniqueNonDeliveryServices.Item(service) = 1
                End If

                For Each quarter In record.visitData.Item(service).Keys
                    qNum = getQuarterNum(quarter)
                    
                    Dim count As Long
                    count = record.visitData.Item(service).Item(quarter).count
                    
                    If (count > 0) Then
                        ' Set service types visited in the quarter
                        If serviceType = delivery Then
                            deliveryServiceQuarters(qNum) = True
                        Else
                            nondeliveryServiceQuarters(qNum) = True
                        End If
                    
                        If Not earliestQuarterAdded.exists(serviceType) Then
                            totals.Item(serviceType).Item(uniqueGuestID)(qNum) = totals.Item(serviceType).Item(uniqueGuestID)(qNum) + _
                                                                                 1
                            totals.Item(serviceType).Item(uniqueGuestIDHousehold)(qNum) = totals.Item(serviceType).Item(uniqueGuestIDHousehold)(qNum) + _
                                                                                          record.householdTotal
                            earliestQuarterAdded.Add serviceType, qNum
                        ' Check for edge case where you get an earlier quarter after already counted due to services being out of order
                        ElseIf qNum < earliestQuarterAdded.Item(serviceType) Then
                            Dim laterQNum As Long
                            laterQNum = earliestQuarterAdded.Item(serviceType)
                            totals.Item(serviceType).Item(uniqueGuestID)(laterQNum) = totals.Item(serviceType).Item(uniqueGuestID)(laterQNum) - _
                                                                                      1
                            totals.Item(serviceType).Item(uniqueGuestIDHousehold)(laterQNum) = totals.Item(serviceType).Item(uniqueGuestIDHousehold)(laterQNum) - _
                                                                                               record.householdTotal
                            
                            totals.Item(serviceType).Item(uniqueGuestID)(qNum) = totals.Item(serviceType).Item(uniqueGuestID)(qNum) + _
                                                                                 1
                            totals.Item(serviceType).Item(uniqueGuestIDHousehold)(qNum) = totals.Item(serviceType).Item(uniqueGuestIDHousehold)(qNum) + _
                                                                                          record.householdTotal
                            earliestQuarterAdded.Item(serviceType) = qNum
                        End If
                    End If
                    totals.Item(serviceType).Item(nonUniqueGuestID)(qNum) = totals.Item(serviceType).Item(nonUniqueGuestID)(qNum) + _
                                                                            count
                    totals.Item(serviceType).Item(nonUniqueHousehold)(qNum) = totals.Item(serviceType).Item(nonUniqueHousehold)(qNum) + _
                                                                              count * record.householdTotal
                Next quarter
NextIteration:
            Next service
          
            ' Check for first time client changed services and increment double count total accordingly
            Dim visitedDelivery As Boolean
            visitedDelivery = False
            Dim visitedNonDelivery As Boolean
            visitedNonDelivery = False
            For i = 1 To 4
                If deliveryServiceQuarters(i) Then visitedDelivery = True
                If nondeliveryServiceQuarters(i) Then visitedNonDelivery = True
                
                If visitedDelivery And visitedNonDelivery Then
                    totals.Item(numDoubleCountedAdditionalDeliveryType).Item(uniqueGuestIDHousehold)(i) = totals.Item(numDoubleCountedAdditionalDeliveryType).Item(uniqueGuestIDHousehold)(i) + _
                                                                                                          record.householdTotal
                    Exit For
                End If
            Next i

        End If

        recordProgress = recordProgress + 1
        Application.StatusBar = "Totaling address " & recordProgress & " of " & UBound(addresses.Keys) + 1
    Next key
    
    Dim nonDeliveryTotalHeader As Range
    Set nonDeliveryTotalHeader = SheetUtilities.getNonDeliveryTotalHeaderRng()
    ' Necessary to avoid VBA compile error
    Dim clonedKeys() As Variant
    clonedKeys = uniqueNonDeliveryServices.Keys
    nonDeliveryTotalHeader.value = "Non-Delivery: " & Join(SheetUtilities.sortArr(clonedKeys), ",")
    
    Dim deliveryTotalHeader As Range
    Set deliveryTotalHeader = SheetUtilities.getDeliveryTotalHeaderRng()
    clonedKeys = uniqueDeliveryServices.Keys
    deliveryTotalHeader.value = "Delivery: " & Join(SheetUtilities.sortArr(clonedKeys), ",")
    
    
    For totalServicename = [_TotalServiceTypeFirst] To [_TotalServiceTypeLast]
        Dim totalsRng As Range
        Set totalsRng = SheetUtilities.getInterfaceTotalsRng(totalServicename)
        For i = 1 To 4
            If totalServicename = numDoubleCountedAdditionalDeliveryType Then
                totalsRng.Cells.Item(1, i) = totals.Item(totalServicename).Item(uniqueGuestIDHousehold)(i)
                GoTo NextQuarterForLoop
            End If
            
            totalsRng.Cells.Item(1, i) = totals.Item(totalServicename).Item(uniqueGuestID)(i)
            totalsRng.Cells.Item(2, i) = totals.Item(totalServicename).Item(uniqueGuestIDHousehold)(i)
            totalsRng.Cells.Item(3, i) = totals.Item(totalServicename).Item(nonUniqueGuestID)(i)
            totalsRng.Cells.Item(4, i) = totals.Item(totalServicename).Item(nonUniqueHousehold)(i)
        
            If totalServicename = nonDelivery Then
                totalsRng.Cells.Item(5, i) = totals.Item(totalServicename).Item(rx)(i)
            End If
NextQuarterForLoop:
        Next i
    Next totalServicename
    
    Application.StatusBar = appStatus
End Sub

Public Sub writeAddressesComputeInterfaceTotals(ByVal addresses As Scripting.Dictionary, _
                                       ByVal needsAutocorrect As Scripting.Dictionary, _
                                       ByVal discards As Scripting.Dictionary, _
                                       ByVal autocorrected As Scripting.Dictionary)
    
    SheetUtilities.ClearAllPreserveDate
    
    writeAddresses AddressesSheet.Name, addresses
    writeAddresses AutocorrectAddressesSheet.Name, needsAutocorrect
    writeAddresses DiscardsSheet.Name, discards
    writeAddresses AutocorrectedAddressesSheet.Name, autocorrected
    
    SortAll
    
    computeInterfaceTotals
End Sub

Public Sub addRecords()
    ' TODO import MicroTimer from Module 1
    ' Save application status bar to restore it later
    Dim appStatus As Variant
    If Application.StatusBar = False Then appStatus = False Else appStatus = Application.StatusBar
    
    Application.StatusBar = "Loading addresses"
        
    Dim addresses As Scripting.Dictionary
    Set addresses = loadAddresses(AddressesSheet.Name)
    
    Dim needsAutocorrect As Scripting.Dictionary
    Set needsAutocorrect = loadAddresses(AutocorrectAddressesSheet.Name)
    
    Dim discards As Scripting.Dictionary
    Set discards = loadAddresses(DiscardsSheet.Name)
    
    Dim autocorrected As Scripting.Dictionary
    Set autocorrected = loadAddresses(AutocorrectedAddressesSheet.Name)
       
    Dim recordsToValidate As Scripting.Dictionary
    Set recordsToValidate = New Scripting.Dictionary
    
    Dim i As Long
    i = getPastedInterfaceRecordsRng.row
    Do While i < getBlankRow(InterfaceSheet.Name).row
        Dim recordToAdd As RecordTuple
        Set recordToAdd = loadRecordFromRaw(InterfaceSheet.Range("A" & i))
        
        Dim existsInDict As Scripting.Dictionary
        Set existsInDict = Nothing
        
        Dim existingRecord As RecordTuple
        Set existingRecord = Nothing
        
        If recordsToValidate.exists(recordToAdd.key) Then
            Set existingRecord = recordsToValidate.Item(recordToAdd.key)
            '@Ignore FunctionReturnValueDiscarded
            existingRecord.MergeRecord recordToAdd
        Else
            If addresses.exists(recordToAdd.key) Then
                Set existsInDict = addresses
            ElseIf needsAutocorrect.exists(recordToAdd.key) Then
                Set existsInDict = needsAutocorrect
            ElseIf discards.exists(recordToAdd.key) Then
                Set existsInDict = discards
            End If
            
            Dim changedAddress As Boolean
            changedAddress = False
            
            If Not (existsInDict Is Nothing) Then
                Set existingRecord = existsInDict.Item(recordToAdd.key)
                '@Ignore FunctionReturnValueDiscarded
                changedAddress = existingRecord.MergeRecord(recordToAdd)
                
                If autocorrected.exists(recordToAdd.key) And Not changedAddress Then
                    Set existingRecord = autocorrected.Item(recordToAdd.key)
                    '@Ignore FunctionReturnValueDiscarded
                    existingRecord.MergeRecord recordToAdd
                End If
                
                Set recordToAdd = existingRecord
                If changedAddress Then
                    existsInDict.Remove recordToAdd.key
                    If autocorrected.exists(recordToAdd.key) Then
                        autocorrected.Remove recordToAdd.key
                    End If
                End If
            End If
            
            If changedAddress Or (existsInDict Is Nothing) Then
                If recordToAdd.isCorrectableAddress() Then
                    recordsToValidate.Add recordToAdd.key, recordToAdd
                Else
                    recordToAdd.SetInCity InCityCode.NotCorrectable
                    discards.Add recordToAdd.key, recordToAdd
                End If
            End If
        End If
        
        Application.StatusBar = "Adding record " & (i - 8) & " of " & (getBlankRow(InterfaceSheet.Name).row - 8)
        ' yield execution so Excel remains responsive and user can hit Esc
        DoEvents
        i = i + 1
    Loop
    
    ' Validate recordsToValidate
    i = 1
    Dim key As Variant
    For Each key In recordsToValidate.Keys
        Dim recordToValidate As RecordTuple
        Set recordToValidate = recordsToValidate.Item(key)
        Dim gburgAddress As Scripting.Dictionary
        Set gburgAddress = Lookup.gburgQuery(recordToValidate.GburgFormatRawAddress.Item(addressKey.Full))
        
        recordToValidate.SetValidAddress gburgAddress
        
        If gburgAddress.Item(addressKey.Full) <> vbNullString Then
            ' Valid address
            recordToValidate.SetInCity InCityCode.ValidInCity
            addresses.Add recordToValidate.key, recordToValidate
            ' NOTE choosing not to add to autocorrected since raw address was enough to match
            ' autocorrected is used to save Google lookups, Gaithersburg lookups are free
            ' However, Gaithersburg lookup can change zipcode, format of address, addition of Apt, etc.
        Else
            recordToValidate.SetInCity InCityCode.NotYetAutocorrected
            needsAutocorrect.Add recordToValidate.key, recordToValidate
        End If
        Application.StatusBar = "Validating record " & i & " of " & (UBound(recordsToValidate.Keys) + 1)
        i = i + 1
        DoEvents
    Next key

    Application.StatusBar = "Writing addresses and computing totals"
    
    writeAddressesComputeInterfaceTotals addresses, needsAutocorrect, discards, autocorrected
    computeCountyTotals
    
    Application.StatusBar = appStatus
End Sub


