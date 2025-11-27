VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RecordTuple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "City_Grant_Address_Report.src"
Option Explicit

Public Enum InCityCode
    NotCorrectable = -2
    NotYetAutocorrected = -1
    FailedAutocorrectNotInCity = 0
    FailedAutocorrectInCity = 1
    ValidInCity = 2
    ValidNotInCity = 3
End Enum

Private pInCity As InCityCode

Private pUserVerified As Boolean

Private pValidAddress As String
Private pValidUnitWithNum As String
Private pValidZipcode As String

Private pRawAddress As String
Private pRawUnitWithNum As String
Private pRawCity As String
Private pRawState As String
Private pRawZip As String
Private pGuestID As String
Private pFirstName As String
Private pLastName As String

'BUG does not track household total by quarter
Private pHouseholdTotal As Long
Private pZeroToOneTotal As Long
Private pTwoToSeventeenTotal As Long
Private pEighteenPlusTotal As Long

' Quarter key: (Visit date CStr(Date) key: RxTotal Double)
Private pRxTotal As Scripting.Dictionary

' Service key: (Quarter key: (Visit date CStr(CDate) Collection)
Private pVisitData As Scripting.Dictionary

Private pGburgFormatRawAddress As Scripting.Dictionary
Private pGburgFormatValidAddress As Scripting.Dictionary

Public Property Get key() As String
    key = pGuestID
End Property

Public Property Get InCity() As InCityCode
    InCity = pInCity
End Property

Public Property Get InCityStr() As String
    Select Case pInCity
        Case InCityCode.NotCorrectable
            InCityStr = "Not correctable"
        Case InCityCode.NotYetAutocorrected
            InCityStr = "Not yet autocorrected"
        Case InCityCode.FailedAutocorrectNotInCity
            InCityStr = "Failed autocorrection and geocoded not in city"
        Case InCityCode.FailedAutocorrectInCity
            InCityStr = "Possible but failed autocorrection"
        Case InCityCode.ValidInCity
            InCityStr = "Yes"
        Case InCityCode.ValidNotInCity
            InCityStr = "No"
    End Select
End Property

' Should be String or InCityCode
' VBA properties don't allow overloading
Public Sub SetInCity(ByVal code As Variant)
    Select Case VarType(code)
        Case vbLong
            pInCity = code
            ' if attempted revalidation but failed, set userverified
            If code = InCityCode.FailedAutocorrectInCity Then UserVerified = False
        Case vbString
            If code = "Not correctable" Then
                pInCity = NotCorrectable
            ElseIf code = "Not yet autocorrected" Then
                pInCity = NotYetAutocorrected
            ElseIf code = "Failed autocorrection and geocoded not in city" Then
                pInCity = FailedAutocorrectNotInCity
            ElseIf code = "Possible but failed autocorrection" Then
                pInCity = FailedAutocorrectInCity
            ElseIf code = "Yes" Then
                pInCity = ValidInCity
            ElseIf code = "No" Then
                pInCity = ValidNotInCity
            End If
    End Select
End Sub

Public Property Get UserVerified() As Boolean
    UserVerified = pUserVerified
End Property

Public Property Let UserVerified(ByVal bool As Boolean)
    pUserVerified = bool
End Property

Public Property Get validAddress() As String
    validAddress = pValidAddress
End Property

Public Property Let validAddress(ByVal str As String)
    pValidAddress = str
End Property

Public Property Get validUnitWithNum() As String
    validUnitWithNum = pValidUnitWithNum
End Property

Public Property Let validUnitWithNum(ByVal str As String)
    pValidUnitWithNum = str
End Property

Public Property Get ValidZipcode() As String
    ValidZipcode = pValidZipcode
End Property

Public Property Let ValidZipcode(ByVal str As String)
    pValidZipcode = str
End Property

' Expects AddressKey StreetAddress, UnitType, UnitNum, Zip
Public Sub SetValidAddress(ByVal address As Scripting.Dictionary)
    ' USPS data is in all caps
    ' NOTE This turns O'Neill into O'neill but that shouldn't matter for Gburg db search
    validAddress = StrConv(address.Item(addressKey.streetAddress), vbProperCase)
    validUnitWithNum = StrConv(address.Item(addressKey.unitWithNum), vbProperCase)
    ValidZipcode = StrConv(address.Item(addressKey.zip), vbProperCase)
End Sub

Public Property Get RawAddress() As String
    RawAddress = pRawAddress
End Property

Public Property Let RawAddress(ByVal str As String)
    pRawAddress = str
End Property

Public Property Get RawUnitWithNum() As String
    RawUnitWithNum = pRawUnitWithNum
End Property

Public Property Let RawUnitWithNum(ByVal str As String)
    pRawUnitWithNum = str
End Property

Public Property Get isAutocorrected() As Boolean
    isAutocorrected = ((validAddress <> vbNullString) And _
                            validAddress <> StrConv(RawAddress, vbProperCase)) Or _
                      ((validUnitWithNum <> vbNullString) And _
                            validUnitWithNum <> StrConv(RawUnitWithNum, vbProperCase)) Or _
                      ((ValidZipcode <> vbNullString) And (ValidZipcode <> RawZip) And (RawZip <> vbNullString))
End Property

Public Property Get RawCity() As String
    RawCity = pRawCity
End Property

Public Property Let RawCity(ByVal str As String)
    pRawCity = str
End Property

Public Property Get RawState() As String
    RawState = pRawState
End Property

Public Property Let RawState(ByVal str As String)
    pRawState = str
End Property

Public Property Get RawZip() As String
    RawZip = pRawZip
End Property

Public Property Let RawZip(ByVal str As String)
    pRawZip = str
End Property

Public Property Get guestID() As String
    guestID = pGuestID
End Property

Public Property Let guestID(ByVal str As String)
    pGuestID = str
End Property

Public Property Get FirstName() As String
    FirstName = pFirstName
End Property

Public Property Let FirstName(ByVal str As String)
    pFirstName = str
End Property

Public Property Get LastName() As String
    LastName = pLastName
End Property

Public Property Let LastName(ByVal str As String)
    pLastName = str
End Property

Public Property Get householdTotal() As Long
    householdTotal = pHouseholdTotal
End Property

Public Property Let householdTotal(ByVal num As Long)
    pHouseholdTotal = num
End Property

Public Property Get zeroToOneTotal() As Long
    zeroToOneTotal = pZeroToOneTotal
End Property

Public Property Let zeroToOneTotal(ByVal num As Long)
    pZeroToOneTotal = num
End Property

Public Property Get twoToSeventeenTotal() As Long
    twoToSeventeenTotal = pTwoToSeventeenTotal
End Property

Public Property Let twoToSeventeenTotal(ByVal num As Long)
    pTwoToSeventeenTotal = num
End Property

Public Property Get eighteenPlusTotal() As Long
    eighteenPlusTotal = pEighteenPlusTotal
End Property

Public Property Let eighteenPlusTotal(ByVal num As Long)
    pEighteenPlusTotal = num
End Property


Public Property Get RxTotal() As Scripting.Dictionary
    If pRxTotal Is Nothing Then Set pRxTotal = New Scripting.Dictionary
    Set RxTotal = pRxTotal
End Property

Public Property Set RxTotal(ByVal dict As Scripting.Dictionary)
    Set pRxTotal = dict
End Property

Public Property Get visitData() As Scripting.Dictionary
    If pVisitData Is Nothing Then Set pVisitData = New Scripting.Dictionary
    Set visitData = pVisitData
End Property

Public Property Set visitData(ByVal dict As Scripting.Dictionary)
    Set pVisitData = dict
End Property

Public Sub AddVisit(ByVal visitDate As String, ByVal service As String)
    Dim newVisitData As Scripting.Dictionary
    Set newVisitData = New Scripting.Dictionary
    
    Dim serviceVisits As Scripting.Dictionary
    Set serviceVisits = New Scripting.Dictionary
    
    Dim visits As Collection
    Set visits = New Collection
    
    ' NOTE Adding Date as dictionary key silently changes string formatting on print for some reason
    ' CStr formats Date more nicely for readability
    ' and prevents issues like 4/17/2023 not matching 2023-04-17T04:00:00.000Z
    visits.Add CStr(CDate(visitDate))
    serviceVisits.Add getQuarterStr(visitDate), visits
    newVisitData.Add service, serviceVisits
    
    MergeVisitData newVisitData
End Sub

Public Sub addRx(ByVal visitDate As String, ByVal rx As Double)
    Dim newRx As Scripting.Dictionary
    Set newRx = New Scripting.Dictionary
    
    Dim rxVisits As Scripting.Dictionary
    Set rxVisits = New Scripting.Dictionary
    
    ' NOTE as above use CStr(CDate)
    rxVisits.Add CStr(CDate(visitDate)), rx
    newRx.Add getQuarterStr(visitDate), rxVisits
    
    MergeRx newRx
End Sub

' Returns true if address changed, false otherwise
Public Function MergeRecord(ByVal record As RecordTuple) As Boolean
    MergeVisitData record.visitData
    MergeRx record.RxTotal
    
    FirstName = record.FirstName
    LastName = record.LastName
    householdTotal = record.householdTotal
    zeroToOneTotal = record.zeroToOneTotal
    twoToSeventeenTotal = record.twoToSeventeenTotal
    eighteenPlusTotal = record.eighteenPlusTotal
        
    Dim addressChanged As Boolean
    addressChanged = (record.RawAddress <> RawAddress) Or _
                     (record.RawUnitWithNum <> RawUnitWithNum) Or _
                     (record.RawCity <> RawCity) Or _
                     (record.RawState <> RawState) Or _
                     (record.RawZip <> RawZip)
    
    RawAddress = record.RawAddress
    RawUnitWithNum = record.RawUnitWithNum
    RawCity = record.RawCity
    RawState = record.RawState
    RawZip = record.RawZip
    
    If addressChanged Then
        UserVerified = False
        SetInCity InCityCode.NotYetAutocorrected
        validAddress = vbNullString
        validUnitWithNum = vbNullString
        ValidZipcode = vbNullString
        
        Set pGburgFormatRawAddress = Nothing
        Set pGburgFormatValidAddress = Nothing
    End If
    
    MergeRecord = addressChanged
End Function

Private Sub MergeRx(ByVal rxToMerge As Scripting.Dictionary)
    Dim quarter As Variant
    For Each quarter In rxToMerge.Keys
        If RxTotal.exists(quarter) Then
            Dim visit As Variant
            For Each visit In rxToMerge.Item(quarter).Keys
                If Not RxTotal.Item(quarter).exists(visit) Then
                    RxTotal.Item(quarter).Add visit, rxToMerge.Item(quarter).Item(visit)
                End If
            Next visit
        Else
            RxTotal.Add quarter, rxToMerge.Item(quarter)
        End If
    Next quarter
End Sub

Private Sub MergeVisitData(ByVal visitDataToMerge As Scripting.Dictionary)
    Dim service As Variant
    
    For Each service In visitDataToMerge.Keys
        If visitData.exists(service) Then
            ' Loop through quarters and merge
            Dim quarter As Variant
            For Each quarter In visitDataToMerge.Item(service)
                If visitData.Item(service).exists(quarter) Then
                    Dim visitDateToMerge As Variant
                    For Each visitDateToMerge In visitDataToMerge.Item(service).Item(quarter)
                        Dim visitExists As Boolean
                        Dim existingVisitDate As Variant
                        For Each existingVisitDate In visitData.Item(service).Item(quarter)
                            If visitDateToMerge = existingVisitDate Then
                                visitExists = True
                                Exit For
                            End If
                        Next existingVisitDate
                        
                        If Not visitExists Then
                            visitData.Item(service).Item(quarter).Add visitDateToMerge
                        End If
                    Next visitDateToMerge
                Else
                    visitData.Item(service).Add quarter, visitDataToMerge.Item(service).Item(quarter)
                End If
            Next quarter
        Else
            visitData.Add service, visitDataToMerge.Item(service)
        End If
    Next service
End Sub

' Returns 1-based array of whether record has quarter or not
' Can exclude Rx records
Public Property Get Quarters(Optional ByVal excludeRx As Boolean = False) As Boolean()
    Dim quarterArr(1 To 4) As Boolean
    
    Dim service As Variant
    For Each service In visitData
        If Not (excludeRx And (service = "Rx Asst")) Then
            Dim quarter As Variant
            For Each quarter In visitData.Item(service)
                Select Case quarter
                    Case "Q1"
                        quarterArr(1) = True
                    Case "Q2"
                        quarterArr(2) = True
                    Case "Q3"
                        quarterArr(3) = True
                    Case "Q4"
                        quarterArr(4) = True
                End Select
            Next quarter
        End If
    Next service

    Quarters = quarterArr
End Property


Public Property Get CleanFirstName() As String
    CleanFirstName = CleanString(FirstName)
End Property

Public Property Get CleanLastName() As String
    Dim properLastName As String
    properLastName = CleanString(LastName)
    ' Last names beginning with "Mc" should have next letter capitalized
    If Left$(properLastName, 2) = "Mc" Then
        properLastName = "Mc" & UCase$(Mid$(properLastName, 3, 1)) & Mid$(properLastName, 4, 99)
    End If
    CleanLastName = properLastName
End Property

Public Property Get cleanName() As String
    cleanName = CleanFirstName & " " & CleanLastName
End Property

Public Property Get CleanInitials() As String
    CleanInitials = SheetUtilities.CleanInitials(cleanName)
End Property

' Checks street name from start position onwards for an abbreviated direction with spaces around it
' If direction is found, returns it, otherwise returns ""
'@Ignore AssignedByValParameter
Private Function GetDirection(ByVal start As Long, ByVal street As String) As String
    If (street = "  ") Then
        GetDirection = vbNullString
        Exit Function
    End If
    If (start < 1) Then
        start = 1
    End If
    
    ' Add spaces to avoid e.g. Westland
    street = " " & street & " "
    If InStr(start, street, " NE ", vbTextCompare) Then
        GetDirection = "NE"
    ElseIf InStr(start, street, " NW ", vbTextCompare) Then
        GetDirection = "NW"
    ElseIf InStr(start, street, " SW ", vbTextCompare) Then
        GetDirection = "SW"
    ElseIf InStr(start, street, " SE ", vbTextCompare) Then
        GetDirection = "SE"
    ElseIf InStr(start, street, " N ", vbTextCompare) Then
        GetDirection = "N"
    ElseIf InStr(start, street, " E ", vbTextCompare) Then
        GetDirection = "E"
    ElseIf InStr(start, street, " S ", vbTextCompare) Then
        GetDirection = "S"
    ElseIf InStr(start, street, " W ", vbTextCompare) Then
        GetDirection = "W"
    Else
        GetDirection = vbNullString
    End If
End Function

' Splits unit into zero-based array of abbreviated unit type, unit num, additional words
' If only one word, assumes unit is Apt
'@Ignore AssignedByValParameter
Private Function SplitUnit(ByVal unitWithNum As String) As String()
    unitWithNum = CleanString(unitWithNum)
    
    ' https://pe.usps.com/text/pub28/28apc_003.htm#ep538629
    unitWithNum = Replace(unitWithNum, "Apartment", "Apt")
    unitWithNum = Replace(unitWithNum, "Building", "Bldg")
    unitWithNum = Replace(unitWithNum, "Basement", "Bsmt")
    unitWithNum = Replace(unitWithNum, "Department", "Dept")
    unitWithNum = Replace(unitWithNum, "Floor", "Fl")
    unitWithNum = Replace(unitWithNum, "Front", "Frnt")
    unitWithNum = Replace(unitWithNum, "Hanger", "Hngr")
    unitWithNum = Replace(unitWithNum, "Key", "Key")
    unitWithNum = Replace(unitWithNum, "Lobby", "Lbby")
    unitWithNum = Replace(unitWithNum, "Lot", "Lot")
    unitWithNum = Replace(unitWithNum, "Lower", "Lowr")
    unitWithNum = Replace(unitWithNum, "Ofc", "Office")
    unitWithNum = Replace(unitWithNum, "Penthouse", "Ph")
    unitWithNum = Replace(unitWithNum, "Pier", "Pier")
    unitWithNum = Replace(unitWithNum, "Rear", "Rear")
    unitWithNum = Replace(unitWithNum, "Room", "Rm")
    unitWithNum = Replace(unitWithNum, "Side", "Side")
    unitWithNum = Replace(unitWithNum, "Slip", "Slip")
    unitWithNum = Replace(unitWithNum, "Space", "Spc")
    unitWithNum = Replace(unitWithNum, "Suite", "Ste")
    unitWithNum = Replace(unitWithNum, "Stop", "Stop")
    unitWithNum = Replace(unitWithNum, "Trailer", "Trlr")
    unitWithNum = Replace(unitWithNum, "Unit", "Unit")
    unitWithNum = Replace(unitWithNum, "Upper", "Uppr")
    
    ' If # assume Apt
    unitWithNum = Replace(unitWithNum, "# ", "Apt ")
    unitWithNum = Replace(unitWithNum, "#", "Apt ")
    
    Dim splitArr() As String
    splitArr = Split(unitWithNum, " ")
    
    Select Case UBound(splitArr)
        Case 0
            ' several designators such as upper don't have a secondary number
            ' however, as of 2/8/24 Gaithersburg only has Unit, Bldg, Fl, Ste, Apt
            ' so not going to worry about those
            Select Case splitArr(0)
                Case "Bsmt", "Frnt", "Lbby", "Lowr", "Ofc", "Ph", "Rear", "Side", "Uppr"
                    SplitUnit = Split(splitArr(0) & ",", ",")
                Case Else ' assume Apt
                    SplitUnit = Split("Apt," & unitWithNum, ",")
            End Select
        Case -1
            ' unitWithNum was ""
            SplitUnit = Split(",", ",")
        Case 1
            ' 2 words
            SplitUnit = splitArr
        Case 2
            ' 3 words, assume Apt e.g. 103 Rm 1
            SplitUnit = Split("Apt " & unitWithNum, " ")
        Case Else
            SplitUnit = splitArr
    End Select
End Function

' Returns dictionary of AddressKey full, streetNum, prefixedstreetname, streettype, postfix, unit type, unit num, streetaddress
' - capitalized Street Name with abbr. prefix, cap. abbr. Street Type, cap. abbr. Postfix
'@Ignore AssignedByValParameter
Private Function getFormattedSplitAddr(ByVal returnRaw As Boolean) As Scripting.Dictionary
    Dim addr As String
    Dim unitWithNum As String

    If returnRaw Then
        addr = RawAddress
        unitWithNum = RawUnitWithNum
    Else
        addr = validAddress
        unitWithNum = validUnitWithNum
    End If
    
    addr = CleanString(addr)
    '@Ignore AssignmentNotUsed
    ' probably this bug: https://github.com/rubberduck-vba/Rubberduck/issues/6142
    unitWithNum = CleanString(unitWithNum)
    
    ' BUG these will not show up in autocorrected
    ' - Gaithersburg has O'Neill and Odend'hal, but Soxbox exports as Oneill (and Oneil misspelling) and Odendhal
    ' - Mccausland Pl renamed to Prism Pl
    addr = Replace(addr, "Oneill", "O'neill")
    addr = Replace(addr, "Oneil", "O'neill")
    addr = Replace(addr, "Odendhal", "Odend'hal")
    addr = Replace(addr, "Mccausland", "Prism")
    
    ' Add spaces to left and right of street name (temporarily) for replace operations (avoid e.g. Westland Dr)
    addr = " " & addr & " "
    addr = Replace(addr, " Northwest ", " NW ")
    addr = Replace(addr, " Southwest ", " SW ")
    addr = Replace(addr, " Northeast ", " NE ")
    addr = Replace(addr, " Southeast ", " SE ")
    addr = Replace(addr, " North ", " N ")
    addr = Replace(addr, " South ", " S ")
    addr = Replace(addr, " East ", " E ")
    addr = Replace(addr, " West ", " W ")
    addr = Trim$(addr)
    
    addr = Replace(addr, "Avenue", "Ave")
    addr = Replace(addr, "Boulevard", "Blvd")
    addr = Replace(addr, "Circle", "Cir")
    addr = Replace(addr, "Close", "Cls")
    addr = Replace(addr, "Court", "Ct")
    addr = Replace(addr, "Drive", "Dr")
    addr = Replace(addr, "Highway", "Hwy")
    addr = Replace(addr, "Lane", "Ln")
    addr = Replace(addr, "Mews", "Mews")
    addr = Replace(addr, "Parkway", "Pkwy")
    addr = Replace(addr, "Place", "Pl")
    addr = Replace(addr, "Road", "Rd")
    ' No 'Square' street type in Gaithersburg but it is present in street name
    'addr = Replace(addr, "Square", "Sq")
    addr = Replace(addr, "Street", "St")
    addr = Replace(addr, "Terrace", "Ter")
    '@Ignore AssignmentNotUsed
    ' probably this bug https://github.com/rubberduck-vba/Rubberduck/issues/6142
    addr = Replace(addr, "Way", "Way")

    Dim splitAddr As Scripting.Dictionary
    Set splitAddr = New Scripting.Dictionary

    Dim splitNumArr() As String
    ' Split Street Number from Street name
    splitNumArr = LWordTrim(addr)
    splitAddr.Add addressKey.streetNum, splitNumArr(0)
    
    ' Remove periods, hyphens, not in street name
    ' see https://pe.usps.com/text/pub28/28c2_013.htm & https://pe.usps.com/text/dmm100/addressing-mail.htm
    splitNumArr(1) = Replace(splitNumArr(1), ".", vbNullString)
    splitNumArr(1) = Replace(splitNumArr(1), "-", vbNullString)
    
    Dim fullAddress As String
    fullAddress = splitNumArr(0) & " " & splitNumArr(1)
    splitAddr.Add addressKey.streetAddress, fullAddress
    
    Dim streetPostfix As String
    streetPostfix = GetDirection(Len(addr) - 3, addr)
    Dim splitTypeArr() As String
    ' Remove postfix if it exists
    If streetPostfix <> vbNullString Then
        splitTypeArr = Lookup.RWordTrim(Lookup.RWordTrim(splitNumArr(1))(0))
    Else
        splitTypeArr = Lookup.RWordTrim(splitNumArr(1))
    End If
    
    splitAddr.Add addressKey.PrefixedStreetname, splitTypeArr(0)
    splitAddr.Add addressKey.StreetType, splitTypeArr(1)
    splitAddr.Add addressKey.Postfix, streetPostfix
    
    Dim splitUnitArr() As String
    splitUnitArr = SplitUnit(unitWithNum)
    
    splitAddr.Add addressKey.unitType, splitUnitArr(0)
    ' NOTE remove unit from array and join the rest because sometimes will be given
    ' multiple secondary numbers e.g. Apt 102 # unt
    splitUnitArr(0) = vbNullString
    splitAddr.Add addressKey.unitNum, Trim$(Join(splitUnitArr, " "))
    
    If splitAddr.Item(addressKey.unitType) <> vbNullString Then
        splitAddr.Add addressKey.Full, fullAddress & " " & splitAddr.Item(addressKey.unitType) & _
                      " " & splitAddr.Item(addressKey.unitNum)
    Else
        splitAddr.Add addressKey.Full, fullAddress
    End If
    
    Set getFormattedSplitAddr = splitAddr
End Function

Public Property Get GburgFormatRawAddress() As Scripting.Dictionary
    If pGburgFormatRawAddress Is Nothing Then
        Set pGburgFormatRawAddress = getFormattedSplitAddr(True)
    End If
    Set GburgFormatRawAddress = pGburgFormatRawAddress
End Property

Public Property Get GburgFormatValidAddress() As Scripting.Dictionary
    If pGburgFormatValidAddress Is Nothing Then
        Set pGburgFormatValidAddress = getFormattedSplitAddr(False)
    End If
    Set GburgFormatValidAddress = pGburgFormatValidAddress
End Property

Public Function isCorrectableAddress() As Boolean
    Dim splitAddr() As String
    splitAddr = Split(Trim$(RawAddress), " ")
    Select Case UBound(splitAddr)
        Case -1 'null
            isCorrectableAddress = False
        Case Else
            Dim hasNumber As Boolean
            Dim hasLetter As Boolean
            Dim i As Long
            For i = 1 To Len(RawAddress)
                Dim char As String
                char = Mid$(RawAddress, i, 1)
                If IsNumeric(char) Then hasNumber = True
                If Not IsNumeric(char) Then hasLetter = True
                If hasNumber And hasLetter Then Exit For
            Next
            isCorrectableAddress = hasNumber And hasLetter
    End Select
End Function

Public Function isRxOnly() As Boolean
    Dim services As Variant
    services = visitData.Keys
    If UBound(services) = -1 Then
        isRxOnly = False
    Else
        isRxOnly = (visitData.count = 1) And (services(0) = "Rx Asst")
    End If
End Function
