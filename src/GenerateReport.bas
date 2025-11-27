Attribute VB_Name = "GenerateReport"
'@Folder("City_Grant_Address_Report.src")
Option Explicit

Private Sub writeNonRxReportRecord(ByVal record As RecordTuple)
    Dim row As Range
    
    Set row = SheetUtilities.getBlankRow(NonRxReportSheet.Name)
    
    row.Cells.Item(1, 1) = "Gaithersburg HELP"
    row.Cells.Item(1, 2) = record.GburgFormatValidAddress.Item(addressKey.streetNum)
    row.Cells.Item(1, 3) = record.GburgFormatValidAddress.Item(addressKey.PrefixedStreetname)
    row.Cells.Item(1, 4) = record.GburgFormatValidAddress.Item(addressKey.StreetType)
    row.Cells.Item(1, 5) = record.GburgFormatValidAddress.Item(addressKey.unitType)
    row.Cells.Item(1, 6) = record.GburgFormatValidAddress.Item(addressKey.unitNum)
    row.Cells.Item(1, 7) = "Gaithersburg"
    row.Cells.Item(1, 8) = "MD"
    row.Cells.Item(1, 9) = record.CleanInitials
    row.Cells.Item(1, 10) = record.householdTotal
    row.Cells.Item(1, 11) = record.eighteenPlusTotal
    row.Cells.Item(1, 12) = record.zeroToOneTotal + record.twoToSeventeenTotal
    
    Dim Quarters() As Boolean
    Quarters = record.Quarters(True)
    If Quarters(1) Then row.Cells.Item(1, 13) = "x"
    If Quarters(2) Then row.Cells.Item(1, 14) = "x"
    If Quarters(3) Then row.Cells.Item(1, 15) = "x"
    If Quarters(4) Then row.Cells.Item(1, 16) = "x"
    
    Dim Months(1 To 12) As Boolean
    Dim svc As Variant
    Dim q As Variant
    Dim v As Variant
    Dim d As Date
    Dim m As Long
    
    ' Walk all services, quarters, and visit dates
    For Each svc In record.visitData.Keys
        If svc <> "Rx Asst" Then
            For Each q In record.visitData.Item(svc).Keys
                For Each v In record.visitData.Item(svc).Item(q)
                    If Not IsEmpty(v) And v <> vbNullString Then
                        d = CDate(v)
                        m = Month(d)
                        If m >= 1 And m <= 12 Then
                            Months(m) = True
                        End If
                    End If
                Next v
            Next q
        End If
    Next svc
    
    Const FIRST_MONTH_COL As Long = 17
    
    For m = 1 To 12
        If Months(m) Then
            row.Cells.Item(1, FIRST_MONTH_COL + (m - 1)) = "x"
        End If
    Next m
End Sub

Public Sub generateNonRxReport()
    SheetUtilities.getNonRxReportRng(True).Clear
    
    Dim addresses As Scripting.Dictionary
    Set addresses = records.loadAddresses(AddressesSheet.Name)
    
    Dim address As Variant
    For Each address In addresses.Keys()
        Dim record As RecordTuple
        Set record = addresses.Item(address)
        
        If (record.InCity = ValidInCity) And _
           (record.visitData.count > 0) And _
           (Not record.isRxOnly) Then
           writeNonRxReportRecord record
        End If
    Next address
    
    SheetUtilities.SortSheet NonRxReportSheet.Name
    
    ActiveSheet.Range("A2").Select
End Sub

Public Sub generateRxReport(ByVal records As RxRecords, ByVal addresses As Scripting.Dictionary)
    Dim Name As Variant
    For Each Name In records.guestNames()
        Dim record As RxRecord
        Set record = records.guestRecord(Name)
        
        Dim addressRecord As RecordTuple
        Set addressRecord = addresses.Item(record.guestID)
        
        Dim row As Range
        Set row = SheetUtilities.getBlankRow(RxReportSheet.Name)
        
        row.Cells.Item(1, 1) = "Gaithersburg HELP"
        row.Cells.Item(1, 2) = addressRecord.GburgFormatValidAddress.Item(addressKey.streetNum)
        row.Cells.Item(1, 3) = addressRecord.GburgFormatValidAddress.Item(addressKey.PrefixedStreetname)
        row.Cells.Item(1, 4) = addressRecord.GburgFormatValidAddress.Item(addressKey.StreetType)
        row.Cells.Item(1, 5) = addressRecord.GburgFormatValidAddress.Item(addressKey.unitType)
        row.Cells.Item(1, 6) = addressRecord.GburgFormatValidAddress.Item(addressKey.unitNum)
        row.Cells.Item(1, 7) = "Gaithersburg"
        row.Cells.Item(1, 8) = "MD"
        row.Cells.Item(1, 9) = CleanInitials(Name)
        
        If record.quarter(q1) Then row.Cells.Item(1, 10) = "x"
        If record.quarter(q2) Then row.Cells.Item(1, 11) = "x"
        If record.quarter(q3) Then row.Cells.Item(1, 12) = "x"
        If record.quarter(q4) Then row.Cells.Item(1, 13) = "x"
        
    Next Name
    
    SheetUtilities.SortSheet RxReportSheet.Name
    
    RxReportSheet.Range("A2").Select
End Sub
