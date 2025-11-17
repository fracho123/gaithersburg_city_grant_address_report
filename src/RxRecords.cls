VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RxRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("City_Grant_Address_Report.src")
'@IgnoreModule ProcedureNotUsed
'ProcedureNotUsed is Rubberduck bug
Option Explicit

Private pRecords As Scripting.Dictionary

Private Property Get records() As Scripting.Dictionary
    If pRecords Is Nothing Then
        Set pRecords = New Scripting.Dictionary
    End If
    Set records = pRecords
End Property

Public Property Get guestNames() As Variant()
    guestNames = records.Keys
End Property

Public Property Get exists(ByVal Name As String) As Boolean
    exists = records.exists(Name)
End Property

Public Property Get guestRecord(ByVal Name As String) As RxRecord
    Set guestRecord = records.Item(Name)
End Property

Public Property Let guestRecord(ByVal Name As String, ByVal record As RxRecord)
    If exists(Name) Then
        'Rubberduck bug
        '@Ignore ValueRequired
        records.Item(Name) = record
    Else
        records.Add Name, record
    End If
End Property
