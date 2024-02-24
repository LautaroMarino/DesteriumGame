Attribute VB_Name = "mPremium"
Option Explicit

Public Type tPremium
    ObjIndex As Integer
    Amount As Integer
    RequiredObj As Long
    RequiredAmount As Long
End Type

Public Premiums() As tPremium
Public Premium_Last As Integer


Public Sub Premiums_Load()
        '<EhHeader>
        On Error GoTo Premiums_Load_Err
        '</EhHeader>
        Dim Read As clsIniManager
        Dim A As Long
        Dim Temp As String
    
100     Set Read = New clsIniManager
    
    
102     Read.Initialize (DatPath & "PREMIUM.DAT")
    
    
104     Premium_Last = val(Read.GetValue("INIT", "LAST"))
    
106     ReDim Premiums(1 To Premium_Last) As tPremium
    
    
108     For A = 1 To Premium_Last
110         With Premiums(A)
112             Temp = Read.GetValue("LIST", A)
114             .ObjIndex = val(ReadField(1, Temp, Asc("-")))
116             .Amount = val(ReadField(2, Temp, Asc("-")))
118             .RequiredAmount = val(ReadField(3, Temp, Asc("-")))
120             .RequiredObj = 1466
            End With
122     Next A
    
    
124     Set Read = Nothing
        '<EhFooter>
        Exit Sub

Premiums_Load_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mPremium.Premiums_Load " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

