Attribute VB_Name = "mDrop"
Option Explicit

Public Type tDropData
    ObjIndex As Integer
    Amount(1) As Integer
    Prob As Byte
End Type

Public Type tDrop
    Last As Byte
    Data() As tDropData
End Type

Public DropLast As Integer
Public DropData() As tDrop

Public Sub Drops_Load()
        '<EhHeader>
        On Error GoTo Drops_Load_Err
        '</EhHeader>
        Dim Manager As clsIniManager
        Dim A As Long, B As Long
        Dim Temp As String
    
100     Set Manager = New clsIniManager
            
          Dim FilePath As String
          FilePath = Drops_FilePath
102     Manager.Initialize (FilePath)
    
104     DropLast = val(Manager.GetValue("INIT", "LAST"))
    
106     ReDim DropData(1 To DropLast) As tDrop
    
108     For A = 1 To DropLast
110         With DropData(A)
112             .Last = val(Manager.GetValue(A, "LAST"))
            
114             ReDim .Data(1 To .Last) As tDropData
            
116             For B = 1 To .Last
118                 Temp = Manager.GetValue(A, B)
120                 .Data(B).ObjIndex = val(ReadField(1, Temp, 45))
                      .Data(B).Prob = val(ReadField(2, Temp, 45))
122                 .Data(B).Amount(0) = val(ReadField(3, Temp, 45))
124                 .Data(B).Amount(1) = val(ReadField(4, Temp, 45))
126             Next B
        
            End With
            
128     Next A

          Manager.DumpFile Drops_FilePath_Client
130     Set Manager = Nothing
    
        '<EhFooter>
        Exit Sub

Drops_Load_Err:
        Set Manager = Nothing
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mDrop.Drops_Load " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
