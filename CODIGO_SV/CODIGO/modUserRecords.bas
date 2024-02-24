Attribute VB_Name = "modUserRecords"
'Argentum Online 0.13.0
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Public Sub LoadRecords()
        '<EhHeader>
        On Error GoTo LoadRecords_Err
        '</EhHeader>

        '**************************************************************
        'Author: Amraphen
        'Last Modify Date: 29/11/2010
        'Carga los seguimientos de usuarios.
        '**************************************************************
        Dim Reader As clsIniManager

        Dim tmpStr As String

        Dim i      As Long

        Dim j      As Long

100     Set Reader = New clsIniManager
    
102     If Not FileExist(DatPath & "RECORDS.DAT") Then
104         Call CreateRecordsFile
        End If
    
106     Call Reader.Initialize(DatPath & "RECORDS.DAT")

108     NumRecords = Reader.GetValue("INIT", "NumRecords")

110     If NumRecords Then ReDim Records(1 To NumRecords)
    
112     For i = 1 To NumRecords

114         With Records(i)
116             .Usuario = Reader.GetValue("RECORD" & i, "Usuario")
118             .Creador = Reader.GetValue("RECORD" & i, "Creador")
120             .Fecha = Reader.GetValue("RECORD" & i, "Fecha")
122             .Motivo = Reader.GetValue("RECORD" & i, "Motivo")

124             .NumObs = val(Reader.GetValue("RECORD" & i, "NumObs"))

126             If .NumObs Then ReDim .Obs(1 To .NumObs)
            
128             For j = 1 To .NumObs
130                 tmpStr = Reader.GetValue("RECORD" & i, "Obs" & j)
                
132                 .Obs(j).Creador = ReadField(1, tmpStr, 45)
134                 .Obs(j).Fecha = ReadField(2, tmpStr, 45)
136                 .Obs(j).Detalles = ReadField(3, tmpStr, 45)
138             Next j

            End With

140     Next i

        '<EhFooter>
        Exit Sub

LoadRecords_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modUserRecords.LoadRecords " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub SaveRecords()
        '<EhHeader>
        On Error GoTo SaveRecords_Err
        '</EhHeader>

        '**************************************************************
        'Author: Amraphen
        'Last Modify Date: 29/11/2010
        'Guarda los seguimientos de usuarios.
        '**************************************************************
        Dim Writer As clsIniManager

        Dim tmpStr As String

        Dim i      As Long

        Dim j      As Long

100     Set Writer = New clsIniManager

102     Call Writer.ChangeValue("INIT", "NumRecords", NumRecords)
    
104     For i = 1 To NumRecords

106         With Records(i)
108             Call Writer.ChangeValue("RECORD" & i, "Usuario", .Usuario)
110             Call Writer.ChangeValue("RECORD" & i, "Creador", .Creador)
112             Call Writer.ChangeValue("RECORD" & i, "Fecha", .Fecha)
114             Call Writer.ChangeValue("RECORD" & i, "Motivo", .Motivo)
            
116             Call Writer.ChangeValue("RECORD" & i, "NumObs", .NumObs)
            
118             For j = 1 To .NumObs
120                 tmpStr = .Obs(j).Creador & "-" & .Obs(j).Fecha & "-" & .Obs(j).Detalles
122                 Call Writer.ChangeValue("RECORD" & i, "Obs" & j, tmpStr)
124             Next j

            End With

126     Next i
    
128     Call Writer.DumpFile(DatPath & "RECORDS.DAT")
        '<EhFooter>
        Exit Sub

SaveRecords_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modUserRecords.SaveRecords " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub AddRecord(ByVal UserIndex As Integer, _
                     ByVal Nickname As String, _
                     ByVal Reason As String)
        '**************************************************************
        'Author: Amraphen
        'Last Modify Date: 29/11/2010
        'Agrega un seguimiento.
        '**************************************************************
        '<EhHeader>
        On Error GoTo AddRecord_Err
        '</EhHeader>
100     NumRecords = NumRecords + 1
102     ReDim Preserve Records(1 To NumRecords)
    
104     With Records(NumRecords)
106         .Usuario = UCase$(Nickname)
108         .Fecha = Format(Now, "DD/MM/YYYY hh:mm:ss")
110         .Creador = UCase$(UserList(UserIndex).Name)
112         .Motivo = Reason
114         .NumObs = 0
        End With

        '<EhFooter>
        Exit Sub

AddRecord_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modUserRecords.AddRecord " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub AddObs(ByVal UserIndex As Integer, _
                  ByVal RecordIndex As Integer, _
                  ByVal Obs As String)
        '<EhHeader>
        On Error GoTo AddObs_Err
        '</EhHeader>

        '**************************************************************
        'Author: Amraphen
        'Last Modify Date: 29/11/2010
        'Agrega una observación.
        '**************************************************************
100     With Records(RecordIndex)
102         .NumObs = .NumObs + 1
104         ReDim Preserve .Obs(1 To .NumObs)
        
106         .Obs(.NumObs).Creador = UCase$(UserList(UserIndex).Name)
108         .Obs(.NumObs).Fecha = Now
110         .Obs(.NumObs).Detalles = Obs
        End With

        '<EhFooter>
        Exit Sub

AddObs_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modUserRecords.AddObs " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub RemoveRecord(ByVal RecordIndex As Integer)
        '<EhHeader>
        On Error GoTo RemoveRecord_Err
        '</EhHeader>

        '**************************************************************
        'Author: Amraphen
        'Last Modify Date: 29/11/2010
        'Elimina un seguimiento.
        '**************************************************************
        Dim i As Long
    
100     If RecordIndex = NumRecords Then
102         NumRecords = NumRecords - 1

104         If NumRecords > 0 Then
106             ReDim Preserve Records(1 To NumRecords)
            End If

        Else
108         NumRecords = NumRecords - 1

110         For i = RecordIndex To NumRecords
112             Records(i) = Records(i + 1)
114         Next i

116         ReDim Preserve Records(1 To NumRecords)
        End If

        '<EhFooter>
        Exit Sub

RemoveRecord_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modUserRecords.RemoveRecord " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub CreateRecordsFile()
        '<EhHeader>
        On Error GoTo CreateRecordsFile_Err
        '</EhHeader>

        '**************************************************************
        'Author: Amraphen
        'Last Modify Date: 29/11/2010
        'Crea el archivo de seguimientos.
        '**************************************************************
        Dim intFile As Integer

100     intFile = FreeFile
    
102     Open DatPath & "RECORDS.DAT" For Output As #intFile
104     Print #intFile, "[INIT]"
106     Print #intFile, "NumRecords=0"
108     Close #intFile
        '<EhFooter>
        Exit Sub

CreateRecordsFile_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modUserRecords.CreateRecordsFile " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
