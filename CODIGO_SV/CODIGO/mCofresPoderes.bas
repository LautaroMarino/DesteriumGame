Attribute VB_Name = "mCofresPoderes"
Option Explicit

Private Enum eCofres

    cBronce = 1
    cPlata = 2
    cOro = 3
    cPremium = 4
    cStreamer = 5
End Enum

Public Function UseCofrePoder(ByVal UserIndex As Integer, _
                              ByVal CofreIndex As Byte) As Boolean
        '<EhHeader>
        On Error GoTo UseCofrePoder_Err
        '</EhHeader>

        Dim ObjRequired As Obj

        Dim TempSTR     As String

        Dim Ft          As FontTypeNames
    
100     With UserList(UserIndex)
    
            ' REQUISITOS
102         Select Case CofreIndex
            
                Case eCofres.cStreamer
104                 If .flags.Streamer = 1 Then
106                     WriteConsoleMsg UserIndex, "Ya eres considerado un usuario Streamer", FontTypeNames.FONTTYPE_INFORED
                        Exit Function
                    End If
                
108                 TempSTR = "Servidor> El usuario " & .Name & " ha sido considerado como streamer de la comunidad."
110                 Ft = FontTypeNames.FONTTYPE_GUILD
                
112             Case eCofres.cOro
                
126                 TempSTR = "¡Te has convertido en una Leyenda!"
128                 Ft = FontTypeNames.FONTTYPE_USERGOLD
                
130             Case eCofres.cBronce
                
136                 TempSTR = "¡Te has convertido en un Aventurero!"
138                 Ft = FontTypeNames.FONTTYPE_USERBRONCE
                
140             Case eCofres.cPlata
150                 TempSTR = "¡Te vas convertido en un Héroe!"
152                 Ft = FontTypeNames.FONTTYPE_USERPLATA
            
154             Case eCofres.cPremium

156                 If .flags.Premium = 1 Then
158                     WriteConsoleMsg UserIndex, "Tu personaje ya posee el poder del cofre seleccionado.", FontTypeNames.FONTTYPE_INFO

                        Exit Function

                    End If
                
160                 TempSTR = "Te has convertido en un PERSONAJE PREMIUM"
162                 Ft = FontTypeNames.FONTTYPE_USERPREMIUM
            End Select

            ' APLICAMOS
164         Select Case CofreIndex
                Case eCofres.cStreamer
166                 .flags.Streamer = 1
                
168             Case eCofres.cOro
170                 .flags.Oro = 1

172             Case eCofres.cPremium
174                 .flags.Premium = 1

176             Case eCofres.cPlata
178                 .flags.Plata = 1

180             Case eCofres.cBronce
182                 .flags.Bronce = 1
            End Select
        
                Call WriteConsoleMsg(UserIndex, TempSTR, Ft)
184         'Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(TempSTR, Ft))
186         UseCofrePoder = True
        End With
    
        '<EhFooter>
        Exit Function

UseCofrePoder_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mCofresPoderes.UseCofrePoder " & _
               "at line " & Erl
        
        '</EhFooter>
End Function
