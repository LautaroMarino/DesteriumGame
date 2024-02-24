Attribute VB_Name = "mFotoDenuncias"
Option Explicit
 
'declare Constants and variables
 
'One second of interval.
'Changes to public of the use in call to capture picture.
Public Const FotoD_MAX_INTERVAL                As Long = 60000

'Here save last interval of photo report.
Private FotoD_LastIN                           As Long

'Number of last insult to the Array.
Private Const FOTOD_INSULTMAX                  As Byte = 36

'Container array of insult list.
Private FotoD_InsultList(1 To FOTOD_INSULTMAX) As String
 
Sub FotoD_Initialize()
       
    FotoD_InsultList(1) = "PT"
    FotoD_InsultList(2) = "MANCO"
    FotoD_InsultList(3) = "ASCO"
    FotoD_InsultList(4) = "ASKO"
    FotoD_InsultList(5) = "NW"
    FotoD_InsultList(6) = "FRACA"
    FotoD_InsultList(7) = "FRAKA"
    FotoD_InsultList(8) = "PETE"
    FotoD_InsultList(9) = "DAS PENA"
    FotoD_InsultList(10) = "KB"
    FotoD_InsultList(11) = "KABE"
    FotoD_InsultList(12) = "CABE"
    FotoD_InsultList(13) = "KBIO"
    FotoD_InsultList(14) = "CABIO"
    FotoD_InsultList(15) = "TAS EN LA RUINA"
    FotoD_InsultList(16) = "PUTO"
    FotoD_InsultList(17) = "PUTA"
    FotoD_InsultList(18) = "PAJERO"
    FotoD_InsultList(19) = "PAJERA"
    FotoD_InsultList(20) = "CONCHA"
    FotoD_InsultList(21) = "TU MADRE"
    FotoD_InsultList(22) = "TU MAMA"
    FotoD_InsultList(23) = "HIJO"
       
    FotoD_InsultList(24) = "LA PUTA QUE TE RE MIL PARIO PEDAZO DE FRACA HIJO DE PUTA DAS ASKO AJAJJAJAJAJA"
       
    FotoD_InsultList(25) = "SORETE"
    FotoD_InsultList(26) = "MIERDA"
    FotoD_InsultList(27) = "PELOTUDO"
    FotoD_InsultList(28) = "MOGOLICO"
    FotoD_InsultList(29) = "RETRASADO"
    FotoD_InsultList(30) = "ENFERMO"
    FotoD_InsultList(31) = "DAWN"
    FotoD_InsultList(32) = "SIMIO"
    FotoD_InsultList(33) = "NO TENES VIDA"
    FotoD_InsultList(34) = "CAGADA"
    FotoD_InsultList(35) = "VIRGEN"
    FotoD_InsultList(36) = "PENE"
       
    FotoD_LastIN = 60001
       
End Sub
 
Sub FotoD_Capturar(refString As String)

    Dim LoopX      As Long

    Dim sendString As String
       
    'Whenever we initialize the variable is null.
    sendString = vbNullString
       
    For LoopX = 1 To LastChar
             
        With CharList(LoopX)
             
            'It's char in pc area?
            If FotoD_CharInPCArea(LoopX) Then

                'Analize LastDialog
                If FotoD_DialogIsInsult(LoopX) Then
                    'Save charDialogs and NickName here.
                    sendString = sendString & "," & .Nombre & " : " & .LastDialog
                End If
            End If

        End With
             
    Next LoopX
       
    refString = sendString
       
    If refString <> vbNullString Then
        FotoD_LastIN = FrameTime
    End If
       
End Sub
 
Sub FotoD_SaveLastDialog(ByVal CharIndex As Integer, ByVal DialoG As String)
       
    With CharList(CharIndex)

        If .Nombre = vbNullString Then Exit Sub
        .LastDialog = DialoG
       
    End With
       
End Sub
 
Sub FotoD_RemoveLastDialog(ByVal CharIndex As Integer)

    If CharList(CharIndex).Nombre = vbNullString Then Exit Sub
    CharList(CharIndex).LastDialog = vbNullString
       
End Sub
 
Function FotoD_DialogIsInsult(ByVal CharIndex As Integer) As Boolean

    Dim LoopX As Long
       
    For LoopX = 1 To UBound(FotoD_InsultList())
             
        'Analize charDialogs
        If CharList(CharIndex).LastDialog <> vbNullString Then
            If InStr(1, UCase$(CharList(CharIndex).LastDialog), FotoD_InsultList(LoopX)) Then
                'Insult are found? returns true and exit function!
                FotoD_DialogIsInsult = True
    
                Exit Function
    
            End If
        End If
        
    Next LoopX
       
    FotoD_DialogIsInsult = False
       
End Function
 
Function FotoD_CanSend() As Boolean
       
    If FotoD_LastIN = 60001 Then FotoD_CanSend = True: Exit Function
       
    FotoD_CanSend = (FrameTime - FotoD_LastIN > FotoD_MAX_INTERVAL)
       
End Function
 
Function FotoD_CharInPCArea(ByVal CharIndex As Integer) As Boolean
       
    With CharList(CharIndex)
         
        FotoD_CharInPCArea = (.Pos.X > (UserPos.X - MinXBorder) And .Pos.X < (UserPos.X + MinXBorder) And .Pos.Y > (UserPos.Y - MinYBorder) And .Pos.Y < (UserPos.Y + MinYBorder))
             
    End With
         
End Function

