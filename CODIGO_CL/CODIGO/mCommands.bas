Attribute VB_Name = "mCommands"
Option Explicit

Public Const LAST_COMMAND As Byte = 12

Public Commands(1 To LAST_COMMAND) As String



Public Sub Commands_Load()
    
    ' @ Not add '!' or '/'
    Commands(1) = "SALIR"
    Commands(2) = "COMERCIAR"
    Commands(3) = "AYUDA"
    Commands(4) = "DENUNCIAR"
    Commands(5) = "INFOEVENTO"
    Commands(6) = "ENTRAR"
    Commands(7) = "ENLISTAR"
    Commands(8) = "ABANDONAR"
    Commands(9) = "SKINS"
    Commands(10) = "SHOP"
    Commands(11) = "STREAM"
    Commands(12) = "STREAMLINK"
    
End Sub

Public Sub Commands_Search(ByRef Text As String)
    Exit Sub
    
    Dim Temp As String
    
    Temp = UCase$(Text)
    
    If Not (Left$(Text, 1) = "!" Or Left$(Text, 1) = "/") Then Exit Sub
    
  '  Temp = Replace$(Text, "!", "/")
    
    Dim lIndex As Long
    
    
    ' Recorro los arrays
    For lIndex = 1 To UBound(Commands)

        ' Si coincide con los patrones
        If InStr(2, Temp, Commands(lIndex)) Then
            ' Se fija de que no esté escrito ya el comando completo. Si el usuario aplica un espacio termino de escribir el comando
            If Not InStr(1, " ", Temp) Then
                Temp = "!" & Commands(lIndex)
            End If
        End If

    Next lIndex
    
  '  Temp = Replace$(Text, "/", "!")
    Text = Temp
    
    FrmMain.SendTxt.SelStart = Len(Text)
End Sub
