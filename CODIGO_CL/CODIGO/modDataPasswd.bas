Attribute VB_Name = "mDataPasswd"
Option Explicit

Public Const NUMPASSWD      As Byte = 50
Public Const NUMCHARS      As Byte = 30

Public Const ARCHIVE_PASSWD As String = "data.ini"

Public Const INVALID_SLOT   As Integer = -1

Public Type tPasswd

    Account As String
    Passwd As String

End Type

Public ListPasswd(1 To NUMPASSWD) As tPasswd

Public LastAccount As String

Public Type tAccount
    Name As String
    Passwd As String
    Pjs(NUMCHARS) As String
End Type

Public DataPasswd As tAccount
Function GetParentFolderPath(ByVal path As String) As String
    ' Verificar si la ruta contiene una barra invertida al final
    If Right(path, 1) = "\" Then
        ' Quitar la barra invertida final
        path = Left(path, Len(path) - 1)
    End If
    
    ' Obtener la carpeta padre
    GetParentFolderPath = Left(path, InStrRev(path, "\"))
End Function

Public Function IDecryptText(ByVal strText As String, ByVal strPwd As String) As String
    Dim i As Integer
    Dim C As Integer
    Dim strBuff As String

    strPwd = UCase$(strPwd)

    ' Decrypt string
    If Len(strPwd) > 0 Then
        For i = 1 To Len(strText)
            C = Asc(mid$(strText, i, 1))
            C = C - Asc(mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
            strBuff = strBuff & Chr$(C And &HFF)
        Next i
    Else
        strBuff = strText
    End If
    
    IDecryptText = strBuff
End Function

Public Sub LoadListPasswd()

    On Error GoTo ErrHandler
    
    Dim A        As Long

    Dim filePath As String

    Dim Manager  As clsIniManager

    Dim Temp     As String
    
    Set Manager = New clsIniManager
    
    Dim Parent As String
    Parent = GetParentFolderPath(App.path)
    
    filePath = Parent & ARCHIVE_PASSWD
    
    If FileExist(filePath, vbArchive) Then
        Manager.Initialize filePath
    End If
    
    For A = 1 To NUMPASSWD
        With ListPasswd(A)
            .Account = Manager.GetValue("PASSWD", "EM" & A)
            .Passwd = IDecryptText(Manager.GetValue("PASSWD", "PM" & A), "1")
        End With
    Next A

    Set Manager = Nothing
    Exit Sub
    
ErrHandler:
End Sub


Public Sub SaveNewAccount(ByVal Account As String, _
                          ByVal Passwd As String, _
                          ByVal KillAccount As Boolean)

    Dim Slot As Integer
    
    Slot = SearchSlot(Account)
    
    LastAccount = Account
    
    If Slot = INVALID_SLOT Then
        Slot = FreeSlot
        
        If Slot <> INVALID_SLOT Then

            With ListPasswd(Slot)
                .Account = Account
                .Passwd = Passwd
                
            End With

        Else
            MsgBox "Tus datos no han sido guardado debido a que superaste el máximo de capacidad."

            Exit Sub

        End If

    Else

        With ListPasswd(Slot)

            If KillAccount Then
                .Passwd = vbNullString
                .Account = vbNullString
            Else
                .Passwd = Passwd
                
            End If

        End With

    End If

End Sub

' Cuando el nombre fue encontrado al hacer clic sobre el txtbox de la contraseña ponemos la passwd correspondiente.
Public Function SearchPasswd(ByVal Account As String) As String

    Dim A As Long
  
    For A = 1 To NUMPASSWD

        With ListPasswd(A)

            If LCase$(.Account) = LCase$(Account) Then
                SearchPasswd = .Passwd

                Exit Function

            End If

        End With

    Next A
    
End Function

' Funciones internas
' Slot Libre
Private Function FreeSlot() As Integer

    Dim A As Long
    
    For A = 1 To NUMPASSWD

        If ListPasswd(A).Account = vbNullString Then
            FreeSlot = A
            Exit Function
        End If
    Next A
    
    FreeSlot = INVALID_SLOT
End Function

' Nombre existente
Private Function SearchSlot(ByVal Account As String) As Integer

    For SearchSlot = 1 To NUMPASSWD

        If ListPasswd(SearchSlot).Account = Account Then Exit Function
    Next SearchSlot
    
    SearchSlot = INVALID_SLOT
End Function



