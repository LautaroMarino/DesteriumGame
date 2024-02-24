Attribute VB_Name = "mMessages"
' # Mensajes interactivos para ayudar al personaje online en su estadía. Le brindamos consejos útiles y mensajes relevantes para su entrenamiento y diversión
Option Explicit

Private Const SPAM_FILE_PATH As String = "Spam.dat"
Private Const SPAM_SECTION_INIT As String = "INIT"
Private Const SPAM_SECTION_MESSAGES As String = "MESSAGE"


Public MessageTime_Actual As Long
Public MessageTime As Long
Public MessageLast As Byte
Public MessageSpam() As String

Private MessageIndexesShown() As Boolean
Private TotalMessagesShown As Integer

Public Sub MessageSpam_Load()
    Dim Manager As clsIniManager
    Dim A As Long
    Dim FilePath As String

    On Error Resume Next

    FilePath = DatPath & SPAM_FILE_PATH

    Set Manager = New clsIniManager

    Manager.Initialize FilePath

    MessageTime = val(Manager.GetValue(SPAM_SECTION_INIT, "SPAM_TIME"))
    MessageLast = val(Manager.GetValue(SPAM_SECTION_INIT, "LAST"))

    ReDim MessageSpam(1 To MessageLast) As String
    ReDim MessageIndexesShown(1 To MessageLast) As Boolean
    
    
    For A = 1 To MessageLast
        MessageIndexesShown(A) = False
        MessageSpam(A) = Manager.GetValue(SPAM_SECTION_MESSAGES, CStr(A))
    Next A

    Set Manager = Nothing
    
    On Error GoTo 0
End Sub

Public Sub MessageSpam_SelectedRandom()
    If MessageLast > 0 Then
        ' Verifica si todos los mensajes se han mostrado al menos una vez
        If TotalMessagesShown = MessageLast Then
            ' Si todos los mensajes se han mostrado, reinicia la lista
            ResetMessageIndexes
        End If

        ' Busca un índice aleatorio que aún no se ha mostrado
        Dim randomIndex As Integer
        Do
            randomIndex = Int((MessageLast * Rnd) + 1)
        Loop While MessageIndexesShown(randomIndex)
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(MessageSpam(randomIndex), FontTypeNames.FONTTYPE_RMSG))
        
        ' Marca el mensaje como mostrado
        MessageIndexesShown(randomIndex) = True
        TotalMessagesShown = TotalMessagesShown + 1
    'Else
        'MsgBox "No hay mensajes disponibles.", vbExclamation, "Advertencia"
    End If
End Sub

Private Sub ResetMessageIndexes()
    ' Reinicia la lista de mensajes mostrados
    ReDim MessageIndexesShown(1 To MessageLast) As Boolean
    TotalMessagesShown = 0
End Sub


Public Sub MessageSpam_SpamUser()

End Sub
