VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGuildDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const p_SEGUNDOS = 7

Private p_CANTIDADDIALOGOS As Byte

Private Type t_GuildDlg

    Texto As String
    Segundos As Byte

End Type

Private p_Dialogos() As t_GuildDlg

Private p_Activo     As Boolean

Public Property Let CantidadDialogos(ByVal v As Byte)

    If v > 0 Then
        ReDim Preserve p_Dialogos(1 To v) As t_GuildDlg
        p_CANTIDADDIALOGOS = v
    End If

End Property

''
' Removes all dialogs from screen.

Public Sub RemoveDialogs()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/02/07
    'Removes all dialogs
    '***************************************************
    Dim I As Long
    
    For I = 1 To p_CANTIDADDIALOGOS
        p_Dialogos(I).Texto = vbNullString
    Next I

End Sub

''
' Retrieves the number of dialogs to be displayed on screen.
'
' @return   The number of dialogs to be displayed on screen.

Public Property Get CantidadDialogos() As Byte
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/01/07
    'Retrieves the number of dialogs to be displayed on screen
    '***************************************************
    CantidadDialogos = p_CANTIDADDIALOGOS
End Property

Public Property Let Activo(ByVal v As Boolean)
    p_Activo = v
End Property

Public Property Get Activo() As Boolean
    Activo = p_Activo
End Property

Private Sub Class_Initialize()
    p_CANTIDADDIALOGOS = 5
    ReDim p_Dialogos(1 To p_CANTIDADDIALOGOS) As t_GuildDlg
    p_Activo = True
End Sub

Public Sub Draw()

    Dim I As Byte

    Dim Y As Long
    Dim AddY As Long
    
    If Not p_Activo Then Exit Sub

    #If ModoBig = 0 Then
    Y = 350
    AddY = 10
    #Else
    Y = 350 * 2
    AddY = 20
    #End If
    
    '500 600 700 800 900
    For I = 1 To p_CANTIDADDIALOGOS

        If Trim$(p_Dialogos(I).Texto) <> vbNullString Then
            If Left$(p_Dialogos(I).Texto, 1) = Chr(3) Then
                Call Draw_Text(f_Chat, 14, 260, Y + (I * AddY), To_Depth(6), -1, ARGB(0, 255, 0, 200), FONT_ALIGNMENT_LEFT, mid(p_Dialogos(I).Texto, 2), True)
            Else
                Call Draw_Text(f_Chat, 14, AddY, Y + (I * AddY), To_Depth(6), -1, ARGB(255, 200, 0, 200), FONT_ALIGNMENT_LEFT, p_Dialogos(I).Texto, True)
            End If
        End If

    Next I

End Sub

Public Sub PassTimer()

    Dim I As Byte

    For I = 1 To p_CANTIDADDIALOGOS

        If p_Dialogos(I).Segundos > 0 Then
            p_Dialogos(I).Segundos = p_Dialogos(I).Segundos - 1
        Else
            p_Dialogos(I).Texto = vbNullString
        End If

    Next I
    
End Sub

''
' Splits dialogs into lines fitting properly the render area and inserts them

Public Sub PushBackText(ByVal s As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 07/04/2009
    'Splits dialogs into lines fitting properly the render area and inserts them
    '07/04/2009: Now cuts the string properly in spite of not fitting in the screen. This avoids an infite loop.
    '***************************************************
    Dim str As String

    Dim tmp As Integer
    
    str = s
    
    ' If it's too long to fit, split it
    Do While FrmMain.TextWidth(str) > 500
        tmp = InStrRev(str, " ")
        
        If tmp = 0 Then Exit Do
        
        str = Left$(str, tmp - 1)
    Loop
    
    'Show message and continue with the rest
    Call RealPushBackText(str)
    
    If LenB(str) <> LenB(s) Then
        Call PushBackText(Right$(s, Len(s) - Len(str) - 1))
    End If

End Sub

Private Sub RealPushBackText(ByVal s As String)

    Dim I     As Byte

    Dim Vacio As Boolean

    If p_CANTIDADDIALOGOS > 0 Then
        I = p_CANTIDADDIALOGOS
        Vacio = True

        While I > 0 And Vacio

            Vacio = p_Dialogos(I).Texto = vbNullString

            If Vacio Then I = I - 1

        Wend

        If I = p_CANTIDADDIALOGOS Then
            'hay q scrollear, estamos llenos
            I = 1

            While I < p_CANTIDADDIALOGOS

                p_Dialogos(I) = p_Dialogos(I + 1)
                I = I + 1

            Wend

            p_Dialogos(I).Texto = s
            p_Dialogos(I).Segundos = p_SEGUNDOS
        Else
            p_Dialogos(I + 1).Texto = s
            p_Dialogos(I + 1).Segundos = p_SEGUNDOS
        End If
    End If

End Sub
