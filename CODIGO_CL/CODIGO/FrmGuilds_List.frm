VERSION 5.00
Begin VB.Form FrmGuilds_List 
   BorderStyle     =   0  'None
   Caption         =   "Lista de Clanes"
   ClientHeight    =   7590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtGuild 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   300
      Left            =   1320
      TabIndex        =   3
      Top             =   1180
      Width           =   1935
   End
   Begin VB.ListBox lstGuild 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   3810
      Left            =   600
      TabIndex        =   0
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Image imgLeader 
      Height          =   255
      Left            =   2040
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Image imgUnload 
      Height          =   315
      Left            =   4920
      Picture         =   "FrmGuilds_List.frx":0000
      Top             =   0
      Width           =   330
   End
   Begin VB.Label lblExp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   3315
      TabIndex        =   2
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label lblElv 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   3800
      TabIndex        =   1
      Top             =   2080
      Width           =   735
   End
   Begin VB.Image imgLvl 
      Height          =   255
      Left            =   3600
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Image imgFound 
      Height          =   255
      Left            =   600
      Top             =   6120
      Width           =   1215
   End
End
Attribute VB_Name = "FrmGuilds_List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Picture = LoadPicture(DirInterface & "menucompacto\guilds_list.jpg")
    Call ListarClanes
End Sub

Private Sub ListarClanes()
    Dim A As Long
    
    Dim Exist As Boolean
    
    lstGuild.Clear
    
    For A = 1 To MAX_GUILDS
        If GuildsInfo(A).Lvl > 0 Then
            lstGuild.AddItem GuildsInfo(A).Name
            Exist = True
        End If
    Next A
    
    If Exist Then
        lstGuild.ListIndex = 0
    End If
End Sub

Private Sub imgFound_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    #If ModoBig = 1 Then
        dockForm FrmGuilds_Found.hWnd, FrmMain.PicMenu, True
    #Else
        Call FrmGuilds_Found.Show(, FrmMain)
    #End If
    
    Unload Me
End Sub

Private Sub imgLeader_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Call WriteGuilds_Required(1000)
    Unload Me
End Sub

Private Sub imgLvl_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    #If ModoBig = 1 Then
        dockForm FrmGuilds_Levels.hWnd, FrmMain.PicMenu, True
    #Else
        Call FrmGuilds_Levels.Show(, FrmMain)
    #End If
End Sub

Private Sub imgUnload_Click()
    Call Audio.PlayInterface(SND_CLICK)

    Unload Me
End Sub


Private Sub lstGuild_Click()
    If lstGuild.ListIndex = -1 Then Exit Sub
    
    Dim Slot As Integer
    Slot = SearchGuild(UCase$(lstGuild.List(lstGuild.ListIndex)))
    
    If Slot > 0 Then
        lblElv.Caption = GuildsInfo(Slot).Lvl
        lblExp.Caption = PonerPuntos(GuildsInfo(Slot).Exp)
    End If
    
End Sub

' # Busca el clan en la lista
Private Function SearchGuild(ByVal Name As String) As Integer
    Dim A As Long
    
    For A = 1 To MAX_GUILDS
        If StrComp(UCase$(GuildsInfo(A).Name), Name) = 0 Then
            SearchGuild = A
            Exit Function
        End If
    Next A
End Function
Private Sub txtGuild_Change()

    Dim A As Long
    
    If Len(txtGuild.Text) <= 0 Then
        ListarClanes
    Else
        Call FiltrarListaClanes(txtGuild.Text)
    End If
End Sub

' # Filtra la lista de clanes
Public Sub FiltrarListaClanes(ByRef sCompare As String)

    Dim lIndex As Long, b As Long
    Dim GuildNull As tGuild
    
    lstGuild.Clear
    
    If UBound(GuildsInfo) <> 0 Then

        ' Recorro los arrays
        For lIndex = 1 To UBound(GuildsInfo)
            ' Si coincide con los patrones
            If InStr(1, UCase$(GuildsInfo(lIndex).Name), UCase$(sCompare)) Then
                lstGuild.AddItem GuildsInfo(lIndex).Name
            End If
        Next lIndex

    End If

End Sub
