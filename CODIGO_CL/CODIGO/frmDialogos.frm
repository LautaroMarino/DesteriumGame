VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmDialogos 
   Caption         =   "Dialogos"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5820
   Icon            =   "frmDialogos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5820
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Todos por Defecto"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Por Defecto"
      Height          =   255
      Left            =   3480
      TabIndex        =   17
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Aplicar"
      Height          =   255
      Left            =   4800
      TabIndex        =   16
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   4320
      TabIndex        =   15
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dialogo"
      Height          =   1935
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   1935
      Begin VB.OptionButton Option1 
         Caption         =   "Susurrar"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   10
         Top             =   1560
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Palabras mágicas"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Gritar"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Party"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Clan"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Normal"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin MSComctlLib.Slider b 
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   1200
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   25
      Max             =   255
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider g 
      Height          =   255
      Left            =   2760
      TabIndex        =   1
      Top             =   840
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   25
      Max             =   255
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider r 
      Height          =   255
      Left            =   2760
      TabIndex        =   0
      Top             =   480
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   25
      Max             =   255
      TickStyle       =   3
   End
   Begin VB.Label Label4 
      Caption         =   "Azul:"
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Verde:"
      Height          =   255
      Left            =   2280
      TabIndex        =   13
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Rojo:"
      Height          =   255
      Left            =   2280
      TabIndex        =   12
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Color:"
      Height          =   255
      Left            =   2400
      TabIndex        =   11
      Top             =   1785
      Width           =   975
   End
   Begin VB.Label lblPrueba 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   2400
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
End
Attribute VB_Name = "frmDialogos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'***************************************************
'Author: Juan Dalmasso (CHOTS)
'Last Modify Date: 12/06/2011
'Allow the character to modify the color of each dialog in the game
'***************************************************
Private indice As Byte

Private Sub b_Change()
    lblPrueba.BackColor = RGB(r.Value, g.Value, b.Value)
End Sub

Private Sub Command1_Click()
    'hace un "actualizar" antes de cerrarlo
    ColoresDialogos(indice).r = r.Value
    ColoresDialogos(indice).g = g.Value
    ColoresDialogos(indice).b = b.Value
    
    'los graba en el dialogos.dat
    Call GrabarColores
    
    Unload frmDialogos
End Sub

Private Sub Command2_Click()
    ColoresDialogos(indice).r = r.Value
    ColoresDialogos(indice).g = g.Value
    ColoresDialogos(indice).b = b.Value
End Sub

Private Sub Command3_Click()
    Call PorDefecto(indice)
End Sub

Private Sub Command4_Click()
    Call TodosPorDefecto
    
    r.Value = ColoresDialogos(indice).r
    g.Value = ColoresDialogos(indice).g
    b.Value = ColoresDialogos(indice).b
    
    lblPrueba.BackColor = RGB(r.Value, g.Value, b.Value)
End Sub

Private Sub Form_Load()
    indice = 1
    
    r.Value = ColoresDialogos(indice).r
    g.Value = ColoresDialogos(indice).g
    b.Value = ColoresDialogos(indice).b
    
End Sub

Private Sub g_Change()
    lblPrueba.BackColor = RGB(r.Value, g.Value, b.Value)
End Sub

Private Sub Option1_Click(Index As Integer)
    r.Value = ColoresDialogos(Index).r
    g.Value = ColoresDialogos(Index).g
    b.Value = ColoresDialogos(Index).b
    indice = Index
End Sub

Private Sub r_Change()
    lblPrueba.BackColor = RGB(r.Value, g.Value, b.Value)
End Sub

Private Sub PorDefecto(ByVal I As Byte)

    Dim archivoC As String

    archivoC = IniPath & "DialogosBACKUP.dat"
    
    ColoresDialogos(I).r = CByte(GetVar(archivoC, CStr(I), "R"))
    ColoresDialogos(I).g = CByte(GetVar(archivoC, CStr(I), "G"))
    ColoresDialogos(I).b = CByte(GetVar(archivoC, CStr(I), "B"))
    
    r.Value = ColoresDialogos(I).r
    g.Value = ColoresDialogos(I).g
    b.Value = ColoresDialogos(I).b
    
    lblPrueba.BackColor = RGB(r.Value, g.Value, b.Value)
End Sub

Private Sub TodosPorDefecto()

    Dim I As Byte
    
    For I = 1 To MAXCOLORESDIALOGOS
        Call PorDefecto(I)
    Next I

End Sub

Private Sub GrabarColores()

    Dim archivoC As String

    Dim I        As Byte

    archivoC = IniPath & "Dialogos.dat"
    
    For I = 1 To MAXCOLORESDIALOGOS
        Call WriteVar(archivoC, I, "R", ColoresDialogos(I).r)
        Call WriteVar(archivoC, I, "G", ColoresDialogos(I).g)
        Call WriteVar(archivoC, I, "B", ColoresDialogos(I).b)
    Next I
    
End Sub
