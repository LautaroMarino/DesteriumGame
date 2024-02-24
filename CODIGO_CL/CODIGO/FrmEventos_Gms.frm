VERSION 5.00
Begin VB.Form FrmEventos_Gms 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9255
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmEventos_Gms.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "FrmEventos_Gms.frx":000C
   ScaleHeight     =   2985
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "CONFIRMAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7455
      TabIndex        =   1
      Top             =   2415
      Width           =   1725
   End
   Begin VB.ListBox lstEvents 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   2130
      Left            =   0
      TabIndex        =   0
      Top             =   210
      Width           =   9255
   End
End
Attribute VB_Name = "FrmEventos_Gms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Frame5_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub lblInfoEvent_Click()

End Sub

Private Sub Form_Load()
    
    ' Lista de Eventos
    Call Events_SetList
End Sub

Private Sub Events_SetList()
    Dim A As Long
    
    lstEvents.Clear
    
    lstEvents.AddItem "1vs1» 4 Cupos. 500 Rojas. Premio: 250.000 ORO"
    lstEvents.AddItem "1vs1» 8 Cupos. 1000 Rojas. Premio: 500.000 ORO"
    lstEvents.AddItem "1vs1» 8 Cupos. 1000 Rojas. Premio: 500.000 ORO Y 20 Fragmentos Premium."
    lstEvents.AddItem "1vs1» 16 Cupos. 1500 Rojas. Premio: 500.000 ORO Y 20 Fragmentos Premium."
End Sub

