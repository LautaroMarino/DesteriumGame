VERSION 5.00
Begin VB.Form FrmConectando 
   BorderStyle     =   0  'None
   Caption         =   "Conectando al Servidor..."
   ClientHeight    =   1425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6915
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmConectando.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "FrmConectando.frx":000C
   ScaleHeight     =   1425
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tReconnect 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   360
      Top             =   480
   End
End
Attribute VB_Name = "FrmConectando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)


    If KeyCode = vbKeyEscape Then
        prgRun = False
        Exit Sub
    End If
    
End Sub

Public Sub Reconnect_Socket()
        Account.Email = LastDataAccount
        Account.Passwd = LastDataPasswd
        Prepare_And_Connect E_MODO.e_LoginAccount
End Sub
Private Sub tReconnect_Timer()

    If Not IsConnected Then
        Reconnect_Socket
    Else
        tReconnect.Enabled = False
    End If
    
End Sub
