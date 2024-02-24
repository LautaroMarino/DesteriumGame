VERSION 5.00
Begin VB.Form frmCargando 
   AutoRedraw      =   -1  'True
   BackColor       =   &H000080FF&
   BorderStyle     =   0  'None
   ClientHeight    =   2235
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmCargando.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   ScaleHeight     =   149
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   960
      Top             =   600
   End
   Begin VB.PictureBox LOGO 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   7200
      Left            =   960
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   0
      Top             =   9480
      Width           =   9600
   End
   Begin VB.PictureBox imgLoading 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   675
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   321
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1515
      Width           =   4815
   End
End
Attribute VB_Name = "frmCargando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Picture = LoadPicture(App.path & "\resource\interface\main\load.jpg")
    imgLoading.Picture = LoadPicture(App.path & "\resource\interface\main\loading.jpg")
    
    imgLoading.Width = 0
    
    Timer1.Enabled = True
    
End Sub

Private Sub Timer1_Timer()

    Static A As Long
    
    imgLoading.Width = imgLoading.Width + 1
    
    If imgLoading.Width >= 307 Then
        imgLoading.Width = 307
        Timer1.Enabled = False
    End If
    
End Sub
