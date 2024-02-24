VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmCriatura_Events 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   ClientHeight    =   7605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5220
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmCriature_Events.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   507
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   348
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picEvents 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      DrawStyle       =   3  'Dash-Dot
      BeginProperty Font 
         Name            =   "Booter - Five Zero"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1725
      Index           =   0
      Left            =   345
      MousePointer    =   99  'Custom
      Picture         =   "frmCriature_Events.frx":000C
      ScaleHeight     =   115
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   305
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   975
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Timer tDraw 
      Interval        =   50
      Left            =   840
      Top             =   240
   End
   Begin VB.PictureBox PicChar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1845
      Left            =   1710
      ScaleHeight     =   123
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   123
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3540
      Width           =   1845
   End
   Begin VB.Timer tSecond 
      Interval        =   1050
      Left            =   360
      Top             =   240
   End
   Begin RichTextLib.RichTextBox Console 
      Height          =   1665
      Left            =   360
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Consola de Comercio"
      Top             =   5640
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   2937
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmCriature_Events.frx":ABEC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmCriatura_Events"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Private Heading As Byte
Private clsFormulario As clsFormMovementManager

Public LastIndex1     As Integer

Public LastIndex2     As Integer

Public LasActionBuy   As Boolean

Private ClickNpcInv   As Boolean

Private lIndex        As Byte
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_EXSTYLE = -20
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT As Long = &H20&

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyEscape Then
            Unload Me
            Exit Sub
        End If
End Sub

Private Sub Form_Load()

    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
          
    Me.Picture = LoadPicture(App.path & "\resource\interface\events\events_info.jpg")

    Call SetWindowLong(Console.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
    
    'g_Captions(eCaption.cPicEvent) = wGL_Graphic.Create_Device_From_Display(picInv.hWnd, picInv.ScaleWidth, picInv.ScaleHeight)
    g_Captions(eCaption.cPicEventChar) = wGL_Graphic.Create_Device_From_Display(PicChar.hWnd, PicChar.ScaleWidth, PicChar.ScaleHeight)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set InvEvent = Nothing
    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.cPicEvent))
    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.cPicEventChar))
    
    Call Audio.DeleteSource(SOURCE_INTERFACE, True)
End Sub

Private Sub PicInv_Click()
    Call InvEvent.DrawInventory
End Sub

Private Sub tDraw_Timer()
    Render_CharPrizeObj
End Sub

Private Sub tSecond_Timer()
    If Events_TimeInit > 0 Then
        Events_TimeInit = Events_TimeInit - 1
        
        If Events_TimeInit = 0 Then
           ' lblTime.Caption = "Inscripciones Abiertas"
          '  lblTime.ForeColor = vbGreen
        Else
           ' lblTime.Caption = SecondsToHMS(Events_TimeInit)
        '    lblTime.ForeColor = vbWhite
        End If
        
    End If

End Sub
