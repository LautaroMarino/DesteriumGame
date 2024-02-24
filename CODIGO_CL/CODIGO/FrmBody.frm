VERSION 5.00
Begin VB.Form FrmBody 
   Caption         =   "Arreglador Bodys 3000"
   ClientHeight    =   4050
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6510
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmBody.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4050
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAPLICARAL 
      Caption         =   "APLICAR AL BODY"
      Height          =   360
      Left            =   1080
      TabIndex        =   18
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox y4 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1200
      TabIndex        =   17
      Text            =   "0"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox x4 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1200
      TabIndex        =   16
      Text            =   "0"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox y3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1200
      TabIndex        =   15
      Text            =   "0"
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox x3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1200
      TabIndex        =   14
      Text            =   "0"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox y2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1200
      TabIndex        =   13
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox x2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1200
      TabIndex        =   12
      Text            =   "0"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox y1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1200
      TabIndex        =   11
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox x1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1200
      TabIndex        =   10
      Text            =   "0"
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txt 
      Height          =   285
      Left            =   1200
      TabIndex        =   9
      Text            =   "0"
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblCliclea 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliclea un Char y se actualizarán los valores"
      Height          =   195
      Left            =   2880
      TabIndex        =   23
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label lblOESTE 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OESTE"
      Height          =   195
      Left            =   2760
      TabIndex        =   22
      Top             =   2880
      Width           =   510
   End
   Begin VB.Label lblESTE 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ESTE"
      Height          =   195
      Left            =   2760
      TabIndex        =   21
      Top             =   1680
      Width           =   510
   End
   Begin VB.Label lblSUR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SUR"
      Height          =   195
      Index           =   1
      Left            =   2880
      TabIndex        =   20
      Top             =   2280
      Width           =   510
   End
   Begin VB.Label lblSUR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NORTE"
      Height          =   195
      Index           =   0
      Left            =   2760
      TabIndex        =   19
      Top             =   1080
      Width           =   510
   End
   Begin VB.Label lblY4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   600
      TabIndex        =   8
      Top             =   3000
      Width           =   210
   End
   Begin VB.Label lblX4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   600
      TabIndex        =   7
      Top             =   2760
      Width           =   210
   End
   Begin VB.Label lblY3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   600
      TabIndex        =   6
      Top             =   2400
      Width           =   210
   End
   Begin VB.Label lblX3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   600
      TabIndex        =   5
      Top             =   2160
      Width           =   210
   End
   Begin VB.Label lblY3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   600
      TabIndex        =   4
      Top             =   1800
      Width           =   210
   End
   Begin VB.Label lblX3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   600
      TabIndex        =   3
      Top             =   1560
      Width           =   210
   End
   Begin VB.Label lblX3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Width           =   210
   End
   Begin VB.Label lblX3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   210
   End
   Begin VB.Label lblBody 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Body"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   420
   End
End
Attribute VB_Name = "FrmBody"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAPLICARAL_Click()

    Dim Body As Integer
    Body = Val(txt.Text)
                
    If TempCharIndex > 0 Then
        
        CharList(TempCharIndex).Body.BodyOffSet(1).X = Val(FrmBody.x1.Text)
        CharList(TempCharIndex).Body.BodyOffSet(1).Y = Val(FrmBody.y1.Text)
                    
        CharList(TempCharIndex).Body.BodyOffSet(2).X = Val(FrmBody.x2.Text)
        CharList(TempCharIndex).Body.BodyOffSet(2).Y = Val(FrmBody.y2.Text)
        
        CharList(TempCharIndex).Body.BodyOffSet(3).X = Val(FrmBody.x3.Text)
        CharList(TempCharIndex).Body.BodyOffSet(3).Y = Val(FrmBody.y3.Text)
        
        CharList(TempCharIndex).Body.BodyOffSet(4).X = Val(FrmBody.x4.Text)
        CharList(TempCharIndex).Body.BodyOffSet(4).Y = Val(FrmBody.y4.Text)
    Else
        Call MsgBox("Selecciona un char")
    End If
End Sub
