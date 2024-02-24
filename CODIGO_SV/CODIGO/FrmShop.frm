VERSION 5.00
Begin VB.Form FrmShop 
   Caption         =   "Lista de Transacciones"
   ClientHeight    =   3975
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8715
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton ButtonRechace 
      Caption         =   "Rechazar"
      Height          =   360
      Left            =   3120
      TabIndex        =   3
      Top             =   3480
      Width           =   990
   End
   Begin VB.CommandButton ButtonAccept 
      Caption         =   "Aceptar"
      Height          =   360
      Left            =   240
      TabIndex        =   2
      Top             =   3480
      Width           =   990
   End
   Begin VB.ListBox lstShop 
      Appearance      =   0  'Flat
      Height          =   2760
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   3855
   End
   Begin VB.Label lblRef 
      BackStyle       =   0  'Transparent
      Caption         =   "Referencias"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   4320
      TabIndex        =   5
      Top             =   600
      Width           =   4260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Referencias"
      Height          =   195
      Left            =   4320
      TabIndex        =   4
      Top             =   240
      Width           =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lista de Transacciones"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1620
   End
End
Attribute VB_Name = "FrmShop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ButtonAccept_Click()
        '<EhHeader>
        On Error GoTo ButtonAccept_Click_Err
        '</EhHeader>

100     If lstShop.ListIndex = -1 Then Exit Sub
    
102     'If MsgBox("¿Estás seguro que deseas aceptar la transacción de " & lstShop.List(lstShop.ListIndex) & "?", vbYesNo) = vbYes Then
104         Call mShop.Transaccion_Accept((val(ReadField(1, lstShop.List(lstShop.ListIndex), Asc("|")))))
        'End If
        '<EhFooter>
        Exit Sub

ButtonAccept_Click_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.FrmShop.ButtonAccept_Click " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub ButtonRechace_Click()
        '<EhHeader>
        On Error GoTo ButtonRechace_Click_Err
        '</EhHeader>
100      If lstShop.ListIndex = -1 Then Exit Sub
     
102    ' If MsgBox("¿Estás seguro que deseas rechazar la transacción de " & lstShop.List(lstShop.ListIndex) & "?", vbYesNo) = vbYes Then
104         Call mShop.Transaccion_Clear(val(ReadField(1, lstShop.List(lstShop.ListIndex), Asc("|"))))

       ' End If
        '<EhFooter>
        Exit Sub

ButtonRechace_Click_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.FrmShop.ButtonRechace_Click " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub lstShop_Click()
        '<EhHeader>
        On Error GoTo lstShop_Click_Err
        '</EhHeader>
100     If lstShop.ListIndex = -1 Then Exit Sub
    
   
    
        Dim Slot As Long
    
102     Slot = val(ReadField(1, lstShop.List(lstShop.ListIndex), Asc("|")))
    
104     lblRef.Caption = ShopWaiting(Slot).Bank

    
        '<EhFooter>
        Exit Sub

lstShop_Click_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.FrmShop.lstShop_Click " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
