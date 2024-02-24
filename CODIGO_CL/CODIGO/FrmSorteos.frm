VERSION 5.00
Begin VB.Form FrmSorteos 
   Caption         =   "Sorterix"
   ClientHeight    =   8025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8475
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
   ScaleHeight     =   8025
   ScaleWidth      =   8475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Lista de Sorteos"
      Height          =   4935
      Left            =   240
      TabIndex        =   12
      Top             =   3000
      Width           =   8055
      Begin VB.ListBox lstLottery 
         Height          =   3960
         Left            =   240
         TabIndex        =   19
         Top             =   600
         Width           =   2535
      End
      Begin VB.Frame Frame3 
         Caption         =   "Informacion"
         Height          =   4095
         Left            =   3120
         TabIndex        =   13
         Top             =   480
         Width           =   4575
         Begin VB.Label lblCancel 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            Caption         =   "CANCELAR SORTEO"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   240
            Left            =   120
            TabIndex        =   18
            Top             =   3720
            Width           =   1905
         End
         Begin VB.Label lblSpam 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            Caption         =   "ENVIAR SPAM"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   240
            Left            =   3000
            TabIndex        =   17
            Top             =   3720
            Width           =   1425
         End
         Begin VB.Label lblDateFinish 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Sorteo:"
            Height          =   315
            Left            =   240
            TabIndex        =   16
            Top             =   1800
            Width           =   3090
         End
         Begin VB.Label lblDateInitial 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Inicio: "
            Height          =   315
            Left            =   240
            TabIndex        =   15
            Top             =   1440
            Width           =   3090
         End
         Begin VB.Label lblDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Desc: "
            Height          =   1035
            Left            =   240
            TabIndex        =   14
            Top             =   360
            Width           =   3090
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nuevo Sorteo"
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      Begin VB.TextBox txtDate 
         Height          =   285
         Left            =   1680
         TabIndex        =   21
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox txtObj 
         Height          =   285
         Index           =   1
         Left            =   3720
         TabIndex        =   9
         Top             =   1965
         Width           =   855
      End
      Begin VB.TextBox txtObj 
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   8
         Top             =   1965
         Width           =   855
      End
      Begin VB.TextBox txtChar 
         Height          =   285
         Left            =   1920
         TabIndex        =   7
         Top             =   1545
         Width           =   3495
      End
      Begin VB.TextBox txtDesc 
         Height          =   525
         Left            =   1920
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   945
         Width           =   5895
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha del Sorteo"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblINIT 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "INICIAR SORTEO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   240
         Left            =   6240
         TabIndex        =   11
         Top             =   2400
         Width           =   1545
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad"
         Height          =   195
         Index           =   4
         Left            =   2880
         TabIndex        =   10
         Top             =   1965
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Objeto a Sortear:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1965
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Personaje a Sortear:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   1605
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción del sorteo"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   1005
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre del Sorteo"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   645
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmSorteos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub lblINIT_Click()
        
        
        Dim TempLottery As tLottery
        
        
        If Not IsValidDateFormat(txtDate.Text) Then
            Call MsgBox("El formato de la fecha es incorrecto. UTILICE DD/MM/YYYY HH:MM y agrega 'hs'. Ejemplo: 01/01/2023 20:00hs")
            Exit Sub
        End If
        
        If Len(txtName.Text) <= 5 Then
            Call MsgBox("Elige un nombre más largo.")
            Exit Sub
        End If
        
        If Len(txtDesc.Text) <= 20 Then
            Call MsgBox("Elige una descripción más larga.")
            Exit Sub
        End If
        
                
        If Len(txtChar.Text) <= 0 And Len(txtObj(0).Text) <= 0 Then
            Call MsgBox("¡Elige algo para sortear!")
            Exit Sub
        End If
        
        If Len(txtObj(0).Text) > 0 And Len(txtObj(1).Text) <= 0 Then
            Call MsgBox("¡Elige la cantidad del objeto a sortear!")
            Exit Sub
        End If
        
        With TempLottery
            .Name = txtName.Text
            .Desc = txtDesc.Text
            .DateFinish = txtDate.Text
            .PrizeChar = txtChar.Text
            .PrizeObj = Val(txtObj(0).Text)
            .PrizeObjAmount = Val(txtObj(1).Text)
        End With
        
        Call WriteLotteryNew(TempLottery)
End Sub

Function IsValidDateFormat(inputString As String) As Boolean
    Dim strParts() As String
    Dim dateParts() As String
    Dim timeParts() As String
    Dim isValid As Boolean
    Dim suffix As String
    
    isValid = False

    ' Split the string into date and time parts
    strParts = Split(inputString, " ")
    If UBound(strParts) = 1 Then
        ' Split the date into day, month, year
        dateParts = Split(strParts(0), "/")
        If UBound(dateParts) = 2 Then
            If IsNumeric(dateParts(0)) And IsNumeric(dateParts(1)) And IsNumeric(dateParts(2)) Then
                If (CInt(dateParts(0)) > 0 And CInt(dateParts(0)) <= 31) And _
                   (CInt(dateParts(1)) > 0 And CInt(dateParts(1)) <= 12) And _
                   (CInt(dateParts(2)) >= 0) Then

                    ' Check the time
                    If Len(strParts(1)) > 2 Then
                        suffix = Right(strParts(1), 2)
                        If suffix = "hs" Then
                            timeParts = Split(Left(strParts(1), Len(strParts(1)) - 2), ":")
                            If UBound(timeParts) = 1 Then
                                If IsNumeric(timeParts(0)) And IsNumeric(timeParts(1)) Then
                                    If (CInt(timeParts(0)) >= 0 And CInt(timeParts(0)) <= 23) And _
                                       (CInt(timeParts(1)) >= 0 And CInt(timeParts(1)) <= 59) Then
                                        isValid = True
                                    End If
                                End If
                            End If
                        End If
                    End If

                End If
            End If
        End If
    End If

    IsValidDateFormat = isValid
End Function
