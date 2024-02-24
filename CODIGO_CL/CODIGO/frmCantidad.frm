VERSION 5.00
Begin VB.Form FrmCantidad 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1950
   ClientLeft      =   1635
   ClientTop       =   4410
   ClientWidth     =   3990
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawMode        =   7  'Invert
   Icon            =   "frmCantidad.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1430
      TabIndex        =   0
      Text            =   "1"
      Top             =   600
      Width           =   1150
   End
   Begin VB.Image imgMenos 
      Height          =   330
      Left            =   840
      Top             =   600
      Width           =   390
   End
   Begin VB.Image imgMas 
      Height          =   330
      Left            =   2640
      Top             =   600
      Width           =   390
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   2160
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   360
      Top             =   1200
      Width           =   1575
   End
End
Attribute VB_Name = "FrmCantidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Dragged As Boolean

Private X      As Single

Private Y      As Single

Private tX     As Byte

Private tY     As Byte

Private MouseX As Long

Private MouseY As Long

Private Sub Form_Unload(Cancel As Integer)
    MirandoCantidad = False
End Sub

Private Sub Form_Load()
    MirandoCantidad = True
    
    
    Me.Picture = LoadPicture(App.path & "\resource\interface\bank\soltarobjeto.jpg")
End Sub

Public Sub SetDropGround()
    Dragged = False
End Sub

Public Sub SetDropDragged(ByVal NewX As Single, ByVal NewY As Single)
    Dragged = True
    X = NewX
    Y = NewY
End Sub

Private Sub Image1_Click()
    
    If LenB(FrmCantidad.Text1) > 0 Then
        Call Audio.PlayInterface(SND_CLICK)
        
        If Not IsNumeric(FrmCantidad.Text1) Then
            Unload Me
            Exit Sub  'Should never happen
        End If
        
        If Dragged Then
              
            If Text1 > Inventario.Amount(Inventario.SelectedItem) Then
                ShowConsoleMsg "No tienes esa cantidad!", 65, 190, 156, False, False
                Unload Me

                Exit Sub

            End If
              
            ConvertCPtoTP X, Y, tX, tY
            WriteDragToPos tX, tY, Inventario.SelectedItem, Text1
            FrmCantidad.Text1.Text = vbNullString
              
        Else
            Call WriteDrop(Inventario.SelectedItem, FrmCantidad.Text1.Text)
            FrmCantidad.Text1.Text = vbNullString
        End If
    End If

    Unload Me
End Sub

Private Sub Image2_Click()

    Call Audio.PlayInterface(SND_CLICK)
    
    If Inventario.SelectedItem = 0 Then Exit Sub
    
    If Not IsNumeric(FrmCantidad.Text1) Then Exit Sub  'Should never happen
    
    If Dragged Then 'drag and drop
        ConvertCPtoTP X, Y, tX, tY
        WriteDragToPos tX, tY, Inventario.SelectedItem, Inventario.Amount(Inventario.SelectedItem)
        FrmCantidad.Text1.Text = vbNullString
        Unload Me
          
    Else 'tirar al piso :D

        If Inventario.SelectedItem <> FLAGORO Then
            Call WriteDrop(Inventario.SelectedItem, Inventario.Amount(Inventario.SelectedItem))
            Unload Me
        Else

            If UserGLD > 10000 Then
                Call WriteDrop(Inventario.SelectedItem, 10000)
                Unload Me
            Else
                Call WriteDrop(Inventario.SelectedItem, UserGLD)
                Unload Me
            End If
        End If
    End If

    FrmCantidad.Text1.Text = ""
          
    Unload Me
End Sub

Private Sub Text1_Change()

    On Error GoTo ErrHandler

    If Val(FrmCantidad.Text1) < 0 Then
        FrmCantidad.Text1 = "1"
    End If
          
          
    If Val(FrmCantidad.Text1) > 10000 Then
        FrmCantidad.Text1 = "10000"
    End If
          
    Exit Sub
          
ErrHandler:
    'If we got here the user may have pasted (Shift + Insert) a REALLY large number, causing an overflow, so we set amount back to 1
    FrmCantidad.Text1 = "1"
End Sub


Private Function CheckAdding(ByVal Value As Long) As Long
    If Value <= 100 Then
         CheckAdding = Value + 1
    ElseIf Value <= 1000 Then
        CheckAdding = Value + 100
    ElseIf Value <= 10000 Then
        CheckAdding = Value + 1000
    Else
        CheckAdding = MAX_INVENTORY_OBJS
    End If
End Function
Private Sub imgMas_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Text1.Text = CheckAdding(Text1.Text)
    
    If Val(Text1.Text) > MAX_INVENTORY_OBJS Then
        Text1.Text = MAX_INVENTORY_OBJS
    End If
    
End Sub
Private Function CheckAdding_Menos(ByVal Value As Long) As Long
    If Value <= 100 Then
         CheckAdding_Menos = Value - 1
    ElseIf Value <= 1000 Then
        CheckAdding_Menos = Value - 100
    ElseIf Value <= 10000 Then
        CheckAdding_Menos = Value - 1000
    Else
        CheckAdding_Menos = 1
    End If
End Function
Private Sub imgMenos_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Text1.Text = CheckAdding_Menos(Text1.Text)
    
    If Val(Text1.Text) < 1 Then
        Text1.Text = 1
    End If
    
End Sub

