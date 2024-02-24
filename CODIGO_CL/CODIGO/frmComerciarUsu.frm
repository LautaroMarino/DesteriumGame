VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmComerciarUsu 
   BorderStyle     =   0  'None
   ClientHeight    =   8850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9960
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmComerciarUsu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmComerciarUsu.frx":000C
   ScaleHeight     =   590
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   664
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picInvEldhirOfertaOtro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   5610
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5550
      Width           =   960
   End
   Begin VB.PictureBox picInvEldhirOfertaProp 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   5580
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1440
      Width           =   960
   End
   Begin VB.PictureBox picInvEldhirProp 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   3450
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1440
      Width           =   960
   End
   Begin VB.PictureBox picInvOroProp 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   3450
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   7
      Top             =   930
      Width           =   960
   End
   Begin VB.TextBox txtAgregar 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   4500
      TabIndex        =   6
      Top             =   2295
      Width           =   1035
   End
   Begin VB.PictureBox picInvOroOfertaOtro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   5610
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   5
      Top             =   5040
      Width           =   960
   End
   Begin VB.PictureBox picInvOfertaOtro 
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
      Height          =   3150
      Left            =   7080
      ScaleHeight     =   210
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   175
      TabIndex        =   4
      Top             =   5040
      Width           =   2625
   End
   Begin VB.PictureBox picInvOfertaProp 
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
      Height          =   3150
      Left            =   7080
      ScaleHeight     =   210
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   175
      TabIndex        =   3
      Top             =   930
      Width           =   2625
   End
   Begin VB.TextBox SendTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   495
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   7965
      Width           =   6060
   End
   Begin VB.PictureBox picInvComercio 
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
      Height          =   3150
      Left            =   480
      ScaleHeight     =   210
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   175
      TabIndex        =   1
      Top             =   945
      Width           =   2625
   End
   Begin VB.PictureBox picInvOroOfertaProp 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   5580
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   0
      Top             =   930
      Width           =   960
   End
   Begin RichTextLib.RichTextBox CommerceConsole 
      Height          =   1620
      Left            =   495
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   6150
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   2858
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmComerciarUsu.frx":23C6C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image imgCancelar 
      Height          =   360
      Left            =   360
      Tag             =   "1"
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Image imgRechazar 
      Height          =   360
      Left            =   8220
      Tag             =   "2"
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Image imgConfirmar 
      Height          =   360
      Left            =   7440
      Tag             =   "2"
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Image imgAceptar 
      Height          =   360
      Left            =   6750
      Tag             =   "2"
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Image imgAgregar 
      Height          =   255
      Left            =   4800
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image imgQuitar 
      Height          =   255
      Left            =   4800
      Top             =   2760
      Width           =   375
   End
End
Attribute VB_Name = "frmComerciarUsu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
' frmComerciarUsu.frm
'
'**************************************************************

'**************************************************************************
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
'**************************************************************************

Option Explicit

Private clsFormulario         As clsFormMovementManager

Private cBotonAceptar         As clsGraphicalButton

Private cBotonCancelar        As clsGraphicalButton

Private cBotonRechazar        As clsGraphicalButton

Private cBotonConfirmar       As clsGraphicalButton

Public LastButtonPressed      As clsGraphicalButton

Private Const GOLD_OFFER_SLOT As Byte = INV_OFFER_SLOTS + 1

Private Const ELDHIR_OFFER_SLOT As Byte = INV_OFFER_SLOTS + 2

Private sCommerceChat         As String

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub imgAceptar_Click()

    If Not cBotonAceptar.IsEnabled Then Exit Sub  ' Deshabilitado
          
    Call WriteUserCommerceOk
    HabilitarAceptarRechazar False
          
End Sub

Private Sub imgAgregar_Click()
         
    ' No tiene seleccionado ningun item
    If InvComUsu.SelectedItem = 0 Then
        Call PrintCommerceMsg("¡No tienes ningún item seleccionado!", FontTypeNames.FONTTYPE_FIGHT)

        Exit Sub

    End If
          
    ' Numero invalido
    If Not IsNumeric(txtAgregar.Text) Then Exit Sub
          
    HabilitarConfirmar True
          
    Dim OfferSlot As Byte

    Dim Amount    As Long

    Dim InvSlot   As Byte
              
    With InvComUsu

        
        If .SelectedItem = FLAGORO Then
            If Val(txtAgregar.Text) > InvOroComUsu(0).Amount(1) Then
                Call PrintCommerceMsg("¡No tienes esa cantidad!", FontTypeNames.FONTTYPE_FIGHT)

                Exit Sub

            End If
                  
            Amount = InvOroComUsu(1).Amount(1) + Val(txtAgregar.Text)
          
            ' Le aviso al otro de mi cambio de oferta
            Call WriteUserCommerceOffer(FLAGORO, Val(txtAgregar.Text), GOLD_OFFER_SLOT)
                  
            ' Actualizo los inventarios
            Call InvOroComUsu(0).ChangeSlotItemAmount(1, InvOroComUsu(0).Amount(1) - Val(txtAgregar.Text))
            Call InvOroComUsu(1).ChangeSlotItemAmount(1, Amount)
                  
            Call PrintCommerceMsg("¡Agregaste " & Val(txtAgregar.Text) & " moneda" & IIf(Val(txtAgregar.Text) = 1, "", "s") & " de oro a tu oferta!!", FontTypeNames.FONTTYPE_GUILD)
            'Call InvOroComUsu(0).DrawInventory
            'Call InvOroComUsu(1).DrawInventory
            
        ElseIf .SelectedItem = FLAGELDHIR Then
            If Val(txtAgregar.Text) > InvEldhirComUsu(0).Amount(1) Then
                Call PrintCommerceMsg("¡No tienes esa cantidad!", FontTypeNames.FONTTYPE_FIGHT)

                Exit Sub

            End If
                  
            Amount = InvEldhirComUsu(1).Amount(1) + Val(txtAgregar.Text)
          
            ' Le aviso al otro de mi cambio de oferta
            Call WriteUserCommerceOffer(FLAGELDHIR, Val(txtAgregar.Text), ELDHIR_OFFER_SLOT)
                  
            ' Actualizo los inventarios
            Call InvEldhirComUsu(0).ChangeSlotItemAmount(1, InvEldhirComUsu(0).Amount(1) - Val(txtAgregar.Text))
            Call InvEldhirComUsu(1).ChangeSlotItemAmount(1, Amount)
                  
            Call PrintCommerceMsg("¡Agregaste " & Val(txtAgregar.Text) & " moneda" & IIf(Val(txtAgregar.Text) = 1, "", "s") & " de Eldhir a tu oferta!!", FontTypeNames.FONTTYPE_GUILD)
            'Call InvEldhirComUsu(0).DrawInventory
            'Call InvEldhirComUsu(1).DrawInventory
        ElseIf .SelectedItem > 0 Then

            If Val(txtAgregar.Text) > .Amount(.SelectedItem) Then
                Call PrintCommerceMsg("¡No tienes esa cantidad!", FontTypeNames.FONTTYPE_FIGHT)

                Exit Sub

            End If
                   
            OfferSlot = CheckAvailableSlot(.SelectedItem, Val(txtAgregar.Text))
                  
            ' Hay espacio o lugar donde sumarlo?
            If OfferSlot > 0 Then
                  
                Call PrintCommerceMsg("¡Agregaste " & Val(txtAgregar.Text) & " " & .ItemName(.SelectedItem) & " a tu oferta!!", FontTypeNames.FONTTYPE_GUILD)
                      
                ' Le aviso al otro de mi cambio de oferta
                Call WriteUserCommerceOffer(.SelectedItem, Val(txtAgregar.Text), OfferSlot)
                      
                ' Actualizo el inventario general de comercio
                Call .ChangeSlotItemAmount(.SelectedItem, .Amount(.SelectedItem) - Val(txtAgregar.Text))
                      
                Amount = InvOfferComUsu(0).Amount(OfferSlot) + Val(txtAgregar.Text)
                      
                ' Actualizo los inventarios
                If InvOfferComUsu(0).ObjIndex(OfferSlot) > 0 Then
                    ' Si ya esta el item, solo actualizo su cantidad en el invenatario
                    Call InvOfferComUsu(0).ChangeSlotItemAmount(OfferSlot, Amount)
                Else
                    InvSlot = .SelectedItem
                    ' Si no agrego todo
                    Call InvOfferComUsu(0).SetItem(OfferSlot, .ObjIndex(InvSlot), Amount, 0, .GrhIndex(InvSlot), .ObjType(InvSlot), .MaxHit(InvSlot), .MinHit(InvSlot), .MaxDef(InvSlot), .MinDef(InvSlot), .Valor(InvSlot), .ItemName(InvSlot), .ValorAzul(InvSlot), .CanUse(InvSlot), .MinHitMag(InvSlot), .MaxHitMag(InvSlot), .MinDefMag(InvSlot), .MaxDefMag(InvSlot))
                    
                End If
                
                'Call InvComUsu.DrawInventory
                'Call InvOfferComUsu(0).DrawInventory
                'Call InvOfferComUsu(1).DrawInventory
            End If
        End If

    End With

End Sub

Private Sub imgCancelar_Click()
    Call WriteUserCommerceEnd
End Sub

Private Sub imgConfirmar_Click()

    If Not cBotonConfirmar.IsEnabled Then Exit Sub  ' Deshabilitado
          
    HabilitarConfirmar False
    imgAgregar.visible = False
    imgQuitar.visible = False
    txtAgregar.Enabled = False
          
    Call PrintCommerceMsg("¡Has confirmado tu oferta! Ya no puedes cambiarla.", FontTypeNames.FONTTYPE_CONSE)
    Call WriteUserCommerceConfirm
End Sub

Private Sub imgQuitar_Click()

    Dim Amount     As Long

    Dim InvComSlot As Byte

    ' No tiene seleccionado ningun item
    If InvOfferComUsu(0).SelectedItem = 0 Then
        Call PrintCommerceMsg("¡No tienes ningún ítem seleccionado!", FontTypeNames.FONTTYPE_FIGHT)

        Exit Sub

    End If
          
    ' Numero invalido
    If Not IsNumeric(txtAgregar.Text) Then Exit Sub

    ' Comparar con el inventario para distribuir los items
    If InvOfferComUsu(0).SelectedItem = FLAGORO Then
        Amount = IIf(Val(txtAgregar.Text) > InvOroComUsu(1).Amount(1), InvOroComUsu(1).Amount(1), Val(txtAgregar.Text))
        ' Estoy quitando, paso un valor negativo
        Amount = Amount * (-1)
              
        ' No tiene sentido que se quiten 0 unidades
        If Amount <> 0 Then
            ' Le aviso al otro de mi cambio de oferta
            Call WriteUserCommerceOffer(FLAGORO, Amount, GOLD_OFFER_SLOT)
                  
            ' Actualizo los inventarios
            Call InvOroComUsu(0).ChangeSlotItemAmount(1, InvOroComUsu(0).Amount(1) - Amount)
            Call InvOroComUsu(1).ChangeSlotItemAmount(1, InvOroComUsu(1).Amount(1) + Amount)
              
            Call PrintCommerceMsg("¡¡Quitaste " & Amount * (-1) & " moneda" & IIf(Val(txtAgregar.Text) = 1, "", "s") & " de oro de tu oferta!!", FontTypeNames.FONTTYPE_GUILD)
            
            'Call InvOroComUsu(0).DrawInventory
            'Call InvOroComUsu(1).DrawInventory
        End If

    ElseIf InvOfferComUsu(0).SelectedItem = FLAGELDHIR Then
        Amount = IIf(Val(txtAgregar.Text) > InvEldhirComUsu(1).Amount(1), InvEldhirComUsu(1).Amount(1), Val(txtAgregar.Text))
        ' Estoy quitando, paso un valor negativo
        Amount = Amount * (-1)
              
        ' No tiene sentido que se quiten 0 unidades
        If Amount <> 0 Then
            ' Le aviso al otro de mi cambio de oferta
            Call WriteUserCommerceOffer(FLAGELDHIR, Amount, ELDHIR_OFFER_SLOT)
                  
            ' Actualizo los inventarios
            Call InvEldhirComUsu(0).ChangeSlotItemAmount(1, InvEldhirComUsu(0).Amount(1) - Amount)
            Call InvEldhirComUsu(1).ChangeSlotItemAmount(1, InvEldhirComUsu(1).Amount(1) + Amount)
              
            Call PrintCommerceMsg("¡¡Quitaste " & Amount * (-1) & " moneda" & IIf(Val(txtAgregar.Text) = 1, "", "s") & " de Eldhir de tu oferta!!", FontTypeNames.FONTTYPE_GUILD)
            
           'Call InvEldhirComUsu(0).DrawInventory
            'Call InvEldhirComUsu(1).DrawInventory
        End If
    
    
    Else
        Amount = IIf(Val(txtAgregar.Text) > InvOfferComUsu(0).Amount(InvOfferComUsu(0).SelectedItem), InvOfferComUsu(0).Amount(InvOfferComUsu(0).SelectedItem), Val(txtAgregar.Text))
        ' Estoy quitando, paso un valor negativo
        Amount = Amount * (-1)
              
        ' No tiene sentido que se quiten 0 unidades
        If Amount <> 0 Then

            With InvOfferComUsu(0)
                      
                Call PrintCommerceMsg("¡¡Quitaste " & Amount * (-1) & " " & .ItemName(.SelectedItem) & " de tu oferta!!", FontTypeNames.FONTTYPE_GUILD)
          
                ' Le aviso al otro de mi cambio de oferta
                Call WriteUserCommerceOffer(0, Amount, .SelectedItem)
                  
                ' Actualizo el inventario general
                Call UpdateInvCom(.ObjIndex(.SelectedItem), Abs(Amount))
                       
                ' Actualizo el inventario de oferta
                If .Amount(.SelectedItem) + Amount = 0 Then
                    ' Borro el item
                    Call .SetItem(.SelectedItem, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, vbNullString, 0, False, 0, 0, 0, 0)
                Else
                    ' Le resto la cantidad deseada
                    Call .ChangeSlotItemAmount(.SelectedItem, .Amount(.SelectedItem) + Amount)
                End If
                
               ' Call InvOfferComUsu(0).DrawInventory
               ' Call InvOfferComUsu(1).DrawInventory

            End With

        End If
    End If
          
    ' Si quito todos los items de la oferta, no puede confirmarla
    If Not HasAnyItem(InvOfferComUsu(0)) And Not HasAnyItem(InvOroComUsu(1)) And Not HasAnyItem(InvEldhirComUsu(1)) Then HabilitarConfirmar (False)
End Sub

Private Sub imgRechazar_Click()

    If Not cBotonRechazar.IsEnabled Then Exit Sub  ' Deshabilitado
          
    Call WriteUserCommerceReject
End Sub

Public Sub Form_LoadDetails()
    LoadButtons

    frmComerciarUsu.SendTxt.Text = vbNullString
    
    Call PrintCommerceMsg("> Una vez termines de formar tu oferta, debes presionar en ""Confirmar"", tras lo cual ya no podrás modificarla.", FontTypeNames.FONTTYPE_GUILDMSG)
    Call PrintCommerceMsg("> Luego que el otro usuario confirme su oferta, podrás aceptarla o rechazarla. Si la rechazas, se terminará el comercio.", FontTypeNames.FONTTYPE_GUILDMSG)
    Call PrintCommerceMsg("> Cuando ambos acepten la oferta del otro, se realizará el intercambio.", FontTypeNames.FONTTYPE_GUILDMSG)
    Call PrintCommerceMsg("> Si se intercambian más ítems de los que pueden entrar en tu inventario, es probable que caigan al suelo, así que presta mucha atención a esto.", FontTypeNames.FONTTYPE_GUILDMSG)
          
    HabilitarConfirmar False
    imgAgregar.visible = True
    imgQuitar.visible = True
    txtAgregar.Enabled = True
    txtAgregar.Text = ""
End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me

    'Me.Picture = LoadPicture(DirGraficos & "VentanaComercioUsuario.jpg")
          
    Call Form_LoadDetails


 ' Ventana de comercio con USUARIOS
    g_Captions(eCaption.cInvComUsu) = wGL_Graphic.Create_Device_From_Display(frmComerciarUsu.picInvComercio.hWnd, frmComerciarUsu.picInvComercio.ScaleWidth, frmComerciarUsu.picInvComercio.ScaleHeight)
    g_Captions(eCaption.cInvOfferComUsu1) = wGL_Graphic.Create_Device_From_Display(frmComerciarUsu.picInvOfertaProp.hWnd, frmComerciarUsu.picInvOfertaProp.ScaleWidth, frmComerciarUsu.picInvOfertaProp.ScaleHeight)
    g_Captions(eCaption.cInvOfferComUsu2) = wGL_Graphic.Create_Device_From_Display(frmComerciarUsu.picInvOfertaOtro.hWnd, frmComerciarUsu.picInvOfertaOtro.ScaleWidth, frmComerciarUsu.picInvOfertaOtro.ScaleHeight)
    g_Captions(eCaption.cInvOroComUsu1) = wGL_Graphic.Create_Device_From_Display(frmComerciarUsu.picInvOroProp.hWnd, frmComerciarUsu.picInvOroProp.ScaleWidth, frmComerciarUsu.picInvOroProp.ScaleHeight)
    g_Captions(eCaption.cInvOroComUsu2) = wGL_Graphic.Create_Device_From_Display(frmComerciarUsu.picInvOroOfertaProp.hWnd, frmComerciarUsu.picInvOroOfertaProp.ScaleWidth, frmComerciarUsu.picInvOroOfertaProp.ScaleHeight)
    g_Captions(eCaption.cInvOroComUsu3) = wGL_Graphic.Create_Device_From_Display(frmComerciarUsu.picInvOroOfertaOtro.hWnd, frmComerciarUsu.picInvOroOfertaOtro.ScaleWidth, frmComerciarUsu.picInvOroOfertaOtro.ScaleHeight)
    
    g_Captions(eCaption.cInvEldhirComUsu1) = wGL_Graphic.Create_Device_From_Display(frmComerciarUsu.picInvEldhirProp.hWnd, frmComerciarUsu.picInvEldhirProp.ScaleWidth, frmComerciarUsu.picInvEldhirProp.ScaleHeight)
    g_Captions(eCaption.cInvEldhirComUsu2) = wGL_Graphic.Create_Device_From_Display(frmComerciarUsu.picInvEldhirOfertaProp.hWnd, frmComerciarUsu.picInvEldhirOfertaProp.ScaleWidth, frmComerciarUsu.picInvEldhirOfertaProp.ScaleHeight)
    g_Captions(eCaption.cInvEldhirComUsu3) = wGL_Graphic.Create_Device_From_Display(frmComerciarUsu.picInvEldhirOfertaOtro.hWnd, frmComerciarUsu.picInvEldhirOfertaOtro.ScaleWidth, frmComerciarUsu.picInvEldhirOfertaOtro.ScaleHeight)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.cInvComUsu))
    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.cInvOfferComUsu1))
    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.cInvOfferComUsu2))
    
    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.cInvOroComUsu1))
    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.cInvOroComUsu2))
    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.cInvOroComUsu3))
    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.cInvEldhirComUsu1))
    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.cInvEldhirComUsu2))
    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.cInvEldhirComUsu3))
End Sub

Private Sub LoadButtons()

    Dim GrhPath As String

    GrhPath = DirGraficos
          
    Set cBotonAceptar = New clsGraphicalButton
    Set cBotonConfirmar = New clsGraphicalButton
    Set cBotonRechazar = New clsGraphicalButton
    Set cBotonCancelar = New clsGraphicalButton
          
    Set LastButtonPressed = New clsGraphicalButton
    
    Call cBotonAceptar.Initialize(imgAceptar, vbNullString, vbNullString, vbNullString, Me, vbNullString, True)
                                    
    Call cBotonConfirmar.Initialize(imgConfirmar, vbNullString, vbNullString, vbNullString, Me, vbNullString, True)
                                        
    Call cBotonRechazar.Initialize(imgRechazar, vbNullString, vbNullString, vbNullString, Me, vbNullString, True)
                                        
    Call cBotonCancelar.Initialize(ImgCancelar, vbNullString, vbNullString, vbNullString, Me)
                                        
End Sub

Private Sub Form_LostFocus()
    Me.SetFocus
End Sub

Private Sub SubtxtAgregar_Change()

    If Val(txtAgregar.Text) < 1 Then txtAgregar.Text = "1"

    If Val(txtAgregar.Text) > 2147483647 Then txtAgregar.Text = "2147483647"
End Sub

Private Sub picInvComercio_Click()
    Call InvOroComUsu(0).DeselectItem
    Call InvEldhirComUsu(0).DeselectItem
    
   ' InvComUsu.DrawInventory
End Sub

Private Sub picInvComercio_MouseMove(Button As Integer, _
                                     Shift As Integer, _
                                     X As Single, _
                                     Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub picInvOfertaOtro_MouseMove(Button As Integer, _
                                       Shift As Integer, _
                                       X As Single, _
                                       Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub picInvOfertaProp_Click()
    InvOroComUsu(1).DeselectItem
    InvEldhirComUsu(1).DeselectItem
End Sub

Private Sub picInvOfertaProp_MouseMove(Button As Integer, _
                                       Shift As Integer, _
                                       X As Single, _
                                       Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub picInvOroOfertaOtro_Click()
    ' No se puede seleccionar el oro que oferta el otro :P
    InvOroComUsu(2).DeselectItem
    InvEldhirComUsu(2).DeselectItem
End Sub
Private Sub picInvEldhirOfertaOtro_Click()
    ' No se puede seleccionar el oro que oferta el otro :P
    InvOroComUsu(2).DeselectItem
    InvEldhirComUsu(2).DeselectItem
End Sub
Private Sub picInvOroOfertaProp_Click()
    InvOfferComUsu(0).SelectGold
End Sub
Private Sub picInvEldhirOfertaProp_Click()
    InvOfferComUsu(0).SelectEldhir
End Sub

Private Sub picInvOroProp_Click()
    InvComUsu.SelectGold
End Sub
Private Sub picInvEldhirProp_Click()
    InvComUsu.SelectEldhir
End Sub
Private Sub SendTxt_Change()

    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 03/10/2009
    '**************************************************************
    If Len(SendTxt.Text) > 160 Then
        sCommerceChat = "Soy un cheater, avisenle a un gm"
    Else

        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim I         As Long

        Dim TempStr   As String

        Dim CharAscii As Integer
              
        For I = 1 To Len(SendTxt.Text)
            CharAscii = Asc(mid$(SendTxt.Text, I, 1))

            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                TempStr = TempStr & Chr$(CharAscii)
            End If

        Next I
              
        If TempStr <> SendTxt.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendTxt.Text = TempStr
        End If
              
        sCommerceChat = SendTxt.Text
        
        
    End If

End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)

    If Not (KeyAscii = vbKeyBack) And Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then KeyAscii = 0
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)

    'Send text
    If KeyCode = vbKeyReturn Then
        If LenB(sCommerceChat) <> 0 Then Call WriteCommerceChat(sCommerceChat)
              
        sCommerceChat = ""
        SendTxt.Text = ""
        KeyCode = 0
    End If

End Sub

Private Sub txtAgregar_Change()

    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 03/10/2009
    '**************************************************************
    'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
    Dim I         As Long

    Dim TempStr   As String

    Dim CharAscii As Integer
          
    For I = 1 To Len(txtAgregar.Text)
        CharAscii = Asc(mid$(txtAgregar.Text, I, 1))
              
        If CharAscii >= 48 And CharAscii <= 57 Then
            TempStr = TempStr & Chr$(CharAscii)
        End If

    Next I
          
    If TempStr <> txtAgregar.Text Then
        'We only set it if it's different, otherwise the event will be raised
        'constantly and the client will crush
        txtAgregar.Text = TempStr
    End If

End Sub

Private Sub txtAgregar_KeyDown(KeyCode As Integer, Shift As Integer)

    If Not ((KeyCode >= 48 And KeyCode <= 57) Or KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Or (KeyCode >= 37 And KeyCode <= 40)) Then
        KeyCode = 0
    End If

End Sub

Private Sub txtAgregar_KeyPress(KeyAscii As Integer)

    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or (KeyAscii >= 37 And KeyAscii <= 40)) Then
        'txtCant = KeyCode
        KeyAscii = 0
    End If

End Sub

Private Function CheckAvailableSlot(ByVal InvSlot As Byte, ByVal Amount As Long) As Byte

    '***************************************************
    'Author: ZaMa
    'Last Modify Date: 30/11/2009
    'Search for an available slot to put an item. If found returns the slot, else returns 0.
    '***************************************************
    Dim Slot As Long

    On Error GoTo err

    ' Primero chequeo si puedo sumar esa cantidad en algun slot que ya tenga ese item
    For Slot = 1 To INV_OFFER_SLOTS

        If InvComUsu.ObjIndex(InvSlot) = InvOfferComUsu(0).ObjIndex(Slot) Then
            If InvOfferComUsu(0).Amount(Slot) + Amount <= MAX_INVENTORY_OBJS Then
                ' Puedo sumarlo aca
                CheckAvailableSlot = Slot

                Exit Function

            End If
        End If

    Next Slot
          
    ' No lo puedo sumar, me fijo si hay alguno vacio
    For Slot = 1 To INV_OFFER_SLOTS

        If InvOfferComUsu(0).ObjIndex(Slot) = 0 Then
            ' Esta vacio, lo dejo aca
            CheckAvailableSlot = Slot

            Exit Function

        End If

    Next Slot

    Exit Function

err:
    Debug.Print "Slot: " & Slot
End Function

Public Sub UpdateInvCom(ByVal ObjIndex As Integer, ByVal Amount As Long)

    Dim Slot            As Byte

    Dim RemainingAmount As Long

    Dim DifAmount       As Long
          
    RemainingAmount = Amount
          
    For Slot = 1 To MAX_INVENTORY_SLOTS
              
        If InvComUsu.ObjIndex(Slot) = ObjIndex Then
            DifAmount = Inventario.Amount(Slot) - InvComUsu.Amount(Slot)

            If DifAmount > 0 Then
                If RemainingAmount > DifAmount Then
                    RemainingAmount = RemainingAmount - DifAmount
                    Call InvComUsu.ChangeSlotItemAmount(Slot, Inventario.Amount(Slot))
                    
                Else
                    Call InvComUsu.ChangeSlotItemAmount(Slot, InvComUsu.Amount(Slot) + RemainingAmount)
                    'Call InvComUsu.DrawInventory

                    Exit Sub

                End If
            End If
        End If

    Next Slot

End Sub

Public Sub PrintCommerceMsg(ByRef msg As String, ByVal FontIndex As Integer)
          
    With FontTypes(FontIndex)
        Call AddtoRichTextBox(frmComerciarUsu.CommerceConsole, msg, .red, .green, .blue, .bold, .italic)
    End With
          
End Sub

Public Function HasAnyItem(ByRef Inventory As clsGrapchicalInventory) As Boolean

    Dim Slot As Long
          
    For Slot = 1 To Inventory.MaxObjs

        If Inventory.Amount(Slot) > 0 Then HasAnyItem = True: Exit Function
    Next Slot
          
End Function

Public Sub HabilitarConfirmar(ByVal Habilitar As Boolean)
    Call cBotonConfirmar.EnableButton(Habilitar)
End Sub

Public Sub HabilitarAceptarRechazar(ByVal Habilitar As Boolean)
    Call cBotonAceptar.EnableButton(Habilitar)
    Call cBotonRechazar.EnableButton(Habilitar)
End Sub

