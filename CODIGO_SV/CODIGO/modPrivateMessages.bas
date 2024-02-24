Attribute VB_Name = "modPrivateMessages"
Option Explicit

Public Sub AgregarMensaje(ByVal UserIndex As Integer, _
                          ByRef Autor As String, _
                          ByRef Mensaje As String)
        '<EhHeader>
        On Error GoTo AgregarMensaje_Err
        '</EhHeader>

        '***************************************************
        'Author: Amraphen
        'Last Modification: 04/08/2011
        'Agrega un nuevo mensaje privado a un usuario online.
        '***************************************************
        Dim LoopC As Long

100     With UserList(UserIndex)

102         If .UltimoMensaje < MAX_PRIVATE_MESSAGES Then
104             .UltimoMensaje = .UltimoMensaje + 1
            Else

106             For LoopC = 1 To MAX_PRIVATE_MESSAGES - 1
108                 .Mensajes(LoopC) = .Mensajes(LoopC + 1)
                Next

            End If
        
110         With .Mensajes(.UltimoMensaje)
112             .Contenido = UCase$(Autor) & ": " & Mensaje & " (" & Now & ")"
114             .Nuevo = True
            End With
        
116         Call WriteConsoleMsg(UserIndex, "¡Has recibido un mensaje privado de un Game Master!", FontTypeNames.FONTTYPE_GM)
        End With

        '<EhFooter>
        Exit Sub

AgregarMensaje_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modPrivateMessages.AgregarMensaje " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub AgregarMensajeOFF(ByRef Destinatario As String, _
                             ByRef Autor As String, _
                             ByRef Mensaje As String)
        '<EhHeader>
        On Error GoTo AgregarMensajeOFF_Err
        '</EhHeader>

        '***************************************************
        'Author: Amraphen
        'Last Modification: 04/08/2011
        'Agrega un nuevo mensaje privado a un usuario offline.
        '***************************************************
        Dim UltimoMensaje As Byte

        Dim Charfile      As String

        Dim Contenido     As String

        Dim LoopC         As Long

100     Charfile = CharPath & Destinatario & ".chr"
102     UltimoMensaje = CByte(GetVar(Charfile, "MENSAJES", "UltimoMensaje"))
104     Contenido = UCase$(Autor) & ": " & Mensaje & " (" & Now & ")"

106     If UltimoMensaje < MAX_PRIVATE_MESSAGES Then
108         UltimoMensaje = UltimoMensaje + 1
        Else

110         For LoopC = 1 To MAX_PRIVATE_MESSAGES - 1
112             Call WriteVar(Charfile, "MENSAJES", "MSJ" & LoopC, GetVar(Charfile, "MENSAJES", "MSJ" & LoopC + 1))
114             Call WriteVar(Charfile, "MENSAJES", "MSJ" & LoopC & "_NUEVO", GetVar(Charfile, "MENSAJES", "MSJ" & LoopC + 1 & "_NUEVO"))
116         Next LoopC

        End If
        
118     Call WriteVar(Charfile, "MENSAJES", "MSJ" & UltimoMensaje, Contenido)
120     Call WriteVar(Charfile, "MENSAJES", "MSJ" & UltimoMensaje & "_NUEVO", 1)
    
122     Call WriteVar(Charfile, "MENSAJES", "UltimoMensaje", UltimoMensaje)
        '<EhFooter>
        Exit Sub

AgregarMensajeOFF_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modPrivateMessages.AgregarMensajeOFF " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Function TieneMensajesNuevos(ByVal UserIndex As Integer) As Boolean
        '<EhHeader>
        On Error GoTo TieneMensajesNuevos_Err
        '</EhHeader>

        '***************************************************
        'Author: Amraphen
        'Last Modification: 04/08/2011
        'Determina si el usuario tiene mensajes nuevos.
        '***************************************************
        Dim LoopC As Long

100     For LoopC = 1 To MAX_PRIVATE_MESSAGES

102         If UserList(UserIndex).Mensajes(LoopC).Nuevo Then
104             TieneMensajesNuevos = True

                Exit Function

            End If

106     Next LoopC
    
108     TieneMensajesNuevos = False
        '<EhFooter>
        Exit Function

TieneMensajesNuevos_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modPrivateMessages.TieneMensajesNuevos " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Sub GuardarMensajes(ByRef IUser As User, ByRef Manager As clsIniManager)
        '<EhHeader>
        On Error GoTo GuardarMensajes_Err
        '</EhHeader>

        '***************************************************
        'Author: Amraphen
        'Last Modification: 04/08/2011
        'Guarda los mensajes del usuario.
        '***************************************************
        Dim LoopC As Long
    
100     With IUser
102         Call Manager.ChangeValue("MENSAJES", "UltimoMensaje", CStr(.UltimoMensaje))
        
104         For LoopC = 1 To MAX_PRIVATE_MESSAGES
106             Call Manager.ChangeValue("MENSAJES", "MSJ" & LoopC, .Mensajes(LoopC).Contenido)

108             If .Mensajes(LoopC).Nuevo Then
110                 Call Manager.ChangeValue("MENSAJES", "MSJ" & LoopC & "_NUEVO", 1)
                Else
112                 Call Manager.ChangeValue("MENSAJES", "MSJ" & LoopC & "_NUEVO", 0)
                End If

114         Next LoopC

        End With

        '<EhFooter>
        Exit Sub

GuardarMensajes_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modPrivateMessages.GuardarMensajes " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub CargarMensajes(ByVal UserIndex As Integer, ByRef Manager As clsIniManager)
        '<EhHeader>
        On Error GoTo CargarMensajes_Err
        '</EhHeader>

        '***************************************************
        'Author: Amraphen
        'Last Modification: 04/08/2011
        'Carga los mensajes del usuario.
        '***************************************************
        Dim LoopC As Long

100     With UserList(UserIndex)
102         .UltimoMensaje = val(Manager.GetValue("MENSAJES", "UltimoMensaje"))
        
104         For LoopC = 1 To MAX_PRIVATE_MESSAGES

106             With .Mensajes(LoopC)
108                 .Nuevo = val(Manager.GetValue("MENSAJES", "MSJ" & LoopC & "_NUEVO"))
110                 .Contenido = CStr(Manager.GetValue("MENSAJES", "MSJ" & LoopC))
                End With

112         Next LoopC

        End With

        '<EhFooter>
        Exit Sub

CargarMensajes_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modPrivateMessages.CargarMensajes " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub LimpiarMensajeSlot(ByVal UserIndex As Integer, ByVal Slot As Byte)
        '<EhHeader>
        On Error GoTo LimpiarMensajeSlot_Err
        '</EhHeader>

        '***************************************************
        'Author: Amraphen
        'Last Modification: 04/08/2011
        'Limpia el un mensaje de un usuario online.
        '***************************************************
100     With UserList(UserIndex).Mensajes(Slot)
102         .Contenido = vbNullString
104         .Nuevo = False
        End With

        '<EhFooter>
        Exit Sub

LimpiarMensajeSlot_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modPrivateMessages.LimpiarMensajeSlot " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub LimpiarMensajes(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo LimpiarMensajes_Err
        '</EhHeader>

        '***************************************************
        'Author: Amraphen
        'Last Modification: 04/08/2011
        'Limpia los mensajes del slot.
        '***************************************************
        Dim LoopC As Long

100     With UserList(UserIndex)
102         .UltimoMensaje = 0
        
104         For LoopC = 1 To MAX_PRIVATE_MESSAGES
106             Call LimpiarMensajeSlot(UserIndex, LoopC)
108         Next LoopC

        End With

        '<EhFooter>
        Exit Sub

LimpiarMensajes_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modPrivateMessages.LimpiarMensajes " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub BorrarMensaje(ByVal UserIndex As Integer, ByVal Slot As Byte)
        '<EhHeader>
        On Error GoTo BorrarMensaje_Err
        '</EhHeader>

        '***************************************************
        'Author: Amraphen
        'Last Modification: 04/08/2011
        'Borra un mensaje de un usuario.
        '***************************************************
        Dim LoopC As Long

100     With UserList(UserIndex)

102         If Slot > .UltimoMensaje Or Slot < 1 Then Exit Sub

104         If Slot = .UltimoMensaje Then
106             Call LimpiarMensajeSlot(UserIndex, Slot)
            Else

108             For LoopC = Slot To MAX_PRIVATE_MESSAGES - 1
110                 .Mensajes(LoopC) = .Mensajes(LoopC + 1)
112             Next LoopC

114             Call LimpiarMensajeSlot(UserIndex, .UltimoMensaje)
            End If
        
116         .UltimoMensaje = .UltimoMensaje - 1
        End With

        '<EhFooter>
        Exit Sub

BorrarMensaje_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modPrivateMessages.BorrarMensaje " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub BorrarMensajeOFF(ByVal UserName As String, ByVal Slot As Byte)
        '<EhHeader>
        On Error GoTo BorrarMensajeOFF_Err
        '</EhHeader>

        '***************************************************
        'Author: Amraphen
        'Last Modification: 04/08/2011
        'Borra un mensaje de un usuario.
        '***************************************************
        Dim Charfile      As String

        Dim UltimoMensaje As Byte

        Dim LoopC         As Long

100     Charfile = CharPath & UserName & ".chr"
    
102     UltimoMensaje = GetVar(Charfile, "MENSAJES", "UltimoMensaje")
    
104     If Slot > UltimoMensaje Or Slot < 1 Then Exit Sub
    
106     If Slot = UltimoMensaje Then
108         Call WriteVar(Charfile, "MENSAJES", "MSJ" & Slot, vbNullString)
110         Call WriteVar(Charfile, "MENSAJES", "MSJ" & Slot & "_Nuevo", vbNullString)
        Else

112         For LoopC = Slot To UltimoMensaje - 1
114             Call WriteVar(Charfile, "MENSAJES", "MSJ" & LoopC, GetVar(Charfile, "MENSAJES", "MSJ" & LoopC + 1))
116             Call WriteVar(Charfile, "MENSAJES", "MSJ" & LoopC & "_NUEVO", GetVar(Charfile, "MENSAJES", "MSJ" & LoopC + 1 & "_NUEVO"))
118         Next LoopC

120         Call WriteVar(Charfile, "MENSAJES", "MSJ" & UltimoMensaje, vbNullString)
122         Call WriteVar(Charfile, "MENSAJES", "MSJ" & UltimoMensaje & "_Nuevo", vbNullString)
        End If
    
124     Call WriteVar(Charfile, "MENSAJES", "UltimoMensaje", UltimoMensaje - 1)
        '<EhFooter>
        Exit Sub

BorrarMensajeOFF_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modPrivateMessages.BorrarMensajeOFF " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub LimpiarMensajesOFF(ByVal UserName As String)
        '<EhHeader>
        On Error GoTo LimpiarMensajesOFF_Err
        '</EhHeader>

        '***************************************************
        'Author: Amraphen
        'Last Modification: 18/08/2011
        'Borra los mensajes de un usuario offline.
        '***************************************************
        Dim Charfile      As String

        Dim UltimoMensaje As Byte

        Dim LoopC         As Long

100     Charfile = CharPath & UserName & ".chr"
    
102     UltimoMensaje = GetVar(Charfile, "MENSAJES", "UltimoMensaje")
    
104     If UltimoMensaje > 0 Then

106         For LoopC = 1 To UltimoMensaje
108             Call WriteVar(Charfile, "MENSAJES", "MSJ" & LoopC, vbNullString)
110             Call WriteVar(Charfile, "MENSAJES", "MSJ" & LoopC & "_NUEVO", vbNullString)
112         Next LoopC
        
114         Call WriteVar(Charfile, "MENSAJES", "UltimoMensaje", 0)
        End If

        '<EhFooter>
        Exit Sub

LimpiarMensajesOFF_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.modPrivateMessages.LimpiarMensajesOFF " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
