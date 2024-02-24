Attribute VB_Name = "mMao"
' Author @ Lautaro
' Sistema de Mercado Estilo Tierras del Sur
' 03/07/2022 03:20hs-05:31hs | 14:10-19:26

' DISEÑO
' Esto se guarda en la cuenta
'[SALE]
'Last = 5
'1=SLOT-1-LION
'2=SLOT-2-GORO-LION
'3=SLOT-1-GORO
'4=SLOT-1-KIM DOTCOM
'5=SLOT-3-LION-KIM DOTCOM-GORO

'[SALEOFFER]
' (Carga el Slot del Mercado en la oferta)
'Last = 5
'1=SLOT-ACCOUNT-ORO-1-LION
'2=SLOT-ACCOUNT-ORO-2-GORO-LION
'3=SLOT-ACCOUNT-ORO-1-GORO
'4=SLOT-ACCOUNT-ORO-1-KIM DOTCOM
'5=SLOT-ACCOUNT-ORO-3-LION-KIM DOTCOM-GORO

' En el Mercader.DAT
'[INIT]
'Last = 5

'[SALE]
'1=ACCOUNT-ORO-1-LION
'2=ACCOUNT-ORO-2-GORO-LION
'3=ACCOUNT-ORO-1-GORO
'4=ACCOUNT-ORO-1-KIM DOTCOM
'5=ACCOUNT-ORO-3-LION-KIM DOTCOM-GORO

Option Explicit

Public Const MERCADER_MAX_LIST As Integer = 255 '
Public Const MERCADER_MAX_GLD As Long = 2000000000 ' 2.000.000.000
Public Const MERCADER_MAX_DSP             As Long = 100000 '100.000

Public Const MERCADER_GLD_SALE As Long = 1500 ' 3.000 pide de base de Monedas de oro.
Public Const MERCADER_MIN_LVL As Byte = 15    ' Pide Nivel 15 para poder ser publicado.

Public Const MERCADER_MAX_OFFER As Byte = 50
Public Const MERCADER_OFFER_TIME As Long = 120000


Public Type tMercaderObj
    ObjIndex As Integer
    Amount As Integer
End Type

Public Type tMercaderCharInfo
    Name As String
    
    Body As Integer
    Head As Integer
    Weapon As Integer
    Shield As Integer
    Helm As Integer
    
    Elv As Byte
    Exp As Long
    Elu As Long
    
    Hp As Integer
    Constitucion As Byte
    
    Class As Byte
    Raze As Byte
    
    Faction As Byte
    FactionRange As Byte
    FragsCiu As Integer
    FragsCri As Integer
    FragsOther As Integer
    
    Gld As Long
    GuildIndex As Integer
    
    Object() As tMercaderObj
    Bank() As tMercaderObj
    Spells() As Byte
    Skills() As Byte
End Type

Public Type tMercaderChar

    Desc As String
    Dsp As Long
    Gld As Long
    Account As String
    Count As Byte
    NameU() As String
    Info() As tMercaderCharInfo
End Type

Public Type tMercader
    Chars As tMercaderChar
    
    LastOffer As Byte
    Offer(1 To MERCADER_MAX_OFFER) As tMercaderChar
    OfferTime(1 To MERCADER_MAX_OFFER) As Long
    Slot As Integer
End Type

Public MercaderList(1 To MERCADER_MAX_LIST) As tMercader

' Path del Mercado.
Private Function FilePath()
    FilePath = DatPath & "mercader.ini"
End Function

' Guardamos la información del Mercado
Private Sub Mercader_Save(ByVal Slot As Byte)

        '<EhHeader>
        On Error GoTo Mercader_Save_Err

        '</EhHeader>
        Dim Manager As clsIniManager
    
100     Set Manager = New clsIniManager
        
        If FileExist(FilePath) Then
102         Manager.Initialize (FilePath)

        End If
            
104     With MercaderList(Slot)
106         Call Manager.ChangeValue("SALE" & Slot, "ACCOUNT", .Chars.Account)
108         Call Manager.ChangeValue("SALE" & Slot, "GLD", CStr(.Chars.Gld))
        
110         Call Manager.ChangeValue("SALE" & Slot, "CHAR", CStr(.Chars.Count))
        
112         If .Chars.Count > 0 Then
114             Call Manager.ChangeValue("SALE" & Slot, "CHARS", Mercader_Generate_Text_Chars(.Chars.NameU))
            Else
                Call Manager.ChangeValue("SALE" & Slot, "CHARS", vbNullString)

            End If
        
116         Manager.DumpFile FilePath

        End With
    
118     Set Manager = Nothing
        '<EhFooter>
        Exit Sub

Mercader_Save_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.mMao.Mercader_Save " & "at line " & Erl

        

        '</EhFooter>
End Sub

' Guardamos la información del Mercado en la Cuenta del Personaje
Private Sub Mercader_SaveUser(ByVal UserIndex As Integer, _
                              ByVal Slot As Integer, _
                              ByRef Mercader As tMercaderChar)
        '<EhHeader>
        On Error GoTo Mercader_SaveUser_Err
        '</EhHeader>
                              
        Dim Manager As clsIniManager
        Dim A As Long
        Dim FileP As String
    
100     Set Manager = New clsIniManager
    
102     FileP = AccountPath & UserList(UserIndex).Account.Email & ".acc"
    
104     Call Manager.Initialize(FileP)
    
106     Call Manager.ChangeValue("SALE", "LAST", CStr(Slot))
    
108     Call Manager.DumpFile(FileP)
    
110     Set Manager = Nothing
    
        '<EhFooter>
        Exit Sub

Mercader_SaveUser_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mMao.Mercader_SaveUser " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

' Cargamos la Lista de Mercado
Public Sub Mercader_Load()
        '<EhHeader>
        On Error GoTo Mercader_Load_Err
        '</EhHeader>
        Dim Manager As clsIniManager
        Dim A       As Long, B As Long
    
100     Set Manager = New clsIniManager
    
102     If FileExist(FilePath, vbArchive) Then
104         Manager.Initialize FilePath
        End If
    
        Dim Temp    As String
        Dim TempA() As String
    
106     For A = 1 To MERCADER_MAX_LIST
108         With MercaderList(A)
110             .Chars.Account = Manager.GetValue("SALE" & A, "ACCOUNT")
112             .Chars.Gld = val(Manager.GetValue("SALE" & A, "GLD"))
114             .Chars.Count = val(Manager.GetValue("SALE" & A, "CHAR"))
            
116             If .Chars.Count > 0 Then
118                 Temp = Manager.GetValue("SALE" & A, "CHARS")
120                 TempA = Split(Temp, "-")
                
                      ReDim .Chars.NameU(1 To .Chars.Count) As String
                      ReDim .Chars.Info(1 To .Chars.Count) As tMercaderCharInfo
                      
122                 For B = 1 To .Chars.Count
124                     .Chars.NameU(B) = TempA(B - 1)
126                     .Chars.Info(B) = Mercader_SetChar(B, .Chars.NameU(B), 0)
                    Next
    
                End If
            
            End With
128     Next A
    
130     Set Manager = Nothing
        '<EhFooter>
        Exit Sub

Mercader_Load_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mMao.Mercader_Load " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
' Busca un slot libre en el mercado (Máximo=MERCADER_MAX_LIST)
Private Function Mercader_FreeSlot(Optional ByVal Premium As Byte = 0) As Integer
        '<EhHeader>
        On Error GoTo Mercader_FreeSlot_Err
        '</EhHeader>
        Dim A As Long
    
110     For A = 1 To MERCADER_MAX_LIST
112         If MercaderList(A).Chars.Count = 0 Then
114             Mercader_FreeSlot = A
                Exit Function
            End If
116     Next A
        '<EhFooter>
        Exit Function

Mercader_FreeSlot_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mMao.Mercader_FreeSlot " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

' Chequeo de 'Hack' al postear personajes erroneos.
' Devuelve la lista de nicks cargados de la cuenta.
' Calculo de cuanto le sale en Monedas de oRO automatico. Realizado con Formula de * según cant + precio base
Public Function Mercader_CheckingChar(ByVal UserIndex As Integer, _
                                      ByRef Chars() As Byte, _
                                      ByRef SumaLvls As Long) As Boolean
        '<EhHeader>
        On Error GoTo Mercader_CheckingChar_Err
        '</EhHeader>
        Dim sChars() As String
    
        Dim A        As Long
        Dim UserName As String
    
        Dim TempLvls As Long
        Dim Temp As Long
    
100     For A = LBound(Chars) To UBound(Chars)
102         If Chars(A) = 1 Then
104             UserName = UCase$(UserList(UserIndex).Account.Chars(A).Name)
                
106             If UserName = vbNullString Then
                    Exit Function ' No hay nada en ese slot
                End If
            
108             Temp = val(GetVar(CharPath & UserName & ".chr", "STATS", "ELV"))
            
110             If Temp < MERCADER_MIN_LVL Then
                    Exit Function ' No tiene el nivel correspondiente
                End If
            
112             TempLvls = TempLvls + Temp
            
114             Temp = val(GetVar(CharPath & UserName & ".chr", "FLAGS", "BAN"))
            
116             If Temp > 0 Then
                    Exit Function ' El personaje está baneado. Cliente message
                End If
            End If
118     Next A
    
120     SumaLvls = TempLvls
122     Mercader_CheckingChar = True
    
        '<EhFooter>
        Exit Function

Mercader_CheckingChar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mMao.Mercader_CheckingChar " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

' Chequeo de 'Hack' al postear personajes erroneos.
Public Function Mercader_CheckingChar_Offer(ByRef Chars() As String, _
                                            ByVal sAccount As String) As Boolean
        '<EhHeader>
        On Error GoTo Mercader_CheckingChar_Offer_Err
        '</EhHeader>
    
        Dim A        As Long
        Dim Temp As String
    
100     For A = LBound(Chars) To UBound(Chars)
        
102         Temp = GetVar(CharPath & Chars(A) & ".chr", "INIT", "ACCOUNTNAME")
        
104         If Not StrComp(Temp, sAccount) = 0 Then
                Exit Function ' No está más en la cuenta
            End If
        
106     Next A
    
108     Mercader_CheckingChar_Offer = True
    
        '<EhFooter>
        Exit Function

Mercader_CheckingChar_Offer_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mMao.Mercader_CheckingChar_Offer " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

' Comprueba si la publicación es válida
' FORMULA =  BASE DE ORO * CANTIDAD DE PJS * SUMA DE NIVELES
'
Public Function Mercader_CheckingNew(ByVal UserIndex As Integer, _
                                     ByRef Chars() As Byte, _
                                     ByRef Mercader As tMercaderChar, _
                                     ByRef SaleCost As Long, _
                                     ByVal Blocked As Byte) As Boolean

        '<EhHeader>
        On Error GoTo Mercader_CheckingNew_Err

        '</EhHeader>
        Dim Suma_Lvls As Long

        Dim A         As Long

        Dim tUser     As Integer
    
100     If Mercader.Gld > MERCADER_MAX_GLD Or Mercader.Gld < 1 Then

            Exit Function ' No se permite tanto ORO. Mensaje informativo en el Cliente

        End If
    
102     If Not Mercader_CheckingChar(UserIndex, Chars, Suma_Lvls) Then Exit Function

        ' No tiene suficiente Oro en cuenta para realizar la publicación.
104
        If UserList(UserIndex).Account.Premium > 2 Then
            SaleCost = 0
        Else
            SaleCost = MERCADER_GLD_SALE * Suma_Lvls
            
106         If UserList(UserIndex).Account.Gld < SaleCost Then
                ' Mensaje en el Cliente
                Exit Function
            End If

        End If
        
        Dim LastChar As Byte
    
        ' Setting Chars String
108     For A = LBound(Chars) To UBound(Chars)

110         If Chars(A) = 1 Then
112             LastChar = LastChar + 1
                ReDim Preserve Mercader.NameU(1 To LastChar) As String
                ReDim Preserve Mercader.Info(1 To LastChar) As tMercaderCharInfo
                  
114             Mercader.NameU(LastChar) = UCase$(UserList(UserIndex).Account.Chars(A).Name)
116             Mercader.Count = Mercader.Count + 1
            
118             If Blocked = 1 Then
120                 tUser = NameIndex(Mercader.NameU(LastChar))
                
122                 If tUser > 0 Then
124                     Call WriteErrorMsg(tUser, "Tu personaje pasará a estar bloqueado debido a una Publicación/Oferta.")
126                     Call WriteDisconnect(tUser)
128                     Call FlushBuffer(tUser)
130                     Call CloseSocket(tUser)

                    End If

                End If
            
132             Mercader.Info(LastChar) = Mercader_SetChar(LastChar, Mercader.NameU(LastChar), Blocked)

            End If

134     Next A
        
        ' Posible Hack de los PREMIUM, que no creo que paguen para estafar pero por las dudas!
        If LastChar > 1 Then
            If UserList(UserIndex).Account.Premium < 2 Then
                Exit Function

            End If

        End If
    
136     Mercader_CheckingNew = True
        '<EhFooter>
        Exit Function

Mercader_CheckingNew_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.mMao.Mercader_CheckingNew " & "at line " & Erl

        

        '</EhFooter>
End Function

' Nueva combinación de PJS
Public Sub Mercader_AddList(ByVal UserIndex As Integer, _
                            ByRef Chars() As Byte, _
                            ByRef Mercader As tMercaderChar, _
                            ByVal Blocked As Byte)

        '<EhHeader>
        On Error GoTo Mercader_AddList_Err

        '</EhHeader>
        Dim Slot     As Long

        Dim SaleCost As Long

        Dim SlotUser As Long
        
100     Slot = Mercader_FreeSlot(UserList(UserIndex).Account.Premium)
    
102     If Slot = 0 Then
104         Call WriteErrorMsg(UserIndex, "¡No hay más espacio en el Mercado Central!")
        Else

106         If UserList(UserIndex).Account.MercaderSlot > 0 Then
108             Call WriteErrorMsg(UserIndex, "¡Tienes una publicación en curso! Primero deberás quitarla para poder realizar otra.")
                Exit Sub

            End If
        
110         If Mercader_CheckingNew(UserIndex, Chars, Mercader, SaleCost, Blocked) Then
                    If SaleCost = 0 And UserList(UserIndex).Account.Premium < 3 Then
                        ' Intentó publicar 0 Personajes, no deberia.
                        Exit Sub
    
                    End If
                    
112             MercaderList(Slot).Chars = Mercader
                  MercaderList(Slot).Chars.Account = UserList(UserIndex).Account.Email
116             UserList(UserIndex).Account.MercaderSlot = Slot
                  UserList(UserIndex).Account.Gld = UserList(UserIndex).Account.Gld - SaleCost
                  
118             Call Mercader_SaveUser(UserIndex, Slot, Mercader)
120             Call Mercader_Save(Slot)
            
                If UserList(UserIndex).Account.Premium > 2 Then
                    Call Mercader_MessageDiscord(MercaderList(Slot).Chars.NameU)
                End If
                
124             Call WriteErrorMsg(UserIndex, "Publicación Exitosa. Un correo ha sido enviado con la información de la publicación.")
                  Call WriteUpdateStatusMAO(UserIndex, 1)
                  Call WriteAccountInfo(UserIndex)
                  Call Logs_Security(eLog.eGeneral, eMercader, "Cuenta: " & UserList(UserIndex).Account.Email & " con IP: " & UserList(UserIndex).IpAddress & " ha realizado una PUBLICACION. PIDE ORO: " & MercaderList(Slot).Chars.Gld)
            End If

        End If
    
        '<EhFooter>
        Exit Sub

Mercader_AddList_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.mMao.Mercader_AddList " & "at line " & Erl

        

        '</EhFooter>
End Sub

' Agregamos una Nueva Oferta a la cuenta
' PROTAGONISTA: El comprador
' Formula para aceptar la OFERTA PONER ABAJO: MAXIMO DE PJS POR CUENTA - LOS OCUPADOS ACTUALES - LOS PJS QUE OFERTO YO
Public Sub Mercader_AddOffer(ByVal UserIndex As Integer, _
                             ByRef Chars() As Byte, _
                             ByVal MercaderSlot As Byte, _
                             ByRef Mercader As tMercaderChar, _
                             ByVal Blocked As Byte)
        '<EhHeader>
        On Error GoTo Mercader_AddOffer_Err
        '</EhHeader>

        Dim tUser     As Integer

        Dim FilePath  As String

        Dim SlotOffer As Integer
    
        Dim SaleCost As Long
    
100     With MercaderList(MercaderSlot)
102         If .Chars.Gld > Mercader.Gld Or .Chars.Gld > UserList(UserIndex).Account.Gld Then

                Exit Sub ' La publicación dice que pide un mínimo de Oro y el usuario quiere ofrecer menos

            End If
        
104         If .LastOffer = MERCADER_MAX_OFFER Then
106             Call WriteErrorMsg(UserIndex, "Parece que el usuario ha recibido demasiadas ofertas y debe seleccionar algunas. Pídele que limpie su lista.")
                Exit Sub

            End If
    
108         SlotOffer = Mercader_SlotOffer(MercaderSlot, UserList(UserIndex).Account.Email)
    
110         If SlotOffer = -1 Then
112             Call WriteErrorMsg(UserIndex, "¡Ya has ofrecido a esta publicación!")
                Exit Sub

            End If
        
114         If Mercader_CheckingNew(UserIndex, Chars, Mercader, SaleCost, Blocked) Then
                  If SaleCost = 0 And Mercader.Gld = 0 Then
                        Exit Sub ' No puede ofrecer NADA a la publicación. Está hackeando el sistema
                  End If
                  
116             .Offer(SlotOffer) = Mercader
118             .OfferTime(SlotOffer) = GetTime
120             .LastOffer = .LastOffer + 1
            
122             Call WriteErrorMsg(UserIndex, "Tu oferta ha sido enviada. ¡Espera prontas noticias del creador de la publicación!")
            
124             tUser = CheckEmailLogged(MercaderList(MercaderSlot).Chars.Account)
            
126             If tUser > 0 Then
128                 If UserList(tUser).flags.UserLogged Then
130                     Call WriteConsoleMsg(tUser, "Has recibido una nueva oferta por tu publicación. Dirigete a la Boveda para decidir si quieres aceptarla o no.", FontTypeNames.FONTTYPE_INFOGREEN)
                    End If
                End If
                
                'Call WriteSendMercaderOffer(MercaderSlot, SlotOffer, MERCADER_OFFER_TIME)
                Call WriteUpdateStatusMAO(UserIndex, 1)
                
                Call Logs_Security(eLog.eGeneral, eMercader, "Cuenta: " & UserList(UserIndex).Account.Email & " con IP: " & UserList(UserIndex).IpAddress & " ha realizado una oferta a " & MercaderList(MercaderSlot).Chars.Account & ". Ofrece ORO: " & .Offer(SlotOffer).Gld)
            End If
       
        End With

        '<EhFooter>
        Exit Sub

Mercader_AddOffer_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mMao.Mercader_AddOffer " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub


' Busca un nombre de personaje en la publicación de la cuenta
Public Function Mercader_CheckUsers(ByVal Mercader As Integer, ByVal UserName As String) As Boolean
        '<EhHeader>
        On Error GoTo Mercader_CheckUsers_Err
        '</EhHeader>
        Dim A As Long

102     If Mercader > 0 Then
104         With MercaderList(Mercader)
            
106             For A = 1 To .Chars.Count
            
108                 If UCase$(.Chars.NameU(A)) = UserName Then
110                     Mercader_CheckUsers = True
                        Exit Function
                    End If
                
112             Next A
            
            End With
    
        End If
        '<EhFooter>
        Exit Function

Mercader_CheckUsers_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mMao.Mercader_CheckUsers " & _
               "at line " & Erl
        
        '</EhFooter>
End Function
' Leemos la Oferta seleccionada
' Protagonista: El que publico (vendedor)
Public Sub Mercader_AcceptOffer(ByVal UserIndex As Integer, _
                                ByVal MercaderSlot As Integer, _
                                ByVal SlotOffer As Byte)

        '<EhHeader>
        On Error GoTo Mercader_AcceptOffer_Err

        '</EhHeader>
                                
        Dim tUser             As Integer

        Dim FilePath          As String

        Dim Gld               As Long

        Dim MercaderNull      As tMercader

        Dim Temp              As String

        Dim SumaLvls          As Long

        Dim SlotsDisponibles  As Long

        Dim SlotsDisponiblesB As Long

        Dim TempChars         As Long

        Dim A                 As Long
        
        Dim NullOffer As tMercaderChar
        
100     FilePath = AccountPath & MercaderList(MercaderSlot).Offer(SlotOffer).Account & ".acc"
            
        ' La oferta caducó
        If (GetTime - MercaderList(MercaderSlot).OfferTime(SlotOffer)) >= MERCADER_OFFER_TIME Then
            MercaderList(MercaderSlot).Offer(SlotOffer) = NullOffer
            MercaderList(MercaderSlot).OfferTime(SlotOffer) = 0
            Call WriteErrorMsg(UserIndex, "¡La oferta caducó!")
            Exit Sub
        End If
        
          ' Caso hipotetico
          ' Personaje nro1 = Publica 5 personajes. Tiene un total de 10
          ' Personaje nro2= Ofrece 5 personajes. Tiene un total de 10
          ' Personajenro1 = 10 - 5 +5 = 10
          SlotsDisponibles = (UserList(UserIndex).Account.CharsAmount - MercaderList(MercaderSlot).Chars.Count + MercaderList(MercaderSlot).Offer(SlotOffer).Count)
          
104     If SlotsDisponibles > ACCOUNT_MAX_CHARS Then
106         Call WriteErrorMsg(UserIndex, "No tienes espacio para recibir la Oferta.")
            Exit Sub

        End If
    
108     tUser = CheckEmailLogged(MercaderList(MercaderSlot).Offer(SlotOffer).Account)
    
        ' El pibe de la oferta
110     If tUser > 0 Then
              SlotsDisponibles = (UserList(tUser).Account.CharsAmount - MercaderList(MercaderSlot).Offer(SlotOffer).Count) + MercaderList(MercaderSlot).Chars.Count
              
114         Gld = UserList(tUser).Account.Gld
        Else
118         Gld = val(GetVar(FilePath, "INIT", "GLD"))
120         TempChars = val(GetVar(FilePath, "INIT", "CHARSAMOUNT"))
            
            
            SlotsDisponibles = (TempChars - MercaderList(MercaderSlot).Offer(SlotOffer).Count) + MercaderList(MercaderSlot).Chars.Count
          
        End If
    
124     If SlotsDisponibles > ACCOUNT_MAX_CHARS Then ' MercaderList(MercaderSlot).Chars.Count
126         Call WriteErrorMsg(UserIndex, "Parece que la persona no tiene espacio para recibir nuevos personajes.")
            Exit Sub

        End If
    
        ' El pibe de la Oferta no tiene el oro que tenia en un principio
128     If MercaderList(MercaderSlot).Offer(SlotOffer).Gld > Gld Then
130         Call WriteErrorMsg(UserIndex, "El usuario te ha ofrecido oro y luego lo ha utilizado por lo cual no tiene para pagarte. ¡Lamentamos lo sucedido!")
        
132         If tUser > 0 Then
134             If UserList(tUser).flags.UserLogged = True Then
136                 Call WriteConsoleMsg(tUser, "Parece ser que no tienes el oro suficiente y la oferta enviada recientemente no puede ser aceptada.", FontTypeNames.FONTTYPE_INFORED)

                End If

            End If
        
            Exit Sub

        End If

138     If MercaderList(MercaderSlot).Offer(SlotOffer).Count > 0 Then
140         If Not Mercader_CheckingChar_Offer(MercaderList(MercaderSlot).Offer(SlotOffer).NameU, MercaderList(MercaderSlot).Offer(SlotOffer).Account) Then
142             Call WriteErrorMsg(UserIndex, "La solicitud ha expirado por alguna razón. Comprueba que la persona siga disponiendo de la Oferta")

                Exit Sub    ' El pibe de la oferta no tiene mas los personajes

            End If

        End If
    
        ' Quitamos el Oro de la cuenta y lo agregamos a la otra
144     If MercaderList(MercaderSlot).Offer(SlotOffer).Gld > 0 Then
146         If tUser > 0 Then
148             UserList(tUser).Account.Gld = UserList(tUser).Account.Gld - MercaderList(MercaderSlot).Offer(SlotOffer).Gld
            Else
152             Call WriteVar(FilePath, "INIT", "GLD", CStr(Gld - MercaderList(MercaderSlot).Offer(SlotOffer).Gld))
    
            End If
        
154         UserList(UserIndex).Account.Gld = UserList(UserIndex).Account.Gld + MercaderList(MercaderSlot).Offer(SlotOffer).Gld
156         Call WriteErrorMsg(UserIndex, "Se ha depositado en tu cuenta algunas Monedas de Oro por una Venta que acaba de ser confirmada.")

        End If
    
        ' Le quitamos los Pjs al flaco de la venta. Está online
158     Call Mercader_UpdateCharsAccount(UserIndex, MercaderList(MercaderSlot).Chars.NameU, True)
    
        ' Le metemos los pjs de la oferta en caso de que haya y no sea solo oro
160     If MercaderList(MercaderSlot).Offer(SlotOffer).Count > 0 Then

            ' Quitamos los personajes de la oferta de la cuenta
164         If tUser > 0 Then
166             Call Mercader_UpdateCharsAccount(tUser, MercaderList(MercaderSlot).Offer(SlotOffer).NameU, True)
              
            Else
170             Call Mercader_RemoveCharsAccount_Offline(MercaderList(MercaderSlot).Offer(SlotOffer).Account, MercaderList(MercaderSlot).Offer(SlotOffer).NameU, True)

            End If
             
            
            Call Mercader_UpdateCharsAccount(UserIndex, MercaderList(MercaderSlot).Offer(SlotOffer).NameU, False)
        
        End If
        
        ' Agregamos los personajes que compró
        If tUser > 0 Then
            Call Mercader_UpdateCharsAccount(tUser, MercaderList(MercaderSlot).Chars.NameU, False)
        Else
            Call Mercader_RemoveCharsAccount_Offline(MercaderList(MercaderSlot).Offer(SlotOffer).Account, MercaderList(MercaderSlot).Chars.NameU, False)
        End If
    
        ' Quitamos la publicacion necesaria en la oferta
174     If MercaderList(MercaderSlot).Offer(SlotOffer).Count > 0 Then

            For A = 1 To MercaderList(MercaderSlot).Offer(SlotOffer).Count
                Call Mercader_SearchPublications_User(MercaderList(MercaderSlot).Offer(SlotOffer).Account, MercaderList(MercaderSlot).Offer(SlotOffer).NameU(A))
            Next A

        End If
        
        Call Logs_Security(eLog.eGeneral, eMercader, "Cuenta: " & UserList(UserIndex).Account.Email & " con IP: " & UserList(UserIndex).IpAddress & " ha confirmado la oferta de " & MercaderList(MercaderSlot).Offer(SlotOffer).Count & ". por ORO: " & MercaderList(MercaderSlot).Offer(SlotOffer).Gld)
         
        ' Quitamos la publicación de la VENTA
        MercaderList(MercaderSlot) = MercaderNull
        
        'Call WriteSendMercaderOffer(MercaderSlot, SlotOffer, 0)
        Call Mercader_Remove(MercaderSlot, UserList(UserIndex).Account.Email)
        
        Call WriteLoggedAccount(UserIndex, UserList(UserIndex).Account.Chars)
        Call mAccount.SaveDataAccount(UserIndex, UserList(UserIndex).Account.Email, UserList(UserIndex).IpAddress)
        Call Mercader_Save(MercaderSlot)
        
        '<EhFooter>
        Exit Sub

Mercader_AcceptOffer_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.mMao.Mercader_AcceptOffer " & "at line " & Erl

        

        '</EhFooter>
End Sub

' Removemos los personajes de la Cuenta
Public Sub Mercader_UpdateCharsAccount(ByVal UserIndex As Integer, _
                                       ByRef Chars() As String, _
                                       ByVal Killed As Boolean)

        '<EhHeader>
        On Error GoTo Mercader_UpdateCharsAccount_Err

        '</EhHeader>
        Dim A           As Long, B As Long

        Dim NullChar    As tAccountChar

        Dim tUser       As Integer
        
        Dim CharsAmount As Byte
        
        
        CharsAmount = UserList(UserIndex).Account.CharsAmount
        ' Mientras no haya completado la cantidad de chars a poner
              
104     For B = LBound(Chars) To UBound(Chars)
            For A = 1 To ACCOUNT_MAX_CHARS
106             tUser = NameIndex(Chars(B))
            
108             If tUser > 0 Then
110                 Call WriteDisconnect(tUser)
112                 Call FlushBuffer(tUser)
114                 Call CloseSocket(tUser)

                End If
            
116             If Killed Then
118                 If StrComp(UCase$(UserList(UserIndex).Account.Chars(A).Name), UCase$(Chars(B))) = 0 Then
120                     UserList(UserIndex).Account.Chars(A) = NullChar
                          Call WriteVar(AccountPath & UserList(UserIndex).Account.Email & ".acc", "CHARS", CStr(A), vbNullString)
                          Call WriteVar(CharPath & Chars(B) & ".chr", "FLAGS", "BLOCKED", "0")
                          Call WriteVar(CharPath & Chars(B) & ".chr", "INIT", "ACCOUNTNAME", vbNullString)
                          Call WriteVar(CharPath & Chars(B) & ".chr", "INIT", "ACCOUNTSLOT", "0")
                        CharsAmount = CharsAmount - 1
                        Exit For

                    End If

                Else

124                 If Len(UserList(UserIndex).Account.Chars(A).Name) = 0 Then
126                     UserList(UserIndex).Account.Chars(A).Name = Chars(B)
                        Call Login_Char_LoadInfo(UserIndex, A, Chars(B))
                        Call WriteVar(AccountPath & UserList(UserIndex).Account.Email & ".acc", "CHARS", CStr(A), Chars(B))
                        Call WriteVar(CharPath & Chars(B) & ".chr", "INIT", "ACCOUNTNAME", UserList(UserIndex).Account.Email)
                        Call WriteVar(CharPath & Chars(B) & ".chr", "INIT", "ACCOUNTSLOT", CStr(A))
                        CharsAmount = CharsAmount + 1
                        Exit For

                    End If
                
                End If
                
130         Next A
        
        

132 Next B

        UserList(UserIndex).Account.CharsAmount = CharsAmount
    
    '<EhFooter>
    Exit Sub

Mercader_UpdateCharsAccount_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mMao.Mercader_UpdateCharsAccount " & "at line " & Erl

    

    '</EhFooter>
End Sub

' Removemos los personajes de la Cuenta Offline
Public Sub Mercader_RemoveCharsAccount_Offline(ByVal Account As String, _
                                               ByRef Chars() As String, _
                                               ByVal Killed As Boolean)

        '<EhHeader>
        On Error GoTo Mercader_RemoveCharsAccount_Offline_Err

        '</EhHeader>
        Dim A           As Long, B As Long

        Dim FilePath    As String

        Dim CharsAmount As Byte
        
100     FilePath = AccountPath & Account & ".acc"
          
        CharsAmount = val(GetVar(FilePath, "INIT", "CHARSAMOUNT"))
          
102

104     For B = LBound(Chars) To UBound(Chars)
            For A = 1 To ACCOUNT_MAX_CHARS
                  
                If Killed Then
106                 If StrComp(UCase$(GetVar(FilePath, "CHARS", A)), Chars(B)) = 0 Then
108                     Call WriteVar(FilePath, "CHARS", A, vbNullString)
                          Call WriteVar(CharPath & Chars(B) & ".chr", "FLAGS", "BLOCKED", "0")
                        CharsAmount = CharsAmount - 1
                        Exit For

                    End If
                
                Else
                                
                    If GetVar(FilePath, "CHARS", A) = vbNullString Then
                        Call WriteVar(FilePath, "CHARS", A, Chars(B))
                        Call WriteVar(CharPath & Chars(B) & ".chr", "INIT", "ACCOUNTNAME", Account)
                        Call WriteVar(CharPath & Chars(B) & ".chr", "INIT", "ACCOUNTSLOT", CStr(A))
                        CharsAmount = CharsAmount + 1
                        Exit For

                    End If

                End If

110         Next A
112     Next B

        Call WriteVar(FilePath, "INIT", "CHARSAMOUNT", CStr(CharsAmount))
    
        '<EhFooter>
        Exit Sub

Mercader_RemoveCharsAccount_Offline_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.mMao.Mercader_RemoveCharsAccount_Offline " & "at line " & Erl

        

        '</EhFooter>
End Sub

' Limpiamos un Slot del Mercado
Public Sub Mercader_Remove(ByVal Slot As Integer, ByVal Account As String)

        '<EhHeader>
        On Error GoTo Mercader_Remove_Err

        '</EhHeader>
        Dim MercaderNull As tMercader

        Dim Temp         As String

        Dim A            As Long
        
        Dim tUser As Integer
        
        For A = 1 To MercaderList(Slot).Chars.Count

            If val(GetVar(CharPath & UCase$(MercaderList(Slot).Chars.NameU(A)) & ".chr", "FLAGS", "BLOCKED")) > 0 Then
                Call WriteVar(CharPath & UCase$(MercaderList(Slot).Chars.NameU(A)) & ".chr", "FLAGS", "BLOCKED", "0")

            End If

        Next A
            
        tUser = CheckEmailLogged(Account)
        
        If tUser > 0 Then
            UserList(tUser).Account.MercaderSlot = 0
            Call WriteVar(AccountPath & Account & ".acc", "SALE", "LAST", "0")
        Else
            Call WriteVar(AccountPath & MercaderList(Slot).Chars.Account & ".acc", "SALE", "LAST", "0")

        End If
          
100     MercaderList(Slot) = MercaderNull
102     Call Mercader_Save(Slot)
        '<EhFooter>
        Exit Sub

Mercader_Remove_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.mMao.Mercader_Remove " & "at line " & Erl

        

        '</EhFooter>
End Sub

' Buscamos las publicaciones donde el usuario tenga pjs y las sacamos.
Public Sub Mercader_SearchPublications_User(ByVal Account As String, _
                                                                   ByVal User As String, _
                                                                   Optional ByVal BanAccount As Boolean = False)
        '<EhHeader>
        On Error GoTo Mercader_SearchPublications_User_Err
        '</EhHeader>

        Dim A As Long, B As Long, C As Long
        Dim SlotMercader As Integer
        Dim tUser As Integer
        
        If User <> vbNullString Then
            tUser = NameIndex(User)
        End If
        
        If tUser > 0 Then
            SlotMercader = UserList(tUser).Account.MercaderSlot
        Else
            SlotMercader = val(GetVar(AccountPath & Account & ".acc", "SALE", "LAST"))
        End If
        
        If SlotMercader = 0 Then Exit Sub
        If MercaderList(SlotMercader).Chars.Count = 0 Then Exit Sub ' Ya fue removida, debido a que otro personaje se involucraba

        ' Control de Baneo de Cuenta entera
        If BanAccount Then
            Mercader_Remove SlotMercader, Account
            Exit Sub
        End If
        
100     With MercaderList(SlotMercader)
    
102         For A = 1 To .Chars.Count
104             If StrComp(.Chars.NameU(A), User) = 0 Then
106                 Mercader_Remove SlotMercader, Account
                      
                      Exit Sub
                End If
108         Next A
        
        End With
        '<EhFooter>
        Exit Sub

Mercader_SearchPublications_User_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mMao.Mercader_SearchPublications_User " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

' Reiniciamos un Slot de Oferta
' Genera la lista de personajes separados con "-"
Public Function Mercader_Generate_Text_Chars(ByRef Users() As String) As String
        '<EhHeader>
        On Error GoTo Mercader_Generate_Text_Chars_Err
        '</EhHeader>
        Dim A As Long
        Dim Temp As String

100     For A = LBound(Users) To UBound(Users)
102         If Users(A) <> vbNullString Then
104             Temp = Temp & Users(A) & "-"
            End If
106     Next A
    
108     Temp = Left$(Temp, Len(Temp) - 1)
    
110     Mercader_Generate_Text_Chars = Temp
        '<EhFooter>
        Exit Function

Mercader_Generate_Text_Chars_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mMao.Mercader_Generate_Text_Chars " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

' Genera un Slot Libre para la Oferta
Private Function Mercader_SlotOffer(ByVal MercaderSlot As Integer, ByVal Account As String) As Integer

    '<EhHeader>
    On Error GoTo Mercader_SlotOffer_Err

    '</EhHeader>
    Dim A As Long
    Dim Temp As Integer
    
    For A = 1 To MERCADER_MAX_OFFER

        With MercaderList(MercaderSlot).Offer(A)
            
            If .Account = vbNullString And Temp = 0 Then
                Temp = A
            End If
            
            If StrComp(.Account, Account) = 0 Then
                Mercader_SlotOffer = -1
                Exit Function
            End If
        End With

    Next A
   
   Mercader_SlotOffer = Temp
    '<EhFooter>
    Exit Function

Mercader_SlotOffer_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mMao.Mercader_SlotOffer " & "at line " & Erl

    

    '</EhFooter>
End Function

' # Enviamos un mensaje al canal de discord [SOLO TIERS 3]
Public Sub Mercader_MessageDiscord(ByRef Chars() As String)
    
    On Error GoTo ErrHandler
    
    Dim A As Long
    Dim Users As String
    For A = LBound(Chars) To UBound(Chars)

        If Chars(A) <> vbNullString Then
            Users = Users & Mercader_GenerateText(Chars(A), True)
            
            If A < UBound(Chars) Then
                Users = Users & " | "
            End If
        End If
    Next A
    
     WriteMessageDiscord CHANNEL_MERCADER, "**Nueva publicación:** " & Users
    
    Exit Sub
    
ErrHandler:
End Sub


Private Function Mercader_SetChar(ByVal SlotChar As Byte, _
                                  ByVal Char As String, _
                                  ByVal Blocked As Byte, _
                                  Optional ByVal IsOffer As Boolean = False) As tMercaderCharInfo
        '<EhHeader>
        On Error GoTo Mercader_SetChar_Err
        '</EhHeader>
    
        Dim Manager As clsIniManager

100     Set Manager = New clsIniManager
        
102     Dim Charfile     As String: Charfile = CharPath & Char & ".chr"
        Dim Temp As tMercaderCharInfo
        Dim promedio As Long
        Dim A As Long
        Dim ln As String
        
        
104     Manager.Initialize Charfile

106     With Temp
108         .Class = val(Manager.GetValue("INIT", "CLASE"))
110         .Raze = val(Manager.GetValue("INIT", "RAZA"))
112         .Elv = val(Manager.GetValue("STATS", "ELV"))
114         .Exp = val(Manager.GetValue("STATS", "EXP"))
116         .Elu = val(Manager.GetValue("STATS", "ELU"))
118         .Hp = val(Manager.GetValue("STATS", "MAXHP"))
120         .Constitucion = val(Manager.GetValue("ATRIBUTOS", "AT" & eAtributos.Constitucion))
        
        
122         .Body = val(Manager.GetValue("INIT", "BODY"))
124         .Head = val(Manager.GetValue("INIT", "HEAD"))
126         .Weapon = val(Manager.GetValue("INIT", "ARMA"))
128         .Helm = val(Manager.GetValue("INIT", "CASCO"))
130         .Shield = val(Manager.GetValue("INIT", "ESCUDO"))
        
132         .Gld = val(Manager.GetValue("STATS", "GLD"))
134         .GuildIndex = val(Manager.GetValue("GUILD", "GUILDINDEX"))
        
136         .Faction = val(Manager.GetValue("FACTION", "STATUS"))
138         .FactionRange = val(Manager.GetValue("FACTION", "RANGE"))
        
            Dim Tempito As Long
        
140         If .Faction = 0 Then
142             Tempito = val(Manager.GetValue("REP", "PROMEDIO"))
            
144             If Tempito < 0 Then
146                 .Faction = 4
                Else
148                 .Faction = 3
                End If
            
            End If
        
150         .FragsCri = val(Manager.GetValue("FACTION", "FRAGSCRI"))
152         .FragsCiu = val(Manager.GetValue("FACTION", "FRAGSCIU"))
                
                
                
              ReDim .Bank(1 To MAX_BANCOINVENTORY_SLOTS) As tMercaderObj
              
154         For A = 1 To MAX_BANCOINVENTORY_SLOTS
156             ln = Manager.GetValue("BANCOINVENTORY", "OBJ" & A)
158             .Bank(A).ObjIndex = CInt(ReadField(1, ln, 45))
160             .Bank(A).Amount = CInt(ReadField(2, ln, 45))
162         Next A

              ReDim .Object(1 To MAX_INVENTORY_SLOTS) As tMercaderObj
164         For A = 1 To MAX_INVENTORY_SLOTS
166             ln = Manager.GetValue("INVENTORY", "OBJ" & A)
168             .Object(A).ObjIndex = val(ReadField(1, ln, 45))
170             .Object(A).Amount = val(ReadField(2, ln, 45))
172         Next A
            
                ReDim .Spells(1 To 35) As Byte
                
                For A = 1 To 35
                    .Spells(A) = val(Manager.GetValue("HECHIZOS", "H" & A))
                Next A
                
                ReDim .Skills(1 To NUMSKILLS) As Byte
                
                For A = 1 To NUMSKILLS
                    .Skills(A) = val(Manager.GetValue("SKILLS", "SK" & A))
                Next A
                
        End With
    
174     If Blocked = 1 Then
176         Call Manager.ChangeValue("FLAGS", "BLOCKED", "1")

              If IsOffer Then
                    Call Manager.ChangeValue("FLAGS", "OFFERTIME", Format$(Now, "dd/mm/yyyy hh:mm:ss"))
              End If
              
178         Call Manager.DumpFile(Charfile)
        End If
    
180     Mercader_SetChar = Temp
182     Set Manager = Nothing
        '<EhFooter>
        Exit Function

Mercader_SetChar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mMao.Mercader_SetChar " & _
               "at line " & Erl & " NICK: " & Char
        
        '</EhFooter>
End Function

Public Function Mercader_GenerateText(ByVal Char As String, Optional ByVal IsDiscord As Boolean = False) As String
        '<EhHeader>
        On Error GoTo Mercader_GenerateText_Err
        '</EhHeader>
    
        Dim Reader As clsIniManager

100     Set Reader = New clsIniManager
    
        Dim Class As eClass

        Dim Raze         As eRaza

        Dim Elv          As Byte

        Dim Exp          As Long

        Dim Elu          As Long

        Dim Ups          As Single

        Dim Penas        As Byte

        Dim Hp           As Integer

        Dim Constitution As Byte
    
102     Dim Charfile     As String: Charfile = CharPath & Char & ".chr"

        Dim Text         As String
    
        Dim TextUps As String
    
104     Reader.Initialize Charfile
    
106     Class = val(Reader.GetValue("INIT", "CLASE"))
108     Raze = val(Reader.GetValue("INIT", "RAZA"))
110     Elv = val(Reader.GetValue("STATS", "ELV"))
112     Exp = val(Reader.GetValue("STATS", "EXP"))
114     Elu = val(Reader.GetValue("STATS", "ELU"))
116     Hp = val(Reader.GetValue("STATS", "MAXHP"))
118     Constitution = val(Reader.GetValue("ATRIBUTOS", "AT" & eAtributos.Constitucion))
120     Ups = Hp - getVidaIdeal(Elv, Class, Constitution)
    
122     If Ups > 0 Then
124         TextUps = "+" & Ups
126     ElseIf Ups < 0 Then
128         TextUps = Ups
130     ElseIf Ups = 0 Then
132         TextUps = "Prom"
        End If
        
        If IsDiscord Then
            Text = "**" & UCase$(Char) & "**." & ListaClases(Class) & "." & ListaRazas(Raze) & "." & Elv & "**(" & TextUps & ")**"
        Else
            Text = UCase$(Char) & "." & ListaClases(Class) & "." & ListaRazas(Raze) & "." & Elv & "(" & TextUps & ")"
        End If
        
134
136     If Elv <> STAT_MAXELV Then
138         Text = Text & "(" & Round(CDbl(Exp) * CDbl(100) / CDbl(Elu), 2) & "%)"
        End If
    
140     Mercader_GenerateText = Text
    
142     Set Reader = Nothing
    
        Exit Function

Mercader_GenerateText_Err:
        Set Reader = Nothing
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mMao.Mercader_GenerateText " & _
               "at line " & Erl
        
        '</EhFooter>
End Function
