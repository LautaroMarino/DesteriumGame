Attribute VB_Name = "mAccount_Mercader"
Option Explicit

' Subasta de Objetos
' Subasta de Personajes
' Compra-Venta de Personajes

' # CONSTANTES DE MERCADO #

' Compra-Venta de Personajes
Public Const MAX_MERCADER_OFFER  As Long = 20

Public Const MAX_MERCADER_GLD    As Long = 500000000

Public Const MAX_MERCADER_ELDHIR As Long = 10000

Public Const MAX_MERCADER_CHARS  As Byte = 5

Public Const SALECHAR_ARCHIVE    As String = "MERCADER.DAT"

Public Const MAX_MERCADER_SLOT   As Integer = 200

Public Const MIN_MERCADER_LEVEL  As Integer = 25

Public Type tAccountMercader_CharSale_Info

    Name As String
    Class As Byte
    Raze As Byte
    Elv As Byte
    Exp As Long
    Elu As Long
    
    Hp As Integer
    Man As Integer
    Ups As Integer
    Promedio As Long
    
    Body As Integer
    Head As Integer
    Helm As Integer
    Genero As Byte
    Shield As Integer
    Weapon As Integer
    
    Gld As Long
    Eldhir As Long
    
    Bronce As Byte
    Plata As Byte
    Oro As Byte
    Premium As Byte
    
    Object(1 To MAX_INVENTORY_SLOTS) As UserOBJ
    ObjectBank(1 To MAX_BANCOINVENTORY_SLOTS) As UserOBJ
    Spells(1 To 35) As String
    Meditations(1 To MAX_MEDITATION) As Byte
    
    Ban As Byte
    Penas As Byte
    PenasText() As String
    
    FragsCiu As Integer
    FragsCri As Integer
    FragsOther As Integer
    FragsNpc As Integer
    
    Faction As Byte
    FactionEx As Byte
    
    GuildName As String
    GuildRange As String
    
    Points As Long
    Retos1Jugados As Long
    Retos1Ganados As Long
    
    Retos2Jugados As Long
    Retos2Ganados As Long
    
    Retos3Jugados As Long
    Retos3Ganados As Long
    
    EventosJugados As Long
    EventosGanados As Long
    
    Bandido As Long
    Asesino As Long
    Nobleza As Long
    
    
    NumQuests As Byte
    Quests() As Byte
End Type

Public Type tAccountMercader_CharSale_Offer

    Account As String
    Char(MAX_MERCADER_CHARS - 1) As String
    CharText(MAX_MERCADER_CHARS - 1) As String
    FastText(MAX_MERCADER_CHARS - 1) As String
    CharAmount As Integer
    Gld As Long
    Eldhir As Long
    LastTime As Long

End Type

Public Type tAccountMercader_CharSale

    Account As String
    Char(MAX_MERCADER_CHARS - 1) As String
    CharText(MAX_MERCADER_CHARS - 1) As String
    
    CharInfo(MAX_MERCADER_CHARS - 1) As tAccountMercader_CharSale_Info
    
    CharAmount As Integer
    Gld As Long
    Eldhir As Long
    Blocked As Byte
    Offers(1 To MERCADER_MAX_OFFER) As tAccountMercader_CharSale_Offer

End Type

Private Mercader_NullOffer As tAccountMercader_CharSale_Offer
Public MercaderDS(1 To MAX_MERCADER_SLOT) As tAccountMercader_CharSale


' # COMPRA VENTA DE PERSONAJES
Public Sub SaleChar_CreateFile()

    On Error GoTo ErrHandler

    Dim intFile As Integer

    Dim A       As Long, B As Long
    
    intFile = FreeFile
    
    Open DatPath & SALECHAR_ARCHIVE For Output As #intFile
    
    For A = 1 To MAX_MERCADER_SLOT
        Print #intFile, "[" & A & "]"
        Print #intFile, "CHAR="
        Print #intFile, "GLD=0"
        Print #intFile, "ELDHIR=0"
        Print #intFile, "BLOCKED=0"
        Print #intFile, "ACCOUNT="
    Next A

    Close #intFile
    
    Exit Sub

ErrHandler:
    Close #intFile
    Call LogError("SaleChar_CreateFile ERROR")
End Sub

Public Sub SaleChar_Load()

    On Error GoTo ErrHandler

    Dim A    As Long, B As Long, C As Long, List() As String

    Dim Read As clsIniManager
    
    If Not FileExist(DatPath & SALECHAR_ARCHIVE) Then
        Call SaleChar_CreateFile
    End If
    
    Set Read = New clsIniManager
    
    Call Read.Initialize(DatPath & SALECHAR_ARCHIVE)
    
    For A = 1 To MAX_MERCADER_SLOT

        With MercaderDS(A)
            .Account = Read.GetValue(A, "ACCOUNT")
            List = Split(Read.GetValue(A, "CHAR"), "-")
            
            For B = LBound(List) To UBound(List)
                .Char(B) = List(B)
                .CharInfo(B).Name = List(B)
                
                If .Char(B) <> vbNullString Then
                    If PersonajeExiste(.Char(B)) Then
                        '.CharText(B) = SaleChar_Text(.Char(B))
                       ' .CharInfo(B) = SaleChar_ListInfo_Load_Offline(.Char(B))
                    End If
                End If
                
                If .Char(B) <> vbNullString Then .CharAmount = .CharAmount + 1
            Next B
            
            .Gld = val(Read.GetValue(A, "GLD"))
            .Eldhir = val(Read.GetValue(A, "ELDHIR"))
            .Blocked = val(Read.GetValue(A, "BLOCKED"))
            

        End With

    Next A
    
    Set Read = Nothing
    
    Exit Sub

ErrHandler:
    Set Read = Nothing
    Call LogError("SaleChar_Load ERROR")
End Sub

Public Sub SaleChar_Save()

    On Error GoTo ErrHandler

    Dim A       As Long, B As Long, C As Long, TempChar As String

    Dim Manager As clsIniManager
    
    Set Manager = New clsIniManager
    
    Call Manager.Initialize(DatPath & SALECHAR_ARCHIVE)
    
    For A = 1 To MAX_MERCADER_SLOT

        With MercaderDS(A)

            For C = LBound(.Char) To UBound(.Char)
                TempChar = TempChar & .Char(C) & "-"
            Next C
                
            If Len(TempChar) > 0 Then
                TempChar = Left$(TempChar, Len(TempChar) - 1)
            End If
             
            Call Manager.ChangeValue(A, "CHAR", TempChar)
            Call Manager.ChangeValue(A, "ACCOUNT", CStr(.Account))
            Call Manager.ChangeValue(A, "GLD", CStr(.Gld))
            Call Manager.ChangeValue(A, "ELDHIR", CStr(.Eldhir))
            Call Manager.ChangeValue(A, "BLOCKED", CStr(.Blocked))

        End With

    Next A
    
    Manager.DumpFile (DatPath & SALECHAR_ARCHIVE)
    Set Manager = Nothing
    
    Exit Sub

ErrHandler:
    Set Manager = Nothing
    Call LogError("SaleChar_CreateFile ERROR")
End Sub

Public Sub SaleChar_SaveSlot(ByVal Slot As Integer)

    On Error GoTo ErrHandler

    Dim B       As Long, C As Long, TempChar As String

    Dim Manager As clsIniManager
    
    Set Manager = New clsIniManager
    
    Call Manager.Initialize(DatPath & SALECHAR_ARCHIVE)
    
    With MercaderDS(Slot)

        If .Char(0) <> vbNullString Then

            For C = LBound(.Char) To UBound(.Char)
                TempChar = TempChar & .Char(C) & "-"
                
                '.CharInfo(C) = SaleChar_ListInfo_Load_Offline(Slot, C)
            Next C
                
            If Len(TempChar) > 0 Then
                TempChar = Left$(TempChar, Len(TempChar) - 1)
            End If
            
        End If
          
        Call Manager.ChangeValue(Slot, "CHAR", TempChar)
        Call Manager.ChangeValue(Slot, "ACCOUNT", CStr(.Account))
        Call Manager.ChangeValue(Slot, "GLD", CStr(.Gld))
        Call Manager.ChangeValue(Slot, "ELDHIR", CStr(.Eldhir))
        Call Manager.ChangeValue(Slot, "BLOCKED", CStr(.Blocked))

    End With
    
    Manager.DumpFile (DatPath & SALECHAR_ARCHIVE)
    Set Manager = Nothing
    
    Exit Sub

ErrHandler:
    Set Manager = Nothing
    Call LogError("SaleChar_CreateFile ERROR")
End Sub

' MERCADER :: Busca un slot libre para una nueva publicación
Private Function SaleChar_FreeMercader() As Integer

    Dim A As Long
    
    For A = 1 To MAX_MERCADER_SLOT

        If MercaderDS(A).Account = vbNullString Then
            SaleChar_FreeMercader = A

            Exit Function

        End If

    Next A

End Function

' MERCADER :: Busca un slot repetido en la oferta dentro de una publicación.
Private Function SaleChar_Search_Offer(ByVal SlotMercader As Integer, _
                                             ByVal Account As String) As Boolean

    Dim A As Long
    
    For A = 1 To MERCADER_MAX_OFFER

        If StrComp(MercaderDS(SlotMercader).Offers(A).Account, Account) = 0 Then
            SaleChar_Search_Offer = True

            Exit Function

        End If

    Next A

End Function

' MERCADER :: Busca un slot libre para una nueva oferta dentro de una publicación.
Private Function SaleChar_FreeMercader_Offer(ByVal SlotMercader As Integer, _
                                             ByVal Account As String) As Integer

    Dim A As Long
    
    For A = 1 To MERCADER_MAX_OFFER

        If MercaderDS(SlotMercader).Offers(A).Account = vbNullString Then
            SaleChar_FreeMercader_Offer = A

            Exit Function

        End If

    Next A

End Function

' MERCADER :: Setea el Slot del mercader en la cuenta de la persona.
Private Sub SaleChar_SetSlot(ByVal UserIndex As Integer, _
   ByVal Slot As Integer)
                             
    UserList(UserIndex).Account.SaleCharSlot = Slot
    Call WriteVar(AccountPath & UserList(UserIndex).Account.Email & ACCOUNT_FORMAT, "INIT", "SALECHARSLOT", CStr(Slot))
End Sub

' MERCADER :: Comprueba si el personaje que desea ingresar no está bloqueado por el mao.
Public Function SaleChar_CheckChars(ByVal UserIndex As Integer, _
                                    ByVal Char As String) As Boolean
    
    Dim Slot As Byte

    Dim A    As Long

    Slot = UserList(UserIndex).Account.SaleCharSlot
    
    If Slot > 0 Then
        If MercaderDS(Slot).Blocked > 0 Then

            With MercaderDS(Slot)

                For A = LBound(.Char) To UBound(.Char)

                    If StrComp(.Char(A), Char) = 0 Then

                        Exit Function

                    End If

                Next A
        
            End With

        End If
    End If
    
    SaleChar_CheckChars = True
End Function

' MERCADER :: Comprueba si una cuenta esta posteada
Public Sub SaleChar_CheckBan(ByVal Char As String)
    
    Dim Slot     As Byte

    Dim A        As Long

    Dim FilePath As String
    
    Slot = SaleChar_CheckSlot(GetVar(CharPath & Char & ".chr", "INIT", "ACCOUNTNAME"))
    
    If Slot > 0 Then

        For A = 1 To MAX_MERCADER_CHARS

            If StrComp(UCase$(MercaderDS(Slot).Char(A)), Char) = 0 Then
            
                Call SaleChar_Reset(Slot)
                Call SaleChar_SaveSlot(Slot)
                
                Exit Sub

            End If

        Next A

    End If

End Sub

' MERCADER :: Comprueba si una cuenta esta posteada
Public Function SaleChar_CheckSlot(ByVal Account As String) As Byte
    
    Dim Slot     As Byte

    Dim A        As Long

    Dim FilePath As String
    
    FilePath = AccountPath & Account & ACCOUNT_FORMAT
    
    Slot = val(GetVar(FilePath, "INIT", "SALECHARSLOT"))
    
    SaleChar_CheckSlot = Slot
End Function

' MERCADER :: Comprobaciones para publicar/ofertar personajes y/o oro/eldhires. (Anti Hacking principal)
Public Function SaleChar_CharValid(ByVal UserIndex As Integer, _
                                    ByRef Char() As String, _
                                    ByVal Gld As Long, _
                                    ByVal Eldhir As Long) As Boolean
    
    Dim A           As Long

    Dim TempAccount As String
    
    If Gld < 0 Or Gld > MAX_MERCADER_GLD Then

        ' Anti Hacking:: Oro inválido.
        Exit Function

    End If
    
    If Eldhir < 0 Or Eldhir > MAX_MERCADER_ELDHIR Then

        ' Anti Hacking:: Eldhir inválido.
        Exit Function

    End If
    
    If UBound(Char) > (MAX_MERCADER_CHARS - 1) Then

        ' Anti Hacking:: No puede publicar más de 5 personajes
        Exit Function

    End If
    
    If UBound(Char) = -1 And Gld = 0 And Eldhir = 0 Then

        ' Anti Hacking:: No puede hacer algo nulo
        Exit Function

    End If
    
    If Not SaleChar_CharValid_Chars(UserIndex, UserList(UserIndex).Account.Email, Char) Then
        
        Exit Function

    End If
    
    SaleChar_CharValid = True
    
End Function

' MERCADER :: Comprobaciones para publicar/ofertar/aceptar personajes. (Anti Hacking secundario)
Private Function SaleChar_CharValid_Chars(ByVal UserIndex As Integer, _
                                          ByVal Account As String, _
                                          ByRef Char() As String) As Boolean
    
    Dim A           As Long

    Dim TempAccount As String
    
    For A = LBound(Char) To UBound(Char)

        If Len(Char(A)) > 0 Then
            If Not PersonajeExiste(Char(A)) Then
                ' Anti hacking:: El personaje no existe.
                Call WriteErrorMsg(UserIndex, "¡El Personaje " & Char(A) & " ya no existe!")

                Exit Function

            End If
    
            TempAccount = LCase$(GetVar(CharPath & UCase$(Char(A)) & ".chr", "INIT", "ACCOUNTNAME"))
            
            If Not StrComp(Account, TempAccount) = 0 Then
                ' Anti Hacking:: El personaje no pertenece a la cuenta.
                Call WriteErrorMsg(UserIndex, "El Personaje " & Char(A) & " no pertenece a la cuenta. ¡Podría haberse vendido!")

                Exit Function

            End If
            
                        
            If (CheckUserLogged(UCase$(Char(A)))) Then
                Call WriteErrorMsg(UserIndex, "El Personaje " & Char(A) & " se encuentra conectado.")
                ' Anti Hacking:: El personaje está ONLINE

                Exit Function

            End If
            
            If val(GetVar(CharPath & UCase$(Char(A)) & ".chr", "FLAGS", "BAN")) > 0 Then
                Call WriteErrorMsg(UserIndex, "El Personaje " & Char(A) & " se encuentra baneado.")

                Exit Function

            End If
            
            If val(GetVar(CharPath & UCase$(Char(A)) & ".chr", "STATS", "ELV")) < 25 Then
                Call WriteErrorMsg(UserIndex, "El Personaje " & Char(A) & " es newbie.")

                Exit Function

            End If
            

            
            
        End If

    Next A
    
    SaleChar_CharValid_Chars = True
    
End Function

' MERCADER :: Comprobaciones para publicar/ofertar/aceptar personajes. (Anti Hacking secundario)
Private Function SaleChar_CharValid_Money(ByVal UserIndex As Integer, _
                                          ByVal Account As String, _
                                          ByVal Gld As Long, _
                                          ByVal Eldhir As Long, _
                                          Optional ByVal ID As Integer = 0) As Boolean
    
    On Error GoTo ErrHandler
    
    Dim TempGld    As Long

    Dim TempEldhir As Long
    
    Dim AccountIndex As Integer
    
    AccountIndex = CheckEmailLogged(LCase$(Account))
    
    If AccountIndex > 0 Then
        TempGld = UserList(AccountIndex).Account.Gld
        TempEldhir = UserList(AccountIndex).Account.Eldhir
    Else
        TempGld = val(GetVar(AccountPath & Account & ACCOUNT_FORMAT, "INIT", "GLD"))
        TempEldhir = val(GetVar(AccountPath & Account & ACCOUNT_FORMAT, "INIT", "ELDHIR"))
    End If
    

        
    If ID > 0 Then
        If TempGld < MercaderDS(ID).Gld Then
            Call WriteErrorMsg(UserIndex, "El vendedor solicita un mínimo de " & MercaderDS(ID).Gld & " Monedas de Oro.")

            Exit Function

        End If
        
        If TempEldhir < MercaderDS(ID).Eldhir Then
            Call WriteErrorMsg(UserIndex, "El vendedor solicita un mínimo de " & MercaderDS(ID).Gld & " Monedas de Eldhir.")

            Exit Function

        End If
    End If
    
    If TempGld < Gld Then
        Call WriteErrorMsg(UserIndex, "Las Monedas de Oro deben estar en la boveda de la cuenta.")

        Exit Function

    End If

    If TempEldhir < Eldhir Then
        Call WriteErrorMsg(UserIndex, "Las Monedas de Eldhir deben estar en la boveda de la cuenta.")

        Exit Function

    End If
    
    SaleChar_CharValid_Money = True
    
    Exit Function
    
ErrHandler:
    SaleChar_CharValid_Money = False
    
End Function

' MERCADER :: Agregamos una nueva publicación.
Public Sub SaleChar_Add(ByVal UserIndex As Integer, _
                        ByRef UserName() As String, _
                        ByVal Gld As Long, _
                        ByVal Eldhir As Long, _
                        ByVal Blocked As Byte, _
                        ByVal Key As String, _
                        ByVal KeyMao As String)

    On Error GoTo ErrHandler

    Dim KeyTemp      As String

    Dim SlotMercader As Integer

    Dim A            As Long

    If UserList(UserIndex).Account.SaleCharSlot > 0 Then
        Call WriteErrorMsg(UserIndex, "Ya tienes una publicación en curso.")

        Exit Sub

    End If
    
    If Not StrComp(KeyMao, UserList(UserIndex).Account.KeyMao) = 0 Then
        Call WriteErrorMsg(UserIndex, "La clave de seguridad ingresada no es válida.")

        Exit Sub

    End If
    
    With UserList(UserIndex)

        If UBound(UserName) > 0 Then
            If Not .Account.Premium Then
                Call WriteErrorMsg(UserIndex, "Tu cuenta debe ser premium para poder realizar una venta de más de un personaje.")

                Exit Sub

            End If
        End If
        
        KeyTemp = GetVar(AccountPath & .Account.Email & ACCOUNT_FORMAT, "INIT", "KEY")
        
        If Not StrComp(KeyTemp, Key) = 0 Then
            Call WriteErrorMsg(UserIndex, "La Clave Pin no coincide con el de la cuenta. ¡Evita bloquear tu cuenta!")

            Exit Sub

        End If


        SlotMercader = SaleChar_FreeMercader
        
        If SlotMercader = 0 Then
            Call WriteErrorMsg(UserIndex, "No hay más lugar disponible para realizar publicaciones. Espera a que alguna termine.")

            Exit Sub

        End If
        
        If Not SaleChar_CharValid_Chars(UserIndex, UserList(UserIndex).Account.Email, UserName) Then
    
            Exit Sub
    
        End If
        
        Dim TemporalNicks As String
        
        ' Actualización en tiempo real del mercader
        With MercaderDS(SlotMercader)
            .Blocked = Blocked
            
            '.Char = UserName
            .Gld = Gld
            .Eldhir = Eldhir
            .Account = UserList(UserIndex).Account.Email
            
            For A = LBound(UserName) To UBound(UserName)
                .Char(A) = UserName(A)
               ' .CharText(A) = SaleChar_Text(.Char(A))
                '.CharInfo(A) = SaleChar_ListInfo_Load_Offline(.Char(A))
                
                TemporalNicks = TemporalNicks & .Char(A) & "-"
                If .Char(A) <> vbNullString Then
                    .CharAmount = .CharAmount + 1
                End If
                    
            Next A
            
        End With
        
        ' Actualización del SLOT en la cuenta
        Call SaleChar_SetSlot(UserIndex, SlotMercader)

        ' Actualización del SLOT del MERCADER
        Call SaleChar_SaveSlot(SlotMercader)
        
        Call WriteErrorMsg(UserIndex, "Publicación realizada exitosamente. Podrás visualizarla de una manera muy simple. ¡Fíjate!")

        Call Logs_Security(eSecurity, eMercader, "Cuenta " & .Account.Email & " con IP: " & .IP & " publicó " & TemporalNicks & ". Oro: " & Gld & " y Eldhir: " & Eldhir)
    End With
    
    Exit Sub

ErrHandler:
    Call LogError("SaleChar_Add ERROR")
End Sub

Public Function SaleChar_Text(ByVal Char As String)

    On Error GoTo ErrHandler
    
    Dim Reader As clsIniManager

    Set Reader = New clsIniManager
    
    Dim Class As eClass

    Dim Raze         As eRaza

    Dim Elv          As Byte

    Dim Exp          As Long

    Dim Elu          As Long

    Dim Ups          As Integer

    Dim Penas        As Byte

    Dim Hp           As Integer

    Dim Constitution As Byte
    
    Dim Charfile     As String: Charfile = CharPath & Char & ".chr"

    Dim Text         As String
    
    Dim TextUps As String
    
    Reader.Initialize Charfile
    
    Class = val(Reader.GetValue("INIT", "CLASE"))
    Raze = val(Reader.GetValue("INIT", "RAZA"))
    Elv = val(Reader.GetValue("STATS", "ELV"))
    Exp = val(Reader.GetValue("STATS", "EXP"))
    Elu = val(Reader.GetValue("STATS", "ELU"))
    Hp = val(Reader.GetValue("STATS", "MAXHP"))
    Constitution = val(Reader.GetValue("ATRIBUTOS", "AT" & eAtributos.Constitucion))
    Ups = UserCheckPromedy(Elv, Hp, Class, Constitution)
    
    If Ups > 0 Then
        TextUps = "+" & Ups
    ElseIf Ups < 0 Then
        TextUps = Ups
    ElseIf Ups = 0 Then
        TextUps = "Prom"
    End If
    
    Text = UCase$(Char) & "." & ListaClases(Class) & "." & ListaRazas(Raze) & "." & Elv & "(" & TextUps & ")"
    If Elv <> STAT_MAXELV Then
        Text = Text & "(" & Round(CDbl(Exp) * CDbl(100) / CDbl(Elu), 2) & "%)"
    End If
    
    SaleChar_Text = Text
    
    Set Reader = Nothing
    
    Exit Function

ErrHandler:
    Set Reader = Nothing
End Function

' MERCADER :: Agregamos una nueva oferta en la publicación.
Public Sub SaleChar_AddOffer(ByVal UserIndex As Integer, _
                             ByVal ID As Integer, _
                             ByVal Char As String, _
                             ByVal Gld As Long, _
                             ByVal Eldhir As Long, _
                             ByVal Key As String, _
                             ByVal KeyMao As String)

    On Error GoTo ErrHandler
    
    Dim UserName() As String

    Dim Slot       As Byte

    Dim TempChar   As String

    Dim A          As Long
    
    If ID = 0 Or ID > MAX_MERCADER_SLOT Then

        ' Anti Hacking
        Exit Sub

    End If
    
    UserName = Split(Char, "-")
    
    If Not SaleChar_CharValid(UserIndex, UserName, Gld, Eldhir) Then

        ' Anti Hacking
        Exit Sub

    End If
    
    If Not StrComp(KeyMao, UserList(UserIndex).Account.KeyMao) = 0 Then
        Call WriteErrorMsg(UserIndex, "La clave de seguridad ingresada no es válida.")

        Exit Sub

    End If
    
    If StrComp(MercaderDS(ID).Account, UserList(UserIndex).Account.Email) = 0 Then
        Call WriteErrorMsg(UserIndex, "¡No puedes ofrecerte a ti mismo!")

        Exit Sub

    End If
    
    If MercaderDS(ID).Account = vbNullString Then
        Call WriteErrorMsg(UserIndex, "La publicación que has seleccionado ha terminado justo antes de que aprietes el botón. ")
        ' Reenviar lista de mercado

        Exit Sub

    End If
    
    If Not SaleChar_CharValid_Money(UserIndex, UserList(UserIndex).Account.Email, Gld, Eldhir, ID) Then

        Exit Sub

    End If
    
    If SaleChar_Search_Offer(ID, UserList(UserIndex).Account.Email) Then
        Call WriteErrorMsg(UserIndex, "¡Ya has ofrecido a esta publicación!")
        Exit Sub
    End If
    
    
    Slot = SaleChar_FreeMercader_Offer(ID, UserList(UserIndex).Account.Email)
    
    If Slot > 0 Then

        With MercaderDS(ID).Offers(Slot)
            '.Char = UserName
            
            If Not UBound(UserName) = -1 Then
                For A = LBound(UserName) To UBound(UserName)
                    .Char(A) = UserName(A)
                    
                    If .Char(A) <> vbNullString Then
              
                        .CharAmount = .CharAmount + 1
                    End If
                    
                Next A
            End If
            
            
            .Gld = Gld
            .Eldhir = Eldhir
            .Account = UserList(UserIndex).Account.Email
            .LastTime = GetTime
        End With
        
    Else
        Call WriteErrorMsg(UserIndex, "El usuario recibió muchas ofertas o bien ya has ofrecido en esta publicación.")

        Exit Sub
    End If
    
    Call WriteErrorMsg(UserIndex, "Has enviado la oferta. Espera prontas noticias para concretar el cambio.")
    
    Call Logs_Security(eSecurity, eMercader, "Cuenta " & UserList(UserIndex).Account.Email & " con IP: " & UserList(UserIndex).IP & " ofertó a la cuenta " & MercaderDS(ID).Account & " los personajes " & Char & ". Oro: " & Gld & " y Eldhir: " & Eldhir)
    Exit Sub

ErrHandler:
    Call LogError("SaleChar_AddOffer ERROR")
End Sub

' MERCADER :: Aceptamos la oferta seleccionada.
Public Sub SaleChar_AcceptOffer(ByVal UserIndex As Integer, _
                                ByVal Slot As Integer, _
                                ByVal SlotOffer As Byte)
    
    On Error GoTo ErrHandler
    
    Dim TempChars As Byte

    Dim IsPremium As Byte
    
    If Slot <= 0 Or Slot > MAX_MERCADER_SLOT Then

        ' Anti Hacking:: El slot de la publicación es inválido.
        Exit Sub

    End If
    
    If SlotOffer <= 0 Or SlotOffer > MERCADER_MAX_OFFER Then

        ' Anti Hacking:: El slot de la oferta es inválido.
        Exit Sub

    End If
10
    If UserList(UserIndex).Account.SaleCharSlot <> Slot Then

        ' Anti Hacking:: El slot de la publicación es inválido.
        Exit Sub

    End If
    
    ' Comprueba que los personajes de la oferta sigan vigentes.
    
    If Not SaleChar_CharValid_Chars(UserIndex, MercaderDS(Slot).Offers(SlotOffer).Account, MercaderDS(Slot).Offers(SlotOffer).Char) Then

        Exit Sub

    End If
20
    ' Comprueba que las Monedas de la oferta sigan vigentes.
    If Not SaleChar_CharValid_Money(UserIndex, MercaderDS(Slot).Offers(SlotOffer).Account, MercaderDS(Slot).Offers(SlotOffer).Gld, MercaderDS(Slot).Offers(SlotOffer).Eldhir) Then

        Exit Sub

    End If

    ' Comprueba que los personajes de la de la publicación estén vigentes.
    If Not SaleChar_CharValid_Chars(UserIndex, MercaderDS(Slot).Account, MercaderDS(Slot).Char) Then

        Exit Sub

    End If
    
30
    ' ¿Va a recibir personaje a cambio?
    If MercaderDS(Slot).Offers(SlotOffer).CharAmount > 0 Then
        
        ' Personaje dueño de la publicación. Comprueba espacio necesario para alojar personajes.
        If UserList(UserIndex).Account.CharsAmount - MercaderDS(Slot).CharAmount + MercaderDS(Slot).Offers(SlotOffer).CharAmount > (ACCOUNT_MAX_CHARS) Then
            Call WriteErrorMsg(UserIndex, "No tienes espacio suficiente para recibir nuevos personajes.")
    
            Exit Sub
    
        End If
    End If
40
    ' Personaje que realizo la oferta. Comprueba espacio ncesario para alojar personajes.
    TempChars = val(GetVar(AccountPath & MercaderDS(Slot).Offers(SlotOffer).Account & ACCOUNT_FORMAT, "INIT", "CHARSAMOUNT"))
        
    If TempChars - MercaderDS(Slot).Offers(SlotOffer).CharAmount + MercaderDS(Slot).CharAmount > (ACCOUNT_MAX_CHARS) Then
        Call WriteErrorMsg(UserIndex, "La persona que envió la oferta no tiene espacio suficiente para recibir nuevos personajes.")

        Exit Sub

    End If
    
    Call Logs_Security(eSecurity, eMercader, "Cuenta " & MercaderDS(Slot).Account & " aceptó la oferta de " & MercaderDS(Slot).Offers(SlotOffer).Account)
    
    Dim tComprador As Integer
    
    Call WriteDisconnect(UserIndex, True)
50

    ' Vendedor recibe oro (Está siempre online)
    UserList(UserIndex).Account.Gld = UserList(UserIndex).Account.Gld + MercaderDS(Slot).Offers(SlotOffer).Gld
    UserList(UserIndex).Account.Eldhir = UserList(UserIndex).Account.Eldhir + MercaderDS(Slot).Offers(SlotOffer).Eldhir
    
                                
    ' Cerramos nuestra cuenta
    Call Protocol.Kick(UserIndex, "¡Mercado de Personaje: Tu cuenta se ve involucrada a la hora de concretar una venta. Por lo cual serás expulsado del juego y al ingresar, tendrás el nuevo cambio!")
60
    ' Le cerramos la cuenta en caso de estar logeada. (comprador)
    tComprador = CheckEmailLogged(MercaderDS(Slot).Offers(SlotOffer).Account)
    If tComprador > 0 Then
        UserList(tComprador).Account.Gld = UserList(tComprador).Account.Gld - MercaderDS(Slot).Offers(SlotOffer).Gld
        UserList(tComprador).Account.Eldhir = UserList(tComprador).Account.Eldhir - MercaderDS(Slot).Offers(SlotOffer).Eldhir
    
                                
        Call Protocol.Kick(tComprador, "¡Mercado de Personaje: Tu cuenta se ve involucrada a la hora de concretar una venta. Por lo cual serás expulsado del juego y al ingresar, tendrás el nuevo cambio!")
    Else
        ' Comprador pierde el orelio
        Call Mercader_ModifyMoney(AccountPath & MercaderDS(Slot).Offers(SlotOffer).Account & ACCOUNT_FORMAT, _
                                -MercaderDS(Slot).Offers(SlotOffer).Gld, _
                                -MercaderDS(Slot).Offers(SlotOffer).Eldhir)
    End If
70
    ' Personaje de la venta
    Call Mercader_ModifyChar(MercaderDS(Slot).Account, MercaderDS(Slot).Char, 0, vbNullString)
    
    Call Mercader_ModifyChar(MercaderDS(Slot).Offers(SlotOffer).Account, _
                                 MercaderDS(Slot).Char, _
                                 255, _
                                 MercaderDS(Slot).Offers(SlotOffer).Account)
    
80
    ' Quitamos los personajes de la cuenta n°2
    If MercaderDS(Slot).Offers(SlotOffer).Char(0) <> vbNullString Then
        Call Mercader_ModifyChar(MercaderDS(Slot).Offers(SlotOffer).Account, MercaderDS(Slot).Offers(SlotOffer).Char, 0, vbNullString)
        
        ' Le damos el personaje al vendedor
        Call Mercader_ModifyChar(MercaderDS(Slot).Account, _
                                  MercaderDS(Slot).Offers(SlotOffer).Char, _
                                 255, _
                                 MercaderDS(Slot).Offers(SlotOffer).Account)
    End If
    
90
    
    If MercaderDS(Slot).Offers(SlotOffer).Account <> vbNullString Then
        Dim SaleCharSlot As Integer
        SaleCharSlot = val(GetVar(AccountPath & MercaderDS(Slot).Offers(SlotOffer).Account & ACCOUNT_FORMAT, "INIT", "SALECHARSLOT"))
        
        If SaleCharSlot > 0 Then
            Call WriteVar(AccountPath & MercaderDS(Slot).Offers(SlotOffer).Account & ACCOUNT_FORMAT, "INIT", "SALECHARSLOT", "0")
            Call SaleChar_Reset(Slot)
            Call SaleChar_SaveSlot(Slot)
        End If
    End If
    
    Call SaleChar_SetSlot(UserIndex, 0)
    Call SaleChar_Reset(Slot)
    Call SaleChar_SaveSlot(Slot)
100
    
    
    Exit Sub

ErrHandler:
    Call LogError("SaleChar_AcceptOffer ERROR " & Err.description & " en linea " & Erl)
End Sub

' MERCADEr :: Modificamos el Oro y/o de la cuenta
Private Sub Mercader_ModifyMoney(ByVal FilePath As String, _
                                ByVal Gld As Long, _
                                ByVal Eldhir As Long)
                                
                                               
    Dim TempGld As Long
    Dim TempEldhir As Long
    
    TempGld = val(GetVar(FilePath, "INIT", "GLD"))
    TempEldhir = val(GetVar(FilePath, "INIT", "ELDHIR"))
    
    Call WriteVar(FilePath, "INIT", "GLD", CStr(TempGld + Gld))
    Call WriteVar(FilePath, "INIT", "ELDHIR", CStr(TempEldhir + Eldhir))
End Sub
' MERCADER :: Modificamos los personajes de la cuenta
Private Sub Mercader_ModifyChar(ByVal Account As String, _
                                ByRef Char() As String, _
                                ByVal Slot As Byte, _
                                ByVal AccountName As String)
                                 
    Dim A        As Long

    Dim TempSlot As Byte

    Dim cant     As Byte
    
    Dim CantChars As Integer
    
    
    For A = LBound(Char) To UBound(Char)
        If Char(A) <> vbNullString Then
            Call WriteVar(CharPath & UCase$(Char(A)) & ".chr", "INIT", "ACCOUNTNAME", AccountName)
            
            ' Si lo borra entonces >> Slot=0
            If Slot = 0 Then
                TempSlot = val(GetVar(CharPath & UCase$(Char(A)) & ".chr", "INIT", "ACCOUNTSLOT"))
                
                Call WriteVar(AccountPath & Account & ACCOUNT_FORMAT, "CHARS", CStr(TempSlot), vbNullString)
                Call WriteVar(CharPath & UCase$(Char(A)) & ".chr", "INIT", "ACCOUNTSLOT", CStr(Slot))
                
                CantChars = CantChars - 1
            Else
            
                CantChars = CantChars + 1
                TempSlot = Account_SearchFreeChar_Offline(AccountPath & Account & ACCOUNT_FORMAT)
                
                Call WriteVar(CharPath & UCase$(Char(A)) & ".chr", "INIT", "ACCOUNTSLOT", CStr(TempSlot))
                Call WriteVar(CharPath & UCase$(Char(A)) & ".chr", "INIT", "ACCOUNTNAME", Account)
                
                Call WriteVar(AccountPath & Account & ACCOUNT_FORMAT, "CHARS", CStr(TempSlot), Char(A))
            End If
            
            
        End If
    Next A
    
    cant = val(GetVar(AccountPath & Account & ACCOUNT_FORMAT, "INIT", "CHARSAMOUNT"))
    Call WriteVar(AccountPath & Account & ACCOUNT_FORMAT, "INIT", "CHARSAMOUNT", CStr(cant + CantChars))
End Sub

' MERCADER :: Borramos los personajes de la cuenta
Private Sub Mercader_DeleteChar(ByVal Account As String, _
   ByRef Char() As String)
                                 
    'If Not VarType(Char) = vbArray Then Exit Sub
    
    Dim A        As Long

    Dim TempSlot As Byte

    Dim cant     As Byte
    
    
    cant = val(GetVar(AccountPath & Account & ACCOUNT_FORMAT, "INIT", "CHARSAMOUNT"))
    
    For A = LBound(Char) To UBound(Char)
    
        TempSlot = val(GetVar(CharPath & UCase$(Char(A)) & ".chr", "INIT", "ACCOUNTSLOT"))
        
        Call WriteVar(CharPath & UCase$(Char(A)) & ".chr", "INIT", "ACCOUNTSLOT", "0")
        Call WriteVar(CharPath & UCase$(Char(A)) & ".chr", "INIT", "ACCOUNTNAME", vbNullString)
        Call WriteVar(AccountPath & Account & ACCOUNT_FORMAT, "CHARS", TempSlot, vbNullString)
    Next A
    
    Call WriteVar(AccountPath & Account & ACCOUNT_FORMAT, "INIT", "CHARSAMOUNT", CStr(cant - UBound(Char) + 1))
End Sub

' MERCADER :: Seteamos los personajes en una cuenta online
Private Sub Mercader_SetChar(ByVal UserIndex As Integer, _
                             ByVal Account As String, _
                             ByRef Char() As String)
    
    Dim A    As Long, B As Long

    Dim Slot As Byte
    
    With UserList(UserIndex)
        
        For A = LBound(Char) To UBound(Char)
        
            For B = 1 To ACCOUNT_MAX_CHARS

                If .Account.Chars(B).Name = vbNullString Then
                    .Account.Chars(B).Name = Char(A)
                    .Account.CharsAmount = .Account.CharsAmount + 1
                    
                    Call WriteVar(AccountPath & Account & ACCOUNT_FORMAT, "CHARS", A, Char(A))
                    Call WriteVar(CharPath & UCase$(Char(B)) & ".chr", "INIT", "ACCOUNTSLOT", CStr(B))
                    Call WriteVar(CharPath & UCase$(Char(B)) & ".chr", "INIT", "ACCOUNTNAME", Account)
                    
                    Exit For

                End If

            Next B
            
        Next A
        
        Call WriteVar(AccountPath & Account & ACCOUNT_FORMAT, "INIT", "CHARSAMOUNT", CStr(.Account.CharsAmount))
    End With

    ' ¡¡ UPGRADE PANEL DE CUENTAS DEL CLIENTE!! WriteLoginAccount()
    
End Sub

' MERCADER :: Seteamos los personajes en una cuenta offline
Private Sub Mercader_SetChar_Offline(ByVal UserIndex As Integer, _
                                     ByVal Account As String, _
                                     ByRef Char() As String)
    
    Dim A                                 As Long, B As Long

    Dim Slot                              As Byte

    Dim CopyChars(1 To ACCOUNT_MAX_CHARS) As String

    Dim CharsAmount                       As Byte
    
    For B = 1 To ACCOUNT_MAX_CHARS
        CopyChars(B) = GetVar(AccountPath & Account & ACCOUNT_FORMAT, "CHARS", A)
    Next B
    
    CharsAmount = val(GetVar(AccountPath & Account & ACCOUNT_FORMAT, "INIT", "CHARSAMOUNT"))
    
    For A = LBound(Char) To UBound(Char)
        For B = 1 To ACCOUNT_MAX_CHARS
        
            If CopyChars(B) = vbNullString Then
                Call WriteVar(AccountPath & Account & ACCOUNT_FORMAT, "CHARS", A, Char(A))
                Call WriteVar(CharPath & UCase$(Char(A)) & ".chr", "INIT", "ACCOUNTSLOT", CStr(B))
                Call WriteVar(CharPath & UCase$(Char(A)) & ".chr", "INIT", "ACCOUNTNAME", Account)
                CharsAmount = CharsAmount + 1
                
                Exit For

            End If

        Next B
        
    Next A
    
End Sub

Public Sub SaleChar_RemoveOffer(ByVal UserIndex As Integer, _
                                ByVal Slot As Byte, _
                                ByVal SlotOffer As Byte)
                                
    On Error GoTo ErrHandler
    
    If Slot <= 0 Or Slot > MAX_MERCADER_SLOT Then

        ' Anti Hacking:: El slot del mercado es inválido.
        Exit Sub

    End If
    
    If SlotOffer <= 0 Or SlotOffer > MERCADER_MAX_OFFER Then

        ' Anti Hacking:: El slot de la oferta es inválido.
        Exit Sub

    End If
    
    If UserList(UserIndex).Account.SaleCharSlot <> Slot Then

        ' Anti Hacking:: El slot de la publicación es inválido.
        Exit Sub

    End If
    
    Exit Sub

ErrHandler:
    Call LogError("SaleChar_RemoveOffer ERROR")
    
End Sub

Public Sub SaleChar_Remove(ByVal UserIndex As Integer)
    
    Dim Slot    As Integer

    Dim Account As String
    
    Slot = UserList(UserIndex).Account.SaleCharSlot
    
    If Slot = 0 Then
        Call WriteErrorMsg(UserIndex, "No tienes ninguna publicación vigente.")

        Exit Sub

    End If
        
    Call SaleChar_SetSlot(UserIndex, 0)
    Call SaleChar_Reset(Slot)
    Call SaleChar_SaveSlot(Slot)
    
    Call WriteErrorMsg(UserIndex, "¡Publicación eliminada!")
End Sub

Public Sub SaleChar_Reset(ByVal Slot As Integer)

    Dim Char() As String

    Dim A      As Long, B As Long
    
    With MercaderDS(Slot)
        .Account = vbNullString
        .Blocked = 0
        '.Char = Char
        .Gld = 0
        .Eldhir = 0
        
        For A = 0 To MAX_MERCADER_CHARS - 1
            .Char(A) = vbNullString
            .CharText(A) = vbNullString
        Next A
        
        For A = 1 To MERCADER_MAX_OFFER
            For B = 0 To MAX_MERCADER_CHARS - 1
                .Offers(A).Char(B) = vbNullString
            Next B
        
            '.Offers(A).Char = Char
            .Offers(A).Gld = 0
            .Offers(A).Eldhir = 0
            .Offers(A).Account = vbNullString
            .Offers(A).LastTime = 0
        Next A

    End With

End Sub

Public Sub SaleChar_ListInfo(ByVal UserIndex As Integer, ByVal MercaderSelected As Long)
    
    On Error GoTo ErrHandler
    
    Dim A As Long
    
    If MercaderSelected <= 0 Or MercaderSelected > MAX_MERCADER_SLOT Then

        ' Anti Hacking ::
        Exit Sub

    End If
    
    If MercaderDS(MercaderSelected).Account = vbNullString Then
        Call WriteErrorMsg(UserIndex, "Parece que la publicación ya no está vigente.")
        Call WriteMercader_List(UserIndex)

        Exit Sub

    End If
    
    Call WriteMercader_ListInfo(UserIndex, MercaderSelected, 0, MercaderDS(MercaderSelected).CharInfo)
    
    Exit Sub

ErrHandler:
    Call LogError("SaleChar_ListInfo ERROR")
    
End Sub

Public Sub SaleChar_SendOffers(ByVal UserIndex As Integer)
    
    On Error GoTo ErrHandler
    
    If UserList(UserIndex).Account.SaleCharSlot = 0 Then
        Call WriteErrorMsg(UserIndex, "No tienes ninguna publicación vigente.")
    Else

        If Not Interval_InfoChar(UserIndex) Then
            Call WriteErrorMsg(UserIndex, "Debes esperar unos segundos antes de solicitar otra vez la lista")
        Else
            Call WriteMercader_ListOffer(UserIndex)
        End If
    End If
                
    Exit Sub

ErrHandler:
    Call LogError("SaleChar_SendOffers ERROR")
    
End Sub

Public Sub SaleChar_SendOfferInfo(ByVal UserIndex As Integer, ByVal OfferSelected As Long)
    
    On Error GoTo ErrHandler
    
    Dim A            As Long

    Dim SlotMercader As Byte

    Dim CharInfo(4)  As tAccountMercader_CharSale_Info
    
    If OfferSelected <= 0 Or OfferSelected > MERCADER_MAX_OFFER Then

        ' Anti Hacking ::
        Exit Sub

    End If
    
    If Not Interval_InfoChar(UserIndex) Then
        Call WriteErrorMsg(UserIndex, "Debes esperar unos segundos antes de solicitar más información de personajes.")

        Exit Sub

    End If
    
    SlotMercader = UserList(UserIndex).Account.SaleCharSlot
    
    For A = 0 To 4

        With MercaderDS(SlotMercader).Offers(OfferSelected)

            If .Char(A) <> vbNullString Then
                CharInfo(A) = SaleChar_ListInfo_Load_Offline(.Char(A))
            End If

        End With

    Next A
    
    Call WriteMercader_ListInfo(UserIndex, SlotMercader, OfferSelected, CharInfo)
    
    Exit Sub

ErrHandler:
    Call LogError("SaleChar_ListInfo ERROR")
    
End Sub

Public Function SaleChar_ListInfo_Load_Offline(ByVal UserName As String) As tAccountMercader_CharSale_Info

    Dim Read         As clsIniManager

    Dim A            As Long
    
    Dim Temp         As tAccountMercader_CharSale_Info

    Dim Constitution As Byte

    Dim Text         As String

    If Not PersonajeExiste(UserName) Then

        ' Anti Hacking o Error fatal
        Exit Function

    End If
        
    Set Read = New clsIniManager
        
    Call Read.Initialize(CharPath & UserName & ".chr")
        
    Temp.Name = UserName
        
    Temp.Hp = val(Read.GetValue("STATS", "MAXHP"))
    Temp.Man = val(Read.GetValue("STATS", "MAXMAN"))
    Constitution = val(Read.GetValue("ATRIBUTOS", "AT" & eAtributos.Constitucion))
        
    Temp.Class = val(Read.GetValue("INIT", "CLASE"))
    Temp.Raze = val(Read.GetValue("INIT", "RAZA"))
    Temp.Promedio = val(Read.GetValue("REP", "PROMEDIO"))
        
    Temp.Body = val(Read.GetValue("INIT", "BODY"))
    Temp.Head = val(Read.GetValue("INIT", "HEAD"))
    Temp.Helm = val(Read.GetValue("INIT", "CASCO"))
    Temp.Shield = val(Read.GetValue("INIT", "ESCUDO"))
    Temp.Weapon = val(Read.GetValue("INIT", "ARMA"))
    Temp.Genero = val(Read.GetValue("INIT", "GENERO"))
        
    Temp.Elv = val(Read.GetValue("STATS", "ELV"))
    Temp.Exp = val(Read.GetValue("STATS", "EXP"))
    Temp.Elu = val(Read.GetValue("STATS", "ELU"))
        
    Temp.Faction = val(Read.GetValue("FACTION", "STATUS"))
    Temp.FactionEx = val(Read.GetValue("FACTION", "EXFACTION"))
        
    Temp.FragsCiu = val(Read.GetValue("FACTION", "FRAGSCIU"))
    Temp.FragsCri = val(Read.GetValue("FACTION", "FRAGSCRI"))
    Temp.FragsOther = val(Read.GetValue("FACTION", "FRAGSOTHER"))
    Temp.FragsNpc = val(Read.GetValue("MUERTES", "NPCSMUERTES"))
        
    Temp.Gld = val(Read.GetValue("STATS", "GLD"))
    Temp.Eldhir = val(Read.GetValue("STATS", "ELDHIR"))
        
    Text = val(Read.GetValue("GUILD", "GUILDINDEX"))
        
    If val(Text) > 0 Then
        Temp.GuildName = GuildsInfo(val(Text)).Name
            
        Text = Read.GetValue("GUILD", "GUILDRANGE")
        Temp.GuildRange = Guilds_PrepareRangeName(val(Text))
    End If
        
    For A = 1 To MAX_INVENTORY_SLOTS
        Text = Read.GetValue("INVENTORY", "OBJ" & A)
            
        Temp.Object(A).ObjIndex = val(ReadField(1, Text, Asc("-")))
        Temp.Object(A).Amount = val(ReadField(2, Text, Asc("-")))
    Next A
        
    For A = 1 To MAX_BANCOINVENTORY_SLOTS
        Text = Read.GetValue("BANCOINVENTORY", "OBJ" & A)
            
        Temp.ObjectBank(A).ObjIndex = val(ReadField(1, Text, Asc("-")))
        Temp.ObjectBank(A).Amount = val(ReadField(2, Text, Asc("-")))
    Next A

    For A = 1 To 35
        Text = val(Read.GetValue("HECHIZOS", "H" & A))
            
        If Text <> 0 Then
            Temp.Spells(A) = Hechizos(Text).Nombre
        End If

    Next A
        
    For A = 1 To MAX_MEDITATION
        Text = val(Read.GetValue("MEDITATION", A))
            
        If Text > 0 Then
            Temp.Meditations(A) = Meditation(A).FX
        End If

    Next A
        
    Temp.Ban = val(Read.GetValue("FLAGS", "BAN"))
    Temp.Penas = val(Read.GetValue("PENAS", "CANT"))
        
    If Temp.Penas > 0 Then
        ReDim Temp.PenasText(0 To Temp.Penas - 1) As String
            
        For A = 0 To Temp.Penas - 1
            Temp.PenasText(A) = Read.GetValue("PENAS", "P" & A + 1)
        Next A
            
    End If
        
    Temp.Bronce = val(Read.GetValue("FLAGS", "BRONCE"))
    Temp.Plata = val(Read.GetValue("FLAGS", "PLATA"))
    Temp.Oro = val(Read.GetValue("FLAGS", "ORO"))
    Temp.Premium = val(Read.GetValue("FLAGS", "PREMIUM"))
        
    Temp.Ups = UserCheckPromedy(Temp.Elv, Temp.Hp, Temp.Class, Constitution)
        
    Temp.Points = val(Read.GetValue("RANKING", "POINTS"))
    Temp.Retos1Ganados = val(Read.GetValue("RANKING", "RETOS1GANADOS"))
    Temp.Retos1Jugados = val(Read.GetValue("RANKING", "RETOS1JUGADOS"))
    Temp.Retos2Ganados = val(Read.GetValue("RANKING", "RETOS2GANADOS"))
    Temp.Retos2Jugados = val(Read.GetValue("RANKING", "RETOS2JUGADOS"))
    Temp.Retos3Ganados = val(Read.GetValue("RANKING", "RETOS3GANADOS"))
    Temp.Retos3Jugados = val(Read.GetValue("RANKING", "RETOS3JUGADOS"))
        
    Temp.EventosGanados = val(Read.GetValue("RANKING", "TORNEOSGANADOS"))
    Temp.EventosJugados = val(Read.GetValue("RANKING", "TORNEOSJUGADOS"))
        
    Temp.Bandido = val(Read.GetValue("REP", "BANDIDO"))
    Temp.Asesino = val(Read.GetValue("REP", "ASESINO"))
    Temp.Nobleza = val(Read.GetValue("REP", "NOBLES"))
    
    Dim TempSTR As String
    TempSTR = Read.GetValue("QUESTS", "QuestsDone")
    
    Temp.NumQuests = val(ReadField(1, TempSTR, 45))
    
    If Temp.NumQuests > 0 Then
        ReDim Temp.Quests(1 To Temp.NumQuests)

        For A = 1 To Temp.NumQuests
            Temp.Quests(A) = val(ReadField(A + 1, TempSTR, 45))
        Next A
    End If
    
    Set Read = Nothing
        
    SaleChar_ListInfo_Load_Offline = Temp
    
    Exit Function

ErrHandler:
    Set Read = Nothing
    Call LogError("SaleChar_ListInfo_Load_Offline ERROR")
    
End Function

Public Sub Mercader_Loop()

On Error GoTo ErrHandler

    Dim A As Long
    Dim B As Long
    
    For A = 1 To MAX_MERCADER_SLOT
        For B = 1 To MERCADER_MAX_OFFER
            Call Mercader_ResetOffer(A, B)
        Next B
    Next A

Exit Sub
ErrHandler:
    Call LogError("Error MercadeR_Loop")

End Sub

Private Sub Mercader_ResetOffer(ByVal Slot As Byte, ByVal SlotOffer As Byte)
    Dim NullOffer As tAccountMercader_CharSale_Offer
    MercaderDS(Slot).Offers(SlotOffer) = NullOffer

End Sub
