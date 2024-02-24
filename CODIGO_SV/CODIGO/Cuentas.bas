Attribute VB_Name = "Cuentas"
Option Explicit

' AYUDA
' CREAR NUEVA CUENTA: mCuenta.CreateAccount Account, Passwd, Pin, Email
' BORRAR CUENTA : mCuenta.KillAccount Account, Passwd, Pin, Email
' AGREGAR PERSONAJE A CUENTA: mCuenta.AddCharAccount UserName, Account, Passwd, Pin, Email
' RECUPERAR CONTRASEÑA : mCuenta.RecoverAccount Account, Pin, Email
' BORRAR PERSONAJE : mCuenta.KillCharAccount UserName, Account, Passwd, Pin, Email
' CAMBIAR CONTRASEÑA : mCuenta.ChangePasswdAccount Account, Pin, Email, OldPasswd, NewPasswd

Private Const DIR_ACCOUNT = "\ACCOUNT\"
Private Const FORMAT_ACCOUNT = ".DS"
Public Const MAX_PJS_ACCOUNT = 30

' Lista de Recuperación por tiempo
Private Type tRecover
    Account As String
    Email As String
End Type

Private Recovers(1 To 20) As tRecover
Public Const MAX_OBJ_PREMIUM As Byte = 60

''''''''''''''''''''''''''''''''''''''

Private Type tObj
    ObjIndex As Long
    Amount As Integer
End Type

Private Type tBank
    Login As Byte
    Obj(1 To MAX_OBJ_PREMIUM) As tObj
End Type

' Configuración de la Cuenta
Private Type tCuenta
    UserName(1 To MAX_PJS_ACCOUNT) As String
    Email As String
    Pin As String
    Passwd As String
    Bank As tBank
End Type

'''''''''''''''''''''''''''''''''''''''''


Private Const MAX_AMOUNT_PREMIUM = 100000


' ¿Existe la CUENTA?
Private Function ExistAccount(ByVal Account As String) As Boolean

    If FileExist(App.Path & DIR_ACCOUNT & UCase$(Account) & FORMAT_ACCOUNT) Then
        ExistAccount = True
        Exit Function
    End If
    
End Function

Public Function IsPremiumAccount(ByVal Account As Integer) As Boolean
    IsPremiumAccount = LoadDataAccount(Account, "ACCOUNT", "PREMIUM")
End Function

' Personaje tiene cuenta VIP?
Public Function IsPremium(ByVal UserIndex As Integer) As Boolean
    IsPremium = LoadDataAccount(UserList(UserIndex).Account, "ACCOUNT", "PREMIUM")
    
    If IsPremium Then
        UserList(UserIndex).IsPremium = True
        WriteConsoleMsg UserIndex, "Tu cuenta es Premium. Sigue disfrutando de los beneficios que te brinda Argentum Online", FontTypeNames.FONTTYPE_CITIZEN
    Else
        UserList(UserIndex).IsPremium = False
        WriteConsoleMsg UserIndex, "Tu cuenta NO Premium. ¡Enterate de los beneficios!", FontTypeNames.FONTTYPE_CITIZEN
    End If
End Function

' Chequeamos alguna información de la cuenta
Private Function CheckData(ByVal DataUser As String, _
                                ByVal Main As String, _
                                ByVal Account As String) As Boolean
    
    Dim FilePath As String
    Dim Temp As String
    
    FilePath = App.Path & DIR_ACCOUNT & UCase$(Account) & FORMAT_ACCOUNT
    
    If DataUser = GetVar(FilePath, "ACCOUNT", Main) Then
        CheckData = True
        Exit Function
    End If
End Function

' Obtenemos la contraseña de una Cuenta
Private Function GetPasswd(ByVal Account As String) As String
    Dim FilePath As String
    
    If Not ExistAccount(Account) Then Exit Function
    
    FilePath = App.Path & DIR_ACCOUNT & UCase$(Account) & FORMAT_ACCOUNT
    
    GetPasswd = GetVar(FilePath, "ACCOUNT", "PASSWD")
End Function

' Guardamos información de la Cuenta
Public Function SaveDataAccount(ByVal Account As String, _
                            ByVal Main As String, _
                            ByVal Var As String, _
                            ByVal Value As String)
    
    Dim FilePath As String
    FilePath = App.Path & DIR_ACCOUNT & UCase$(Account) & FORMAT_ACCOUNT
    
    WriteVar FilePath, Main, Var, Value
End Function

' Cargamos información de la cuenta
Public Function LoadDataAccount(ByVal Account As String, _
                            ByVal Main As String, _
                            ByVal Var As String) As String
    
    Dim FilePath As String
    FilePath = App.Path & DIR_ACCOUNT & UCase$(Account) & FORMAT_ACCOUNT
    
    LoadDataAccount = GetVar(FilePath, Main, Var)
End Function
' Buscamos el Personaje
Public Function SearchCharAccount(ByVal Account As String, ByVal UserName As String) As Byte
    Dim LoopC As Integer
    Dim Temp As String
    Dim FilePath As String
    
    FilePath = App.Path & DIR_ACCOUNT & UCase$(Account) & FORMAT_ACCOUNT
    
    For LoopC = 1 To MAX_PJS_ACCOUNT
        If UCase$(GetVar(FilePath, "ACCOUNT", "PERSONAJE" & LoopC)) = UCase$(UserName) Then
            SearchCharAccount = LoopC
            Exit For
        End If
    Next LoopC
End Function


' Creamos la Cuenta
Public Sub CreateAccount(ByVal UserIndex As Integer, _
                            ByVal Account As String, _
                            ByVal Passwd As String, _
                            ByVal Pin As String, _
                            ByVal Email As String)
                            
    Dim FilePath As String
    Dim LoopC As Integer
    
    FilePath = App.Path & DIR_ACCOUNT & UCase$(Account) & FORMAT_ACCOUNT
    
    If Len(Account) > 15 Then Exit Sub
    If Not AsciiValidos(Account, True) Then
        Call WriteErrorMsg(UserIndex, "Nombre inválido.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
    
    If ExistAccount(Account) Then
        ' El nombre de la cuenta ya está en uso.
        Exit Sub
    End If
    
    SaveDataAccount Account, "ACCOUNT", "PASSWD", Passwd
    SaveDataAccount Account, "ACCOUNT", "EMAIL", Email
    SaveDataAccount Account, "ACCOUNT", "PIN", Pin
    
    For LoopC = 1 To MAX_PJS_ACCOUNT
        SaveDataAccount Account, "ACCOUNT", "Personaje" & LoopC, "0"
    Next LoopC
    
    For LoopC = 1 To MAX_OBJ_PREMIUM
        SaveDataAccount Account, "BANK", "OBJ" & LoopC, "0-0-0"
    Next LoopC
    
    Protocol.WriteErrorMsg UserIndex, "La cuenta ha sido creada exitosamente"
End Sub

' Logeamos la Cuenta
Public Sub LoginAccount(ByVal UserIndex As Integer, _
                            ByVal Account As String, _
                            ByVal Passwd As String, _
                            Optional ByVal Deslogged As Boolean = False)
                                                 
                                                 
On Error GoTo ErrHandler

    Dim FileAccount As String
    Dim FileChar As String
    Dim Chars(1 To MAX_PJS_ACCOUNT) As tCuentaUser
    Dim LoopC As Integer
10

    FileAccount = App.Path & DIR_ACCOUNT & UCase$(Account) & FORMAT_ACCOUNT

    If Not Deslogged Then
        If Not ExistAccount(Account) Or _
          Not CheckData(Passwd, "PASSWD", Account) Then
            WriteErrorMsg UserIndex, "La cuenta no existe o la contraseña es inválida."
            Exit Sub
        End If
    End If
    
    For LoopC = 1 To MAX_PJS_ACCOUNT
        With Chars(LoopC)
            .Name = GetVar(FileAccount, "ACCOUNT", "PERSONAJE" & LoopC)

            If .Name <> "0" Then
32                FileChar = CharPath & UCase$(.Name) & ".chr"

33                .Ban = val(GetVar(FileChar, "FLAGS", "BAN"))
34                .clase = val(GetVar(FileChar, "INIT", "CLASE"))
35                .raza = val(GetVar(FileChar, "INIT", "RAZA"))
36                .Elv = val(GetVar(FileChar, "STATS", "ELV"))

40                .body = val(GetVar(FileChar, "INIT", "BODY"))
42                .Head = val(GetVar(FileChar, "INIT", "HEAD"))
                  .Helm = val(GetVar(FileChar, "INIT", "CASCO"))
45                .Shield = val(GetVar(FileChar, "INIT", "ESCUDO"))
                  .Weapon = val(GetVar(FileChar, "INIT", "ARMA"))
50
            Else
                  .Ban = 0
                  .clase = 0
                  .raza = 0
                  .Elv = 0
                  .body = 0
                  .Head = 0
                  .Helm = 0
                  .Weapon = 0
                  .Shield = 0
            End If
55
            
            WriteAccount_Data UserIndex, LoopC, Chars(LoopC)
        End With
    Next LoopC
    
    
    'WriteErrorMsg UserIndex, "Conectado"
    
  Exit Sub
  
ErrHandler:
  LogError "Error LoginAccount linea " & Erl
  
End Sub

' Logiamos el Personaje de la cuenta
Public Sub LoginCharAccount(ByVal UserIndex As Integer, _
                                ByVal Account As String, _
                                ByVal Passwd As String, _
                                ByVal UserName As String)
                            
    Dim FilePath As String

    FilePath = App.Path & DIR_ACCOUNT & UCase$(Account) & FORMAT_ACCOUNT
    
    If Not ExistAccount(Account) Then Exit Sub
    If Not CheckData(Passwd, "PASSWD", Account) Then Exit Sub
    If UCase$(Account) <> UCase$(GetVar(CharPath & UserName & ".chr", "INIT", "ACCOUNT")) Then Exit Sub
    
    If BANCheck(UserName) Then
        Call WriteErrorMsg(UserIndex, "Se te ha prohibido la entrada a Argentum Online.")
        Exit Sub
    End If
    
    If UserList(UserIndex).flags.UserLogged Then
        Call LogCheating("El usuario " & UserList(UserIndex).Name & " ha intentado loguear a " & UserName & " desde la IP " & UserList(UserIndex).ip)
        'Kick player ( and leave character inside :D )!
        Call CloseSocketSL(UserIndex)
        Call Cerrar_Usuario(UserIndex)
        Exit Sub
    End If
    
    'Controlamos no pasar el maximo de usuarios
    If NumUsers >= MaxUsers Then
        Call WriteErrorMsg(UserIndex, "El servidor ha alcanzado el máximo de usuarios soportado, por favor vuelva a intertarlo más tarde.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
    
    '¿Este IP ya esta conectado?
    If AllowMultiLogins = 0 Then
        If CheckForSameIP(UserIndex, UserList(UserIndex).ip) = True Then
            Call WriteErrorMsg(UserIndex, "No es posible usar más de un personaje al mismo tiempo.")
            Call FlushBuffer(UserIndex)
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
    End If
    
    '¿Existe el personaje?
    If Not FileExist(CharPath & UCase$(UserName) & ".chr", vbNormal) Then
        Call WriteErrorMsg(UserIndex, "El personaje no existe.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
    
    '¿Ya esta conectado el personaje?
    If CheckForSameName(UserName) Then
        If UserList(NameIndex(UserName)).Counters.Saliendo Then
            Call WriteErrorMsg(UserIndex, "El usuario está saliendo.")
        Else
            Call WriteErrorMsg(UserIndex, "Perdón, un usuario con el mismo nombre se ha logueado.")
        End If
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
    

    ConnectUser UserIndex, UserName, UserName
    UserList(UserIndex).Account = Account
End Sub

' Creamos un Nuevo Personaje
Public Sub CreateCharAccount(ByVal UserIndex As Integer, _
                                ByVal Account As String, _
                                ByVal Passwd As String, _
                                ByVal UserName As String, _
                                ByVal UserClase As Byte, _
                                ByVal UserRaza As Byte, _
                                ByVal UserSexo As Byte, _
                                ByVal UserFaccion As Byte)
                                
    Dim FilePath As String
    
    FilePath = App.Path & DIR_ACCOUNT & UCase$(Account) & FORMAT_ACCOUNT
    
    If Not ExistAccount(Account) Then Exit Sub
    If Not CheckData(Passwd, "PASSWD", Account) Then Exit Sub
    
    If Len(UserName) > 9 Then Exit Sub
    
    If Not AsciiValidos(UserName, True) Or LenB(UserName) = 0 Then
        WriteErrorMsg UserIndex, "Nombre inválido."
        Exit Sub
    End If
    
    If PersonajeExiste(UserName) Then
        WriteErrorMsg UserIndex, "El personaje ya existe."
        Exit Sub
    End If

    If UserList(UserIndex).flags.UserLogged Then
        Call LogCheating("El usuario " & UserList(UserIndex).Name & " ha intentado crear a " & UserName & " desde la IP " & UserList(UserIndex).ip)
              
        Call CloseSocketSL(UserIndex)
        Call Cerrar_Usuario(UserIndex)
              
        Exit Sub
    End If

    If ServerSoloGMs > 0 Then
        If (UserList(UserIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)) = 0 Then
            Call WriteErrorMsg(UserIndex, "Servidor restringido a administradores. Por favor reintente en unos momentos.")
            Call FlushBuffer(UserIndex)
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
    End If
    
    If UserClase > 9 Then Exit Sub

    ConnectNewUser UserIndex, UserName, Account, UserRaza, UserSexo, UserClase, UserFaccion
    AddCharAccount UserIndex, Account, UserName
End Sub

' Borramos la Cuenta ¡NO SE USA!
Public Sub KillAccount(ByVal Account As String, _
                            ByVal Passwd As String, _
                            ByVal Pin As String, _
                            ByVal Email As String)
    
    Dim FilePath As String
    
    FilePath = App.Path & DIR_ACCOUNT & UCase$(Account) & FORMAT_ACCOUNT
    
    If Not ExistAccount(Account) Then
        ' La cuenta que deseas eliminar no existe.
        Exit Sub
    End If
    
    If Not CheckData(Passwd, "PASSWD", Account) Then Exit Sub
    If Not CheckData(Pin, "PIN", Account) Then Exit Sub
    If Not CheckData(Email, "EMAIL", Account) Then Exit Sub
    
    Kill (App.Path & DIR_ACCOUNT & UCase$(Account) & FORMAT_ACCOUNT)
End Sub

' Agregamos el personaje a nuestra cuenta
Public Function AddCharAccount(ByVal UserIndex As Integer, _
                                ByVal Account As String, _
                                ByVal UserName As String) As Boolean
    
    If Account_CheckFreeUser(UCase$(Account)) = 0 Then
        Call WriteErrorMsg(UserIndex, "No tienes más espacio en tu cuenta.")
        AddCharAccount = False
        Exit Function
    End If
    
    SaveDataAccount Account, "ACCOUNT", "PERSONAJE" & Account_SearchFreeUser(UCase$(Account)), UserName
    
    AddCharAccount = True
End Function

' Borramos el personaje de la cuenta
Public Sub KillCharAccount(ByVal UserIndex As Integer, _
                            ByVal Account As String, _
                            ByVal Passwd As String, _
                            ByVal Index As Byte)
    
    Dim FilePath As String
    Dim Temp As String
    Dim UserName As String
    
    UserName = LoadDataAccount(Account, "ACCOUNT", "PERSONAJE" & Index)
    FilePath = App.Path & "\CHARFILE\" & UCase$(UserName) & ".chr"
    
    If Not CheckData(Passwd, "PASSWD", Account) Then Exit Sub
    
    If val(GetVar(FilePath, "FLAGS", "BAN")) = 1 Then
        WriteErrorMsg UserIndex, "No puedes borrar personajes baneados"
        Exit Sub
    End If
    
    SaveDataAccount Account, "ACCOUNT", "PERSONAJE" & Index, "0"
    Kill FilePath
    LoginAccount UserIndex, Account, Passwd
    
End Sub

' Recuperamos la Cuenta
Public Sub RecoverAccount(ByVal UserIndex As Integer, _
                            ByVal Account As String, _
                            ByVal Pin As String, _
                            ByVal Email As String)

    Dim Temp As String
    
    If Not ExistAccount(Account) Then Exit Sub
    If Not CheckData(Pin, "PIN", Account) Then Exit Sub
    If Not CheckData(Email, "EMAIL", Account) Then Exit Sub
    
    Temp = GeneratePasswd
    SaveDataAccount Account, "ACCOUNT", "PASSWD", Temp
    
    WriteErrorMsg UserIndex, "Cuenta: " & Account & vbCrLf & " Contraseña nueva: " & Temp
End Sub

Public Sub UpdateAccountUserName(ByVal UserName As String, _
                                  ByVal Account As String, _
                                    Optional ByVal NewNick As String = vbNullString)

    If Not ExistAccount(Account) Then Exit Sub
    
    Dim Slot As Byte
    
    Slot = SearchCharAccount(Account, UserName)
    
    SaveDataAccount Account, "ACCOUNT", "PERSONAJE" & Slot, "0"
    
    If NewNick <> vbNullString Then
        SaveDataAccount Account, "ACCOUNT", "PERSONAJE" & Slot, NewNick
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema de recuperación DE CUENTAS
' Cambiamos la contraseña de la Cuenta
Public Sub ChangePasswdAccount(ByVal Account As String, _
                                ByVal Pin As String, _
                                ByVal Email As String, _
                                ByVal OldPasswd As String, _
                                ByVal NewPasswd As String)
    
    If Not CheckData(OldPasswd, "PASSWD", Account) Then Exit Sub
    If Not CheckData(Email, "EMAIL", Account) Then Exit Sub
    If Not CheckData(Pin, "PIN", Account) Then Exit Sub
    
    SaveDataAccount Account, "ACCOUNT", "PASSWD", NewPasswd
    ' Tu contraseña ha sido modificada a NewPasswd
End Sub

Private Function GeneratePasswd() As String
    Randomize
    GeneratePasswd = Int(Rnd(1) * 10000) & Int(Rnd(1) * 10) & Int(Rnd(1) * 1000)
End Function

' Nos dice los slot libres de los que disponemos
Public Function Account_CheckFreeUser(ByVal Account As String) As Byte
    
    Dim A As Long
    
    For A = 1 To MAX_PJS_ACCOUNT
        If LoadDataAccount(Account, "ACCOUNT", "Personaje" & A) = "0" Then
            Account_CheckFreeUser = Account_CheckFreeUser + 1
        End If
    Next A
End Function

Public Function Account_SearchFreeUser(ByVal Account As String) As Byte
    Dim A As Long
    
    For A = 1 To MAX_PJS_ACCOUNT
        If LoadDataAccount(Account, "ACCOUNT", "Personaje" & A) = "0" Then
            Account_SearchFreeUser = A
            Exit Function
        End If
    Next A
End Function
