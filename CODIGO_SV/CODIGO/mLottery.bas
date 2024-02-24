Attribute VB_Name = "mLottery"
Option Explicit

Public LotteryLast As Long

' Spam cada una hora de los sorteos existentes. [Es importante hacer de a pocos]
Public Const LOTTERY_LAST_SPAM As Long = 60000


' Máximo de Personajes participes de un sorteo.
Public Const LOTTERY_MAX_CHARS As Long = 1000


' Máximo de usuarios que busca como ganador. Si hay 10 que estan offline el TDs es una reverenda poronga
Public Const LOTTERY_MAX_CHANCES As Long = 10


Type tLottery
    Name As String
    Desc As String
    
    DateInitial As String ' Inicio del sorteo
    DateFinish As String ' Fecha en la que se realiza el sorteo
    
    PrizeChar As String ' Personaje que va a ser sorteado
    PrizeObj As Integer ' Objeto que va a ser sorteado
    PrizeObjAmount As Integer   ' Cantidad del objeto que recibe
    
    CharLast As Integer
    Chars() As String
    LastSpam As Long            ' Tiempo que hace que spameo en la consola sobre el sorteo.
End Type

Public Lottery() As tLottery

' Carga la lista de sorteos vigentes.
Public Sub Lottery_Load()
    Dim Manager As clsIniManager
    Dim FilePath As String
    Dim A As Long
    Dim Temp As String
    
    Set Manager = New clsIniManager
    
    FilePath = DatPath & "lottery.dat"
    
    Manager.Initialize FilePath
    
    LotteryLast = val(Manager.GetValue("INIT", "LAST"))

    ReDim Lottery(0 To LotteryLast) As tLottery
    
    For A = 1 To LotteryLast
        With Lottery(A)
            .Name = Manager.GetValue(A, "NAME")
            .Desc = Manager.GetValue(A, "DESC")
            
            .DateInitial = Manager.GetValue(A, "DATEINITIAL")
            .DateFinish = Manager.GetValue(A, "DATEFINISH")
            
            .PrizeChar = Manager.GetValue(A, "PRIZECHAR")
            
            Temp = Manager.GetValue(A, "PRIZEOBJ")
            .PrizeObj = val(ReadField(1, Temp, 45))
            .PrizeObjAmount = val(ReadField(2, Temp, 45))
            
            
        End With
    
    Next A
    
    Set Manager = Nothing
End Sub

' Guarda la información del sorteo vigente. con opcional de guardado de usuarios cada WorldSave
Public Sub Lottery_Save()

    Dim Manager  As clsIniManager

    Dim FilePath As String

    Dim A        As Long

    Dim Temp     As String
    
    Set Manager = New clsIniManager
    
    FilePath = DatPath & "lottery.dat"
    
    For A = 1 To LotteryLast

        With Lottery(A)
            Call Manager.ChangeValue(A, "NAME", .Name)
            Call Manager.ChangeValue(A, "DESC", .Desc)
            
            Call Manager.ChangeValue(A, "DATEINITIAL", .DateInitial)
            Call Manager.ChangeValue(A, "DATEFINISH", .DateFinish)
            
            Call Manager.ChangeValue(A, "PRIZECHAR", .PrizeChar)
            Call Manager.ChangeValue(A, "PRIZEOBJ", .PrizeObj & "-" & .PrizeObjAmount)

            ' Guardamos los usuarios que participaron
        End With
    Next A
    
    Call Manager.DumpFile(FilePath)
    
    Set Manager = Nothing

End Sub

' Comprueba si un sorteo está por finalizar o realiza los mensajes de spam correspondientes
Public Sub Lottery_Loop()

    Dim A As Long

    Dim T As String
    
    T = Now
    
    For A = 1 To LotteryLast

        With Lottery(A)
        
            If DateDiff("s", T, .DateFinish) <= 0 Then
                Call Lottery_Finish(A)
            Else

                If (GetTime - .LastSpam) < LOTTERY_LAST_SPAM Then
                    SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & "» " & .Desc & " Sortea " & .DateFinish, FontTypeNames.FONTTYPE_SERVER)
                End If

            End If
        End With
    Next A

    End Sub

' Busca un nuevo slot para el SORTEO
Private Function Lottery_FreeSlot() As Integer
    Dim A As Integer
    
    For A = LBound(Lottery) To UBound(Lottery)
        If Lottery(A).Name = vbNullString Then
            Lottery_FreeSlot = A
            Exit Function
        End If
    Next A
    
End Function

' Inicia un nuevo sorteo
Public Sub Lottery_New(ByRef TempLottery As tLottery)
    
    Dim Slot As Integer
    
    Slot = Lottery_FreeSlot
    
    If Slot = 0 Then
        Slot = UBound(Lottery) + 1
        ReDim Preserve Lottery(LBound(Lottery) To Slot)
    End If
    
    Lottery(Slot) = TempLottery
    
    With Lottery(Slot)
        .DateInitial = Format(Now, "dd/mm/yyyy HH:MM")
        
        Call Logs_Security(eSecurity, eLottery, "Lottery_New:: Sorteo nro° " & Slot & " iniciado. Personaje: " & .PrizeChar & " Objeto: " & IIf(.PrizeObj > 0, ObjData(.PrizeObj).Name, "NINGUNO") & " Cantidad: " & .PrizeObjAmount)
    End With
   
   Dim Temp As String
   Temp = "«NUEVO SORTEO» Enterate de un nuevo sorteo disponible tipeando el comando /SORTEOS"
    SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg(Temp, FontTypeNames.FONTTYPE_SERVER)
    
    Call Lottery_Save
End Sub

' La loteria se cancela
Private Sub Lottery_Cancel(ByVal Slot As Integer, ByVal ShowMessage As Boolean)
    Dim LotteryNull As tLottery
    Lottery(Slot) = LotteryNull
    
    If ShowMessage Then
        SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Sorteo nro° " & Slot & " cancelado por falta de participantes o bien porque no se encontró ningún ganador.", FontTypeNames.FONTTYPE_SERVER)
    End If
    
    Call Logs_Security(eSecurity, eLottery, "Lottery_Cancel:: Sorteo nro° " & Slot & " cancelado por falta de participantes o bien porque no se encontró ningún ganador.")
End Sub

' Busca posibles ganadores que esten ONLINE y que tengan espacio en la cuenta.
Private Function Searching_Winner(ByVal Slot As Integer) As Integer
    
    Dim NotWin As Boolean
    Dim Chances As Integer
    Dim UserWin As Integer
    Dim tUser As Integer
    
    With Lottery(Slot)
        While NotWin = False And Chances <= LOTTERY_MAX_CHANCES
            Chances = Chances + 1

            UserWin = RandomNumber(1, .CharLast)
            
            tUser = NameIndex(.Chars(UserWin))
            
            If tUser > 0 Then
                If Not (UserList(tUser).Account.CharsAmount) = ACCOUNT_MAX_CHARS Then
                    Searching_Winner = tUser
                    Exit Function
                End If
            End If
        Wend
     End With
     
End Function

' El sorteo finaliza y el personaje es entregado.
Private Sub Lottery_Finish(ByVal Slot As Integer)

    With Lottery(Slot)

        ' Si no hay participantes no se sortea y se cancela
        If .CharLast = 0 Then
            Call Lottery_Cancel(Slot, True)
            Exit Sub

        End If

        Dim Chars(0) As String

        Dim tUser    As String
            
        tUser = Searching_Winner(Slot)
            
        If tUser <= 0 Then
            Call Lottery_Cancel(Slot, False)
            Exit Sub

        End If

        Chars(0) = .PrizeChar
    
        Call Mercader_UpdateCharsAccount(tUser, Chars, False)
        Call SaveDataAccount(tUser, UserList(tUser).Account.Email, UserList(tUser).IpAddress)
        Call WriteLoggedAccount(tUser, UserList(tUser).Account.Chars)
            
        Call Logs_Security(eSecurity, eLottery, "Lottery_Finish:: Sorteo nro° " & Slot & " concretado y premios otorgados al personaje " & UserList(tUser).Name & ".")
        
        Call WriteConsoleMsg(tUser, "¡Has ganado el sorteo! Que buena racha has tenido. Esperamos que hayas disfrutado del sorteo y te esperamos en el próximo", FontTypeNames.FONTTYPE_ANGEL)
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡El usuario " & UserList(tUser).Name & " se ha llevado el sorteo nro°" & Slot & " ¡Felicitaciones!", FontTypeNames.FONTTYPE_ANGEL))

    End With
    
    Call Lottery_Cancel(Slot, False)

End Sub

