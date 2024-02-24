Attribute VB_Name = "mShop"
Option Explicit

Public Const MAX_TRANSACCION As Byte = 100

Public Type tShopWaiting
    Email As String
    Promotion As Byte
    Bank As String
End Type

Public Type tShop
    Name As String
    Gld As Long
    Dsp As Long
    Desc As String
    ObjIndex As Integer
    ObjAmount As Integer
    Points As Integer
    
End Type

Public Shop() As tShop
Public ShopWaiting(MAX_TRANSACCION) As tShopWaiting
Public ShopLast As Integer

Public Type tShopChars
    Name As String
    Elv As Byte
    Constitucion As Byte
    Class As Byte
    Raze As Byte
    Head As Integer
    Hp As Integer
    Man As Integer
    Dsp As Integer
    Porc As Byte
End Type

Public ShopCharLast As Integer
Public ShopChars() As tShopChars

Public Sub Shop_Load_Chars_Index(ByRef Char As tShopChars)
    
    Dim ManagerChar As clsIniManager, FilePath_Char As String
    Set ManagerChar = New clsIniManager
    
    Dim Exp As Long, Elu As Long
    
    With Char
        Set ManagerChar = New clsIniManager
        FilePath_Char = CharPath & .Name & ".chr"
            
        ManagerChar.Initialize FilePath_Char
                
        .Elv = val(ManagerChar.GetValue("STATS", "ELV"))
                  
        If .Elv <> STAT_MAXELV Then
            Exp = val(ManagerChar.GetValue("STATS", "EXP"))
            Elu = val(ManagerChar.GetValue("STATS", "ELU"))
            .Porc = Int(Exp) * CDbl(100) / CDbl(Elu)
        End If
        
        .Head = val(ManagerChar.GetValue("INIT", "HEAD"))
        .Class = val(ManagerChar.GetValue("INIT", "CLASE"))
        .Raze = val(ManagerChar.GetValue("INIT", "RAZA"))
        .Hp = val(ManagerChar.GetValue("STATS", "MAXHP"))
        .Man = val(ManagerChar.GetValue("STATS", "MAXMAN"))
        
        Set ManagerChar = Nothing

    End With

End Sub

Public Sub Shop_Load_Chars()
        '<EhHeader>
        On Error GoTo Shop_Load_Chars_Err
        '</EhHeader>
    
        Dim A As Long
        Dim FilePath As String, FilePath_Char As String
        Dim Manager As clsIniManager, ManagerChar As clsIniManager
        Dim Temp As String
        Dim Exp As Long, Elu As Long
        
100     FilePath = DatPath & "CHARS.ini"
    
102     Set Manager = New clsIniManager

104     Manager.Initialize FilePath
    
106     ShopCharLast = val(Manager.GetValue("INIT", "LAST"))

108     ReDim ShopChars(0 To ShopCharLast) As tShopChars
    
    
110     For A = 1 To ShopCharLast
112         With ShopChars(A)
                  Temp = Manager.GetValue("CHARS", A)
114             .Name = ReadField(1, Temp, 45)
                  .Dsp = val(ReadField(2, Temp, 45))
                  
116             Set ManagerChar = New clsIniManager
118             FilePath_Char = CharPath & .Name & ".chr"
            
120             ManagerChar.Initialize FilePath_Char
                
                  .Elv = val(ManagerChar.GetValue("STATS", "ELV"))
                  
                  If .Elv <> STAT_MAXELV Then
                        Exp = val(ManagerChar.GetValue("STATS", "EXP"))
                        Elu = val(ManagerChar.GetValue("STATS", "ELU"))
                        .Porc = Int(Exp) * CDbl(100) / CDbl(Elu)
                 Else
                        .Porc = A
                  End If
                  
                .Head = val(ManagerChar.GetValue("INIT", "HEAD"))
122             .Class = val(ManagerChar.GetValue("INIT", "CLASE"))
124             .Raze = val(ManagerChar.GetValue("INIT", "RAZA"))
126             .Hp = val(ManagerChar.GetValue("STATS", "MAXHP"))
128             .Man = val(ManagerChar.GetValue("STATS", "MAXMAN"))
            
130             Set ManagerChar = Nothing
            End With
    
132     Next A

134     Set Manager = Nothing
    
        '<EhFooter>
        Exit Sub

Shop_Load_Chars_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mShop.Shop_Load_Chars " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub SaveShopChars()

        '<EhHeader>
        On Error GoTo Shop_Save_Char_Err

        '</EhHeader>
    
        Dim FilePath As String
        Dim A As Long
        
        Dim Manager  As clsIniManager

100     FilePath = DatPath & "CHARS.ini"
    
102     Set Manager = New clsIniManager
        
          Call Manager.ChangeValue("INIT", "LAST", CStr(ShopCharLast))
          
          For A = 1 To ShopCharLast
112         With ShopChars(A)
                Call Manager.ChangeValue("CHARS", CStr(A), .Name & "-" & CStr(.Dsp))
    
            End With
        Next A
        
        Manager.DumpFile FilePath
    
134     Set Manager = Nothing
    
        '<EhFooter>
        Exit Sub

Shop_Save_Char_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.mShop.Shop_Save_Char " & "at line " & Erl
        
        

        '</EhFooter>
End Sub

Public Sub Shop_Load()
        '<EhHeader>
        On Error GoTo Shop_Load_Err
        '</EhHeader>
        Dim A As Long
        Dim FilePath As String
        Dim Manager As clsIniManager
        Dim Temp As String
    
100     FilePath = DatPath & "SHOP.ini"
    
102     Set Manager = New clsIniManager
    
104     Manager.Initialize FilePath
    
106     ShopLast = val(Manager.GetValue("INIT", "LAST"))

108     ReDim Shop(1 To ShopLast) As tShop
    
110     For A = 1 To ShopLast
112         With Shop(A)
114             .Name = Manager.GetValue(A, "NAME")
116             .Desc = Manager.GetValue(A, "DESC")
118             .Gld = val(Manager.GetValue(A, "GLD"))
120             .Dsp = val(Manager.GetValue(A, "DSP"))
            
122             Temp = Manager.GetValue(A, "OBJINDEX")
124             .ObjIndex = val(ReadField(1, Temp, 45))
126             .ObjAmount = val(ReadField(2, Temp, 45))
            
128             .Points = val(Manager.GetValue(A, "POINTS"))
            End With
    
130     Next A
    
132     Set Manager = Nothing
    
134     Call DataServer_Generate_Shop
        '<EhFooter>
        Exit Sub

Shop_Load_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mShop.Shop_Load " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

' Transacciones que el Admin debe habilitar

Private Function Transaccion_FreeSlot() As Integer
        '<EhHeader>
        On Error GoTo Transaccion_FreeSlot_Err
        '</EhHeader>

        Dim A As Long
    
        Transaccion_FreeSlot = -1
        
100     For A = 0 To MAX_TRANSACCION

102         If ShopWaiting(A).Email = vbNullString Then
104             Transaccion_FreeSlot = A
                Exit Function

            End If

106     Next A

        '<EhFooter>
        Exit Function

Transaccion_FreeSlot_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mShop.Transaccion_FreeSlot " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Sub Transaccion_Add(ByVal UserIndex As Integer, ByRef Waiting As tShopWaiting)
        '<EhHeader>
        On Error GoTo Transaccion_Add_Err
        '</EhHeader>
    
        Dim Slot As Integer
    
100     Slot = Transaccion_FreeSlot
    
102     If Slot = -1 Then
104         Call WriteErrorMsg(UserIndex, "Ocurrió un error grave con la transacción. Espera unos momentos y vuelve a intentar.")
            Exit Sub
        End If
        
        'If Not AsciiValidos(Waiting.Bank) Then
             'Call WriteErrorMsg(UserIndex, "Evita números, simbolos y tildes. Escribe el nombre de la persona (la que figura en la cuenta que ingreso el pago)")
             'Exit Sub
       ' End If
        
        If Not CheckMailString(Waiting.Email) Then
             Call WriteErrorMsg(UserIndex, "Email inválido. Corrobora que no tenga espacios ni caracteres inválidos.")
             Exit Sub
        End If

        If Waiting.Promotion < 0 Or Waiting.Promotion > 5 Then Exit Sub
        
106     ShopWaiting(Slot) = Waiting
108     FrmShop.lstShop.AddItem Slot & "|" & ShopWaiting(Slot).Email & "|" & ShopWaiting(Slot).Promotion
          
110     Call WriteErrorMsg(UserIndex, "¡Has confirmado una nueva transacción. Te pedimos que si aún no has enviado el dinero, lo hagas de manera inmediata así la validación es en los próximos minutos...")
        '<EhFooter>
        Exit Sub

Transaccion_Add_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mShop.Transaccion_Add " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub Transaccion_Accept(ByVal Slot As Byte)
        '<EhHeader>
        On Error GoTo Transaccion_Accept_Err
        '</EhHeader>
    
        ' // Le damos lo que pagó
        Dim tUser As Integer
        Dim FilePath As String
        Dim TempDSP As Long
        
100     tUser = CheckEmailLogged(ShopWaiting(Slot).Email)
        
        If tUser > 0 Then
            
            With UserList(tUser)
                .Account.Eldhir = .Account.Eldhir + Cantidad_Dsp(ShopWaiting(Slot).Promotion)
                
                Call WriteConsoleMsg(tUser, "¡Has recibido la suma de " & Cantidad_Dsp(ShopWaiting(Slot).Promotion) & " DSP. ¡Disfrutalas!", FontTypeNames.FONTTYPE_INFOGREEN)
                Call WriteAccountInfo(tUser)
            End With

        Else
            FilePath = AccountPath & ShopWaiting(Slot).Email & ".acc"
            
            If Not FileExist(FilePath, vbArchive) Then
                Call MsgBox("¡No existe la cuenta!")
                Exit Sub
            End If
            
            TempDSP = val(GetVar(FilePath, "INIT", "ELDHIR"))
            
            Call WriteVar(FilePath, "INIT", "ELDHIR", CStr(TempDSP + Cantidad_Dsp(ShopWaiting(Slot).Promotion)))
        End If
        
        ' // Generamos LOG
102     Call Logs_Security(eLog.eGeneral, eLogSecurity.eShop, "Carga de DSP: " & Cantidad_Dsp(ShopWaiting(Slot).Promotion) & " DSP en la cuenta de " & ShopWaiting(Slot).Email & ".")
    
        ' Quitamos de la lista
104     Call Transaccion_Clear(Slot)

        FrmShop.lblRef.Caption = "Carga de DSP: " & Cantidad_Dsp(ShopWaiting(Slot).Promotion) & " DSP en la cuenta de " & ShopWaiting(Slot).Email & "."
        '<EhFooter>
        Exit Sub

Transaccion_Accept_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mShop.Transaccion_Accept " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub Transaccion_Clear(ByVal Slot As Byte)
        '<EhHeader>
        On Error GoTo Transaccion_Clear_Err
        '</EhHeader>
    
        Dim NullShop As tShopWaiting
    
100     ShopWaiting(Slot) = NullShop
    
102     FrmShop.lstShop.RemoveItem FrmShop.lstShop.ListIndex
    
        '<EhFooter>
        Exit Sub

Transaccion_Clear_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mShop.Transaccion_Clear " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Function Cantidad_Dsp(ByVal Promotion As Byte) As Long

    Select Case Promotion

        Case 0
            Cantidad_Dsp = 250

        Case 1
            Cantidad_Dsp = 500

        Case 2
            Cantidad_Dsp = 1000

        Case 3
            Cantidad_Dsp = 2000
            
        Case 4
            Cantidad_Dsp = 4000
            
        Case 5
            Cantidad_Dsp = 8000
            
        Case Else
            Cantidad_Dsp = 0
            
    End Select

End Function

Public Function ApplyDiscount(ByVal UserIndex As Integer, ByVal Price As Long)
        '<EhHeader>
        On Error GoTo ApplyDiscount_Err
        '</EhHeader>
100     Select Case UserList(UserIndex).Account.Premium
    
            Case 0
102             ApplyDiscount = Price
104         Case 1
106             ApplyDiscount = Price - (Price * 0.05)
108         Case 2
110             ApplyDiscount = Price - (Price * 0.07)
112         Case 3
114             ApplyDiscount = Price - (Price * 0.1)
        End Select
        '<EhFooter>
        Exit Function

ApplyDiscount_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mShop.ApplyDiscount " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function
Public Sub ConfirmItem(ByVal UserIndex As Integer, _
                       ByVal ID As Integer, _
                       ByVal SelectedValue As Byte)

        '<EhHeader>
        On Error GoTo ConfirmItem_Err

        '</EhHeader>
    
        Dim Obj    As Obj

        Dim Random As Byte

        Dim Sound  As Integer
    
100     With Shop(ID)
            
                     Obj.Amount = .ObjAmount
         Obj.ObjIndex = .ObjIndex

            
            If Obj.ObjIndex <> 880 Then
            If SelectedValue = 0 And .Gld = 0 Then Exit Sub ' Quiere comprar por ORO, y el item no sale ORO
            If SelectedValue = 1 And .Dsp = 0 Then Exit Sub ' Quiere comprar por DSP, y el item no sale DSP
              
102       If (.Gld > UserList(UserIndex).Account.Gld) And SelectedValue = 0 Then Exit Sub
            If (.Dsp > UserList(UserIndex).Account.Eldhir) And SelectedValue = 1 Then Exit Sub
104
              End If
              

110         If Obj.ObjIndex = 880 Then

                ' Solicita Puntos de Torneo [CANJE]
                If .Points > UserList(UserIndex).Stats.Points Then Exit Sub
112                 UserList(UserIndex).Account.Eldhir = UserList(UserIndex).Account.Eldhir + Obj.Amount
                    UserList(UserIndex).Stats.Points = UserList(UserIndex).Stats.Points - .Points
                  
114         ElseIf Obj.ObjIndex = 9999 Then
                Exit Sub
            
116         ElseIf Obj.ObjIndex = 9998 Then

                Dim MeditationSelected As Byte

118             MeditationSelected = val(ReadField(2, .Name, Asc(" ")))
                    
                UserList(UserIndex).MeditationUser(MeditationSelected) = 1
120             Call mMeditations.Meditation_Select(UserIndex, MeditationSelected)
                
                Call Logs_Security(eLog.eGeneral, eLogSecurity.eShop, "Usuario: " & UserList(UserIndex).Name & "» COMPRA DE MEDITACION: " & MeditationSelected)
            Else

124             If Not MeterItemEnInventario(UserIndex, Obj, True) Then Exit Sub
                  
                Call Logs_Security(eLog.eGeneral, eLogSecurity.eShop, "Usuario: " & UserList(UserIndex).Name & "» COMPRA DE ITEM: " & ObjData(Obj.ObjIndex).Name & " (x" & Obj.Amount & ")")

            End If
            
            
            If SelectedValue = 0 Then
126             UserList(UserIndex).Account.Gld = UserList(UserIndex).Account.Gld - ApplyDiscount(UserIndex, .Gld)
128         ElseIf SelectedValue = 1 Then
130             UserList(UserIndex).Account.Eldhir = UserList(UserIndex).Account.Eldhir - ApplyDiscount(UserIndex, .Dsp)
            End If
              
132         Random = RandomNumber(1, 100)
        
134         If Random <= 25 Then
136             Sound = eSound.sChestDrop1
138         ElseIf Random <= 50 Then
140             Sound = eSound.sChestDrop2
            Else
142             Sound = eSound.sChestDrop3

            End If
        
144         Call SendData(SendTarget.ToOne, UserIndex, PrepareMessagePlayEffect(Sound, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
146         Call WriteAccountInfo(UserIndex)
148
    
        End With

        '<EhFooter>
        Exit Sub

ConfirmItem_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.mShop.ConfirmItem " & "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub ConfirmTier(ByVal UserIndex As Integer, ByVal Tier As Byte)

        '<EhHeader>
        On Error GoTo ConfirmTier_Err

        '</EhHeader>
100     If Tier <= 0 Or Tier > 3 Then Exit Sub
    
        Dim Price As Long
    
102     Select Case Tier
    
            Case 1
104             Price = 100

106         Case 2
108             Price = 250

110         Case 3
112             Price = 450

        End Select
    
114     With UserList(UserIndex)

116         If .Account.Eldhir < Price Then Exit Sub

118         If .Account.Premium > 0 Then
                  If .Account.Premium < Tier Then
                        Call WriteConsoleMsg(UserIndex, "¡Debes esperar a que se venca el Tier Inferior o hablar con un administrador para realizar un Upgrade!", FontTypeNames.FONTTYPE_INFORED)
                        Call WriteErrorMsg(UserIndex, "¡Debes esperar a que se venca el Tier Inferior o hablar con un administrador para realizar un Upgrade!")
                        Exit Sub
                  End If
                  
                  
120             .Account.DatePremium = DateAdd("m", 1, .Account.DatePremium)

124             Call Logs_Security(eLog.eGeneral, eLogSecurity.eShop, "Usuario: " & UserList(UserIndex).Name & "» ACTUALIZA A TIER:  " & Tier)
            Else
126             .Account.DatePremium = DateAdd("m", 1, Now)
            
130             Call Logs_Security(eLog.eGeneral, eLogSecurity.eShop, "Usuario: " & UserList(UserIndex).Name & "» COMPRA DE TIER:  " & Tier)

            End If
            
                   .Account.Eldhir = .Account.Eldhir - Price
                   
            .Account.Premium = Tier
            
            Call WriteErrorMsg(UserIndex, "Tiempo PREMIUM actualizado hasta " & .Account.DatePremium & ".")
122         Call WriteConsoleMsg(UserIndex, "Tiempo PREMIUM actualizado hasta " & .Account.DatePremium & ".", FontTypeNames.FONTTYPE_USERPREMIUM)
         
            Call SaveDataAccount(UserIndex, .Account.Email, .IpAddress)
132         Call SendData(SendTarget.ToOne, UserIndex, PrepareMessagePlayEffect(eSound.sVictory2, .Pos.X, .Pos.Y))

        End With
          
134     Call WriteAccountInfo(UserIndex)
        '<EhFooter>
        Exit Sub

ConfirmTier_Err:
        LogError Err.description & vbCrLf & "in ServidorArgentum.mShop.ConfirmTier " & "at line " & Erl

        

        '</EhFooter>
End Sub

Public Sub ConfirmChar(ByVal UserIndex As Integer, ByVal ID As Byte)
        '<EhHeader>
        On Error GoTo ConfirmChar_Err
        '</EhHeader>
    
100     If UserList(UserIndex).Account.Eldhir < ShopChars(ID).Dsp Then Exit Sub
    
        Dim NullShopChar As tShopChars

        Dim Chars(0)     As String

102     If ShopChars(ID).Elv = 0 Then
104         Call WriteErrorMsg(UserIndex, "¡Parece que el personaje ha sido vendido!")
106         Call WriteShopChars(UserIndex)
            Exit Sub

        End If

108     If (UserList(UserIndex).Account.CharsAmount) = ACCOUNT_MAX_CHARS Then
110         Call WriteErrorMsg(UserIndex, "No tienes espacio para recibir nuevos personajes.")
            Exit Sub

        End If
    
112     UserList(UserIndex).Account.Eldhir = UserList(UserIndex).Account.Eldhir - ShopChars(ID).Dsp
    
114     Chars(0) = ShopChars(ID).Name
    
116     Call Logs_Security(eLog.eGeneral, eLogSecurity.eShop, "Cuenta: " & UserList(UserIndex).Account.Email & "» COMPRA EL PERSONAJE:  " & ShopChars(ID).Name & " a " & ShopChars(ID).Dsp & " DSP")
    
118     Call Mercader_UpdateCharsAccount(UserIndex, Chars, False)
          Call UpdateShopChars(ID)
          
          
124     Call mAccount.SaveDataAccount(UserIndex, UserList(UserIndex).Account.Email, UserList(UserIndex).IpAddress)
126     Call WriteLoggedAccount(UserIndex, UserList(UserIndex).Account.Chars)

128     Call WriteShopChars(UserIndex)
          Call WriteErrorMsg(UserIndex, "¡Has comprado el personaje: " & Chars(0) & "!")
        '<EhFooter>
        Exit Sub

ConfirmChar_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mShop.ConfirmChar " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub UpdateShopChars(ByVal Slot As Byte)
        '<EhHeader>
        On Error GoTo UpdateShopChars_Err
        '</EhHeader>

        Dim A                       As Long
        Dim Temp() As tShopChars
100     ReDim Temp(1 To ShopCharLast) As tShopChars
            
        ' Copia para no repetir
102     For A = 1 To ShopCharLast
104         Temp(A) = ShopChars(A)
106     Next A
    
108     ShopCharLast = ShopCharLast - 1
    
        ' Movemos +1 a los usuarios desde esta posición.
110     For A = Slot To ShopCharLast
112         ShopChars(A) = Temp(A + 1)
114     Next A
    
    
    
116     Call SaveShopChars

        '<EhFooter>
        Exit Sub

UpdateShopChars_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mShop.UpdateShopChars " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub Shop_CharAdd(ByRef Char As tShopChars)
        '<EhHeader>
        On Error GoTo Shop_CharAdd_Err
        '</EhHeader>
    
100     ShopCharLast = ShopCharLast + 1
    
102     ReDim Preserve ShopChars(0 To ShopCharLast) As tShopChars
          
          Call Shop_Load_Chars_Index(Char)
104     ShopChars(ShopCharLast) = Char
106     Call SaveShopChars

        '<EhFooter>
        Exit Sub

Shop_CharAdd_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mShop.Shop_CharAdd " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub
