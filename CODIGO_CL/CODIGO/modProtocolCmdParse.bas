Attribute VB_Name = "ProtocolCmdParse"
'Exodo Online
'
'Copyright (C) 2006 Juan Martín Sotuyo Dodero (Maraxus)
'Copyright (C) 2006 Alejandro Santos (AlejoLp)

'
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
'
'Exodo Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'

Option Explicit

#If ClienteGM = 1 Then

Public Enum eSearchData
    eMac = 1
    eDisk = 2
    eIpAddress = 3
End Enum

#End If

Public Enum eNumber_Types

    ent_Byte
    ent_Integer
    ent_Long
    ent_Trigger

End Enum

Public Sub AuxWriteWhisper(ByVal UserName As String, ByVal Mensaje As String)

    '***************************************************
    'Author: Unknown
    'Last Modification: 03/12/2010
    '03/12/2010: Enanoh - Ahora se envía el nick en vez del index del usuario.
    '***************************************************
    If LenB(UserName) = 0 Then Exit Sub
    
    If (InStrB(UserName, "+") <> 0) Then
        UserName = Replace$(UserName, "+", " ")
    End If
    
    UserName = UCase$(UserName)
    
    Call WriteWhisper(UserName, Mensaje)
    
End Sub

''
' Interpreta, valida y ejecuta el comando ingresado .
'
' @param    RawCommand El comando en version String
' @remarks  None Known.

Public Sub ParseUserCommand(ByVal RawCommand As String)

    '***************************************************
    'Author: Alejandro Santos (AlejoLp)
    'Last Modification: 16/11/2009
    'Interpreta, valida y ejecuta el comando ingresado
    '26/03/2009: ZaMa - Flexibilizo la cantidad de parametros de /nene,  /onlinemap y /telep
    '16/11/2009: ZaMa - Ahora el /ct admite radio
    '18/09/2010: ZaMa - Agrego el comando /mod username vida xxx
    '***************************************************
    Dim TmpArgos()         As String
    
    Dim Comando            As String

    Dim ArgumentosAll()    As String

    Dim ArgumentosRaw      As String

    Dim Argumentos2()      As String

    Dim Argumentos3()      As String

    Dim Argumentos4()      As String

    Dim CantidadArgumentos As Long

    Dim notNullArguments   As Boolean
    
    Dim tmpArr()           As String

    Dim tmpInt             As Integer
    
    ' TmpArgs: Un array de a lo sumo dos elementos,
    ' el primero es el comando (hasta el primer espacio)
    ' y el segundo elemento es el resto. Si no hay argumentos
    ' devuelve un array de un solo elemento
    TmpArgos = Split(RawCommand, " ", 2)
    
    Comando = Trim$(UCase$(TmpArgos(0)))
    
    If UBound(TmpArgos) > 0 Then
        ' El string en crudo que este despues del primer espacio
        ArgumentosRaw = TmpArgos(1)
        
        'veo que los argumentos no sean nulos
        notNullArguments = LenB(Trim$(ArgumentosRaw))
        
        ' Un array separado por blancos, con tantos elementos como
        ' se pueda
        ArgumentosAll = Split(TmpArgos(1), " ")
        
        ' Cantidad de argumentos. En ESTE PUNTO el minimo es 1
        CantidadArgumentos = UBound(ArgumentosAll) + 1
        
        ' Los siguientes arrays tienen A LO SUMO, COMO MAXIMO
        ' 2, 3 y 4 elementos respectivamente. Eso significa
        ' que pueden tener menos, por lo que es imperativo
        ' preguntar por CantidadArgumentos.
        
        Argumentos2 = Split(TmpArgos(1), " ", 2)
        Argumentos3 = Split(TmpArgos(1), " ", 3)
        Argumentos4 = Split(TmpArgos(1), " ", 4)
    Else
        CantidadArgumentos = 0

    End If
    
    ' Sacar cartel APESTA!! (y es ilógico, estás diciendo una pausa/espacio  :rolleyes: )
    If Comando = "" Then Comando = " "
    
    If Left$(Comando, 1) = "!" Then
        Comando = Replace$(Comando, "!", "/")

    End If
    
    If Left$(Comando, 1) = "/" Then
        ' Comando normal
        
        Select Case Comando

            Case "/ONLINE"
                Call WriteOnline
            
            Case "/ALQUILAR"
                Call WriteAlquilar(1)
                Call ShowConsoleMsg("¡¡RECUERDA!! Con el comando /BALANCEALQUILER podrás saber en todo momento cuanto llevas recaudado. Luego de que se termine el contrato tus items iran a tu boveda ¡¡¡¡¡¡¡Recuerda tener LUGAR DISPONIBLE!!!!!!!")
                 Call ShowConsoleMsg("¡¡RECUERDA!! Con el comando /RECLAMARGANANCIA podrás ir retirando las ganancias a diario. Sino el último día se te dará")
            Case "/BALANCEALQUILER"
                Call WriteAlquilar(2)
                
            Case "/RECLAMARGANANCIA"
                Call WriteAlquilar(3)
                
            Case "/SALIR"

                If UserParalizado Then 'Inmo

                    With FontTypes(FontTypeNames.FONTTYPE_WARNING)
                        Call ShowConsoleMsg("No puedes salir estando paralizado.", .red, .green, .blue, .bold, .italic)

                    End With

                    Exit Sub

                End If
                
                Call WriteQuit

            Case "/MEDITAR"
                
                If UserEstado = 1 Then 'Muerto

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

                    End With

                    Exit Sub

                End If

                Call RequestMeditate
        
            Case "/CONSULTA"
                Call WriteConsultation
            
            Case "/RESUCITAR"
                Call WriteResucitate
                
            Case "/CURAR"
                Call WriteHeal
                              
            Case "/EST"
                Call WriteRequestStats
            
            Case "/AYUDA"
                Call WriteHelp
                
            Case "/COMERCIAR"

                If UserEstado = 1 Then 'Muerto

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

                    End With

                    Exit Sub
                
                ElseIf Comerciando Then 'Comerciando

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("Ya estás comerciando", .red, .green, .blue, .bold, .italic)

                    End With

                    Exit Sub

                End If

                Call WriteCommerceStart
                
            Case "/BOVEDA"

                If UserEstado = 1 Then 'Muerto

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

                    End With

                    Exit Sub

                End If

                Call WriteBankStart(0)
            
            Case "/COMPARTIRNPC"

                If UserEstado = 1 Then 'Muerto

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

                    End With

                    Exit Sub

                End If
                
                Call WriteShareNpc
                
            Case "/NOCOMPARTIRNPC"

                If UserEstado = 1 Then 'Muerto

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

                    End With

                    Exit Sub

                End If
                
                Call WriteStopSharingNpc
        
            Case "/PMSG"

                'Ojo, no usar notNullArguments porque se usa el string vacio para borrar cartel.
                If CantidadArgumentos > 0 Then
                    Call WritePartyMessage(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba un mensaje.")

                End If
            
            Case "/CMSG"

                'Ojo, no usar notNullArguments porque se usa el string vacio para borrar cartel.
                If CantidadArgumentos > 0 Then
                    Call WriteGuilds_Talk(ArgumentosRaw, False)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba un mensaje.")

                End If
                
            Case "/BMSG"

                If notNullArguments Then
                    Call WriteCouncilMessage(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba un mensaje.")

                End If
            
            Case "/DESC"

                If UserEstado = 1 Then 'Muerto

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

                    End With

                    Exit Sub

                End If
                
                Call WriteChangeDescription(ArgumentosRaw)
               
            Case "/PENAS"

                If notNullArguments Then
                    Call WritePunishments(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    'Call ShowConsoleMsg("Faltan parámetros. Utilice /penas NICKNAME.")
                    Call WritePunishments(UserName)

                End If
            
            Case "/APOSTAR"

                If UserEstado = 1 Then 'Muerto

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

                    End With

                    Exit Sub

                End If

                If notNullArguments Then
                    If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Integer) Then
                        Call WriteGamble(ArgumentosRaw)
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Cantidad incorrecta. Utilice /apostar CANTIDAD.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /apostar CANTIDAD.")

                End If
                
            Case "/DENUNCIAR"

                If notNullArguments Then
                    Call WriteDenounce(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Formule su denuncia.")

                End If
                
            Case "/CHAT"

                If notNullArguments Then
                    Call WriteGMMessage(ArgumentosRaw)

                End If
                
            Case "/SHOWNAME"
                Call WriteShowName
            
                #If ClienteGM = 1 Then
                
                Case "/MODOEVENTO"
                    Call WriteChangeModoArgentum
                    
                Case "/DATA"

                    If notNullArguments Then
                        Call WriteSendDataUser(ArgumentosRaw)

                    End If
                
                Case "/MAC"

                    If notNullArguments Then
                        Call WriteSearchDataUser(eSearchData.eMac, ArgumentosRaw)

                    End If
                    
                Case "/IPADDRESS"

                    If notNullArguments Then
                        Call WriteSearchDataUser(eSearchData.eIpAddress, ArgumentosRaw)

                    End If
                
                Case "/DISK"

                    If notNullArguments Then
                        Call WriteSearchDataUser(eSearchData.eDisk, ArgumentosRaw)

                    End If

                #End If
            
            Case "/HORA"
                Call Protocol.WriteServerTime
            
            Case "/DONDE"

                If notNullArguments Then
                    Call WriteWhere(ArgumentosRaw, False)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /DONDE NICKNAME.")

                End If
                
            Case "/DONDECLAN"

                If notNullArguments Then
                    Call WriteWhere(ArgumentosRaw, True)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /DONDECLAN CLAN.")

                End If
                
            Case "/NENE"

                If notNullArguments Then
                    If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Integer) Then
                        Call WriteCreaturesInMap(ArgumentosRaw)
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Mapa incorrecto. Utilice /nene MAPA.")

                    End If

                Else
                    'Por default, toma el mapa en el que esta
                    Call WriteCreaturesInMap(UserMap)

                End If
                
            Case "/TELEP"

                If notNullArguments And CantidadArgumentos >= 4 Then
                    If ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(3), eNumber_Types.ent_Byte) Then
                        Call WriteWarpChar(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2), ArgumentosAll(3))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Valor incorrecto. Utilice /telep NICKNAME MAPA X Y.")

                    End If

                ElseIf CantidadArgumentos = 3 Then

                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Byte) Then
                        'Por defecto, si no se indica el nombre, se teletransporta el mismo usuario
                        Call WriteWarpChar("YO", ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2))
                    ElseIf ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Byte) Then
                        'Por defecto, si no se indica el mapa, se teletransporta al mismo donde esta el usuario
                        Call WriteWarpChar(ArgumentosAll(0), UserMap, ArgumentosAll(1), ArgumentosAll(2))
                    Else
                        'No uso ningun formato por defecto
                        Call ShowConsoleMsg("Valor incorrecto. Utilice /telep NICKNAME MAPA X Y.")

                    End If

                ElseIf CantidadArgumentos = 2 Then

                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Byte) Then
                        ' Por defecto, se considera que se quiere unicamente cambiar las coordenadas del usuario, en el mismo mapa
                        Call WriteWarpChar("YO", UserMap, ArgumentosAll(0), ArgumentosAll(1))
                    Else
                        'No uso ningun formato por defecto
                        Call ShowConsoleMsg("Valor incorrecto. Utilice /telep NICKNAME MAPA X Y.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /telep NICKNAME MAPA X Y.")

                End If
                
            Case "/SILENCIAR"

                If notNullArguments Then
                    Call WriteSilence(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /silenciar NICKNAME.")

                End If
                
            Case "/SHOW"

                If notNullArguments Then

                    Select Case UCase$(ArgumentosAll(0))
                        
                        Case "DENUNCIAS"
                            Call WriteShowDenouncesList

                    End Select

                End If
                
            Case "/DENUNCIAS"
                Call WriteEnableDenounces
                
            Case "/IRA"

                If notNullArguments Then
                    Call WriteGoToChar(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /ira NICKNAME.")

                End If
        
            Case "/INVISIBLE"
                Call WriteInvisible
                
            Case "/PANELGM"
                Call WriteGMPanel
                
            Case "/CARCEL"

                If notNullArguments Then
                    tmpArr = Split(ArgumentosRaw, "@")

                    If UBound(tmpArr) = 2 Then
                        If ValidNumber(tmpArr(2), eNumber_Types.ent_Byte) Then
                            Call WriteJail(tmpArr(0), tmpArr(1), tmpArr(2))
                        Else
                            'No es numerico
                            Call ShowConsoleMsg("Tiempo incorrecto. Utilice /carcel NICKNAME@MOTIVO@TIEMPO.")

                        End If

                    Else
                        'Faltan los parametros con el formato propio
                        Call ShowConsoleMsg("Formato incorrecto. Utilice /carcel NICKNAME@MOTIVO@TIEMPO.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /carcel NICKNAME@MOTIVO@TIEMPO.")

                End If
                
            Case "/RMATA"
                Call WriteKillNPC
                
            Case "/ADVERTENCIA"

                If notNullArguments Then
                    tmpArr = Split(ArgumentosRaw, "@", 2)

                    If UBound(tmpArr) = 1 Then
                        Call WriteWarnUser(tmpArr(0), tmpArr(1))
                    Else
                        'Faltan los parametros con el formato propio
                        Call ShowConsoleMsg("Formato incorrecto. Utilice /advertencia NICKNAME@MOTIVO.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /advertencia NICKNAME@MOTIVO.")

                End If
            
            Case "/INFOGM"

                If notNullArguments Then
                    Call WriteRequestCharInfo(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /INFOGM NICKNAME.")

                End If
                
            Case "/INV"

                If notNullArguments Then
                    Call WriteRequestCharInventory(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /inv NICKNAME.")

                End If
                
            Case "/BOV"

                If notNullArguments Then
                    Call WriteRequestCharBank(ArgumentosRaw, 1)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /BOV NICKNAME.")

                End If
               
            Case "/BOVCUENTA"

                If notNullArguments Then
                    Call WriteRequestCharBank(ArgumentosRaw, 2)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /BOVCUENTA NICKNAME.")

                End If
                
            Case "/REVIVIR"

                If notNullArguments Then
                    Call WriteReviveChar(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /revivir NICKNAME.")

                End If
                
            Case "/ONLINEGM"
                Call WriteOnlineGM
                
            Case "/ONLINEMAP"

                If notNullArguments Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) Then
                        Call WriteOnlineMap(ArgumentosAll(0))
                    Else
                        Call ShowConsoleMsg("Mapa incorrecto.")

                    End If

                Else
                    Call WriteOnlineMap(UserMap)

                End If
                
            Case "/PERDONCAOS"

                If notNullArguments Then
                    Call WriteForgive(ArgumentosRaw, False)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /PERDONCAOS NICKNAME.")

                End If
                
            Case "/PERDONARMADA"

                If notNullArguments Then
                    Call WriteForgive(ArgumentosRaw, True)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /PERDONARMADA NICKNAME.")

                End If
                
            Case "/ECHAR"

                If notNullArguments Then
                    Call WriteKick(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /echar NICKNAME.")

                End If
                
            Case "/EJECUTAR"

                If notNullArguments Then
                    Call WriteExecute(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /ejecutar NICKNAME.")

                End If
                
            Case "/UNBAN"

                If notNullArguments Then
                    Call WriteUnbanChar(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /unban NICKNAME.")

                End If
                
            Case "/SEGUIR"
                Call WriteNPCFollow
                
            Case "/SUM"

                If notNullArguments Then
                    Call WriteSummonChar(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /sum NICKNAME.")

                End If
                
            Case "/CC"
                Call WriteSpawnListRequest
                
            Case "/RESETINV"
                Call WriteResetNPCInventory
                
            Case "/LIMPIAR"
                Call WriteCleanWorld
                
            Case "/RMSG"

                If notNullArguments Then
                    Call WriteServerMessage(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba un mensaje.")

                End If
            
            Case "/MAPMSG"

                If notNullArguments Then
                    Call WriteMapMessage(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba un mensaje.")

                End If
                
            Case "/NICK2IP"

                If notNullArguments Then
                    Call WriteNickToIP(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /nick2ip NICKNAME.")

                End If
                
            Case "/IP2NICK"

                If notNullArguments Then
                    If validipv4str(ArgumentosRaw) Then
                        Call WriteIPToNick(str2ipv4l(ArgumentosRaw))
                    Else
                        'No es una IP
                        Call ShowConsoleMsg("IP incorrecta. Utilice /ip2nick IP.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /ip2nick IP.")

                End If
                
            Case "/CT"

                If notNullArguments And CantidadArgumentos >= 3 Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Byte) Then
                        
                        If CantidadArgumentos = 3 Then
                            Call WriteTeleportCreate(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2))
                        Else

                            If ValidNumber(ArgumentosAll(3), eNumber_Types.ent_Byte) Then
                                Call WriteTeleportCreate(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2), ArgumentosAll(3))
                            Else
                                'No es numerico
                                Call ShowConsoleMsg("Valor incorrecto. Utilice /ct MAPA X Y RADIO(Opcional).")

                            End If

                        End If

                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Valor incorrecto. Utilice /ct MAPA X Y RADIO(Opcional).")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /ct MAPA X Y RADIO(Opcional).")

                End If
                
            Case "/DT"
                Call WriteTeleportDestroy
            
            Case "/MP3"

                If notNullArguments Then

                    'elegir el mapa es opcional
                    If CantidadArgumentos = 1 Then
                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) Then
                            'eviamos un mapa nulo para que tome el del usuario.
                            Call WriteForceMIDIToMap(ArgumentosAll(0), 0)
                        Else
                            'No es numerico
                            Call ShowConsoleMsg("MP3 incorrecto. Utilice /MP3 NUM MAPA, siendo el mapa opcional.")

                        End If

                    Else

                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) Then
                            Call WriteForceMIDIToMap(ArgumentosAll(0), ArgumentosAll(1))
                        Else
                            'No es numerico
                            Call ShowConsoleMsg("MP3 incorrecto. Utilice /MP3 NUM MAPA, siendo el mapa opcional.")

                        End If

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Utilice /forcemidimap MIDI MAPA, siendo el mapa opcional.")

                End If
                
            Case "/WAV"

                If notNullArguments Then

                    'elegir la posicion es opcional
                    If CantidadArgumentos = 1 Then
                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) Then
                            'eviamos una posicion nula para que tome la del usuario.
                            Call WriteForceWAVEToMap(ArgumentosAll(0), 0, 0, 0)
                        Else
                            'No es numerico
                            Call ShowConsoleMsg("Utilice /WAV NUM MAP X Y, siendo los últimos 3 opcionales.")

                        End If

                    ElseIf CantidadArgumentos = 4 Then

                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(3), eNumber_Types.ent_Byte) Then
                            Call WriteForceWAVEToMap(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2), ArgumentosAll(3))
                        Else
                            'No es numerico
                            Call ShowConsoleMsg("Utilice /WAV NUM MAP X Y, siendo los últimos 3 opcionales.")

                        End If

                    Else
                        'Avisar que falta el parametro
                        Call ShowConsoleMsg("Utilice /WAV WAV NUM X Y, siendo los últimos 3 opcionales.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Utilice /forcewavmap WAV MAP X Y, siendo los últimos 3 opcionales.")

                End If
                
            Case "/REALMSG"

                If notNullArguments Then
                    Call WriteRoyalArmyMessage(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba un mensaje.")

                End If
                 
            Case "/CAOSMSG"

                If notNullArguments Then
                    Call WriteChaosLegionMessage(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba un mensaje.")

                End If

            Case "/TALKAS"

                If notNullArguments Then
                    Call WriteTalkAsNPC(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba un mensaje.")

                End If
        
            Case "/MASSDEST"
                Call WriteDestroyAllItemsInArea
    
            Case "/CONSEJO"

                If notNullArguments Then
                    Call WriteAcceptRoyalCouncilMember(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /CONSEJO NICKNAME.")

                End If
                
            Case "/CONCILIO"

                If notNullArguments Then
                    Call WriteAcceptChaosCouncilMember(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /CONCILIO NICKNAME.")

                End If
                
            Case "/PISO"
                Call WriteItemsInTheFloor
                
            Case "/KICKCONSE"

                If notNullArguments Then
                    Call WriteCouncilKick(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /kickconse NICKNAME.")

                End If
                
            Case "/TRIGGER"

                If notNullArguments Then
                    If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Trigger) Then
                        Call WriteSetTrigger(ArgumentosRaw)
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Numero incorrecto. Utilice /trigger NUMERO.")

                    End If

                Else
                    'Version sin parametro
                    Call WriteAskTrigger

                End If
                
            Case "/BANIPLIST"
                Call WriteBannedIPList
                
            Case "/BANIPRELOAD"
                Call WriteBannedIPReload
                
            Case "/BANIP"

                If CantidadArgumentos >= 2 Then
                    If validipv4str(ArgumentosAll(0)) Then
                        Call WriteBanIP(True, str2ipv4l(ArgumentosAll(0)), vbNullString, Right$(ArgumentosRaw, Len(ArgumentosRaw) - Len(ArgumentosAll(0)) - 1))
                    Else
                        'No es una IP, es un nick
                        Call WriteBanIP(False, str2ipv4l("0.0.0.0"), ArgumentosAll(0), Right$(ArgumentosRaw, Len(ArgumentosRaw) - Len(ArgumentosAll(0)) - 1))

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /banip IP motivo o /banip nick motivo.")

                End If
                
            Case "/UNBANIP"

                If notNullArguments Then
                    If validipv4str(ArgumentosRaw) Then
                        Call WriteUnbanIP(str2ipv4l(ArgumentosRaw))
                    Else
                        'No es una IP
                        Call ShowConsoleMsg("IP incorrecta. Utilice /unbanip IP.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /unbanip IP.")

                End If
                
            Case "/FIANZA"

                If notNullArguments Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Long) Then
                        Call WriteFianza(ArgumentosAll(0))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Objeto incorrecto. Utilice /FIANZA CANTIDAD.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /FIANZA CANTIDAD.")

                End If
                
            'Case "/HOGAR"
                'Call WriteHome
                
            Case "/CI"

                If notNullArguments Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Long) Then
                        Call WriteCreateItem(ArgumentosAll(0))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Objeto incorrecto. Utilice /ci OBJETO.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /ci OBJETO.")

                End If
                
            Case "/DEST"
                Call WriteDestroyItems
                
            Case "/NOCAOS"

                If notNullArguments Then
                    Call WriteChaosLegionKick(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /nocaos NICKNAME.")

                End If
    
            Case "/NOREAL"

                If notNullArguments Then
                    Call WriteRoyalArmyKick(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /noreal NICKNAME.")

                End If
    
            Case "/MP3ALL"

                If notNullArguments Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) Then
                        Call WriteForceMIDIAll(ArgumentosAll(0))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Midi incorrecto. Utilice /MP3ALL NUM.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /MP3ALL NUM.")

                End If
    
            Case "/WAVALL"

                If notNullArguments Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) Then
                        Call WriteForceWAVEAll(ArgumentosAll(0))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Wav incorrecto. Utilice /WAVALL NUM.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /WAVALL NUM.")

                End If
                
            Case "/BLOQ"
                Call WriteTileBlockedToggle
                
            Case "/MATA"
                Call WriteKillNPCNoRespawn
        
            Case "/MASSKILL"
                Call WriteKillAllNearbyNPCs
                
            Case "/LASTIP"

                If notNullArguments Then
                    Call WriteLastIP(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /lastip NICKNAME.")

                End If
                
            Case "/SMSG"

                If notNullArguments Then
                    Call WriteSystemMessage(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba un mensaje.")

                End If
                
            Case "/ACC"

                If notNullArguments Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) Then
                        Call WriteCreateNPC(ArgumentosAll(0))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Npc incorrecto. Utilice /acc NPC.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /acc NPC.")

                End If
                
            Case "/RACC"

                If notNullArguments Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) Then
                        Call WriteCreateNPCWithRespawn(ArgumentosAll(0))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Npc incorrecto. Utilice /racc NPC.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /racc NPC.")

                End If
        
            Case "/HABILITAR"
                Call WriteServerOpenToUsersToggle
            
            Case "/APAGAR"
                Call WriteTurnOffServer
                
            Case "/CONDEN"

                If notNullArguments Then
                    Call WriteTurnCriminal(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /conden NICKNAME.")

                End If
                
            Case "/RAJAR"

                If notNullArguments Then
                    Call WriteResetFactions(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /rajar NICKNAME.")

                End If
                
            Case "/CREARPRETORIANOS"
            
                If CantidadArgumentos = 3 Then
                    
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Byte) Then
                       
                        Call WriteCreatePretorianClan(Val(ArgumentosAll(0)), Val(ArgumentosAll(1)), Val(ArgumentosAll(2)))
                    Else
                        'Faltan o sobran los parametros con el formato propio
                        Call ShowConsoleMsg("Formato incorrecto. Utilice /CrearPretorianos MAPA X Y.")

                    End If
                    
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /CrearPretorianos MAPA X Y.")

                End If
                
            Case "/ELIMINARPRETORIANOS"
            
                If CantidadArgumentos = 1 Then
                    
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) Then
                       
                        Call WriteDeletePretorianClan(Val(ArgumentosAll(0)))
                    Else
                        'Faltan o sobran los parametros con el formato propio
                        Call ShowConsoleMsg("Formato incorrecto. Utilice /EliminarPretorianos MAPA.")

                    End If
                    
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /EliminarPretorianos MAPA.")

                End If
                
            Case "/CR"

                If CantidadArgumentos = 1 Then
                    
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) Then
                       
                        Call WriteCountDown(Val(ArgumentosAll(0)), False)
                    Else
                        'Faltan o sobran los parametros con el formato propio
                        Call ShowConsoleMsg("Formato incorrecto. Utilice /CR NUMERO.")

                    End If
                    
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /CR NUMERO.")

                End If
                
            Case "/CRMAP"

                If CantidadArgumentos = 1 Then
                    
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) Then
                       
                        Call WriteCountDown(Val(ArgumentosAll(0)), True)
                    Else
                        'Faltan o sobran los parametros con el formato propio
                        Call ShowConsoleMsg("Formato incorrecto. Utilice /CRMAP NUMERO.")

                    End If
                    
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /CRMAP NUMERO.")

                End If
            
            Case "/DOBACKUP"
                Call WriteDoBackup
                
            Case "/GUARDAMAPA"
                Call WriteSaveMap
                
            Case "/MODMAPINFO" ' PK, BACKUP

                If CantidadArgumentos > 1 Then

                    Select Case UCase$(ArgumentosAll(0))

                        Case "PK" ' "/MODMAPINFO PK"
                            Call WriteChangeMapInfoPK(ArgumentosAll(1) = "1")
                            
                        Case "NIVELMIN" ' /MODMAPINFO NIVELMIN
                            Call WriteChangeMapInfoLvl(ArgumentosAll(1))
                            
                        Case "BACKUP" ' "/MODMAPINFO BACKUP"
                            Call WriteChangeMapInfoBackup(ArgumentosAll(1) = "1")
                        
                        Case "RESTRINGIR" '/MODMAPINFO RESTRINGIR
                            Call WriteChangeMapInfoRestricted(ArgumentosAll(1))
                        
                        Case "MAGIASINEFECTO" '/MODMAPINFO MAGIASINEFECTO
                            Call WriteChangeMapInfoNoMagic(ArgumentosAll(1) = "1")
                        
                        Case "INVISINEFECTO" '/MODMAPINFO INVISINEFECTO
                            Call WriteChangeMapInfoNoInvi(ArgumentosAll(1) = "1")
                        
                        Case "RESUSINEFECTO" '/MODMAPINFO RESUSINEFECTO
                            Call WriteChangeMapInfoNoResu(ArgumentosAll(1) = "1")
                        
                        Case "TERRENO" '/MODMAPINFO TERRENO
                            Call WriteChangeMapInfoLand(ArgumentosAll(1))
                        
                        Case "ZONA" '/MODMAPINFO ZONA
                            Call WriteChangeMapInfoZone(ArgumentosAll(1))
                            
                        Case "ROBONPC" '/MODMAPINFO ROBONPC
                            Call WriteChangeMapInfoStealNpc(ArgumentosAll(1) = "1")
                            
                        Case "OCULTARSINEFECTO" '/MODMAPINFO OCULTARSINEFECTO
                            Call WriteChangeMapInfoNoOcultar(ArgumentosAll(1) = "1")
                            
                        Case "INVOCARSINEFECTO" '/MODMAPINFO INVOCARSINEFECTO
                            Call WriteChangeMapInfoNoInvocar(ArgumentosAll(1) = "1")
                            
                        Case "LIMPIEZA"
                            Call WriteChangeMapInfoLimpieza(Val(ArgumentosAll(1)))
                            
                        Case "CAENITEMS"
                            Call WriteChangeMapInfoItems(Val(ArgumentosAll(1)))
                            
                        Case "EXP"
                            Call WriteChangeMapInfoExp(Val(ArgumentosAll(1)))
                            
                        Case "ATAQUE"
                            Call WriteChangeMapInfoAttack(Val(ArgumentosAll(1)))

                    End Select

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parametros. Opciones: PK, BACKUP, RESTRINGIR, MAGIASINEFECTO, INVISINEFECTO, RESUSINEFECTO, TERRENO, ZONA")

                End If
                
            Case "/GRABAR"
                Call WriteSaveChars
                
            Case "/BOTDELAYSUMMON"

                If notNullArguments And CantidadArgumentos >= 1 Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Long) Then
                        Call WriteStreamerBotSetting(Val(ArgumentosAll(0)), 0, 0)
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Valor incorrecto. Utilice /BOTDELAYSUMMON DELAY")

                    End If
                    
                Else
                    'Avisar que falta el parametro
                     Call ShowConsoleMsg("Valor incorrecto. Utilice /BOTDELAYSUMMON DELAY")
                End If
        
           Case "/BOTDELAYINDEX"

                If notNullArguments And CantidadArgumentos >= 1 Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Long) Then
                        Call WriteStreamerBotSetting(0, 0, Val(ArgumentosAll(0)))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Valor incorrecto. Utilice /BOTDELAYINDEX DELAY")

                    End If
                    
                Else
                    'Avisar que falta el parametro
                     Call ShowConsoleMsg("Valor incorrecto. Utilice /BOTDELAYINDEX DELAY")
                End If
     
     
            Case "/BOTMODE"

                If notNullArguments And CantidadArgumentos >= 1 Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) Then
                        Call WriteStreamerBotSetting(0, Val(ArgumentosAll(0)), 0)
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Valor incorrecto. Utilice /BOTMODE MODE")
                        Call ShowConsoleMsg("1=ZONA SEGURA, 2=EVENTOS AUTOMATICOS, 3=RETOS,4=RETOS RAPIDOS, 5=AGITES EN INSEGURA, 6=MIXED, 7=LISTA DE SEGUIDOS", 150, 150, 150, , True)
                    End If
                    
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Valor incorrecto. Utilice /BOTMODE MODE")
                End If
                
            Case "/BOTINITIAL"

                If notNullArguments And CantidadArgumentos >= 3 Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Long) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Long) Then
                        Call WriteStreamerBotSetting(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Valor incorrecto. Utilice /BOTINITIAL DELAY MODE DELAYINDEX")

                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Valor incorrecto. Utilice /BOTINITIAL DELAY MODE DELAYINDEX")

                End If
                
            Case "/LIVE"
                Call WriteRequiredLive
                
            Case "/CHATCOLOR"

                If notNullArguments And CantidadArgumentos >= 3 Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Byte) Then
                        Call WriteChatColor(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Valor incorrecto. Utilice /chatcolor R G B.")

                    End If

                ElseIf Not notNullArguments Then    'Go back to default!
                    Call WriteChatColor(0, 255, 0)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /chatcolor R G B.")

                End If
            
            Case "/IGNORADO"
                Call WriteIgnored
                
            Case "/TORNEO"

                If notNullArguments Then
                    Call WriteParticipeEvent(ArgumentosRaw)
                Else

                    'Avisar que falta el parametro
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("Escribe /TORNEO y separado el nombre del evento que aparece en consola.", .red, .green, .blue, .bold, .italic)

                    End With

                End If

            Case "/RETOSON"
                FightOn = True
                Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_SEGURO_RETOS_ACTIVADO, 0, 255, 0, True, False, True)
                
            Case "/RETOSOFF"
                FightOn = False
                Call AddtoRichTextBox(FrmMain.RecTxt, MENSAJE_SEGURO_RETOS_DESACTIVADO, 255, 0, 0, True, False, True)

            Case "/MAPA"
                Call FrmMapa.Show(vbModeless, FrmMain)
            
            Case "/SUBASTAR"
                Call FrmSubasta.Show(, FrmMain)
                
            Case "/SIPARTY"
                Call WritePartyClient(3)
                
            Case "/RESET"
                
                If Account.Premium = 0 Then
                    Call ShowConsoleMsg("Debes tener al menos el Tier 1 para poder reiniciar tu personaje. Consulta por donaciones en /SHOP")
                Else
                    
                    If UserLvl >= 3 Then
                        If MsgBox("¿Estás seguro que deseas reiniciar tu personaje a Nivel 1? ¡¡TODO TU INVENTARIO SE REINICIARÁ! NO perderás hechizos NI tampoco el oro obtenido. ¿ACEPTAR?", vbYesNo) = vbYes Then
                            Call WriteUserEditation
    
                        End If

                    Else
                        Call ShowConsoleMsg("Solo puedes reiniciar tu personaje a partir del nivel 3.")

                    End If
                    
                End If
                

            Case "/SHOP"
                Call FrmShop.Show(, FrmMain)
                
            Case "/SKINS"
                Call WriteRequiredSkins(0, 0)
                
            Case "/STREAM"
                Call WriteModoStream
            
            Case "/STREAMLINK"

                If notNullArguments Then
                    Call WriteStreamerLink(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /STREAMERLINK LINK DE TU STREAM. Ejemplo: /STREAMERLINK https://www.twitch.tv/desteriumgame")

                End If
                
            Case "/MISIONES"
                Call WriteQuestRequired(0)
                
            'Case "/RULETA"
                'Call FrmRuleta.Show(, FrmMain)
                
            Case "/CASTILLO"
                Call WriteCastle
                
            Case "/NORTE"
                Call WriteCastle(1)
                
            Case "/SUR"
                Call WriteCastle(3)
            Case "/ESTE"
                Call WriteCastle(2)
            Case "/OESTE"
                Call WriteCastle(4)
                
            Case "/PING"
                Call WritePing
                
                #If Testeo = 1 Then

                Case "/BODY"
                    Call FrmBody.Show(, FrmMain)
                #End If
                
            Case "/INFOSUBASTA"
                Call WriteAuction_Info
            
            Case "/INVASION"

                If notNullArguments And CantidadArgumentos >= 1 Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) Then
                        Call WriteGoInvation(ArgumentosAll(0))
                    Else
                        Call ShowConsoleMsg("Valor incorrecto. Utilice '/INVASION' seguido de la invasión a la que desees ingresar.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Valor incorrecto. Utilice '/INVASION' seguido de la invasión a la que desees ingresar.")

                End If
                
            Case "/OFRECER"

                If notNullArguments And CantidadArgumentos >= 2 Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Long) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) Then
                        Call WriteAuction_Offer(ArgumentosAll(0), ArgumentosAll(1))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Valor incorrecto. Utilice /OFRECER ORO ELDHIR")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /OFRECER ORO ELDHIR")

                End If
                
            Case "/PODER"
                Call WriteWherePower
                
            Case "/SUPLICAR"
                Call WriteForgive_Faction
                
            Case "/KICK"

                If notNullArguments Then
                    Call WriteGuilds_Kick(ArgumentosRaw)
                Else

                    'Avisar que falta el parametro
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("Debes ingresar el nombre del personaje a eliminar.", .red, .green, .blue, .bold, .italic)

                    End With

                End If
                
            
            Case "/INFO"

                If notNullArguments Then
                    Call WriteRequiredStatsUser(197, ArgumentosRaw)
                Else

                    'Avisar que falta el parametro
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("Debes ingresar el nombre del personaje del que quieras saber la información.", .red, .green, .blue, .bold, .italic)

                    End With

                End If
                
            Case "/CLAN"

                If notNullArguments Then
                    Call WriteGuilds_Invitation(ArgumentosRaw, 0)
                Else

                    'Avisar que falta el parametro
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("Debes ingresar el nombre del personaje a invitar.", .red, .green, .blue, .bold, .italic)

                    End With

                End If
                
            Case "/SICLAN"

                If notNullArguments Then
                    Call WriteGuilds_Invitation(ArgumentosRaw, 1)
                Else

                    'Avisar que falta el parametro
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("Debes ingresar el nombre del Lider que te invitó a su clan", .red, .green, .blue, .bold, .italic)

                    End With

                End If
            
            Case "/ONLINECLAN"
                Call WriteGuilds_Online
                
            Case "/SALIRCLAN"
                Call WriteGuilds_Abandonate
                
            Case "/FUNDARCLAN"
                Call Guilds_FounderNew
                
            Case "/RECLAMAR"
                Call WriteRetos_RewardObj
            
            Case "/ABANDONAR"
                Call WriteAbandonateFaction
                
            Case "/BARRAS"
                ControlActivated = Not ControlActivated
                
            Case "/ENLISTAR"
                Call WriteEnlist
                
            Case "/RECOMPENSA"
                Call WriteReward
                
            Case "/TORNEOS"
                Call WriteInfoEvento
            

            Case "/RETOS"
                #If ModoBig = 1 Then
                    dockForm FrmRetos.hWnd, FrmMain.PicMenu, True
                #Else
                    Call FrmRetos.Show(, FrmMain)
                #End If

                FrmRetos.TypeFight = eTypeFight.eSend
            
            Case "/DESAFIO"
                Call WriteEntrarDesafio(0)
                
            Case "/SALIRDESAFIO"
                Call WriteEntrarDesafio(1)
                
            Case "/GLOBAL"

                If notNullArguments Then
                    Call WriteChatGlobal(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /GLOBAL DIALOGO.")

                End If
                
            Case "/DV"

                If notNullArguments Then
                    Call WriteGiveBackUser(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /DV NICK.")

                End If
                
            Case "/SETDIALOG"

                If notNullArguments Then
                    Call WriteSetDialog(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /SETDIALOG DIALOGO.")

                End If
                
            Case "/ACEPTAR"

                If notNullArguments Then
                    Call WriteAcceptFight(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /ACEPTAR NICK.")

                End If
                
            Case "/OBJ"

                If notNullArguments Then
                    Call WriteSearchObj(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /OBJ TAG.")

                End If
                
                #If ClienteGM = 1 Then
            
                Case "/PRO"

                    If notNullArguments Then
                        Call WriteSolicitaSeguridad(ArgumentosRaw, 0)

                    End If
                
                Case "/APAGA"

                    If notNullArguments Then
                        Call WriteSolicitaSeguridad(ArgumentosRaw, 3)

                    End If
            
                Case "/REINICIA"

                    If notNullArguments Then
                        Call WriteSolicitaSeguridad(ArgumentosRaw, 4)

                    End If
                
                Case "/CIERRA"

                    If notNullArguments Then
                        Call WriteSolicitaSeguridad(ArgumentosRaw, 5)

                    End If
                
                Case "/VERHD"

                    If notNullArguments Then
                        Call WriteSolicitaSeguridad(ArgumentosRaw, 6)

                    End If
                
                Case "/VERMAC"

                    If notNullArguments Then
                        Call WriteSolicitaSeguridad(ArgumentosRaw, 7)

                    End If
                
                Case "/BANHD"

                    If notNullArguments Then
                        If Copy_HD <> 0 Then
                            Call WriteSolicitaSeguridad(ArgumentosRaw & "|" & CStr(Copy_HD), 8)
                            Copy_HD = 0

                        End If
                    
                    End If
                
                Case "/BANMAC"

                    If notNullArguments Then
                        If Copy_MAC <> vbNullString Then
                            Call WriteSolicitaSeguridad(ArgumentosRaw & "|" & Copy_MAC, 9)
                            Copy_MAC = vbNullString

                        End If

                    End If
                
                Case "/UNBANHD"

                    If notNullArguments Then
                        Call WriteSolicitaSeguridad(ArgumentosRaw, 10)

                    End If
                
                Case "/UNBANMAC"

                    If notNullArguments Then
                        Call WriteSolicitaSeguridad(ArgumentosRaw, 11)

                    End If
                
                Case "/BANMAIL"

                    If notNullArguments Then
                        Call WriteSolicitaSeguridad(ArgumentosRaw, 12)

                    End If
                
                Case "/UNBANMAIL"

                    If notNullArguments Then
                        Call WriteSolicitaSeguridad(ArgumentosRaw, 13)

                    End If

                #End If
                
            Case "/GLOBALSTATUS"
                Call WriteCheckingGlobal
                
            Case "/IMPERSONAR"
                Call WriteImpersonate
                
            Case "/MIMETIZAR"
                Call WriteImitate

        End Select
        
    ElseIf Left$(Comando, 1) = "\" Then

        If UserEstado = 1 Then 'Muerto

            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

            End With

            Exit Sub

        End If

        ' Mensaje Privado
        Call AuxWriteWhisper(complexNameToSimple(mid$(Comando, 2), False), ArgumentosRaw)
        
    ElseIf Left$(Comando, 1) = "-" Then

        If UserEstado = 1 Then 'Muerto

            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

            End With

            Exit Sub

        End If

        ' Gritar
        Call WriteYell(mid$(RawCommand, 2))
        
    Else
        ' Hablar
        Call WriteTalk(RawCommand)

    End If

End Sub

''
' Show a console message.
'
' @param    Message The message to be written.
' @param    red Sets the font red color.
' @param    green Sets the font green color.
' @param    blue Sets the font blue color.
' @param    bold Sets the font bold style.
' @param    italic Sets the font italic style.

Public Sub ShowConsoleMsg(ByVal Message As String, _
                          Optional ByVal red As Integer = 255, _
                          Optional ByVal green As Integer = 255, _
                          Optional ByVal blue As Integer = 255, _
                          Optional ByVal bold As Boolean = False, _
                          Optional ByVal italic As Boolean = False)
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 01/03/07
    '
    '***************************************************
    Call AddtoRichTextBox(FrmMain.RecTxt, Message, red, green, blue, bold, italic)
End Sub

''
' Returns whether the number is correct.
'
' @param    Numero The number to be checked.
' @param    Tipo The acceptable type of number.

Public Function ValidNumber(ByVal Numero As String, _
                            ByVal Tipo As eNumber_Types) As Boolean

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 01/06/07
    '
    '***************************************************
    Dim minimo As Long

    Dim maximo As Long
    
    If Not IsNumeric(Numero) Then Exit Function
    
    Select Case Tipo

        Case eNumber_Types.ent_Byte
            minimo = 0
            maximo = 255

        Case eNumber_Types.ent_Integer
            minimo = -32768
            maximo = 32767

        Case eNumber_Types.ent_Long
            minimo = -2147483648#
            maximo = 2147483647
        
        Case eNumber_Types.ent_Trigger
            minimo = 0
            maximo = 10
    End Select
    
    If Val(Numero) >= minimo And Val(Numero) <= maximo Then ValidNumber = True
End Function

''
' Returns whether the ip format is correct.
'
' @param    IP The ip to be checked.

Private Function validipv4str(ByVal IP As String) As Boolean

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 01/06/07
    '
    '***************************************************
    Dim tmpArr() As String
    
    tmpArr = Split(IP, ".")
    
    If UBound(tmpArr) <> 3 Then Exit Function

    If Not ValidNumber(tmpArr(0), eNumber_Types.ent_Byte) Or Not ValidNumber(tmpArr(1), eNumber_Types.ent_Byte) Or Not ValidNumber(tmpArr(2), eNumber_Types.ent_Byte) Or Not ValidNumber(tmpArr(3), eNumber_Types.ent_Byte) Then Exit Function
    
    validipv4str = True
End Function

''
' Converts a string into the correct ip format.
'
' @param    IP The ip to be converted.

Private Function str2ipv4l(ByVal IP As String) As Byte()

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 07/26/07
    'Last Modified By: Rapsodius
    'Specify Return Type as Array of Bytes
    'Otherwise, the default is a Variant or Array of Variants, that slows down
    'the function
    '***************************************************
    Dim tmpArr() As String

    Dim bArr(3)  As Byte
    
    tmpArr = Split(IP, ".")
    
    bArr(0) = CByte(tmpArr(0))
    bArr(1) = CByte(tmpArr(1))
    bArr(2) = CByte(tmpArr(2))
    bArr(3) = CByte(tmpArr(3))

    str2ipv4l = bArr
End Function

''
' Do an Split() in the /AEMAIL in onother way
'
' @param text All the comand without the /aemail
' @return An bidimensional array with user and mail

Private Function AEMAILSplit(ByRef Text As String) As String()

    '***************************************************
    'Author: Lucas Tavolaro Ortuz (Tavo)
    'Useful for AEMAIL BUG FIX
    'Last Modification: 07/26/07
    'Last Modified By: Rapsodius
    'Specify Return Type as Array of Strings
    'Otherwise, the default is a Variant or Array of Variants, that slows down
    'the function
    '***************************************************
    Dim tmpArr(0 To 1) As String

    Dim Pos            As Byte
    
    Pos = InStr(1, Text, "-")
    
    If Pos <> 0 Then
        tmpArr(0) = mid$(Text, 1, Pos - 1)
        tmpArr(1) = mid$(Text, Pos + 1)
    Else
        tmpArr(0) = vbNullString
    End If
    
    AEMAILSplit = tmpArr
End Function
