Attribute VB_Name = "mSecurity"

' SecurityKey
Public Const MAX_KEY             As Integer = 500

Public SecurityKey(0 To MAX_KEY) As Byte

Public SecurityKey_Number        As Integer

' INTERVALOS
Public Enum eInterval

    iUseItem = 0
    iUseItemClick = 1
    iUseSpell = 2

End Enum

Public Type tInterval

    Default As Long
    Modify As Long
    UseInvalid As Byte
    ModifyTime As Long

End Type

Public Const MAX_INTERVAL                   As Byte = 2

Public Intervalos(0 To MAX_INTERVAL)        As tInterval

Public DefaultIntervalos(0 To MAX_INTERVAL) As Long

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal nCode As Long) As Integer

' Declaraciones del Api
'*********************************************************************************
  
' Enumera los procesos
  
' Retorna un array que contiene la lista de id de los procesos
Private Declare Function EnumProcesses _
                Lib "PSAPI.DLL" (ByRef lpidProcess As Long, _
                                 ByVal cB As Long, _
                                 ByRef cbNeeded As Long) As Long
  
' Abre un proceso para poder obtener el path ( Retorna el handle )
Private Declare Function OpenProcess _
                Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, _
                                    ByVal bInheritHandle As Long, _
                                    ByVal dwProcessId As Long) As Long
  
' Obtiene el nombre del proceso a partir de un handle _
  obtenido con EnumProcesses

Private Declare Function GetModuleFileNameExA _
                Lib "PSAPI.DLL" (ByVal hProcess As Long, _
                                 ByVal hModule As Long, _
                                 ByVal lpFileName As String, _
                                 ByVal nSize As Long) As Long
  
' Cierra y libera el proceso abierto con OpenProcess
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long

' Nombre de las ventanas
Declare Function GetWindowText _
        Lib "user32" _
        Alias "GetWindowTextA" (ByVal hWnd As Long, _
                                ByVal lpString As String, _
                                ByVal cch As Long) As Long
                                                            
' Esta es la función Api que busca las ventanas y retorna su handle o Hwnd
Private Declare Function GetWindow _
                Lib "user32" (ByVal hWnd As Long, _
                              ByVal wFlag As Long) As Long
                                    
'Esta función Api devuelve un valor  Boolean indicando si la ventana es una ventana visible
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
                                    
'Esta función retorna el número de caracteres del caption de la ventana
Private Declare Function GetWindowTextLength _
                Lib "user32" _
                Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
        
'Constantes para buscar las ventanas mediante el Api GetWindow
Private Const GW_HWNDFIRST = 0&

Private Const GW_HWNDNEXT = 2&

Private Const GW_CHILD = 5&

' Constantes
  
Private Const PROCESS_VM_READ           As Long = (&H10)

Private Const PROCESS_QUERY_INFORMATION As Long = (&H400)
  
' Rutina que recorre todos los procesos abiertos y devuelve el _
  nombre y path de los procesos  para listarlos en un control ListBox

'*********************************************************************************
Sub Enumerar_Procesos()

    Dim Array_Procesos() As Long
    Dim Buffer           As String
    Dim i_Procesos       As Long
    Dim ret              As Long
    Dim Ruta             As String
    Dim t_cbNeeded       As Long
    Dim Handle_Proceso   As Long
    Dim i                As Long
    Dim Sender           As Boolean
    ReDim Array_Procesos(250) As Long
      
    ' Obtiene un array con los id de los procesos
    ret = EnumProcesses(Array_Procesos(1), 1000, t_cbNeeded)
  
    i_Procesos = t_cbNeeded / 4
      
    ' Recorre todos los procesos
    For i = 1 To i_Procesos
        ' Lo abre y devuelve el handle
        Handle_Proceso = OpenProcess(PROCESS_QUERY_INFORMATION + PROCESS_VM_READ, 0, Array_Procesos(i))
              
        If Handle_Proceso <> 0 Then
            ' Crea un buffer para almacenar el nombre y ruta
            Buffer = Space(255)
                  
            ' Le pasa el Buffer al Api y el Handle
            ret = GetModuleFileNameExA(Handle_Proceso, 0, Buffer, 255)
            ' Le elimina los espacios nulos a la cadena devuelta
            Ruta = Left(Buffer, ret)
              
            ' Cierra el proceso abierto
            ret = CloseHandle(Handle_Proceso)
                  
            If Ruta_Valid(UCase$(Ruta)) Then
                Ruta = Replace(Ruta, "C:\Windows\System32\", "\Sys\")
                    
                Call WriteSendListSecurity(Ruta, 1)
                'Sender = True
            End If
        End If

        DoEvents
    Next
    
    'If Sender Then modNetwork.Flush
  
End Sub

Private Function Ruta_Valid(Ruta As String) As Boolean

    Select Case Ruta

        Case vbNullString
            Ruta_Valid = False

        Case "C:\WINDOWS\SYSTEM32\SVCHOST.EXE"
            Ruta_Valid = False

        Case "C:\WINDOWS\SYSTEM32\RUNTIMEBROKER.EXE"
            Ruta_Valid = False

        Case Else
            Ruta_Valid = True
    End Select
    
End Function

Public Sub Enumerar_Ventanas()

    Dim buf As Long, handle As Long, titulo As String, lenT As Long, ret As Long
    Dim Sender As Boolean
    
    'Obtenemos el Hwnd de la primera ventana, usando la constante GW_HWNDFIRST
    handle = GetWindow(FrmMain.hWnd, GW_HWNDFIRST)

    'Este bucle va a recorrer todas las ventanas.
    'cuando GetWindow devielva un 0, es por que no hay mas
    Do While handle <> 0

        'Tenemos que comprobar que la ventana es una de tipo visible
        If IsWindowVisible(handle) Then
            'Obtenemos el número de caracteres de la ventana
            lenT = GetWindowTextLength(handle)

            'si es el número anterior es mayor a 0
            If lenT > 0 Then
                'Creamos un buffer. Este buffer tendrá el tamaño con la variable LenT
                titulo = String$(lenT, 0)
                'Ahora recuperamos el texto de la ventana en el buffer que le enviamos
                'y también debemos pasarle el Hwnd de dicha ventana
                ret = GetWindowText(handle, titulo, lenT + 1)
                titulo$ = Left$(titulo, ret)
                
                Call WriteSendListSecurity(titulo$, 2)
                'Sender = True
            End If
        End If

        'Buscamos con GetWindow la próxima ventana usando la constante GW_HWNDNEXT
        handle = GetWindow(handle, GW_HWNDNEXT)
    Loop
    
    'If Sender Then modNetwork.Flush

End Sub

Public Function AoDefMacrer(ByVal KeyCode As Integer) As Boolean

    If Not GetAsyncKeyState(KeyCode) < 0 Then
        AoDefMacrer = True
    Else
        AoDefMacrer = False
    End If

End Function

' Intervalos Default
Public Sub SetIntervalos()
    
    DefaultIntervalos(eInterval.iUseItem) = 2
    DefaultIntervalos(eInterval.iUseItemClick) = 5
    DefaultIntervalos(eInterval.iUseSpell) = 1000
    
End Sub

' Restamos tiempo de los intervalos
Public Sub LoopInterval()

    Dim A As Long
    
    For A = 0 To MAX_INTERVAL

        If Intervalos(A).Modify > 0 Then Intervalos(A).Modify = Intervalos(A).Modify - 1
    Next A
    
End Sub

' Chequeamos si un intervalo sigue descontando
Public Function CheckInterval(ByVal iType As eInterval) As Boolean
    
    Dim Time As Long
    
    Time = FrameTime - Intervalos(iType).ModifyTime
    
    ' ShowConsoleMsg Time
    
    If (Time) <= 230 Then

        Exit Function

    End If
    
    'If Intervalos(iType).Modify > 0 Then Exit Function
    
    CheckInterval = True
    
End Function

' Asignamos al intervalo el tiempo para descontarlo
Public Sub AssignedInterval(ByVal iType As eInterval)
                                
    'Intervalos(iType).Modify = DefaultIntervalos(iType)
    Intervalos(iType).ModifyTime = FrameTime
    
End Sub

Public Function CheckPrincipales() As Boolean

    If Not FrmMain.visible Then Exit Function
          
    Dim strTemp   As String

    Dim CheatName As String

    strTemp = LstPscGS
          
    If InStr(UCase$(strTemp), UCase$("XMouseButtonControl")) <> 0 Then
        WriteDenounce "Estoy usando un posible CHEAT: X-MouseButton"
        ShowConsoleMsg "Recuerda que las configuraciones del X-MouseButton no están permitidas en estas Tierras. Prefiririamos que lo cerraras para evitar baneos de personajes. Gracias", 200, 200, 200, True
    End If
          
    If InStr(UCase$(strTemp), UCase$("Macro")) <> 0 Then
        WriteDenounce "Estoy usando un posible CHEAT: Macro"
    End If
        
    If InStr(UCase$(strTemp), UCase$("Razer")) <> 0 Then
        ShowConsoleMsg "Recuerda que las configuraciones del MouseRazer no están permitidas en estas Tierras. Prefiririamos que lo cerraras para evitar baneos de personajes. Gracias", 200, 200, 200, True
    End If
              
    If InStr(UCase$(strTemp), UCase$("Cheat")) <> 0 Then
        WriteDenounce "Estoy usando un posible CHEAT: Cheat"
    End If
          
End Function

Public Sub Initialize_Security()
    SecurityKey(0) = 67
    SecurityKey(1) = 24
    SecurityKey(2) = 47
    SecurityKey(3) = 147
    SecurityKey(4) = 29
    SecurityKey(5) = 81
    SecurityKey(6) = 110
    SecurityKey(7) = 94
    SecurityKey(8) = 105
    SecurityKey(9) = 166
    SecurityKey(10) = 4
    SecurityKey(11) = 27
    SecurityKey(12) = 245
    SecurityKey(13) = 252
    SecurityKey(14) = 85
    SecurityKey(15) = 111
    SecurityKey(16) = 94
    SecurityKey(17) = 204
    SecurityKey(18) = 5
    SecurityKey(19) = 66
    SecurityKey(20) = 131
    SecurityKey(21) = 201
    SecurityKey(22) = 11
    SecurityKey(23) = 123
    SecurityKey(24) = 57
    SecurityKey(25) = 195
    SecurityKey(26) = 7
    SecurityKey(27) = 10
    SecurityKey(28) = 64
    SecurityKey(29) = 203
    SecurityKey(30) = 213
    SecurityKey(31) = 44
    SecurityKey(32) = 118
    SecurityKey(33) = 152
    SecurityKey(34) = 98
    SecurityKey(35) = 234
    SecurityKey(36) = 75
    SecurityKey(37) = 41
    SecurityKey(38) = 190
    SecurityKey(39) = 227
    SecurityKey(40) = 117
    SecurityKey(41) = 172
    SecurityKey(42) = 115
    SecurityKey(43) = 76
    SecurityKey(44) = 229
    SecurityKey(45) = 159
    SecurityKey(46) = 22
    SecurityKey(47) = 53
    SecurityKey(48) = 249
    SecurityKey(49) = 53
    SecurityKey(50) = 27
    SecurityKey(51) = 14
    SecurityKey(52) = 243
    SecurityKey(53) = 251
    SecurityKey(54) = 237
    SecurityKey(55) = 105
    SecurityKey(56) = 170
    SecurityKey(57) = 187
    SecurityKey(58) = 62
    SecurityKey(59) = 1
    SecurityKey(60) = 127
    SecurityKey(61) = 160
    SecurityKey(62) = 156
    SecurityKey(63) = 252
    SecurityKey(64) = 147
    SecurityKey(65) = 156
    SecurityKey(66) = 70
    SecurityKey(67) = 109
    SecurityKey(68) = 55
    SecurityKey(69) = 7
    SecurityKey(70) = 61
    SecurityKey(71) = 29
    SecurityKey(72) = 165
    SecurityKey(73) = 137
    SecurityKey(74) = 158
    SecurityKey(75) = 139
    SecurityKey(76) = 237
    SecurityKey(77) = 57
    SecurityKey(78) = 46
    SecurityKey(79) = 49
    SecurityKey(80) = 49
    SecurityKey(81) = 6
    SecurityKey(82) = 55
    SecurityKey(83) = 9
    SecurityKey(84) = 21
    SecurityKey(85) = 86
    SecurityKey(86) = 147
    SecurityKey(87) = 80
    SecurityKey(88) = 114
    SecurityKey(89) = 238
    SecurityKey(90) = 6
    SecurityKey(91) = 138
    SecurityKey(92) = 74
    SecurityKey(93) = 18
    SecurityKey(94) = 208
    SecurityKey(95) = 198
    SecurityKey(96) = 1
    SecurityKey(97) = 237
    SecurityKey(98) = 131
    SecurityKey(99) = 35
    SecurityKey(100) = 202
    SecurityKey(101) = 50
    SecurityKey(102) = 19
    SecurityKey(103) = 147
    SecurityKey(104) = 95
    SecurityKey(105) = 4
    SecurityKey(106) = 44
    SecurityKey(107) = 223
    SecurityKey(108) = 245
    SecurityKey(109) = 20
    SecurityKey(110) = 200
    SecurityKey(111) = 236
    SecurityKey(112) = 111
    SecurityKey(113) = 9
    SecurityKey(114) = 0
    SecurityKey(115) = 60
    SecurityKey(116) = 210
    SecurityKey(117) = 36
    SecurityKey(118) = 34
    SecurityKey(119) = 183
    SecurityKey(120) = 249
    SecurityKey(121) = 36
    SecurityKey(122) = 5
    SecurityKey(123) = 172
    SecurityKey(124) = 137
    SecurityKey(125) = 103
    SecurityKey(126) = 153
    SecurityKey(127) = 19
    SecurityKey(128) = 40
    SecurityKey(129) = 83
    SecurityKey(130) = 194
    SecurityKey(131) = 21
    SecurityKey(132) = 234
    SecurityKey(133) = 244
    SecurityKey(134) = 103
    SecurityKey(135) = 205
    SecurityKey(136) = 12
    SecurityKey(137) = 230
    SecurityKey(138) = 197
    SecurityKey(139) = 81
    SecurityKey(140) = 229
    SecurityKey(141) = 118
    SecurityKey(142) = 10
    SecurityKey(143) = 236
    SecurityKey(144) = 25
    SecurityKey(145) = 4
    SecurityKey(146) = 31
    SecurityKey(147) = 174
    SecurityKey(148) = 16
    SecurityKey(149) = 171
    SecurityKey(150) = 197
    SecurityKey(151) = 39
    SecurityKey(152) = 167
    SecurityKey(153) = 36
    SecurityKey(154) = 227
    SecurityKey(155) = 111
    SecurityKey(156) = 37
    SecurityKey(157) = 232
    SecurityKey(158) = 30
    SecurityKey(159) = 105
    SecurityKey(160) = 112
    SecurityKey(161) = 149
    SecurityKey(162) = 171
    SecurityKey(163) = 73
    SecurityKey(164) = 128
    SecurityKey(165) = 147
    SecurityKey(166) = 97
    SecurityKey(167) = 84
    SecurityKey(168) = 21
    SecurityKey(169) = 247
    SecurityKey(170) = 19
    SecurityKey(171) = 231
    SecurityKey(172) = 165
    SecurityKey(173) = 168
    SecurityKey(174) = 28
    SecurityKey(175) = 187
    SecurityKey(176) = 153
    SecurityKey(177) = 192
    SecurityKey(178) = 59
    SecurityKey(179) = 103
    SecurityKey(180) = 184
    SecurityKey(181) = 53
    SecurityKey(182) = 162
    SecurityKey(183) = 39
    SecurityKey(184) = 228
    SecurityKey(185) = 184
    SecurityKey(186) = 73
    SecurityKey(187) = 219
    SecurityKey(188) = 4
    SecurityKey(189) = 221
    SecurityKey(190) = 136
    SecurityKey(191) = 83
    SecurityKey(192) = 65
    SecurityKey(193) = 125
    SecurityKey(194) = 229
    SecurityKey(195) = 201
    SecurityKey(196) = 117
    SecurityKey(197) = 88
    SecurityKey(198) = 42
    SecurityKey(199) = 175
    SecurityKey(200) = 224
    SecurityKey(201) = 255
    SecurityKey(202) = 187
    SecurityKey(203) = 171
    SecurityKey(204) = 29
    SecurityKey(205) = 242
    SecurityKey(206) = 39
    SecurityKey(207) = 225
    SecurityKey(208) = 85
    SecurityKey(209) = 5
    SecurityKey(210) = 253
    SecurityKey(211) = 112
    SecurityKey(212) = 179
    SecurityKey(213) = 8
    SecurityKey(214) = 225
    SecurityKey(215) = 63
    SecurityKey(216) = 24
    SecurityKey(217) = 166
    SecurityKey(218) = 223
    SecurityKey(219) = 249
    SecurityKey(220) = 15
    SecurityKey(221) = 142
    SecurityKey(222) = 254
    SecurityKey(223) = 86
    SecurityKey(224) = 3
    SecurityKey(225) = 209
    SecurityKey(226) = 25
    SecurityKey(227) = 157
    SecurityKey(228) = 175
    SecurityKey(229) = 139
    SecurityKey(230) = 234
    SecurityKey(231) = 102
    SecurityKey(232) = 215
    SecurityKey(233) = 198
    SecurityKey(234) = 104
    SecurityKey(235) = 165
    SecurityKey(236) = 54
    SecurityKey(237) = 155
    SecurityKey(238) = 83
    SecurityKey(239) = 228
    SecurityKey(240) = 183
    SecurityKey(241) = 154
    SecurityKey(242) = 13
    SecurityKey(243) = 208
    SecurityKey(244) = 232
    SecurityKey(245) = 108
    SecurityKey(246) = 171
    SecurityKey(247) = 247
    SecurityKey(248) = 171
    SecurityKey(249) = 183
    SecurityKey(250) = 76
    SecurityKey(251) = 208
    SecurityKey(252) = 46
    SecurityKey(253) = 66
    SecurityKey(254) = 169
    SecurityKey(255) = 252
    SecurityKey(256) = 30
    SecurityKey(257) = 90
    SecurityKey(258) = 238
    SecurityKey(259) = 203
    SecurityKey(260) = 24
    SecurityKey(261) = 116
    SecurityKey(262) = 200
    SecurityKey(263) = 2
    SecurityKey(264) = 97
    SecurityKey(265) = 19
    SecurityKey(266) = 192
    SecurityKey(267) = 220
    SecurityKey(268) = 214
    SecurityKey(269) = 237
    SecurityKey(270) = 199
    SecurityKey(271) = 78
    SecurityKey(272) = 38
    SecurityKey(273) = 73
    SecurityKey(274) = 18
    SecurityKey(275) = 143
    SecurityKey(276) = 62
    SecurityKey(277) = 171
    SecurityKey(278) = 40
    SecurityKey(279) = 216
    SecurityKey(280) = 5
    SecurityKey(281) = 179
    SecurityKey(282) = 57
    SecurityKey(283) = 104
    SecurityKey(284) = 74
    SecurityKey(285) = 67
    SecurityKey(286) = 177
    SecurityKey(287) = 204
    SecurityKey(288) = 250
    SecurityKey(289) = 224
    SecurityKey(290) = 13
    SecurityKey(291) = 93
    SecurityKey(292) = 151
    SecurityKey(293) = 91
    SecurityKey(294) = 237
    SecurityKey(295) = 10
    SecurityKey(296) = 229
    SecurityKey(297) = 176
    SecurityKey(298) = 107
    SecurityKey(299) = 88
    SecurityKey(300) = 231
    SecurityKey(301) = 46
    SecurityKey(302) = 172
    SecurityKey(303) = 166
    SecurityKey(304) = 9
    SecurityKey(305) = 216
    SecurityKey(306) = 180
    SecurityKey(307) = 182
    SecurityKey(308) = 159
    SecurityKey(309) = 12
    SecurityKey(310) = 127
    SecurityKey(311) = 105
    SecurityKey(312) = 142
    SecurityKey(313) = 98
    SecurityKey(314) = 77
    SecurityKey(315) = 202
    SecurityKey(316) = 73
    SecurityKey(317) = 215
    SecurityKey(318) = 61
    SecurityKey(319) = 78
    SecurityKey(320) = 0
    SecurityKey(321) = 43
    SecurityKey(322) = 29
    SecurityKey(323) = 90
    SecurityKey(324) = 19
    SecurityKey(325) = 135
    SecurityKey(326) = 129
    SecurityKey(327) = 6
    SecurityKey(328) = 205
    SecurityKey(329) = 99
    SecurityKey(330) = 18
    SecurityKey(331) = 33
    SecurityKey(332) = 79
    SecurityKey(333) = 167
    SecurityKey(334) = 41
    SecurityKey(335) = 117
    SecurityKey(336) = 202
    SecurityKey(337) = 16
    SecurityKey(338) = 157
    SecurityKey(339) = 76
    SecurityKey(340) = 242
    SecurityKey(341) = 214
    SecurityKey(342) = 216
    SecurityKey(343) = 50
    SecurityKey(344) = 175
    SecurityKey(345) = 140
    SecurityKey(346) = 49
    SecurityKey(347) = 253
    SecurityKey(348) = 21
    SecurityKey(349) = 71
    SecurityKey(350) = 117
    SecurityKey(351) = 11
    SecurityKey(352) = 150
    SecurityKey(353) = 2
    SecurityKey(354) = 199
    SecurityKey(355) = 203
    SecurityKey(356) = 118
    SecurityKey(357) = 65
    SecurityKey(358) = 171
    SecurityKey(359) = 127
    SecurityKey(360) = 128
    SecurityKey(361) = 245
    SecurityKey(362) = 93
    SecurityKey(363) = 64
    SecurityKey(364) = 248
    SecurityKey(365) = 160
    SecurityKey(366) = 103
    SecurityKey(367) = 66
    SecurityKey(368) = 208
    SecurityKey(369) = 185
    SecurityKey(370) = 114
    SecurityKey(371) = 89
    SecurityKey(372) = 30
    SecurityKey(373) = 82
    SecurityKey(374) = 93
    SecurityKey(375) = 188
    SecurityKey(376) = 206
    SecurityKey(377) = 248
    SecurityKey(378) = 140
    SecurityKey(379) = 9
    SecurityKey(380) = 148
    SecurityKey(381) = 219
    SecurityKey(382) = 131
    SecurityKey(383) = 138
    SecurityKey(384) = 37
    SecurityKey(385) = 46
    SecurityKey(386) = 179
    SecurityKey(387) = 183
    SecurityKey(388) = 167
    SecurityKey(389) = 209
    SecurityKey(390) = 147
    SecurityKey(391) = 252
    SecurityKey(392) = 102
    SecurityKey(393) = 46
    SecurityKey(394) = 243
    SecurityKey(395) = 188
    SecurityKey(396) = 200
    SecurityKey(397) = 96
    SecurityKey(398) = 141
    SecurityKey(399) = 149
    SecurityKey(400) = 131
    SecurityKey(401) = 155
    SecurityKey(402) = 222
    SecurityKey(403) = 230
    SecurityKey(404) = 13
    SecurityKey(405) = 200
    SecurityKey(406) = 52
    SecurityKey(407) = 142
    SecurityKey(408) = 84
    SecurityKey(409) = 111
    SecurityKey(410) = 7
    SecurityKey(411) = 247
    SecurityKey(412) = 176
    SecurityKey(413) = 218
    SecurityKey(414) = 140
    SecurityKey(415) = 83
    SecurityKey(416) = 22
    SecurityKey(417) = 120
    SecurityKey(418) = 136
    SecurityKey(419) = 38
    SecurityKey(420) = 142
    SecurityKey(421) = 127
    SecurityKey(422) = 98
    SecurityKey(423) = 5
    SecurityKey(424) = 231
    SecurityKey(425) = 213
    SecurityKey(426) = 125
    SecurityKey(427) = 157
    SecurityKey(428) = 169
    SecurityKey(429) = 49
    SecurityKey(430) = 196
    SecurityKey(431) = 246
    SecurityKey(432) = 75
    SecurityKey(433) = 125
    SecurityKey(434) = 135
    SecurityKey(435) = 249
    SecurityKey(436) = 166
    SecurityKey(437) = 127
    SecurityKey(438) = 133
    SecurityKey(439) = 49
    SecurityKey(440) = 170
    SecurityKey(441) = 185
    SecurityKey(442) = 74
    SecurityKey(443) = 206
    SecurityKey(444) = 80
    SecurityKey(445) = 142
    SecurityKey(446) = 187
    SecurityKey(447) = 239
    SecurityKey(448) = 207
    SecurityKey(449) = 165
    SecurityKey(450) = 239
    SecurityKey(451) = 33
    SecurityKey(452) = 19
    SecurityKey(453) = 147
    SecurityKey(454) = 64
    SecurityKey(455) = 34
    SecurityKey(456) = 107
    SecurityKey(457) = 180
    SecurityKey(458) = 162
    SecurityKey(459) = 235
    SecurityKey(460) = 130
    SecurityKey(461) = 89
    SecurityKey(462) = 52
    SecurityKey(463) = 238
    SecurityKey(464) = 144
    SecurityKey(465) = 41
    SecurityKey(466) = 21
    SecurityKey(467) = 157
    SecurityKey(468) = 209
    SecurityKey(469) = 193
    SecurityKey(470) = 121
    SecurityKey(471) = 43
    SecurityKey(472) = 54
    SecurityKey(473) = 158
    SecurityKey(474) = 252
    SecurityKey(475) = 150
    SecurityKey(476) = 91
    SecurityKey(477) = 61
    SecurityKey(478) = 53
    SecurityKey(479) = 229
    SecurityKey(480) = 186
    SecurityKey(481) = 128
    SecurityKey(482) = 143
    SecurityKey(483) = 174
    SecurityKey(484) = 30
    SecurityKey(485) = 84
    SecurityKey(486) = 84
    SecurityKey(487) = 220
    SecurityKey(488) = 90
    SecurityKey(489) = 145
    SecurityKey(490) = 11
    SecurityKey(491) = 175
    SecurityKey(492) = 58
    SecurityKey(493) = 33
    SecurityKey(494) = 4
    SecurityKey(495) = 4
    SecurityKey(496) = 186
    SecurityKey(497) = 101
    SecurityKey(498) = 49
    SecurityKey(499) = 215
    SecurityKey(500) = 118
End Sub

