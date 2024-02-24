Attribute VB_Name = "mSecurity"
'

Option Explicit

Public Enum eInterval

    iUseItem = 0
    iUseItemClick = 1
    iUseSpell = 2

End Enum

Public Type tInterval

    Default As Long
    ModifyTimer As Long
    UseInvalid As Byte

End Type

' Maximo de Intervalos
Public Const MAX_INTERVAL As Byte = 2

' Calcular Tiempos
Public Declare Function timeGetTime Lib "winmm.dll" () As Long

' Constants Key Code Packets
Public Const MAX_KEY_PACKETS    As Byte = 2

Public Const MAX_KEY_CHANGE     As Byte = 30

' Constants Pointers
Public Const MAX_POINTERS       As Byte = 2

Public Const LIMIT_POINTER      As Byte = 9

Public Const LIMIT_FLOD_POINTER As Byte = 10

' Position of the pointer
Public Enum ePoint

    Point_Spell = 1
    Point_Inv = 2

End Enum

' Cursor X, Y
Public Type tPoint

    X(LIMIT_POINTER) As Long
    Y(LIMIT_POINTER) As Long
    cant(LIMIT_POINTER) As Byte

End Type

Public Type tPackets

    Value As Byte
    cant As Byte

End Type

' Key Code of the special Packets
Public Enum eKeyPackets

    Key_UseItem = 0
    Key_UseSpell = 1
    Key_UseWeapon = 2

End Enum

' Check Key Code
Public Function CheckKeyPacket(ByVal UserIndex As Integer, _
                               ByVal Packet As eKeyPackets, _
                               ByVal KeyPacket As Long) As Boolean
        '<EhHeader>
        On Error GoTo CheckKeyPacket_Err
        '</EhHeader>
                            
100     With UserList(UserIndex)
    
102         If .KeyPackets(Packet).Value = 0 Then
                'UpdateKeyPacket UserIndex, Packet
104             CheckKeyPacket = True

                Exit Function

            End If
        
106         If .KeyPackets(Packet).Value <> KeyPacket Then Exit Function
        
108         .KeyPackets(Packet).cant = .KeyPackets(Packet).cant + 1
        
110         If .KeyPackets(Packet).cant = MAX_KEY_CHANGE Then
                'UpdateKeyPacket UserIndex, Packet
            End If
        
        End With
    
112     CheckKeyPacket = True
        '<EhFooter>
        Exit Function

CheckKeyPacket_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mSecurity.CheckKeyPacket " & _
               "at line " & Erl
        
        '</EhFooter>
End Function
                        
' Reset Key Code
Public Sub ResetKeyPackets(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo ResetKeyPackets_Err
        '</EhHeader>

        Dim A As Long
    
100     With UserList(UserIndex)
    
102         For A = 0 To MAX_KEY_PACKETS
104             .KeyPackets(A).Value = 0
106             .KeyPackets(A).cant = 0
108         Next A
        
        End With

        '<EhFooter>
        Exit Sub

ResetKeyPackets_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mSecurity.ResetKeyPackets " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

' Reset Pointer
Public Sub ResetPointer(ByVal UserIndex As Integer, ByVal Point As ePoint)
        '<EhHeader>
        On Error GoTo ResetPointer_Err
        '</EhHeader>

        Dim A As Long
    
100     With UserList(UserIndex).Pointers(Point)
    
102         For A = 0 To LIMIT_POINTER
104             .cant(A) = 0
106             .X(A) = 0
108             .Y(A) = 0
110         Next A
        
        End With

        '<EhFooter>
        Exit Sub

ResetPointer_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mSecurity.ResetPointer " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub UpdatePointer(ByVal UserIndex As Integer, _
                         ByVal Point As ePoint, _
                         ByVal X As Long, _
                         ByVal Y As Long, _
                         ByVal Ident As String)
        '<EhHeader>
        On Error GoTo UpdatePointer_Err
        '</EhHeader>
                        
        Dim A           As Integer

        Dim PointerCero As Byte
    
100     With UserList(UserIndex).Pointers(Point)

102         For A = 0 To LIMIT_POINTER
            
                ' Pointer cero
104             If PointerCero = 0 And (.X(A) = 0 And .Y(A) = 0) Then
                
106                 PointerCero = A
                End If
            
                ' Pointer repetido
108             If .X(A) = X And .Y(A) = Y Then
110                 .cant(A) = .cant(A) + 1
                
                    ' (Máximo permitido en el Point de igualdad)
112                 If .cant(A) = LIMIT_FLOD_POINTER Then
114                     Call Logs_Security(eLog.eSecurity, eLogSecurity.eAntiCheat, Ident & ": Misma posición de marcado " & UserList(UserIndex).Account.Email & " NICK: " & UserList(UserIndex).Name & " IP: " & UserList(UserIndex).IpAddress)
                    
116                     Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(Ident & ": Misma posición de marcado " & UserList(UserIndex).Account.Email & " NICK: " & UserList(UserIndex).Name & " IP: " & UserList(UserIndex).IpAddress, FontTypeNames.FONTTYPE_SERVER, eMessageType.Admin))
                    
118                     ResetPointer UserIndex, Point

                        Exit Sub

                    End If
                
                    Exit Sub

                End If

120         Next A
        
122         If PointerCero = 0 Then
124             ResetPointer UserIndex, Point
126             .X(0) = X
128             .Y(0) = Y
130             .cant(0) = 1
            Else
132             .X(PointerCero) = X
134             .Y(PointerCero) = Y
136             .cant(PointerCero) = 1
            End If
        
        End With

        '<EhFooter>
        Exit Sub

UpdatePointer_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mSecurity.UpdatePointer " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub Initialize_Security()
        '<EhHeader>
        On Error GoTo Initialize_Security_Err
        '</EhHeader>
100     SecurityKey(0) = 67
102     SecurityKey(1) = 24
104     SecurityKey(2) = 47
106     SecurityKey(3) = 147
108     SecurityKey(4) = 29
110     SecurityKey(5) = 81
112     SecurityKey(6) = 110
114     SecurityKey(7) = 94
116     SecurityKey(8) = 105
118     SecurityKey(9) = 166
120     SecurityKey(10) = 4
122     SecurityKey(11) = 27
124     SecurityKey(12) = 245
126     SecurityKey(13) = 252
128     SecurityKey(14) = 85
130     SecurityKey(15) = 111
132     SecurityKey(16) = 94
134     SecurityKey(17) = 204
136     SecurityKey(18) = 5
138     SecurityKey(19) = 66
140     SecurityKey(20) = 131
142     SecurityKey(21) = 201
144     SecurityKey(22) = 11
146     SecurityKey(23) = 123
148     SecurityKey(24) = 57
150     SecurityKey(25) = 195
152     SecurityKey(26) = 7
154     SecurityKey(27) = 10
156     SecurityKey(28) = 64
158     SecurityKey(29) = 203
160     SecurityKey(30) = 213
162     SecurityKey(31) = 44
164     SecurityKey(32) = 118
166     SecurityKey(33) = 152
168     SecurityKey(34) = 98
170     SecurityKey(35) = 234
172     SecurityKey(36) = 75
174     SecurityKey(37) = 41
176     SecurityKey(38) = 190
178     SecurityKey(39) = 227
180     SecurityKey(40) = 117
182     SecurityKey(41) = 172
184     SecurityKey(42) = 115
186     SecurityKey(43) = 76
188     SecurityKey(44) = 229
190     SecurityKey(45) = 159
192     SecurityKey(46) = 22
194     SecurityKey(47) = 53
196     SecurityKey(48) = 249
198     SecurityKey(49) = 53
200     SecurityKey(50) = 27
202     SecurityKey(51) = 14
204     SecurityKey(52) = 243
206     SecurityKey(53) = 251
208     SecurityKey(54) = 237
210     SecurityKey(55) = 105
212     SecurityKey(56) = 170
214     SecurityKey(57) = 187
216     SecurityKey(58) = 62
218     SecurityKey(59) = 1
220     SecurityKey(60) = 127
222     SecurityKey(61) = 160
224     SecurityKey(62) = 156
226     SecurityKey(63) = 252
228     SecurityKey(64) = 147
230     SecurityKey(65) = 156
232     SecurityKey(66) = 70
234     SecurityKey(67) = 109
236     SecurityKey(68) = 55
238     SecurityKey(69) = 7
240     SecurityKey(70) = 61
242     SecurityKey(71) = 29
244     SecurityKey(72) = 165
246     SecurityKey(73) = 137
248     SecurityKey(74) = 158
250     SecurityKey(75) = 139
252     SecurityKey(76) = 237
254     SecurityKey(77) = 57
256     SecurityKey(78) = 46
258     SecurityKey(79) = 49
260     SecurityKey(80) = 49
262     SecurityKey(81) = 6
264     SecurityKey(82) = 55
266     SecurityKey(83) = 9
268     SecurityKey(84) = 21
270     SecurityKey(85) = 86
272     SecurityKey(86) = 147
274     SecurityKey(87) = 80
276     SecurityKey(88) = 114
278     SecurityKey(89) = 238
280     SecurityKey(90) = 6
282     SecurityKey(91) = 138
284     SecurityKey(92) = 74
286     SecurityKey(93) = 18
288     SecurityKey(94) = 208
290     SecurityKey(95) = 198
292     SecurityKey(96) = 1
294     SecurityKey(97) = 237
296     SecurityKey(98) = 131
298     SecurityKey(99) = 35
300     SecurityKey(100) = 202
302     SecurityKey(101) = 50
304     SecurityKey(102) = 19
306     SecurityKey(103) = 147
308     SecurityKey(104) = 95
310     SecurityKey(105) = 4
312     SecurityKey(106) = 44
314     SecurityKey(107) = 223
316     SecurityKey(108) = 245
318     SecurityKey(109) = 20
320     SecurityKey(110) = 200
322     SecurityKey(111) = 236
324     SecurityKey(112) = 111
326     SecurityKey(113) = 9
328     SecurityKey(114) = 0
330     SecurityKey(115) = 60
332     SecurityKey(116) = 210
334     SecurityKey(117) = 36
336     SecurityKey(118) = 34
338     SecurityKey(119) = 183
340     SecurityKey(120) = 249
342     SecurityKey(121) = 36
344     SecurityKey(122) = 5
346     SecurityKey(123) = 172
348     SecurityKey(124) = 137
350     SecurityKey(125) = 103
352     SecurityKey(126) = 153
354     SecurityKey(127) = 19
356     SecurityKey(128) = 40
358     SecurityKey(129) = 83
360     SecurityKey(130) = 194
362     SecurityKey(131) = 21
364     SecurityKey(132) = 234
366     SecurityKey(133) = 244
368     SecurityKey(134) = 103
370     SecurityKey(135) = 205
372     SecurityKey(136) = 12
374     SecurityKey(137) = 230
376     SecurityKey(138) = 197
378     SecurityKey(139) = 81
380     SecurityKey(140) = 229
382     SecurityKey(141) = 118
384     SecurityKey(142) = 10
386     SecurityKey(143) = 236
388     SecurityKey(144) = 25
390     SecurityKey(145) = 4
392     SecurityKey(146) = 31
394     SecurityKey(147) = 174
396     SecurityKey(148) = 16
398     SecurityKey(149) = 171
400     SecurityKey(150) = 197
402     SecurityKey(151) = 39
404     SecurityKey(152) = 167
406     SecurityKey(153) = 36
408     SecurityKey(154) = 227
410     SecurityKey(155) = 111
412     SecurityKey(156) = 37
414     SecurityKey(157) = 232
416     SecurityKey(158) = 30
418     SecurityKey(159) = 105
420     SecurityKey(160) = 112
422     SecurityKey(161) = 149
424     SecurityKey(162) = 171
426     SecurityKey(163) = 73
428     SecurityKey(164) = 128
430     SecurityKey(165) = 147
432     SecurityKey(166) = 97
434     SecurityKey(167) = 84
436     SecurityKey(168) = 21
438     SecurityKey(169) = 247
440     SecurityKey(170) = 19
442     SecurityKey(171) = 231
444     SecurityKey(172) = 165
446     SecurityKey(173) = 168
448     SecurityKey(174) = 28
450     SecurityKey(175) = 187
452     SecurityKey(176) = 153
454     SecurityKey(177) = 192
456     SecurityKey(178) = 59
458     SecurityKey(179) = 103
460     SecurityKey(180) = 184
462     SecurityKey(181) = 53
464     SecurityKey(182) = 162
466     SecurityKey(183) = 39
468     SecurityKey(184) = 228
470     SecurityKey(185) = 184
472     SecurityKey(186) = 73
474     SecurityKey(187) = 219
476     SecurityKey(188) = 4
478     SecurityKey(189) = 221
480     SecurityKey(190) = 136
482     SecurityKey(191) = 83
484     SecurityKey(192) = 65
486     SecurityKey(193) = 125
488     SecurityKey(194) = 229
490     SecurityKey(195) = 201
492     SecurityKey(196) = 117
494     SecurityKey(197) = 88
496     SecurityKey(198) = 42
498     SecurityKey(199) = 175
500     SecurityKey(200) = 224
502     SecurityKey(201) = 255
504     SecurityKey(202) = 187
506     SecurityKey(203) = 171
508     SecurityKey(204) = 29
510     SecurityKey(205) = 242
512     SecurityKey(206) = 39
514     SecurityKey(207) = 225
516     SecurityKey(208) = 85
518     SecurityKey(209) = 5
520     SecurityKey(210) = 253
522     SecurityKey(211) = 112
524     SecurityKey(212) = 179
526     SecurityKey(213) = 8
528     SecurityKey(214) = 225
530     SecurityKey(215) = 63
532     SecurityKey(216) = 24
534     SecurityKey(217) = 166
536     SecurityKey(218) = 223
538     SecurityKey(219) = 249
540     SecurityKey(220) = 15
542     SecurityKey(221) = 142
544     SecurityKey(222) = 254
546     SecurityKey(223) = 86
548     SecurityKey(224) = 3
550     SecurityKey(225) = 209
552     SecurityKey(226) = 25
554     SecurityKey(227) = 157
556     SecurityKey(228) = 175
558     SecurityKey(229) = 139
560     SecurityKey(230) = 234
562     SecurityKey(231) = 102
564     SecurityKey(232) = 215
566     SecurityKey(233) = 198
568     SecurityKey(234) = 104
570     SecurityKey(235) = 165
572     SecurityKey(236) = 54
574     SecurityKey(237) = 155
576     SecurityKey(238) = 83
578     SecurityKey(239) = 228
580     SecurityKey(240) = 183
582     SecurityKey(241) = 154
584     SecurityKey(242) = 13
586     SecurityKey(243) = 208
588     SecurityKey(244) = 232
590     SecurityKey(245) = 108
592     SecurityKey(246) = 171
594     SecurityKey(247) = 247
596     SecurityKey(248) = 171
598     SecurityKey(249) = 183
600     SecurityKey(250) = 76
602     SecurityKey(251) = 208
604     SecurityKey(252) = 46
606     SecurityKey(253) = 66
608     SecurityKey(254) = 169
610     SecurityKey(255) = 252
612     SecurityKey(256) = 30
614     SecurityKey(257) = 90
616     SecurityKey(258) = 238
618     SecurityKey(259) = 203
620     SecurityKey(260) = 24
622     SecurityKey(261) = 116
624     SecurityKey(262) = 200
626     SecurityKey(263) = 2
628     SecurityKey(264) = 97
630     SecurityKey(265) = 19
632     SecurityKey(266) = 192
634     SecurityKey(267) = 220
636     SecurityKey(268) = 214
638     SecurityKey(269) = 237
640     SecurityKey(270) = 199
642     SecurityKey(271) = 78
644     SecurityKey(272) = 38
646     SecurityKey(273) = 73
648     SecurityKey(274) = 18
650     SecurityKey(275) = 143
652     SecurityKey(276) = 62
654     SecurityKey(277) = 171
656     SecurityKey(278) = 40
658     SecurityKey(279) = 216
660     SecurityKey(280) = 5
662     SecurityKey(281) = 179
664     SecurityKey(282) = 57
666     SecurityKey(283) = 104
668     SecurityKey(284) = 74
670     SecurityKey(285) = 67
672     SecurityKey(286) = 177
674     SecurityKey(287) = 204
676     SecurityKey(288) = 250
678     SecurityKey(289) = 224
680     SecurityKey(290) = 13
682     SecurityKey(291) = 93
684     SecurityKey(292) = 151
686     SecurityKey(293) = 91
688     SecurityKey(294) = 237
690     SecurityKey(295) = 10
692     SecurityKey(296) = 229
694     SecurityKey(297) = 176
696     SecurityKey(298) = 107
698     SecurityKey(299) = 88
700     SecurityKey(300) = 231
702     SecurityKey(301) = 46
704     SecurityKey(302) = 172
706     SecurityKey(303) = 166
708     SecurityKey(304) = 9
710     SecurityKey(305) = 216
712     SecurityKey(306) = 180
714     SecurityKey(307) = 182
716     SecurityKey(308) = 159
718     SecurityKey(309) = 12
720     SecurityKey(310) = 127
722     SecurityKey(311) = 105
724     SecurityKey(312) = 142
726     SecurityKey(313) = 98
728     SecurityKey(314) = 77
730     SecurityKey(315) = 202
732     SecurityKey(316) = 73
734     SecurityKey(317) = 215
736     SecurityKey(318) = 61
738     SecurityKey(319) = 78
740     SecurityKey(320) = 0
742     SecurityKey(321) = 43
744     SecurityKey(322) = 29
746     SecurityKey(323) = 90
748     SecurityKey(324) = 19
750     SecurityKey(325) = 135
752     SecurityKey(326) = 129
754     SecurityKey(327) = 6
756     SecurityKey(328) = 205
758     SecurityKey(329) = 99
760     SecurityKey(330) = 18
762     SecurityKey(331) = 33
764     SecurityKey(332) = 79
766     SecurityKey(333) = 167
768     SecurityKey(334) = 41
770     SecurityKey(335) = 117
772     SecurityKey(336) = 202
774     SecurityKey(337) = 16
776     SecurityKey(338) = 157
778     SecurityKey(339) = 76
780     SecurityKey(340) = 242
782     SecurityKey(341) = 214
784     SecurityKey(342) = 216
786     SecurityKey(343) = 50
788     SecurityKey(344) = 175
790     SecurityKey(345) = 140
792     SecurityKey(346) = 49
794     SecurityKey(347) = 253
796     SecurityKey(348) = 21
798     SecurityKey(349) = 71
800     SecurityKey(350) = 117
802     SecurityKey(351) = 11
804     SecurityKey(352) = 150
806     SecurityKey(353) = 2
808     SecurityKey(354) = 199
810     SecurityKey(355) = 203
812     SecurityKey(356) = 118
814     SecurityKey(357) = 65
816     SecurityKey(358) = 171
818     SecurityKey(359) = 127
820     SecurityKey(360) = 128
822     SecurityKey(361) = 245
824     SecurityKey(362) = 93
826     SecurityKey(363) = 64
828     SecurityKey(364) = 248
830     SecurityKey(365) = 160
832     SecurityKey(366) = 103
834     SecurityKey(367) = 66
836     SecurityKey(368) = 208
838     SecurityKey(369) = 185
840     SecurityKey(370) = 114
842     SecurityKey(371) = 89
844     SecurityKey(372) = 30
846     SecurityKey(373) = 82
848     SecurityKey(374) = 93
850     SecurityKey(375) = 188
852     SecurityKey(376) = 206
854     SecurityKey(377) = 248
856     SecurityKey(378) = 140
858     SecurityKey(379) = 9
860     SecurityKey(380) = 148
862     SecurityKey(381) = 219
864     SecurityKey(382) = 131
866     SecurityKey(383) = 138
868     SecurityKey(384) = 37
870     SecurityKey(385) = 46
872     SecurityKey(386) = 179
874     SecurityKey(387) = 183
876     SecurityKey(388) = 167
878     SecurityKey(389) = 209
880     SecurityKey(390) = 147
882     SecurityKey(391) = 252
884     SecurityKey(392) = 102
886     SecurityKey(393) = 46
888     SecurityKey(394) = 243
890     SecurityKey(395) = 188
892     SecurityKey(396) = 200
894     SecurityKey(397) = 96
896     SecurityKey(398) = 141
898     SecurityKey(399) = 149
900     SecurityKey(400) = 131
902     SecurityKey(401) = 155
904     SecurityKey(402) = 222
906     SecurityKey(403) = 230
908     SecurityKey(404) = 13
910     SecurityKey(405) = 200
912     SecurityKey(406) = 52
914     SecurityKey(407) = 142
916     SecurityKey(408) = 84
918     SecurityKey(409) = 111
920     SecurityKey(410) = 7
922     SecurityKey(411) = 247
924     SecurityKey(412) = 176
926     SecurityKey(413) = 218
928     SecurityKey(414) = 140
930     SecurityKey(415) = 83
932     SecurityKey(416) = 22
934     SecurityKey(417) = 120
936     SecurityKey(418) = 136
938     SecurityKey(419) = 38
940     SecurityKey(420) = 142
942     SecurityKey(421) = 127
944     SecurityKey(422) = 98
946     SecurityKey(423) = 5
948     SecurityKey(424) = 231
950     SecurityKey(425) = 213
952     SecurityKey(426) = 125
954     SecurityKey(427) = 157
956     SecurityKey(428) = 169
958     SecurityKey(429) = 49
960     SecurityKey(430) = 196
962     SecurityKey(431) = 246
964     SecurityKey(432) = 75
966     SecurityKey(433) = 125
968     SecurityKey(434) = 135
970     SecurityKey(435) = 249
972     SecurityKey(436) = 166
974     SecurityKey(437) = 127
976     SecurityKey(438) = 133
978     SecurityKey(439) = 49
980     SecurityKey(440) = 170
982     SecurityKey(441) = 185
984     SecurityKey(442) = 74
986     SecurityKey(443) = 206
988     SecurityKey(444) = 80
990     SecurityKey(445) = 142
992     SecurityKey(446) = 187
994     SecurityKey(447) = 239
996     SecurityKey(448) = 207
998     SecurityKey(449) = 165
1000     SecurityKey(450) = 239
1002     SecurityKey(451) = 33
1004     SecurityKey(452) = 19
1006     SecurityKey(453) = 147
1008     SecurityKey(454) = 64
1010     SecurityKey(455) = 34
1012     SecurityKey(456) = 107
1014     SecurityKey(457) = 180
1016     SecurityKey(458) = 162
1018     SecurityKey(459) = 235
1020     SecurityKey(460) = 130
1022     SecurityKey(461) = 89
1024     SecurityKey(462) = 52
1026     SecurityKey(463) = 238
1028     SecurityKey(464) = 144
1030     SecurityKey(465) = 41
1032     SecurityKey(466) = 21
1034     SecurityKey(467) = 157
1036     SecurityKey(468) = 209
1038     SecurityKey(469) = 193
1040     SecurityKey(470) = 121
1042     SecurityKey(471) = 43
1044     SecurityKey(472) = 54
1046     SecurityKey(473) = 158
1048     SecurityKey(474) = 252
1050     SecurityKey(475) = 150
1052     SecurityKey(476) = 91
1054     SecurityKey(477) = 61
1056     SecurityKey(478) = 53
1058     SecurityKey(479) = 229
1060     SecurityKey(480) = 186
1062     SecurityKey(481) = 128
1064     SecurityKey(482) = 143
1066     SecurityKey(483) = 174
1068     SecurityKey(484) = 30
1070     SecurityKey(485) = 84
1072     SecurityKey(486) = 84
1074     SecurityKey(487) = 220
1076     SecurityKey(488) = 90
1078     SecurityKey(489) = 145
1080     SecurityKey(490) = 11
1082     SecurityKey(491) = 175
1084     SecurityKey(492) = 58
1086     SecurityKey(493) = 33
1088     SecurityKey(494) = 4
1090     SecurityKey(495) = 4
1092     SecurityKey(496) = 186
1094     SecurityKey(497) = 101
1096     SecurityKey(498) = 49
1098     SecurityKey(499) = 215
1100     SecurityKey(500) = 118
         '<EhFooter>
         Exit Sub

Initialize_Security_Err:
         LogError Err.description & vbCrLf & _
                "in ServidorArgentum.mSecurity.Initialize_Security " & _
                "at line " & Erl
         
         '</EhFooter>
End Sub


Private Function PacketID_Change(ByVal Selected As Byte) As Integer
        '<EhHeader>
        On Error GoTo PacketID_Change_Err
        '</EhHeader>
    
        Dim Temp     As Integer

        Dim KeyText  As String

        Dim KeyValue As String
    
100     Select Case Selected

            Case 75
102             KeyValue = "GAHBDEWIDKFLSQ2DIWJNE"

104         Case 150
106             KeyValue = "AGSQEFHFFDFSDQETUHFLSJNE"

108         Case 99
110             KeyValue = "13SDDJS2s"

112         Case 105
114             KeyValue = "ADSDEWEFFDFGRT"
        End Select
    
116     Temp = 127
118     Temp = Temp Xor 45
    
120     If Len(KeyValue) > 10 Then
122         Temp = Temp Xor 4 Xor Selected
        Else
124         Temp = Temp Xor 75
        End If
    
        '<EhFooter>
        Exit Function

PacketID_Change_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mSecurity.PacketID_Change " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Public Function ReadPacketID(ByVal PacketID As Integer) As Integer
        '<EhHeader>
        On Error GoTo ReadPacketID_Err
        '</EhHeader>
    
        Dim KeyTempOne   As Integer

        Dim KeyTempTwo   As Integer

        Dim KeyTempThree As Integer
    
100     Dim KeyOne       As String: KeyOne = "137"

102     Dim KeyTwo       As String: KeyTwo = "215"

104     Dim KeyThree     As String: KeyThree = "45"

106     Dim KeyFour      As String: KeyFour = "12"

108     Dim KeyFive      As String: KeyFive = "197"
    
110     PacketID = PacketID Xor 127
112     KeyTempOne = 127
114     PacketID = PacketID Xor 67
116     PacketID = PacketID Xor Len(KeyOne)
118     KeyTempOne = KeyTempOne Xor 12
    
120     PacketID = PacketID Xor PacketID_Change(99)
    
122     If PacketID Then
124         PacketID = PacketID Xor Len(KeyTwo)
126         PacketID = PacketID Xor Len(KeyThree)
        
128         PacketID = PacketID Xor PacketID_Change(75)
        Else
130         PacketID = PacketID Xor Len(KeyOne)
132         PacketID = PacketID Xor Len(KeyThree)
134         PacketID = PacketID Xor PacketID_Change(99)
        End If
    
136     KeyTempOne = KeyTempOne Xor PacketID
    
138     If KeyTempOne > 55 Then
140         KeyTempTwo = KeyTempTwo Xor 49
142         KeyTempThree = KeyTempThree Xor 75
144     ElseIf KeyTempOne > 150 Then
146         KeyTempTwo = KeyTempTwo Xor 49
148         KeyTempThree = KeyTempThree Xor 75
150     ElseIf KeyTempOne > 250 Then
152         KeyTempTwo = KeyTempTwo Xor 49
        End If
    
154     PacketID = PacketID Xor KeyOne
156     KeyTempTwo = KeyTempTwo Xor KeyTempOne Xor PacketID_Change(150)
158     KeyTempThree = KeyTempOne Xor KeyTempTwo
160     PacketID = PacketID Xor 75 Xor PacketID_Change(105)
    
162     KeyTempTwo = PacketID Xor KeyTempThree
164     PacketID = PacketID Xor 21
    
166     PacketID = PacketID Xor Len(KeyFive)
    
168     ReadPacketID = PacketID
        '<EhFooter>
        Exit Function

ReadPacketID_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.mSecurity.ReadPacketID " & _
               "at line " & Erl
        
        '</EhFooter>
End Function
