Attribute VB_Name = "GameIni"
'Exodo Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
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
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Public Type tCabecera 'Cabecera de los con

    Desc As String * 255
    CRC As Long
    MagicWord As Long

End Type

Public Type tGameIni

    Puerto As Long
    Musica As Byte
    FX As Byte
    tip As Byte
    Password As String
    Name As String
    DirGraficos As String
    DirSonidos As String
    DirMusica As String
    DirMapas As String
    NumeroDeBMPs As Long
    NumeroMapas As Integer

End Type


Public Const MAX_SETUP_MODS As Byte = 11

Public Enum eSetupMods
    SETUP_MODOVSYNC = 1   ' Activa Modo vsync
    SETUP_CURSORES = 2 ' Desactiva los cursores Gráficos
    SETUP_SOUND3D = 3 ' Fuerza a realizar efectos 3D sobre los efectos principales del juego
    SETUP_PERSONAJEOCULTOENINVI = 4 ' Ve el personaje sin intermitencia (evita entorpecer al usuario de saber donde esta)
    SETUP_MOVERPANTALLA = 5 ' Permite mover la pantalla cuando no está en 800x600
    SETUP_MOVERSEHABLAR = 6 ' Permite No moverse al hablar
    SETUP_BOTONLANZAR = 7 ' Lanzar el hechizo con intervalo respetado
    SETUP_MASTERSOUND = 8 ' Quitar todos los Sonidos
    SETUP_PANTALLACOMPLETA = 9  'Pantalla Completa
    SETUP_INTERFAZTDS = 10          'Interfaz TDS
    SETUP_INTERFAZMODERNA = 11  'Interfaz Moderna
End Enum

Public Type tSetupMods

    
    bMasterSound As Byte
    bSoundMusic     As Byte
    bSoundEffect    As Byte
    bSoundInterface As Byte
    
    bValueSoundMusic As Byte
    bValueSoundEffect As Byte
    bValueSoundInterface As Byte
    bValueSoundMaster As Byte
    
    bResolution As Byte
    
    bFps As Integer
    bAlpha As Integer
    bConfig(1 To MAX_SETUP_MODS)   As Byte ' Configs rapidas

    CursorGeneral As String
    CursorSpell As String
    CursorInv As String
    CursorHand As String
End Type

Public ClientSetup   As tSetupMods

Public MiCabecera    As tCabecera

Public Config_Inicio As tGameIni

Public Sub IniciarCabecera(ByRef Cabecera As tCabecera)
    Cabecera.Desc = "Exodo Online by Lautaro"
    Cabecera.CRC = Rnd * 100
    Cabecera.MagicWord = Rnd * 10
End Sub

Public Function LeerGameIni() As tGameIni

    Dim N       As Integer

    Dim GameIni As tGameIni

    N = FreeFile
    Open IniPath & "Inicio.con" For Binary As #N
    Get #N, , MiCabecera
    
    Get #N, , GameIni
    
    Close #N
    LeerGameIni = GameIni
End Function

Public Sub EscribirGameIni(ByRef GameIniConfiguration As tGameIni)
    On Local Error Resume Next

    Dim N As Integer

    N = FreeFile
    Open IniPath & "Inicio.con" For Binary As #N
    Put #N, , MiCabecera
    Put #N, , GameIniConfiguration
    Close #N
End Sub

