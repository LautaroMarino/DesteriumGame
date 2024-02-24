VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Desterium ModTDS"
   ClientHeight    =   7650
   ClientLeft      =   1950
   ClientTop       =   1515
   ClientWidth     =   11070
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000004&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7650
   ScaleWidth      =   11070
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.CommandButton Command10 
      Caption         =   "Command10"
      Height          =   495
      Left            =   5640
      TabIndex        =   33
      Top             =   6000
      Width           =   2535
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Entrega de tunicas"
      Height          =   2655
      Left            =   5160
      TabIndex        =   13
      Top             =   1440
      Width           =   4695
      Begin VB.CommandButton cmbOtorgar 
         BackColor       =   &H00FFC0FF&
         Caption         =   "DAR TITAN"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txtObj 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   360
         TabIndex        =   30
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   360
         TabIndex        =   29
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         Height          =   285
         Left            =   2040
         TabIndex        =   32
         Top             =   480
         Visible         =   0   'False
         Width           =   1485
      End
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Command9"
      Height          =   255
      Left            =   5400
      TabIndex        =   28
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Evento Testin"
      Height          =   360
      Left            =   6480
      TabIndex        =   27
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox txtMascota 
      Height          =   315
      Left            =   8400
      TabIndex        =   26
      Text            =   "1"
      Top             =   4320
      Width           =   1935
   End
   Begin VB.CommandButton Command7 
      Caption         =   "SUM MIS MASCOTAS"
      Height          =   360
      Left            =   8160
      TabIndex        =   25
      Top             =   5280
      Width           =   2655
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Agregar MASCOTA"
      Height          =   360
      Left            =   8400
      TabIndex        =   24
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton cmbSHOP 
      BackColor       =   &H00FFC0FF&
      Caption         =   "SHOP"
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Borrar Ofertas"
      Height          =   255
      Left            =   8520
      TabIndex        =   18
      Top             =   3720
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "cerrar invasion"
      Height          =   615
      Left            =   8520
      TabIndex        =   17
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "iniciar invasion"
      Height          =   615
      Left            =   8520
      TabIndex        =   16
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Happy Hour"
      Height          =   1230
      Left            =   105
      TabIndex        =   14
      Top             =   4305
      Width           =   2775
      Begin VB.CheckBox chkParty 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Party Time (25% BONUS)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   110
         TabIndex        =   23
         Top             =   480
         Width           =   2505
      End
      Begin VB.CheckBox chkHappy 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Exp x2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   105
         TabIndex        =   15
         Top             =   210
         Width           =   1065
      End
      Begin VB.Label lblBots 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1695
         TabIndex        =   21
         Top             =   270
         Width           =   315
      End
      Begin VB.Label lblAyudin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   345
         Index           =   0
         Left            =   2175
         TabIndex        =   20
         Top             =   210
         Width           =   255
      End
      Begin VB.Label lblAyudin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   345
         Index           =   1
         Left            =   1365
         TabIndex        =   19
         Top             =   210
         Width           =   255
      End
   End
   Begin VB.TextBox txtNumUsers 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "0"
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdSystray 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Systray"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton cmdCerrarServer 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Cerrar Servidor"
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6600
      Width           =   3495
   End
   Begin VB.CommandButton cmdConfiguracion 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Configuración General"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6120
      Width           =   4935
   End
   Begin VB.Timer tPiqueteC 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   3000
      Top             =   2580
   End
   Begin VB.CommandButton cmdDump 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Crear Log Crítico de Usuarios"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5640
      Width           =   4935
   End
   Begin VB.Timer FX 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   3960
      Top             =   2580
   End
   Begin VB.Timer Auditoria 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3960
      Top             =   3060
   End
   Begin VB.Timer GameTimer 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   3960
      Top             =   2100
   End
   Begin VB.Timer AutoSave 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   3000
      Top             =   3060
   End
   Begin VB.Timer KillLog 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   4440
      Top             =   2100
   End
   Begin VB.Timer TIMER_AI 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   4455
      Top             =   2580
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Mensajea todos los clientes (Solo testeo)"
      Height          =   3615
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4935
      Begin VB.Timer tControlHechizos 
         Enabled         =   0   'False
         Left            =   360
         Top             =   1920
      End
      Begin VB.Timer TimerGuardarUsuarios 
         Enabled         =   0   'False
         Interval        =   3000
         Left            =   360
         Top             =   1440
      End
      Begin VB.Timer TimerFlush 
         Interval        =   10
         Left            =   1680
         Top             =   1440
      End
      Begin VB.TextBox txtChat 
         BackColor       =   &H00C0FFFF&
         Height          =   2175
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   1320
         Width           =   4695
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Enviar por Consola"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Enviar por Pop-Up"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox BroadMsg 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Label Escuch 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   285
      Left            =   3840
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label CantUsuarios 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Número de usuarios jugando:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   2460
   End
   Begin VB.Label txStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   15
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuMostrar 
         Caption         =   "&Mostrar"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.12.2
'Copyright (C) 2002 Márquez Pablo Ignacio
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
'Argentum Online is based on Baronsoft's VB6 Online RPG
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

Private tHechizosMinutesCounter As Byte
Public ESCUCHADAS As Long

Private Type NOTIFYICONDATA

    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64

End Type
   
Const NIM_ADD = 0

Const NIM_DELETE = 2

Const NIF_MESSAGE = 1

Const NIF_ICON = 2

Const NIF_TIP = 4

Const WM_MOUSEMOVE = &H200

Const WM_LBUTTONDBLCLK = &H203

Const WM_RBUTTONUP = &H205

Private Declare Function GetWindowThreadProcessId _
                Lib "user32" (ByVal hWnd As Long, _
                              lpdwProcessId As Long) As Long

Private Declare Function Shell_NotifyIconA _
                Lib "SHELL32" (ByVal dwMessage As Long, _
                               lpData As NOTIFYICONDATA) As Integer

Private Function setNOTIFYICONDATA(hWnd As Long, _
                                   ID As Long, _
                                   flags As Long, _
                                   CallbackMessage As Long, _
                                   Icon As Long, _
                                   Tip As String) As NOTIFYICONDATA
        '<EhHeader>
        On Error GoTo setNOTIFYICONDATA_Err
        '</EhHeader>

        Dim nidTemp As NOTIFYICONDATA

100     nidTemp.cbSize = Len(nidTemp)
102     nidTemp.hWnd = hWnd
104     nidTemp.uID = ID
106     nidTemp.uFlags = flags
108     nidTemp.uCallbackMessage = CallbackMessage
110     nidTemp.hIcon = Icon
112     nidTemp.szTip = Tip & Chr$(0)

114     setNOTIFYICONDATA = nidTemp
        '<EhFooter>
        Exit Function

setNOTIFYICONDATA_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.frmMain.setNOTIFYICONDATA " & _
               "at line " & Erl
        
        '</EhFooter>
End Function

Sub CheckIdleUser()
        '<EhHeader>
        On Error GoTo CheckIdleUser_Err
        '</EhHeader>

        Dim iUserIndex As Long
        Dim CheckingMapsOns As Boolean
        Static MinutosOns As Integer
        
        MinutosOns = MinutosOns + 1
        
        If MinutosOns = 30 Then
            CheckingMapsOns = True
            MinutosOns = 0
        End If
        
100     For iUserIndex = 1 To LastUser
102         If iUserIndex <> SLOT_TERMINAL_ARCHIVE Then
         
104             With UserList(iUserIndex)
                    
                    
                    
                    
                    'Actualiza el contador de inactividad
106                 If .flags.Traveling = 0 Then
108                     .Counters.IdleCount = .Counters.IdleCount + 1
                    End If
                    
                    'Conexion activa? y es un usuario loggeado?
110                 If .flags.UserLogged Then
                        ' # Chequea el on fire del mapa
                        Map_CheckFire .Pos.Map
                        
                        ' Chequea si está en un mapa por horario y lo regresa a la ciudad principal
                        If Not CheckMap_HourDay(iUserIndex, .Pos) Then
                            Call EventWarpUser(iUserIndex, Ullathorpe.Map, Ullathorpe.X, Ullathorpe.Y)
                        End If
                            
                        If CheckingMapsOns Then
                            If Not CheckMap_Onlines(iUserIndex, .Pos) Then
                                Call EventWarpUser(iUserIndex, Ullathorpe.Map, Ullathorpe.X, Ullathorpe.Y)
                            End If
                        End If
                        
                        
112                     If .Counters.IdleCount >= IdleLimit Then
114                         Call WriteShowMessageBox(iUserIndex, "Demasiado tiempo inactivo. Has sido desconectado.")

                            'mato los comercios seguros
116                         If .ComUsu.DestUsu > 0 Then
118                             If UserList(.ComUsu.DestUsu).flags.UserLogged Then
120                                 If UserList(.ComUsu.DestUsu).ComUsu.DestUsu = iUserIndex Then
122                                     Call WriteConsoleMsg(.ComUsu.DestUsu, "Comercio cancelado por el otro usuario.", FontTypeNames.FONTTYPE_TALK)
124                                     Call FinComerciarUsu(.ComUsu.DestUsu)
126                                     Call FlushBuffer(.ComUsu.DestUsu) 'flush the buffer to send the message right away
                                    End If
                                End If

128                             Call FinComerciarUsu(iUserIndex)
                            End If
                    
                            'Kick player ( and leave character inside :D )!
130                         Call WriteDisconnect(iUserIndex)
132                         Call FlushBuffer(iUserIndex)
134                         Call CloseSocket(iUserIndex)
            
                        End If
                    End If
            
                End With
            End If
136     Next iUserIndex

        '<EhFooter>
        Exit Sub

CheckIdleUser_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.frmMain.CheckIdleUser " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub Auditoria_Timer()

    On Error GoTo errhand

    ' Sistema de eventos segundos
    LoopEvent                        ' Desarrollo de los eventos en curso.
    'Events_Automatic_Loop     ' Eventos que hace JARVIS
    Events_Loop_Check           ' Eventos DEFINIDOS= TITAN DEL MES
    ''''''''''''''''''''''''''''''
    Static centinelSecs As Byte

    centinelSecs = centinelSecs + 1

    If centinelSecs = 5 Then
        Call RetoFast_Loop
        centinelSecs = 0
    End If
    
    Call PasarSegundo 'sistema de desconexion de 10 segs

    Exit Sub

errhand:

    Call LogError("Error en Timer Auditoria. Err: " & Err.description & " - " & Err.number)

    

End Sub
Private Function CheckMultipleTimes() As Boolean
    ' Definir las horas y minutos objetivo en un array
    Dim targetTimes As Variant
    targetTimes = Array("00:00")

    ' Obtener la hora y los minutos actuales
    Dim currentHour As Integer
    Dim currentMinute As Integer
    currentHour = Hour(Now)
    currentMinute = Minute(Now)

    ' Convertir la hora actual a un formato comparable
    Dim currentTime As String
    currentTime = Format(Now, "hh:mm")

    ' Verificar si la hora actual está dentro de la ventana permitida de 2 minutos para alguna de las horas objetivo
    If IsInArray(currentTime, targetTimes) Then
        ' Verificar si han pasado al menos 2 minutos desde la última ejecución
        If DateDiff("n", lastRunTime, Now) >= 2 Then
            ' Actualizar el tiempo de la última ejecución
            lastRunTime = Now
            CheckMultipleTimes = True
        Else
            ' No han pasado 2 minutos desde la última ejecución
            CheckMultipleTimes = False
        End If
    Else
        ' La hora actual no está dentro de la ventana permitida
        CheckMultipleTimes = False
    End If
End Function

Private Function IsInArray(valueToFind As Variant, arr As Variant) As Boolean
    ' Función auxiliar para verificar si un valor está en un array
    Dim element As Variant
    For Each element In arr
        If element = valueToFind Then
            IsInArray = True
            Exit Function
        End If
    Next element
    IsInArray = False
End Function
Private Function CheckMidnight() As Boolean
    ' Obtener la hora y los minutos actuales
    Dim currentHour As Integer
    Dim currentMinute As Integer
    currentHour = Hour(Now)
    currentMinute = Minute(Now)

    ' Definir la hora objetivo (00:00)
    Dim targetHour As Integer
    Dim targetMinute As Integer
    targetHour = 0
    targetMinute = 0

    ' Verificar si la hora actual está dentro de la ventana permitida de 2 minutos
    If currentHour = targetHour And Abs(currentMinute - targetMinute) <= 1 Then
        ' Verificar si han pasado al menos 2 minutos desde la última ejecución
        If DateDiff("n", lastRunTime, Now) >= 2 Then
            ' Actualizar el tiempo de la última ejecución
            lastRunTime = Now
            CheckMidnight = True
        Else
            ' No han pasado 2 minutos desde la última ejecución
            CheckMidnight = False
        End If
    Else
        ' La hora actual no está dentro de la ventana permitida
        CheckMidnight = False
    End If
End Function

Private Sub AutoSave_Timer()

    On Error GoTo ErrHandler

    'fired every minute
    Static Minutos          As Long

    Static MinutosLatsClean As Long

    Static MinsClearChar    As Long
    
    Static SpamMessage As Long
    
    Minutos = Minutos + 1
    SpamMessage = SpamMessage + 1
    
    ' Spam de mensajes cada 15 minutos
    If SpamMessage >= mMessages.MessageTime Then
        SpamMessage = 0
        mMessages.MessageSpam_SelectedRandom
    End If
    
    
    
    #If Testeo = 0 Then
        ' # Actualizamos la info rapida del servidor
        Call MySql_UpdateServer
    #End If
    
    
   ' If 0 = 1 Then 'CheckMultipleTimes
       ' AutoRestart = True
       ' cmdCerrarServer_Click
        ' Exit Sub
   ' End If
            
    ' Checking Rank Reset
   ' Call mRank.Ranking_Loop_Reset
    
    ' Update Stats Web
    Call WriteUpdateStats

    If Minutos = MinutosWs - 1 Then
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Worldsave en 1 minuto ...", FontTypeNames.FONTTYPE_VENENO))
    End If

    If Minutos >= MinutosWs Then
        Call ES.DoBackUp
        Call SaveRecords
        Call aClon.VaciarColeccion
        Minutos = 0
    End If
    
    If MinutosLatsClean >= 15 Then
        If HappyHour Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡HappyHour Activado! Exp x2 ¡Entrená tu personaje!", FontTypeNames.FONTTYPE_USERBRONCE))
        End If
        
        If PartyTime Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("PartyTime» Los miembros de la party reciben 25% de experiencia extra.", FontTypeNames.FONTTYPE_USERBRONCE))
        End If
        
        If Power.UserIndex = 0 Then
            Call Power_Search_All
        End If
        
        MinutosLatsClean = 0
        Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales
        Call LimpiarMundo
        'CountDownLimpieza = 5
    Else
        MinutosLatsClean = MinutosLatsClean + 1
    End If
    
    Call PurgarPenas
    Call CheckIdleUser

    ' Torneos automáticos cada 30 minutos
    'Call EventosDS.CheckEvent_Time_Auto

    
    

    'Power.Active = Power_CheckTime
    '<<<<<-------- Log the number of users online ------>>>
    Dim N As Integer

    N = FreeFile()
    Open LogPath & "numusers.log" For Output Shared As N
    Print #N, NumUsers + UsersBot
    Close #N
    '<<<<<-------- Log the number of users online ------>>>

    Exit Sub

ErrHandler:
    Call LogError("Error en TimerAutoSave " & Err.number & ": " & Err.description)

    

End Sub



Private Sub chkHappy_Click()
 CheckHappyHour

End Sub

Private Sub chkParty_Click()
    CheckPartyTime
End Sub

Private Sub cmbShop_Click()
        '<EhHeader>
        On Error GoTo cmbShop_Click_Err
        '</EhHeader>
100     Call FrmShop.Show
        '<EhFooter>
        Exit Sub

cmbShop_Click_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.frmMain.cmbShop_Click " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub cmdCerrarServer_Click()
        '<EhHeader>
        On Error GoTo cmdCerrarServer_Click_Err
        '</EhHeader>
            
        If AutoRestart = False Then
100         If MsgBox("¿Está seguro que desea hacer WorldSave, guardar pjs y cerrar?", vbYesNo, "Apagar Magicamente") = vbNo Then Exit Sub
        
102         Me.MousePointer = 11
        End If
        
104     FrmStat.Show
    
        ' Cancelamos retos
106     Call Retos_Reset_All
    
        ' Cancelamos eventos
108     Call Eventos_Reset_All
    
        ' Cancelamos Retos Fast
110     Call Fast_Reset_All
    
        ' Cancelamos la Subasta
112     Call Auction_Cancel
    
        ' Eventos automáticos
114     Call Events_Data_Predetermined

            'commit experiencia
118     Call DistributeExpAndGldGroups
        
        'WorldSave
116     Call ES.DoBackUp

        ' Guardamos todos los usuarios y cuentas conectadas
        Call GuardarUsuarios_Close
    
        Dim A As Long
    
122     For A = 1 To LastUser
            If AutoRestart Then
                Call WriteUpdateClient(A)
                Call FlushBuffer(A)
            End If
            
124         Call Protocol.Kick(A)
        Next
        
        Call Server.Close
        
        If AutoRestart Then
            AutoRestart = False
            Call WriteVar(IniPath & "Server.ini", "INIT", "lastRunTime", lastRunTime)
            Shell App.Path & "\FILEZILLA\UpdateArchiveFTP.exe", vbNormalFocus
            
            If GetCurrentProcessName = "Desterium.exe" Then
                Shell App.Path & "\Desterium1.exe", vbNormalFocus
            Else
                Shell App.Path & "\Desterium.exe", vbNormalFocus
            End If
            
            
        End If
        
        
        
        'Chauuu
126     Unload frmMain

        '<EhFooter>
        Exit Sub

cmdCerrarServer_Click_Err:
        LogError Err.description & vbCrLf & _
               "in cmdCerrarServer_Click " & _
               "at line " & Erl

        '</EhFooter>
End Sub

Private Sub cmdConfiguracion_Click()
    frmServidor.Visible = True
    'Create_Stats_General
End Sub

Private Sub cmdSystray_Click()
    SetSystray
End Sub

Private Sub Command1_Click()
        '<EhHeader>
        On Error GoTo Command1_Click_Err
        '</EhHeader>
100     Call SendData(SendTarget.ToAll, 0, PrepareMessageShowMessageBox(BroadMsg.Text))
        ''''''''''''''''SOLO PARA EL TESTEO'''''''
        ''''''''''SE USA PARA COMUNICARSE CON EL SERVER'''''''''''
102     txtChat.Text = txtChat.Text & vbNewLine & "Servidor> " & BroadMsg.Text
        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.frmMain.Command1_Click " & _
           "at line " & Erl
    
    '</EhFooter>
End Sub

Public Sub InitMain(ByVal f As Byte)

    If f = 1 Then
        Call SetSystray
    Else
        frmMain.Show
    End If

End Sub

Private Sub Command10_Click()

 
            ' Selecciona una frase aleatoria
            Dim randomIndex As Integer
            randomIndex = Int((UBound(FrasesOnFire) + 1) * Rnd)
            Dim Mensaje As String
            Mensaje = Replace(FrasesOnFire(randomIndex), "{Mapa}", "**" & "Dungeon Magma" & "**")
            ' Convierte la cadena a UTF-8 antes de mostrarla
           ' Mensaje = StrConv(Mensaje, vbFromUnicode)

            Call MsgBox(Mensaje)
  'MySql_UpdateServer
    
End Sub

Private Sub Command2_Click()
    
    
    
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> " & BroadMsg.Text, FontTypeNames.FONTTYPE_SERVER, eMessageType.Admin))
    ''''''''''''''''SOLO PARA EL TESTEO'''''''
    ''''''''''SE USA PARA COMUNICARSE CON EL SERVER'''''''''''
    txtChat.Text = txtChat.Text & vbNewLine & "Servidor> " & BroadMsg.Text
End Sub

Private Sub Command3_Click()
'    Call Invations_New(1)


End Sub

Private Sub Command4_Click()
    ' Call Invations_Close(1)
   
    Dim tUser As Integer
   Dim X As Integer
   Dim Y As Integer
   
    tUser = NameIndex("LION")
   
    If tUser > 0 Then
        X = UserList(tUser).Pos.X
        Y = UserList(tUser).Pos.Y
        Call SendData(SendTarget.ToOne, tUser, PrepareMessageCreateFXMap(RandomNumber(X - 5, X + 5), RandomNumber(Y - 5, Y + 5), 126, RandomNumber(1, 5)))

    End If

End Sub

Private Sub Command6_Click()
    
    Dim UserIndex As Integer
    UserIndex = NameIndex("LION")
    
    If UserIndex > 0 Then
        If BotIntelligence_Add(UserIndex, "Carmen", eClass.Mage, eRaza.Humano, 75, True) Then
            Call WriteConsoleMsg(UserIndex, " ¡Has invocado una nueva mascota! Vaya que maravilloso muchacho", FontTypeNames.FONTTYPE_INFOGREEN)
        End If
    End If
End Sub

Private Sub Command7_Click()
    Dim UserIndex As Integer
    UserIndex = NameIndex("LION")
    
    If UserIndex > 0 Then
        If BotIntelligence_Spawn(UserList(UserIndex).BotIntelligence(1), UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, BOT_MOVEMENT_DEFAULT, BOT_MODE_MIXED) Then
            Call WriteConsoleMsg(UserIndex, " ¡Has SPAWNEADO una de tus mascotas! Vaya que maravilloso muchacho", FontTypeNames.FONTTYPE_INFOGREEN)
        End If
        
        If BotIntelligence_Spawn(UserList(UserIndex).BotIntelligence(2), UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, BOT_MOVEMENT_DEFAULT, BOT_MODE_MIXED) Then
            Call WriteConsoleMsg(UserIndex, " ¡Has SPAWNEADO una de tus mascotas! Vaya que maravilloso muchacho", FontTypeNames.FONTTYPE_INFOGREEN)
        End If
    End If
End Sub

Private Sub Command8_Click()
    Events_SetConfig
End Sub

Private Sub Command9_Click()

    Dim Player As tPlayerData
    
    Player.ConsecutiveWins = 1
    Player.GamesPlayed = 54
    Player.GamesWon = 34
    Player.PlayerName = "ABC"
    
    Call Reward_Process_User(10, Player)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        '<EhHeader>
        On Error GoTo Form_MouseMove_Err
        '</EhHeader>
   
100     If Not Visible Then

102         Select Case X \ Screen.TwipsPerPixelX
                
                Case WM_LBUTTONDBLCLK
104                 WindowState = vbNormal
106                 Visible = True

                    Dim hProcess As Long

108                 GetWindowThreadProcessId hWnd, hProcess
110                 AppActivate hProcess

112             Case WM_RBUTTONUP
114                 hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf AppHook, App.hInstance, App.ThreadID)
116                 PopupMenu mnuPopUp

118                 If hHook Then UnhookWindowsHookEx hHook: hHook = 0
            End Select

        End If
   
        '<EhFooter>
        Exit Sub

Form_MouseMove_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.frmMain.Form_MouseMove " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub QuitarIconoSystray()
        '<EhHeader>
        On Error GoTo QuitarIconoSystray_Err
        '</EhHeader>


        'Borramos el icono del systray
        Dim i   As Integer

        Dim nid As NOTIFYICONDATA

100     nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, vbNull, frmMain.Icon, "")

102     i = Shell_NotifyIconA(NIM_DELETE, nid)

        '<EhFooter>
        Exit Sub

QuitarIconoSystray_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.frmMain.QuitarIconoSystray " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo Form_Unload_Err
        '</EhHeader>

        'Save stats!!!
        'Call Statistics.DumpStatistics

100     Call QuitarIconoSystray

        Dim LoopC As Integer

102     For LoopC = 1 To LastUser
104         Call Protocol.Kick(LoopC)
        Next

        'Log
        Dim N As Integer

106     N = FreeFile
108     Open LogPath & "Main.log" For Append Shared As #N
110     Print #N, Date & " " & Time & " server cerrado."
112     Close #N

114     End

116     Set SonidosMapas = Nothing

        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        LogError Err.description & vbCrLf & _
               "in ServidorArgentum.frmMain.Form_Unload " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub FX_Timer()

    On Error GoTo hayerror

    'Call SonidosMapas.ReproducirSonidosDeMapas

    Exit Sub

hayerror:

End Sub

Private Sub GameTimer_Timer()

    '********************************************************
    'Author: Unknown
    'Last Modify Date: -
    '********************************************************
    Dim iUserIndex   As Long

    Dim bEnviarStats As Boolean

    Dim bEnviarAyS   As Boolean
    
    On Error GoTo hayerror

    '<<<<<< Procesa eventos de los usuarios >>>>>>
    For iUserIndex = 1 To LastUser

        With UserList(iUserIndex)

            'Conexion activa?
            If .ConnIDValida Then
                '¿User valido?

                If .flags.UserLogged Then

                    '[Alejo-18-5]
                    bEnviarStats = False
                    bEnviarAyS = False

                    If .flags.Paralizado = 1 Then Call EfectoParalisisUser(iUserIndex)
                    If .flags.Ceguera = 1 Or .flags.Estupidez Then Call EfectoCegueEstu(iUserIndex)
                    
                    If .flags.Muerto = 0 Then
                        
                        If .Counters.BuffoAceleration > 0 Then Call EfectoAceleracion(iUserIndex)
                        
                        If (.flags.Privilegios And PlayerType.User) Then Call EfectoLava(iUserIndex)

                        If (.flags.Privilegios And PlayerType.User) Then Call EfectoFrio(iUserIndex)

                        'If .flags.Desnudo <> 0 And (.flags.Privilegios And PlayerType.User) <> 0 Then Call EfectoFrio(iUserIndex)

                        If .flags.Meditando Then Call DoMeditar(iUserIndex)

                        If .flags.Envenenado <> 0 And (.flags.Privilegios And PlayerType.User) <> 0 Then Call EfectoVeneno(iUserIndex)
                        
                        If .flags.Incinerado <> 0 And (.flags.Privilegios And PlayerType.User) <> 0 Then Call User_EfectoIncineracion(iUserIndex)
                        
                        If .flags.AdminInvisible <> 1 Then
                            If .flags.Invisible = 1 Then Call EfectoInvisibilidad(iUserIndex)
                            If .flags.Oculto = 1 Then Call DoPermanecerOculto(iUserIndex)
                        End If

                        If .flags.Mimetizado = 1 Or .flags.Transform = 1 Or .flags.TransformVIP Then Call EfectoMimetismo(iUserIndex)

                        Call DuracionPociones(iUserIndex)

                        Call HambreYSed(iUserIndex, bEnviarAyS)

                        If .flags.Hambre = 0 And .flags.Sed = 0 Then
                            If Not .flags.Descansar Then
                                Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloSinDescansar)

                                If bEnviarStats Then
                                    Call WriteUpdateSta(iUserIndex)
                                    bEnviarStats = False
                                End If

                            Else
                                Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloDescansar)

                                If bEnviarStats Then
                                    Call WriteUpdateSta(iUserIndex)
                                    bEnviarStats = False
                                End If

                                'termina de descansar automaticamente
                                If .Stats.MaxSta = .Stats.MinSta Then
                                    Call WriteRestOK(iUserIndex)
                                    'Call WriteConsoleMsg(iUserIndex, "Has terminado de descansar.", FontTypeNames.FONTTYPE_INFO)
                                    .flags.Descansar = False
                                End If

                            End If
                        End If

                        If bEnviarAyS Then Call WriteUpdateHungerAndThirst(iUserIndex)

                        If .MascotaIndex > 0 Then Call TiempoInvocacion(iUserIndex)
                    Else

                        If .flags.Traveling <> 0 Then Call TravelingEffect(iUserIndex)
                    End If 'Muerto

                End If 'UserLogged

                'If there is anything to be sent, we send it
                Call FlushBuffer(iUserIndex)
            End If

        End With

    Next iUserIndex

    Exit Sub

hayerror:
    LogError ("Error en GameTimer: " & Err.description & " UserIndex = " & iUserIndex & " en linea " & Erl)
End Sub

Private Sub lblAyudin_Click(Index As Integer)

    Select Case Index

        Case 1

            If UsersBot = 0 Then Exit Sub
            lblBots.Caption = val(lblBots.Caption) - 1
            UsersBot = UsersBot - 1

        Case 0
            lblBots.Caption = val(lblBots.Caption) + 1
            UsersBot = UsersBot + 1
    End Select
    
End Sub

Private Sub mnusalir_Click()
    Call cmdCerrarServer_Click
End Sub

Private Sub KillLog_Timer()
    '<EhHeader>
    On Error GoTo KillLog_Timer_Err
    '</EhHeader>

    If FileExist(LogPath & "connect.log", vbNormal) Then Kill LogPath & "connect.log"
    If FileExist(LogPath & "haciendo.log", vbNormal) Then Kill LogPath & "haciendo.log"
    If FileExist(LogPath & "stats.log", vbNormal) Then Kill LogPath & "stats.log"
    If FileExist(LogPath & "Asesinatos.log", vbNormal) Then Kill LogPath & "Asesinatos.log"
    If FileExist(LogPath & "HackAttemps.log", vbNormal) Then Kill LogPath & "HackAttemps.log"

    '<EhFooter>
    Exit Sub

KillLog_Timer_Err:
    LogError Err.description & vbCrLf & _
           "in ServidorArgentum.frmMain.KillLog_Timer " & _
           "at line " & Erl
    
    '</EhFooter>
End Sub

Private Sub SetSystray()

    Dim i   As Integer

    Dim S   As String

    Dim nid As NOTIFYICONDATA
    
    S = "ARGENTUM-ONLINE"
    nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, frmMain.Icon, S)
    i = Shell_NotifyIconA(NIM_ADD, nid)
        
    If WindowState <> vbMinimized Then WindowState = vbMinimized
    Visible = False

End Sub

Private Sub tControlHechizos_Timer()
    '<EhHeader>
    On Error GoTo ErrHandler
    '</EhHeader>

Dim UserIndex As Integer
    'Reseteo control de hechizos
    tHechizosMinutesCounter = tHechizosMinutesCounter + 1
    
    If tHechizosMinutesCounter = 2 Then
        For UserIndex = 1 To LastUser
            With UserList(UserIndex)
                UserList(UserIndex).Counters.controlHechizos.HechizosTotales = 0
                UserList(UserIndex).Counters.controlHechizos.HechizosCasteados = 0
            End With
        Next UserIndex
        tHechizosMinutesCounter = 0
    End If
    
    Exit Sub
ErrHandler:
    
End Sub

Private Sub TIMER_AI_Timer()

    On Error GoTo ErrorHandler

    Dim NpcIndex As Long

    Dim mapa     As Integer

    Dim e_p      As Integer
    
    'Barrin 29/9/03
    If Not haciendoBK And Not EnPausa Then

        'Update NPCs
        For NpcIndex = 1 To LastNPC
            
            With Npclist(NpcIndex)

                If .flags.NPCActive Then 'Nos aseguramos que sea INTELIGENTE!
                    
                    ' Chequea si contiua teniendo dueño
                    If .Owner > 0 Then Call ValidarPermanenciaNpc(NpcIndex)
                
                    If .flags.Incinerado > 0 Then Call Npc_EfectoIncineracion(NpcIndex)
                    
                    If .flags.Paralizado = 1 Then
                        Call EfectoParalisisNpc(NpcIndex)
                    Else
                        
                        ' Preto? Tienen ai especial
                        If .NPCtype = eNPCType.Pretoriano Then
                           ' If .ClanIndex Then
                               ' If Intervalo_CriatureVelocity(NpcIndex) Then
                                  '  Call ClanPretoriano(.ClanIndex).PerformPretorianAI(NpcIndex)

                              '  End If

                           ' End If
                        
                        Else
                            
                            'Usamos AI si hay algun user en el mapa
                            If .flags.Inmovilizado = 1 Then
                                Call EfectoParalisisNpc(NpcIndex)

                            End If
                                
                            mapa = .Pos.Map
                            
                            If mapa > 0 Then
                                If MapInfo(mapa).NumUsers > 0 Then
                                    
                                    If .Movement = IntelligenceMax Then
                                        Call BotIntelligence_AI(NpcIndex)
                                    ElseIf .Movement <> TipoAI.Estatico Then

                                        If Intervalo_CriatureVelocity(NpcIndex) Then
                                            Call NpcAI(NpcIndex)

                                        End If

                                    End If

                                End If

                            End If

                        End If
                        
                    End If

                End If

            End With

        Next NpcIndex

    End If
    
    Exit Sub

ErrorHandler:
    Call LogError("Error en TIMER_AI_Timer " & Npclist(NpcIndex).Name & " mapa:" & Npclist(NpcIndex).Pos.Map)
    Call MuereNpc(NpcIndex, 0)

End Sub

Private Sub TimerFlush_Timer()

    On Error GoTo ErrHandler

    Dim A As Long
    
    For A = 1 To LastUser

        With UserList(A)
            If .ConnIDValida Then
                If .flags.UserLogged Then
                    Call FlushBuffer(A)
                End If
            End If
        End With
    Next A
    
        Exit Sub

ErrHandler:
    Call LogError("Error en TimerFlush_Timer " & Err.number & ": " & Err.description & ": Linea: " & Erl)
End Sub

Public Sub TimerGuardarUsuarios_Timer()

On Error GoTo ErrHandler
    Dim i             As Integer

    Dim NotSave       As Boolean
   
    Dim UserGuardados As Long
    Dim EventClass As Long
    Dim Save As Boolean
    
    
    For i = 1 To LastUser
2
        Save = True
        
        If UserList(i).flags.UserLogged Then
4

            If GetTime - UserList(i).Counters.LastSave > IntervaloGuardarUsuarios Then
6
                EventClass = UserList(i).flags.SlotEvent
                
                If EventClass > 0 Then
                    If Events(EventClass).ChangeClass > 0 Then
                        Save = False
                    End If
                End If
                
                If Save Then
                    If UserList(i).flags.ModoStream Then
                        If UserList(i).flags.StreamUrl = vbNullString Then
                            Call WriteConsoleMsg(i, "Setea una URL con /STREAMURL asi se spamea por consola", FontTypeNames.FONTTYPE_STREAM)
                        Else
                            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("EN VIVO» " & UserList(i).flags.StreamUrl, FontTypeNames.FONTTYPE_STREAM))
                        End If
8
                    End If
10
                    Call SaveUser(UserList(i), CharPath & UCase$(UserList(i).Name) & ".chr", False)
                    Call SaveDataAccount(i, UserList(i).Account.Email, UserList(i).IpAddress)
12
                    If Not EsGm(i) Then Call WriteUpdateUserData(UserList(i))
                    
                    UserGuardados = UserGuardados + 1
14
                    If UserGuardados > NumUsers / IntervaloGuardarUsuarios * IntervaloTimerGuardarUsuarios Then Exit For
16             End If
            End If

        End If

    Next i
18
    Exit Sub

ErrHandler:
    Call LogError("Error en TimerGuardarUsuarios_Timer " & Err.number & ": " & Err.description & ": Linea: " & Erl)
End Sub

Private Sub tPiqueteC_Timer()

    Dim NuevaA As Boolean

    Dim NuevoL As Boolean

    Dim GI     As Integer

    Dim i      As Long
    
    On Error GoTo ErrHandler

    For i = 1 To LastUser

        With UserList(i)

            If .flags.UserLogged Then
                If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.ANTIPIQUETE Then
                    If .flags.Muerto = 0 Then
                        .Counters.PiqueteC = .Counters.PiqueteC + 1
                        
                        If .flags.SlotEvent > 0 Then
                            If .Counters.PiqueteC > 5 Then
                                If Events(.flags.SlotEvent).Modality = eModalityEvent.Teleports Then
                                    Call WriteConsoleMsg(i, "¡¡¡Estás obstruyendo el teleport. Has sido respawneado!!!", FontTypeNames.FONTTYPE_INFO)
                                    Call EventWarpUser(i, 65, 25, 45)
                                End If
                            End If
                        Else
                            Call WriteConsoleMsg(i, "¡¡¡Estás obstruyendo la vía pública, muévete o serás encarcelado!!!", FontTypeNames.FONTTYPE_INFO)
                            
                            If .Counters.PiqueteC > 23 Then
                                .Counters.PiqueteC = 0
                                Call Encarcelar(i, TIEMPO_CARCEL_PIQUETE)
                            End If
                        End If
                        
                    Else
                        .Counters.PiqueteC = 0
                    End If

                Else
                    .Counters.PiqueteC = 0
                End If

                Call FlushBuffer(i)
            End If

        End With

    Next i

    Exit Sub

ErrHandler:
    Call LogError("Error en tPiqueteC_Timer " & Err.number & ": " & Err.description & ": Linea: " & Erl)
End Sub

