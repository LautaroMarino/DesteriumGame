VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'------------------------------------------------------------------------------
' cSubclass                                                         (21/Mar/99)
'
' Clase para subclasificar ventanas (formularios o controles con hWnd)
'
' Última revisión:  08/Ago/99
' Revisado (probado) para publicar en mis páginas                   (18/Abr/01)
'
' ©Guillermo 'guille' Som, 1999-2001
'------------------------------------------------------------------------------
' Basado en el código de un ejemplo del 26/Jun/98
' para detectar los mensajes del menú seleccionado.
'
' También he usado como fuente de inspiración un artículo de Francesco Balena:
'   Subclass Forms to Create New Events
' Aunque no me gustó la forma de enfocar la forma de gestionar varias instancias
' de la clase y me decidí por el enfoque del array.
' La verdad es que algunos de los eventos producidos están "inspirados" en el
' artículo de Balena.
' Todo hay que decirlo: La clase casi la terminé antes de conseguir el código
' del citado artículo, si hubiese tenido ese código antes, seguramente la
' hubiese hecho de otra forma... ¡menos mal que no fue así!
' ya que al menos me "esforcé" en hacerlo por mis medios...
'------------------------------------------------------------------------------
Option Explicit

' Para los TwipsPerPixels del objeto Screen
Private twX As Long, twY As Long
' Para poder interceptar el mensaje WM_MOUSELEAVE
Private mtTME As tTRACKMOUSEEVENT

Private mWndType As Object              ' La ventana que se subclasifica
Private mhWnd As Long                   ' El hWnd de la ventana subclasificada
Private mPrevWndProc As Long            ' La función anterior de procesamiento
                                        ' de mensajes de la ventana a subclasificar

'------------------------------------------------------------------------------
' Enumeración con los mensajes de windows (Window Messages)
'
' Esta lista está sacada de WinUser.h,
' algunas declaraciones están en el fichero Win32API.txt
'
' En los casos que se indique #if ... es que son para otras versiones de Windows:
' #if(WINVER >= 0x0400)     Será Windows NT 4 y superior y Windows 98
' #if(WINVER >= 0x0500)     Será Windows 2000 y superior (Windows NT 5) (creo)
Public Enum eWSCWM

    WM_NULL = &H0
    WM_CREATE = &H1
    WM_DESTROY = &H2
    WM_MOVE = &H3
    WM_SIZE = &H5
    
    WM_ACTIVATE = &H6
    
    WM_SETFOCUS = &H7
    WM_KILLFOCUS = &H8
    WM_ENABLE = &HA
    WM_SETREDRAW = &HB
    WM_SETTEXT = &HC
    WM_GETTEXT = &HD
    WM_GETTEXTLENGTH = &HE
    WM_PAINT = &HF
    WM_CLOSE = &H10
    WM_QUERYENDSESSION = &H11
    WM_QUIT = &H12
    WM_QUERYOPEN = &H13
    WM_ERASEBKGND = &H14
    WM_SYSCOLORCHANGE = &H15
    WM_ENDSESSION = &H16
    WM_SHOWWINDOW = &H18
    WM_WININICHANGE = &H1A
    ' #if(WINVER >= 0x0400)
    WM_SETTINGCHANGE = WM_WININICHANGE
    ' #endif /* WINVER >= 0x0400 */
    WM_DEVMODECHANGE = &H1B
    WM_ACTIVATEAPP = &H1C
    WM_FONTCHANGE = &H1D
    WM_TIMECHANGE = &H1E
    WM_CANCELMODE = &H1F
    WM_SETCURSOR = &H20
    WM_MOUSEACTIVATE = &H21
    WM_CHILDACTIVATE = &H22
    WM_QUEUESYNC = &H23
    
    WM_GETMINMAXINFO = &H24
    
    WM_PAINTICON = &H26
    WM_ICONERASEBKGND = &H27
    WM_NEXTDLGCTL = &H28
    WM_SPOOLERSTATUS = &H2A
    WM_DRAWITEM = &H2B
    WM_MEASUREITEM = &H2C
    WM_DELETEITEM = &H2D
    WM_VKEYTOITEM = &H2E
    WM_CHARTOITEM = &H2F
    WM_SETFONT = &H30
    WM_GETFONT = &H31
    WM_SETHOTKEY = &H32
    WM_GETHOTKEY = &H33
    WM_QUERYDRAGICON = &H37
    WM_COMPAREITEM = &H39
    ' #if(WINVER >= 0x0500)
    WM_GETOBJECT = &H3D
    ' #endif /* WINVER >= 0x0500 */
    WM_COMPACTING = &H41
    
    WM_WINDOWPOSCHANGING = &H46
    WM_WINDOWPOSCHANGED = &H47
    
    WM_POWER = &H48
    
    WM_COPYDATA = &H4A
    WM_CANCELJOURNAL = &H4B
    
    ' #if(WINVER >= 0x0400)
    WM_NOTIFY = &H4E
    WM_INPUTLANGCHANGEREQUEST = &H50
    WM_INPUTLANGCHANGE = &H51
    WM_TCARD = &H52
    WM_HELP = &H53
    WM_USERCHANGED = &H54
    WM_NOTIFYFORMAT = &H55
    '
    '#define NFR_ANSI                             1
    '#define NFR_UNICODE                          2
    '#define NF_QUERY                             3
    '#define NF_REQUERY                           4
    '
    WM_CONTEXTMENU = &H7B
    WM_STYLECHANGING = &H7C
    WM_STYLECHANGED = &H7D
    WM_DISPLAYCHANGE = &H7E
    WM_GETICON = &H7F
    WM_SETICON = &H80
    ' #endif /* WINVER >= 0x0400 */
    '
    WM_NCCREATE = &H81
    WM_NCDESTROY = &H82
    WM_NCCALCSIZE = &H83
    WM_NCHITTEST = &H84
    WM_NCPAINT = &H85
    WM_NCACTIVATE = &H86
    WM_GETDLGCODE = &H87
    WM_NCMOUSEMOVE = &HA0
    WM_NCLBUTTONDOWN = &HA1
    WM_NCLBUTTONUP = &HA2
    WM_NCLBUTTONDBLCLK = &HA3
    WM_NCRBUTTONDOWN = &HA4
    WM_NCRBUTTONUP = &HA5
    WM_NCRBUTTONDBLCLK = &HA6
    WM_NCMBUTTONDOWN = &HA7
    WM_NCMBUTTONUP = &HA8
    WM_NCMBUTTONDBLCLK = &HA9
    '
    'WM_KEYFIRST = &H100
    WM_KEYDOWN = &H100
    WM_KEYUP = &H101
    WM_CHAR = &H102
    WM_DEADCHAR = &H103
    WM_SYSKEYDOWN = &H104
    WM_SYSKEYUP = &H105
    WM_SYSCHAR = &H106
    WM_SYSDEADCHAR = &H107
    'WM_KEYLAST = &H108
    '
    ' #if(WINVER >= 0x0400)
    WM_IME_STARTCOMPOSITION = &H10D
    WM_IME_ENDCOMPOSITION = &H10E
    WM_IME_COMPOSITION = &H10F
    'WM_IME_KEYLAST = &H10F
    ' #endif /* WINVER >= 0x0400 */
    '
    WM_INITDIALOG = &H110
    WM_COMMAND = &H111
    WM_SYSCOMMAND = &H112
    WM_TIMER = &H113
    WM_HSCROLL = &H114
    WM_VSCROLL = &H115
    WM_INITMENU = &H116
    WM_INITMENUPOPUP = &H117
    WM_MENUSELECT = &H11F
    WM_MENUCHAR = &H120
    WM_ENTERIDLE = &H121
    '
    ' #if(WINVER >= 0x0500)
    WM_MENURBUTTONUP = &H122
    WM_MENUDRAG = &H123
    WM_MENUGETOBJECT = &H124
    WM_UNINITMENUPOPUP = &H125
    WM_MENUCOMMAND = &H126
    ' #endif /* WINVER >= 0x0500 */
    '
    WM_CTLCOLORMSGBOX = &H132
    WM_CTLCOLOREDIT = &H133
    WM_CTLCOLORLISTBOX = &H134
    WM_CTLCOLORBTN = &H135
    WM_CTLCOLORDLG = &H136
    WM_CTLCOLORSCROLLBAR = &H137
    WM_CTLCOLORSTATIC = &H138
    
    'WM_MOUSEFIRST = &H200
    WM_MOUSEMOVE = &H200
    WM_LBUTTONDOWN = &H201
    WM_LBUTTONUP = &H202
    WM_LBUTTONDBLCLK = &H203
    WM_RBUTTONDOWN = &H204
    WM_RBUTTONUP = &H205
    WM_RBUTTONDBLCLK = &H206
    WM_MBUTTONDOWN = &H207
    WM_MBUTTONUP = &H208
    WM_MBUTTONDBLCLK = &H209
    ' #if (_WIN32_WINNT >= 0x0400) || (_WIN32_WINDOWS > 0x0400)
    WM_MOUSEWHEEL = &H20A
    'WM_MOUSELAST = &H20A
    ' #else
    'WM_MOUSELAST = &H209
    ' #endif /* if (_WIN32_WINNT < 0x0400) */
    
    WM_PARENTNOTIFY = &H210
    WM_ENTERMENULOOP = &H211
    WM_EXITMENULOOP = &H212
    
    ' #if(WINVER >= 0x0400)
    WM_NEXTMENU = &H213
    WM_SIZING = &H214
    WM_CAPTURECHANGED = &H215
    WM_MOVING = &H216
    
    WM_POWERBROADCAST = &H218
    WM_DEVICECHANGE = &H219
    ' #endif /* WINVER >= 0x0400 */
    
    WM_MDICREATE = &H220
    WM_MDIDESTROY = &H221
    WM_MDIACTIVATE = &H222
    WM_MDIRESTORE = &H223
    WM_MDINEXT = &H224
    WM_MDIMAXIMIZE = &H225
    WM_MDITILE = &H226
    WM_MDICASCADE = &H227
    WM_MDIICONARRANGE = &H228
    WM_MDIGETACTIVE = &H229
    WM_MDISETMENU = &H230
    WM_DROPFILES = &H233
    WM_MDIREFRESHMENU = &H234
    
    ' #if(WINVER >= 0x0400)
    WM_IME_SETCONTEXT = &H281
    WM_IME_NOTIFY = &H282
    WM_IME_CONTROL = &H283
    WM_IME_COMPOSITIONFULL = &H284
    WM_IME_SELECT = &H285
    WM_IME_CHAR = &H286
    ' #endif /* WINVER >= 0x0400 */
    ' #if(WINVER >= 0x0500)
    WM_IME_REQUEST = &H288
    ' #endif /* WINVER >= 0x0500 */
    ' #if(WINVER >= 0x0400)
    WM_IME_KEYDOWN = &H290
    WM_IME_KEYUP = &H291
    ' #endif /* WINVER >= 0x0400 */
    '
    ' #if(_WIN32_WINNT >= 0x0400)
    WM_MOUSEHOVER = &H2A1
    WM_MOUSELEAVE = &H2A3
    ' #endif /* _WIN32_WINNT >= 0x0400 */
    
    WM_CUT = &H300
    WM_COPY = &H301
    WM_PASTE = &H302
    WM_CLEAR = &H303
    WM_UNDO = &H304
    WM_RENDERFORMAT = &H305
    WM_RENDERALLFORMATS = &H306
    WM_DESTROYCLIPBOARD = &H307
    WM_DRAWCLIPBOARD = &H308
    WM_PAINTCLIPBOARD = &H309
    WM_VSCROLLCLIPBOARD = &H30A
    WM_SIZECLIPBOARD = &H30B
    WM_ASKCBFORMATNAME = &H30C
    WM_CHANGECBCHAIN = &H30D
    WM_HSCROLLCLIPBOARD = &H30E
    WM_QUERYNEWPALETTE = &H30F
    WM_PALETTEISCHANGING = &H310
    WM_PALETTECHANGED = &H311
    WM_HOTKEY = &H312
    '
    ' #if(WINVER >= 0x0400)
    WM_PRINT = &H317
    WM_PRINTCLIENT = &H318
    '
    WM_HANDHELDFIRST = &H358
    WM_HANDHELDLAST = &H35F
    '
    WM_AFXFIRST = &H360
    WM_AFXLAST = &H37F
    ' #endif /* WINVER >= 0x0400 */
    '
    WM_PENWINFIRST = &H380
    WM_PENWINLAST = &H38F
    '
    ' #if(WINVER >= 0x0400)
    WM_APP = &H8000
    ' #endif /* WINVER >= 0x0400 */
    
    ' NOTE: All Message Numbers below 0x0400 are RESERVED.
    
    ' Private Window Messages Start Here:
    WM_USER = &H400
End Enum

Public Enum eWSCHitTest
    ' WM_NCHITTEST and MOUSEHOOKSTRUCT Mouse Position Codes
    HTERROR = (-2)
    HTTRANSPARENT = (-1)
    HTNOWHERE = 0
    HTCLIENT = 1
    HTCAPTION = 2
    HTSYSMENU = 3
    HTGROWBOX = 4
    HTSIZE = HTGROWBOX
    HTMENU = 5
    HTHSCROLL = 6
    HTVSCROLL = 7
    HTMINBUTTON = 8
    HTMAXBUTTON = 9
    HTLEFT = 10
    HTRIGHT = 11
    HTTOP = 12
    HTTOPLEFT = 13
    HTTOPRIGHT = 14
    HTBOTTOM = 15
    HTBOTTOMLEFT = 16
    HTBOTTOMRIGHT = 17
    HTBORDER = 18
    HTREDUCE = HTMINBUTTON
    HTZOOM = HTMAXBUTTON
    HTSIZEFIRST = HTLEFT
    HTSIZELAST = HTBOTTOMRIGHT
End Enum

Public Enum eWSCMF
    ' Menú Flags para WM_MENUSELECT
    'MF_UNCHECKED = &H0&
    MF_GRAYED = &H1&
    MF_DISABLED = &H2&
    MF_BITMAP = &H4&
    MF_CHECKED = &H8&
    MF_POPUP = &H10&
    MF_HILITE = &H80&
    MF_OWNERDRAW = &H100&
    MF_SYSMENU = &H2000&
    MF_MOUSESELECT = &H8000&
End Enum

' Valores de fuSource para el mensaje WM_ENTERIDLE
Public Enum eWSCMSFG
    MSGF_DIALOGBOX = 0
    MSGF_MENU = 2
End Enum

' Mensajes varios
Public Enum eWSCMisc
    ' WM_ACTIVATE state values
    WA_INACTIVE = 0
    WA_ACTIVE = 1
    WA_CLICKACTIVE = 2
    
    ' wParam for WM_POWER window message and DRV_POWER driver notification
    PWR_OK = 1
    PWR_FAIL = (-1)
    PWR_SUSPENDREQUEST = 1
    PWR_SUSPENDRESUME = 2
    PWR_CRITICALRESUME = 3
    
    ' WM_SYNCTASK Commands
    ST_BEGINSWP = 0
    ST_ENDSWP = 1
    
    ' SendMessageTimeout values
    SMTO_NORMAL = &H0
    SMTO_BLOCK = &H1
    SMTO_ABORTIFHUNG = &H2
    
    ' WM_MOUSEACTIVATE Return Codes
    MA_ACTIVATE = 1
    MA_ACTIVATEANDEAT = 2
    MA_NOACTIVATE = 3
    MA_NOACTIVATEANDEAT = 4
    
    ' WM_SIZE message wParam values
    SIZE_RESTORED = 0
    SIZE_MINIMIZED = 1
    SIZE_MAXIMIZED = 2
    SIZE_MAXSHOW = 3
    SIZE_MAXHIDE = 4
    
    ' WM_NCCALCSIZE return flags
    WVR_ALIGNTOP = &H10
    WVR_ALIGNLEFT = &H20
    WVR_ALIGNBOTTOM = &H40
    WVR_ALIGNRIGHT = &H80
    WVR_HREDRAW = &H100
    WVR_VREDRAW = &H200
    WVR_REDRAW = (WVR_HREDRAW Or WVR_VREDRAW)
    WVR_VALIDRECTS = &H400
    
    ' Key State Masks for Mouse Messages
    MK_LBUTTON = &H1
    MK_RBUTTON = &H2
    MK_SHIFT = &H4
    MK_CONTROL = &H8
    MK_MBUTTON = &H10
    
    ' Constantes para el menú del sistema
    SC_RESTORE = &HF120&
    SC_MOVE = &HF010&
    SC_SIZE = &HF000&
    SC_MINIMIZE = &HF020&
    SC_MAXIMIZE = &HF030&
    SC_CLOSE = &HF060&
    
    ' Track Mouse Event
    HOVER_DEFAULT = &HFFFFFFFF
    
    ' #if(_WIN32_WINNT >= 0x0400)
    WHEEL_DELTA = 120                   ' /* Value for rolling one detent */
    ' #endif /* _WIN32_WINNT >= 0x0400 */
    ' #if(_WIN32_WINNT >= 0x0400)
    '#define UINT_MAX      =&hffffffff    /* maximum unsigned int value */
    '#define WHEEL_PAGESCROLL                (UINT_MAX) /* Scroll one page */
    WHEEL_PAGESCROLL = &HFFFFFFFF
    ' #endif /* _WIN32_WINNT >= 0x0400 */
    
    '// begin_pbt
    PBT_APMQUERYSUSPEND = &H0           ' Request for permission to suspend.
    PBT_APMQUERYSTANDBY = &H1
    '
    PBT_APMQUERYSUSPENDFAILED = &H2     ' Suspension request denied.
    PBT_APMQUERYSTANDBYFAILED = &H3
    '
    PBT_APMSUSPEND = &H4                ' System is suspending operation.
    PBT_APMSTANDBY = &H5
    '
    PBT_APMRESUMECRITICAL = &H6         ' Operation resuming after critical suspension.
    PBT_APMRESUMESUSPEND = &H7          ' Operation resuming after suspension.
    PBT_APMRESUMESTANDBY = &H8
    '
    PBTF_APMRESUMEFROMFAILURE = &H1&
    '
    PBT_APMBATTERYLOW = &H9             ' Battery power is low.
    PBT_APMPOWERSTATUSCHANGE = &HA      ' Power status has changed.
    '
    PBT_APMOEMEVENT = &HB               ' OEM-defined event occurred.
    PBT_APMRESUMEAUTOMATIC = &H12       ' Operation resuming automatically after event.
    '// end_pbt
    
End Enum

' WM_WINDOWPOSCHANGING/CHANGED struct pointed to by lParam
Private Type WINDOWPOS
    hWnd As Long
    hWndInsertAfter As Long
    X As Long
    Y As Long
    cx As Long
    cy As Long
    Flags As Long
End Type

Public Enum eTME
    ' Track Mouse Event
    TME_HOVER = &H1
    TME_LEAVE = &H2
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum

Private Type tTRACKMOUSEEVENT
    cbSize As Long
    dwFlags As eTME 'Long
    hWndTrack As Long
    dwHoverTime As Long
End Type

Private Declare Function TrackMouseEvent Lib "user32" _
    (lpEventTrack As tTRACKMOUSEEVENT) As Long

'Private Declare Function DefWindowProc Lib "User32" Alias "DefWindowProcA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (Destination As Any, Source As Any, ByVal Length As Long)

'------------------------------------------------------------------------------
' *** Los eventos producidos por la clase ***

'------------------------------------------------------------------------------
' Eventos genéricos, no directamente relacionados con la ventana subclasificada

' Permite interceptar todos los mensajes producidos
Public Event msg(ByVal uMSG As eWSCWM, ByVal wParam As Long, ByVal lParam As Long)

' Indica el menú seleccionado
Public Event MenuSelected(ByVal mnuItem As Long, mnuFlags As eWSCMF, ByVal hMenu As Long)
' Ha cambiado el tamaño del objeto Screen y/o el número de colores
Public Event DisplayChange(ByVal BitsPerPixel As Long, ByVal cxScreen As Long, ByVal cyScreen As Long)
' Se ha añadido o quitado alguna fuente del sistema
Public Event FontChange()
' El sistema indica que hay pocos recursos disponibles
Public Event LowMemory()
' Este evento se producirá mientras se está seleccionando un menú o en un diálogo
' dentro del formulario
'Public Event EnterIDLE(ByVal IsIDLE As Boolean, ByVal fuSource As Long, ByVal hWnd As Long)
'Public Event EnterIDLE(ByVal IsIDLE As Boolean, ByVal fuSource As eWSCMSFG, ByVal hWnd As Long)

'------------------------------------------------------------------------------
' Eventos directamente relacionados con la ventana subclasificada

' La ventana ha cambiado de posición, indica la posición y tamaño
Public Event WindowPosChanged(ByVal wLeft As Long, ByVal wTop As Long, ByVal wWidth As Long, ByVal wHeight As Long)
' Se ha movido esta ventana, devuelve la posición en twips
Public Event Move(ByVal wLeft As Long, ByVal wTop As Long)
' El ratón sale de la ventana subclasificada
Public Event MouseLeave()
' Indica de que control sale, si la ventana contiene otros controles
Public Event MouseLeaveOn(unControl As Object)
' El ratón entra en la ventana
Public Event MouseEnter()
' Indica en que control entra, si la ventana contiene controles
Public Event MouseEnterOn(unControl As Object)
' Indica si se ha movido la rueda del ratón y los parámetros correspondientes
Public Event MouseWheel(ByVal wKeys As Long, ByVal zDelta As Long, ByVal xPos As Long, ByVal yPos As Long)
'
Public Event SetCursor(unControl As Object, ByVal HitTest As eWSCHitTest, ByVal MouseMsg As Long)
' Se activa la ventana
Public Event Activate()
' Se desactiva la ventana
Public Event Deactivate()

Private Sub Class_Initialize()
    ' Esto no es necesario, pero lo uso para saber que cuando se inicia la clase
    ' estas variables locales están a "cero"
    mhWnd = 0
    mPrevWndProc = 0
    
    '--------------------------------------------------------------------------
    ' Para devolver los valores en twips,
    ' ya que el valor devuelto originalmente será en pixels
    '
    ' Se hace en este evento para no tener que calcularlo en cada mensaje.
    '$Por hacer: (22/Mar/99)
    '   Si se cambia la resolución habría que cambiar estos valores...
    '--------------------------------------------------------------------------
    twX = Screen.TwipsPerPixelX
    twY = Screen.TwipsPerPixelY
End Sub

Private Sub Class_Terminate()
    ' Comprobar si están iniciadas las variables
    ' y de ser así dejar de subclasificar el formulario
    If mhWnd Then
        mMouse.unHook Me
    End If
End Sub

Public Sub Hook(Optional ByVal NewValue As Object = Nothing)
    '--------------------------------------------------------------------------
    ' Este método se debe usar para iniciar o terminar la subclasificación
    '
    ' El parámetro se usará para indicar la ventana a subclasificar,
    ' se podría haber usado un valor Long para el hWnd de la ventana, pero,
    ' es que este objeto se usa para comprobar los controles incluidos
    ' ya que en los mensajes procesados se puede procesar si el ratón sale o
    ' entra en algún control incluido en la colección controls,
    ' por tanto sólo tiene utilidad real si ese objeto tiene una colección Controls
    ' (por ejemplo un formulario o un control de usuario)
    '
    ' Nota: Se puede subclasificar un form o cualquier otro control con hWnd
    '--------------------------------------------------------------------------
    
    On Local Error GoTo ErrHook
    
    If NewValue Is Nothing Then
        ' Quitar el 'gancho' del form
        If Me.hWnd <> 0 Then
            mMouse.unHook Me
        End If
    
    ' Si no tiene hWnd producirá error
    ElseIf NewValue.hWnd = 0 Then
        '
    Else
        ' Si ya había una copia funcionando, quitar el gancho       (08/Ago/99)
        If Me.hWnd <> 0 Then
            mMouse.unHook Me
        End If
        ' Asignar el nuevo gancho a la ventana a subclasificar
        mMouse.Hook Me, NewValue
        
        Set mWndType = NewValue
        
    End If
    
    Exit Sub
ErrHook:
    ' Nada que hacer si se pasa un parámero erróneo.
    ' Devolver un error indicándolo
    On Local Error GoTo 0
    With err
        .Description = "El parámetro debe ser una ventana (que tenga la propiedad hWnd) " & _
                       "o Nothing para dejar de interceptar los mensajes." & _
                       vbCrLf & vbCrLf & _
                       "El tipo del parámetro es: " & TypeName(NewValue)
                       
        .Source = "cSubclass.Hook"
        .Raise .Number
    End With
End Sub

Friend Sub unMSG(ByVal uMSG As eWSCWM, ByVal wParam As Long, ByVal lParam As Long)
        ' Este procedimiento será el que reciba los mensajes y se encargará
        ' de producir los eventos correspondientes
        '<EhHeader>
        On Error GoTo unMSG_Err
        '</EhHeader>
    
        ' Usar variables estáticas para mayor rapidez
        Static tWP As WINDOWPOS
        Static bMouseEnter As Boolean
        Static ControlAnt As Control        ' Control en el que se produjo el MouseMove
                                            ' para los casos en los que la ventana
                                            ' subclasificada tenga otros controles
    
        ' Si se produce un error que no se pare la aplicación...
        On Error Resume Next
    
    '------------------------------------------------------------------------------
    ' Si está en modo IDLE (no hay mensajes)
    'If uMSG = WM_ENTERIDLE Then
    '    '--------------------------------------------------------------------------
    '    ' Se ha entrado en un estado en el que no hay mensajes
    '
    '    '
    '    ' fuSource = wParam;    // idle-source flag
    '    ' hwnd = (HWND) lParam; // handle of dialog box or owner window
    '    '
    '    ' Value of wParam.  Specifies whether the message is the result of a dialog box
    '    '                   or a menu being displayed.
    '    ' This parameter can be one of the following values:
    '    '   MSGF_DIALOGBOX    The system is idle because a dialog box is displayed.
    '    '   MSGF_MENU         The system is idle because a menu is displayed.
    '    '
    '    RaiseEvent EnterIDLE(True, wParam, lParam)
    '    '
    ''------------------------------------------------------------------------------
    '' Hay otros mensajes aparte del IDLE
    'Else
    '    '--------------------------------------------------------------------------
    '    ' Indicar que ya no está en modo IDLE
    '    RaiseEvent EnterIDLE(False, 0&, 0&)
        
        '--------------------------------------------------------------------------
        ' Evento genérico
100     RaiseEvent msg(uMSG, wParam, lParam)
    
    
        ' Según el tipo de mensaje, producir eventos...
102     Select Case uMSG
    
        '--------------------------------------------------------------------------
        ' Selección de un menú
        Case WM_MENUSELECT
            ' Los parámetros son:   mnuItem         mnuFlags        mnuHandle
104         RaiseEvent MenuSelected(LoWord(wParam), HiWord(wParam), lParam)
    
        '--------------------------------------------------------------------------
        ' Cambio en la posición y tamaño de la ventana
        ' La posición y tamaño se indicará en twips
106     Case WM_WINDOWPOSCHANGED 'WM_WINDOWPOSCHANGING
            '----------------------------------------------------------------------
            ' Nota: Los valores de twX y twY se asignan en el evento Initialize
        
            '--------------------------------------------------------------------------
            ' Copiar el valor referenciado por lParam en un UDT
108         CopyMemory tWP, ByVal lParam, Len(tWP)
110         With tWP
                '                           Left      Top       Width      Height
112             RaiseEvent WindowPosChanged(.X * twX, .Y * twY, .cx * twX, .cy * twY)
            End With
    
        '--------------------------------------------------------------------------
        ' Se ha movido la ventana
        ' La posición se indicará en twips
114     Case WM_MOVE
116         With tWP
118             .X = LoWord(lParam)
120             .Y = HiWord(lParam)
122             RaiseEvent Move(.X * twX, .Y * twY)
            End With
    
        '--------------------------------------------------------------------------
        ' Ha cambiado la resolución de la pantalla o el número de colores
124     Case WM_DISPLAYCHANGE
            '                         BitsPerPixels
            '                                 New horizontal resolution of the screen
            '                                                 New vertical resolution of the screen
126         RaiseEvent DisplayChange(wParam, LoWord(lParam), HiWord(lParam))
    
        '--------------------------------------------------------------------------
        ' Se han añadido o quitado fuentes del sistema
128     Case WM_FONTCHANGE
130         RaiseEvent FontChange
    
        '--------------------------------------------------------------------------
        ' El sistema está compactando la memoria porque tiene poca disponible
132     Case WM_COMPACTING
134         RaiseEvent LowMemory
    
        '--------------------------------------------------------------------------
        ' The WM_ACTIVATEAPP message is sent when a window belonging to a different
        ' application than the active window is about to be activated.
        ' The message is sent to the application whose window is being activated
        ' and to the application whose window is being deactivated.
        ' fActive = (BOOL) wParam;        // activation flag
        '
        ' Specifies whether the window is being activated or deactivated.
        ' This parameter is TRUE if the window is being activated;
        ' it is FALSE if the window is being deactivated.
        '
        ' Si se activa o desactiva la ventana
136     Case WM_ACTIVATEAPP
138         If wParam Then
140             RaiseEvent Activate
            Else
142             RaiseEvent Deactivate
            End If
        
        '--------------------------------------------------------------------------
        ' Si se ha movido la rueda de los ratos con ruedas
144     Case WM_MOUSEWHEEL
            'fwKeys = LOWORD(wParam);    // key flags
            'zDelta = (short) HIWORD(wParam);    // wheel rotation
            'xPos = (short) LOWORD(lParam);    // horizontal position of pointer
            'yPos = (short) HIWORD(lParam);    // vertical position of pointer
146         RaiseEvent MouseWheel(LoWord(wParam), HiWord(wParam), LoWord(lParam), HiWord(lParam))
            '
            ' fwKeys:
            ' Indicates whether various virtual keys are down.
            ' This parameter can be any combination of the following values:
            ' (estos valores están en la enumeración: eWSCMisc)
            ' MK_CONTROL    Set if the ctrl key is down.
            ' MK_LBUTTON    Set if the left mouse button is down.
            ' MK_MBUTTON    Set if the middle mouse button is down.
            ' MK_RBUTTON    Set if the right mouse button is down.
            ' MK_SHIFT      Set if the shift key is down.
            '
            ' zDelta:
            ' Indicates the distance that the wheel is rotated,
            ' expressed in multiples or divisions of WHEEL_DELTA, which is 120.
            ' A positive value indicates that the wheel was rotated forward,
            ' away from the user;
            ' a negative value indicates that the wheel was rotated backward,
            ' toward the user.
    
        '--------------------------------------------------------------------------
        ' Cuando el ratón sale del área de la ventana indicada
148     Case WM_MOUSELEAVE
            ' Si este mensaje se intercepta en un control incluido en el form
            ' además del propio form, cuando se entre en un control se producirá
            ' un MouseLeave en el form (y viceversa)
150         RaiseEvent MouseLeave
152         bMouseEnter = False
    
        '--------------------------------------------------------------------------
        ' Cuando el ratón se mueve, es porque está en la ventana,
        ' comprobar si es la primera vez que entra y así lanzar el evento MouseEnter
154     Case WM_MOUSEMOVE
            ' Cuando se produce un MouseMove, es decir que el mouse está en el control
            ' hay que indicar que se vuelva a detectar la salida del control
156         If Not bMouseEnter Then
158             bMouseEnter = True
160             With mtTME
162                 .dwFlags = TME_LEAVE
164                 .cbSize = Len(mtTME)
166                 .hWndTrack = mhWnd
                End With
168             Call TrackMouseEvent(mtTME)
                '
170             RaiseEvent MouseEnter
                '
                ' Comprobar si está asignado el control anterior
172             If Not (ControlAnt Is Nothing) Then
174                 RaiseEvent MouseLeaveOn(ControlAnt)
                End If
            End If
        
        '--------------------------------------------------------------------------
        ' The WM_SETCURSOR message is sent to a window if the mouse causes the cursor
        ' to move within a window and mouse input is not captured.
        ' Parámetros:
        '   hwnd = (HWND) wParam;       // handle to window with cursor
        '   nHittest = LOWORD(lParam);  // hit-test code
        '   wMouseMsg = HIWORD(lParam); // mouse-message identifier
176     Case WM_SETCURSOR
            Dim I As Long
            Dim tControl As Control
            Dim nHittest As Long
            Dim wMouseMsg As Long
        
178         nHittest = LoWord(lParam)
180         wMouseMsg = HiWord(lParam)
        
            ' Este bucle sólo funcionará si la ventana subclasificada
            ' tiene una colección de controles, (por ejemplo un form)
182         For Each tControl In mWndType.Controls
184             With tControl
                    ' Si no es este control o el control no soporta el hWnd
186                 If .hWnd <> wParam Then
                        '
                    Else
188                     RaiseEvent SetCursor(tControl, nHittest, wMouseMsg)
                        ' Si el mensaje del ratón es que se mueve...
190                     If wMouseMsg = WM_MOUSEMOVE Then
                            ' Comprobar si ya había entrado en el control
192                         If ControlAnt.hWnd <> wParam Then
                                ' Si ya se había asignado
194                             If Not (ControlAnt Is Nothing) Then
                                    ' Se ha salido de este control
196                                 RaiseEvent MouseLeaveOn(ControlAnt)
                                End If
                            End If
                            ' Asignar este control para posteriores comprobaciones
198                         Set ControlAnt = tControl
200                         RaiseEvent MouseEnterOn(tControl)
                        End If
                        Exit For
                    End If
                End With
            Next
    
        End Select
    
    '    '
    'End If

        '<EhFooter>
        Exit Sub

unMSG_Err:
        err.Raise vbObjectError + 100, _
                  "ARGENTUM.clsMouse.unMSG", _
                  "clsMouse component failure"
                  
                   On Error Resume Next
        '</EhFooter>
End Sub

Friend Property Get hWnd() As Long
    hWnd = mhWnd
End Property

Friend Property Let hWnd(ByVal NewValue As Long)
    mhWnd = NewValue
End Property

Friend Property Get PrevWndProc() As Long
    PrevWndProc = mPrevWndProc
End Property

Friend Property Let PrevWndProc(ByVal NewValue As Long)
    mPrevWndProc = NewValue
End Property

Private Function LoWord(ByVal Numero As Long) As Long
    ' Devuelve el LoWord del número pasado como parámetro
    LoWord = Numero And &HFFFF&
End Function

Private Function HiWord(ByVal Numero As Long) As Long
    ' Devuelve el HiWord del número pasado como parámetro
    HiWord = Numero \ &H10000 And &HFFFF&
End Function

