Attribute VB_Name = "mClanes"
' Módulo de Seguridad

Option Explicit

Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetModuleFileNameA Lib "kernel32" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long


Private Const WH_SHELL As Long = 10&
Private Const HSHELL_LOADDLL As Long = 20&

Private hHook As Long

' Función del gancho de Windows para capturar la carga de DLL
Private Function ShellProc(ByVal nCode As Long, ByVal wParam As Long, lParam As Long) As Long
    If nCode = 0 And wParam = HSHELL_LOADDLL Then
        Dim ModuleName As String
        ModuleName = Space(256)
        GetModuleFileNameA lParam, ModuleName, Len(ModuleName)
        ModuleName = Left$(ModuleName, InStr(ModuleName, vbNullChar) - 1)
        
        TempModuleName = ModuleName
    End If
    
    ShellProc = CallNextHookEx(0, nCode, wParam, ByVal lParam)
End Function

' Inicia el monitoreo de la carga de DLL en tiempo de ejecución
Public Sub StartMonitoring()
    If hHook = 0 Then
        hHook = SetWindowsHookEx(WH_SHELL, AddressOf ShellProc, GetModuleHandleA(App.EXEName), GetCurrentThreadId)
        If hHook = 0 Then
            TempModuleName = "Error StartHook"
        End If
    End If
End Sub

' Detiene el monitoreo de la carga de DLL en tiempo de ejecución
Public Sub StopMonitoring()
    If hHook <> 0 Then
        UnhookWindowsHookEx hHook
        hHook = 0
    End If
End Sub
