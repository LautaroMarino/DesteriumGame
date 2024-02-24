VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClanes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Clase ModuleMonitor

Option Explicit

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetModuleFileNameA Lib "kernel32" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

'Private WithEvents Timer1 As clsTimer
Private CurrentModules As Collection

Public Event ModuleLoaded(ByVal moduleName As String)

Public Sub StartMonitoring()
    'Set Timer1 = New clsTimer
    'Timer1.Interval = intervalInSeconds * 1000 ' Convertir segundos a milisegundos
    Set CurrentModules = New Collection
   ' Timer1.Enabled = True
End Sub

Public Sub SearchMonitoring()
    Dim hModule As Long
    Dim moduleName As String
    Dim moduleNameLength As Long
    Dim shouldContinue As Boolean
    
    shouldContinue = True
    
    Do While shouldContinue
        hModule = GetModuleHandle(vbNullString)
        If hModule = 0 Then
            shouldContinue = False ' Salir del bucle cuando hModule sea 0
        Else
            moduleNameLength = GetModuleFileNameA(hModule, moduleName, Len(moduleName))
            moduleName = Left(moduleName, moduleNameLength)
            
            ' Comprobar si el módulo está en la colección actual
            If Not ModuleExists(moduleName) Then
                ' Este módulo es nuevo y se ha cargado en tiempo de ejecución
                'CurrentModules.Add moduleName
                'RaiseEvent ModuleLoaded(moduleName)
                Call WriteDenounce("[SEGURIDAD]: Posible inyección de librería.")
            End If
        End If
    Loop
End Sub


Private Function ModuleExists(moduleName As String) As Boolean
    Dim module As Variant
    ModuleExists = False
    For Each module In CurrentModules
        If module = moduleName Then
            ModuleExists = True
            Exit Function
        End If
    Next module
End Function
