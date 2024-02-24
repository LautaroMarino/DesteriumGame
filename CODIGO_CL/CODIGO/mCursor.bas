Attribute VB_Name = "mCursor"

Option Explicit

Public Declare Function LoadCursorFromFile Lib "user32" Alias _
    "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Public Declare Function SetSystemCursor Lib "user32" _
    (ByVal hcur As Long, ByVal ID As Long) As Long
Public Declare Function GetCursor Lib "user32" () As Long
Public Declare Function CopyIcon Lib "user32" (ByVal hcur As Long) As Long

Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long

Public Const IDC_ARROW = 32512&
Public Const IDC_IBEAM = 32513&
Public Const IDC_WAIT = 32514&
Public Const IDC_CROSS = 32515&
Public Const IDC_UPARROW = 32516&
Public Const IDC_ICON = 32641&
Public Const IDC_SIZENWSE = 32642&
Public Const IDC_SIZENESW = 32643&
Public Const IDC_SIZEWE = 32644&
Public Const IDC_SIZENS = 32645&
Public Const IDC_SIZEALL = 32646&
Public Const IDC_NO = 32648&
Public Const IDC_HAND = 32649&
Public Const IDC_APPSTARTING = 32650&

Public lngOldCursor As Long, lngNewCursor As Long
Public lngDefaultCursor As Long

Public Const SPIF_UPDATEINIFILE = &H1
Public Const SPIF_SENDWININICHANGE = &H2
Public Const SPI_SETCURSORS = &H57 'restores sys cursors

Public Declare Function SystemParametersInfo Lib "user32" _
Alias "SystemParametersInfoA" _
(ByVal uAction As Long, ByVal uParam As Long, _
lpvParam As Any, ByVal fuWinIni As Long) As Long

' Seteamos los cursores Default
Public Sub Cursores_ResotreDefault()


    SystemParametersInfo SPI_SETCURSORS, 0&, ByVal 0&, _
        (SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)
    
        
End Sub
Public Sub StartAnimatedCursor(AniFilePath As String, ID As Long)
    
    If ClientSetup.bConfig(eSetupMods.SETUP_CURSORES) = 1 Then Exit Sub
    
    'If lngNewCursor = 0 Then
    
        If InStr(1, AniFilePath, "") Then
            lngNewCursor = LoadCursorFromFile(AniFilePath)
        Else
            lngNewCursor = LoadCursorFromFile(App.path & _
               "" & AniFilePath)
        End If
    
        SetSystemCursor lngNewCursor, ID

   ' End If
End Sub

Public Sub RestoreLastCursor(ID As Long)

    ' Restore last cursor
    SetSystemCursor lngNewCursor, ID
    
    lngOldCursor = 0
    lngNewCursor = 0
End Sub
