Attribute VB_Name = "mClientSetup"
Option Explicit
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (destination As Any, ByVal Length As Long)
Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, pbData As Any, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, phHash As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal hHash As Long, ByVal dwParam As Long, pbData As Any, pdwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (phProv As Long, ByVal pszContainer As Long, ByVal pszProvider As Long, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long

Private Const PROV_RSA_FULL As Long = 1
Private Const CRYPT_VERIFYCONTEXT As Long = &HF0000000
Private Const CALG_MD5 As Long = 32771

Private hashInitialized As Boolean
Private savedHashValue As String

Private Function CalculateHash(ByVal data As String) As String

    Dim hCryptProv     As Long

    Dim hHash          As Long

    Dim result         As Long

    Dim hashData()     As Byte

    Dim hashDataLength As Long
    
    result = CryptAcquireContext(hCryptProv, 0, 0, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT)

    If result = 0 Then
        ' Manejar error al adquirir el contexto criptográfico
        ' ...
        Exit Function

    End If
    
    result = CryptCreateHash(hCryptProv, CALG_MD5, 0, 0, hHash)

    If result = 0 Then
        ' Manejar error al crear el hash
        ' ...
        CryptReleaseContext hCryptProv, 0
        Exit Function

    End If
    
    result = CryptHashData(hHash, ByVal StrPtr(data), Len(data), 0)

    If result = 0 Then
        ' Manejar error al calcular el hash
        ' ...
        CryptDestroyHash hHash
        CryptReleaseContext hCryptProv, 0
        Exit Function

    End If
    
    result = CryptGetHashParam(hHash, 21, 0, hashDataLength, 0)

    If result = 0 Then
        ' Manejar error al obtener la longitud del hash
        ' ...
        CryptDestroyHash hHash
        CryptReleaseContext hCryptProv, 0
        Exit Function

    End If
    
    ReDim hashData(hashDataLength - 1) As Byte
    
    result = CryptGetHashParam(hHash, 21, hashData(0), hashDataLength, 0)

    If result = 0 Then
        ' Manejar error al obtener el hash
        ' ...
        CryptDestroyHash hHash
        CryptReleaseContext hCryptProv, 0
        Exit Function

    End If

    CalculateHash = StrConv(hashData, vbUnicode)

    CryptDestroyHash hHash
    CryptReleaseContext hCryptProv, 0

End Function
Private Function GetMemoryHash() As String
Dim processId As Long
Dim memoryData As String
processId = GetCurrentProcessId()

' Leer los datos de memoria relevantes
' Esto puede incluir valores críticos, estructuras de datos o secciones específicas de la memoria
' En este ejemplo, solo se muestra un valor de ejemplo
memoryData = CStr(processId) & "example_data"

GetMemoryHash = CalculateHash(memoryData)
End Function
Private Sub CheckMemory()
Dim currentHashValue As String

currentHashValue = GetMemoryHash()

If hashInitialized Then
    If currentHashValue <> savedHashValue Then
        MsgBox "Modificación de memoria detectada"
    End If
Else
    savedHashValue = currentHashValue
    hashInitialized = True
End If
End Sub
Public Function PATH_CLIENTSETUP() As String
    
    ' ## En testeo cargo la PC DE MI CASA
    ' G:\2024\0_2024_V4\AO
    #If Testeo = 1 Then
        PATH_CLIENTSETUP = Replace$(App.path, "\AO", "") & "\config.ini"
    #Else
        PATH_CLIENTSETUP = Replace$(App.path, "Argentum Game\AO4", "Argentum Game\") & "config.ini"
        
        If Not FileExist(PATH_CLIENTSETUP, vbDirectory) Then
            PATH_CLIENTSETUP = Replace$(App.path, "Desterium Game\AO4", "Desterium Game\") & "config.ini"
        End If
    #End If
   
    Debug.Print PATH_CLIENTSETUP
End Function

Public Sub ILoadClientSetup()
        '<EhHeader>
        On Error GoTo ILoadClientSetup_Err
        '</EhHeader>
 
 
        Dim A As Long

100
    
        ' Start Cursor
102     Call StartAnimatedCursor(App.path & "\resource\cursor\" & ClientSetup.CursorGeneral, IDC_ARROW)
104     Call StartAnimatedCursor(App.path & "\resource\cursor\" & ClientSetup.CursorSpell, IDC_CROSS)
106     Call StartAnimatedCursor(App.path & "\resource\cursor\" & ClientSetup.CursorHand, IDC_HAND)
    
108         If FileExist(PATH_CLIENTSETUP, vbArchive) Then

110             ClientSetup.CursorGeneral = GetVar(PATH_CLIENTSETUP, "CURSOR", "GENERAL")
112             ClientSetup.CursorHand = GetVar(PATH_CLIENTSETUP, "CURSOR", "HAND")
114             ClientSetup.CursorInv = GetVar(PATH_CLIENTSETUP, "CURSOR", "INV")
116             ClientSetup.CursorSpell = GetVar(PATH_CLIENTSETUP, "CURSOR", "SPELL")

118             ClientSetup.bMasterSound = Val(GetVar(PATH_CLIENTSETUP, "SOUND", "MASTER"))
120             ClientSetup.bSoundMusic = Val(GetVar(PATH_CLIENTSETUP, "SOUND", "MUSIC"))
122             ClientSetup.bSoundEffect = Val(GetVar(PATH_CLIENTSETUP, "SOUND", "EFFECT"))
124             ClientSetup.bSoundInterface = Val(GetVar(PATH_CLIENTSETUP, "SOUND", "INTERFACE"))

126             ClientSetup.bValueSoundMusic = Val(GetVar(PATH_CLIENTSETUP, "SOUND", "VALUEMUSIC"))
128             ClientSetup.bValueSoundEffect = Val(GetVar(PATH_CLIENTSETUP, "SOUND", "VALUEEFFECT"))
130             ClientSetup.bValueSoundInterface = Val(GetVar(PATH_CLIENTSETUP, "SOUND", "VALUEINTERFACE"))
132             ClientSetup.bValueSoundMaster = Val(GetVar(PATH_CLIENTSETUP, "SOUND", "VALUEMASTER"))
                  
                  ClientSetup.bResolution = Val(GetVar(PATH_CLIENTSETUP, "VIDEO", "RESOLUTION"))
                  
134             ClientSetup.bFps = Val(GetVar(PATH_CLIENTSETUP, "VIDEO", "FPS"))
136             ClientSetup.bAlpha = Val(GetVar(PATH_CLIENTSETUP, "VIDEO", "ALPHA"))
                  ClientSetup.bResolution = Val(GetVar(PATH_CLIENTSETUP, "VIDEO", "RESOLUTION"))
                   
138             For A = 1 To MAX_SETUP_MODS
140                 ClientSetup.bConfig(A) = Val(GetVar(PATH_CLIENTSETUP, "CONFIG", CStr(A)))
                Next
                
            Else
142             Call MsgBox("Hubo un error crítico al cargar las opciones de juego. Contacta a los Administradores del Juego", vbCritical, App.Title)
                'End
            End If
        '<EhFooter>
        Exit Sub

ILoadClientSetup_Err:
        LogError err.Description & vbCrLf & _
               "in ARGENTUM.mClientSetup.ILoadClientSetup " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
    End Sub
    
Public Sub ISaveClientSetup()
        '<EhHeader>
        On Error GoTo ISaveClientSetup_Err
        '</EhHeader>
     
        Dim A As Long
     
100     Call WriteVar(PATH_CLIENTSETUP, "CURSOR", "GENERAL", ClientSetup.CursorGeneral)
102     Call WriteVar(PATH_CLIENTSETUP, "CURSOR", "HAND", ClientSetup.CursorHand)
104     Call WriteVar(PATH_CLIENTSETUP, "CURSOR", "INV", ClientSetup.CursorInv)
106     Call WriteVar(PATH_CLIENTSETUP, "CURSOR", "SPELL", ClientSetup.CursorSpell)

108     Call WriteVar(PATH_CLIENTSETUP, "SOUND", "MASTER", CStr(ClientSetup.bMasterSound))
110     Call WriteVar(PATH_CLIENTSETUP, "SOUND", "MUSIC", CStr(ClientSetup.bSoundMusic))
112     Call WriteVar(PATH_CLIENTSETUP, "SOUND", "EFFECT", CStr(ClientSetup.bSoundEffect))
114     Call WriteVar(PATH_CLIENTSETUP, "SOUND", "INTERFACE", CStr(ClientSetup.bSoundInterface))

116     Call WriteVar(PATH_CLIENTSETUP, "SOUND", "VALUEMUSIC", CStr(ClientSetup.bValueSoundMusic))
118     Call WriteVar(PATH_CLIENTSETUP, "SOUND", "VALUEEFFECT", CStr(ClientSetup.bValueSoundEffect))
120     Call WriteVar(PATH_CLIENTSETUP, "SOUND", "VALUEINTERFACE", CStr(ClientSetup.bValueSoundInterface))
          Call WriteVar(PATH_CLIENTSETUP, "SOUND", "VALUEMASTER", CStr(ClientSetup.bValueSoundMaster))
          
122     Call WriteVar(PATH_CLIENTSETUP, "VIDEO", "FPS", CStr(ClientSetup.bFps))
124     Call WriteVar(PATH_CLIENTSETUP, "VIDEO", "ALPHA", CStr(ClientSetup.bAlpha))
          Call WriteVar(PATH_CLIENTSETUP, "VIDEO", "RESOLUTION", CStr(ClientSetup.bResolution))
          
126     For A = 1 To MAX_SETUP_MODS
128         Call WriteVar(PATH_CLIENTSETUP, "CONFIG", CStr(A), CStr(ClientSetup.bConfig(A)))
        Next

        '<EhFooter>
        Exit Sub

ISaveClientSetup_Err:
        LogError err.Description & vbCrLf & _
               "in ARGENTUM.mClientSetup.ISaveClientSetup " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

