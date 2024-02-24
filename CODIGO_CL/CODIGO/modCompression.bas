Attribute VB_Name = "modCompression"
Option Explicit

Public Const PNG_SOURCE_FILE_EXT As String = ".png"

Public Const BMP_SOURCE_FILE_EXT As String = ".bmp"


#If ModoBig = 1 Then

    Public Const GRH_RESOURCE_FILE As String = "graphics2.ao"

#Else

    Public Const GRH_RESOURCE_FILE As String = "graphics.ao"

#End If

Public Const GRH_RESOURCE_FILE_DEFAULT As String = "graphics.ao"


Public Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" ( _
    ByVal lpPictDesc As picDesc, _
    ByRef riid As Any, _
    ByVal fOwn As Long, _
    ByRef ppvObj As IPicture) As Long

Public Type picDesc
    Size As Long
    Type As Long
    hPic As Long
    hPal As Long
End Type

Public Const AVATARS_RESOURCE_FILE As String = "avatars.ao"

Public Const GRH_PATCH_FILE        As String = "patch.ao"

Public Const MAPS_SOURCE_FILE_EXT  As String = ".map"

Public Const MAPS_RESOURCE_FILE    As String = "Mapas.AO"

Public Const MAPS_PATCH_FILE       As String = "Mapas.PATCH"

Public GrhDatContra()              As Byte ' Contraseña

Public GrhUsaContra                As Boolean ' Usa Contraseña?

Public MapsDatContra()             As Byte ' Contraseña

Public MapsUsaContra               As Boolean  ' Usa Contraseña?

'This structure will describe our binary file's
'size, number and version of contained files
Public Type FILEHEADER

    lngNumFiles As Long                 'How many files are inside?
    lngFileSize As Long                 'How big is this file? (Used to check integrity)
    lngFileVersion As Long              'The resource version (Used to patch)

End Type

'This structure will describe each file contained
'in our binary file
Public Type INFOHEADER

    lngFileSize As Long             'How big is this chunk of stored data?
    lngFileStart As Long            'Where does the chunk start?
    strFileName As String * 16      'What's the name of the file this data came from?
    lngFileSizeUncompressed As Long 'How big is the file compressed

End Type

Private Enum PatchInstruction

    Delete_File
    Create_File
    Modify_File

End Enum

Private Declare Function compress _
                Lib "zlib.dll" (dest As Any, _
                                destlen As Any, _
                                src As Any, _
                                ByVal srclen As Long) As Long

Private Declare Function uncompress _
                Lib "zlib.dll" (dest As Any, _
                                destlen As Any, _
                                src As Any, _
                                ByVal srclen As Long) As Long

Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (ByRef dest As Any, _
                                       ByRef Source As Any, _
                                       ByVal byteCount As Long)

'BitMaps Strucures
Public Type BITMAPFILEHEADER

    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long

End Type

Public Type BITMAPINFOHEADER

    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long

End Type

Public Type RGBQUAD

    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte

End Type

Public Type BITMAPINFO

    bmiHeader As BITMAPINFOHEADER
    bmiColors(255) As RGBQUAD

End Type

Private Const BI_RGB       As Long = 0

Private Const BI_RLE8      As Long = 1

Private Const BI_RLE4      As Long = 2

Private Const BI_BITFIELDS As Long = 3

Private Const BI_JPG       As Long = 4

Private Const BI_PNG       As Long = 5

Private Declare Function CreateStreamOnHGlobal _
                Lib "ole32" (ByVal hGlobal As Long, _
                             ByVal fDeleteOnRelease As Long, _
                             ppstm As Any) As Long

Private Declare Function GlobalAlloc _
                Lib "kernel32" (ByVal uFlags As Long, _
                                ByVal dwBytes As Long) As Long

Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare Function OleLoadPicture _
                Lib "olepro32" (pStream As Any, _
                                ByVal lSize As Long, _
                                ByVal fRunmode As Long, _
                                riid As Any, _
                                ppvObj As Any) As Long
                                
'To get free bytes in drive
Private Declare Function GetDiskFreeSpace _
                Lib "kernel32" _
                Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, _
                                             FreeBytesToCaller As Currency, _
                                             BytesTotal As Currency, _
                                             FreeBytesTotal As Currency) As Long

Public Sub GenerateContra(ByVal Contra As String, Optional Modo As Byte = 0)
    '***************************************************
    'Author: ^[GS]^
    'Last Modification: 17/06/2012 - ^[GS]^
    '
    '***************************************************

    On Error Resume Next

    Dim LoopC As Byte

    If Modo = 0 Then
        Erase GrhDatContra
    ElseIf Modo = 1 Then
        Erase MapsDatContra

    End If
    
    If LenB(Contra) <> 0 Then
        If Modo = 0 Then
            ReDim GrhDatContra(Len(Contra) - 1)

            For LoopC = 0 To UBound(GrhDatContra)
                GrhDatContra(LoopC) = Asc(mid(Contra, LoopC + 1, 1))
            Next LoopC

            GrhUsaContra = True
        ElseIf Modo = 1 Then
            ReDim MapsDatContra(Len(Contra) - 1)

            For LoopC = 0 To UBound(MapsDatContra)
                MapsDatContra(LoopC) = Asc(mid(Contra, LoopC + 1, 1))
            Next LoopC

            MapsUsaContra = True

        End If

    Else

        If Modo = 0 Then
            GrhUsaContra = False
        ElseIf Modo = 1 Then
            MapsUsaContra = False

        End If

    End If
    
End Sub

Private Function General_Drive_Get_Free_Bytes(ByVal DriveName As String) As Currency

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 6/07/2004
    '
    '**************************************************************
    Dim retval As Long

    Dim FB     As Currency

    Dim BT     As Currency

    Dim FBT    As Currency
    
    retval = GetDiskFreeSpace(Left$(DriveName, 2), FB, BT, FBT)
    
    General_Drive_Get_Free_Bytes = FB * 10000 'convert result to actual size in bytes

End Function

''
' Sorts the info headers by their file name. Uses QuickSort.
'
' @param    InfoHead() The array of headers to be ordered.
' @param    first The first index in the list.
' @param    last The last index in the list.

Private Sub Sort_Info_Headers(ByRef InfoHead() As INFOHEADER, _
                              ByVal First As Long, _
                              ByVal Last As Long)

    '*****************************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modify Date: 08/20/2007
    'Sorts the info headers by their file name using QuickSort.
    '*****************************************************************
    Dim aux  As INFOHEADER

    Dim min  As Long

    Dim max  As Long

    Dim comp As String
    
    min = First
    max = Last
    
    comp = InfoHead((min + max) \ 2).strFileName
    
    Do While min <= max
        Do While InfoHead(min).strFileName < comp And min < Last
            min = min + 1
        Loop

        Do While InfoHead(max).strFileName > comp And max > First
            max = max - 1
        Loop

        If min <= max Then
            aux = InfoHead(min)
            InfoHead(min) = InfoHead(max)
            InfoHead(max) = aux
            min = min + 1
            max = max - 1

        End If

    Loop
    
    If First < max Then Call Sort_Info_Headers(InfoHead, First, max)
    If min < Last Then Call Sort_Info_Headers(InfoHead, min, Last)

End Sub

''
' Searches for the specified InfoHeader.
'
' @param    ResourceFile A handler to the data file.
' @param    InfoHead The header searched.
' @param    FirstHead The first head to look.
' @param    LastHead The last head to look.
' @param    FileHeaderSize The bytes size of a FileHeader.
' @param    InfoHeaderSize The bytes size of a InfoHeader.
'
' @return   True if found.
'
' @remark   File must be already open.
' @remark   InfoHead must have set its file name to perform the search.

Private Function BinarySearch(ByRef ResourceFile As Integer, _
                              ByRef InfoHead As INFOHEADER, _
                              ByVal FirstHead As Long, _
                              ByVal LastHead As Long, _
                              ByVal FileHeaderSize As Long, _
                              ByVal InfoHeaderSize As Long) As Boolean

    '*****************************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modify Date: 08/21/2007
    'Searches for the specified InfoHeader
    '*****************************************************************
    Dim ReadingHead  As Long

    Dim ReadInfoHead As INFOHEADER
    
    Do Until FirstHead > LastHead
        ReadingHead = (FirstHead + LastHead) \ 2

        Get ResourceFile, FileHeaderSize + InfoHeaderSize * (ReadingHead - 1) + 1, ReadInfoHead

        If InfoHead.strFileName = ReadInfoHead.strFileName Then
            InfoHead = ReadInfoHead
            BinarySearch = True

            Exit Function

        Else

            If InfoHead.strFileName < ReadInfoHead.strFileName Then
                LastHead = ReadingHead - 1
            Else
                FirstHead = ReadingHead + 1

            End If

        End If

    Loop

End Function

''
' Retrieves the InfoHead of the specified graphic file.
'
' @param    ResourcePath The resource file folder.
' @param    FileName The graphic file name.
' @param    InfoHead The InfoHead where data is returned.
'
' @return   True if found.

Private Function Get_InfoHeader(ByRef ResourcePath As String, _
                                ByRef FileName As String, _
                                ByRef InfoHead As INFOHEADER, _
                                Optional Modo As Byte = 0) As Boolean

    '*****************************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modify Date: 16/07/2012 - ^[GS]^
    'Retrieves the InfoHead of the specified graphic file
    '*****************************************************************
    Dim ResourceFile As Integer

    Dim FileHead     As FILEHEADER
    
    On Local Error GoTo ErrHandler
    
    'Set InfoHeader we are looking for
    InfoHead.strFileName = UCase$(FileName)
   
    'Open the binary file
    ResourceFile = FreeFile()
    Open ResourcePath For Binary Access Read Lock Write As ResourceFile
    'Extract the FILEHEADER
    Get ResourceFile, 1, FileHead
        
    'Check the file for validity
    If LOF(ResourceFile) <> FileHead.lngFileSize Then
        MsgBox "Archivo de recursos dañado. " & ResourcePath, , "Error"
        Close ResourceFile

        Exit Function

    End If
        
    'Search for it!
    If BinarySearch(ResourceFile, InfoHead, 1, FileHead.lngNumFiles, Len(FileHead), Len(InfoHead)) Then
        Get_InfoHeader = True

    End If
        
    Close ResourceFile

    Exit Function

ErrHandler:
    Close ResourceFile
    
    Call MsgBox("Error al intentar leer el archivo " & ResourcePath & ". Razón: " & err.Number & " : " & err.Description, vbOKOnly, "Error")

End Function

''
' Compresses binary data avoiding data loses.
'
' @param    data() The data array.

Private Sub Compress_Data(ByRef data() As Byte, Optional Modo As Byte = 0)

    '*****************************************************************
    'Author: Juan Martín Dotuyo Dodero
    'Last Modify Date: 17/07/2012 - ^[GS]^
    'Compresses binary data avoiding data loses
    '*****************************************************************
    Dim Dimensions As Long

    Dim DimBuffer  As Long

    Dim BufTemp()  As Byte

    Dim LoopC      As Long
    
    Dimensions = UBound(data) + 1
    
    ' The worst case scenario, compressed info is 1.06 times the original - see zlib's doc for more info.
    DimBuffer = Dimensions * 1.06
    
    ReDim BufTemp(DimBuffer)
    
    Call compress(BufTemp(0), DimBuffer, data(0), Dimensions)
    
    Erase data
    
    ReDim data(DimBuffer - 1)
    ReDim Preserve BufTemp(DimBuffer - 1)
    
    data = BufTemp
    
    Erase BufTemp
    
    ' GSZAO - Seguridad
    If Modo = 0 And GrhUsaContra = True Then
        If UBound(GrhDatContra) <= UBound(data) And UBound(GrhDatContra) <> 0 Then

            For LoopC = 0 To UBound(GrhDatContra)
                data(LoopC) = data(LoopC) Xor GrhDatContra(LoopC)
            Next LoopC

        End If

    ElseIf Modo = 1 And MapsUsaContra = True Then

        If UBound(MapsDatContra) <= UBound(data) And UBound(MapsDatContra) <> 0 Then

            For LoopC = 0 To UBound(MapsDatContra)
                data(LoopC) = data(LoopC) Xor MapsDatContra(LoopC)
            Next LoopC

        End If

    End If

    ' GSZAO - Seguridad
    
End Sub

''
' Decompresses binary data.
'
' @param    data() The data array.
' @param    OrigSize The original data size.

Private Sub Decompress_Data(ByRef data() As Byte, _
                            ByVal OrigSize As Long, _
                            Optional Modo As Byte = 0)

    '*****************************************************************
    'Author: Juan Martín Dotuyo Dodero
    'Last Modify Date: 16/07/2012 - ^[GS]^
    'Decompresses binary data
    '*****************************************************************
    Dim BufTemp() As Byte

    Dim LoopC     As Integer
    
    ReDim BufTemp(OrigSize - 1)
    
    ' GSZAO - Seguridad
    If Modo = 0 And GrhUsaContra = True Then
        If UBound(GrhDatContra) <= UBound(data) And UBound(GrhDatContra) <> 0 Then

            For LoopC = 0 To UBound(GrhDatContra)
                data(LoopC) = data(LoopC) Xor GrhDatContra(LoopC)
            Next LoopC

        End If

    ElseIf Modo = 1 And MapsUsaContra = True Then

        If UBound(MapsDatContra) <= UBound(data) And UBound(MapsDatContra) <> 0 Then

            For LoopC = 0 To UBound(MapsDatContra)
                data(LoopC) = data(LoopC) Xor MapsDatContra(LoopC)
            Next LoopC

        End If

    End If

    ' GSZAO - Seguridad
    
    Call uncompress(BufTemp(0), OrigSize, data(0), UBound(data) + 1)
    
    ReDim data(OrigSize - 1)
    
    data = BufTemp
    
    Erase BufTemp

End Sub

''
' Retrieves a byte array with the compressed data from the specified file.
'
' @param    ResourcePath The resource file folder.
' @param    InfoHead The header specifiing the graphic file info.
' @param    data() The byte array to return data.
'
' @return   True if no error occurred.
'
' @remark   InfoHead must not be encrypted.
' @remark   Data is not desencrypted.

Public Function Get_File_RawData(ByRef ResourcePath As String, _
                                 ByRef InfoHead As INFOHEADER, _
                                 ByRef data() As Byte, _
                                 Optional Modo As Byte = 0) As Boolean

    '*****************************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modify Date: 16/07/2012 - ^[GS]^
    'Retrieves a byte array with the compressed data from the specified file
    '*****************************************************************
    Dim ResourceFile As Integer
    
    On Local Error GoTo ErrHandler
    
    'Size the Data array
    ReDim data(InfoHead.lngFileSize - 1)
    
    'Open the binary file
    ResourceFile = FreeFile
    Open ResourcePath For Binary Access Read Lock Write As ResourceFile
    'Get the data
    Get ResourceFile, InfoHead.lngFileStart, data
    'Close the binary file
    Close ResourceFile
    
    Get_File_RawData = True

    Exit Function

ErrHandler:
    Close ResourceFile

End Function

''
' Extract the specific file from a resource file.
'
' @param    ResourcePath The resource file folder.
' @param    InfoHead The header specifiing the graphic file info.
' @param    data() The byte array to return data.
'
' @return   True if no error occurred.
'
' @remark   Data is desencrypted.

Public Function Extract_File(ByRef ResourcePath As String, _
                             ByRef InfoHead As INFOHEADER, _
                             ByRef data() As Byte, _
                             Optional Modo As Byte = 0) As Boolean
    '*****************************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modify Date: 14/09/2012 - ^[GS]^
    'Extract the specific file from a resource file
    '*****************************************************************
    On Local Error GoTo ErrHandler
    
    If Get_File_RawData(ResourcePath, InfoHead, data, Modo) Then
        'Decompress all data
        'If InfoHead.lngFileSize < InfoHead.lngFileSizeUncompressed Then ' GSZAO
        Call Decompress_Data(data, InfoHead.lngFileSizeUncompressed, Modo)
        'End If
        
        Extract_File = True

    End If

    Exit Function

ErrHandler:
    Call MsgBox("Error al intentar decodificar recursos. Razón: " & err.Number & " : " & err.Description, vbOKOnly, "Error")

End Function

''
' Extracts all files from a resource file.
'
' @param    ResourcePath The resource file folder.
' @param    OutputPath The folder where graphic files will be extracted.
' @param    PrgBar The control that shows the process state.
'
' @return   True if no error occurred.

Public Function Extract_Files(ByRef ResourcePath As String, _
                              ByRef OutputPath As String, _
                              ByRef prgBar As ProgressBar, _
                              Optional Modo As Byte = 0) As Boolean

    '*****************************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modify Date: 17/07/2012 - ^[GS]^
    'Extracts all files from a resource file
    '*****************************************************************
    Dim LoopC         As Long

    Dim ResourceFile  As Integer

    Dim OutputFile    As Integer

    Dim SourceData()  As Byte

    Dim FileHead      As FILEHEADER

    Dim InfoHead()    As INFOHEADER

    Dim RequiredSpace As Currency
    
    On Local Error GoTo ErrHandler
    
    'Open the binary file
    ResourceFile = FreeFile()
    Open ResourcePath For Binary Access Read Lock Write As ResourceFile
    'Extract the FILEHEADER
    Get ResourceFile, 1, FileHead
    
    'Check the file for validity
    If LOF(ResourceFile) <> FileHead.lngFileSize Then
        Call MsgBox("Archivo de recursos dañado. " & ResourcePath, , "Error")
        Close ResourceFile

        Exit Function

    End If
        
    'Size the InfoHead array
    ReDim InfoHead(FileHead.lngNumFiles - 1)
        
    'Extract the INFOHEADER
    Get ResourceFile, , InfoHead
        
    'Check if there is enough hard drive space to extract all files
    For LoopC = 0 To UBound(InfoHead)
            
        RequiredSpace = RequiredSpace + InfoHead(LoopC).lngFileSizeUncompressed
    Next LoopC
        
    If RequiredSpace >= General_Drive_Get_Free_Bytes(Left$(App.path, 3)) Then
        Erase InfoHead
        Close ResourceFile
        Call MsgBox("No hay suficiente espacio en el disco para extraer los archivos.", , "Error")

        Exit Function

    End If

    Close ResourceFile
    
    'Update progress bar
    If Not prgBar Is Nothing Then
        prgBar.Value = 0
        prgBar.max = FileHead.lngNumFiles + 1

    End If
    
    'Extract all of the files from the binary file
    For LoopC = 0 To UBound(InfoHead)

        'Extract this file
        If Extract_File(ResourcePath, InfoHead(LoopC), SourceData) Then

            'Destroy file if it previuosly existed
            If FileExist(OutputPath & InfoHead(LoopC).strFileName, vbNormal) Then
                Call Kill(OutputPath & InfoHead(LoopC).strFileName)

            End If
            
            'Save it!
            OutputFile = FreeFile()
            Open OutputPath & InfoHead(LoopC).strFileName For Binary As OutputFile
            Put OutputFile, , SourceData
            Close OutputFile
            
            Erase SourceData
        Else
            Erase SourceData
            Erase InfoHead
            
            Call MsgBox("No se pudo extraer el archivo " & InfoHead(LoopC).strFileName, vbOKOnly, "Error")

            Exit Function

        End If
            
        'Update progress bar
        If Not prgBar Is Nothing Then prgBar.Value = prgBar.Value + 1
        DoEvents
    Next LoopC
    
    Erase InfoHead
    Extract_Files = True

    Exit Function

ErrHandler:
    Close ResourceFile
    Erase SourceData
    Erase InfoHead
    
    Call MsgBox("No se pudo extraer el archivo binario correctamente. Razón: " & err.Number & " : " & err.Description, vbOKOnly, "Error")

End Function

''
' Retrieves a byte array with the specified file data.
'
' @param    ResourcePath The resource file folder.
' @param    FileName The graphic file name.
' @param    data() The byte array to return data.
'
' @return   True if no error occurred.
'
' @remark   Data is desencrypted.

Public Function Get_File_Data(ByRef ResourcePath As String, _
                              ByRef FileName As String, _
                              ByRef data() As Byte, _
                              Optional Modo As Byte = 0) As Boolean

    '*****************************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modify Date: 16/07/2012 - ^[GS]^
    'Retrieves a byte array with the specified file data
    '*****************************************************************
    Dim InfoHead As INFOHEADER
    
    If Get_InfoHeader(ResourcePath, FileName, InfoHead, Modo) Then
        'Extract!
        Get_File_Data = Extract_File(ResourcePath, InfoHead, data, Modo)
    Else
        Get_File_Data = False

        'Call MsgBox("No se se encontro el recurso " & FileName)
    End If

End Function

''
' Retrieves image file data.
'
' @param    ResourcePath The resource file folder.
' @param    FileName The graphic file name.
' @param    bmpInfo The bitmap info structure.
' @param    data() The byte array to return data.
'
' @return   True if no error occurred.

Public Function Get_Image(ByRef ResourcePath As String, _
                          ByRef FileName As String, _
                          ByRef data() As Byte, _
                          Optional SoloPNG As Boolean = True) As Boolean

    '*****************************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modify Date: 09/10/2012 - ^[GS]^
    'Retrieves image file data
    '*****************************************************************
    Dim InfoHead  As INFOHEADER

    Dim ExistFile As Boolean
    
    ExistFile = False
    
    If SoloPNG = True Then
        If Get_InfoHeader(ResourcePath, FileName & ".PNG", InfoHead, 0) Then ' ¿BMP?
            FileName = FileName & ".PNG"
            ExistFile = True

        End If

    Else

        If Get_InfoHeader(ResourcePath, FileName & ".BMP", InfoHead, 0) Then ' ¿BMP?
            FileName = FileName & ".BMP"
            ExistFile = True
        ElseIf Get_InfoHeader(ResourcePath, FileName & ".PNG", InfoHead, 0) Then ' Existe PNG?
            FileName = FileName & ".PNG" ' usamos el PNG
            ExistFile = True

        End If

    End If
    
    If ExistFile = True Then
        If Extract_File(ResourcePath, InfoHead, data, 0) Then Get_Image = True
    Else
        'Call LogError("Get_Image::No se encontro el recurso " & FileName)
        Call MsgBox("Get_Image::No se encontro el recurso " & FileName)

    End If

End Function

''
' Compare two byte arrays to detect any difference.
'
' @param    data1() Byte array.
' @param    data2() Byte array.
'
' @return   True if are equals.

Private Function Compare_Datas(ByRef data1() As Byte, ByRef data2() As Byte) As Boolean

    '*****************************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modify Date: 02/11/2007
    'Compare two byte arrays to detect any difference
    '*****************************************************************
    Dim Length As Long

    Dim act    As Long
    
    Length = UBound(data1) + 1
    
    If (UBound(data2) + 1) = Length Then

        While act < Length

            If data1(act) Xor data2(act) Then Exit Function
            
            act = act + 1

        Wend
        
        Compare_Datas = True

    End If

End Function

''
' Retrieves the next InfoHeader.
'
' @param    ResourceFile A handler to the resource file.
' @param    FileHead The reource file header.
' @param    InfoHead The returned header.
' @param    ReadFiles The number of headers that have already been read.
'
' @return   False if there are no more headers tu read.
'
' @remark   File must be already open.
' @remark   Used to walk through the resource file info headers.
' @remark   The number of read files will increase although there is nothing else to read.
' @remark   InfoHead is encrypted.

Private Function ReadNext_InfoHead(ByRef ResourceFile As Integer, _
                                   ByRef FileHead As FILEHEADER, _
                                   ByRef InfoHead As INFOHEADER, _
                                   ByRef ReadFiles As Long) As Boolean
    '*****************************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modify Date: 08/24/2007
    'Reads the next InfoHeader
    '*****************************************************************

    If ReadFiles < FileHead.lngNumFiles Then
        'Read header
        Get ResourceFile, Len(FileHead) + Len(InfoHead) * ReadFiles + 1, InfoHead
        
        'Update
        ReadNext_InfoHead = True

    End If
    
    ReadFiles = ReadFiles + 1

End Function

''
' Compares two resource versions and makes a patch file.
'
' @param    NewResourcePath The actual reource file folder.
' @param    OldResourcePath The previous reource file folder.
' @param    OutputPath The patchs file folder.
' @param    PrgBar The control that shows the process state.
'
' @return   True if no error occurred.

Public Function Make_Patch(ByRef NewResourcePath As String, _
                           ByRef OldResourcePath As String, _
                           ByRef OutputPath As String, _
                           ByRef prgBar As ProgressBar, _
                           Optional Modo As Byte = 0) As Boolean

    '*****************************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modify Date: 17/07/2012 - ^[GS]^
    'Compares two resource versions and make a patch file
    '*****************************************************************
    Dim NewResourceFile     As Integer

    Dim NewResourceFilePath As String

    Dim NewFileHead         As FILEHEADER

    Dim NewInfoHead         As INFOHEADER

    Dim NewReadFiles        As Long

    Dim NewReadNext         As Boolean
    
    Dim OldResourceFile     As Integer

    Dim OldResourceFilePath As String

    Dim OldFileHead         As FILEHEADER

    Dim OldInfoHead         As INFOHEADER

    Dim OldReadFiles        As Long

    Dim OldReadNext         As Boolean
    
    Dim OutputFile          As Integer

    Dim OutputFilePath      As String

    Dim data()              As Byte

    Dim auxData()           As Byte

    Dim Instruction         As Byte
    
    'Set up the error handler
    On Local Error GoTo ErrHandler

    If Modo = 0 Then
        NewResourceFilePath = NewResourcePath
        OldResourceFilePath = OldResourcePath
        OutputFilePath = OutputPath & GRH_PATCH_FILE
    ElseIf Modo = 1 Then
        NewResourceFilePath = NewResourcePath
        OldResourceFilePath = OldResourcePath
        OutputFilePath = OutputPath & MAPS_PATCH_FILE

    End If
    
    'Open the old binary file
    OldResourceFile = FreeFile
    Open OldResourceFilePath For Binary Access Read Lock Write As OldResourceFile
        
    'Get the old FileHeader
    Get OldResourceFile, 1, OldFileHead

    'Check the file for validity
    If LOF(OldResourceFile) <> OldFileHead.lngFileSize Then
        Call MsgBox("Archivo de recursos anterior dañado. " & OldResourceFilePath, , "Error")
        Close OldResourceFile

        Exit Function

    End If
        
    'Open the new binary file
    NewResourceFile = FreeFile()
    Open NewResourceFilePath For Binary Access Read Lock Write As NewResourceFile
            
    'Get the new FileHeader
    Get NewResourceFile, 1, NewFileHead

    'Check the file for validity
    If LOF(NewResourceFile) <> NewFileHead.lngFileSize Then
        Call MsgBox("Archivo de recursos anterior dañado. " & NewResourceFilePath, , "Error")
        Close NewResourceFile
        Close OldResourceFile

        Exit Function

    End If
            
    'Destroy file if it previuosly existed
    If LenB(Dir(OutputFilePath, vbNormal)) <> 0 Then Kill OutputFilePath
            
    'Open the patch file
    OutputFile = FreeFile()
    Open OutputFilePath For Binary Access Read Write As OutputFile
                
    If Not prgBar Is Nothing Then
        prgBar.Value = 0
        prgBar.max = (OldFileHead.lngNumFiles + NewFileHead.lngNumFiles) + 1

    End If
                
    'put previous file version (unencrypted)
    Put OutputFile, , OldFileHead.lngFileVersion
                
    'Put the new file header
    Put OutputFile, , NewFileHead

    'Try to read old and new first files
    If ReadNext_InfoHead(OldResourceFile, OldFileHead, OldInfoHead, OldReadFiles) And ReadNext_InfoHead(NewResourceFile, NewFileHead, NewInfoHead, NewReadFiles) Then
                    
        'Update
        prgBar.Value = prgBar.Value + 2
                    
        Do 'Main loop

            'Comparisons are between encrypted names, for ordering issues
            If OldInfoHead.strFileName = NewInfoHead.strFileName Then

                'Get old file data
                Call Get_File_RawData(OldResourcePath, OldInfoHead, auxData, Modo)
                            
                'Get new file data
                Call Get_File_RawData(NewResourcePath, NewInfoHead, data, Modo)
                            
                If Not Compare_Datas(data, auxData) Then
                    'File was modified
                    Instruction = PatchInstruction.Modify_File
                    Put OutputFile, , Instruction
                                
                    'Write header
                    Put OutputFile, , NewInfoHead
                                
                    'Write data
                    Put OutputFile, , data

                End If
                            
                'Read next OldResource
                If Not ReadNext_InfoHead(OldResourceFile, OldFileHead, OldInfoHead, OldReadFiles) Then

                    Exit Do

                End If
                            
                'Read next NewResource
                If Not ReadNext_InfoHead(NewResourceFile, NewFileHead, NewInfoHead, NewReadFiles) Then
                    'Reread last OldInfoHead
                    OldReadFiles = OldReadFiles - 1

                    Exit Do

                End If
                            
                'Update
                If Not prgBar Is Nothing Then prgBar.Value = prgBar.Value + 2
                        
            ElseIf OldInfoHead.strFileName < NewInfoHead.strFileName Then
                            
                'File was deleted
                Instruction = PatchInstruction.Delete_File
                Put OutputFile, , Instruction
                Put OutputFile, , OldInfoHead
                            
                'Read next OldResource
                If Not ReadNext_InfoHead(OldResourceFile, OldFileHead, OldInfoHead, OldReadFiles) Then
                    'Reread last NewInfoHead
                    NewReadFiles = NewReadFiles - 1

                    Exit Do

                End If
                            
                'Update
                If Not prgBar Is Nothing Then prgBar.Value = prgBar.Value + 1
                        
            Else
                            
                'New file
                Instruction = PatchInstruction.Create_File
                Put OutputFile, , Instruction
                Put OutputFile, , NewInfoHead
                                     
                'Get file data
                Call Get_File_RawData(NewResourcePath, NewInfoHead, data, Modo)
                            
                'Write data
                Put OutputFile, , data
                            
                'Read next NewResource
                If Not ReadNext_InfoHead(NewResourceFile, NewFileHead, NewInfoHead, NewReadFiles) Then
                    'Reread last OldInfoHead
                    OldReadFiles = OldReadFiles - 1

                    Exit Do

                End If
                            
                'Update
                If Not prgBar Is Nothing Then prgBar.Value = prgBar.Value + 1

            End If
                        
            DoEvents
        Loop
                
    Else
        'if at least one is empty
        OldReadFiles = 0
        NewReadFiles = 0

    End If
                
    'Read everything?
    While ReadNext_InfoHead(OldResourceFile, OldFileHead, OldInfoHead, OldReadFiles)

        'Delete file
        Instruction = PatchInstruction.Delete_File
        Put OutputFile, , Instruction
        Put OutputFile, , OldInfoHead
                    
        'Update
        If Not prgBar Is Nothing Then prgBar.Value = prgBar.Value + 1
        DoEvents

    Wend
                
    'Read everything?
    While ReadNext_InfoHead(NewResourceFile, NewFileHead, NewInfoHead, NewReadFiles)

        'Create file
        Instruction = PatchInstruction.Create_File
        Put OutputFile, , Instruction
        Put OutputFile, , NewInfoHead
                    
        'Get file data
        Call Get_File_RawData(NewResourcePath, NewInfoHead, data, Modo)
        'Write data
        Put OutputFile, , data
                    
        'Update
        If Not prgBar Is Nothing Then prgBar.Value = prgBar.Value + 1
        DoEvents

    Wend
            
    'Close the patch file
    Close OutputFile
        
    'Close the new binary file
    Close NewResourceFile
    
    'Close the old binary file
    Close OldResourceFile
    
    Make_Patch = True

    Exit Function

ErrHandler:
    Close OutputFile
    Close NewResourceFile
    Close OldResourceFile
    
    Call MsgBox("No se pudo terminar de crear el parche. Razón: " & err.Number & " : " & err.Description, vbOKOnly, "Error")

End Function

''
' Follows patches instructions to update a resource file.
'
' @param    ResourcePath The reource file folder.
' @param    PatchPath The patch file folder.
' @param    PrgBar The control that shows the process state.
'
' @return   True if no error occurred.
Public Function Apply_Patch(ByRef ResourcePath As String, _
                            ByRef PatchPath As String, _
                            Optional Modo As Byte = 0) As Boolean

    '*****************************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modify Date: 17/07/2012 - ^[GS]^
    'Follows patches instructions to update a resource file
    '*****************************************************************
    Dim ResourceFile       As Integer

    Dim ResourceFilePath   As String

    Dim FileHead           As FILEHEADER

    Dim InfoHead           As INFOHEADER

    Dim ResourceReadFiles  As Long

    Dim EOResource         As Boolean

    Dim PatchFile          As Integer

    Dim PatchFilePath      As String

    Dim PatchFileHead      As FILEHEADER

    Dim PatchInfoHead      As INFOHEADER

    Dim Instruction        As Byte

    Dim OldResourceVersion As Long

    Dim OutputFile         As Integer

    Dim OutputFilePath     As String

    Dim data()             As Byte

    Dim WrittenFiles       As Long

    Dim DataOutputPos      As Long

    On Local Error GoTo ErrHandler

    If Modo = 0 Then
        ResourceFilePath = ResourcePath
        PatchFilePath = PatchPath
        OutputFilePath = ResourcePath & "tmp"
    ElseIf Modo = 1 Then
        ResourceFilePath = ResourcePath
        PatchFilePath = PatchPath
        OutputFilePath = ResourcePath & "tmp"

    End If
    
    'Open the old binary file
    ResourceFile = FreeFile()
    Open ResourceFilePath For Binary Access Read Lock Write As ResourceFile
        
    'Read the old FileHeader
    Get ResourceFile, , FileHead

    'Check the file for validity
    If LOF(ResourceFile) <> FileHead.lngFileSize Then
        Call MsgBox("Archivo de recursos anterior dañado. " & ResourceFilePath, , "Error")
        Close ResourceFile

        Exit Function

    End If
        
    'Open the patch file
    PatchFile = FreeFile()
    Open PatchFilePath For Binary Access Read Lock Write As PatchFile
            
    'Get previous file version
    Get PatchFile, , OldResourceVersion
            
    'Check the file version
    If OldResourceVersion <> FileHead.lngFileVersion Then
        Call MsgBox("Incongruencia en versiones.", , "Error")
        Close ResourceFile
        Close PatchFile

        Exit Function

    End If
            
    'Read the new FileHeader
    Get PatchFile, , PatchFileHead
            
    'Destroy file if it previuosly existed
    If FileExist(OutputFilePath, vbNormal) Then Call Kill(OutputFilePath)
            
    'Open the patch file
    OutputFile = FreeFile()
    Open OutputFilePath For Binary Access Read Write As OutputFile
                
    'Save the file header
    Put OutputFile, , PatchFileHead
                
    'Update
    DataOutputPos = Len(FileHead) + Len(InfoHead) * PatchFileHead.lngNumFiles + 1
                
    'Process loop
    While Loc(PatchFile) < LOF(PatchFile)
                    
        'Get the instruction
        Get PatchFile, , Instruction
        'Get the InfoHead
        Get PatchFile, , PatchInfoHead
                    
        Do
            EOResource = Not ReadNext_InfoHead(ResourceFile, FileHead, InfoHead, ResourceReadFiles)
                        
            'Comparison is performed among encrypted names for ordering issues
            If Not EOResource And InfoHead.strFileName < PatchInfoHead.strFileName Then
                            
                'GetData and update InfoHead
                Call Get_File_RawData(ResourcePath, InfoHead, data, Modo)
                InfoHead.lngFileStart = DataOutputPos
                                           
                'Save file!
                Put OutputFile, Len(FileHead) + Len(InfoHead) * WrittenFiles + 1, InfoHead
                Put OutputFile, DataOutputPos, data
                            
                'Update
                DataOutputPos = DataOutputPos + UBound(data) + 1
                WrittenFiles = WrittenFiles + 1
            Else

                Exit Do

            End If

        Loop
                    
        Select Case Instruction

                'Delete
            Case PatchInstruction.Delete_File

                If InfoHead.strFileName <> PatchInfoHead.strFileName Then
                    err.Description = "Incongruencia en archivos de recurso"
                    GoTo ErrHandler

                End If
                        
                'Create
            Case PatchInstruction.Create_File

                If (InfoHead.strFileName > PatchInfoHead.strFileName) Or EOResource Then
                                
                    'Get file data
                    ReDim data(PatchInfoHead.lngFileSize - 1)
                    Get PatchFile, , data
                                
                    'Save it
                    Put OutputFile, Len(FileHead) + Len(InfoHead) * WrittenFiles + 1, PatchInfoHead
                    Put OutputFile, DataOutputPos, data
                                
                    'Reanalize last Resource InfoHead
                    EOResource = False
                    ResourceReadFiles = ResourceReadFiles - 1
                                
                    'Update
                    DataOutputPos = DataOutputPos + UBound(data) + 1
                    WrittenFiles = WrittenFiles + 1
                Else
                    err.Description = "Incongruencia en archivos de recurso"
                    GoTo ErrHandler

                End If
                        
                'Modify
            Case PatchInstruction.Modify_File

                If InfoHead.strFileName = PatchInfoHead.strFileName Then

                    'Get file data
                    ReDim data(PatchInfoHead.lngFileSize - 1)
                    Get PatchFile, , data
                                             
                    'Save it
                    Put OutputFile, Len(FileHead) + Len(InfoHead) * WrittenFiles + 1, PatchInfoHead
                    Put OutputFile, DataOutputPos, data
                                
                    'Update
                    DataOutputPos = DataOutputPos + UBound(data) + 1
                    WrittenFiles = WrittenFiles + 1
                Else
                    err.Description = "Incongruencia en archivos de recurso"
                    GoTo ErrHandler

                End If

        End Select
                    
        DoEvents

    Wend
                
    'Read everything?
    While ReadNext_InfoHead(ResourceFile, FileHead, InfoHead, ResourceReadFiles)

        'GetData and update InfoHeader
        Call Get_File_RawData(ResourcePath, InfoHead, data, Modo)
        InfoHead.lngFileStart = DataOutputPos
                    
        'Save file!
        Put OutputFile, Len(FileHead) + Len(InfoHead) * WrittenFiles + 1, InfoHead
        Put OutputFile, DataOutputPos, data
                    
        'Update
        DataOutputPos = DataOutputPos + UBound(data) + 1
        WrittenFiles = WrittenFiles + 1
        DoEvents

    Wend
            
    'Close the patch file
    Close OutputFile
        
    'Close the new binary file
    Close PatchFile
    
    'Close the old binary file
    Close ResourceFile
    
    'Check integrity
    If (PatchFileHead.lngNumFiles = WrittenFiles) Then

        'Replace File
        Call Kill(ResourceFilePath)
        Name OutputFilePath As ResourceFilePath

    Else
        err.Description = "Falla al procesar parche"
        GoTo ErrHandler

    End If
    
    Apply_Patch = True

    Exit Function

ErrHandler:
    Close OutputFile
    Close PatchFile
    Close ResourceFile

    'Destroy file if created
    If FileExist(OutputFilePath, vbNormal) Then Call Kill(OutputFilePath)
    
    Call MsgBox("No se pudo parchear. Razón: " & err.Number & " : " & err.Description, vbOKOnly, "Error")

End Function

Private Function AlignScan(ByVal inWidth As Long, ByVal inDepth As Integer) As Long
    '*****************************************************************
    'Author: Unknown
    'Last Modify Date: Unknown
    '*****************************************************************
    AlignScan = (((inWidth * inDepth) + &H1F) And Not &H1F&) \ &H8

End Function

''
' Retrieves the version number of a given resource file.
'
' @param    ResourceFilePath The resource file complete path.
'
' @return   The version number of the given file.

Public Function GetVersion(ByVal ResourceFilePath As String) As Long

    '*****************************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 11/23/2008
    '
    '*****************************************************************
    Dim ResourceFile As Integer

    Dim FileHead     As FILEHEADER
    
    ResourceFile = FreeFile()
    Open ResourceFilePath For Binary Access Read Lock Write As ResourceFile
    'Extract the FILEHEADER
    Get ResourceFile, 1, FileHead
        
    Close ResourceFile
    
    GetVersion = FileHead.lngFileVersion

End Function

Public Function TestZLib() As Boolean ' GSZAO

    '*****************************************************************
    'Author: ^[GS]^
    'Last Modify Date: 19/06/2011
    '*****************************************************************
    On Error GoTo ErrHandler
    
    Dim data() As Byte

    Dim lnD    As Integer
    
    data = ""
    lnD = UBound(data) + 1
    Call Decompress_Data(data, lnD)
    
    TestZLib = True
    
    Exit Function

ErrHandler:

    TestZLib = False

End Function


Public Function ArrayToPicture(inArray() As Byte, _
                               offset As Long, _
                               Size As Long) As IPicture
          
    Dim o_hMem        As Long

    Dim o_lpMem       As Long

    Dim aGUID(0 To 3) As Long

    Dim IIStream      As IUnknown
          
    aGUID(0) = &H7BF80980
    aGUID(1) = &H101ABF32
    aGUID(2) = &HAA00BB8B
    aGUID(3) = &HAB0C3000
          
    o_hMem = GlobalAlloc(&H2&, Size)

    If Not o_hMem = 0& Then
        o_lpMem = GlobalLock(o_hMem)

        If Not o_lpMem = 0& Then
            CopyMemory ByVal o_lpMem, inArray(offset), Size
            Call GlobalUnlock(o_hMem)

            If CreateStreamOnHGlobal(o_hMem, 1&, IIStream) = 0& Then
                Call OleLoadPicture(ByVal ObjPtr(IIStream), 0&, 0&, aGUID(0), ArrayToPicture)

            End If

        End If

    End If

End Function
