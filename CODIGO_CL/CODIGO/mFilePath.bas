Attribute VB_Name = "mFilePath"
Option Explicit

Public Function Init_FilePath() As String
    #If Classic = 1 Then
        Init_FilePath = App.path & "\resource\init\"
    #Else
        Init_FilePath = App.path & "\resource\initclassic\"
    #End If
End Function

Public Function Maps_FilePath() As String
    #If Classic = 1 Then
        Maps_FilePath = "\resource\maps\"
    #Else
        Maps_FilePath = "\resource\mapsclassic\"
    #End If
End Function
Public Function MiniMap_FilePath() As String

#If ModoBig > 0 Then
    #If Classic = 1 Then
        MiniMap_FilePath = App.path & "\resource\minimapbig\"
    #Else
        MiniMap_FilePath = App.path & "\resource\minimapbigclassic\"
    #End If

#Else
    #If Classic = 1 Then
        MiniMap_FilePath = App.path & "\resource\minimap\"
    #Else
        MiniMap_FilePath = App.path & "\resource\minimapclassic\"
    #End If
    
    #End If
End Function
Public Function Drops_FilePath() As String
    #If Classic = 1 Then
            Drops_FilePath = IniPath & "server\server_drops.ind"
    #Else
            Drops_FilePath = IniPath & "server\server_drops_classic.ind"
    #End If
End Function
Public Sub Folder_Renew(ByVal filePath As String)
        '<EhHeader>
        On Error GoTo Delete_file_Err
        '</EhHeader>

        'declaras el tipo de variable
        Dim m_fso As FileSystemObject

        'seteas la variable del objeto
100     Set m_fso = New FileSystemObject

        'comprobas si existe la carpeta
102     If m_fso.FolderExists(filePath) Then

            ' Por si se genera algun error
            On Error Resume Next

            ' Como la carpeta existe la borras
104         m_fso.DeleteFolder Left$(filePath, Len(filePath) - 1)

            ' Creas la carpeta de nuevo
106         m_fso.CreateFolder filePath

        Else
            'Si no existe la creas
108         m_fso.CreateFolder filePath

        End If
            
            
        Set m_fso = Nothing
        '-----------------
        '<EhFooter>
        Exit Sub

Delete_file_Err:
           Set m_fso = Nothing
        LogError err.Description & vbCrLf & _
               "in ARGENTUM.mFilePath.Delete_file " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
