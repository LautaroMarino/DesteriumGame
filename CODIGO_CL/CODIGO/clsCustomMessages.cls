VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCustomMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' clsCustomMessages.cls - User-defined messages with shortcut.
' Allows the creation and serialization of user defined messages to be used with shortcuts.
'
' Designed and implemented by Juan Mart�n Sotuyo Dodero (Maraxus)
' (juansotuyo@gmail.com)
'**************************************************************

'**************************************************************************
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
'**************************************************************************

Option Explicit
Option Base 0



''
' Defines the number of messages that can be created.
Private Const MESSAGE_COUNT               As Byte = 10

''
' Defines the maximum length allowed for any custom message.
Private Const MAX_MESSAGE_LENGTH          As Byte = 120

''
' The error code thrown when the message being set is invalid.
Private Const INVALID_MESSAGE             As Long = vbObjectError + 123

''
' Relative path from the application in which custom message are stored.
Private Const MESSAGE_FILE_PATH           As String = "/resource/init/messages.txt"

''
'Array of defined messages.
Private CustomMessages(MESSAGE_COUNT - 1) As String

''
' Class constructor. Initializes a new clsCustomMessage

Private Sub Class_Initialize()

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 07/26/07
    'Initializes the messages by loading them from the config file
    '***************************************************
    On Error GoTo ErrHandler

    Dim handle As Integer

    Dim I      As Long
    
    'Load custom messages
    handle = FreeFile()
    
    Open App.path & MESSAGE_FILE_PATH For Input As handle
    
    Do While (I < MESSAGE_COUNT) And (Not EOF(handle))
        Line Input #handle, CustomMessages(I)
        
        'Purge invalid messages
        If Not MessageIsValid(CustomMessages(I)) Then
            CustomMessages(I) = vbNullString
        End If
        
        I = I + 1
    Loop
    
    Close handle
    
    'If the file does not exist (only possible error),
    'then we have nothing to worry. All messages will be empty
ErrHandler:
End Sub

''
' Class desconstructor. Serializes messages to

Private Sub Class_Terminate()

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 07/26/07
    'Stores all messages back into config file for later use.
    '***************************************************
    Dim handle        As Integer

    Dim I             As Long

    Dim directories() As String

    Dim path          As String
    
    'Make sure the path exists
    directories = Split(MESSAGE_FILE_PATH, "/")
    path = App.path
    
    For I = 0 To UBound(directories()) - 1
        path = path & "/" & directories(I)
        
        If LenB(Dir(path, vbDirectory)) = 0 Then
            Call MkDir(path)
        End If

    Next I
    
    'Serialize custom messages
    handle = FreeFile()
    
    Open App.path & MESSAGE_FILE_PATH For Output As handle
    
    For I = 0 To MESSAGE_COUNT - 1
        Print #handle, CustomMessages(I)
    Next I
    
    Close handle
End Sub

''
' Retrieves the requested message.
'
' @param    index   The message index to retrieve.
' @return   The message requested.

Public Property Get Message(ByVal Index As Integer) As String

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 07/26/07
    'Retrieves the requested message.
    '***************************************************
    If Index < 0 Or Index >= MESSAGE_COUNT Then

        Exit Property

    End If
    
    Message = CustomMessages(Index)
End Property

''
' Sets the requested message.
'
' @param    index   The message index to be set.
' @param    nValue  The message to be stored.

Public Property Let Message(ByVal Index As Integer, ByVal nValue As String)

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 07/26/07
    'Sets the given message index with the requested message.
    '***************************************************
    If Index < 0 Or Index >= MESSAGE_COUNT Then

        Exit Property

    End If
    
    If Not MessageIsValid(nValue) Then
        Call err.Raise(INVALID_MESSAGE)

        Exit Property

    End If
    
    CustomMessages(Index) = nValue
End Property

''
' Retrieves the error code of invalid messages.
'
' @return   The error code triggered when an invalid message is found.

Public Property Get InvalidMessageErrCode() As Long
    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 07/26/07
    'Retrieves the INVALID_MESSAGE error code.
    '***************************************************
    InvalidMessageErrCode = INVALID_MESSAGE
End Property

''
' Checks if the given message is valid or not.
'
' @param    msg   The message to be checked.
' @return   True if the message is valid, False otherwise.

Private Function MessageIsValid(ByRef msg As String) As Boolean

    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
    'Last Modification: 08/07/07
    'Last Modified by: Rapsodius
    '
    'Erased line "msg=StrConv(msg,vbUnicode)", because
    'it was causing problems and detected as many messages
    'as wrong with nothing bad on them
    '***************************************************
    Dim I     As Long

    Dim Char  As String * 1

    Dim ascii As Byte
    
    'Get rid of spaces
    msg = Trim$(msg)
        
    'Empty messages are valid
    If LenB(msg) = 0 Then
        MessageIsValid = True

        Exit Function

    End If
    
    For I = 1 To Len(msg)
        Char = mid$(msg, I, 1)
        ascii = Asc(Char)
        
        'check for invalid chars...
        If ascii < 31 Then Exit Function
        
        If ascii > 126 And Char <> "�" And Char <> "�" And Char <> "�" And Char <> "�" And Char <> "�" And Char <> "�" And Char <> "�" And Char <> "�" And Char <> "�" And Char <> "�" And Char <> "�" And Char <> "�" And Char <> "�" And Char <> "�" And Char <> "�" And Char <> "�" And Char <> "�" And Char <> "�" And Char <> "�" And Char <> "�" And Char <> "�" And Char <> "�" And Char <> "�" And Char <> "�" And Char <> "�" And Char <> "�" And Char <> "�" And Char <> "�" And Char <> "�" And Char <> "�" And Char <> "�" And Char <> "�" And Char <> "�" And Char <> "�" And Char <> "�" And Char <> "�" And Char <> "�" And Char <> "�" And Char <> "�" And Char <> "�" And Char <> "�" And Char <> "�" And Char <> "�" And Char <> "�" Then

            Exit Function

        End If

    Next I
    
    'Discard extremely long messages
    If Len(msg) > MAX_MESSAGE_LENGTH Then Exit Function
    
    MessageIsValid = True
End Function
