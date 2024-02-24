Attribute VB_Name = "MySql"
Option Explicit

Public SQL As ADODB.Connection
Public RS As ADODB.Recordset

Private Const SERVER_INDEX As Byte = 1

Public Sub MySql_Open()
    On Error GoTo ErrHandler
    
    If SQL Is Nothing Then
        Set SQL = New ADODB.Connection
    End If
    
    If SQL.State = adStateClosed Then
        SQL.ConnectionString = "Driver={MYSQL ODBC 3.51 Driver};" & "SERVER=45.235.98.111;" & "DATABASE=argentumgame;" & "UID=desterium;PWD=U3pFj*xWUMMIOFt); OPTION=3"
        SQL.CursorLocation = adUseClient
        SQL.Open
    End If
    
    Exit Sub
ErrHandler:
End Sub

' # Actualiza la información del servidor.
Public Sub MySql_UpdateServer()
    On Error GoTo ErrHandler
    
    
    
    If SQL.State = adStateClosed Then
        MySql_Open
    End If
    
    Set RS = New ADODB.Recordset
    SQL.Execute "UPDATE servers SET onlines = '" & NumUsers + UsersBot & "', record = '" & RECORDusuarios & "' WHERE ID = '" & SERVER_INDEX & "'"
    
    Exit Sub
ErrHandler:
    LogError "Error en MySql_UpdateServer(): " & Err.description
End Sub
