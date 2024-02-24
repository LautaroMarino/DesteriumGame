Attribute VB_Name = "mUpdateWeb"
Option Explicit

 Public Function Cabecera_Init() As String

    Dim Text As String

    Text = "<html><head>"
    Text = Text & "<meta http-equiv=" + Chr(34) + "Expires" + Chr(34) + " content=" + Chr(34) + "0" + Chr(34) + ">" + vbCrLf
    Text = Text & "<meta http-equiv=" + Chr(34) + "Last-Modified" + Chr(34) + " content=" + Chr(34) + "0" + Chr(34) + ">" + vbCrLf
    Text = Text & "<meta http-equiv=" + Chr(34) + "Cache-Control" + Chr(34) + " content=" + Chr(34) + "no-cache, mustrevalidate" + Chr(34) + ">" + vbCrLf
    Text = Text & "<meta http-equiv=" + Chr(34) + "Pragma" + Chr(34) + " content=" + Chr(34) + "no-cache" + Chr(34) + ">" + vbCrLf
    Text = Text & "</head> <body>"

    Cabecera_Init = Text
End Function

Public Function Cabecera_Table_Init() As String

        Dim Text As String

        Text = "<html><head>"
        Text = Text & "<meta http-equiv=" + Chr(34) + "Expires" + Chr(34) + " content=" + Chr(34) + "0" + Chr(34) + ">"
        Text = Text & "<meta http-equiv=" + Chr(34) + "Last-Modified" + Chr(34) + " content=" + Chr(34) + "0" + Chr(34) + ">"
        Text = Text & "<meta http-equiv=" + Chr(34) + "Cache-Control" + Chr(34) + " content=" + Chr(34) + "no-cache, mustrevalidate" + Chr(34) + ">"
        Text = Text & "<meta http-equiv=" + Chr(34) + "Pragma" + Chr(34) + " content=" + Chr(34) + "no-cache" + Chr(34) + ">"

        Text = Text & "<link rel = " + Chr(34) + "stylesheet" + Chr(34) + " type=" + Chr(34) + "Text/css" + Chr(34) + "href=" + Chr(34) + "CSS.css" + Chr(34) + ">"
        Text = Text & "<style type = " + Chr(34) + "text/css" + Chr(34) + ">"
        Text = Text & "body { background-color: #202020; }"

        Text = Text & "table{border-collapse:collapse;}"
        Text = Text & "th, tr, td{"
        Text = Text & "border:1px solid #000;"
        Text = Text & "width:150px;"
        Text = Text & "height:45px;"
         Text = Text & "vertical-align: middle;"
         Text = Text & "Text-align: center;"
        Text = Text & "}"
        Text = Text & "td{"
        Text = Text & "     color: #E3E1E1;"
        Text = Text & "     text-shadow: 2px 2px #000000;"
        Text = Text & "}"
        Text = Text & "th{"
        Text = Text & " Color: #fff;"
        Text = Text & " background-Color:  #252525;"
        Text = Text & "}"

        Text = Text & "tr:     nth-child(odd) td{"
        Text = Text & " background-Color: #6f6a6a;"
         
        Text = Text & "}"
        Text = Text & "</style>"

        Text = Text & "</head>"

        Cabecera_Table_Init = Text
    End Function

    Public Function Cabecera_End() As String
        Cabecera_End = "</body>" + vbCrLf + "</html>"
    End Function
    

 Public Sub Create_Stats_General()
 
    On Error GoTo ErrHandler
    Dim Text As String
    Dim Onlines As String
    Dim Record As String
    Dim Version As String
    
    Onlines = CStr(NumUsers + UsersBot)
    Record = CStr(RECORDusuarios)
    Version = Version
    Text = Cabecera_Init()


    Text = Text & "<font color=" + Chr(34) + "white" + Chr(34) + "> Actualmente hay</font>"
    Text = Text & "<font color=" + Chr(34) + "green" + Chr(34) + "> <b>" + Onlines + " </b></font>"
    Text = Text & "<font color=" + Chr(34) + "white" + Chr(34) + "> Jugadores en linea.</font>"

    Text = Text & "<font color=" + Chr(34) + "white" + Chr(34) + "> El record de usuarios conectados simultaneamente es de </font>"
    Text = Text & "<font color=" + Chr(34) + "teal" + Chr(34) + "> <b>" + Record + ". </b></font>"

    Text = Text & "<font color=" + Chr(34) + "white" + Chr(34) + "> La versión actual es la <b>" + Version + "</b></font>"

    Text = Text & Cabecera_End()

    Open App.Path & "\STATS\general.html" For Append As #1
    Print #1, Text
    Close #1

    Exit Sub
ErrHandler:
    Call LogError("Ocurrió u nerror al cargar estadísticas Create_Stats_General")
    
End Sub

