Attribute VB_Name = "mStats"
Option Explicit

Public Sub Stats_Generate()
    
    ' Generamos las estadísticas generales
    Call Stats_Generate_General
    
    ' Generamos las estadísticas de Ranking
    Call Stats_Generate_Rank
    
    ' Generamos las estadísticas de Castillos
    Call Stats_Generate_Castle
    
End Sub


Private Sub Stats_Generate_General()
    Dim strTemp As String
    Dim File As Integer
    Const FilePath As String = "/stats/stats.html"
    
    Const Comillas_Html As String = "&quot;"
    
    strTemp = "<body>" & vbCrLf
    strTemp = strTemp & "<p style=&quot;color:rgb(255,255,255);&quot;><b>Jugadores online:</b> " & NumUsers & "</p>" & vbCrLf
    strTemp = strTemp & "</body>"
      
    File = FreeFile

    Open App.Path & FilePath For Output As File
          Write #File, strTemp
    Close
End Sub

Private Sub Stats_Generate_Rank()

End Sub

Private Sub Stats_Generate_Castle()

End Sub
