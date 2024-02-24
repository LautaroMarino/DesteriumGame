Attribute VB_Name = "mInvasion"
Option Explicit

Public Type t_Rectangle
    X1 As Integer
    Y1 As Integer
    X2 As Integer
    Y2 As Integer
End Type

Type t_SpawnBox
    TopLeft As WorldPos
    BottomRight As WorldPos
    Heading As eHeading
    CoordMuralla As Integer
    LegalBox As t_Rectangle
End Type

' WyroX: Devuelve si X e Y están dentro del Rectangle
Public Function InsideRectangle(r As t_Rectangle, ByVal X As Integer, ByVal Y As Integer) As Boolean
100     If X < r.X1 Then Exit Function
102     If X > r.X2 Then Exit Function
104     If Y < r.Y1 Then Exit Function
106     If Y > r.Y2 Then Exit Function
108     InsideRectangle = True
End Function
