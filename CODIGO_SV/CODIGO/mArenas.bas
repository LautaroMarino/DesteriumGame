Attribute VB_Name = "mArenas"
Option Explicit

' Cargamos las arenas para disputar los diferentes eventos del juego. Cada una tiene su identificador valido

Public Type tArenas
    Used As Boolean
    
    Map As Integer
    X As Integer
    Y As Integer
    
    TileAddX As Integer
    TileAddY As Integer
    
    MinUsers As Byte
    MaxUsers As Byte
    
    Terreno As Byte
    Plante As Byte
    Tipo As Byte
    
End Type

Public ArenaLast As Integer
Public ArenaLastConfig As Integer
Public Arenas() As tArenas

Public Sub Arenas_Load()
    On Error GoTo ErrHandler
    
    Dim Manager As clsIniManager
    Dim FilePath As String
    Dim A As Long, B As Long, C As Long, D As Long
    Dim Temp As String
    Dim sMap() As String, sX() As String, sY() As String, AddX As Integer, AddY As Integer, MaxUsers As Byte, MinUsers As Byte
    Dim MapLast As Long
    
    Dim Terreno As Byte, Tipo As Byte, Plante As Byte
    
    FilePath = DatPath & "Arenas.ini"
    
    Set Manager = New clsIniManager
    Manager.Initialize (FilePath)
    
    ArenaLastConfig = val(Manager.GetValue("INIT", "LAST"))
    
    MapLast = 0
    
    For A = 1 To ArenaLastConfig
        Temp = Manager.GetValue(A, "Map")
        sMap = Split(Temp, "-")
        
        Temp = Manager.GetValue(A, "X")
        sX = Split(Temp, "-")
        
        Temp = Manager.GetValue(A, "Y")
        sY = Split(Temp, "-")
        
        AddX = val(Manager.GetValue(A, "AddX"))
        AddY = val(Manager.GetValue(A, "AddY"))
        
        MinUsers = val(Manager.GetValue(A, "MinUsers"))
        MaxUsers = val(Manager.GetValue(A, "MaxUsers"))
        
        Terreno = val(Manager.GetValue(A, "Terreno"))
        Tipo = val(Manager.GetValue(A, "Tipo"))
        Plante = val(Manager.GetValue(A, "Plante"))
        
        For B = 0 To UBound(sMap)
            For C = 0 To UBound(sX)
                For D = 0 To UBound(sY)
                    MapLast = MapLast + 1
                    ReDim Preserve Arenas(1 To MapLast)
                    
                    With Arenas(MapLast)
                        .Map = CLng(sMap(B))
                        .X = CLng(sX(C))
                        .Y = CLng(sY(D))
                        .TileAddX = AddX
                        .TileAddY = AddY
                        .MinUsers = MinUsers
                        .MaxUsers = MaxUsers
                        .Terreno = Terreno
                        .Tipo = Tipo
                        .Plante = Plante
                    End With
                Next D
            Next C
        Next B
    Next A


    Set Manager = Nothing
    
    Exit Sub

ErrHandler:
    ' Manejo de errores
    Set Manager = Nothing
End Sub




' # Busca una arena libre para disputar en el evento.
' # Tipos
' 0 = Retos rapidos
' 1 = Retos
'
Public Function Arenas_Free(ByVal Users As Byte, ByVal Tipo As Byte, Optional ByVal Terreno As Byte = 0) As Integer


    On Error GoTo ErrHandler
    
    Dim FreeArenas() As Integer
    Dim FreeCount As Integer
    Dim A As Integer
    Dim RandIndex As Integer
    
    ' Inicializar con 0
    Arenas_Free = 0
    
    
    
    ' Contar la cantidad de arenas libres y almacenar sus índices
    FreeCount = 0
    For A = LBound(Arenas) To UBound(Arenas)
                                   '   2             6                          2           6
        If Not Arenas(A).Used And _
            (Arenas(A).MaxUsers >= Users And Arenas(A).MinUsers <= Users) And _
            ((Arenas(A).Terreno > 0 And Arenas(A).Terreno = Terreno) Or Arenas(A).Terreno = 0) And _
            (Arenas(A).Tipo = Tipo) Then
            
            FreeCount = FreeCount + 1
            ReDim Preserve FreeArenas(1 To FreeCount)
            FreeArenas(FreeCount) = A
            
        End If
    Next A
    
    
    
    ' Si hay al menos una arena libre, generar un número aleatorio para seleccionar una arena libre
    If FreeCount > 0 Then
        RandIndex = RandomNumber(1, FreeCount)
        Arenas_Free = FreeArenas(RandIndex)
    End If
    
    Exit Function
ErrHandler:
    
End Function





