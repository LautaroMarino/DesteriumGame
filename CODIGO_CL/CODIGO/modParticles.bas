Attribute VB_Name = "mParticle"
' programado por maTih.-
 
Option Explicit

Public Type tImageData
    Bytes() As Byte
End Type

Public ImageData As tImageData
 
Public Const DegreeToRadian As Single = 0.01745329251994 'Pi / 180
Public Const RadianToDegree As Single = 57.2958279087977 '180 / Pi
 
Type Effect_Type
     FX_Grh     As grh      '< FxGrh.
     Fx_Index   As Integer  '< Indice del fx.
     ViajeChar  As Integer  '< CharIndex al que viaja.
     Viaje_X    As Single   '< X hacia donde se dirije.
     End_Effect As Integer  '< FX De la explosión.
     End_Loops  As Integer  '< Loops del fx de la explosión.
     Viaje_Y    As Single   '< Y hacia donde se dirije.
     ViajeSpeed As Single   '< Velocidad de viaje.
     Now_Moved  As Long     '< Tiempo del movimiento actual.
     Last_Move  As Long     '< Tiempo del último movimiento.
     Now_X      As Integer  '< Posición X actual
     Now_Y      As Integer  '< Posición Y actual
     Slot_Used  As Boolean  '< Si está usandose este slot.
End Type
 
Const NO_INDEX = -1         '< Índice no válido.
 
Public Effect() As Effect_Type
 
Public Sub Initialize()
 
'
' @ Inicializa el array de efectos.
 
ReDim Effect(1 To 255) As Effect_Type
 
End Sub
 
Public Sub Terminate_Index(ByVal effect_Index As Integer)
 
'
' @ Destruye un indice del array
 
Dim clear_Index As Effect_Type
 
'Si es un slot válido
If (effect_Index <> 0) And (effect_Index <= UBound(Effect())) Then
    Effect(effect_Index) = clear_Index
End If
 
End Sub
 
Public Function Effect_Begin(ByVal CharIndex As Integer, _
                             ByVal Fx_Index As Integer, _
                             ByVal Bind_Speed As Single, _
                             ByVal X As Single, _
                             ByVal Y As Single, _
                             Optional ByVal explosion_FX_Index As Integer = -1, _
                             Optional ByVal explosion_FX_Loops As Integer = -1) As Integer
 
'
' @ Inicia un nuevo efecto y devuelve el index.
 
Effect_Begin = GetFreeIndex()
 
'Si hay efecto
If (Effect_Begin <> 0) Then
    
    With Effect(Effect_Begin)
         .Now_X = CInt(X)
         .Now_Y = CInt(Y)
         .ViajeChar = CharIndex
         .Fx_Index = Fx_Index
         
         .ViajeSpeed = Bind_Speed
         
         'Explosión?
         If (explosion_FX_Index <> -1) Then
            .End_Effect = explosion_FX_Index
            .End_Loops = explosion_FX_Loops
         End If
         
         'Hay fx
         If (.Fx_Index <> 0) Then
            'Inicializa la animación.
            InitGrh .FX_Grh, FxData(Fx_Index).Animacion
         End If
         
         .Slot_Used = True
         
    End With
    
End If
 
End Function
 
Public Sub Effect_Render_All()
 
'
' @ Dibuja todos los efectos
 
Dim i   As Long
 
For i = 1 To UBound(Effect())
    With Effect(i)
         
         If .Slot_Used Then
            Effect_Render_Slot CInt(i)
         End If
         
    End With
Next i
 
End Sub
 
Public Sub Effect_Render_Slot(ByVal effect_Index As Integer)
 
'
' @ Renderiza un efecto.
 
With Effect(effect_Index)
 
     Dim target_Angle   As Single
     Dim Time As Long
     
     Time = FrameTime
     .Now_Moved = Time
     
     'Controla el intervalo de vuelo
     If (.Last_Move + 10) < .Now_Moved Then
        .Last_Move = Time
        
        'Si tiene char de destino.
        If (.ViajeChar <> 0) Then
     
            'Actualiza la pos de destino.
            .Viaje_X = CharList(.ViajeChar).NowPosX
            .Viaje_Y = CharList(.ViajeChar).NowPosY
 
        End If
        
      End If
      
     'Actualiza el ángulo.
     target_Angle = Engine_GetAngle(.Now_X, .Now_Y, CInt(.Viaje_X), CInt(.Viaje_Y))
    
     'Actualiza la posición del efecto.
     .Now_X = (.Now_X + Sin(target_Angle * DegreeToRadian) * .ViajeSpeed)
     .Now_Y = (.Now_Y - Cos(target_Angle * DegreeToRadian) * .ViajeSpeed)
     
     'Si hay posición dibuja.
     If (.Now_X <> 0) And (.Now_Y <> 0) Then
        Call Draw_Grh(.FX_Grh, .Now_X, .Now_Y, To_Depth(7), 0, 1)
        
        'Check si terminó.
        If (.FX_Grh.started = 0) Then .Fx_Index = 0: .Slot_Used = False
        
        'Llegó a destino?
        If (.Now_X = .Viaje_X) And (.Now_Y = .Viaje_Y) Then
            'Inicializa la explosión : p
            If (.End_Effect <> 0) And (.End_Loops <> 0) Then
                Call SetCharacterFx(.ViajeChar, .End_Effect, .End_Loops)
            End If
           .Slot_Used = False
        End If
        
     End If
End With
 
End Sub
 
Public Function Engine_GetAngle(ByVal CenterX As Integer, ByVal CenterY As Integer, ByVal TargetX As Integer, ByVal TargetY As Integer) As Single
'************************************************************
'Gets the angle between two points in a 2d plane
'More info: [url=http://www.vbgore.com/GameClient.TileEn]http://www.vbgore.com/GameClient.TileEn[/url] ... e_GetAngle" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'************************************************************
Dim SideA As Single
Dim SideC As Single
 
    On Error GoTo ErrOut
 
    'Check for horizontal lines (90 or 270 degrees)
    If CenterY = TargetY Then
 
        'Check for going right (90 degrees)
        If CenterX < TargetX Then
            Engine_GetAngle = 90
 
            'Check for going left (270 degrees)
        Else
            Engine_GetAngle = 270
        End If
 
        'Exit the function
        Exit Function
 
    End If
 
    'Check for horizontal lines (360 or 180 degrees)
    If CenterX = TargetX Then
 
        'Check for going up (360 degrees)
        If CenterY > TargetY Then
            Engine_GetAngle = 360
 
            'Check for going down (180 degrees)
        Else
            Engine_GetAngle = 180
        End If
 
        'Exit the function
        Exit Function
 
    End If
 
    'Calculate Side C
    SideC = Sqr(Abs(TargetX - CenterX) ^ 2 + Abs(TargetY - CenterY) ^ 2)
 
    'Side B = CenterY
 
    'Calculate Side A
    SideA = Sqr(Abs(TargetX - CenterX) ^ 2 + TargetY ^ 2)
 
    'Calculate the angle
    Engine_GetAngle = (SideA ^ 2 - CenterY ^ 2 - SideC ^ 2) / (CenterY * SideC * -2)
    Engine_GetAngle = (Atn(-Engine_GetAngle / Sqr(-Engine_GetAngle * Engine_GetAngle + 1)) + 1.5708) * 57.29583
 
    'If the angle is >180, subtract from 360
    If TargetX < CenterX Then Engine_GetAngle = 360 - Engine_GetAngle
 
    'Exit function
 
Exit Function
 
    'Check for error
ErrOut:
 
    'Return a 0 saying there was an error
    Engine_GetAngle = 0
 
Exit Function
 
End Function
 
Public Function GetFreeIndex() As Integer
 
'
' @ Devuelve un índice para un nuevo FX.
 
Dim i   As Long
 
For i = 1 To UBound(Effect())
    'No está usado.
    If Not Effect(i).Slot_Used Then
       GetFreeIndex = CInt(i)
       Exit Function
    End If
Next i
 
GetFreeIndex = NO_INDEX
 
End Function


Public Function Get_PixelX_Of_Char(ByVal CharIndex As Integer)
    Get_PixelX_Of_Char = CharList(CharIndex).NowPosX
End Function

Public Function Get_PixelY_Of_Char(ByVal CharIndex As Integer)
    Get_PixelY_Of_Char = CharList(CharIndex).NowPosY
End Function
