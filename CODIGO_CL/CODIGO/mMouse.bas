Attribute VB_Name = "mMouse"
'------------------------------------------------------------------------------
' Módulo para subclasificación (subclassing)                        (26/Jun/98)
' Revisado (probado) para publicar en mis páginas                   (18/Abr/01)
'
' Modificado para usar con la clase clsMouse                    (21/Mar/99)
'
' ©Guillermo 'guille' Som, 1998-2001
'
' Para más información sobre subclasificación:
' En la documentación de Visual Basic:
'   Pasar punteros de función a los procedimientos de DLL y a las bibliotecas de tipos
' En la MSDN Library (o en la Knowledge Base):
'   HOWTO: Subclass a UserControl
'       Article ID: Q179398
'   HOWTO: Hook Into a Window's Messages Using AddressOf
'       Article ID: Q168795
'   HOWTO: Build a Windows Message Handler with AddressOf in VB5
'       Article ID: Q170570
'------------------------------------------------------------------------------
Option Explicit

' Un array de la clase que se usará para subclasificar ventanas
' y el último elemento de clases en el array; empieza a contar por uno
Private mWSC() As clsMouse          ' Array de clases
Private mnWSC As Long                   ' Número de ventanas subclasificadas

Private Const GWL_WNDPROC = (-4&)

Public Declare Function CallWindowProc Lib "User32" Alias "CallWindowProcA" _
    (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, _
    ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long

Public Function WndProc(ByVal hWnd As Long, ByVal uMSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    ' Los mensajes de windows llegarán aquí.
    ' Lo que hay que hacer es "capturar" los que se necesiten,
    ' en este caso se devuelven los mensajes a la clase, usando para
    ' ello un procedimiento público llamado unMSG con los siguientes parámetros:
    ' ByVal uMSG As Long, ByVal wParam As Long, ByVal lParam As Long
    '
    ' Para un mejor uso, usar la clase en el formato:
    '   Dim WithEvents laClase As clsMouse
    '
    Static i As Long
    
    ' Buscar el índice de esta clase en el array
    ' NOTA: Esto se hará para cada uno de los mensajes recibidos,
    '       por tanto, no sería conveniente tener demasiadas ventanas o controles
    '       subclasificados, con idea de que no tarde demasiado en procesarlos
    '$Por hacer:
    '   Sería conveniente poner un límite máximo de ventanas a subclasificar
    i = IndiceClase(hWnd)
    
    If i Then
        With mWSC(i)
            WndProc = CallWindowProc(.PrevWndProc, hWnd, uMSG, wParam, lParam)
            ' Producir el evento del mensaje recibido
            .unMSG uMSG, wParam, lParam
        End With
    End If
End Function

Public Sub Hook(ByVal WSC As clsMouse, ByVal unControl As Object)
    ' Subclasificar la ventana o control indicado
    
    '--------------------------------------------------------------------------
    ' Nota:
    ' En este procedimiento no se hace chequeo de que el objeto pasado tenga
    ' la propiedad hWnd, ya que se comprueba en el método Hook de la clase,
    ' por tanto no se debería llamar a este método sin antes hacer una comprobación
    ' de que estamos pasando un objeto-ventana (que tenga la propiedad hWnd)
    '
    
    ' Comprobar si ya está subclasificada esta ventana
    Dim claseActual As Long
    Dim claseLibre As Long
    
    ' Buscar el índice de esta clase en el array
    ' y si hay alguna clase liberada anteriormente
    claseActual = IndiceClase(unControl.hWnd, claseLibre)
    
    If claseActual = 0 Then
        ' Si hay un índice que ya no se usa...
        If claseLibre Then
            ' se usará ese índice
            claseActual = claseLibre
        Else
            ' Crear una nueva clase
            mnWSC = mnWSC + 1
            ReDim Preserve mWSC(1 To mnWSC)
            claseActual = mnWSC
        End If
    End If
    
    ' Aquí se está haciendo referencia a una clase ya existente,
    ' para que no queden referencias "sueltas", en el evento Terminate de la clase
    ' se llama al procedimiento de liberación de la subclasificación en el que se
    ' borrará la referencia a la clase indicada, por tanto no se debería modificar
    ' esa forma de actuar.
    '--------------------------------------------------------------------------
    ' Nota:
    '   En lugar de hacer una referencia a la clase, se podría usar un puntero a
    '   la misma usando ObjPtr, pero esto implicaría usar la función CopyMemory
    '   para poder acceder a las propiedades de la clase, y no sé si esto
    '   incrementaría el tiempo de procesamiento, pero...
    '   "los expertos" así lo hacen... así que se supone que tendrá sus ventajas;
    '   aunque si se siguen "las reglas" indicadas, no tendría que dar problemas.
    '   Además la intención de esta clase es formar parte de un componente (DLL)
    '   y el código no estaría disponible a las aplicaciones cliente...
    '   Por eso, te aconsejo que no hagas experimentos,
    '   si no sabes las consecuencias que esa pruebas pueden tener, el que avisa...
    '
    ' Ver el siguiente artículo en la Knowledge Base de Microsoft para un ejemplo
    ' de un UserControl subclasificado usando punteros a objetos:
    '   HOWTO: Subclass a UserControl, Article ID: Q179398
    '
    Set mWSC(claseActual) = WSC
    
    ' Subclasificar la ventana, (form o control), pasada como parámetro y
    ' guardar el procedimiento anterior
    With mWSC(claseActual)
        .hWnd = unControl.hWnd
        .PrevWndProc = SetWindowLong(.hWnd, GWL_WNDPROC, AddressOf WndProc)
    End With
End Sub

Public Sub unHook(ByVal WSC As clsMouse)
    ' Des-subclasificar la clase indicada
    Static claseActual As Long
    
    ' Buscar el índice de esta clase en el array
    claseActual = IndiceClase(WSC.hWnd)
    
    ' Si ya estaba subclasificada esta clase
    If claseActual Then
        With mWSC(claseActual)
            ' Restaurar la función anterior de procesamiento de mensajes
            Call SetWindowLong(.hWnd, GWL_WNDPROC, .PrevWndProc)
            ' Poner a cero el indicador de que se está usando
            .hWnd = 0&
        End With
        
        ' Quitar la referencia a esta clase
        Set mWSC(claseActual) = Nothing
        
        ' Si es la última del array...
        If mnWSC = claseActual Then
            ' Eliminar este item y ajustar el array
            mnWSC = mnWSC - 1
            ' Si no hay más, eliminar el array
            If mnWSC = 0 Then
                Erase mWSC
            Else
                ' Ajustar el número de elementos del array
                ReDim Preserve mWSC(1 To mnWSC)
            End If
        End If
    End If
End Sub

Private Function IndiceClase(ByVal elhWnd As Long, Optional ByRef Libre As Long = 0) As Long
    ' Este procedimiento buscará el índice de la clase que tiene el hWnd indicado
    ' También, si se especifica, devolverá el índice de una clase que esté libre.
    ' Nota: Es importante que el último parámetro sea por referencia,
    '       ya que en él se devolverá el valor del índice libre.
    '
    Static i As Long
    
    IndiceClase = 0
    
    ' Recorrer todo el array
    For i = 1 To mnWSC
        With mWSC(i)
            ' Si coinciden los hWnd, es que ya se está usando una subclasificación
            If .hWnd = elhWnd Then
                ' usar esta misma clase
                ' pero si el hWnd es cero, será uno libre
                If elhWnd = 0 Then
                    Libre = i
                Else
                    IndiceClase = i
                End If
                Exit For
                
            ' Comprobar si hay algún "hueco" en el array,
            ' por ejemplo de una clase previamente liberada.
            ' Hay que tener en cuenta que estos procedimientos están en un BAS
            ' y sus valores se mantienen entre varias llamadas a las clases.
            ElseIf .hWnd = 0& Then
                Libre = i
            End If
        End With
    Next
End Function

