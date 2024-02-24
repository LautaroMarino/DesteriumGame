VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCola"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Argentum Online 0.12.2
'Copyright (C) 2002 Márquez Pablo Ignacio
'
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
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit
'                    Metodos publicos
'
' Public sub Push(byval i as variant) mete el elemento i
' al final de la cola.
'
' Public Function Pop As Variant: quita de la cola el primer elem
' y lo devuelve
'
' Public Function VerElemento(ByVal Index As Integer) As Variant
' muestra el elemento numero Index de la cola sin quitarlo
'
' Public Function PopByVal() As Variant: muestra el primer
' elemento de la cola sin quitarlo
'
' Public Property Get Longitud() As Integer: devuelve la
' cantidad de elementos que tiene la cola.

Private Const FRENTE = 1

Private MAXELEM As Integer

Private Cola    As Collection

Public Property Let MaxLenght(ByVal max As Integer)
    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/11/2010
    'Sets the max queue lenght.
    '***************************************************
    MAXELEM = max
End Property

Public Sub Reset()
        '<EhHeader>
        On Error GoTo Reset_Err
        '</EhHeader>

        Dim i As Integer

100     For i = 1 To Me.Longitud
102         Cola.Remove FRENTE
104     Next i

        '<EhFooter>
        Exit Sub

Reset_Err:
        Err.Raise vbObjectError + 100, _
                  "ServidorArgentum.cCola.Reset", _
                  "cCola component failure"
        '</EhFooter>
End Sub

Public Property Get Longitud() As Integer
    Longitud = Cola.Count
End Property

Private Function IndexValido(ByVal i As Integer) As Boolean
    IndexValido = i >= 1 And i <= Me.Longitud
End Function

Private Sub Class_Initialize()
    Set Cola = New Collection
End Sub

Public Function VerElemento(ByVal Index As Integer, _
                            Optional ByVal ConvertUpper As Boolean = True) As String
        '<EhHeader>
        On Error GoTo VerElemento_Err
        '</EhHeader>

100     If IndexValido(Index) Then

            'Pablo
102         If ConvertUpper Then
104             VerElemento = UCase$(Cola.Item(Index))
            Else
106             VerElemento = Cola.Item(Index)
            End If
        
            '/Pablo
            'VerElemento = Cola(Index)
        Else
108         VerElemento = 0
        End If

        '<EhFooter>
        Exit Function

VerElemento_Err:
        Err.Raise vbObjectError + 100, _
                  "ServidorArgentum.cCola.VerElemento", _
                  "cCola component failure"
        '</EhFooter>
End Function

Public Sub Push(ByVal Detalle As String, Optional ByVal ConvertUpper As Boolean = True)
        '<EhHeader>
        On Error GoTo Push_Err
        '</EhHeader>

        '***************************************************
        'Author: Unknown
        'Last Modification: 14/11/2010
        '14/11/2010: ZaMa - Ahora se valida si supera un maximo previamente establecido.
        '***************************************************

        'Mete elemento en la cola
    
        Dim aux As String

100     aux = Time$ & " "
    
102     If ConvertUpper Then
104         aux = aux & UCase$(Detalle)
        Else
106         aux = aux & Detalle
        End If
    
        ' Validate that not exceed max lenght
108     If MAXELEM <> 0 Then
110         If Cola.Count = MAXELEM Then
112             Call Cola.Remove(FRENTE) ' Quito el primer elemento
            End If
        End If
    
114     Call Cola.Add(aux)
        
        '<EhFooter>
        Exit Sub

Push_Err:
        Err.Raise vbObjectError + 100, _
                  "ServidorArgentum.cCola.Push", _
                  "cCola component failure"
        '</EhFooter>
End Sub

Public Function Pop() As String
        '<EhHeader>
        On Error GoTo Pop_Err
        '</EhHeader>

        'Quita elemento de la cola
100     If Cola.Count > 0 Then
102         Pop = Cola(FRENTE)
104         Call Cola.Remove(FRENTE)
        Else
106         Pop = 0
        End If

        '<EhFooter>
        Exit Function

Pop_Err:
        Err.Raise vbObjectError + 100, _
                  "ServidorArgentum.cCola.Pop", _
                  "cCola component failure"
        '</EhFooter>
End Function

Public Function PopByVal() As String
        '<EhHeader>
        On Error GoTo PopByVal_Err
        '</EhHeader>

        'Call LogTarea("PopByVal SOS")

        'Quita elemento de la cola
100     If Cola.Count > 0 Then
102         PopByVal = Cola.Item(1)
        Else
104         PopByVal = 0
        End If

        '<EhFooter>
        Exit Function

PopByVal_Err:
        Err.Raise vbObjectError + 100, _
                  "ServidorArgentum.cCola.PopByVal", _
                  "cCola component failure"
        '</EhFooter>
End Function

Public Function Existe(ByVal Nombre As String) As Boolean
        '<EhHeader>
        On Error GoTo Existe_Err
        '</EhHeader>
  
        Dim V                  As String

        Dim i                  As Integer

        Dim NombreEnMayusculas As String

100     NombreEnMayusculas = UCase$(Nombre)
 
102     For i = 1 To Me.Longitud
            'Pablo
104         V = mid$(Me.VerElemento(i), 10, Len(Me.VerElemento(i)))

            '/Pablo
            'V = Me.VerElemento(i)
106         If V = NombreEnMayusculas Then
108             Existe = True

                Exit Function

            End If

        Next

110     Existe = False

        '<EhFooter>
        Exit Function

Existe_Err:
        Err.Raise vbObjectError + 100, _
                  "ServidorArgentum.cCola.Existe", _
                  "cCola component failure"
        '</EhFooter>
End Function

Public Sub Quitar(ByVal Nombre As String)
        '<EhHeader>
        On Error GoTo Quitar_Err
        '</EhHeader>

        Dim V                  As String

        Dim i                  As Integer

        Dim NombreEnMayusculas As String
 
100     NombreEnMayusculas = UCase$(Nombre)
 
102     For i = 1 To Me.Longitud
            'Pablo
104         V = mid$(Me.VerElemento(i), 10, Len(Me.VerElemento(i)))

            '/Pablo
            'V = Me.VerElemento(i)
106         If V = NombreEnMayusculas Then
108             Call Cola.Remove(i)

                Exit Sub

            End If

110     Next i
 
        '<EhFooter>
        Exit Sub

Quitar_Err:
        Err.Raise vbObjectError + 100, _
                  "ServidorArgentum.cCola.Quitar", _
                  "cCola component failure"
        '</EhFooter>
End Sub

Public Sub QuitarIndex(ByVal Index As Integer)
        '<EhHeader>
        On Error GoTo QuitarIndex_Err
        '</EhHeader>

100     If IndexValido(Index) Then Call Cola.Remove(Index)
        '<EhFooter>
        Exit Sub

QuitarIndex_Err:
        Err.Raise vbObjectError + 100, _
                  "ServidorArgentum.cCola.QuitarIndex", _
                  "cCola component failure"
        '</EhFooter>
End Sub

Private Sub Class_Terminate()
    'Destruimos el objeto Cola
    Set Cola = Nothing
End Sub
