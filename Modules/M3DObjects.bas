Attribute VB_Name = "M3DObjects"
Option Explicit


' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
'                                           '
'                 z ^                       '
'               2___|____1                  '
'               /|  |   /|                  '
'             3/_|____0/ |                  '
'              | |6____|_|5   --> x         '
'              | /     | /                  '
'              |/______|/                   '
'             7   /    4                    '
'                /                          '
'               y                           '
'                                           '
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Public Function CreateCube(Optional ByVal size_x As Double, Optional ByVal size_y As Double = -1, Optional ByVal size_z As Double = -1) As Object3D
    Dim X As Double: X = IIf(size_x = 0, 0.5, Abs(size_x) / 2)
    Dim Y As Double: Y = IIf(size_y <= 0, X, Abs(size_y) / 2)
    Dim Z As Double: Z = IIf(size_z <= 0, X, Abs(size_z) / 2)
    Dim i As Long
    With CreateCube
        'entweder so:
'        ReDim .points(0 To 7)
'        .points(i) = New_Point3(X, Y, Z):    i = i + 1
'        .points(i) = New_Point3(X, -Y, Z):   i = i + 1
'        .points(i) = New_Point3(-X, -Y, Z):  i = i + 1
'        .points(i) = New_Point3(-X, Y, Z):   i = i + 1
'        .points(i) = New_Point3(X, Y, -Z):   i = i + 1
'        .points(i) = New_Point3(X, -Y, -Z):  i = i + 1
'        .points(i) = New_Point3(-X, -Y, -Z): i = i + 1
'        .points(i) = New_Point3(-X, Y, -Z):  i = 0
'        ReDim .areas(0 To 5)
'        .areas(i) = New_Area(0, 1, 2, 3): i = i + 1
'        .areas(i) = New_Area(4, 5, 6, 7): i = i + 1
'        .areas(i) = New_Area(0, 1, 5, 4): i = i + 1
'        .areas(i) = New_Area(1, 2, 6, 5): i = i + 1
'        .areas(i) = New_Area(2, 3, 7, 6): i = i + 1
'        .areas(i) = New_Area(3, 0, 4, 7): i = i + 1
        'oder so:
        Dim s1: s1 = v(v(X, Y, Z), v(X, -Y, Z), v(-X, -Y, Z), v(-X, Y, Z), _
                       v(X, Y, -Z), v(X, -Y, -Z), v(-X, -Y, -Z), v(-X, Y, -Z))
        Dim s2: s2 = v(v(0, 1, 2, 3), v(4, 5, 6, 7), v(0, 1, 5, 4), _
                       v(1, 2, 6, 5), v(2, 3, 7, 6), v(3, 0, 4, 7))
        .points = ParsePoints(s1)
        .areas = ParseAreas(s2)
    End With
    CreateNormales CreateCube
End Function

Public Function CreateIcosahedron() As Object3D
    '
    Dim p As Point3
    'erzeugt ein Ikosaeder (engl: icosahedron)
    'die X und Z-Koordinaten des angegebenen Punktes werden verwendet
    Dim X As Double: X = Abs(p.X)
    Dim Z As Double: Z = Abs(p.Z)
    If X = 0 And Z = 0 Then X = 0.525731112119134: Z = 0.85065080835204
        
    Dim s1: s1 = v(v(-X, 0#, Z), v(X, 0#, Z), v(-X, 0#, -Z), v(X, 0#, -Z), _
                   v(0#, Z, X), v(0#, Z, -X), v(0#, -Z, X), v(0#, -Z, -X), _
                   v(Z, X, 0#), v(-Z, X, 0#), v(Z, -X, 0#), v(-Z, -X, 0#))
    Dim s2: s2 = v(v(0, 4, 1), v(0, 9, 4), v(9, 5, 4), v(4, 5, 8), v(4, 8, 1), _
                   v(8, 10, 1), v(8, 3, 10), v(5, 3, 8), v(5, 2, 3), v(2, 7, 3), _
                   v(7, 10, 3), v(7, 6, 10), v(7, 11, 6), v(11, 0, 6), v(0, 1, 6), _
                   v(6, 1, 10), v(9, 0, 11), v(9, 11, 2), v(9, 2, 5), v(7, 2, 11))
                   
    Dim icosphere As Object3D
    With icosphere
        .points = ParsePoints(s1)
        .areas = ParseAreas(s2)
        'halt, so hat man zwar die Normalen aber noch nicht die Indizes in den Triangles
        '.normales = CreateNormales(icosphere)
    End With
    CreateNormales icosphere
    CreateIcosahedron = icosphere
End Function

