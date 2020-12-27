Attribute VB_Name = "M3D"
Option Explicit

'Public Type Point 3' siehe MMatrix

Public Type Area
    iP1 As Long 'index zum 1. Punkt der Fläche im Array aus Point3
    iP2 As Long 'index zum 2. Punkt der Fläche im Array aus Point3
    iP3 As Long 'index zum 3. Punkt der Fläche im Array aus Point3
    iP4 As Long 'index zum 4. Punkt der Fläche im Array aus Point3 'falls -1 dann ist es ein Dreieck, sonst ist es ein ebenes Viereck
    iNorm As Long 'index zur Flächennormalen
End Type
Public Type Object3D
    points()     As Point3   'die Punkte
    areas()      As Area     'Indices zu den Punkten
    normales()   As Point3   'die Flächennormalen der Areas
    Projection() As Point2L  'die Punkte in projizierten Bildkoordinaten
End Type

Public Pi
Public Pi2
Public Const eps  As Double = 0.000000000001
Private HashTable As Collection
Public sc As Double 'für DrawPoint3_xy

Public Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As Point2L, ByVal nCount As Long) As Long
        
Public Declare Function Polyline Lib "gdi32" (ByVal hDC As Long, lpPoint As Point2L, ByVal nCount As Long) As Long

' c'tors
Public Function New_Area(ByVal i1 As Long, ByVal i2 As Long, ByVal i3 As Long, Optional ByVal i4 As Long = -1) As Area
    With New_Area:        .iP1 = i1:        .iP2 = i2:        .iP3 = i3:        .iP4 = i4:    End With
End Function

Public Function New_Area_va(v) As Area

    If Not IsArray(v) Then Exit Function
    Dim n As Long: n = UBound(v) - LBound(v) + 1
    
    With New_Area_va
        .iP1 = CLng(v(0)): If n < 2 Then Exit Function
        .iP2 = CLng(v(1)): If n < 3 Then Exit Function
        .iP3 = CLng(v(2)): If n < 4 Then Exit Function
        .iP4 = CLng(v(3))
    End With
    
End Function

Sub CreateNormales(obj As Object3D)
    With obj
        Dim i As Long: i = LBound(.areas)
        Dim u As Long: u = UBound(.areas)
        ReDim .normales(i To u) As Point3
        Dim a As Area
        For i = i To u
            .areas(i).iNorm = i
            a = .areas(i)
            .normales(i) = Normalize(.points(a.iP1), .points(a.iP2), .points(a.iP3))
        Next
    End With
End Sub

Public Function Normale(p As Point3) As Point3
    Dim d As Double: d = VBA.Sqr(p.X * p.X + p.Y * p.Y + p.Z * p.Z)
    If d = 0 Then Exit Function
    With Normale
        .X = p.X / d
        .Y = p.Y / d
        .Z = p.Z / d
    End With
End Function
Public Function Normalize(p1 As Point3, p2 As Point3, p3 As Point3) As Point3
    Normalize = NormCrossProd(Point3_Subtract(p1, p2), Point3_Subtract(p2, p3))
End Function
Public Function NormCrossProd(v1 As Point3, v2 As Point3) As Point3
    With NormCrossProd
        .X = v1.Y * v2.Z - v1.Z * v2.Y
        .Y = v1.Z * v2.X - v1.X * v2.Z
        .Z = v1.X * v2.Y - v1.Y * v2.X
    End With
    NormCrossProd = Normale(NormCrossProd)
End Function

'Hilfsfunktion
Public Function v(ParamArray p())
    v = p
End Function

'Parsers
Public Function ParsePoints(VArr) As Point3()
    Dim i As Long
    ReDim p(LBound(VArr) To UBound(VArr)) As Point3
    For i = LBound(VArr) To UBound(VArr)
        p(i) = New_Point3_va(VArr(i))
    Next
    ParsePoints = p
End Function
Public Function ParseAreas(VArr) As Area()
    Dim i As Long
    ReDim a(LBound(VArr) To UBound(VArr)) As Area
    For i = LBound(VArr) To UBound(VArr)
        a(i) = New_Area_va(VArr(i))
    Next
    ParseAreas = a
End Function

'Zeichenroutinen
Public Function DrawObj3D_xy(aPB As PictureBox, obj As Object3D, color As Long)

    Dim i As Long
    Dim a As Area
    Dim p1 As Point3, p2 As Point3, p3 As Point3, p4 As Point3
    'Dim sc As Double: sc = 37 * 5
    If sc = 0 Then sc = 37
    Dim tx As Double: tx = aPB.ScaleWidth / 2
    Dim ty As Double: ty = aPB.ScaleHeight / 2
    Dim c As Long
    aPB.ForeColor = color
    With obj
    
        For i = LBound(.areas) To UBound(.areas)
            a = .areas(i)
            p1 = .points(a.iP1)
            p2 = .points(a.iP2)
            p3 = .points(a.iP3)
            If a.iP4 >= 0 Then
                p4 = .points(a.iP4)
            End If
            aPB.Line (tx + p1.X * sc, ty + p1.Y * sc)-(tx + p2.X * sc, ty + p2.Y * sc)
            aPB.Line -(tx + p3.X * sc, ty + p3.Y * sc)
            If a.iP4 >= 0 Then
                aPB.Line -(tx + p4.X * sc, ty + p4.Y * sc)
            End If
            aPB.Line -(tx + p1.X * sc, ty + p1.Y * sc)
            'c = c + 3
        Next
        
    End With
    'Debug.Print c / 2
End Function

Public Function DrawPoint3_xy(aPB As PictureBox, p As Point3, color As Long)
    If sc = 0 Then sc = 37
    Dim tx As Double: tx = aPB.ScaleWidth / 2
    Dim ty As Double: ty = aPB.ScaleHeight / 2
    aPB.ForeColor = color
    aPB.Circle (tx + p.X * sc, ty + p.Y * sc), 3
End Function

Public Function DrawObj3D_xz(aPB As PictureBox, obj As Object3D, color As Long)

    Dim i As Long
    Dim a As Area
    Dim p1 As Point3, p2 As Point3, p3 As Point3, p4 As Point3
    'Dim sc As Double: sc = 37 * 5
    If sc = 0 Then sc = 37
    Dim tx As Double: tx = aPB.ScaleWidth / 2
    Dim tz As Double: tz = aPB.ScaleHeight / 2
    aPB.ForeColor = color
    With obj
    
        For i = LBound(.areas) To UBound(.areas)
            a = .areas(i)
            p1 = .points(a.iP1)
            p2 = .points(a.iP2)
            p3 = .points(a.iP3)
            If a.iP4 >= 0 Then
                p4 = .points(a.iP4)
            End If
            aPB.Line (tx + p1.X * sc, tz + -p1.Z * sc)-(tx + p2.X * sc, tz + -p2.Z * sc)
            aPB.Line -(tx + p3.X * sc, tz + -p3.Z * sc)
            If a.iP4 >= 0 Then
                aPB.Line -(tx + p4.X * sc, tz + -p4.Z * sc)
            End If
            aPB.Line -(tx + p1.X * sc, tz + -p1.Z * sc)
        Next
        
    End With
    
End Function

Public Function DrawPoint3_xz(aPB As PictureBox, p As Point3, color As Long)
    If sc = 0 Then sc = 37
    Dim tx As Double: tx = aPB.ScaleWidth / 2
    Dim tz As Double: tz = aPB.ScaleHeight / 2
    aPB.ForeColor = color
    aPB.Circle (tx + p.X * sc, tz + p.Z * sc), 3
End Function

Public Sub DrawObj3D_projected(aPB As PictureBox, aProj As Matrix34, obj As Object3D, color As Long)
    'Debug.Print "Drawing with Polyline"
    Dim a As Area
    Dim n As Long: n = (UBound(obj.areas) - LBound(obj.areas) + 1) * 4
    ReDim p(0 To n - 1) As Point2L
    ReDim aa(0 To n - 1) As Long
    Dim i As Long, c As Long
    For i = LBound(obj.areas) To UBound(obj.areas)
        a = obj.areas(i)
        p(c) = Point3_Projection(obj.points(a.iP1), aProj): c = c + 1
        p(c) = Point3_Projection(obj.points(a.iP2), aProj): c = c + 1
        aa(c) = 3
        p(c) = Point3_Projection(obj.points(a.iP3), aProj): c = c + 1
        If a.iP4 >= 0 Then
            aa(c) = 4
            p(c) = Point3_Projection(obj.points(a.iP4), aProj)
        End If
        c = c + 1
    Next
    Dim rv As Long
    Dim hDC As Long: hDC = aPB.hDC
    'hmm, ja wie löst man das Problem?
    '3 oder 4 zeichnen lassen?
    For i = 0 To UBound(p) Step 4
        rv = Polyline(hDC, p(i), aa(i + 3))
    Next
    ''wie schauts aus mit den Flächennormalen?
End Sub
Public Function DrawPoint3_projected(aPB As PictureBox, aProj As Matrix34, pt As Point3, color As Long)
    If sc = 0 Then sc = 37
    Dim tx As Double: tx = aPB.ScaleWidth / 2
    Dim ty As Double: ty = aPB.ScaleHeight / 2
    aPB.ForeColor = color
    Dim p As Point2L: p = Point3_Projection(pt, aProj)
    aPB.Circle (p.X, p.Y), 3
End Function

Public Sub ZoomIn(Optional ByVal zoomfact As Double = 1.5)
    sc = sc * zoomfact
End Sub
Public Sub ZoomOut(Optional ByVal zoomfact As Double = 1.5)
    sc = sc / zoomfact
End Sub

Public Function Distance(p As Point3) As Double
    Distance = Sqr(p.X * p.X + p.Y * p.Y + p.Z * p.Z)
End Function

Public Function Rotate_xy(points() As Point3, ByVal x0 As Double, ByVal y0 As Double, ByVal alpha_rad As Double) As Point3()
    Dim i As Long
    ReDim p(LBound(points) To UBound(points)) As Point3
    Dim sin_a As Double: sin_a = Sin(alpha_rad)
    Dim cos_a As Double: cos_a = Cos(alpha_rad)
    Dim dx As Double, dy As Double
    For i = LBound(p) To UBound(p)
        dx = points(i).X - x0
        dy = points(i).Y - y0
        p(i).X = x0 + dx * cos_a - dy * sin_a
        p(i).Y = y0 + dx * sin_a + dy * cos_a
        p(i).Z = points(i).Z
    Next
    Rotate_xy = p
End Function

