'                          _
'   __ _ ___ ___ _ __  ___| |_ _ _ _  _ 
'  / _` / -_) _ \ '  \/ -_)  _| '_| || |
'  \__, \___\___/_|_|_\___|\__|_|  \_, |
'  |___/                           |__/ 
' 
' Geometry calculation class
'
'Imports Microsoft.VisualBasic
Imports System.Globalization
Imports System.Text.RegularExpressions
Imports WallnerMild.Core

Namespace Geom

    Public Class Point
        Private p_x As Double
        Private p_y As Double

        Public Sub New()
            p_x = 0
            p_y = 0
        End Sub
        Public Sub New(x As Double, y As Double)
            p_x = x
            p_y = y
        End Sub

        ''' <summary>
        ''' x value
        ''' </summary>
        ''' <returns></returns>
        Public Property x As Double
            Get
                Return p_x
            End Get
            Set(value As Double)
                p_x = value
            End Set
        End Property
        ''' <summary>
        ''' y value
        ''' </summary>
        ''' <returns></returns>
        Public Property y As Double
            Get
                Return p_y
            End Get
            Set(value As Double)
                p_y = value
            End Set
        End Property
        ''' <summary>
        ''' convert to vector
        ''' </summary>
        ''' <returns></returns>
        Public Function toVector() As Vector
            Return New Vector(x, y)
        End Function
        ''' <summary>
        ''' Rotate point about a point
        ''' </summary>
        ''' <param name="center"></param>
        ''' <param name="alphaRad"></param>
        ''' <returns></returns>
        Public Function rotate(center As Point, alphaRad As Double) As Point
            Dim p1 As New Point
            Dim p2 As New Point

            ' Translate 
            p1.x = Me.x - center.x
            p1.y = Me.y - center.y
            ' Rotate
            p2.x = Math.Cos(alphaRad) * p1.x - Math.Sin(alphaRad) * p1.y
            p2.y = Math.Sin(alphaRad) * p1.x + Math.Cos(alphaRad) * p1.y
            ' Translate back
            p2.x = p2.x - center.x
            p2.y = p2.y - center.y

            Return p2

        End Function
    End Class

    ''' <summary>
    ''' Vector class for analysis
    ''' </summary>
    Public Class Vector
        Private p_x As Double
        Private p_y As Double

        Public Sub New()
            p_x = 0
            p_y = 0
        End Sub
        Public Sub New(x As Double, y As Double)
            p_x = x
            p_y = y
        End Sub
        Public Sub New(P1 As Point, P2 As Point)
            p_x = P2.x - P1.x
            p_y = P2.y - P1.y
        End Sub
        ''' <summary>
        ''' x value
        ''' </summary>
        ''' <returns></returns>
        Public Property x As Double
            Get
                Return p_x
            End Get
            Set(value As Double)
                p_x = value
            End Set
        End Property
        ''' <summary>
        ''' y value
        ''' </summary>
        ''' <returns></returns>
        Public Property y As Double
            Get
                Return p_y
            End Get
            Set(value As Double)
                p_y = value
            End Set
        End Property

        ''' <summary>
        ''' vector normal
        ''' </summary>
        ''' <param name="uniform"></param>
        ''' <returns></returns>
        Public Function normal(Optional uniform As Boolean = True) As Vector
            Dim n As New Vector
            n.x = -y
            n.y = x
            If uniform Then
                n = n.divide(length)
            End If
            Return n
        End Function

        ''' <summary>
        ''' add a vector
        ''' </summary>
        ''' <param name="vect"></param>
        Public Function add(vect As Vector) As Vector
            Dim v As New Vector
            v.x = Me.x + vect.x
            v.y = Me.y + vect.y
            Return v
        End Function

        ''' <summary>
        ''' add a vector
        ''' </summary>
        ''' <param name="vect"></param>
        Public Function subtract(vect As Vector) As Vector
            Dim v As New Vector
            v.x = Me.x - vect.x
            v.y = Me.y - vect.y

            Return v
        End Function

        ''' <summary>
        ''' divide vector by value
        ''' </summary>
        ''' <param name="value"></param>
        Public Function divide(value As Double) As Vector
            Dim v As New Vector
            v.x = Me.x / value
            v.y = Me.y / value
            Return v
        End Function

        ''' <summary>
        ''' multiply vector by value
        ''' </summary>
        ''' <param name="value"></param>
        Public Function multiply(value As Double) As Vector
            Dim v As New Vector
            v.x = Me.x * value
            v.y = Me.y * value
            Return v
        End Function

        ''' <summary>
        ''' Return Unit Vector
        ''' </summary>
        ''' <returns></returns>
        Public Function unitVector() As Vector
            Dim v As New Vector
            v.x = Me.x
            v.y = Me.y
            If v.length <> 0 Then
                v = v.divide(v.length)
            End If
            Return v
        End Function
        ''' <summary>
        ''' vector length
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property length() As Double
            Get
                Return Math.Sqrt(x ^ 2 + y ^ 2)
            End Get
        End Property

        ''' <summary>
        ''' Angle in radians
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property angle() As Double
            Get
                Return Math.Atan2(y, x) '* 180 / Math.PI
            End Get
        End Property
        ''' <summary>
        ''' rotate
        ''' </summary>
        ''' <param name="alphaRad">angle in radians</param>
        ''' <returns></returns>
        Public Function rotate(alphaRad As Double) As Vector
            Dim v As New Vector
            v.x = Math.Cos(alphaRad) * Me.x - Math.Sin(alphaRad) * Me.y
            v.y = Math.Sin(alphaRad) * Me.x + Math.Cos(alphaRad) * Me.y
            Return v

        End Function
        ''' <summary>
        ''' Distance to second point
        ''' </summary>
        ''' <param name="p2"></param>
        ''' <returns></returns>
        Public Function distance(p2 As Vector) As Double
            Dim v As New Vector
            v.x = Me.x
            v.y = Me.y
            v.subtract(p2)
            Return v.length
        End Function
        ''' <summary>
        ''' Return Point
        ''' </summary>
        ''' <returns></returns>
        Public Function toPoint() As Point
            Return New Point(Me.x, Me.y)
        End Function
    End Class
    Public Class rectangle
        Private p_p1 As New Point
        Private p_p2 As New Point
        Private p_isEmpty As Boolean

        Public Sub New()
            p_isEmpty = True
        End Sub

        Public Sub New(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double)
            p_p1.x = X1
            p_p1.y = Y1
            p_p2.x = X2
            p_p2.y = Y2
            p_isEmpty = False
        End Sub
        ''' <summary>
        ''' first point
        ''' </summary>
        ''' <returns></returns>
        Public Property P1() As Point
            Get
                Return p_p1
            End Get
            Set(value As Point)
                p_p1 = value
                If p_isEmpty Then p_p2 = value
                p_isEmpty = False
            End Set
        End Property
        ''' <summary>
        ''' second point
        ''' </summary>
        ''' <returns></returns>
        Public Property P2() As Point
            Get
                Return p_p2
            End Get
            Set(value As Point)
                p_p2 = value
                If p_isEmpty Then p_p1 = value
                p_isEmpty = False
            End Set
        End Property
        ''' <summary>
        ''' minimum x-Coordinate
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property xMin() As Double
            Get
                xMin = p1.x
                If P2.x < xMin Then xMin = P2.x
            End Get
        End Property
        ''' <summary>
        ''' minimum y-Coordinate
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property yMin() As Double
            Get
                yMin = p1.y
                If P2.y < yMin Then yMin = P2.y
            End Get
        End Property

        ''' <summary>
        ''' maximum x-Coordinate
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property xMax() As Double
            Get
                xMax = p1.x
                If P2.x > xMax Then xMax = P2.x
            End Get
        End Property
        ''' <summary>
        ''' maximum y-Coordinate
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property yMax() As Double
            Get
                yMax = p1.y
                If P2.y > yMax Then yMax = P2.y
            End Get
        End Property
        ''' <summary>
        ''' expand rectangle to include given point
        ''' </summary>
        ''' <param name="p"></param>
        Public Sub includePoint(ByVal p As Point)
            If p_isEmpty Then
                P1.x = p.x
                P1.y = p.y

                P2.x = p.x
                P2.y = p.y
                p_isEmpty = False
            Else
                If P1.x < P2.x Then
                    If p.x < P1.x Then P1.x = p.x
                Else
                    If p.x < P2.x Then P2.x = p.x
                End If

                If P1.x > P2.x Then
                    If p.x > P1.x Then P1.x = p.x
                Else
                    If p.x > P2.x Then P2.x = p.x
                End If

                If P1.y < P2.y Then
                    If p.y < P1.y Then P1.y = p.y
                Else
                    If p.y < P2.y Then P2.y = p.y
                End If

                If P1.y > P2.y Then
                    If p.y > P1.y Then P1.y = p.y
                Else
                    If p.y > P2.y Then P2.y = p.y
                End If
            End If
        End Sub

        ''' <summary>
        ''' Min Point
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property pMin() As Point
            Get
                Return New Point(xMin, yMin)
            End Get
        End Property
        ''' <summary>
        ''' Max Point
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property pMax() As Point
            Get
                Return New Point(xMax, yMax)
            End Get
        End Property

    End Class
    Public Class polygon
        Private p_points As New WMDictionary(Of Integer, Point)
        Private p_closed As Boolean

        Public Sub New()
            p_closed = True
        End Sub
        ''' <summary>
        ''' get or set point list by a string. Syntax is: [number] [x] [y]; [number] [x] [y]; ...
        ''' </summary>
        ''' <returns></returns>
        ''' <example> 1 1.0 2.0; 2 1.0 5.0; 3 3.0.5.0; 4 3.0 0.0</example>
        Public Property definitionString() As String
            Get
                definitionString = ""
                For Each kvp In p_points
                    definitionString &= String.Format(CultureInfo.InvariantCulture, "{0,3} {1} {2}; ", kvp.Key, kvp.Value.x, kvp.Value.y)
                Next
            End Get
            Set(value As String)
                Dim p_newPoints As New WMDictionary(Of Integer, Point)
                Dim pos As Integer

                Try
                    For Each pItem In Split(value, ";")
                        pos += 1
                        Dim kxy() As String
                        pItem = Regex.Replace(Trim(pItem), "\s{2,}", " ")

                        kxy = Split(Trim(pItem), " ")

                        Select Case kxy.Count
                            Case 2
                                Dim i As Integer
                                i = points.ElementAt(points.Count - 1).Key + 1
                                Dim x As Double
                                Dim y As Double
                                x = 0
                                x = Double.Parse(kxy(0), CultureInfo.InvariantCulture)
                                y = 0
                                y = Double.Parse(kxy(1), CultureInfo.InvariantCulture)
                                p_newPoints.Add(i, New Point(x, y))
                            Case 3
                                Dim i As Integer
                                i = Integer.Parse(kxy(0), CultureInfo.InvariantCulture)
                                Dim x As Double
                                Dim y As Double
                                x = 0
                                x = Double.Parse(kxy(1), CultureInfo.InvariantCulture)
                                y = 0
                                y = Double.Parse(kxy(2), CultureInfo.InvariantCulture)
                                p_newPoints.Add(i, New Point(x, y))
                            Case Else
                                ' ignored
                        End Select
                    Next
                Catch ex As Exception
                    Throw New FormatException(String.Format("Malformatted definition string at ; position {0}", pos), ex)
                End Try

            End Set
        End Property



        ''' <summary>
        ''' Dictionary of points
        ''' </summary>
        ''' <returns></returns>
        Public Property points() As WMDictionary(Of Integer, Point)
            Get
                Return p_points
            End Get
            Set(value As WMDictionary(Of Integer, Point))
                p_points = value
            End Set
        End Property
        ''' <summary>
        ''' Get Point at given Index
        ''' </summary>
        ''' <param name="index"></param>
        ''' <returns></returns>
        ''' <exception cref="System.ArgumentOutOfRangeException">Thrown when <paramref name="index"/> is out of range</exception>
        Public ReadOnly Property pointAt(index As Integer) As Point
            Get
                If index >= 0 And index < points.Count Then
                    pointAt = points.ElementAt(index).Value
                Else
                    Throw New ArgumentOutOfRangeException("index", index, String.Format("Given index {0} is out of range [{1},{2}].", index, 0, points.Count - 1))
                End If
            End Get
        End Property


        ''' <summary>
        ''' boolean value for closed state
        ''' </summary>
        ''' <returns></returns>
        Public Property closed As Boolean
            Get
                Return p_closed
            End Get
            Set(value As Boolean)
                p_closed = value
            End Set
        End Property

        ''' <summary>
        ''' test if point is in polygon (if polygon is not closed, it only includes points if it forms one or more loops by itself)
        ''' </summary>
        ''' <param name="testPoint">point to test for inclusion</param>
        ''' <returns></returns>
        Public Function pointInPolygon(testPoint As Point) As Boolean
            Dim i As Integer           ' start edge point index
            Dim k As Integer           ' end edge point index
            Dim cn As Integer = 0      ' crossing counter 

            Dim P_i As Point
            Dim P_k As Point

            ' loop throug all edges of the polygon
            For j = 0 To points.Count - 1
                i = j
                If j = points.Count - 1 Then

                    If closed Then
                        'Last edge connects last node with first node
                        k = 0
                    Else
                        ' there is no last edge
                        k = j
                    End If
                Else
                    k = j + 1
                End If

                ' point i and j as polygon edge

                P_i = points(points.ElementAt(i).Key)
                P_k = points(points.ElementAt(k).Key)

                ' check if e crosses upward ala Rule #1
                If (((P_i.y <= testPoint.y) And (P_k.y > testPoint.y)) Or      ' upward crossing
                    ((P_i.y > testPoint.y) And (P_k.y <= testPoint.y))) Then   ' downward crossing
                    ' compute  the actual edge-ray intersect x-coordinate
                    Dim vt As Double = (testPoint.y - P_i.y) / (P_k.y - P_i.y)
                    If (testPoint.x < P_i.x + vt * (P_k.x - P_i.x)) Then   ' x < intersect
                        cn += 1
                    End If
                End If
            Next

            Return Not (cn Mod 2 = 0)   ' or cn and 1
        End Function
        ''' <summary>
        ''' Intersect with a Segment and change P1 or P2 to clip polygon bounds
        ''' </summary>
        ''' <param name="P1">Byref Point 1</param>
        ''' <param name="P2">Byref Point 2</param>
        ''' <returns>
        ''' 2 if segment is completely inside
        ''' 1 if point P1 is inside and P2 is changed to intersect the polygon edge
        ''' -1 if point P2 is inside and P1 is changed to intersect the polygon edge
        ''' -2 if segmeint is completely outside
        ''' </returns>
        Public Function intersectWithSegment(ByRef P1 As Point, ByRef P2 As Point) As Integer
            Dim retValue As Integer

            If Me.pointInPolygon(P1) Then
                If Me.pointInPolygon(P2) Then
                    Return 2
                Else
                    ' P1 is inside, P2 is outside
                    retValue = 1
                End If
            Else
                If Me.pointInPolygon(P2) Then
                    ' P2 is inside, P1 is outside
                    retValue = -1
                Else
                    Return -2
                End If
            End If

            '
            ' loop throug all edges of the polygon
            '

            Dim i As Integer           ' start edge point index
            Dim k As Integer           ' end edge point index

            Dim P_i As Point
            Dim P_k As Point

            Dim pLine As line
            Dim sLine As line
            Dim PIntersect As Point
            Dim contact As Boolean

            For j = 0 To points.Count - 1
                i = j
                If j = points.Count - 1 Then

                    If closed Then
                        'Last edge connects last node with first node
                        k = 0
                    Else
                        ' there is no last edge
                        k = j
                    End If
                Else
                    k = j + 1
                End If

                ' point i and j as polygon edge

                P_i = points(points.ElementAt(i).Key)
                P_k = points(points.ElementAt(k).Key)

                pLine = New line(P_i, P_k)
                sLine = New line(P1, P2)
                PIntersect = pLine.intersect(sLine, True)
                If IsNothing(PIntersect) Then
                    'Return 0
                Else
                    If retValue = 1 Then
                        P2 = PIntersect
                        contact = True
                        Exit For
                    ElseIf retValue = -1 Then
                        P1 = PIntersect
                        contact = True
                        Exit For
                    End If
                End If
            Next

            Return retValue

        End Function

    End Class

    Public Class line
        Private p_Start As New Vector
        Private p_direction As New Vector

        ''' <summary>
        ''' Start Point
        ''' </summary>
        ''' <returns></returns>
        Public Property P As Vector
            Get
                Return p_Start
            End Get
            Set(value As Vector)
                p_Start = value
            End Set
        End Property
        ''' <summary>
        ''' End Point
        ''' </summary>
        ''' <returns></returns>
        Public Property P_End As Vector
            Get
                P_End = New Vector
                P_End = p_Start.add(d)
            End Get
            Set(value As Vector)
                d.x = value.x - P.x
                d.y = value.y - P.y
            End Set
        End Property
        ''' <summary>
        ''' direction
        ''' </summary>
        ''' <returns></returns>
        Public Property d As Vector
            Get
                Return p_direction
            End Get
            Set(value As Vector)
                p_direction = value
            End Set
        End Property
        ''' <summary>
        ''' StartPoint x
        ''' </summary>
        ''' <returns></returns>
        Public Property px As Double
            Get
                Return p_Start.x
            End Get
            Set(value As Double)
                p_Start.x = value
            End Set
        End Property
        ''' <summary>
        ''' Start Point y
        ''' </summary>
        ''' <returns></returns>
        Public Property py As Double
            Get
                Return p_Start.y
            End Get
            Set(value As Double)
                p_Start.y = value
            End Set
        End Property
        ''' <summary>
        ''' Direction x
        ''' </summary>
        ''' <returns></returns>
        Public Property dx As Double
            Get
                Return p_direction.x
            End Get
            Set(value As Double)
                p_direction.x = value
            End Set
        End Property
        ''' <summary>
        ''' Direction y
        ''' </summary>
        ''' <returns></returns>
        Public Property dy As Double
            Get
                Return p_direction.y
            End Get
            Set(value As Double)
                p_direction.y = value
            End Set
        End Property
        ''' <summary>
        ''' New
        ''' </summary>
        Public Sub New()
            p_Start = New Vector
            p_direction = New Vector
        End Sub
        ''' <summary>
        ''' New by Start Point coordinates and direction vector
        ''' </summary>
        ''' <param name="Px">Start Point X</param>
        ''' <param name="Py">Start Point Y</param>
        ''' <param name="dx">direction vector x</param>
        ''' <param name="dy">direction vector y</param>
        Public Sub New(Px As Double, Py As Double, dx As Double, dy As Double)
            p_Start = New Vector(Px, Py)
            p_direction = New Vector(dx, dy)
        End Sub
        ''' <summary>
        ''' New Line by Start and End-Point
        ''' </summary>
        ''' <param name="P1"></param>
        ''' <param name="P2"></param>
        Public Sub New(P1 As Point, P2 As Point)
            p_Start = New Vector(P1.x, P1.y)
            p_direction = New Vector(P1, P2)
        End Sub
        ''' <summary>
        ''' parallel line
        ''' </summary>
        ''' <param name="distance">positive right</param>
        ''' <returns></returns>
        Public Function offsetLine(distance As Double) As line
            Dim newLine As New line

            newLine = Me.MemberwiseClone

            'offsetLine.P = Me.P
            'offsetLine.d = Me.d

            Dim offsetVector As New Vector
            offsetVector = newLine.d
            offsetVector = offsetVector.normal(True)
            offsetVector = offsetVector.multiply(-distance)

            newLine.P = newLine.P.add(offsetVector)
            Return newLine

        End Function

        ''' <summary>
        ''' intersection with a second line
        ''' </summary>
        ''' <param name="other"></param>
        ''' <param name="withinDirectionVectorLength"></param>
        ''' <returns></returns>
        Public Function intersect(other As line, Optional withinDirectionVectorLength As Boolean = False) As Point
            Dim TOL As Double = 0.000001

            'Dim n As Double
            Dim m As Double
            Dim n As Double
            intersect = New Point
            'n = -(dx * (Py - other.Py) + dy * other.Px - dy * Px) / (other.dx * dy - dx * other.dy)
            If (other.dx * dy - dx * other.dy) = 0 Then
                Return Nothing
            Else
                m = -(other.dx * (py - other.py) + other.dy * other.px - other.dy * px) / (other.dx * dy - dx * other.dy)

                If withinDirectionVectorLength Then
                    n = -(dx * (other.py - py) + dy * px - dy * other.px) / (dx * other.dy - other.dx * dy)
                    If Not (m >= 0 - TOL And m <= 1 + TOL) Or Not (n >= 0 - TOL And n <= 1 + TOL) Then
                        Return Nothing
                    End If
                End If

                intersect.x = px + m * dx
                intersect.y = py + m * dy
            End If

        End Function
        ''' <summary>
        ''' get closest point on line
        ''' </summary>
        ''' <param name="P"></param>
        ''' <returns></returns>
        Public Function closestPoint(P As Vector) As Point
            Dim g2 As New line
            closestPoint = New Point

            g2.P = P
            g2.d = Me.d.normal
            closestPoint = Me.intersect(g2)
        End Function

    End Class
End Namespace
