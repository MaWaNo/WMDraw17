Imports Microsoft.VisualBasic

Namespace WallnerMild.Geom

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
        Public ReadOnly Property normal(Optional uniform As Boolean = True) As Vector
            Get
                Dim n As New Vector
                n.x = -y
                n.y = x
                If uniform Then
                    n.divide(length)
                End If
                Return n
            End Get
        End Property

        ''' <summary>
        ''' add a vector
        ''' </summary>
        ''' <param name="vect"></param>
        Public Sub add(vect As Vector)
            x += vect.x
            y += vect.y
        End Sub

        ''' <summary>
        ''' add a vector
        ''' </summary>
        ''' <param name="vect"></param>
        Public Sub subtract(vect As Vector)
            x -= vect.x
            y -= vect.y
        End Sub

        ''' <summary>
        ''' divide vector by value
        ''' </summary>
        ''' <param name="value"></param>
        Public Sub divide(value As Double)
            x = x / value
            y = y / value
        End Sub

        ''' <summary>
        ''' multiply vector by value
        ''' </summary>
        ''' <param name="value"></param>
        Public Sub multiply(value As Double)
            x = x * value
            y = y * value
        End Sub

        ''' <summary>
        ''' unify vector
        ''' </summary>
        Public Sub unify()
            If length > 0 Then
                Me.divide(length)
            End If
        End Sub
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
    End Class

    Public Class line
        Private p_Start As Vector
        Private p_direction As Vector

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
        ''' <param name="Px"></param>
        ''' <param name="Py"></param>
        ''' <param name="dx"></param>
        ''' <param name="dy"></param>
        Public Sub New(Px As Double, Py As Double, dx As Double, dy As Double)
            p_Start = New Vector(Px, Py)
            p_direction = New Vector(dx, dy)
        End Sub
        ''' <summary>
        ''' parallel line
        ''' </summary>
        ''' <param name="distance">positive right</param>
        ''' <returns></returns>
        Public Function offsetLine(distance As Double) As line
            offsetLine = New line
            offsetLine.P = Me.P
            offsetLine.d = Me.d

            Dim offsetVector As Vector
            offsetVector = p_direction.normal(True)
            offsetVector.multiply(-distance)

            offsetLine.P.add(offsetVector)
        End Function

        ''' <summary>
        ''' intersection with a second line
        ''' </summary>
        ''' <param name="other"></param>
        ''' <returns>nothing if parallel lines</returns>
        Public Function intersect(other As line) As Vector
            'Dim n As Double
            Dim m As Double
            intersect = New Vector
            'n = -(dx * (Py - other.Py) + dy * other.Px - dy * Px) / (other.dx * dy - dx * other.dy)
            If (other.dx * dy - dx * other.dy) = 0 Then
                Return Nothing
            Else
                m = -(other.dx * (py - other.py) + other.dy * other.px - other.dy * px) / (other.dx * dy - dx * other.dy)
                intersect.x = px + m * dx
                intersect.y = py + m * dy
            End If

        End Function
        ''' <summary>
        ''' get closest point on line
        ''' </summary>
        ''' <param name="P"></param>
        ''' <returns></returns>
        Public Function closestPoint(P As Vector) As Vector
            Dim g2 As New line
            closestPoint = New Vector

            g2.P = P
            g2.d = Me.d.normal
            closestPoint = Me.intersect(g2)
        End Function

    End Class
End Namespace
