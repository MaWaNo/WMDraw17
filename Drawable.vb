
Imports System.Runtime.InteropServices

Namespace WallnerMild.Draw

    <ComVisible(False)>
    Public Interface Drawable
        Inherits IComparable(Of Drawable)

        ' Define the delegate function for the comparisons.
        Delegate Function contextCoordinates(ByVal p As Point) As Point
        Delegate Function contextSize(ByVal inputSize As size) As size

        ''' <summary>
        ''' Drawing Order
        ''' </summary>
        ''' <returns></returns>
        Property zIndex As Long
        Property pen As pen

        Function boundingRectangle(Optional ref As Reference = Reference.world) As Double()
        Sub draw(contextobject As ContextObject, contextCoordinatesDelegate As contextCoordinates, Optional contextSizeDelegate As contextSize = Nothing)

    End Interface

    ''' <summary>
    ''' Geometric references for coordinates or sizes
    ''' </summary>
    Public Enum Reference
        ''' <summary>
        ''' world coordinates
        ''' </summary>
        world
        ''' <summary>
        ''' coordinates in mm
        ''' </summary>
        contextMillimeters
        ''' <summary>
        ''' context specific units (usually pixels)
        ''' </summary>
        contextUnits
        ''' <summary>
        ''' Fraction of the drawing context
        ''' </summary>
        contextFraction
    End Enum
    ''' <summary>
    ''' Point display types
    ''' </summary>
    Public Enum PointDisplay
        ''' <summary>
        ''' </summary>
        circle
        ''' <summary>
        ''' </summary>
        plus
        ''' <summary>
        ''' </summary>
        x
        dot
        invisible
    End Enum
    <ComVisible(False)>
    Public Class size
        Private p_width As Double
        Private p_height As Double

        Private p_Reference As New Reference

        Public Sub New()
            p_width = 0
            p_height = 0
            p_Reference = Reference.contextUnits
        End Sub

        Public Property width As Double
            Get
                Return p_width
            End Get
            Set(value As Double)
                p_width = value
            End Set
        End Property

        Public Property height As Double
            Get
                Return p_height
            End Get
            Set(value As Double)
                p_height = value
            End Set
        End Property
        Public ReadOnly Property average As Double
            Get
                average = (width + height) / 2
            End Get
        End Property

        Public Property Reference As Reference
            Get
                Return p_Reference
            End Get
            Set(value As Reference)
                p_Reference = value
            End Set
        End Property
    End Class

#Region "Point Class"
    ''' <summary>
    ''' Point as a drawable item
    ''' </summary>
    <ComVisible(False)>
    Public Class Point
        Implements Drawable

        Private p_zIndex As Long
        Private p_pen As pen

        ''' <summary>
        ''' CompareTo as implementation of IComparable
        ''' </summary>
        ''' <param name="other"></param>
        ''' <returns></returns>
        Public Function CompareTo(ByVal other As Drawable) As Integer Implements System.IComparable(Of Drawable).CompareTo
            If p_zIndex = other.zIndex Then
                Return 0
            Else
                If p_zIndex < other.zIndex Then
                    Return -1
                Else
                    Return 1
                End If
            End If
        End Function
        ''' <summary>
        ''' x-Coordinate in world coordinates
        ''' </summary>
        Private p_x As Double
        ''' <summary>
        ''' y-Coordinate in world coordinates
        ''' </summary>
        Private p_y As Double
        ''' <summary>
        ''' Coordinate reference
        ''' </summary>
        Private p_coordinateReference As Reference
        ''' <summary>
        ''' Display type
        ''' </summary>
        Private p_display As WallnerMild.Draw.PointDisplay
        ''' <summary>
        ''' Size
        ''' </summary>
        Private p_displaySize As New size

        ''' <summary>
        ''' Point at position 0/0
        ''' </summary>
        Sub New()
            p_x = 0
            p_y = 0
            p_display = PointDisplay.dot
            p_displaySize.width = 1
            p_displaySize.height = 1
        End Sub
        ''' <summary>
        ''' Pint by given world coordinates
        ''' </summary>
        ''' <param name="x"></param>
        ''' <param name="y"></param>
        Sub New(x As Double, y As Double)
            p_x = x
            p_y = y
            p_display = PointDisplay.dot
            p_displaySize.width = 2
            p_displaySize.height = 2
            p_displaySize.Reference = Reference.contextMillimeters

            p_coordinateReference = Reference.world
        End Sub
        ''' <summary>
        ''' point by coordinates in defined reference coordinates
        ''' </summary>
        ''' <param name="x"></param>
        ''' <param name="y"></param>
        ''' <param name="ref"></param>
        Sub New(x As Double, y As Double, ref As Reference)
            p_x = x
            p_y = y
            p_coordinateReference = ref

            p_display = PointDisplay.dot
            p_displaySize.width = 2
            p_displaySize.height = 2
            p_displaySize.Reference = Reference.contextMillimeters

            p_displaySize.Reference = Reference.contextMillimeters
        End Sub
        ''' <summary>
        ''' point by coordinates in defined reference coordinates with defined display options
        ''' </summary>
        ''' <param name="x"></param>
        ''' <param name="y"></param>
        ''' <param name="ref"></param>
        ''' <param name="display"></param>
        ''' <param name="displaySize"></param>
        ''' <param name="displayReference"></param>
        Sub New(x As Double, y As Double, ref As Reference, display As PointDisplay, displaySize As Double, displayReference As Reference)
            p_x = x
            p_y = y
            p_coordinateReference = ref

            p_display = PointDisplay.dot
            p_displaySize.width = 2
            p_displaySize.height = 2
            p_displaySize.Reference = Reference.contextMillimeters

            p_displaySize.Reference = Reference.contextMillimeters
        End Sub

        ''' <summary>
        ''' Point in world-coordinates and defined display type and size
        ''' </summary>
        ''' <param name="x">x-coordinate in world coordinates</param>
        ''' <param name="y">y-coordinate in world coordinates</param>
        ''' <param name="display">display type</param>
        ''' <param name="displaySize">display size</param>
        Sub New(x As Double, y As Double, display As PointDisplay, displaySize As Double, displayReference As Reference)
            p_x = x
            p_y = y
            p_coordinateReference = Reference.world

            p_display = display
            p_displaySize.width = displaySize
            p_displaySize.height = displaySize
            p_displaySize.Reference = displayReference
        End Sub
        ''' <summary>
        ''' coordinate reference
        ''' </summary>
        ''' <returns></returns>
        Public Property coordinateReference As Reference
            Get
                Return p_coordinateReference
            End Get
            Set(value As Reference)
                p_coordinateReference = value
            End Set
        End Property
        ''' <summary>
        ''' draw the item
        ''' </summary>
        ''' <param name="contextobject"></param>
        ''' <param name="contextCoordinatesDelegate"></param>
        Public Sub draw(contextobject As ContextObject, contextCoordinatesDelegate As Drawable.contextCoordinates, Optional contextSizeDelegate As Drawable.contextSize = Nothing) Implements Drawable.draw

            If TypeOf contextobject.Item Is System.Windows.Controls.Canvas Then
                Dim myCanvas As New System.Windows.Controls.Canvas
                myCanvas = TryCast(contextobject.Item, System.Windows.Controls.Canvas)
                With myCanvas
                    Dim myPoint As New System.Windows.Shapes.Ellipse

                    With myPoint
                        Dim p1 As New Point
                        Dim s1 As New size
                        s1 = contextSizeDelegate(Me.displaySize)
                        '
                        ' calculate coordinates in local system
                        ' by a call-back to the calling drawing object
                        '
                        p1 = contextCoordinatesDelegate(Me)

                        .Width = s1.width
                        .Height = s1.height

                        .Stroke = System.Windows.Media.Brushes.Green
                        .StrokeThickness = 1

                        myCanvas.Children.Add(myPoint)

                        'position
                        System.Windows.Controls.Canvas.SetLeft(myPoint, p1.x - (s1.width / 2))
                        System.Windows.Controls.Canvas.SetTop(myPoint, p1.y - (s1.height / 2))

                    End With

                End With
            Else
                Throw New NotImplementedException()
            End If

        End Sub
        ''' <summary>
        ''' bounding rectangle
        ''' </summary>
        ''' <param name="ref"></param>
        ''' <returns></returns>
        Public Function boundingRectangle(Optional ref As Reference = Reference.world) As Double() Implements Drawable.boundingRectangle

            Dim r(3) As Double

            Dim p1 As New Point
            Dim p2 As New Point

            r(0) = p1.x
            r(1) = p1.y
            r(2) = p1.x
            r(3) = p1.y

            Return r
        End Function


        ''' <summary>
        ''' x-coordinate (in Coordinate-reference units)
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
        ''' y-coordinate (in Coordinate-reference units)
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
        ''' display type
        ''' </summary>
        Public Property display As PointDisplay
            Get
                Return p_display
            End Get
            Set(value As PointDisplay)
                p_display = value
            End Set
        End Property

        ''' <summary>
        ''' display size (in displayReference units)
        ''' </summary>
        ''' <returns></returns>
        Public Property displaySize As size
            Get
                Return p_displaySize
            End Get
            Set(value As size)
                p_displaySize = value
            End Set
        End Property

        Public Property pen As pen Implements Drawable.pen
            Get
                Return p_pen
            End Get
            Set(value As pen)
                p_pen = value
            End Set
        End Property

        Public Property zIndex As Long Implements Drawable.zIndex
            Get
                Return p_zIndex
            End Get
            Set(value As Long)
                p_zIndex = value
            End Set
        End Property
    End Class

#End Region
#Region "Beam Class"
    <ComVisible(False)>
    Public Class BeamUniformLoad
        Inherits Line
        Implements Drawable

        Public Enum loadOrientationType
            localOrientation
            globalOrientation
        End Enum

        Public Enum loadedLengthType
            horizontalProjection
            realLength
        End Enum

        Private Const LOADOFFSET = 20

        Private p_offsetIndex As Integer
        Private p_orientation As loadOrientationType
        Private p_loadedLength As loadedLengthType

        Sub New(startX As Double, startY As Double, endX As Double, endY As Double, loadedLength As loadedLengthType, orientation As loadOrientationType, offset As Integer)
            Me.startPoint = New Point(startX, startY)
            Me.endPoint = New Point(endX, endY)
            Me.pen = New pen
            Me.pen.thickness = 3
            Me.p_loadedLength = loadedLength
            Me.p_orientation = orientation
            Me.p_offsetIndex = offset
        End Sub

        Public Property loadedLength As loadedLengthType
            Get
                Return p_loadedLength
            End Get
            Set(value As loadedLengthType)
                p_loadedLength = value
            End Set
        End Property

        Public Property orientation As loadOrientationType
            Get
                Return p_orientation
            End Get
            Set(value As loadOrientationType)
                p_orientation = value
            End Set
        End Property

        Public Overrides Sub draw(contextobject As ContextObject, contextCoordinatesDelegate As Drawable.contextCoordinates, Optional contextSizeDelegate As Drawable.contextSize = Nothing) Implements Drawable.draw
            If TypeOf contextobject.Item Is System.Windows.Controls.Canvas Then
                Dim myCanvas As New System.Windows.Controls.Canvas
                myCanvas = TryCast(contextobject.Item, System.Windows.Controls.Canvas)
                With myCanvas
                    Dim line1 As New System.Windows.Shapes.Line
                    Dim line2 As New System.Windows.Shapes.Line

                    '
                    ' calculate coordinates in local system
                    ' by a call-back to the calling drawing object
                    '
                    Dim p1 As New Point
                    Dim p2 As New Point
                    Dim t As size

                    Dim vx As Double
                    Dim vy As Double
                    Dim tmp As Double

                    p1 = contextCoordinatesDelegate(startPoint)
                    p2 = contextCoordinatesDelegate(endPoint)

                    t = contextSizeDelegate(Me.pen.size)
                    If t.width < 1 Then t.width = 1

                    vx = p2.x - p1.x
                    vy = p2.y - p1.y

                    '
                    ' normal vector
                    '
                    tmp = vx
                    vx = -vy
                    vy = tmp
                    '
                    ' uniform vector
                    '
                    tmp = Math.Sqrt(vx ^ 2 + vy ^ 2)
                    If tmp <> 0 Then
                        vx = vx / tmp
                        vy = vy / tmp
                    End If

                    With line2
                        .X1 = p1.x - vx * LOADOFFSET * Me.p_offsetIndex
                        .Y1 = p1.y - vy * LOADOFFSET * Me.p_offsetIndex
                        .X2 = p2.x - vx * LOADOFFSET * Me.p_offsetIndex
                        .Y2 = p2.y - vy * LOADOFFSET * Me.p_offsetIndex
                        .Stroke = Me.pen.stroke
                        .StrokeThickness = 1
                    End With

                    myCanvas.Children.Add(line2)
                End With
            Else
                Throw New NotImplementedException()
            End If
        End Sub
    End Class
    <ComVisible(False)>
    Public Class Beam
        Inherits Line
        Implements Drawable

        Public Sub New()
            MyBase.New
            Me.pen.thickness = 3
        End Sub

        Sub New(startX As Double, startY As Double, endX As Double, endY As Double)
            Me.startPoint = New Point(startX, startY)
            Me.endPoint = New Point(endX, endY)
            Me.pen = New pen
            Me.pen.thickness = 3
        End Sub

        Public Overrides Sub draw(contextobject As ContextObject, contextCoordinatesDelegate As Drawable.contextCoordinates, Optional contextSizeDelegate As Drawable.contextSize = Nothing) Implements Drawable.draw
            If TypeOf contextobject.Item Is System.Windows.Controls.Canvas Then
                Dim myCanvas As New System.Windows.Controls.Canvas
                myCanvas = TryCast(contextobject.Item, System.Windows.Controls.Canvas)
                With myCanvas
                    Dim line1 As New System.Windows.Shapes.Line
                    Dim line2 As New System.Windows.Shapes.Line

                    '
                    ' calculate coordinates in local system
                    ' by a call-back to the calling drawing object
                    '
                    Dim p1 As New Point
                    Dim p2 As New Point
                    Dim t As size

                    Dim vx As Double
                    Dim vy As Double
                    Dim tmp As Double

                    p1 = contextCoordinatesDelegate(startPoint)
                    p2 = contextCoordinatesDelegate(endPoint)

                    t = contextSizeDelegate(Me.pen.size)
                    If t.width < 1 Then t.width = 1

                    vx = p2.x - p1.x
                    vy = p2.y - p1.y

                    '
                    ' normal vector
                    '
                    tmp = vx
                    vx = -vy
                    vy = tmp
                    '
                    ' uniform vector
                    '
                    tmp = Math.Sqrt(vx ^ 2 + vy ^ 2)
                    If tmp <> 0 Then
                        vx = vx / tmp
                        vy = vy / tmp
                    End If

                    With line1
                        .X1 = p1.x
                        .Y1 = p1.y
                        .X2 = p2.x
                        .Y2 = p2.y
                        .Stroke = Me.pen.stroke
                        .StrokeThickness = t.width
                        .StrokeDashArray = Me.pen.dashArray
                    End With

                    With line2
                        .X1 = p1.x + vx * t.width * 1.5
                        .Y1 = p1.y + vy * t.width * 1.5
                        .X2 = p2.x + vx * t.width * 1.5
                        .Y2 = p2.y + vy * t.width * 1.5
                        .Stroke = Me.pen.stroke
                        .StrokeThickness = 1

                        ' ----    ----    ----    ----
                        Dim sda As New Windows.Media.DoubleCollection
                        sda.Add(4)
                        sda.Add(4)
                        .StrokeDashArray = sda
                    End With

                    myCanvas.Children.Add(line1)
                    myCanvas.Children.Add(line2)
                End With
            Else
                Throw New NotImplementedException()
            End If
        End Sub

    End Class
#End Region
#Region "Text Class"
    Public Class Text
        Implements Drawable

        Private p_text As String
        Private p_position As Point
        Private p_angle As Double


        ''' <summary>
        ''' Text to be written
        ''' </summary>
        ''' <returns></returns>
        Public Property text As String
            Get
                Return p_text
            End Get
            Set(value As String)
                p_text = value
            End Set
        End Property
        Public Property angle() As Double
            Get
                Return p_angle
            End Get
            Set(value As Double)
                p_angle = value
            End Set
        End Property
        Public Property position() As Point
            Get
                Return p_position
            End Get
            Set(value As Point)
                p_position = value
            End Set
        End Property
        Public Property zIndex As Long Implements Drawable.zIndex
            Get
                'Throw New NotImplementedException()
            End Get
            Set(value As Long)
                'Throw New NotImplementedException()
            End Set
        End Property

        Private Property Drawable_pen As pen Implements Drawable.pen
            Get
                'Throw New NotImplementedException()
            End Get
            Set(value As pen)
                'Throw New NotImplementedException()
            End Set
        End Property

        Public Sub draw(contextobject As ContextObject, contextCoordinatesDelegate As Drawable.contextCoordinates, Optional contextSizeDelegate As Drawable.contextSize = Nothing) Implements Drawable.draw
            If TypeOf contextobject.Item Is System.Windows.Controls.Canvas Then
                Dim myCanvas As New System.Windows.Controls.Canvas
                myCanvas = TryCast(contextobject.Item, System.Windows.Controls.Canvas)
                With myCanvas
                    Dim tb As New Windows.Controls.TextBlock

                    With tb
                        .Text = p_text

                        '
                        ' calculate coordinates in local system
                        ' by a call-back to the calling drawing object
                        '
                        'p1 = contextCoordinatesDelegate(startPoint)
                        '                        t = contextSizeDelegate(Me.pen.size)
                    End With

                    Dim myRotateTransform As New Windows.Media.RotateTransform
                    myRotateTransform.Angle = angle
                    tb.RenderTransform = myRotateTransform

                    Dim p As New Point

                    '
                    ' calculate coordinates in local system
                    ' by a call-back to the calling drawing object
                    '
                    p = contextCoordinatesDelegate(p_position)
                    myCanvas.Children.Add(tb)
                    myCanvas.SetLeft(tb, p.x)
                    myCanvas.SetTop(tb, p.y)
                End With
            Else
                Throw New NotImplementedException()
            End If
        End Sub

        Public Function boundingRectangle(Optional ref As Reference = Reference.world) As Double() Implements Drawable.boundingRectangle
            'Throw New NotImplementedException()
            Dim r(3) As Double
            Return r
        End Function

        Public Function CompareTo(other As Drawable) As Integer Implements IComparable(Of Drawable).CompareTo
            'Throw New NotImplementedException()
        End Function
    End Class

#End Region
#Region "Line Class"
    ''' <summary>
    ''' as a drawable item
    ''' </summary>
    <ComVisible(False)>
    Public Class Line
        Implements Drawable

        Private p_pen As pen
        Private p_zIndex As Long
        ''' <summary>
        ''' Start Point
        ''' </summary>
        Private p_startPoint As Point
        ''' <summary>
        ''' End Point
        ''' </summary>
        Private p_endPoint As Point
        Sub New()
            Me.startPoint = New Point
            Me.endPoint = New Point
        End Sub
        ''' <summary>
        ''' Line from start and end points in world coordinates
        ''' </summary>
        ''' <param name="startPoint"></param>
        ''' <param name="endPont"></param>
        Sub New(startPoint As Point, endPont As Point)
            Me.startPoint = startPoint
            Me.endPoint = endPoint
        End Sub
        ''' <summary>
        ''' Line from start and end coordinates in world coordinates
        ''' </summary>
        ''' <param name="startX"></param>
        ''' <param name="startY"></param>
        ''' <param name="endX"></param>
        ''' <param name="endY"></param>
        Sub New(startX As Double, startY As Double, endX As Double, endY As Double)
            Me.startPoint = New Point(startX, startY)
            Me.endPoint = New Point(endX, endY)
        End Sub
        ''' <summary>
        ''' Line from start and end coordinates in coordinates in defined reference 
        ''' </summary>
        ''' <param name="startX"></param>
        ''' <param name="startY"></param>
        ''' <param name="endX"></param>
        ''' <param name="endY"></param>
        ''' <param name="coordinateReference"></param>
        Sub New(startX As Double, startY As Double, endX As Double, endY As Double, coordinateReference As Reference)
            Me.startPoint = New Point(startX, startY, coordinateReference)
            Me.endPoint = New Point(endX, endY, coordinateReference)
        End Sub

        ''' <summary>
        ''' place on canvas
        ''' </summary>
        ''' <param name="contextobject"></param>
        ''' <param name="contextCoordinatesDelegate"></param>
        Public Overridable Sub draw(contextobject As ContextObject, contextCoordinatesDelegate As Drawable.contextCoordinates, Optional contextSizeDelegate As Drawable.contextSize = Nothing) Implements Drawable.draw

            If TypeOf contextobject.Item Is System.Windows.Controls.Canvas Then
                Dim myCanvas As New System.Windows.Controls.Canvas
                myCanvas = TryCast(contextobject.Item, System.Windows.Controls.Canvas)
                With myCanvas
                    Dim line1 As New System.Windows.Shapes.Line
                    With line1
                        Dim p1 As New Point
                        Dim p2 As New Point
                        Dim t As size

                        '
                        ' calculate coordinates in local system
                        ' by a call-back to the calling drawing object
                        '
                        p1 = contextCoordinatesDelegate(startPoint)
                        p2 = contextCoordinatesDelegate(endPoint)

                        t = contextSizeDelegate(Me.pen.size)
                        If t.width < 1 Then t.width = 1

                        .X1 = p1.x
                        .Y1 = p1.y
                        .X2 = p2.x
                        .Y2 = p2.y
                        .Stroke = Me.pen.stroke
                        .StrokeThickness = t.width
                        .StrokeDashArray = Me.pen.dashArray
                    End With

                    myCanvas.Children.Add(line1)
                End With
            Else
                Throw New NotImplementedException()
            End If

        End Sub

        Private Function boundingRectangle(Optional ref As Reference = Reference.world) As Double() Implements Drawable.boundingRectangle

            Dim r(3) As Double

            Dim p1 As New Point
            Dim p2 As New Point

            If startPoint.coordinateReference = ref Then
                r(0) = startPoint.x
                r(1) = startPoint.y
            End If

            If startPoint.coordinateReference = ref Then
                r(2) = endPoint.x
                r(3) = endPoint.y
            Else
                r(2) = startPoint.x
                r(3) = startPoint.y
            End If

            ' sort points
            For i As Integer = 0 To 1
                If r(0 + i) > r(2 + i) Then
                    Dim t As Double = r(0 + i)
                    r(0 + i) = r(2 + i)
                    r(2 + i) = t
                End If
            Next

            Return r
        End Function

        Public Property startPoint As Point
            Get
                Return p_startPoint
            End Get
            Set(value As Point)
                p_startPoint = value
            End Set
        End Property

        Public Property endPoint As Point
            Get
                Return p_endPoint
            End Get
            Set(value As Point)
                p_endPoint = value
            End Set
        End Property

        Public Property pen As pen Implements Drawable.pen
            Get
                Return p_pen
            End Get
            Set(value As pen)
                p_pen = value
            End Set
        End Property

        Public Property zIndex As Long Implements Drawable.zIndex
            Get
                Return p_zIndex
            End Get
            Set(value As Long)
                p_zIndex = value
            End Set
        End Property


        ''' <summary>
        ''' CompareTo as implementation of IComparable
        ''' </summary>
        ''' <param name="other"></param>
        ''' <returns></returns>
        Public Function CompareTo(ByVal other As Drawable) As Integer Implements System.IComparable(Of Drawable).CompareTo
            If p_zIndex = other.zIndex Then
                Return 0
            Else
                If p_zIndex < other.zIndex Then
                    Return -1
                Else
                    Return 1
                End If
            End If
        End Function

    End Class
#End Region

    <ComVisible(False)>
    Public Class Support
        Implements Drawable

        Const SUPPORT_HEIGHT_MM As Double = 7
        Const SUPPORT_WIDTH_MM As Double = 7
        Const SUPPORT_ANGLE_DEG As Double = 30

        Private p_position As New Point
        Private p_size As New size
        Private p_angle As Double
        Private p_caption As String

        ' stiffness values
        Private p_cx As Double
        Private p_cy As Double

        Private p_zIndex As Long

        Private p_pen As pen

        ''' <summary>
        ''' CompareTo as implementation of IComparable
        ''' </summary>
        ''' <param name="other"></param>
        ''' <returns></returns>
        Public Function CompareTo(ByVal other As Drawable) As Integer Implements System.IComparable(Of Drawable).CompareTo
            If p_zIndex = other.zIndex Then
                Return 0
            Else
                If p_zIndex < other.zIndex Then
                    Return -1
                Else
                    Return 1
                End If
            End If
        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="x"></param>
        ''' <param name="y"></param>
        ''' <param name="angle"></param>
        Public Sub New(x As Double, y As Double, angle As Double)
            With p_position
                .x = x
                .y = y
                .coordinateReference = Reference.world
            End With
            With p_size
                .width = 5
                .height = 5
                .Reference = Reference.contextMillimeters
            End With
            p_angle = angle
        End Sub
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="x"></param>
        ''' <param name="y"></param>
        ''' <param name="angle"></param>
        Public Sub New(x As Double, y As Double, angle As Double, caption As String)
            With p_position
                .x = x
                .y = y
                .coordinateReference = Reference.world
            End With
            With p_size
                .width = 5
                .height = 5
                .Reference = Reference.contextMillimeters
            End With
            p_caption = caption
            p_angle = angle
        End Sub
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="x"></param>
        ''' <param name="y"></param>
        ''' <param name="coordinateReference"></param>
        ''' <param name="size"></param>
        ''' <param name="sizeReference"></param>
        ''' <param name="angle"></param>
        Public Sub New(x As Double, y As Double, coordinateReference As Reference, size As Double, sizeReference As Reference, angle As Double)
            With p_position
                .x = x
                .y = y
                .coordinateReference = coordinateReference
            End With
            With p_size
                .width = size
                .height = size
                .Reference = sizeReference
            End With
            p_angle = angle
        End Sub

        Public Property position As Point
            Get
                Return p_position
            End Get
            Set(value As Point)
                p_position = value
            End Set
        End Property

        Public Property size As size
            Get
                Return p_size
            End Get
            Set(value As size)
                p_size = value
            End Set
        End Property
        Public Property caption As String
            Get
                Return p_caption
            End Get
            Set(value As String)
                p_caption = value
            End Set
        End Property
        Public Property angle As Double
            Get
                Return p_angle
            End Get
            Set(value As Double)
                p_angle = value
            End Set
        End Property

        Public Sub draw(contextobject As ContextObject, contextCoordinatesDelegate As Drawable.contextCoordinates, Optional contextSizeDelegate As Drawable.contextSize = Nothing) Implements Drawable.draw

            If TypeOf contextobject.Item Is System.Windows.Controls.Canvas Then
                Dim myCanvas As New System.Windows.Controls.Canvas
                myCanvas = TryCast(contextobject.Item, System.Windows.Controls.Canvas)
                With myCanvas
                    Dim myPolyline As New System.Windows.Shapes.Polyline

                    Dim p1 As New Point
                    Dim s1 As New size

                    Dim r As Double
                    Dim h As Double
                    Dim sa As Double = 0.5      ' sin(30)
                    Dim ca As Double = 0.866    ' cos(30)

                    With myPolyline

                        s1 = contextSizeDelegate(Me.size)
                        p1 = contextCoordinatesDelegate(Me.position)

                        .Stroke = System.Windows.Media.Brushes.Blue
                        .StrokeThickness = 2

                        r = s1.average / 5
                        h = s1.average

                        myPolyline.Points.Add(New System.Windows.Point(p1.x + r * sa, p1.y + r * ca))
                        myPolyline.Points.Add(New System.Windows.Point(p1.x + h * sa, p1.y + h * ca))
                        myPolyline.Points.Add(New System.Windows.Point(p1.x + h / 2, p1.y + h * ca))
                        myPolyline.Points.Add(New System.Windows.Point(p1.x - h / 2, p1.y + h * ca))
                        myPolyline.Points.Add(New System.Windows.Point(p1.x - h * sa, p1.y + h * ca))
                        myPolyline.Points.Add(New System.Windows.Point(p1.x - r * sa, p1.y + r * ca))

                        myPolyline.Fill = System.Windows.Media.Brushes.White
                        myCanvas.Children.Add(myPolyline)

                    End With

                    myPolyline = New System.Windows.Shapes.Polyline

                    With myPolyline
                        Dim a As Double

                        For i As Integer = -3 To 3
                            a = 90 * Math.PI / 180 * i / 3
                            myPolyline.Points.Add(New System.Windows.Point(p1.x + r * Math.Sin(a), p1.y + r * Math.Cos(a)))
                        Next

                        .Stroke = System.Windows.Media.Brushes.Blue
                        .StrokeThickness = 2
                        .Fill = System.Windows.Media.Brushes.White
                        myCanvas.Children.Add(myPolyline)

                    End With

                    If p_caption <> "" Then
                        Dim tb As New Windows.Controls.TextBlock

                        With tb
                            .Text = p_caption

                            '
                            ' calculate coordinates in local system
                            ' by a call-back to the calling drawing object
                            '
                            'p1 = contextCoordinatesDelegate(startPoint)
                            '                        t = contextSizeDelegate(Me.pen.size)
                        End With

                        Dim myRotateTransform As New Windows.Media.RotateTransform
                        myRotateTransform.Angle = angle
                        tb.RenderTransform = myRotateTransform

                        Dim p As New Point

                        '
                        ' calculate coordinates in local system
                        ' by a call-back to the calling drawing object
                        '

                        myCanvas.Children.Add(tb)
                        myCanvas.SetLeft(tb, p1.x + h / 2)
                        myCanvas.SetTop(tb, p1.y + h * ca)
                    End If
                End With
            Else
                Throw New NotImplementedException()
            End If

        End Sub

        Public Function boundingRectangle(Optional ref As Reference = Reference.world) As Double() Implements Drawable.boundingRectangle
            Return Me.position.boundingRectangle(ref)
        End Function

        Public Property zIndex As Long Implements Drawable.zIndex
            Get
                Return p_zIndex
            End Get
            Set(value As Long)
                p_zIndex = value
            End Set
        End Property

        Public Property pen As pen Implements Drawable.pen
            Get
                Return p_pen
            End Get
            Set(value As pen)
                p_pen = value
            End Set
        End Property
    End Class
    <ComVisible(False)>
    Public Class Circle
        Implements Drawable

        Private p_zIndex As Long
        ''' <summary>
        ''' CompareTo as implementation of IComparable
        ''' </summary>
        ''' <param name="other"></param>
        ''' <returns></returns>
        Public Function CompareTo(ByVal other As Drawable) As Integer Implements System.IComparable(Of Drawable).CompareTo
            If p_zIndex = other.zIndex Then
                Return 0
            Else
                If p_zIndex < other.zIndex Then
                    Return -1
                Else
                    Return 1
                End If
            End If
        End Function

        Public Property midPoint As Point
            Get
                Return Nothing
            End Get
            Set(value As Point)
            End Set
        End Property

        Public Property pen As pen
            Get
                Return Nothing
            End Get
            Set(value As pen)
            End Set
        End Property

        Public Property radius As Double
            Get
                Return Nothing
            End Get
            Set(value As Double)
            End Set
        End Property

        Public Property fill As fill
            Get
                Return Nothing
            End Get
            Set(value As fill)
            End Set
        End Property



        Public Sub draw(contextobject As ContextObject, contextCoordinatesDelegate As Drawable.contextCoordinates, Optional contextSizeDelegate As Drawable.contextSize = Nothing) Implements Drawable.draw
            Throw New NotImplementedException()
        End Sub

        Public Function boundingRectangle(Optional ref As Reference = Reference.world) As Double() Implements Drawable.boundingRectangle
            Throw New NotImplementedException()
        End Function

        Public Property zIndex As Long Implements Drawable.zIndex
            Get
                Return p_zIndex
            End Get
            Set(value As Long)
                p_zIndex = value
            End Set
        End Property

        Private Property Drawable_pen As pen Implements Drawable.pen
            Get
                Throw New NotImplementedException()
            End Get
            Set(value As pen)
                Throw New NotImplementedException()
            End Set
        End Property
    End Class

    <ComVisible(False)>
    Public Class pen

        'Public stroke As System.Windows.Media.Brush

        Private p_color As System.Windows.Media.Color
        'Private p_thicknessReference As Reference
        Private p_thickness As size
        Private p_dasharray As System.Windows.Media.DoubleCollection

        Public Const DASHARRAY_SOLID = ""
        Public Const DASHARRAY_DASHED_SHORT = "3,2"
        Public Const DASHARRAY_DASHED_MEDIUM = "5,3"
        Public Const DASHARRAY_DASHED_LONG = "6,3"
        Public Const DASHARRAY_POINTS = "1,1"
        Public Const DASHARRAY_DASHDOT = "5,1,1,1"

        Public Sub New()
            p_color = System.Windows.Media.Colors.Black
            p_thickness = New size
            p_thickness.width = 2
            p_thickness.Reference = Reference.contextUnits
            Me.dashString = DASHARRAY_SOLID
        End Sub
        ''' <summary>
        ''' Get the brush object with defined color
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property stroke As System.Windows.Media.Brush
            Get
                Return New System.Windows.Media.SolidColorBrush(Me.color)
            End Get
        End Property
        ''' <summary>
        ''' brush color
        ''' </summary>
        ''' <returns></returns>
        Public Property color As System.Windows.Media.Color
            Get
                Return p_color
            End Get
            Set(value As System.Windows.Media.Color)
                p_color = value
            End Set
        End Property
        ''' <summary>
        ''' Brush thickness in reference units
        ''' </summary>
        ''' <returns></returns>
        Public Property thickness As Double
            Get
                Return p_thickness.width
            End Get
            Set(value As Double)
                p_thickness.width = value
            End Set
        End Property
        ''' <summary>
        ''' Thickness reference
        ''' </summary>
        ''' <returns></returns>
        Public Property thicknessReference As Reference
            Get
                Return p_thickness.Reference
            End Get
            Set(value As Reference)
                p_thickness.Reference = value
            End Set
        End Property

        Public ReadOnly Property size As size
            Get
                Return p_thickness
            End Get
        End Property

        Public ReadOnly Property dashArray() As System.Windows.Media.DoubleCollection
            Get
                Return p_dasharray
            End Get
        End Property


        Public Property dashString() As String
            Get
                Return p_dasharray.ToString
            End Get
            Set(value As String)
                p_dasharray = New System.Windows.Media.DoubleCollection
                'p_dasharray.Clear()

                For Each nStr As String In Split(value, ",")
                    If IsNumeric(nStr) Then
                        p_dasharray.Add(CDbl(nStr))
                    End If
                Next

            End Set
        End Property
    End Class

    <ComVisible(False)>
    Public Class fill
    End Class

End Namespace
