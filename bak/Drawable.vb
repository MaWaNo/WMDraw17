
Imports System.Runtime.InteropServices
Imports System.Windows.Media

Namespace WMDraw

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
        ' Property fill As fill

        Function boundingRectangle(Optional ref As Reference = Reference.world) As Double()
        Sub draw(contextobject As ContextObject, contextCoordinatesDelegate As contextCoordinates, Optional contextSizeDelegate As contextSize = Nothing)

    End Interface


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
    Public Class myMath
        ''' <summary>
        ''' convert degrees to radians
        ''' </summary>
        ''' <param name="angle_degrees"></param>
        ''' <returns></returns>
        Public Shared Function radians(angle_degrees As Double) As Double
            radians = angle_degrees * Math.PI / 180
        End Function

        ''' <summary>
        ''' convert radians to degrees
        ''' </summary>
        ''' <param name="angle_radians">Angle in radians</param>
        ''' <returns></returns>
        Public Shared Function degrees(angle_radians As Double) As Double
            degrees = angle_radians * 180 / Math.PI
        End Function
        ''' <summary>
        ''' include a point in the bounding rectangle
        ''' </summary>
        ''' <param name="r">rectangle (0..xmin, 1...ymin, 2...xmax, 3...ymax</param>
        ''' <param name="p">point to include</param>
        Public Shared Sub includeInBoundingRectangle(r() As Double, p As Point)
            If p.x < r(0) Then r(0) = p.x
            If p.x > r(2) Then r(2) = p.x
            If p.y < r(1) Then r(1) = p.y
            If p.y > r(3) Then r(3) = p.y
        End Sub
    End Class
    ''' <summary>
    ''' Size of an object defined with a Reference
    ''' </summary>
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
        Public Property thickness As Double
            Get
                Return average
            End Get
            Set(value As Double)
                average = value
            End Set
        End Property
        Public Property average As Double
            Get
                average = (Math.Abs(width) + Math.Abs(height)) / 2
            End Get
            Set(value As Double)
                p_height = value
                p_width = value
            End Set
        End Property
        Public ReadOnly Property max As Double
            Get
                max = Math.Max(Math.Abs(width), Math.Abs(height))
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
        Private p_pen As New pen

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
        Private p_display As WMDraw.PointDisplay
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

                    Dim p1 As New Point
                    Dim s1 As New size
                    s1 = contextSizeDelegate(Me.displaySize)
                    p1 = contextCoordinatesDelegate(Me)

                    Select Case display
                        Case PointDisplay.x
                            Dim myLine As New System.Windows.Shapes.Line()
                            With myLine
                                .X1 = p1.x - s1.width / 2
                                .Y1 = p1.y + s1.height / 2

                                .X2 = p1.x + s1.width / 2
                                .Y2 = p1.y - s1.width / 2

                                .Stroke = pen.stroke
                                .StrokeThickness = 1
                            End With

                            myCanvas.Children.Add(myLine)
                            System.Windows.Controls.Canvas.SetLeft(myLine, p1.x)
                            System.Windows.Controls.Canvas.SetTop(myLine, p1.y)

                            Dim myLine2 As New System.Windows.Shapes.Line()
                            With myLine2
                                .X1 = p1.x + s1.width / 2
                                .Y1 = p1.y + s1.height / 2

                                .X2 = p1.x - s1.width / 2
                                .Y2 = p1.y - s1.width / 2

                                .Stroke = pen.stroke
                                .StrokeThickness = 1
                            End With

                            myCanvas.Children.Add(myLine2)
                            System.Windows.Controls.Canvas.SetLeft(myLine2, p1.x)
                            System.Windows.Controls.Canvas.SetTop(myLine2, p1.y)

                        Case PointDisplay.circle, PointDisplay.dot
                            Dim myPoint As New System.Windows.Shapes.Ellipse

                            With myPoint

                                .Width = s1.width
                                .Height = s1.height

                                .Stroke = pen.stroke
                                .StrokeThickness = 1

                                myCanvas.Children.Add(myPoint)

                                'position
                                System.Windows.Controls.Canvas.SetLeft(myPoint, p1.x - (s1.width / 2))
                                System.Windows.Controls.Canvas.SetTop(myPoint, p1.y - (s1.height / 2))

                            End With

                    End Select


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

            r(0) = Me.x
            r(1) = Me.y
            r(2) = Me.x
            r(3) = Me.y

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

        Private p_zindex As Long
        Private p_pen As pen

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
                Return p_zindex
            End Get
            Set(value As Long)
                p_zindex = value
            End Set
        End Property

        Private Property Drawable_pen As pen Implements Drawable.pen
            Get
                Return p_pen
            End Get
            Set(value As pen)
                p_pen = value
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
            Dim r(3) As Double
            'todo: somehow determine text size 

            r(0) = Me.position.x
            r(1) = Me.position.y
            r(2) = Me.position.x
            r(3) = Me.position.y
            Return r
        End Function

        Public Function CompareTo(other As Drawable) As Integer Implements IComparable(Of Drawable).CompareTo
            'Throw New NotImplementedException()
        End Function
    End Class

#End Region

#Region "Force and Moment"
    <ComVisible(False)>
    Public Class ForceArrow
        Implements Drawable
        '
        ' todo: Create a force or arrow drawable and a Moment drawable
        '
        Public Enum forceDirections
            ArrowAtTip
            ArrowAtBottom
        End Enum

        Public Enum forceTypes
            normalForce
            resistanceForce
        End Enum

        Public Enum endPointTypes
            none
            circle
            cross
        End Enum

        Public Enum arrowTipTypes
            defaultForceTip
            halfTipLeft
            halfTipRight
            directionIndicatorLeft
            directionIndicatorRight
        End Enum

        Const DEFAULT_ARROW_WIDTH_MM As Double = 3
        Const DEFAULT_ARROW_HEIGHT_MM As Double = 5

        Const DEFAULT_FORCE_SIZE_MM As Double = 10

        Private p_pen As pen
        Private p_zIndex As Long

        Private p_tipPoint As New Point
        Private p_caption As String

        Private p_arrowSize As New size

        Private p_maxSize As New size

        Private p_Fx As Double                  ' input Fx
        Private p_Fy As Double                  ' input Fy
        Private p_inputFxFy As Boolean

        Private p_F As Double                   ' input F
        Private p_angle As Double               ' input alpha
        Private p_inputFAlpha As Boolean

        Private p_maximumForce As Double
        Private p_isMaximumForceDefined As Boolean

        Private p_ForceType As forceTypes
        Private p_ArrowTipType As arrowTipTypes

        Private p_forceDirection As forceDirections

        Private p_endPointDisplay As WMDraw.PointDisplay

        Sub New()
            ' default force size
            Me.p_maxSize.width = DEFAULT_FORCE_SIZE_MM
            Me.p_maxSize.height = DEFAULT_FORCE_SIZE_MM
            Me.p_maxSize.Reference = Reference.contextMillimeters
            ' default arrow size
            Me.p_arrowSize.width = DEFAULT_ARROW_WIDTH_MM
            Me.p_arrowSize.height = DEFAULT_ARROW_WIDTH_MM
            Me.p_arrowSize.Reference = Reference.contextMillimeters
            ' default maximum force
            p_maximumForce = 1
            p_isMaximumForceDefined = False

            p_ForceType = forceTypes.normalForce
            p_endPointDisplay = PointDisplay.invisible


        End Sub
        ''' <summary>
        ''' new Force Arrow
        ''' </summary>
        ''' <param name="tipX">x-coordinate of the force tip</param>
        ''' <param name="tipY">y-coordinate of the force tip</param>
        ''' <param name="F">intensity</param>
        ''' <param name="angle">angle in degrees</param>
        ''' <param name="caption">force caption</param>
        Sub New(tipX As Double, tipY As Double, F As Double, angle As Double, caption As String)
            Me.New()

            tipPoint.x = tipX
            tipPoint.y = tipY
            tipPoint.coordinateReference = Reference.world

            Me.F = F
            Me.angle = angle

            Me.caption = caption
        End Sub

        Sub New(tipX As Double, tipY As Double, F As Double, angle As Double, caption As String, maxF As Double)
            Me.New(tipX, tipY, F, angle, caption)
            Me.maximumForce = maxF
        End Sub

        ''' <summary>
        ''' Caption (replacements: {0} by F, {1} by angle, {2} by Fx, {3} by Fy)
        ''' </summary>
        ''' <returns></returns>
        Public Property caption As String
            Get
                Return p_caption
            End Get
            Set(value As String)
                p_caption = value
            End Set
        End Property
        ''' <summary>
        ''' Force tip point
        ''' </summary>
        ''' <returns></returns>
        Public Property tipPoint As Point
            Get
                Return p_tipPoint
            End Get
            Set(value As Point)
                p_tipPoint = value
            End Set
        End Property
        '
        ' force in x-direction
        '
        Public Property Fx() As Double
            Get
                If p_inputFxFy Then
                    Return p_Fx
                Else
                    Return F * Math.Cos(myMath.radians(angle))
                End If
            End Get
            Set(value As Double)
                p_Fx = value
                p_inputFxFy = True
                p_inputFAlpha = False
            End Set
        End Property
        '
        ' force in y-direction
        '
        Public Property Fy() As Double
            Get
                If p_inputFxFy Then
                    Return p_Fy
                Else
                    Return F * Math.Sin(myMath.radians(angle))
                End If
            End Get
            Set(value As Double)
                p_Fy = value
                p_inputFxFy = True
                p_inputFAlpha = False
            End Set
        End Property
        ''' <summary>
        ''' Resulting force (in angle direction)
        ''' </summary>
        ''' <returns></returns>
        Public Property F() As Double
            Get
                If p_inputFAlpha Then
                    Return p_F
                Else
                    Return Math.Sqrt(Fx ^ 2 + Fy ^ 2)
                End If
            End Get
            Set(value As Double)
                p_F = value
                p_inputFxFy = False
                p_inputFAlpha = True
            End Set
        End Property
        ''' <summary>
        ''' Angle in degrees
        ''' </summary>
        ''' <returns></returns>
        Public Property angle() As Double
            Get
                If p_inputFAlpha Then
                    Return p_angle
                Else
                    Return myMath.degrees(Math.Atan2(Fy, Fx))
                End If
            End Get
            Set(value As Double)
                p_angle = value
                p_inputFxFy = False
                p_inputFAlpha = True
            End Set
        End Property
        ''' <summary>
        ''' Maximum Force
        ''' </summary>
        ''' <returns></returns>
        Public Property maximumForce() As Double
            Get
                If Not p_isMaximumForceDefined Then
                    Return 1
                Else
                    Return p_maximumForce
                End If
            End Get
            Set(value As Double)
                p_maximumForce = value
                p_isMaximumForceDefined = True
            End Set
        End Property
        ''' <summary>
        ''' maximum Force size
        ''' </summary>
        ''' <returns></returns>
        Public Property MaximumSizeMM() As Double
            Get
                Return p_maxSize.max
            End Get
            Set(value As Double)
                p_maxSize.width = value
                p_maxSize.height = value
            End Set
        End Property
        ''' <summary>
        ''' maximum Force size 
        ''' </summary>
        ''' <returns></returns>
        Public Property MaximumSize() As size
            Get
                Return p_maxSize
            End Get
            Set(value As size)
                p_maxSize = value
            End Set
        End Property
        ''' <summary>
        ''' Force Type
        ''' </summary>
        ''' <returns></returns>
        Public Property ForceType As forceTypes
            Get
                Return p_ForceType
            End Get
            Set(value As forceTypes)
                p_ForceType = value
            End Set
        End Property
        ''' <summary>
        ''' Arrow Tip Type
        ''' </summary>
        ''' <returns></returns>
        Public Property ArrowTipType As arrowTipTypes
            Get
                Return p_ArrowTipType
            End Get
            Set(value As arrowTipTypes)
                p_ArrowTipType = value
            End Set
        End Property
        ''' <summary>
        ''' Force Direction
        ''' </summary>
        ''' <returns></returns>
        Public Property ForceDirection As forceDirections
            Get
                Return p_forceDirection
            End Get
            Set(value As forceDirections)
                p_forceDirection = value
            End Set
        End Property
        ''' <summary>
        ''' EndPoint Tip
        ''' </summary>
        ''' <returns></returns>
        Public Property endPointDisplay As WMDraw.PointDisplay
            Get
                Return p_endPointDisplay
            End Get
            Set(value As WMDraw.PointDisplay)
                p_endPointDisplay = value
            End Set
        End Property
        ''' <summary>
        ''' Arrow Size
        ''' </summary>
        ''' <returns></returns>
        Public Property arrowSize As size
            Get
                Return p_arrowSize
            End Get
            Set(value As size)
                p_arrowSize = value
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

        Private Function boundingRectangle(Optional ref As Reference = Reference.world) As Double() Implements Drawable.boundingRectangle
            Return Me.p_tipPoint.boundingRectangle(ref)
        End Function

        ''' <summary>
        ''' place on canvas
        ''' </summary>
        ''' <param name="contextobject"></param>
        ''' <param name="contextCoordinatesDelegate"></param>
        Public Overridable Sub draw(contextobject As ContextObject, contextCoordinatesDelegate As Drawable.contextCoordinates, Optional contextSizeDelegate As Drawable.contextSize = Nothing) Implements Drawable.draw

            Dim pTip As New Point         ' tip point
            Dim pEnd As New Point         ' end point
            Dim pText As New Point        ' text point

            Dim fSize As New size         ' size
            Dim arrowSize1 As New size    ' arrow
            Dim t As New size             ' pen Thickness

            If TypeOf contextobject.Item Is System.Windows.Controls.Canvas Then
                Dim myCanvas As New System.Windows.Controls.Canvas
                myCanvas = TryCast(contextobject.Item, System.Windows.Controls.Canvas)

                With myCanvas
                    '
                    ' Force
                    '
                    pTip = contextCoordinatesDelegate(Me.p_tipPoint)
                    fSize = contextSizeDelegate(Me.p_maxSize)
                    arrowSize1 = contextSizeDelegate(Me.arrowSize)

                    t = contextSizeDelegate(Me.pen.size)
                    If t.width < 1 Then t.width = 1

                    Dim myPolyline As New System.Windows.Shapes.Polyline

                    With myPolyline
                        '
                        ' shaft
                        '

                        '
                        ' draw in x-position
                        '
                        If p_isMaximumForceDefined Then
                            pEnd.x = pTip.x + fSize.width * F / maximumForce
                            pEnd.y = pTip.y
                        Else
                            pEnd.x = pTip.x + fSize.width * F / F
                            pEnd.y = pTip.y
                        End If

                        Select Case ForceType
                            Case forceTypes.resistanceForce
                                .Points.Add(New System.Windows.Point(pTip.x, pTip.y + arrowSize1.width / 5))
                                .Points.Add(New System.Windows.Point(pEnd.x, pEnd.y + arrowSize1.width / 5))

                                .Points.Add(New System.Windows.Point(pEnd.x, pEnd.y - arrowSize1.width / 5))
                                .Points.Add(New System.Windows.Point(pTip.x, pTip.y - arrowSize1.width / 5))
                            Case Else
                                .Points.Add(New System.Windows.Point(pTip.x, pTip.y))
                                .Points.Add(New System.Windows.Point(pEnd.x, pEnd.y))
                        End Select

                        ' rotate
                        Dim myRotateTransform As New Windows.Media.RotateTransform

                        myRotateTransform.CenterX = pTip.x
                        myRotateTransform.CenterY = pTip.y

                        Select Case ForceDirection
                            Case forceDirections.ArrowAtTip
                                myRotateTransform.Angle = -Me.angle - 180
                            Case Else
                                myRotateTransform.Angle = -Me.angle
                        End Select

                        myPolyline.RenderTransform = myRotateTransform

                        Dim pText1 As Windows.Point
                        pText1 = myRotateTransform.Transform(New Windows.Point(pEnd.x, pEnd.y))
                        pText.x = pText1.X
                        pText.y = pText1.Y

                        .Stroke = Me.pen.stroke
                        .StrokeDashArray = Me.pen.dashArray
                        .StrokeThickness = t.width
                    End With

                    myCanvas.Children.Add(myPolyline)

                    '
                    ' Arrow Tip
                    '
                    myPolyline = New System.Windows.Shapes.Polyline

                    With myPolyline
                        Select Case ForceDirection
                            Case forceDirections.ArrowAtTip

                                Select Case ArrowTipType
                                    Case arrowTipTypes.defaultForceTip
                                        .Points.Add(New System.Windows.Point(pTip.x + arrowSize1.height, pTip.y + arrowSize1.width / 2))
                                        .Points.Add(New System.Windows.Point(pTip.x, pTip.y))
                                        .Points.Add(New System.Windows.Point(pTip.x + arrowSize1.height, pTip.y - arrowSize1.width / 2))
                                    Case arrowTipTypes.halfTipLeft
                                        .Points.Add(New System.Windows.Point(pTip.x + arrowSize1.height, pTip.y + arrowSize1.width / 2))
                                        .Points.Add(New System.Windows.Point(pTip.x, pTip.y))
                                    Case arrowTipTypes.halfTipRight
                                        .Points.Add(New System.Windows.Point(pTip.x, pTip.y))
                                        .Points.Add(New System.Windows.Point(pTip.x + arrowSize1.height, pTip.y - arrowSize1.width / 2))
                                    Case arrowTipTypes.directionIndicatorLeft
                                        .Points.Add(New System.Windows.Point(pTip.x + arrowSize1.height / 2, pTip.y))
                                        .Points.Add(New System.Windows.Point(pTip.x + arrowSize1.height / 2, pTip.y + arrowSize1.width / 2))
                                        .Points.Add(New System.Windows.Point(pTip.x, pTip.y + arrowSize1.width / 2))
                                        .Points.Add(New System.Windows.Point(pTip.x - arrowSize1.height / 2, pTip.y))
                                    Case arrowTipTypes.directionIndicatorRight
                                        .Points.Add(New System.Windows.Point(pTip.x - arrowSize1.height / 2, pTip.y))
                                        .Points.Add(New System.Windows.Point(pTip.x, pTip.y - arrowSize1.width / 2))
                                        .Points.Add(New System.Windows.Point(pTip.x + arrowSize1.height / 2, pTip.y - arrowSize1.width / 2))
                                        .Points.Add(New System.Windows.Point(pTip.x + arrowSize1.height / 2, pTip.y))
                                End Select

                                Dim myRotateTransform As New Windows.Media.RotateTransform

                                myRotateTransform.CenterX = pTip.x
                                myRotateTransform.CenterY = pTip.y
                                If Me.F < 0 Then
                                    myRotateTransform.Angle = -Me.angle
                                Else
                                    myRotateTransform.Angle = -Me.angle - 180
                                End If
                                myPolyline.RenderTransform = myRotateTransform
                            Case Else
                                If Me.F < 0 Then
                                    Select Case ArrowTipType
                                        Case arrowTipTypes.defaultForceTip
                                            .Points.Add(New System.Windows.Point(pEnd.x + arrowSize1.height, pEnd.y + arrowSize1.width / 2))
                                            .Points.Add(New System.Windows.Point(pEnd.x, pEnd.y))
                                            .Points.Add(New System.Windows.Point(pEnd.x + arrowSize1.height, pEnd.y - arrowSize1.width / 2))
                                        Case arrowTipTypes.halfTipLeft
                                            .Points.Add(New System.Windows.Point(pEnd.x + arrowSize1.height, pEnd.y + arrowSize1.width / 2))
                                            .Points.Add(New System.Windows.Point(pEnd.x, pEnd.y))
                                        Case arrowTipTypes.halfTipRight
                                            .Points.Add(New System.Windows.Point(pEnd.x, pEnd.y))
                                            .Points.Add(New System.Windows.Point(pEnd.x + arrowSize1.height, pEnd.y - arrowSize1.width / 2))
                                        Case arrowTipTypes.directionIndicatorLeft
                                            .Points.Add(New System.Windows.Point(pEnd.x + arrowSize1.height / 2, pEnd.y))
                                            .Points.Add(New System.Windows.Point(pEnd.x + arrowSize1.height / 2, pEnd.y + arrowSize1.width / 2))
                                            .Points.Add(New System.Windows.Point(pEnd.x, pEnd.y + arrowSize1.width / 2))
                                            .Points.Add(New System.Windows.Point(pEnd.x - arrowSize1.height / 2, pEnd.y))
                                        Case arrowTipTypes.directionIndicatorRight
                                            .Points.Add(New System.Windows.Point(pEnd.x - arrowSize1.height / 2, pEnd.y))
                                            .Points.Add(New System.Windows.Point(pEnd.x, pEnd.y - arrowSize1.width / 2))
                                            .Points.Add(New System.Windows.Point(pEnd.x + arrowSize1.height / 2, pEnd.y - arrowSize1.width / 2))
                                            .Points.Add(New System.Windows.Point(pEnd.x + arrowSize1.height / 2, pEnd.y))
                                    End Select
                                Else
                                    Select Case ArrowTipType
                                        Case arrowTipTypes.defaultForceTip
                                            .Points.Add(New System.Windows.Point(pEnd.x - arrowSize1.height, pEnd.y + arrowSize1.width / 2))
                                            .Points.Add(New System.Windows.Point(pEnd.x, pTip.y))
                                            .Points.Add(New System.Windows.Point(pEnd.x - arrowSize1.height, pEnd.y - arrowSize1.width / 2))
                                        Case arrowTipTypes.halfTipLeft
                                            .Points.Add(New System.Windows.Point(pEnd.x - arrowSize1.height, pEnd.y + arrowSize1.width / 2))
                                            .Points.Add(New System.Windows.Point(pEnd.x, pTip.y))
                                        Case arrowTipTypes.halfTipRight
                                            .Points.Add(New System.Windows.Point(pEnd.x, pTip.y))
                                            .Points.Add(New System.Windows.Point(pEnd.x - arrowSize1.height, pEnd.y - arrowSize1.width / 2))
                                        Case arrowTipTypes.directionIndicatorLeft
                                            .Points.Add(New System.Windows.Point(pEnd.x + arrowSize1.height, pEnd.y + arrowSize1.width / 2))
                                            .Points.Add(New System.Windows.Point(pEnd.x + arrowSize1.height / 2, pEnd.y + arrowSize1.width / 2))
                                            .Points.Add(New System.Windows.Point(pEnd.x, pEnd.y))
                                        Case arrowTipTypes.directionIndicatorRight
                                            .Points.Add(New System.Windows.Point(pEnd.x, pEnd.y))
                                            .Points.Add(New System.Windows.Point(pEnd.x + arrowSize1.height / 2, pEnd.y - arrowSize1.width / 2))
                                            .Points.Add(New System.Windows.Point(pEnd.x + arrowSize1.height, pEnd.y - arrowSize1.width / 2))
                                    End Select

                                End If


                                Dim myRotateTransform As New Windows.Media.RotateTransform


                                If Me.F < 0 Then
                                    myRotateTransform.CenterX = pTip.x
                                    myRotateTransform.CenterY = pTip.y
                                    myRotateTransform.Angle = -Me.angle '- 180
                                Else
                                    myRotateTransform.CenterX = pTip.x
                                    myRotateTransform.CenterY = pTip.y
                                    myRotateTransform.Angle = -Me.angle
                                End If
                                myPolyline.RenderTransform = myRotateTransform
                        End Select

                        '.Stroke = System.Windows.Media.Brushes.Red
                        .Stroke = Me.pen.stroke
                        .StrokeDashArray = Me.pen.dashArray
                        .StrokeThickness = t.width

                    End With

                    myCanvas.Children.Add(myPolyline)

                    ' end point
                    pText.display = endPointDisplay
                    pText.displaySize = arrowSize
                    pText.draw(contextobject, contextCoordinatesDelegate, contextSizeDelegate)

                    '
                    ' Text
                    '
                    If p_caption <> "" Then
                        Dim tb As New Windows.Controls.TextBlock
                        Dim myCaption As String
                        myCaption = caption

                        ' replaces {0} by F, {1} by angle, {2} by Fx, {3} by Fy
                        myCaption = String.Format(myCaption, F, angle, Fx, Fy)

                        With tb
                            .Text = myCaption
                        End With

                        'Dim myRotateTransform As New Windows.Media.RotateTransform
                        'myRotateTransform.Angle = (startAngle + endAngle) / 2
                        'tb.RenderTransform = myRotateTransform

                        '
                        ' calculate coordinates in local system
                        ' by a call-back to the calling drawing object
                        '

                        myCanvas.Children.Add(tb)
                        myCanvas.SetLeft(tb, pText.x)
                        myCanvas.SetTop(tb, pText.y)
                    End If

                End With
            Else
                Throw New NotImplementedException()
            End If

        End Sub




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

    ''' <summary>
    ''' Moment Arc as a drawable item
    ''' </summary>
    <ComVisible(False)>
    Public Class MomentArc
        Implements Drawable
        '
        ' todo: Create a force or arrow drawable and a Moment drawable
        '
        Const DEFAULT_ARROW_WIDTH_MM As Double = 3
        Const DEFAULT_ARROW_HEIGHT_MM As Double = 5

        Const DEFAULT_RADIUS_MM As Double = 8

        Const DEFAULT_START_ANGLE As Double = 30
        Const DEFAULT_END_ANGLE As Double = 150

        Private p_pen As pen
        Private p_zIndex As Long

        Private p_midPoint As New Point
        Private p_Radius As New size
        Private p_StartAngle As Double
        Private p_EndAngle As Double

        Private p_caption As String

        ''' <summary>
        ''' Arrow Size
        ''' </summary>
        Private p_arrowSize As New size

        Sub New()
            Me.midPoint = New Point
            Me.p_StartAngle = DEFAULT_START_ANGLE
            Me.p_EndAngle = DEFAULT_END_ANGLE

            Me.Radius = DEFAULT_RADIUS_MM
            Me.RadiusReference = Reference.contextMillimeters

            Me.p_arrowSize.width = DEFAULT_ARROW_WIDTH_MM
            Me.p_arrowSize.height = DEFAULT_ARROW_WIDTH_MM
            Me.p_arrowSize.Reference = Reference.contextMillimeters
        End Sub

        ''' <summary>
        ''' Moment
        ''' </summary>
        Sub New(x As Double, y As Double, caption As String)
            Me.New

            Me.midPoint.x = x
            Me.midPoint.y = y
            Me.midPoint.coordinateReference = Reference.world

            Me.caption = caption
            Me.startAngle = DEFAULT_START_ANGLE
            Me.endAngle = DEFAULT_END_ANGLE
        End Sub

        ''' <summary>
        ''' Moment
        ''' </summary>
        Sub New(x As Double, y As Double, caption As String, startAngle As Double, endAngle As Double)

            Me.midPoint.x = x
            Me.midPoint.y = y
            Me.midPoint.coordinateReference = Reference.world
            Me.Radius = DEFAULT_RADIUS_MM
            Me.RadiusReference = Reference.contextMillimeters
            Me.caption = caption
            Me.startAngle = startAngle
            Me.endAngle = endAngle

            Me.p_arrowSize.width = DEFAULT_ARROW_WIDTH_MM
            Me.p_arrowSize.height = DEFAULT_ARROW_WIDTH_MM
            Me.p_arrowSize.Reference = Reference.contextMillimeters

        End Sub

        ''' <summary>
        ''' Moment
        ''' </summary>
        Sub New(x As Double, y As Double, coordinateReference As Reference, caption As String, radius As Double, radiusReference As Reference, startAngle As Double, endAngle As Double)
            Me.midPoint.x = x
            Me.midPoint.y = y
            Me.midPoint.coordinateReference = coordinateReference
            Me.Radius = radius
            Me.RadiusReference = radiusReference
            Me.caption = caption
            Me.startAngle = startAngle
            Me.endAngle = endAngle

            Me.p_arrowSize.width = DEFAULT_ARROW_WIDTH_MM
            Me.p_arrowSize.height = DEFAULT_ARROW_WIDTH_MM
            Me.p_arrowSize.Reference = Reference.contextMillimeters

        End Sub

        ''' <summary>
        ''' place on canvas
        ''' </summary>
        ''' <param name="contextobject"></param>
        ''' <param name="contextCoordinatesDelegate"></param>
        Public Overridable Sub draw(contextobject As ContextObject, contextCoordinatesDelegate As Drawable.contextCoordinates, Optional contextSizeDelegate As Drawable.contextSize = Nothing) Implements Drawable.draw

            Dim pm As New Point         ' arc mid point
            Dim p1 As New Point         ' arc start point
            Dim p2 As New Point         ' arc end point
            Dim r1 As New size          ' radius
            Dim arrowSize1 As New size  ' arrow
            Dim t As New size           ' pen Thickness

            If TypeOf contextobject.Item Is System.Windows.Controls.Canvas Then
                Dim myCanvas As New System.Windows.Controls.Canvas
                myCanvas = TryCast(contextobject.Item, System.Windows.Controls.Canvas)

                With myCanvas
                    '
                    ' Arc
                    '
                    r1 = contextSizeDelegate(Me.p_Radius)
                    pm = contextCoordinatesDelegate(Me.midPoint)
                    arrowSize1 = contextSizeDelegate(Me.p_arrowSize)

                    t = contextSizeDelegate(Me.pen.size)
                    If t.width < 1 Then t.width = 1

                    p1.x = pm.x + r1.max * Math.Cos(myMath.radians(Me.startAngle))
                    p1.y = pm.y - r1.max * Math.Sin(myMath.radians(Me.startAngle))

                    p2.x = pm.x + r1.max * Math.Cos(myMath.radians(Me.endAngle))
                    p2.y = pm.y - r1.max * Math.Sin(myMath.radians(Me.endAngle))

                    '
                    ' https://stackoverflow.com/questions/21957779/arc-segment-between-two-points-and-a-radius
                    '
                    Dim pthFigure As PathFigure = New PathFigure()
                    Dim arcSeg As ArcSegment = New ArcSegment()
                    Dim myPathSegmentCollection As PathSegmentCollection = New PathSegmentCollection()

                    ' startpoint of pathFigure
                    ' P1 (start  Point)
                    pthFigure.StartPoint = New System.Windows.Point(p1.x, p1.y)

                    With arcSeg
                        ' arc starts with last point of pathFigure
                        ' P2 (end Point)
                        .Point = New System.Windows.Point(p2.x, p2.y)
                        ' Radius
                        .Size = New Windows.Size(r1.width, r1.height)
                        ' clarify dual solutions for a circle defined 
                        ' by two points And a radius
                        .IsLargeArc = Math.Abs(endAngle - startAngle) > 180
                        .SweepDirection = SweepDirection.Counterclockwise
                    End With

                    ' add arc to segment collection
                    myPathSegmentCollection.Add(arcSeg)
                    ' add segments to path figure
                    pthFigure.Segments = myPathSegmentCollection

                    Dim pthFigureCollection As PathFigureCollection = New PathFigureCollection()
                    pthFigureCollection.Add(pthFigure)

                    Dim pthGeometry As PathGeometry = New PathGeometry()
                    pthGeometry.Figures = pthFigureCollection

                    Dim arcPath As Windows.Shapes.Path = New Windows.Shapes.Path
                    arcPath.Data = pthGeometry

                    arcPath.Stroke = Me.pen.stroke
                    arcPath.StrokeDashArray = Me.pen.dashArray
                    arcPath.StrokeThickness = t.width

                    myCanvas.Children.Add(arcPath)

                    '
                    ' Arrow Tip
                    '
                    Dim myPolyline As New System.Windows.Shapes.Polyline

                    With myPolyline
                        .Points.Add(New System.Windows.Point(pm.x + r1.max + arrowSize1.width / 2, pm.y + arrowSize1.height))
                        .Points.Add(New System.Windows.Point(pm.x + r1.max, pm.y))
                        .Points.Add(New System.Windows.Point(pm.x + r1.max - arrowSize1.width / 2, pm.y + arrowSize1.height))

                        Dim myRotateTransform As New Windows.Media.RotateTransform

                        myRotateTransform.CenterX = pm.x
                        myRotateTransform.CenterY = pm.y
                        myRotateTransform.Angle = -endAngle
                        myPolyline.RenderTransform = myRotateTransform

                        .Stroke = Me.pen.stroke
                        .StrokeDashArray = Me.pen.dashArray
                        .StrokeThickness = t.width

                    End With

                    myCanvas.Children.Add(myPolyline)

                    '
                    ' Text
                    '
                    If p_caption <> "" Then
                        Dim tb As New Windows.Controls.TextBlock

                        With tb
                            .Text = p_caption
                        End With

                        'Dim myRotateTransform As New Windows.Media.RotateTransform
                        'myRotateTransform.Angle = (startAngle + endAngle) / 2
                        'tb.RenderTransform = myRotateTransform

                        Dim p3 As New Point
                        p3.x = pm.x + r1.max * Math.Cos(myMath.radians((Me.startAngle + Me.endAngle) / 2))
                        p3.y = pm.y - r1.max * Math.Sin(myMath.radians((Me.startAngle + Me.endAngle) / 2))

                        '
                        ' calculate coordinates in local system
                        ' by a call-back to the calling drawing object
                        '

                        myCanvas.Children.Add(tb)
                        myCanvas.SetLeft(tb, p3.x)
                        myCanvas.SetTop(tb, p3.y)
                    End If

                End With
            Else
                Throw New NotImplementedException()
            End If

        End Sub

        Private Function boundingRectangle(Optional ref As Reference = Reference.world) As Double() Implements Drawable.boundingRectangle

            Dim r(3) As Double

            Dim p1 As New Point
            Dim p2 As New Point

            Dim p(2) As Point

            Dim alpha(2) As Double

            Dim Q(1) As Byte
            alpha(0) = startAngle
            alpha(1) = endAngle
            '
            ' Quadrants
            '
            For i As Integer = 0 To 1
                Q(i) = CByte(Math.Floor(alpha(i) / 90) Mod 4)
            Next
            '
            ' sort Quadrants
            '
            If Q(0) > Q(1) Then
                Dim tmp As Byte
                tmp = Q(1)
                Q(1) = Q(0)
                Q(0) = tmp
            End If

            alpha(0) = Me.startAngle
            alpha(1) = Me.endAngle

            If Q(0) <> Q(1) Then
                alpha(2) = (Q(1) - 1) * 90
            Else
                alpha(2) = (alpha(0) + alpha(1)) / 2
            End If
            '
            ' todo: size has to be converted if not world coords
            '
            If midPoint.coordinateReference = ref And p_Radius.Reference = ref Then
                For i As Integer = 0 To 2
                    p(i) = New Point
                    p(i).x = midPoint.x + p_Radius.max * Math.Cos(myMath.radians(alpha(i)))
                    p(i).y = midPoint.y + p_Radius.max * Math.Sin(myMath.radians(alpha(i)))
                    myMath.includeInBoundingRectangle(r, p(i))
                Next
                Return r
            Else
                Return Me.midPoint.boundingRectangle(ref)
            End If


        End Function
        ''' <summary>
        ''' Caption
        ''' </summary>
        ''' <returns></returns>
        Public Property caption As String
            Get
                Return p_caption
            End Get
            Set(value As String)
                p_caption = value
            End Set
        End Property
        ''' <summary>
        ''' Mid-Point
        ''' </summary>
        ''' <returns></returns>
        Public Property midPoint As Point
            Get
                Return p_midPoint
            End Get
            Set(value As Point)
                p_midPoint = value
            End Set
        End Property

        ''' <summary>
        ''' Radius
        ''' </summary>
        ''' <returns></returns>
        Public Property Radius As Double
            Get
                Return p_Radius.max
            End Get
            Set(value As Double)
                p_Radius.width = value
                p_Radius.height = value
            End Set
        End Property
        ''' <summary>
        ''' Radius Reference
        ''' </summary>
        ''' <returns></returns>
        Public Property RadiusReference As Reference
            Get
                Return p_Radius.Reference
            End Get
            Set(value As Reference)
                p_Radius.Reference = value
            End Set
        End Property
        ''' <summary>
        ''' Start Angle
        ''' </summary>
        ''' <returns></returns>
        Public Property startAngle As Double
            Get
                Return p_StartAngle
            End Get
            Set(value As Double)
                p_StartAngle = value
                If p_EndAngle < p_StartAngle Then
                    Dim tmp As Double
                    tmp = p_StartAngle
                    p_EndAngle = p_StartAngle
                    p_StartAngle = tmp
                End If

            End Set
        End Property
        ''' <summary>
        ''' End Angle
        ''' </summary>
        ''' <returns></returns>
        Public Property endAngle As Double
            Get
                Return p_EndAngle
            End Get
            Set(value As Double)
                p_EndAngle = value
                If p_StartAngle > p_EndAngle Then
                    Dim tmp As Double
                    tmp = p_StartAngle
                    p_StartAngle = p_EndAngle
                    p_EndAngle = tmp
                End If
            End Set
        End Property
        ''' <summary>
        ''' Arrow Size
        ''' </summary>
        ''' <returns></returns>
        Public Property arrowSize As size
            Get
                Return p_arrowSize
            End Get
            Set(value As size)
                p_arrowSize = value
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
#Region "Line Class"
    ''' <summary>
    ''' as a drawable item
    ''' </summary>
    <ComVisible(False)>
    Public Class Line
        Implements Drawable
        '
        ' todo: Create a force or arrow drawable and a Moment drawable
        '
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
        ''' <param name="endPoint"></param>
        Sub New(startPoint As Point, endPoint As Point)
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

#If DEBUG Then
                        Debug.Print("Line (" & String.Format("({0:F2}/{1:F2}) - ({2:F2}/{3:F2})", p1.x, p1.y, p2.x, p2.y))
#End If
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
#Region "Line Class"
    ''' <summary>
    ''' as a drawable item
    ''' </summary>
    <ComVisible(False)>
    Public Class Rectangle
        Implements Drawable
        '
        ' todo: Create a force or arrow drawable and a Moment drawable
        '
        Private p_pen As pen
        Private p_fill As fill

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
            p_fill = New fill
            p_pen = New pen
        End Sub

        ''' <summary>
        ''' Line from start and end points in world coordinates
        ''' </summary>
        ''' <param name="startPoint"></param>
        ''' <param name="endPoint"></param>
        Sub New(startPoint As Point, endPoint As Point)
            Me.New()
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
            Me.New()
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
            Me.New()
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
                    Dim rect1 As New System.Windows.Shapes.Rectangle
                    With rect1
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

#If DEBUG Then
                        Debug.Print("Rectangle (" & String.Format("({0:F2}/{1:F2}) - ({2:F2}/{3:F2})", p1.x, p1.y, p2.x, p2.y))
#End If
                        .Width = Math.Abs(p2.x - p1.x)
                        .Height = Math.Abs(p2.y - p1.y)

                        'todo: offer fill solid and hatch

                        .Stroke = Nothing 'Me.pen.stroke
                        .StrokeThickness = t.width
                        .StrokeDashArray = Me.pen.dashArray

                        .Fill = fill.Brush()

                        myCanvas.Children.Add(rect1)
                        Windows.Controls.Canvas.SetLeft(rect1, Math.Min(p1.x, p2.x))
                        Windows.Controls.Canvas.SetTop(rect1, Math.Min(p1.y, p2.y))
                    End With


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

        Private ReadOnly Property width()
            Get
                Return Math.Abs(endPoint.x - startPoint.x)
            End Get
        End Property

        Private ReadOnly Property height()
            Get
                Return Math.Abs(endPoint.y - startPoint.y)
            End Get
        End Property

        Private ReadOnly Property top()
            Get
                Return Math.Min(startPoint.y, endPoint.y)
            End Get
        End Property

        Private ReadOnly Property left()
            Get
                Return Math.Min(startPoint.x, endPoint.x)
            End Get
        End Property

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

        Public Property fill As fill
            Get
                Return p_fill
            End Get
            Set(value As fill)
                p_fill = value
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

    Public Class Screw
        Implements Drawable

        Private p_HeadPosition As New Point
        Private p_l As Double
        Private p_lg As Double
        Private p_angle As Double
        Private p_caption As String


        Private p_zIndex As Long

        Private p_pen As pen

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

        Public Sub draw(contextobject As ContextObject, contextCoordinatesDelegate As Drawable.contextCoordinates, Optional contextSizeDelegate As Drawable.contextSize = Nothing) Implements Drawable.draw
            Throw New NotImplementedException()
        End Sub

        Public Function boundingRectangle(Optional ref As Reference = Reference.world) As Double() Implements Drawable.boundingRectangle
            Throw New NotImplementedException()
        End Function

        Public Function CompareTo(other As Drawable) As Integer Implements IComparable(Of Drawable).CompareTo
            Throw New NotImplementedException()
        End Function
    End Class
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
        ''' <summary>
        ''' angle in degrees
        ''' </summary>
        ''' <returns></returns>
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
                        End With

                        Dim myRotateTransform As New Windows.Media.RotateTransform
                        myRotateTransform.Angle = angle
                        tb.RenderTransform = myRotateTransform

                        'Dim p As New Point

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
        Private p_pen As pen
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
            Dim r(3) As Double
            r(0) = Me.midPoint.x - Me.radius
            r(1) = Me.midPoint.y - Me.radius
            r(2) = Me.midPoint.x + Me.radius
            r(3) = Me.midPoint.y + Me.radius
            Return r
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
                Return p_pen
            End Get
            Set(value As pen)
                p_pen = value
            End Set
        End Property
    End Class

    <ComVisible(False)>
    Public Class pen

        ' todo: allow transparency

        'Public stroke As System.Windows.Media.Brush

        Private p_color As System.Windows.Media.Color
        'Private p_thicknessReference As Reference
        Private p_thickness As size
        Private p_dasharray As System.Windows.Media.DoubleCollection
        Private p_opacity As Double

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
            p_opacity = 1
            Me.dashString = DASHARRAY_SOLID
        End Sub
        ''' <summary>
        ''' Get the brush object with defined color
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property stroke As System.Windows.Media.Brush
            Get
                stroke = New System.Windows.Media.SolidColorBrush(Me.color)
                stroke.Opacity = p_opacity

                Return stroke
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
        Public Property opacity As Double
            Get
                Return p_opacity
            End Get
            Set(value As Double)
                p_opacity = value
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

        Public Enum fillType
            solidColor
            linearHatching
        End Enum

        Private p_filltype As fillType

        Private p_color As System.Windows.Media.Color
        Private p_thickness As Double
        Private p_spacing As Double

        Private p_angle As Double
        Private p_opacity As Double

        Public Sub New()
            p_color = Colors.Black
            p_thickness = 1

            p_opacity = 1
            p_spacing = 5

            p_angle = 45
        End Sub

        Public Sub SetSolidColor(color As Color)
            p_filltype = fillType.solidColor
            p_color = color
        End Sub

        Public Sub setLinearHatch(color As Color, Optional opacity As Double = 1,
                                  Optional hatchThickness_mm As Double = 2,
                                  Optional hatchSpacing_mm As Double = 6,
                                  Optional hatchAngle As Double = 45)
            p_filltype = fillType.linearHatching
            p_color = color

            p_thickness = hatchThickness_mm
            p_spacing = hatchSpacing_mm

            p_opacity = opacity
            p_angle = hatchAngle

        End Sub
        Public Property color() As Color
            Get
                Return p_color
            End Get
            Set(value As Color)
                p_color = value
            End Set
        End Property

        Public ReadOnly Property Brush() As Brush
            Get
                Select Case Me.p_filltype
                    Case fillType.solidColor
                        Dim myBrush As New SolidColorBrush
                        myBrush.Color = Me.p_color
                        myBrush.Opacity = Me.p_opacity
                        Return myBrush
                    Case fillType.linearHatching
                        Dim FillColor As Color
                        Dim HatchThickness As Double
                        Dim HatchDistance As Double
                        Dim HatchAngle As Double

                        FillColor = p_color

                        HatchThickness = p_thickness
                        HatchAngle = p_angle
                        HatchDistance = p_spacing

                        '
                        ' https://stackoverflow.com/questions/42667566/creating-diagonal-pattern-in-wpf
                        ' and
                        ' https://docs.microsoft.com/en-us/dotnet/framework/wpf/graphics-multimedia/wpf-brushes-overview#paint-with-a-drawing
                        '
                        Dim myBrush As New DrawingBrush()
                        Dim myGeometryGroup As New GeometryGroup()

                        Dim TileSize As Double
                        TileSize = HatchThickness + HatchDistance

                        '
                        ' add a horizontal line to the geometry group
                        '
                        myGeometryGroup.Children.Add(New LineGeometry(New Windows.Point(0, HatchThickness / 2), New Windows.Point(TileSize, HatchThickness / 2)))

                        '
                        ' draw geometry with transparent brush and pen as defined
                        '
                        Dim p As New Windows.Media.Pen
                        p.Brush = New SolidColorBrush(FillColor)
                        p.Brush.Opacity = p_opacity
                        p.Thickness = HatchThickness
                        p.StartLineCap = PenLineCap.Square
                        p.EndLineCap = PenLineCap.Square

                        Dim myDrawing As New GeometryDrawing(Nothing, p, myGeometryGroup)

                        '
                        ' apply the drawing to the brush
                        '
                        myBrush.Drawing = myDrawing

                        '
                        ' in case, there is more than one line use a Drawing Group
                        '

                        'Dim myDrawingGroup As New DrawingGroup()
                        'myDrawingGroup.Children.Add(checkers)
                        'myBrush.Drawing = myDrawingGroup

                        ' set viewbox and viewport
                        myBrush.Viewbox = New Windows.Rect(0, 0, TileSize, TileSize)
                        myBrush.ViewboxUnits = BrushMappingMode.Absolute
                        myBrush.Viewport = New Windows.Rect(0, 0, TileSize, TileSize)
                        myBrush.ViewportUnits = BrushMappingMode.Absolute
                        myBrush.TileMode = TileMode.Tile
                        myBrush.Stretch = Stretch.UniformToFill
                        ' rotate
                        myBrush.Transform = New RotateTransform(HatchAngle)

                        Return myBrush
                    Case Else
                        Return Nothing
                End Select
            End Get
        End Property



    End Class

End Namespace
