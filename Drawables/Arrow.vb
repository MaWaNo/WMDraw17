Imports System.Runtime.InteropServices

Namespace WMDraw

    <ComVisible(False)>
    Public Class Arrow
        Implements Drawable

        Const DEFAULT_ARROW_HEIGHT_MM As Double = 2.5
        Const DEFAULT_ARROW_WIDTH_MM As Double = 1
        Const DEFAULT_FORCE_SIZE_MM As Double = 10

        Private p_caption As String

        Private p_endArrowSize As New size

        Private p_endArrowTipType As arrowTipTypes

        Private p_endArrowVisible As Boolean

        ''' <summary>
        ''' End Point
        ''' </summary>
        Private p_endPoint As Point

        Private p_pen As pen

        Private p_startArrowSize As New size

        Private p_startArrowTipType As arrowTipTypes

        Private p_startArrowVisible As Boolean

        ''' <summary>
        ''' Start Point
        ''' </summary>
        Private p_startPoint As Point
        Private p_zIndex As Long

        Sub New()
            ' default arrow size
            p_startArrowVisible = False
            With Me.p_startArrowSize
                .width = DEFAULT_ARROW_WIDTH_MM
                .height = DEFAULT_ARROW_WIDTH_MM
                .Reference = Reference.contextMillimeters
            End With
            Me.p_startArrowTipType = arrowTipTypes.defaultForceTip

            p_endArrowVisible = True
            With Me.p_endArrowSize
                .width = DEFAULT_ARROW_WIDTH_MM
                .height = DEFAULT_ARROW_WIDTH_MM
                .Reference = Reference.contextMillimeters
            End With
            Me.p_endArrowTipType = arrowTipTypes.defaultForceTip
        End Sub

        ''' <summary>
        ''' New Arrow (Pointing form start to end by default)
        ''' </summary>
        ''' <param name="startX">x-coordinate of start point</param>
        ''' <param name="startY">y-coordinate of start point</param>
        ''' <param name="endX">x-coordinate of end point</param>
        ''' <param name="endY">y-coordinate of end point</param>
        ''' <param name="caption">caption</param>
        Sub New(startX As Double, startY As Double, endX As Double, endY As Double, caption As String)
            Me.New()

            StartPoint = New Point(startX, startY)
            EndPoint = New Point(endX, endY)
            Me.caption = caption
        End Sub

        ''' <summary>
        ''' New Arrow (Pointing form start to end by default)
        ''' </summary>
        ''' <param name="startPoint"></param>
        ''' <param name="endPoint"></param>
        ''' <param name="caption"></param>
        Sub New(startPoint As Point, endPoint As Point, caption As String)
            Me.New()

            Me.StartPoint = startPoint
            Me.EndPoint = endPoint
            Me.caption = caption
        End Sub

        ''' <summary>
        ''' New Arrow (Pointing form start to end by default)
        ''' </summary>
        ''' <param name="startX">x-coordinate of start point</param>
        ''' <param name="startY">y-coordinate of start point</param>
        ''' <param name="endX">x-coordinate of end point</param>
        ''' <param name="endY">y-coordinate of end point</param>
        Sub New(startX As Double, startY As Double, endX As Double, endY As Double)
            Me.New()

            StartPoint = New Point(startX, startY)
            EndPoint = New Point(endX, endY)
            Me.caption = ""

        End Sub

        ''' <summary>
        ''' New Arrow (Pointing form start to end by default)
        ''' </summary>
        ''' <param name="startPoint"></param>
        ''' <param name="endPoint"></param>
        Sub New(startPoint As Point, endPoint As Point)
            Me.New()

            Me.StartPoint = startPoint
            Me.EndPoint = endPoint
            Me.caption = ""
        End Sub

        Public Enum arrowTipTypes
            defaultForceTip
            halfTipLeft
            halfTipRight
            directionIndicatorLeft
            directionIndicatorRight
        End Enum

        ' orientation
        Public ReadOnly Property angle() As Double
            Get
                Return myMath.degrees(Math.Atan2((EndPoint.y - StartPoint.y), (EndPoint.x - StartPoint.x)))
            End Get
        End Property

        ''' <summary>
        ''' Caption (replacements: {0} by length, {1} by angle, {2} by length in x-direction, {3} by lenth in y-directon)
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
        ''' End Arrow Size
        ''' </summary>
        ''' <returns></returns>
        Public Property EndArrowSize As size
            Get
                Return p_endArrowSize
            End Get
            Set(value As size)
                p_endArrowSize = value
            End Set
        End Property

        ''' <summary>
        ''' End Arrow Tip Type
        ''' </summary>
        ''' <returns></returns>
        Public Property EndArrowTipType As arrowTipTypes
            Get
                Return p_endArrowTipType
            End Get
            Set(value As arrowTipTypes)
                p_endArrowTipType = value
            End Set
        End Property

        ''' <summary>
        ''' Start Arrow Visibility
        ''' </summary>
        ''' <returns></returns>
        Public Property EndArrowVisible As Boolean
            Get
                Return p_endArrowVisible
            End Get
            Set(value As Boolean)
                p_endArrowVisible = value
            End Set
        End Property

        ''' <summary>
        ''' Start Point
        ''' </summary>
        ''' <returns></returns>
        Public Property EndPoint() As Point
            Get
                Return p_endPoint
            End Get
            Set(value As Point)
                p_endPoint = value
            End Set
        End Property

        Public ReadOnly Property length() As Double
            Get
                Return Math.Sqrt(lengthX ^ 2 + lengthY ^ 2)
            End Get
        End Property

        Public ReadOnly Property lengthX() As Double
            Get
                Return (EndPoint.x - StartPoint.x)
            End Get
        End Property

        Public ReadOnly Property lengthY() As Double
            Get
                Return (EndPoint.y - StartPoint.y)
            End Get
        End Property

        Public Property pen As pen Implements Drawable.pen
            Get
                Return p_pen
            End Get
            Set(value As pen)
                p_pen = value
            End Set
        End Property

        ''' <summary>
        ''' Start Arrow Size
        ''' </summary>
        ''' <returns></returns>
        Public Property StartArrowSize As size
            Get
                Return p_startArrowSize
            End Get
            Set(value As size)
                p_startArrowSize = value
            End Set
        End Property
        ''' <summary>
        ''' Start Arrow Tip Type
        ''' </summary>
        ''' <returns></returns>
        Public Property StartArrowTipType As arrowTipTypes
            Get
                Return p_startArrowTipType
            End Get
            Set(value As arrowTipTypes)
                p_startArrowTipType = value
            End Set
        End Property
        ''' <summary>
        ''' Start Arrow Visibility
        ''' </summary>
        ''' <returns></returns>
        Public Property StartArrowVisible As Boolean
            Get
                Return p_startArrowVisible
            End Get
            Set(value As Boolean)
                p_startArrowVisible = value
            End Set
        End Property

        ''' <summary>
        ''' Start Point
        ''' </summary>
        ''' <returns></returns>
        Public Property StartPoint() As Point
            Get
                Return p_startPoint
            End Get
            Set(value As Point)
                p_startPoint = value
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

        Public Function boundingRectangle(Optional getWordSize As Drawable.estimateWorldSize = Nothing) As Double() Implements Drawable.boundingRectangle
            Dim r(3) As Double
            myMath.includeInBoundingRectangle(r, Me.p_startPoint)
            myMath.includeInBoundingRectangle(r, Me.p_endPoint)
            Return r
        End Function

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
        ''' place on canvas
        ''' </summary>
        ''' <param name="contextobject"></param>
        ''' <param name="contextCoordinatesDelegate"></param>
        Public Overridable Sub draw(contextobject As ContextObject, contextCoordinatesDelegate As Drawable.contextCoordinates, Optional contextSizeDelegate As Drawable.contextSize = Nothing) Implements Drawable.draw

            Dim pTip As New Point         ' tip point
            Dim pOffsetCU As New size       ' force offset in context size
            Dim pOffset As New Point

            Dim pEnd As New Point         ' end point
            Dim pText As New Point        ' text point

            Dim fSize As New size         ' size
            Dim t As New size             ' pen Thickness

            If TypeOf contextobject.Item Is System.Windows.Controls.Canvas Then
                Dim myCanvas As New System.Windows.Controls.Canvas
                myCanvas = TryCast(contextobject.Item, System.Windows.Controls.Canvas)

                With myCanvas
                    '
                    ' Force
                    '
                    Dim p1 As New Point
                    Dim p2 As New Point
                    p1 = contextCoordinatesDelegate(StartPoint)
                    p2 = contextCoordinatesDelegate(EndPoint)

                    t = contextSizeDelegate(Me.pen.size)
                    If t.width > 0 And t.width < 0.1 Then t.width = 0.1

                    Dim line1 As New System.Windows.Shapes.Line

                    With line1
                        .X1 = p1.x
                        .Y1 = p1.y
                        .X2 = p2.x
                        .Y2 = p2.y
                        .Stroke = Me.pen.stroke
                        .StrokeThickness = t.width
                        .StrokeDashArray = Me.pen.dashArray
                    End With

                    myCanvas.Children.Add(line1)

                    Dim arrowPoints As New List(Of Point)
                    Dim arrowSizes As New List(Of size)
                    Dim angles As New List(Of Double)
                    Dim atts As New List(Of arrowTipTypes)

                    If Me.StartArrowVisible Then
                        arrowPoints.Add(p1)
                        arrowSizes.Add(StartArrowSize)
                        atts.Add(StartArrowTipType)
                        angles.Add(Me.angle)
                    End If
                    If EndArrowVisible Then
                        arrowPoints.Add(p2)
                        arrowSizes.Add(EndArrowSize)
                        atts.Add(EndArrowTipType)
                        angles.Add(Me.angle - 180)
                    End If

                    '
                    ' Arrow at Start point
                    '
                    If arrowPoints.Count > 0 Then
                        For k As Integer = 0 To arrowPoints.Count - 1
                            Dim p As Point = arrowPoints(k)
                            Dim sz As size = arrowSizes(k)
                            Dim tt As arrowTipTypes = atts(k)
                            Dim angl As Double = angles(k)

                            Dim myPolyline As New System.Windows.Shapes.Polyline
                            Dim arrowSize1 As New size
                            arrowSize1 = contextSizeDelegate(sz)

                            With myPolyline

                                Select Case tt
                                    Case arrowTipTypes.defaultForceTip
                                        .Points.Add(New System.Windows.Point(p.x + arrowSize1.height, p.y + arrowSize1.width / 2))
                                        .Points.Add(New System.Windows.Point(p.x, p.y))
                                        .Points.Add(New System.Windows.Point(p.x + arrowSize1.height, p.y - arrowSize1.width / 2))
                                    Case arrowTipTypes.halfTipLeft
                                        .Points.Add(New System.Windows.Point(p.x + arrowSize1.height, p.y + arrowSize1.width / 2))
                                        .Points.Add(New System.Windows.Point(p.x, p.y))
                                    Case arrowTipTypes.halfTipRight
                                        .Points.Add(New System.Windows.Point(p.x, p.y))
                                        .Points.Add(New System.Windows.Point(p.x + arrowSize1.height, p.y - arrowSize1.width / 2))
                                    Case arrowTipTypes.directionIndicatorLeft
                                        .Points.Add(New System.Windows.Point(p.x + arrowSize1.height / 2, p.y))
                                        .Points.Add(New System.Windows.Point(p.x + arrowSize1.height / 2, p.y + arrowSize1.width / 2))
                                        .Points.Add(New System.Windows.Point(p.x, p.y + arrowSize1.width / 2))
                                        .Points.Add(New System.Windows.Point(p.x - arrowSize1.height / 2, p.y))
                                    Case arrowTipTypes.directionIndicatorRight
                                        .Points.Add(New System.Windows.Point(p.x - arrowSize1.height / 2, p.y))
                                        .Points.Add(New System.Windows.Point(p.x, p.y - arrowSize1.width / 2))
                                        .Points.Add(New System.Windows.Point(p.x + arrowSize1.height / 2, p.y - arrowSize1.width / 2))
                                        .Points.Add(New System.Windows.Point(p.x + arrowSize1.height / 2, p.y))
                                End Select

                                Dim myRotateTransform As New Windows.Media.RotateTransform
                                myRotateTransform.CenterX = p.x
                                myRotateTransform.CenterY = p.y
                                myRotateTransform.Angle = -angl

                                .RenderTransform = myRotateTransform

                                '.Stroke = System.Windows.Media.Brushes.Red
                                .Stroke = Me.pen.stroke
                                .StrokeDashArray = Me.pen.dashArray
                                .StrokeThickness = t.width

                            End With

                            myCanvas.Children.Add(myPolyline)
                        Next
                    End If
                    '
                    ' Text
                    '
                    If p_caption <> "" Then
                        Dim tb As New Windows.Controls.TextBlock
                        Dim myCaption As String
                        myCaption = caption

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
                        myCanvas.SetLeft(tb, (p1.x + p2.x) / 2)
                        myCanvas.SetTop(tb, (p1.y + p2.y) / 2)
                    End If

                End With
            Else
                Throw New NotImplementedException()
            End If

        End Sub

    End Class

End Namespace