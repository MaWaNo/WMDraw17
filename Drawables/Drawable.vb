Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.Windows.Media
Imports wmg = WallnerMild.Geom

#Const dp = False

Namespace WMDraw

#Region "Drawable Interface"

    Public Interface Drawable
        Inherits IComparable(Of Drawable)

        ' Markus:
        '
        ' delegate methods to calculate the context sizes at time of drawing to device
        ' delegates are pointers to a method of the calling object and solve the problem
        ' of the child object knowing nothing about the parent. The child asks the parent to
        ' perform the delegate method.

        ''' <summary>
        ''' Delegate Function to calculate coordinates at time of drawing to device
        ''' </summary>
        ''' <param name="p"></param>
        ''' <returns></returns>
        ''' <see cref="https://www.codeproject.com/Articles/30458/Delegates-in-VB-NET"/>
        '''
        Delegate Function contextCoordinates(ByVal p As Point) As Point

        ''' <summary>
        ''' Delegate Function to calculate size at time of drawing to device
        ''' </summary>
        ''' <param name="inputSize"></param>
        ''' <returns></returns>
        Delegate Function contextSize(ByVal inputSize As size) As size

        ''' <summary>
        ''' Delegate Function receive Size in World coordinates
        ''' </summary>
        Delegate Function estimateWorldSize(ByVal s As size) As size

        ''' <summary>
        ''' Pen
        ''' </summary>
        ''' <returns></returns>
        Property pen As pen

        ''' <summary>
        ''' Drawing Order
        ''' </summary>
        ''' <returns></returns>
        Property zIndex As Long

        ' Property fill As fill

        ''' <summary>
        ''' Bounding Rectangle to return world-coordinate sizw for proper scaling
        ''' </summary>
        '''
        ''' <returns></returns>
        Function boundingRectangle(Optional getWordSize As estimateWorldSize = Nothing) As Double()

        ''' <summary>
        ''' Actually draw to device
        ''' here the delegate functions of the parent object are consumed
        ''' </summary>
        ''' <param name="contextobject"></param>
        ''' <param name="contextCoordinatesDelegate"></param>
        ''' <param name="contextSizeDelegate"></param>
        Sub draw(contextobject As ContextObject, contextCoordinatesDelegate As contextCoordinates, Optional contextSizeDelegate As contextSize = Nothing)

    End Interface

#End Region

    Public Enum horizontalAlignment
        left
        center
        right
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

    Public Enum verticalAlignment
        top
        center
        bottom
    End Enum

    '
    ' Zusammenfassung:
    '     Implementiert einen Satz vordefinierter Farben.
    Public NotInheritable Class WMColors

        Public Shared ReadOnly Property Black As Color
            Get
                Return Color.FromRgb(0, 0, 0)
            End Get
        End Property

        Public Shared ReadOnly Property DarkGreen As Color
            Get
                Return Color.FromRgb(0, 87, 24)
            End Get
        End Property

        Public Shared ReadOnly Property DarkGrey As Color
            Get
                Return Color.FromRgb(64, 64, 64)
            End Get
        End Property

        Public Shared ReadOnly Property DarkRed As Color
            Get
                Return Color.FromRgb(164, 18, 21)
            End Get
        End Property

        Public Shared ReadOnly Property DarkWallnerMildBlue As Color
            Get
                Return Color.FromRgb(29, 130, 190)
            End Get
        End Property

        Public Shared ReadOnly Property LightGreen As Color
            Get
                Return Color.FromRgb(198, 228, 195)
            End Get
        End Property

        Public Shared ReadOnly Property LightGrey As Color
            Get
                Return Color.FromRgb(223, 223, 223)
            End Get
        End Property

        Public Shared ReadOnly Property LightWallnerMildBlue As Color
            Get
                Return Color.FromRgb(203, 230, 247)
            End Get
        End Property

        Public Shared ReadOnly Property MediumWallnerMildBlue As Color
            Get
                Return Color.FromRgb(203, 230, 247)
            End Get
        End Property

        Public Shared ReadOnly Property MediumWallnerMildBlueTransparent As Color
            Get
                Return Color.FromArgb(&H80, &H9B, &HCE, &HF3)
            End Get
        End Property

        Public Shared ReadOnly Property Olive As Color
            Get
                Return Color.FromRgb(&H73, &H71, &HB)
            End Get
        End Property

        Public Shared ReadOnly Property Red As Color
            Get
                Return Color.FromRgb(&HFF, 0, 0)
            End Get
        End Property

        Public Shared ReadOnly Property Transparent As Color
            Get
                Return Color.FromArgb(0, 255, 255, 255)
            End Get
        End Property

        Public Shared ReadOnly Property White As Color
            Get
                Return Color.FromRgb(255, 255, 255)
            End Get
        End Property
        Public Shared ReadOnly Property WoodDarkBrown As Color
            Get
                Return Color.FromRgb(&HA1, &H88, &H67)
            End Get
        End Property

        Public Shared ReadOnly Property WoodDarkYellow As Color
            Get
                Return Color.FromRgb(&HE6, &HDC, &H87)
            End Get
        End Property

        Public Shared ReadOnly Property WoodLightBrown As Color
            Get
                Return Color.FromRgb(&HE2, &HCF, &HB6)
            End Get
        End Property

        Public Shared ReadOnly Property WoodLightYellow As Color
            Get
                Return Color.FromRgb(&HFF, &HFC, &HE3)
            End Get
        End Property

        Public Shared ReadOnly Property WoodMediumBrown As Color
            Get
                Return Color.FromRgb(&HCE, &HAE, &H84)
            End Get
        End Property

        Public Shared ReadOnly Property WoodMediumYellow As Color
            Get
                Return Color.FromRgb(&HFF, &HF5, &H96)
            End Get
        End Property


        ''' <summary>
        ''' Calculate color values for a color with given opacity
        ''' to look identical to inputColor if placed above the given background color.
        ''' </summary>
        ''' <param name="inputColor"></param>
        ''' <param name="opacity"></param>
        ''' <param name="backgroundColor"></param>
        ''' <returns></returns>
        Public Shared Function CalculateTransparentColor(inputColor As Color, opacity As Double, Optional backgroundColor As Color? = Nothing) As Color
            Dim bgColor As Color
            ' test for nothing and set white as default
            If backgroundColor Is Nothing Then backgroundColor = Colors.White
            bgColor = backgroundColor

            Dim outputColor As Color

            Dim r As Single
            Dim g As Single
            Dim b As Single

            r = -(((CLng(inputColor.R) - CLng(bgColor.R)) * 1.0 - CLng(inputColor.R)) * opacity + (CLng(bgColor.R) - CLng(inputColor.R)) * 1.0) / opacity
            g = -(((CLng(inputColor.G) - CLng(bgColor.G)) * 1.0 - CLng(inputColor.G)) * opacity + (CLng(bgColor.G) - CLng(inputColor.G)) * 1.0) / opacity
            b = -(((CLng(inputColor.B) - CLng(bgColor.B)) * 1.0 - CLng(inputColor.B)) * opacity + (CLng(bgColor.B) - CLng(inputColor.B)) * 1.0) / opacity

            If r < 0 Then r = 0
            If g < 0 Then g = 0
            If b < 0 Then b = 0

            If r > 255 Then r = 255
            If g > 255 Then g = 255
            If b > 255 Then b = 255

            outputColor.R = CByte(r)
            outputColor.G = CByte(g)
            outputColor.B = CByte(b)
            outputColor.A = CByte(opacity * 255)

            Return outputColor

        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="wmd"></param>
        ''' <param name="minValue"></param>
        ''' <param name="maxValue"></param>
        ''' <param name="alpha"></param>
        ''' <returns></returns>
        Public Shared Function DrawRainbowScale(wmd As Drawing, minValue As Double, maxValue As Double, steps As Integer, Optional alpha As Integer = 255,
            Optional fWidth As Single = 0.2, Optional fHeight As Single = 0.9, Optional fMargin As Single = 0.05,
            Optional horizontalAlignment As horizontalAlignment = horizontalAlignment.right)

            Dim p As New pen
            Dim curValue As Double
            Dim lowerValue As Double
            Dim upperValue As Double


            p.color = Colors.White
            p.thickness = 5


            Dim fBorderY As Single = 0.05
            Dim fBorderX As Single = 0.065

            Dim fLeft As Single
            Dim fBottom As Single
            Dim fRight As Single
            Dim fTop As Single

            Select Case horizontalAlignment
                Case horizontalAlignment.left
                    fLeft = fMargin
                    fRight = fMargin + fWidth
                Case horizontalAlignment.center
                    fLeft = 0.5 - fWidth / 2
                    fRight = 0.5 + fWidth / 2
                Case horizontalAlignment.right
                    fLeft = 1 - fMargin - fWidth
                    fRight = 1 - fMargin
            End Select
            fBottom = fMargin
            fTop = fMargin + fHeight

            If fLeft < 0 Then fLeft = 0
            If fRight > 1 Then fRight = 1
            If fBottom < 0 Then fBottom = 0
            If fTop > 1 Then fTop = 1

            Dim fTextSize As Single = 0.02

            Dim outerRect As New Rectangle(fLeft, fBottom, fRight, fTop, Reference.contextFraction)

            Dim oFill As New fill
            oFill.color = Colors.White
            outerRect.fill = oFill
            wmd.add(outerRect)


            For i As Integer = 0 To steps - 1
                lowerValue = minValue + (maxValue - minValue) / steps * i
                upperValue = minValue + (maxValue - minValue) / steps * (i + 1)
                curValue = (lowerValue + upperValue) / 2

                Dim f As New fill
                f.color = GetRainbowColor(curValue, minValue, maxValue, alpha)

                Dim r3 As New Rectangle(fLeft + fBorderX, fBottom + fBorderY + (fTop - fBottom - 2 * fBorderY) / steps * i,
                                        fRight - fBorderX, fBottom + fBorderY + (fTop - fBottom - 2 * fBorderY) / steps * (i + 1), Reference.contextFraction)

                r3.pen = p
                r3.fill = f
                wmd.add(r3)

                Dim t3 As New Text(fLeft + fBorderX, fBottom + fBorderY + (fTop - fBottom - 2 * fBorderY) / steps * i, Reference.contextFraction,
                                   SetSigFigs(lowerValue, 3), fTextSize, Reference.contextFraction)
                t3.horizontalAlignment = horizontalAlignment.left
                t3.verticalAlignment = verticalAlignment.center
                wmd.add(t3)

                Dim t4 As New Text(fRight - fBorderX, fBottom + fBorderY + (fTop - fBottom - 2 * fBorderY) / steps * (i + 1 / 2), Reference.contextFraction,
                                   SetSigFigs(curValue, 3), fTextSize, Reference.contextFraction)
                t4.horizontalAlignment = horizontalAlignment.right
                t4.verticalAlignment = verticalAlignment.center
                wmd.add(t4)

                If i = steps - 1 Then
                    Dim t5 As New Text(fLeft + fBorderX, fBottom + fBorderY + (fTop - fBottom - 2 * fBorderY) / steps * (i + 1), Reference.contextFraction,
                                  SetSigFigs(upperValue, 3), fTextSize, Reference.contextFraction)
                    t5.horizontalAlignment = horizontalAlignment.left
                    t5.verticalAlignment = verticalAlignment.center
                    wmd.add(t5)
                End If

            Next


        End Function
        Public Shared Function GetRainbowColor(value As Single,
                                               Optional minValue As Double = 0, Optional maxValue As Double = 1, Optional alpha As Integer = 255) As Color

            Dim interpolationFactor As Single
            Dim normalizedValue As Single


            ' Define the rainbow colors
            Dim colors As Color() = {
                System.Windows.Media.Colors.DarkBlue,
                System.Windows.Media.Colors.Blue,
                System.Windows.Media.Colors.Cyan,
                System.Windows.Media.Colors.Green,
                System.Windows.Media.Colors.Yellow,
                System.Windows.Media.Colors.Orange,
                System.Windows.Media.Colors.OrangeRed,
                System.Windows.Media.Colors.DarkRed}

            If minValue <> maxValue Then
                normalizedValue = (value - minValue) / (maxValue - minValue)
            End If


            ' Normalize the value to ensure it falls within the expected range
            normalizedValue = Math.Max(0, Math.Min(1, normalizedValue))

            Debug.Print("value is {0} [{1},{2}]. normalized Value is {3}", value, minValue, maxValue, normalizedValue)

            interpolationFactor = (normalizedValue * (colors.Length - 1)) Mod 1
            Debug.Print("interpolationFactor: {0}", interpolationFactor)


            ' Calculate which two colors to interpolate between
            Dim colorIndex As Integer = CInt(Math.Floor(normalizedValue * (colors.Length - 1)))
            Dim colorStart As Color = colors(colorIndex)
            Dim colorEnd As Color = colors(Math.Min(colorIndex + 1, colors.Length - 1))

            ' Calculate the interpolation factor

            ' Interpolate between the two colors
            Dim r As Integer
            Dim g As Integer
            Dim b As Integer

            r = Math.Min((CInt(colorStart.R) + (CInt(colorEnd.R) - CInt(colorStart.R)) * interpolationFactor), 255)
            g = Math.Min((CInt(colorStart.G) + (CInt(colorEnd.G) - CInt(colorStart.G)) * interpolationFactor), 255)
            b = Math.Min((CInt(colorStart.B) + (CInt(colorEnd.B) - CInt(colorStart.B)) * interpolationFactor), 255)

            Dim resultColor As Color = Color.FromArgb(alpha, r, g, b)
            Return resultColor

        End Function

        Public Function ConvertDrawingColorToMediaColor(drawingColor As System.Drawing.Color) As Windows.Media.Color
            Return Windows.Media.Color.FromArgb(drawingColor.A, drawingColor.R, drawingColor.G, drawingColor.B)
        End Function
    End Class

#Region "myMath Class"

    ''' <summary>
    ''' Helper Class to perform Math-Operations for Drawings
    ''' </summary>
    Public Class myMath

        ''' <summary>
        ''' convert radians to degrees
        ''' </summary>
        ''' <param name="angle_radians">Angle in radians</param>
        ''' <returns>angle in degrees</returns>
        Public Shared Function degrees(angle_radians As Double) As Double
            degrees = angle_radians * 180 / Math.PI
        End Function

        ''' <summary>
        ''' include a point in the bounding rectangle
        ''' </summary>
        ''' <param name="r">rectangle by reference (0..xmin, 1...ymin, 2...xmax, 3...ymax</param>
        ''' <param name="p">point to include</param>
        Public Shared Sub includeInBoundingRectangle(ByRef r() As Double, p As Point)
            If p.x < r(0) Then r(0) = p.x
            If p.x > r(2) Then r(2) = p.x
            If p.y < r(1) Then r(1) = p.y
            If p.y > r(3) Then r(3) = p.y
        End Sub

        ''' <summary>
        ''' include a point in the bounding rectangle
        ''' </summary>
        ''' <param name="r">rectangle by reference (0..xmin, 1...ymin, 2...xmax, 3...ymax</param>
        ''' <param name="x">x coordinate</param>
        ''' <param name="y">y coordinate</param>
        Public Shared Sub includeInBoundingRectangle(ByRef r() As Double, x As Double, y As Double)
            If x < r(0) Then r(0) = x
            If x > r(2) Then r(2) = x
            If y < r(1) Then r(1) = y
            If y > r(3) Then r(3) = y
        End Sub

        ''' <summary>
        ''' convert degrees to radians
        ''' </summary>
        ''' <param name="angle_degrees">angle in degrees</param>
        ''' <returns>
        ''' angle in radians
        ''' </returns>
        Public Shared Function radians(angle_degrees As Double) As Double
            radians = angle_degrees * Math.PI / 180
        End Function

    End Class

#End Region

#Region "Size Class"

    ''' <summary>
    ''' Size of an object in defined reference-system
    ''' </summary>
    <ComVisible(False)>
    Public Class size
        Private p_height As Double
        Private p_Reference As New Reference
        Private p_width As Double

        Public Sub New()
            p_width = 0
            p_height = 0
            p_Reference = Reference.contextUnits
        End Sub

        Public Sub New(value As Double)
            p_width = value
            p_height = value
            p_Reference = Reference.world
        End Sub

        Public Sub New(width As Double, height As Double)
            p_width = width
            p_height = height
            p_Reference = Reference.world
        End Sub

        Public Sub New(value As Double, ReferenceAsLong As Long)
            Dim r As Reference
            Try
                r = CType(ReferenceAsLong, Reference)
            Catch ex As Exception
                r = Reference.world
            End Try

            p_width = value
            p_height = value
            p_Reference = r
        End Sub

        Public Sub New(value As Double, Reference As Reference)
            p_width = value
            p_height = value
            p_Reference = Reference
        End Sub

        Public Sub New(width As Double, height As Double, Reference As Reference)
            p_width = width
            p_height = height
            p_Reference = Reference
        End Sub

        ''' <summary>
        ''' average size
        ''' </summary>
        ''' <returns></returns>
        Public Property average As Double
            Get
                If IsNothing(width) Then
                    If IsNothing(height) Then
                        Return 0
                    Else
                        Return height
                    End If
                Else
                    If IsNothing(height) Then
                        Return width
                    End If
                End If
                average = (Math.Abs(width) + Math.Abs(height)) / 2
            End Get
            Set(value As Double)
                p_height = value
                p_width = value
            End Set
        End Property

        ''' <summary>
        ''' height in reference-system
        ''' </summary>
        ''' <returns></returns>
        Public Property height As Double
            Get
                Return p_height
            End Get
            Set(value As Double)
                p_height = value
            End Set
        End Property

        ''' <summary>
        ''' max size
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property max As Double
            Get
                max = Math.Max(Math.Abs(width), Math.Abs(height))
            End Get
        End Property

        ''' <summary>
        ''' reference system
        ''' </summary>
        ''' <returns></returns>
        Public Property Reference As Reference
            Get
                Return p_Reference
            End Get
            Set(value As Reference)
                p_Reference = value
            End Set
        End Property

        ''' <summary>
        ''' Thickness in reference system (average of width and height)
        ''' equivalent to average
        ''' </summary>
        ''' <returns></returns>
        Public Property thickness As Double
            Get
                Try
                    Return average
                Catch ex As Exception
                    Return 0
                End Try
            End Get
            Set(value As Double)
                average = value
            End Set
        End Property

        ''' <summary>
        ''' Width in reference-system
        ''' </summary>
        ''' <returns></returns>
        Public Property width As Double
            Get
                Return p_width
            End Get
            Set(value As Double)
                p_width = value
            End Set
        End Property

        Public Shared Operator /(ByVal a As size, b As Double) As size
            If b <> 0 Then
                Return New size(a.width / b, a.height / b, a.Reference)
            Else
                Throw New DivideByZeroException("size divided by 0")
                Return a
            End If
        End Operator

    End Class

#End Region

#Region "Point Class"

    ''' <summary>
    ''' Point as a drawable item
    ''' </summary>
    <ComVisible(False)>
    Public Class Point
        Implements Drawable

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

        Private p_pen As New pen

        ''' <summary>
        ''' x-Coordinate in world coordinates
        ''' </summary>
        Private p_x As Double

        ''' <summary>
        ''' y-Coordinate in world coordinates
        ''' </summary>
        Private p_y As Double

        Private p_zIndex As Long

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
            p_displaySize.Reference = displayReference
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

        Public Property zIndex As Long Implements Drawable.zIndex
            Get
                Return p_zIndex
            End Get
            Set(value As Long)
                p_zIndex = value
            End Set
        End Property

        Public Shared Operator <>(pointA As Point, pointB As Point) As Boolean
            Return Not (pointA.x = pointB.x And pointA.y = pointB.y)
        End Operator

        Public Shared Operator =(pointA As Point, pointB As Point) As Boolean
            Return (pointA.x = pointB.x And pointA.y = pointB.y)
        End Operator
        ''' <summary>
        ''' bounding rectangle
        ''' </summary>
        '''
        ''' <returns></returns>
        Public Function boundingRectangle(Optional getWordSize As Drawable.estimateWorldSize = Nothing) As Double() Implements Drawable.boundingRectangle

            Dim r(3) As Double

            r(0) = Me.x
            r(1) = Me.y
            r(2) = Me.x
            r(3) = Me.y

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

                    Dim t As size
                    t = contextSizeDelegate(Me.pen.size)
                    If t.width > 0 And t.width < 0.1 Then t.width = 0.1

                    Select Case display
                        Case PointDisplay.x
                            Dim myLine As New System.Windows.Shapes.Line()

                            With myLine
                                .X1 = p1.x - s1.width / 2
                                .Y1 = p1.y + s1.height / 2

                                .X2 = p1.x + s1.width / 2
                                .Y2 = p1.y - s1.width / 2

                                .Stroke = pen.stroke
                                .StrokeThickness = t.width
                            End With

                            myCanvas.Children.Add(myLine)
                            'System.Windows.Controls.Canvas.SetLeft(myLine, p1.x)
                            'System.Windows.Controls.Canvas.SetTop(myLine, p1.y)

                            Dim myLine2 As New System.Windows.Shapes.Line()
                            With myLine2
                                .X1 = p1.x + s1.width / 2
                                .Y1 = p1.y + s1.height / 2

                                .X2 = p1.x - s1.width / 2
                                .Y2 = p1.y - s1.width / 2

                                .Stroke = pen.stroke
                                .StrokeThickness = t.width
                            End With

                            myCanvas.Children.Add(myLine2)
                            'System.Windows.Controls.Canvas.SetLeft(myLine2, p1.x)
                            'System.Windows.Controls.Canvas.SetTop(myLine2, p1.y)

                        Case PointDisplay.circle, PointDisplay.dot
                            Dim myPoint As New System.Windows.Shapes.Ellipse

                            With myPoint

                                .Width = s1.width
                                .Height = s1.height

                                .Stroke = Me.pen.stroke
                                .StrokeThickness = t.width

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

        Public Function rotate(center As Point, angle As Double) As Point
            Dim p As New wmg.Vector
            Dim c As New wmg.Vector
            c.x = center.x
            c.y = center.y
            p.x = Me.x
            p.y = Me.y
            p.subtract(c)
            p.rotate(angle * Math.PI / 180)
            p.add(c)

            rotate = Me
            rotate.x = p.x
            rotate.y = p.y
        End Function

    End Class

#End Region

#Region "Beam Class"

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
                    If t.width > 0 And t.width < 0.1 Then t.width = 0.1

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

    <ComVisible(False)>
    Public Class BeamUniformLoad
        Inherits Line
        Implements Drawable

        Private Const DEFAULT_ARROW_DISTANCE_MM = 3
        Private Const DEFAULT_ARROWTIP_HEIGHT_MM = 1.5
        Private Const DEFAULT_ARROWTIP_WIDTH_MM = 1
        Private Const DEFAULT_LOADBACKGROUND_HEX = "#37c6c6c6"
        Private Const DEFAULT_LOADOFFSET_MM = 10
        Private p_captionEnd As String
        Private p_captionMidpoint As String
        Private p_captionStart As String

        'Private p_offsetIndex As Integer
        Private p_direction As loadDirectionType

        Private p_drawArrows As Boolean
        Private p_fill As fill

        Private p_lengthReference As loadReferenceLength

        Private p_loadIntensityEnd As New size

        Private p_loadIntensityStart As New size

        Private p_loadOffset As New size

        'Private p_font As font
        Sub New()
            p_loadIntensityStart = New size
            p_loadIntensityEnd = New size
            p_loadOffset = New size
            Me.pen = New pen
            p_fill = New fill
            p_fill.color = ColorConverter.ConvertFromString(DEFAULT_LOADBACKGROUND_HEX)
            p_drawArrows = True

            p_loadOffset.width = DEFAULT_LOADOFFSET_MM
            p_loadOffset.height = DEFAULT_LOADOFFSET_MM
            p_loadOffset.Reference = Reference.contextMillimeters

        End Sub

        Sub New(startX As Double, startY As Double, endX As Double, endY As Double, intensityStart As size, intensityEnd As size, lengthReference As loadReferenceLength, direction As loadDirectionType)
            Me.New
            Me.startPoint = New Point(startX, startY)
            Me.endPoint = New Point(endX, endY)
            Me.loadIntensityStart = intensityStart
            Me.loadIntensityEnd = intensityEnd
            Me.p_lengthReference = lengthReference
            Me.p_direction = direction
        End Sub

        Sub New(refBeam As Beam, intensity As size, lengthReference As loadReferenceLength, direction As loadDirectionType)
            Me.New
            Me.startPoint = refBeam.startPoint
            Me.endPoint = refBeam.endPoint
            Me.loadIntensityStart = intensity
            Me.loadIntensityEnd = intensity
            Me.lengthReference = lengthReference
            Me.p_direction = direction
        End Sub

        Public Enum loadDirectionType
            localKS
            globalKS
        End Enum

        Public Enum loadReferenceLength
            horizontalProjection
            realLength
        End Enum

        ''' <summary>
        ''' Caption at end of udl
        ''' </summary>
        ''' <returns></returns>
        Public Property captionEnd As String
            Get
                Return p_captionEnd
            End Get
            Set(value As String)
                p_captionEnd = value
            End Set
        End Property

        ''' <summary>
        ''' Caption at midpoint of udl
        ''' </summary>
        ''' <returns></returns>
        Public Property captionMidpoint As String
            Get
                Return p_captionMidpoint
            End Get
            Set(value As String)
                p_captionMidpoint = value
            End Set
        End Property

        ''' <summary>
        ''' Caption at start of udl
        ''' </summary>
        ''' <returns></returns>
        Public Property captionStart As String
            Get
                Return p_captionStart
            End Get
            Set(value As String)
                p_captionStart = value
            End Set
        End Property

        Public Property direction As loadDirectionType
            Get
                Return p_direction
            End Get
            Set(value As loadDirectionType)
                p_direction = value
            End Set
        End Property

        Public Property drawArrows() As Boolean
            Get
                Return p_drawArrows
            End Get
            Set(value As Boolean)
                p_drawArrows = value
            End Set
        End Property

        Public Property fill() As fill
            Get
                Return p_fill
            End Get
            Set(value As fill)
                p_fill = value
            End Set
        End Property

        Public Property lengthReference As loadReferenceLength
            Get
                Return p_lengthReference
            End Get
            Set(value As loadReferenceLength)
                p_lengthReference = value
            End Set
        End Property

        ''' <summary>
        ''' Load Intensity at start point
        ''' </summary>
        ''' <returns></returns>
        Public Property loadIntensityEnd() As size
            Get
                Return p_loadIntensityEnd
            End Get
            Set(value As size)
                p_loadIntensityEnd = value
            End Set
        End Property

        ''' <summary>
        ''' Load Intensity at start point
        ''' </summary>
        ''' <returns></returns>
        Public Property loadIntensityStart() As size
            Get
                Return p_loadIntensityStart
            End Get
            Set(value As size)
                p_loadIntensityStart = value
            End Set
        End Property

        ''' <summary>
        ''' Distance from reference points to actual load display
        ''' </summary>
        ''' <returns></returns>
        Public Property loadOffset As size
            Get
                Return p_loadOffset
            End Get
            Set(value As size)
                p_loadOffset = value
            End Set
        End Property

        Public Overloads Function boundingRectangle(Optional getWordSize As Drawable.estimateWorldSize = Nothing) As Double() Implements Drawable.boundingRectangle
            Dim r(3) As Double

            Dim p1 As New Point
            Dim p2 As New Point

            ' sort points
            For i As Integer = 0 To 1
                If r(0 + i) > r(2 + i) Then
                    Dim t As Double = r(0 + i)
                    r(0 + i) = r(2 + i)
                    r(2 + i) = t
                End If
            Next

            Dim udlStart As New size
            Dim udlEnd As New size
            Dim loadOffsetCU As New size

            udlStart = getWordSize(Me.loadIntensityStart)
            udlEnd = getWordSize(Me.loadIntensityEnd)
            loadOffsetCU = getWordSize(loadOffset)

            Dim gl1 As New Geom.line
            Dim gl2 As New Geom.line

            gl1 = New Geom.line(startPoint.x, startPoint.y, endPoint.x - startPoint.x, endPoint.y - startPoint.y)
            gl2 = gl1

            Select Case Me.direction
                Case loadDirectionType.localKS
                    gl2 = gl1.offsetLine(loadOffsetCU.height)
                Case loadDirectionType.globalKS
                    gl2.StartPoint.x = gl1.StartPoint.x
                    gl2.StartPoint.y = gl1.StartPoint.y - loadOffsetCU.height
            End Select

            ' unit and normal vectors
            Dim uVect As New Geom.Vector
            Dim nVect As New Geom.Vector
            Dim v1 As New Geom.Vector   ' length direction
            Dim v2 As New Geom.Vector   ' load direction

            uVect = gl1.Direction
            uVect = uVect.unitVector

            nVect = uVect.normal

            Select Case Me.direction
                Case loadDirectionType.localKS
                    v1 = nVect
                Case loadDirectionType.globalKS
                    v1 = New Geom.Vector(0, 1)
            End Select

            Select Case Me.lengthReference
                Case loadReferenceLength.realLength
                    ' vabene
                    v2 = uVect
                Case loadReferenceLength.horizontalProjection
                    Dim minY As Double
                    minY = gl2.StartPoint.y
                    If minY > gl2.EndPoint.y Then minY = gl2.EndPoint.y
                    gl2.StartPoint.y = minY
                    gl2.EndPoint.y = minY
                    gl2.Direction.y = 0
                    v2 = New Geom.Vector(1, 0)
            End Select

            '
            ' boundary polygon
            '
            If startPoint.coordinateReference = Reference.world Then
                r(0) = startPoint.x
                r(1) = startPoint.y
            End If

            If startPoint.coordinateReference = Reference.world Then
                r(2) = endPoint.x
                r(3) = endPoint.y
            Else
                r(2) = startPoint.x
                r(3) = startPoint.y
            End If

            If r(0) > gl2.StartPoint.x Then r(0) = gl2.StartPoint.x
            If r(1) > gl2.StartPoint.y Then r(1) = gl2.StartPoint.y
            If r(2) < gl2.StartPoint.x Then r(2) = gl2.StartPoint.x
            If r(3) < gl2.StartPoint.y Then r(3) = gl2.StartPoint.y

            If r(0) > gl2.StartPoint.x Then r(0) = gl2.StartPoint.x
            If r(1) > gl2.StartPoint.y Then r(1) = gl2.StartPoint.y
            If r(2) < gl2.StartPoint.x Then r(2) = gl2.StartPoint.x
            If r(3) < gl2.StartPoint.y Then r(3) = gl2.StartPoint.y

            If r(0) > gl2.EndPoint.x Then r(0) = gl2.EndPoint.x
            If r(1) > gl2.EndPoint.y Then r(1) = gl2.EndPoint.y
            If r(2) < gl2.EndPoint.x Then r(2) = gl2.EndPoint.x
            If r(3) < gl2.EndPoint.y Then r(3) = gl2.EndPoint.y

            If r(0) > gl2.EndPoint.x + v1.x * udlEnd.max * -1 Then r(0) = gl2.EndPoint.x + v1.x * udlEnd.max * -1
            If r(1) > gl2.EndPoint.y + v1.y * udlEnd.max * +1 Then r(1) = gl2.EndPoint.y + v1.y * udlEnd.max * +1
            If r(2) < gl2.EndPoint.x + v1.x * udlEnd.max * -1 Then r(2) = gl2.EndPoint.x + v1.x * udlEnd.max * -1
            If r(3) < gl2.EndPoint.y + v1.y * udlEnd.max * +1 Then r(3) = gl2.EndPoint.y + v1.y * udlEnd.max * +1

            If r(0) > gl2.StartPoint.x + v1.x * udlStart.max * -1 Then r(0) = gl2.StartPoint.x + v1.x * udlStart.max * -1
            If r(1) > gl2.StartPoint.y + v1.y * udlStart.max * +1 Then r(1) = gl2.StartPoint.y + v1.y * udlStart.max * +1
            If r(2) < gl2.StartPoint.x + v1.x * udlStart.max * -1 Then r(2) = gl2.StartPoint.x + v1.x * udlStart.max * -1
            If r(3) < gl2.StartPoint.y + v1.y * udlStart.max * +1 Then r(3) = gl2.StartPoint.y + v1.y * udlStart.max * +1

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

        Public Overrides Sub draw(contextobject As ContextObject, contextCoordinatesDelegate As Drawable.contextCoordinates, Optional contextSizeDelegate As Drawable.contextSize = Nothing) Implements Drawable.draw
            If TypeOf contextobject.Item Is System.Windows.Controls.Canvas Then
                Dim myCanvas As New System.Windows.Controls.Canvas
                myCanvas = TryCast(contextobject.Item, System.Windows.Controls.Canvas)
                With myCanvas
                    Dim line1 As New System.Windows.Shapes.Line
                    Dim line2 As New System.Windows.Shapes.Line
                    Dim line3 As New System.Windows.Shapes.Line
                    Dim line4 As New System.Windows.Shapes.Line

                    '
                    ' calculate coordinates in localKS system
                    ' by a call-back to the calling drawing object
                    '
                    Dim p1 As New Point
                    Dim p2 As New Point
                    Dim udlStart As New size
                    Dim udlEnd As New size
                    Dim loadOffsetCU As New size

                    Dim t As size

                    Dim tmp As Double

                    p1 = contextCoordinatesDelegate(startPoint)
                    p2 = contextCoordinatesDelegate(endPoint)

                    udlStart = contextSizeDelegate(Me.loadIntensityStart)
                    udlEnd = contextSizeDelegate(Me.loadIntensityEnd)
                    loadOffsetCU = contextSizeDelegate(loadOffset)

                    t = contextSizeDelegate(Me.pen.size)
                    If t.width > 0 And t.width < 0.1 Then t.width = 0.1

                    Dim gl1 As New Geom.line
                    Dim gl2 As New Geom.line

                    gl1 = New Geom.line(p1.x, p1.y, p2.x - p1.x, p2.y - p1.y)
                    gl2 = gl1

                    Select Case Me.direction
                        Case loadDirectionType.localKS
                            gl2 = gl1.offsetLine(loadOffsetCU.height)
                        Case loadDirectionType.globalKS
                            gl2.StartPoint.x = gl1.StartPoint.x
                            gl2.StartPoint.y = gl1.StartPoint.y - loadOffsetCU.height
                    End Select

                    ' unit and normal vectors
                    Dim uVect As New Geom.Vector
                    Dim nVect As New Geom.Vector
                    Dim v1 As New Geom.Vector   ' length direction
                    Dim v2 As New Geom.Vector   ' load direction

                    uVect = gl1.Direction
                    uVect = uVect.unitVector

                    nVect = uVect.normal

                    Dim myPolygon = New System.Windows.Shapes.Polygon

                    Select Case Me.direction
                        Case loadDirectionType.localKS
                            v1 = nVect
                        Case loadDirectionType.globalKS
                            v1 = New Geom.Vector(0, 1)
                    End Select

                    Select Case Me.lengthReference
                        Case loadReferenceLength.realLength
                            ' vabene
                            v2 = uVect
                        Case loadReferenceLength.horizontalProjection
                            Dim minY As Double
                            minY = gl2.StartPoint.y
                            If minY > gl2.EndPoint.y Then minY = gl2.EndPoint.y
                            gl2.StartPoint.y = minY
                            gl2.EndPoint.y = minY
                            gl2.Direction.y = 0
                            v2 = New Geom.Vector(1, 0)
                    End Select

                    '
                    ' boundary polygon
                    '
                    With myPolygon
                        .Points.Add(New System.Windows.Point(gl2.StartPoint.x, gl2.StartPoint.y))
                        .Points.Add(New System.Windows.Point(gl2.EndPoint.x, gl2.EndPoint.y))

                        .Points.Add(New System.Windows.Point(gl2.EndPoint.x + v1.x * udlEnd.max * -1, gl2.EndPoint.y + v1.y * udlEnd.max * -1))
                        .Points.Add(New System.Windows.Point(gl2.StartPoint.x + v1.x * udlStart.max * -1, gl2.StartPoint.y + v1.y * udlStart.max * -1))

                        .Stroke = Me.pen.stroke
                        .StrokeThickness = 1

                        .Fill = p_fill.Brush
                    End With
                    myCanvas.Children.Add(myPolygon)

                    Dim ad As New size
                    ad = contextSizeDelegate(New size(DEFAULT_ARROW_DISTANCE_MM, Reference.contextMillimeters))
                    If ad.average = 0 Then ad.average = 5

                    If drawArrows Then
                        Dim i
                        Dim imax

                        Dim pthFigure As PathFigure = New PathFigure()
                        Dim lineSeg As LineSegment = New LineSegment()
                        Dim myPathSegmentCollection As PathSegmentCollection = New PathSegmentCollection()
                        Dim pthFigureCollection As PathFigureCollection = New PathFigureCollection()

                        Dim arrowHeightCU As size
                        Dim arrowWidthCU As size

                        arrowWidthCU = contextSizeDelegate(New size(DEFAULT_ARROWTIP_WIDTH_MM, Reference.contextMillimeters))
                        arrowHeightCU = contextSizeDelegate(New size(DEFAULT_ARROWTIP_HEIGHT_MM, Reference.contextMillimeters))

                        ' startpoint of pathFigure
                        ' P1 (start  Point)
                        'pthFigure.StartPoint = New System.Windows.Point(p1.x, p1.y)

                        imax = Int(gl2.Direction.length / ad.average)

                        For i = 0 To imax
                            pthFigure = New PathFigure()
                            myPathSegmentCollection = New PathSegmentCollection()
                            lineSeg = New LineSegment()
                            '
                            ' Arrow top   .
                            '             |
                            '
                            pthFigure.StartPoint = New System.Windows.Point(
                                    gl2.StartPoint.x + v2.x * gl2.Direction.length * i / imax + v1.x * udlStart.max * -1 - v1.x * (udlEnd.max - udlStart.max) * i / imax,
                                    gl2.StartPoint.y + v2.y * gl2.Direction.length * i / imax + v1.y * udlStart.max * -1 - v1.y * (udlEnd.max - udlStart.max) * i / imax)
                            '
                            ' Arrow bottom |
                            '              .
                            '
                            lineSeg.Point = New System.Windows.Point(
                                        gl2.StartPoint.x + v2.x * gl2.Direction.length * i / imax,
                                        gl2.StartPoint.y + v2.y * gl2.Direction.length * i / imax)

                            myPathSegmentCollection.Add(lineSeg)
                            pthFigure.Segments = myPathSegmentCollection
                            pthFigureCollection.Add(pthFigure)

                            '
                            ' Arrow hook \
                            '
                            pthFigure = New PathFigure()
                            myPathSegmentCollection = New PathSegmentCollection()
                            lineSeg = New LineSegment()

                            pthFigure.StartPoint = New System.Windows.Point(
                                        gl2.StartPoint.x + v2.x * gl2.Direction.length * i / imax - v2.x * arrowWidthCU.average - v1.x * arrowHeightCU.average,
                                        gl2.StartPoint.y + v2.y * gl2.Direction.length * i / imax - v2.y * arrowWidthCU.average - v1.y * arrowHeightCU.average)

                            lineSeg.Point = New System.Windows.Point(
                                        gl2.StartPoint.x + v2.x * gl2.Direction.length * i / imax,
                                        gl2.StartPoint.y + v2.y * gl2.Direction.length * i / imax)

                            myPathSegmentCollection.Add(lineSeg)

                            lineSeg = New LineSegment()
                            lineSeg.Point = New System.Windows.Point(
                                        gl2.StartPoint.x + v2.x * gl2.Direction.length * i / imax + v2.x * arrowWidthCU.average + v1.x * arrowHeightCU.average,
                                        gl2.StartPoint.y + v2.y * gl2.Direction.length * i / imax - v2.y * arrowWidthCU.average - v1.y * arrowHeightCU.average)

                            ' add line segment to collection
                            myPathSegmentCollection.Add(lineSeg)
                            pthFigure.Segments = myPathSegmentCollection

                            pthFigureCollection.Add(pthFigure)
                        Next

                        Dim pthGeometry As PathGeometry = New PathGeometry()
                        pthGeometry.Figures = pthFigureCollection

                        Dim linePath As Windows.Shapes.Path = New Windows.Shapes.Path
                        linePath.Data = pthGeometry
                        With linePath
                            .Stroke = Me.pen.stroke
                            .StrokeThickness = 0.5
                        End With
                        myCanvas.Children.Add(linePath)
                    End If

                    Dim tb As New Windows.Controls.TextBlock

                    If p_captionStart <> "" Then
                        tb = New Windows.Controls.TextBlock
                        With tb
                            .Text = p_captionStart
                            '.FontSize = textSize.width
                        End With

                        Dim myRotateTransform2 As New Windows.Media.RotateTransform
                        myRotateTransform2.Angle = gl2.Direction.angle
                        tb.RenderTransform = myRotateTransform2
                        myCanvas.Children.Add(tb)

                        Dim myFormattedText As New FormattedText(tb.Text, CultureInfo.CurrentCulture, Windows.FlowDirection.LeftToRight,
                            New Typeface(tb.FontFamily, tb.FontStyle, tb.FontWeight, tb.FontStretch), tb.FontSize, Brushes.Black,
                            New NumberSubstitution(), 1)

                        tb.HorizontalAlignment = Windows.HorizontalAlignment.Center
                        tb.VerticalAlignment = Windows.VerticalAlignment.Center

                        myCanvas.SetLeft(tb, gl2.StartPoint.x + v1.x * udlStart.max * -1)
                        myCanvas.SetTop(tb, gl2.StartPoint.y + v1.y * udlStart.max * -1)
                    End If

                    If p_captionMidpoint <> "" Then
                        tb = New Windows.Controls.TextBlock
                        With tb
                            .Text = p_captionMidpoint
                            '.FontSize = textSize.width
                        End With

                        Dim myRotateTransform2 As New Windows.Media.RotateTransform
                        myRotateTransform2.Angle = gl2.Direction.angle
                        tb.RenderTransform = myRotateTransform2
                        myCanvas.Children.Add(tb)

                        Dim myFormattedText As New FormattedText(tb.Text, CultureInfo.CurrentCulture, Windows.FlowDirection.LeftToRight,
                            New Typeface(tb.FontFamily, tb.FontStyle, tb.FontWeight, tb.FontStretch), tb.FontSize, Brushes.Black,
                            New NumberSubstitution(), 1)

                        tb.HorizontalAlignment = Windows.HorizontalAlignment.Center
                        tb.VerticalAlignment = Windows.VerticalAlignment.Top

                        myCanvas.SetLeft(tb, gl2.StartPointX + gl2.dx / 2 - myFormattedText.Width / 2 * Math.Cos(gl1.Direction.angle * Math.PI / 180) + myFormattedText.Height * Math.Sin(gl1.Direction.angle * Math.PI / 180))
                        myCanvas.SetTop(tb, gl2.StartPointY + gl2.dy / 2 - myFormattedText.Width / 2 * Math.Sin(gl1.Direction.angle * Math.PI / 180) - myFormattedText.Height * Math.Cos(gl1.Direction.angle * Math.PI / 180))

                    End If

                    If p_captionEnd <> "" Then
                        tb = New Windows.Controls.TextBlock
                        With tb
                            .Text = p_captionEnd
                            '.FontSize = textSize.width
                        End With

                        Dim myRotateTransform2 As New Windows.Media.RotateTransform
                        myRotateTransform2.Angle = gl2.Direction.angle
                        tb.RenderTransform = myRotateTransform2
                        myCanvas.Children.Add(tb)

                        Dim myFormattedText As New FormattedText(tb.Text, CultureInfo.CurrentCulture, Windows.FlowDirection.LeftToRight,
                            New Typeface(tb.FontFamily, tb.FontStyle, tb.FontWeight, tb.FontStretch), tb.FontSize, Brushes.Black,
                            New NumberSubstitution(), 1)

                        tb.HorizontalAlignment = Windows.HorizontalAlignment.Right
                        tb.VerticalAlignment = Windows.VerticalAlignment.Center

                        'myCanvas.SetLeft(tb, gl2.EndPoint.x + v1.x * udlEnd.max * -1)
                        'myCanvas.SetTop(tb, gl2.EndPoint.y + v1.y * udlEnd.max * -1)
                        myCanvas.SetLeft(tb, gl2.StartPointX + gl2.dx - myFormattedText.Width * Math.Cos(gl1.Direction.angle * Math.PI / 180) + myFormattedText.Height * Math.Sin(gl1.Direction.angle * Math.PI / 180))
                        myCanvas.SetTop(tb, gl2.StartPointY + gl2.dy - myFormattedText.Width * Math.Sin(gl1.Direction.angle * Math.PI / 180) - myFormattedText.Height * Math.Cos(gl1.Direction.angle * Math.PI / 180))

                    End If
                End With
            Else
                Throw New NotImplementedException()
            End If
        End Sub

    End Class

    <ComVisible(False)>
    Public Class Gridline
        Inherits Line
        Implements Drawable
        Private p_drawSectionWidth As Boolean
        Private p_EndHinge As Boolean
        Private p_hingeSize As size
        Private p_isInOpening As Boolean
        Private p_sectionwidth As New size
        Private p_SectionWidth_Fill As fill
        Private p_SectionWidth_Pen As pen
        Private p_StartHinge As Boolean
        Public Sub New()

            Me.pen = New pen
            Me.pen.thickness = 1
            Me.HingeSize = New size(0.5, 0.5, Reference.contextMillimeters)

            Me.drawSectionWidth = False
            p_SectionWidth_Fill = New fill
            Me.SectionWidth_Fill.color = WMColors.CalculateTransparentColor(WMColors.WoodLightBrown, 0.3)
            Me.SectionWidth_Pen = New pen
            Me.SectionWidth_Pen.color = WMColors.CalculateTransparentColor(WMColors.WoodLightBrown, 0.7)
        End Sub

        Public Sub New(startX As Double, startY As Double, endX As Double, endY As Double)
            Me.New()
            Me.startPoint = New Point(startX, startY)
            Me.endPoint = New Point(endX, endY)
        End Sub

        Public Sub New(startX As Double, startY As Double, endX As Double, endY As Double, startHinge As Boolean, endHinge As Boolean)
            Me.New()
            Me.startPoint = New Point(startX, startY)
            Me.endPoint = New Point(endX, endY)
            Me.StartHinge = startHinge
            Me.EndHinge = endHinge
        End Sub

        ''' <summary>
        ''' set true to draw the section width
        ''' </summary>
        ''' <returns></returns>
        Public Property drawSectionWidth As Boolean
            Get
                Return p_drawSectionWidth
            End Get
            Set(value As Boolean)
                p_drawSectionWidth = value
            End Set
        End Property

        Public Property EndHinge() As Boolean
            Get
                Return p_EndHinge
            End Get
            Set(value As Boolean)
                p_EndHinge = value
            End Set
        End Property

        Public Property HingeSize As size
            Get
                Return p_hingeSize
            End Get
            Set(value As size)
                p_hingeSize = value
            End Set
        End Property

        Public Property isInOpening As Boolean
            Get
                Return p_isInOpening
            End Get
            Set(value As Boolean)
                p_isInOpening = value
            End Set
        End Property

        ''' <summary>
        ''' Section width
        ''' </summary>
        ''' <returns></returns>
        Public Property SectionWidth As size
            Get
                Return p_sectionwidth
            End Get
            Set(value As size)
                p_sectionwidth = value
                If value Is Nothing Then
                    drawSectionWidth = False
                Else
                    drawSectionWidth = True
                End If
            End Set
        End Property

        ''' <summary>
        ''' Fill of section width
        ''' </summary>
        ''' <returns></returns>
        Public Property SectionWidth_Fill() As fill
            Get
                Return p_SectionWidth_Fill
            End Get
            Set(value As fill)
                p_SectionWidth_Fill = value
            End Set
        End Property

        ''' <summary>
        ''' boundary of section width
        ''' </summary>
        ''' <returns></returns>
        Public Property SectionWidth_Pen() As pen
            Get
                Return p_SectionWidth_Pen
            End Get
            Set(value As pen)
                p_SectionWidth_Pen = value
            End Set
        End Property
        Public Property StartHinge() As Boolean
            Get
                Return p_StartHinge
            End Get
            Set(value As Boolean)
                p_StartHinge = value
            End Set
        End Property

        Public Overrides Sub draw(contextobject As ContextObject, contextCoordinatesDelegate As Drawable.contextCoordinates, Optional contextSizeDelegate As Drawable.contextSize = Nothing) Implements Drawable.draw
            If isInOpening Then Return

            If TypeOf contextobject.Item Is System.Windows.Controls.Canvas Then

                Dim myCanvas As New System.Windows.Controls.Canvas
                myCanvas = TryCast(contextobject.Item, System.Windows.Controls.Canvas)

                With myCanvas
                    Dim line1 As New System.Windows.Shapes.Line

                    '
                    ' calculate coordinates in local system
                    ' by a call-back to the calling drawing object
                    '
                    Dim p1 As New Point
                    Dim p2 As New Point
                    Dim t As size
                    Dim s1 As New size
                    s1 = contextSizeDelegate(Me.HingeSize)

                    Dim vx As Double
                    Dim vy As Double
                    Dim tmp As Double

                    Dim nx As Double
                    Dim ny As Double
                    Dim sw As Double

                    p1 = contextCoordinatesDelegate(startPoint)
                    p2 = contextCoordinatesDelegate(endPoint)

                    t = contextSizeDelegate(Me.pen.size)
                    sw = contextSizeDelegate(SectionWidth).thickness

                    'If t.width > 0 and t.width < 0.1 Then t.width = 0.1

                    vx = p2.x - p1.x
                    vy = p2.y - p1.y

                    '
                    ' uniform vector
                    '
                    tmp = Math.Sqrt(vx ^ 2 + vy ^ 2)
                    If tmp <> 0 Then
                        vx = vx / tmp
                        vy = vy / tmp
                    End If

                    nx = -vy
                    ny = vx

                    If drawSectionWidth And SectionWidth.thickness > 0 Then
                        ' draw polygon
                        Dim ps1 As New Point
                        Dim ps2 As New Point
                        Dim ps3 As New Point
                        Dim ps4 As New Point
                        ' vector calculus
                        ps1.x = p1.x + nx * sw
                        ps1.y = p1.y + ny * sw

                        ps2.x = p2.x + nx * sw
                        ps2.y = p2.y + ny * sw

                        ps3.x = p2.x - nx * sw
                        ps3.y = p2.y - ny * sw

                        ps4.x = p1.x - nx * sw
                        ps4.y = p1.y - ny * sw

                        'draw polygon
                        Dim polygon1 As New System.Windows.Shapes.Polygon
                        With polygon1
                            .Points.Add(New Windows.Point(ps1.x, ps1.y))
                            .Points.Add(New Windows.Point(ps2.x, ps2.y))
                            .Points.Add(New Windows.Point(ps3.x, ps3.y))
                            .Points.Add(New Windows.Point(ps4.x, ps4.y))
                            .Points.Add(New Windows.Point(ps1.x, ps1.y))

                            'Dim b As Brush
                            'b = New System.Windows.Media.SolidColorBrush(Color.FromArgb(100, 255, 0, 0))
                            .Fill = Me.SectionWidth_Fill.Brush

                            '.Stroke = New System.Windows.Media.SolidColorBrush(Color.FromArgb(200, 100, 0, 0))
                            .Stroke = Me.SectionWidth_Pen.stroke
                            .StrokeThickness = 1

                        End With
                        myCanvas.Children.Add(polygon1)
                    End If

                    Dim startOffset As Double
                    Dim endOffset As Double

                    If Me.p_StartHinge Then
                        startOffset = s1.thickness
                    Else
                        startOffset = 0
                    End If

                    If Me.p_EndHinge Then
                        endOffset = s1.thickness
                    Else
                        endOffset = 0
                    End If

                    With line1
                        .X1 = p1.x + vx * startOffset
                        .Y1 = p1.y + vy * startOffset

                        .X2 = p2.x - vx * endOffset
                        .Y2 = p2.y - vy * endOffset
                        .Stroke = Me.pen.stroke

                        'Debug.Assert(Me.pen.color = Windows.Media.Colors.Black)

                        .StrokeThickness = t.width
                        .StrokeDashArray = Me.pen.dashArray
                    End With

                    myCanvas.Children.Add(line1)
                    If Me.p_StartHinge Then
                        Dim Ellipse1 As New System.Windows.Shapes.Ellipse
                        With Ellipse1
                            .Width = startOffset
                            .Height = startOffset
                            .Stroke = Me.pen.stroke
                            .StrokeThickness = t.width
                            .StrokeDashArray = Me.pen.dashArray
                            '.Fill = fill.Brush()

                            myCanvas.Children.Add(Ellipse1)
                            Windows.Controls.Canvas.SetLeft(Ellipse1, p2.x - vx * startOffset / 2 - .Width / 2)
                            Windows.Controls.Canvas.SetTop(Ellipse1, p2.y - vy * startOffset / 2 - .Height / 2)
                        End With
                    End If

                    If Me.p_EndHinge Then
                        Dim Ellipse1 As New System.Windows.Shapes.Ellipse
                        With Ellipse1
                            .Width = endOffset
                            .Height = endOffset
                            .Stroke = Me.pen.stroke
                            .StrokeThickness = t.width
                            .StrokeDashArray = Me.pen.dashArray
                            '.Fill = fill.Brush()

                            myCanvas.Children.Add(Ellipse1)
                            Windows.Controls.Canvas.SetLeft(Ellipse1, p2.x - vx * endOffset / 2 - .Width / 2)
                            Windows.Controls.Canvas.SetTop(Ellipse1, p2.y - vy * endOffset / 2 - .Height / 2)
                        End With

                    End If
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

        Private p_angle As Double
        Private p_fill As fill
        Private p_fontSize As New size

        Private p_horizontalAlignment As horizontalAlignment

        ' not used
        Private p_pen As pen
        Private p_position As New Point
        Private p_text As String
        Private p_verticalAlignment As verticalAlignment
        Private p_zindex As Long
        Public Sub New()
            p_fontSize.height = 3
            p_fontSize.Reference = Reference.contextMillimeters
            p_fill = New fill
            p_pen = New pen
            horizontalAlignment = horizontalAlignment.center
            verticalAlignment = verticalAlignment.center
        End Sub

        Public Sub New(x As Double, y As Double, text As String)
            Me.New()
            p_position.x = x
            p_position.y = y
            p_text = text
        End Sub

        Public Sub New(x As Double, y As Double, text As String, fontSize As Double, Optional fontSizeReference As Reference = Reference.contextMillimeters)
            Me.New()
            p_position.x = x
            p_position.y = y
            p_text = text
            p_fontSize.height = fontSize
            p_fontSize.Reference = fontSizeReference
        End Sub

        Public Sub New(x As Double, y As Double, coordinateReference As Reference, text As String)
            Me.New()
            p_position.x = x
            p_position.y = y
            p_position.coordinateReference = coordinateReference
            p_text = text
        End Sub

        Public Sub New(x As Double, y As Double, coordinateReference As Reference, text As String, fontSize As Double, Optional fontSizeReference As Reference = Reference.contextMillimeters)
            Me.New()
            p_position.x = x
            p_position.y = y
            p_position.coordinateReference = coordinateReference
            p_text = text
            p_fontSize.height = fontSize
            p_fontSize.Reference = fontSizeReference
        End Sub

        ''' <summary>
        ''' orientation angle in degrees
        ''' </summary>
        ''' <returns></returns>
        Public Property angle() As Double
            Get
                Return p_angle
            End Get
            Set(value As Double)
                p_angle = value
            End Set
        End Property

        ''' <summary>
        ''' Text fill
        ''' </summary>
        ''' <returns></returns>
        Public Property fill As fill
            Get
                Return p_fill
            End Get
            Set(value As fill)
                p_fill = value
            End Set
        End Property

        ''' <summary>
        ''' Font size
        ''' </summary>
        ''' <returns></returns>
        Public Property fontSize() As size
            Get
                Return p_fontSize
            End Get
            Set(value As size)
                p_fontSize = value
            End Set
        End Property

        ''' <summary>
        ''' horizontal alignment
        ''' </summary>
        ''' <returns></returns>
        Public Property horizontalAlignment As horizontalAlignment
            Get
                Return p_horizontalAlignment
            End Get
            Set(value As horizontalAlignment)
                p_horizontalAlignment = value
            End Set
        End Property

        ''' <summary>
        ''' Position
        ''' </summary>
        ''' <returns></returns>
        Public Property position() As Point
            Get
                Return p_position
            End Get
            Set(value As Point)
                p_position = value
            End Set
        End Property

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

        ''' <summary>
        ''' vertical alignment
        ''' </summary>
        ''' <returns></returns>
        Public Property verticalAlignment As verticalAlignment
            Get
                Return p_verticalAlignment
            End Get
            Set(value As verticalAlignment)
                p_verticalAlignment = value
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

        Public Function boundingRectangle(Optional getWordSize As Drawable.estimateWorldSize = Nothing) As Double() Implements Drawable.boundingRectangle
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

        Public Sub draw(contextobject As ContextObject, contextCoordinatesDelegate As Drawable.contextCoordinates, Optional contextSizeDelegate As Drawable.contextSize = Nothing) Implements Drawable.draw

            If TypeOf contextobject.Item Is System.Windows.Controls.Canvas Then
                Dim myCanvas As New System.Windows.Controls.Canvas
                myCanvas = TryCast(contextobject.Item, System.Windows.Controls.Canvas)

                With myCanvas
                    Dim tb As New Windows.Controls.TextBlock
                    Dim s1 As New size
                    s1 = contextSizeDelegate(Me.fontSize)

                    With tb
                        .Text = p_text
                        .FontSize = s1.height
                        Try
                            .Foreground = fill.Brush
                        Catch ex As Exception
                        End Try

                    End With

                    '
                    ' calculate text size
                    '
                    Dim myFormattedText As New FormattedText(tb.Text, CultureInfo.CurrentCulture, Windows.FlowDirection.LeftToRight,
                        New Typeface(tb.FontFamily, tb.FontStyle, tb.FontWeight, tb.FontStretch), s1.height + 0 * tb.FontSize, Brushes.Black,
                        New NumberSubstitution(), 1)
                    With myFormattedText
                        If Not Double.IsInfinity(tb.MaxHeight) Then
                            .MaxTextHeight = tb.MaxHeight
                        End If
                        If Not Double.IsInfinity(tb.MaxWidth) Then
                            .MaxTextWidth = tb.MaxWidth
                        End If
                        .Trimming = tb.TextTrimming
                    End With

                    tb.HorizontalAlignment = Windows.HorizontalAlignment.Center
                    tb.VerticalAlignment = Windows.VerticalAlignment.Center

                    tb.Margin = New Windows.Thickness(0)

                    Dim p As New Point

                    '
                    ' calculate coordinates in local system
                    ' by a call-back to the calling drawing object
                    '
                    p = contextCoordinatesDelegate(p_position)

                    Dim kx As Double
                    Dim ky As Double

                    Select Case Me.verticalAlignment
                        Case verticalAlignment.top
                            ky = 1.5
                        Case verticalAlignment.bottom
                            ky = 0.5 '-0.5
                        Case verticalAlignment.center
                            ky = 0.5
                    End Select

                    Select Case Me.horizontalAlignment
                        Case horizontalAlignment.left
                            kx = 1
                        Case horizontalAlignment.right
                            kx = 0
                        Case horizontalAlignment.center
                            kx = 0.5
                    End Select

                    Dim myRotateTransform As New Windows.Media.RotateTransform
                    myRotateTransform.Angle = -angle
                    tb.RenderTransform = myRotateTransform

                    myCanvas.Children.Add(tb)

                    myCanvas.SetLeft(tb, p.x - kx * myFormattedText.Width * Math.Cos(angle * Math.PI / 180) - ky * myFormattedText.Height * Math.Sin(angle * Math.PI / 180))
                    myCanvas.SetTop(tb, p.y + kx * myFormattedText.Width * Math.Sin(angle * Math.PI / 180) - ky * myFormattedText.Height * Math.Cos(angle * Math.PI / 180))

                    Dim line1 As New System.Windows.Shapes.Line
                    With line1
                        .X1 = p.x - kx * myFormattedText.Width * Math.Cos(angle * Math.PI / 180) + ky * myFormattedText.Height * Math.Sin(angle * Math.PI / 180)
                        .Y1 = p.y - kx * myFormattedText.Width * Math.Sin(angle * Math.PI / 180) - ky * myFormattedText.Height * Math.Cos(angle * Math.PI / 180)
                        .X2 = p.x + kx * myFormattedText.Width * Math.Cos(angle * Math.PI / 180) - ky * myFormattedText.Height * Math.Sin(angle * Math.PI / 180)
                        .Y2 = p.y + kx * myFormattedText.Width * Math.Sin(angle * Math.PI / 180) + ky * myFormattedText.Height * Math.Cos(angle * Math.PI / 180)
                    End With
                    myCanvas.Children.Add(line1)
                End With
            Else
                Throw New NotImplementedException()
            End If
        End Sub

    End Class

#End Region

#Region "Force and Moment"

    <ComVisible(False)>
    Public Class ForceArrow
        Implements Drawable
        Const DEFAULT_ARROW_HEIGHT_MM As Double = 5

        Const DEFAULT_ARROW_WIDTH_MM As Double = 3

        Const DEFAULT_FORCE_SIZE_MM As Double = 10

        Private p_angle As Double

        Private p_arrowSize As New size

        Private p_ArrowTipType As arrowTipTypes

        Private p_caption As String

        Private p_endPointDisplay As WMDraw.PointDisplay

        Private p_F As Double

        Private p_forceDirection As forceDirections

        Private p_ForceType As forceTypes

        Private p_Fx As Double

        ' input Fx
        Private p_Fy As Double

        ' input F
        ' input alpha
        Private p_inputFAlpha As Boolean

        ' input Fy
        Private p_inputFxFy As Boolean

        Private p_isMaximumForceDefined As Boolean

        Private p_maximumForce As Double

        Private p_maxSize As New size

        Private p_Offset As New size

        Private p_pen As pen

        Private p_tipPoint As New Point

        Private p_zIndex As Long

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

        Public Enum arrowTipTypes
            defaultForceTip
            halfTipLeft
            halfTipRight
            directionIndicatorLeft
            directionIndicatorRight
        End Enum

        Public Enum endPointTypes
            none
            circle
            cross
        End Enum

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
        Public Property MaximumSize() As size
            Get
                Return p_maxSize
            End Get
            Set(value As size)
                p_maxSize = value
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
        ''' Force tip point Offset Vector
        ''' </summary>
        ''' <returns></returns>
        Public Property offset As size
            Get
                Return p_Offset
            End Get
            Set(value As size)
                p_Offset = value
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
            myMath.includeInBoundingRectangle(r, Me.tipPoint)
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
            Dim arrowSize1 As New size    ' arrow
            Dim t As New size             ' pen Thickness

            If TypeOf contextobject.Item Is System.Windows.Controls.Canvas Then
                Dim myCanvas As New System.Windows.Controls.Canvas
                myCanvas = TryCast(contextobject.Item, System.Windows.Controls.Canvas)

                With myCanvas
                    '
                    ' Force
                    '
                    pTip = contextCoordinatesDelegate(Me.tipPoint)
                    fSize = contextSizeDelegate(Me.p_maxSize)
                    arrowSize1 = contextSizeDelegate(Me.arrowSize)

                    t = contextSizeDelegate(Me.pen.size)
                    If t.width > 0 And t.width < 0.1 Then t.width = 0.1

                    Dim myPolyline As New System.Windows.Shapes.Polyline

                    Dim gl1 As New Geom.Line
                    Dim gl2 As New Geom.Line

                    With myPolyline
                        '
                        ' shaft
                        '

                        '
                        ' draw in x-position
                        '
                        If p_isMaximumForceDefined And maximumForce <> 0 Then
                            pEnd.x = pTip.x + fSize.width * F / maximumForce
                            pEnd.y = pTip.y
                        Else
                            pEnd.x = pTip.x + fSize.width * F / F
                            pEnd.y = pTip.y
                        End If

                        If Me.p_Offset.average <> 0 Then
                            pOffsetCU = contextSizeDelegate(Me.p_Offset)
                            gl1 = New Geom.Line(pTip.x, pTip.y, pEnd.x - pTip.x, pEnd.y - pTip.y)
                            gl2 = gl1.OffsetLine(pOffsetCU.height)
                            Debug.Print(String.Format("pTip=({0}/{1}", pTip.x, pTip.y))
                            Debug.Print(String.Format("pEnd=({0}/{1}", pEnd.x, pEnd.y))
                            pTip.x = gl2.StartPointX
                            pTip.y = gl2.StartPointY
                            pEnd.x = gl2.StartPointX + gl2.dx
                            pEnd.y = gl2.StartPointY + gl2.dy
                            Debug.Print(String.Format("after offset by {0}", pOffsetCU.height))
                            Debug.Print(String.Format("pTip=({0}/{1}", pTip.x, pTip.y))
                            Debug.Print(String.Format("pEnd=({0}/{1}", pEnd.x, pEnd.y))
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

    End Class


    ''' <summary>
    ''' Moment Arc as a drawable item
    ''' </summary>
    <ComVisible(False)>
    Public Class MomentArc
        Implements Drawable
        Const DEFAULT_ARROW_HEIGHT_MM As Double = 5

        '
        ' todo: Create a force or arrow drawable and a Moment drawable
        '
        Const DEFAULT_ARROW_WIDTH_MM As Double = 3

        Const DEFAULT_END_ANGLE As Double = 150
        Const DEFAULT_RADIUS_MM As Double = 8

        Const DEFAULT_START_ANGLE As Double = 30

        ''' <summary>
        ''' Arrow Size
        ''' </summary>
        Private p_arrowSize As New size

        Private p_caption As String
        Private p_EndAngle As Double
        Private p_midPoint As New Point
        Private p_pen As pen
        Private p_Radius As New size
        Private p_StartAngle As Double
        Private p_zIndex As Long

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

        Public Property pen As pen Implements Drawable.pen
            Get
                Return p_pen
            End Get
            Set(value As pen)
                p_pen = value
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
            If midPoint.coordinateReference = Reference.world And p_Radius.Reference = Reference.world Then
                For i As Integer = 0 To 2
                    p(i) = New Point
                    p(i).x = midPoint.x + p_Radius.max * Math.Cos(myMath.radians(alpha(i)))
                    p(i).y = midPoint.y + p_Radius.max * Math.Sin(myMath.radians(alpha(i)))
                    myMath.includeInBoundingRectangle(r, p(i))
                Next
                Return r
            Else
                Return Me.midPoint.boundingRectangle()
            End If

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
                    If t.width > 0 And t.width < 0.1 Then t.width = 0.1

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

    End Class

#End Region

#Region "Polygon Class"

    Public Class Polygon
        Implements Drawable

        Private p_closed As Boolean
        Private p_fill As fill
        Private p_fontFill As fill
        Private p_fontSize As New size
        ' {0} ... node number
        ' {1} ... x coordinate
        ' {2} ... y coordinate
        Private p_nodeLabelFormat As String

        Private p_pen As pen
        Private p_points As New List(Of Point)
        Private p_zIndex As Long
        Public Sub New()
            p_points = New List(Of Point)
            p_pen = New pen
            p_fill = New fill
        End Sub

        ''' <summary>
        ''' is polygon closed?
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property closed()
            Get
                Return (Me.points.First = Me.points.Last)
            End Get
        End Property

        ''' <summary>
        ''' Fill
        ''' </summary>
        ''' <returns></returns>
        Public Property fill As fill
            Get
                Return p_fill
            End Get
            Set(value As fill)
                p_fill = value
            End Set
        End Property

        ''' <summary>
        ''' Text fill
        ''' </summary>
        ''' <returns></returns>
        Public Property fontFill As fill
            Get
                Return p_fontFill
            End Get
            Set(value As fill)
                p_fontFill = value
            End Set
        End Property

        ''' <summary>
        ''' Font size
        ''' </summary>
        ''' <returns></returns>
        Public Property fontSize() As size
            Get
                Return p_fontSize
            End Get
            Set(value As size)
                p_fontSize = value
            End Set
        End Property

        ''' <summary>
        ''' define a node label format
        ''' use placeholders: {0} for node number {1:f2} for x coordinate {2:f2} for y coordinate
        ''' </summary>
        ''' <returns></returns>
        Public Property nodeLabelFormat() As String
            Get
                Return p_nodeLabelFormat
            End Get
            Set(value As String)
                Dim tmp As String
                Try
                    tmp = String.Format(value, 1, 10.1, 20.2)
                Catch ex As Exception
                    Throw ex
                End Try
                p_nodeLabelFormat = value
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

        Public Property points As List(Of Point)
            Get
                Return p_points
            End Get
            Set(value As List(Of Point))
                p_points = value
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

            For Each myPoint In Me.points
                myMath.includeInBoundingRectangle(r, myPoint.x, myPoint.y)
            Next
            Return r
        End Function

        ''' <summary>
        ''' close polygon by adding first point to the end of points
        ''' </summary>
        Public Sub close()
            If Not Me.closed Then
                Me.points.Add(Me.points.First)
            End If
        End Sub
        Public Function CompareTo(other As Drawable) As Integer Implements IComparable(Of Drawable).CompareTo
            Throw New NotImplementedException()
        End Function

        Public Sub draw(contextobject As ContextObject, contextCoordinatesDelegate As Drawable.contextCoordinates, Optional contextSizeDelegate As Drawable.contextSize = Nothing) Implements Drawable.draw

            If TypeOf contextobject.Item Is System.Windows.Controls.Canvas Then
                Dim myCanvas As New System.Windows.Controls.Canvas
                myCanvas = TryCast(contextobject.Item, System.Windows.Controls.Canvas)
                With myCanvas
                    Dim polygon1

                    If Me.closed Then
                        polygon1 = New System.Windows.Shapes.Polygon
                    Else
                        polygon1 = New System.Windows.Shapes.Polyline
                    End If

                    With polygon1
                        Dim p1 As New Point
                        Dim t As size

                        t = contextSizeDelegate(Me.pen.size)

                        For Each myPoint In Me.points
                            p1 = contextCoordinatesDelegate(myPoint)
                            .Points.Add(New Windows.Point(x:=p1.x, y:=p1.y))
                        Next

                        .Stroke = Me.pen.stroke
                        .StrokeThickness = t.width
                        .StrokeDashArray = Me.pen.dashArray
                        .Fill = fill.Brush()

                        If t.width > 0 And t.width < 0.1 Then t.width = 0.1

                    End With

                    myCanvas.Children.Add(polygon1)

                    Dim tmp As String
                    Dim i As Integer

                    Try
                        tmp = ""
                        tmp = String.Format(nodeLabelFormat(), 1, 10.1, 20.2)
                        If Len(tmp) > 0 Then

                            Dim p1 As New Point
                            Dim tb As New Windows.Controls.TextBlock
                            Dim s1 As New size
                            s1 = contextSizeDelegate(Me.fontSize)

                            For Each myPoint In Me.points
                                i += 1
                                tmp = String.Format(nodeLabelFormat(), i, myPoint.x, myPoint.y)
                                If Len(tmp) > 0 Then
                                    p1 = contextCoordinatesDelegate(myPoint)
                                    tb = New Windows.Controls.TextBlock
                                    tb.Text = tmp
                                    tb.FontSize = s1.height
                                    tb.Foreground = fontFill.Brush

                                    myCanvas.Children.Add(tb)
                                    myCanvas.SetLeft(tb, p1.x)
                                    myCanvas.SetTop(tb, p1.y)
                                End If
                            Next

                        End If
                    Catch ex As Exception

                    End Try

                End With
            Else
                Throw New NotImplementedException()
            End If

        End Sub

        Public Sub rotate(center As Point, angle As Double)
            For i = 0 To Me.points.Count - 1
                Dim myPoint As Point = points.ElementAt(i)
                points(i) = myPoint.rotate(center, angle * Math.PI / 180)
            Next
        End Sub

    End Class

#End Region

#Region "Line Class"

    ''' <summary>
    ''' Line as a drawable item
    ''' </summary>
    <ComVisible(False)>
    Public Class Line
        Implements Drawable

        ''' <summary>
        ''' End Point
        ''' </summary>
        Private p_endPoint As Point

        Private p_pen As pen

        ''' <summary>
        ''' Start Point
        ''' </summary>
        Private p_startPoint As Point

        Private p_zIndex As Long

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

        Public Property startPoint As Point
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

        Public ReadOnly Property innerPoint(xi As Double) As Point
            Get
                innerPoint = New Point(startPoint.x + (endPoint.x - startPoint.x) * xi, startPoint.y + (endPoint.y - startPoint.y) * xi)
            End Get
        End Property

        Public Function boundingRectangle(Optional getWordSize As Drawable.estimateWorldSize = Nothing) As Double() Implements Drawable.boundingRectangle

            Dim r(3) As Double

            Dim p1 As New Point
            Dim p2 As New Point

            If startPoint.coordinateReference = Reference.world Then
                r(0) = startPoint.x
                r(1) = startPoint.y
            End If

            If startPoint.coordinateReference = Reference.world Then
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
                        If t.width > 0 And t.width < 0.1 Then t.width = 0.1

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

        Public Function rotate(center As Point, angle As Double) As Line
            rotate = Me
            rotate.startPoint = Me.startPoint.rotate(center, angle)
            rotate.endPoint = Me.endPoint.rotate(center, angle)
        End Function

    End Class

#End Region

#Region "Rectangle Class"

    ''' <summary>
    ''' as a drawable item
    ''' </summary>
    <ComVisible(False)>
    Public Class Rectangle
        Implements Drawable

        ''' <summary>
        ''' End Point
        ''' </summary>
        Private p_endPoint As Point

        Private p_fill As fill
        Private p_pen As pen

        ''' <summary>
        ''' Start Point
        ''' </summary>
        Private p_startPoint As Point

        Private p_zIndex As Long

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

        Public Property endPoint As Point
            Get
                Return p_endPoint
            End Get
            Set(value As Point)
                p_endPoint = value
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

        Public Property pen As pen Implements Drawable.pen
            Get
                Return p_pen
            End Get
            Set(value As pen)
                p_pen = value
            End Set
        End Property

        Public Property startPoint As Point
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

        Private ReadOnly Property height()
            Get
                Return Math.Abs(endPoint.y - startPoint.y)
            End Get
        End Property

        Private ReadOnly Property left()
            Get
                Return Math.Min(startPoint.x, endPoint.x)
            End Get
        End Property

        Private ReadOnly Property top()
            Get
                Return Math.Min(startPoint.y, endPoint.y)
            End Get
        End Property

        Private ReadOnly Property width()
            Get
                Return Math.Abs(endPoint.x - startPoint.x)
            End Get
        End Property

        Public Function boundingRectangle(Optional getWordSize As Drawable.estimateWorldSize = Nothing) As Double() Implements Drawable.boundingRectangle

            Dim r(3) As Double

            Dim p1 As New Point
            Dim p2 As New Point

            If startPoint.coordinateReference = Reference.world Then
                r(0) = startPoint.x
                r(1) = startPoint.y
            End If

            If startPoint.coordinateReference = Reference.world Then
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
                        If t.width > 0 And t.width < 0.1 Then t.width = 0.1

                        .Width = Math.Abs(p2.x - p1.x)
                        .Height = Math.Abs(p2.y - p1.y)

                        'todo: offer fill solid and hatch

                        If t.width = 0 Then
                            .Stroke = Nothing
                            .StrokeThickness = 0
                        Else
                            .Stroke = Me.pen.stroke
                            .StrokeThickness = t.width
                            .StrokeDashArray = Me.pen.dashArray
                        End If

                        If Not fill Is Nothing Then
                            .Fill = fill.Brush()
                        End If

                        myCanvas.Children.Add(rect1)
                        Windows.Controls.Canvas.SetLeft(rect1, Math.Min(p1.x, p2.x))
                        Windows.Controls.Canvas.SetTop(rect1, Math.Min(p1.y, p2.y))
                    End With

                End With
            Else
                Throw New NotImplementedException()
            End If

        End Sub

        Public Function toPolygon() As Polygon
            Dim p As New Polygon
            p.points.Add(New Point(Me.startPoint.x, Me.startPoint.y))
            p.points.Add(New Point(Me.startPoint.x, Me.endPoint.y))
            p.points.Add(New Point(Me.endPoint.x, Me.endPoint.y))
            p.points.Add(New Point(Me.startPoint.y, Me.endPoint.y))
            p.close()
            Return p
        End Function
    End Class

#End Region

#Region "Screw Class"

    ''' <summary>
    ''' Screw
    ''' </summary>
    Public Class Screw
        Implements Drawable

        Private p_angle As Double
        Private p_caption As String
        Private p_HeadPosition As New Point
        Private p_l As Double
        Private p_lg As Double
        Private p_pen As pen
        Private p_zIndex As Long

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

        Public Function boundingRectangle(Optional getWordSize As Drawable.estimateWorldSize = Nothing) As Double() Implements Drawable.boundingRectangle
            Throw New NotImplementedException()
        End Function

        Public Function CompareTo(other As Drawable) As Integer Implements IComparable(Of Drawable).CompareTo
            Throw New NotImplementedException()
        End Function

        Public Sub draw(contextobject As ContextObject, contextCoordinatesDelegate As Drawable.contextCoordinates, Optional contextSizeDelegate As Drawable.contextSize = Nothing) Implements Drawable.draw
            Throw New NotImplementedException()
            'todo: not implemented
        End Sub

    End Class

#End Region

#Region "Support Class"

    ''' <summary>
    ''' Support symbol
    ''' </summary>
    <ComVisible(False)>
    Public Class Support
        Implements Drawable

        Const SUPPORT_ANGLE_DEG As Double = 30
        Const SUPPORT_HEIGHT_MM As Double = 7
        Const SUPPORT_WIDTH_MM As Double = 7
        Private p_angle As Double
        Private p_caption As String

        ' stiffness values
        Private p_cx As Double

        Private p_cy As Double
        Private p_pen As pen
        Private p_position As New Point
        Private p_size As New size
        Private p_zIndex As Long

        ''' <summary>
        ''' New Support
        ''' </summary>
        ''' <param name="x">Point X-Coordinate</param>
        ''' <param name="y">Point Y-Coordinate</param>
        ''' <param name="angle">Rotation Angle</param>
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
        ''' New Support
        ''' </summary>
        ''' <param name="x">Point X-Coordinate</param>
        ''' <param name="y">Point Y-Coordinate</param>
        ''' <param name="angle">Rotation Angle</param>
        ''' <param name="caption">Caption</param>
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
        ''' New Support
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

        ''' <summary>
        ''' rotation angle in degrees
        ''' </summary>
        ''' <returns>
        ''' </returns>
        Public Property angle As Double
            Get
                Return p_angle
            End Get
            Set(value As Double)
                p_angle = value
            End Set
        End Property

        ''' <summary>
        ''' Caption
        ''' </summary>
        ''' <returns>
        ''' Support Caption
        ''' </returns>
        Public Property caption As String
            Get
                Return p_caption
            End Get
            Set(value As String)
                p_caption = value
            End Set
        End Property

        ''' <summary>
        ''' Pen
        ''' </summary>
        ''' <returns></returns>
        Public Property pen As pen Implements Drawable.pen
            Get
                Return p_pen
            End Get
            Set(value As pen)
                p_pen = value
            End Set
        End Property

        ''' <summary>
        ''' Position
        ''' </summary>
        ''' <returns>
        ''' Point
        ''' </returns>
        Public Property position As Point
            Get
                Return p_position
            End Get
            Set(value As Point)
                p_position = value
            End Set
        End Property

        ''' <summary>
        ''' Size
        ''' </summary>
        ''' <returns></returns>
        Public Property size As size
            Get
                Return p_size
            End Get
            Set(value As size)
                p_size = value
            End Set
        End Property

        ''' <summary>
        ''' Z-Index
        ''' </summary>
        ''' <returns></returns>
        Public Property zIndex As Long Implements Drawable.zIndex
            Get
                Return p_zIndex
            End Get
            Set(value As Long)
                p_zIndex = value
            End Set
        End Property

        ''' <summary>
        ''' Bounding Rectangle
        ''' </summary>
        '''
        ''' <returns>
        ''' The bounding rectangle - momentarily only the position point
        ''' </returns>
        Public Function boundingRectangle(Optional getWordSize As Drawable.estimateWorldSize = Nothing) As Double() Implements Drawable.boundingRectangle
            Return Me.position.boundingRectangle()
        End Function

        ''' <summary>
        ''' Compare Z-Order
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
        ''' Draw Support on destination device
        ''' </summary>
        ''' <param name="contextobject"></param>
        ''' <param name="contextCoordinatesDelegate"></param>
        ''' <param name="contextSizeDelegate"></param>
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

    End Class

#End Region

#Region "Ellipse Class"

    ''' <summary>
    ''' Ellipse
    ''' </summary>
    <ComVisible(False)>
    Public Class Ellipse

        Implements Drawable

        Private p_angle As Double
        Private p_fill As fill
        Private p_midpoint As Point

        Private p_offsetDirection As wmg.Vector
        Private p_offsetSize As size

        Private p_pen As pen
        Private p_radii As size
        Private p_zIndex As Long

        Sub New()
            Me.midPoint = New Point
            Me.radii = New size
            Me.angle = 0
            p_fill = New fill
            p_fill.color = Colors.Transparent
            p_pen = New pen
            p_offsetDirection = New wmg.Vector(0, 0)
            p_offsetSize = New size(0)
        End Sub

        Sub New(midPoint As Point, radii As size)
            Me.New()
            Me.midPoint = midPoint
            Me.radii = radii
        End Sub

        Sub New(midPoint As Point, radii As size, angle As Double)
            Me.New()
            Me.midPoint = midPoint
            Me.radii = radii
            Me.angle = angle
        End Sub

        Sub New(midpoint As Point, circularRadius As Double)
            Me.New()
            Me.midPoint = midpoint
            Me.radii.width = circularRadius
            Me.radii.height = circularRadius
            Me.radii.Reference = midpoint.coordinateReference
        End Sub

        ''' <summary>
        ''' Rotation angle
        ''' </summary>
        ''' <returns></returns>
        Public Property angle() As Double
            Get
                Return p_angle
            End Get
            Set(value As Double)
                p_angle = value
            End Set
        End Property

        ''' <summary>
        ''' Fill
        ''' </summary>
        ''' <returns></returns>
        Public Property fill As fill
            Get
                Return p_fill
            End Get
            Set(value As fill)
                p_fill = value
            End Set
        End Property

        ''' <summary>
        ''' Mid-Point
        ''' </summary>
        ''' <returns>
        ''' Point Object
        ''' </returns>
        Public Property midPoint As Point
            Get
                Return p_midpoint
            End Get
            Set(value As Point)
                p_midpoint = value
            End Set
        End Property

        ''' <summary>
        ''' Ellipse offset direction vector
        ''' </summary>
        ''' <returns></returns>
        Public Property offsetDirection() As wmg.Vector
            Get
                Return p_offsetDirection
            End Get
            Set(value As wmg.Vector)
                If value.length = 0 Then
                    p_offsetDirection = value
                Else
                    p_offsetDirection = value.unitVector
                End If
            End Set
        End Property

        ''' <summary>
        ''' Ellipse offset size
        ''' </summary>
        ''' <returns></returns>
        Public Property offsetSize() As size
            Get
                Return p_offsetSize
            End Get
            Set(value As size)
                p_offsetSize = value
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

        ''' <summary>
        ''' Radii
        ''' </summary>
        ''' <returns></returns>
        Public Property radii As size
            Get
                Return p_radii
            End Get
            Set(value As size)
                p_radii = value
            End Set
        End Property

        ''' <summary>
        ''' Z-Index
        ''' </summary>
        ''' <returns></returns>
        Public Property zIndex As Long Implements Drawable.zIndex
            Get
                Return p_zIndex
            End Get
            Set(value As Long)
                p_zIndex = value
            End Set
        End Property

        ''' <summary>
        ''' Bounding Rectangle
        ''' </summary>
        '''
        ''' <returns></returns>
        Public Function boundingRectangle(Optional getWordSize As Drawable.estimateWorldSize = Nothing) As Double() Implements Drawable.boundingRectangle
            Dim r(3) As Double

            'todo: consider offset in bounding box of ellipse - if possible (?)

            'If offsetSize.max = 0 Then
            r(0) = Me.midPoint.x - Me.radii.width
            r(1) = Me.midPoint.y - Me.radii.height
            r(2) = Me.midPoint.x + Me.radii.width
            r(3) = Me.midPoint.y + Me.radii.height
            'Else

            'End If

            Return r
        End Function

        ''' <summary>
        ''' Compare z-Index
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
        ''' Draw Ellipse
        ''' </summary>
        ''' <param name="contextobject"></param>
        ''' <param name="contextCoordinatesDelegate"></param>
        ''' <param name="contextSizeDelegate"></param>
        Public Sub draw(contextobject As ContextObject, contextCoordinatesDelegate As Drawable.contextCoordinates, Optional contextSizeDelegate As Drawable.contextSize = Nothing) Implements Drawable.draw

            If TypeOf contextobject.Item Is System.Windows.Controls.Canvas Then
                Dim myCanvas As New System.Windows.Controls.Canvas
                myCanvas = TryCast(contextobject.Item, System.Windows.Controls.Canvas)

                Dim Ellipse1 As New System.Windows.Shapes.Ellipse
                Dim p1 As New Point
                Dim r As New size
                Dim t As size

                With myCanvas

                    With Ellipse1
                        '
                        ' calculate coordinates in local system
                        ' by a call-back to the calling drawing object
                        '
                        p1 = contextCoordinatesDelegate(midPoint)
                        r = contextSizeDelegate(radii)

                        t = contextSizeDelegate(Me.pen.size)
                        If t.width > 0 And t.width < 0.1 Then t.width = 0.1

                        .Width = 2 * r.width
                        .Height = 2 * r.height

                        .Stroke = Me.pen.stroke
                        .StrokeThickness = t.width
                        .StrokeDashArray = Me.pen.dashArray

                        .Fill = fill.Brush()
                    End With

                    If angle <> 0 Then
                        Dim myRotateTransform As New Windows.Media.RotateTransform
                        myRotateTransform.Angle = -angle
                        myRotateTransform.CenterX = r.width
                        myRotateTransform.CenterY = r.height
                        Ellipse1.RenderTransform = myRotateTransform
                    End If

                    myCanvas.Children.Add(Ellipse1)

                    If offsetSize.max = 0 Then
                        Windows.Controls.Canvas.SetLeft(Ellipse1, p1.x - r.width)
                        Windows.Controls.Canvas.SetTop(Ellipse1, p1.y - r.height)
                    Else
                        ' calculate offset position
                        Dim t2 As size
                        Dim off As New wmg.Vector
                        t2 = contextSizeDelegate(offsetSize)
                        off = offsetDirection.multiply(t2.max)

                        '' for debugging: draw a line at the original position
                        ''

                        'Dim line1 As New System.Windows.Shapes.Line
                        'With line1

                        '    '
                        '    ' calculate coordinates in local system
                        '    ' by a call-back to the calling drawing object
                        '    '
                        '    If t.width > 0 and t.width < 0.1 Then t.width = 0.1

                        '    .X1 = p1.x - 1
                        '    .Y1 = p1.y - 1
                        '    .X2 = p1.x + 1
                        '    .Y2 = p1.y + 1
                        '    .Stroke = Me.pen.stroke
                        '    .StrokeThickness = t.width
                        '    .StrokeDashArray = Me.pen.dashArray
                        'End With
                        'myCanvas.Children.Add(line1)

                        Windows.Controls.Canvas.SetLeft(Ellipse1, p1.x - r.width + off.x)
                        Windows.Controls.Canvas.SetTop(Ellipse1, p1.y - r.height + off.y)
                    End If

                End With
            Else
                Throw New NotImplementedException()
            End If
        End Sub

    End Class

#End Region

#Region "Pen Class"

    ''' <summary>
    ''' Pen
    ''' </summary>
    <ComVisible(False)>
    Public Class pen

        ' todo: allow transparency

        'Public stroke As System.Windows.Media.Brush

        Public Const DASHARRAY_DASHDOT = "5,1,1,1"
        Public Const DASHARRAY_DASHED_LONG = "6,3"
        Public Const DASHARRAY_DASHED_MEDIUM = "5,3"
        Public Const DASHARRAY_DASHED_SHORT = "3,2"
        Public Const DASHARRAY_POINTS = "1,1"
        Public Const DASHARRAY_SOLID = ""
        Private p_color As New System.Windows.Media.Color
        Private p_dasharray As System.Windows.Media.DoubleCollection

        Private p_opacity As Double

        'Private p_thicknessReference As Reference
        Private p_thickness As size

        ''' <summary>
        ''' New
        ''' </summary>
        Public Sub New()
            p_color = System.Windows.Media.Colors.Black
            p_thickness = New size
            p_thickness.width = 2
            p_thickness.Reference = Reference.contextUnits
            p_opacity = 1
            Me.dashString = DASHARRAY_SOLID
        End Sub

        ''' <summary>
        ''' color
        ''' </summary>
        ''' <returns>
        ''' Color
        ''' </returns>
        Public Property color As System.Windows.Media.Color
            Get
                Return p_color
            End Get
            Set(value As System.Windows.Media.Color)
                p_color = value
            End Set
        End Property

        ''' <summary>
        ''' Dash array as collection of double values
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property dashArray() As System.Windows.Media.DoubleCollection
            Get
                Return p_dasharray
            End Get
        End Property

        ''' <summary>
        ''' Dash String: , - separated list of dash, space - values
        ''' </summary>
        ''' <returns></returns>
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

        ''' <summary>
        ''' Opacity factor
        ''' </summary>
        ''' <returns></returns>
        Public Property opacity As Double
            Get
                Return p_opacity
            End Get
            Set(value As Double)
                p_opacity = value
            End Set
        End Property

        ''' <summary>
        ''' Size equals thickness
        ''' </summary>
        ''' <returns></returns>
        Public Property size As size
            Get
                Return p_thickness
            End Get
            Set(value As size)
                p_thickness = value
            End Set
        End Property

        ''' <summary>
        ''' Stroke as a brush object with defined color
        ''' </summary>
        ''' <returns>
        ''' Brush
        ''' </returns>
        Public ReadOnly Property stroke As System.Windows.Media.Brush
            Get
                stroke = New System.Windows.Media.SolidColorBrush(Me.color)
                stroke.Opacity = p_opacity

                Return stroke
            End Get
        End Property

        ''' <summary>
        ''' Brush thickness in reference units as set in thicknessReference
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

    End Class

#End Region

#Region "Fill Class"

    <ComVisible(False)>
    Public Class fill
        Private p_angle As Double

        Private p_color As System.Windows.Media.Color

        Private p_filltype As fillType

        Private p_opacity As Double

        Private p_spacing As Double

        Private p_thickness As Double

        ''' <summary>
        ''' New
        ''' </summary>
        Public Sub New()
            p_color = Colors.Black
            p_thickness = 1

            p_opacity = 1
            p_spacing = 5

            p_angle = 45
        End Sub

        ''' <summary>
        ''' Fill type
        ''' </summary>
        Public Enum fillType
            solidColor
            linearHatching
            empty
        End Enum

        ''' <summary>
        ''' Fill Brush depending on fillType
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property Brush() As Brush
            Get
                Select Case Me.p_filltype
                    Case fillType.empty
                        Return Nothing
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

        ''' <summary>
        ''' Color
        ''' </summary>
        ''' <returns></returns>
        Public Property color() As Color
            Get
                Return p_color
            End Get
            Set(value As Color)
                p_color = value
            End Set
        End Property

        Public Property type As fillType
            Get
                Return p_filltype
            End Get
            Set(value As fillType)
                p_filltype = value
            End Set
        End Property
        ''' <summary>
        ''' set filltype Linear Hatch
        ''' </summary>
        ''' <param name="color"></param>
        ''' <param name="opacity"></param>
        ''' <param name="hatchThickness_mm"></param>
        ''' <param name="hatchSpacing_mm"></param>
        ''' <param name="hatchAngle"></param>
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

        ''' <summary>
        ''' Set solid fill
        ''' </summary>
        ''' <param name="color"></param>
        Public Sub SetSolidColor(color As Color)
            p_filltype = fillType.solidColor
            p_color = color
        End Sub
        Public Property opacity As Double
            Get
                Return p_opacity
            End Get
            Set(value As Double)
                p_opacity = value
            End Set
        End Property
    End Class

#End Region
#Region "ScaleBar"
    <ComVisible(False)>
    Public Class ScaleBar
        Implements Drawable

        Private p_position As Point
        Private p_alignment As horizontalAlignment

        Private p_fill As fill
        Private p_pen As New pen
        Private p_size As New size
        Private p_barThickness As New size

        Private p_zIndex As Long
        Private p_nDiv As Integer

        Public Sub New()
            p_position = New Point
            p_position = New Point(1, 0)
            p_position.coordinateReference = Reference.contextFraction

            p_alignment = horizontalAlignment.right

            p_fill = New fill
            p_fill.color = WMColors.Black
            p_fill.opacity = 0.6

            p_pen = New pen
            p_pen.opacity = 0.6

            p_size = New size(30, 30, Reference.contextMillimeters)
            p_barThickness = New size(1.5, 1.5, Reference.contextMillimeters)
            p_nDiv = 5
        End Sub

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
        ''' subdivisions for scale bar
        ''' </summary>
        ''' <returns></returns>
        Public Property nDiv As Integer
            Get
                Return p_nDiv
            End Get
            Set(value As Integer)
                If value < 2 Then value = 2
                p_nDiv = value
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
        ''' <summary>
        ''' Size
        ''' </summary>
        ''' <returns></returns>
        Public Property size As size
            Get
                Return p_size
            End Get
            Set(value As size)
                p_size = value
            End Set
        End Property


        ''' <summary>
        ''' Bar thicknes
        ''' </summary>
        ''' <returns></returns>
        Public Property barThickness As size
            Get
                Return p_barThickness
            End Get
            Set(value As size)
                p_barThickness = value
            End Set
        End Property
        Public Sub draw(contextobject As ContextObject, contextCoordinatesDelegate As Drawable.contextCoordinates, Optional contextSizeDelegate As Drawable.contextSize = Nothing) Implements Drawable.draw

            If TypeOf contextobject.Item Is System.Windows.Controls.Canvas Then
                Dim myCanvas As New System.Windows.Controls.Canvas
                myCanvas = TryCast(contextobject.Item, System.Windows.Controls.Canvas)
                With myCanvas

                    'Dim rect1 As New System.Windows.Shapes.Rectangle

                    'Maßstabsbalken (Skala)

                    Dim p1 As New Point
                    Dim p2 As New Point
                    Dim t As size
                    Dim tSize As size
                    Dim tBar As size

                    '
                    ' calculate coordinates in local system
                    ' by a call-back to the calling drawing object
                    '
                    ' world coordinates
                    p1 = contextCoordinatesDelegate(New Point(0, 0))
                    p2 = contextCoordinatesDelegate(New Point(1, 1))

                    ' one world unit in x-direction 
                    ' is ux drawing units
                    Dim ux As Double = p2.x - p1.x
                    Dim uy As Double = p2.y - p1.y

                    tSize = contextSizeDelegate(Me.size)

                    ' one significant only!
                    ' world units length of the scalebar
                    Dim wux As Double = Math.Floor(tSize.width / ux)

                    wux = WMHelpers.SetSigFigs(wux, 1)
                    Dim oneTickX As Double = wux / nDiv * ux

                    tBar = contextSizeDelegate(Me.barThickness)

                    ' pen size
                    t = contextSizeDelegate(Me.pen.size)
                    If t.width > 0 And t.width < 0.1 Then t.width = 1

                    ' position
                    p1 = contextCoordinatesDelegate(p_position)

                    Select Case p_alignment
                        Case horizontalAlignment.left
                            ' keep position
                        Case horizontalAlignment.center
                            p1.x = p1.x + nDiv * oneTickX / 2
                        Case horizontalAlignment.right
                            p1.x = p1.x - nDiv * oneTickX
                    End Select

                    For i As Integer = 1 To nDiv
                        Dim rect1 As New System.Windows.Shapes.Rectangle

                        Dim tt As New Text
                        tt.fontSize = New size(0.75 * tBar.height, Reference.asIs)
                        tt.position = New Point(p1.x + i * oneTickX, p1.y - tBar.height, Reference.asIs)
                        tt.horizontalAlignment = horizontalAlignment.center
                        tt.verticalAlignment = verticalAlignment.bottom
                        tt.text = String.Format("{0}", i * (wux / nDiv))
                        tt.fill = p_fill

                        tt.draw(contextobject, contextCoordinatesDelegate, contextSizeDelegate)

                        With rect1
                            .Width = oneTickX
                            .Height = tBar.height

                            .Stroke = Me.pen.stroke
                            .StrokeThickness = t.width
                            .StrokeDashArray = Me.pen.dashArray
                            If i Mod 2 = 0 Then
                                .Fill = Nothing
                            Else
                                .Fill = p_fill.Brush()
                                .Fill.Opacity = p_fill.Brush.Opacity
                            End If

                            myCanvas.Children.Add(rect1)
                            Windows.Controls.Canvas.SetLeft(rect1, p1.x + (i - 1) * oneTickX)
                            Windows.Controls.Canvas.SetTop(rect1, p1.y)

                        End With

                    Next

                End With
            Else
                Throw New NotImplementedException()
            End If
        End Sub

        Public Function boundingRectangle(Optional getWordSize As Drawable.estimateWorldSize = Nothing) As Double() Implements Drawable.boundingRectangle
            Dim r(3) As Double

            'Dim p1 As New Point

            If p_position.coordinateReference = Reference.world Then
                r(0) = p_position.x
                r(1) = p_position.y
                r(2) = p_position.x
                r(3) = p_position.y
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

        Public Function CompareTo(other As Drawable) As Integer Implements IComparable(Of Drawable).CompareTo
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

#Region "WoodGrainSymbol Class"

    ''' <summary>
    ''' Line as a drawable item
    ''' </summary>
    <ComVisible(False)>
    Public Class WoodGrainSymbol
        Implements Drawable

        Private p_angle As Double
        Private p_coordinateReference As Reference
        Private p_pen As New pen
        Private p_position As Point
        Private p_size As New size
        Private p_zIndex As Long

        Sub New()
            p_position = New Point(0, 0)
            p_size = New size(8, 1.5, Reference.contextMillimeters)
            p_pen = New pen()
            p_pen.color = Colors.Black
            p_pen.thickness = 0.12
            p_pen.thicknessReference = Reference.contextMillimeters
        End Sub

        Sub New(x As Double, y As Double, angleDeg As Double)
            Me.New()
            p_position = New Point(x, y)
            p_angle = angleDeg
        End Sub

        ''' <summary>
        ''' rotation angle in degrees
        ''' </summary>
        ''' <returns></returns>
        Public Property angle() As Double
            Get
                Return p_angle
            End Get
            Set(value As Double)
                p_angle = value
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

        ''' <summary>
        ''' Position
        ''' </summary>
        ''' <returns>
        ''' Point
        ''' </returns>
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
        Public Property zIndex As Long Implements Drawable.zIndex
            Get
                Return p_zIndex
            End Get
            Set(value As Long)
                p_zIndex = value
            End Set
        End Property

        Public Function boundingRectangle(Optional getWordSize As Drawable.estimateWorldSize = Nothing) As Double() Implements Drawable.boundingRectangle
            Return Me.position.boundingRectangle()
        End Function

        Public Function CompareTo(other As Drawable) As Integer Implements IComparable(Of Drawable).CompareTo
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

        Public Sub draw(contextobject As ContextObject, contextCoordinatesDelegate As Drawable.contextCoordinates, Optional contextSizeDelegate As Drawable.contextSize = Nothing) Implements Drawable.draw
            If TypeOf contextobject.Item Is System.Windows.Controls.Canvas Then
                Dim myCanvas As New System.Windows.Controls.Canvas
                myCanvas = TryCast(contextobject.Item, System.Windows.Controls.Canvas)
                With myCanvas
                    Dim myPolyline As New System.Windows.Shapes.Polyline

                    Dim p1 As New Point
                    Dim s1 As New size

                    Dim h As Double
                    Dim l As Double
                    Dim t As size

                    With myPolyline

                        s1 = contextSizeDelegate(Me.size)
                        p1 = contextCoordinatesDelegate(Me.position)

                        t = contextSizeDelegate(Me.pen.size)

                        h = s1.height / 2
                        l = s1.width / 2

                        .Points.Add(New System.Windows.Point(p1.x - l / 2 + 2.5 * h / 2, p1.y + h / 2))
                        .Points.Add(New System.Windows.Point(p1.x - l / 2, p1.y))
                        .Points.Add(New System.Windows.Point(p1.x + l / 2, p1.y))
                        .Points.Add(New System.Windows.Point(p1.x + l / 2 - 2.5 * h / 2, p1.y - h / 2))

                        .Stroke = Me.pen.stroke
                        .StrokeThickness = t.width
                        .StrokeDashArray = Me.pen.dashArray
                    End With

                    ' orient
                    Dim myRotateTransform As New Windows.Media.RotateTransform

                    myRotateTransform.CenterX = p1.x
                    myRotateTransform.CenterY = p1.y
                    myRotateTransform.Angle = -Me.angle
                    myPolyline.RenderTransform = myRotateTransform

                    myCanvas.Children.Add(myPolyline)

                End With
            Else
                Throw New NotImplementedException()
            End If
        End Sub
    End Class

#End Region

#Region "DimensionLine Class"

    ''' <summary>
    ''' Line as a drawable item
    ''' </summary>
    <ComVisible(False)>
    Public Class DimensionLine
        Implements Drawable

        ''' <summary>
        ''' Alignment
        ''' </summary>
        Private p_Alignment As DimAlignement

        Private p_anchorPointDistance As size

        Private p_dimSymbol As DimSymbols

        ''' <summary>
        ''' End Point
        ''' </summary>
        Private p_endPoint As Point

        ''' <summary>
        ''' offset
        ''' </summary>
        '''
        Private p_offset As size

        Private p_overlength As size

        Private p_pen As pen

        ''' <summary>
        ''' Start Point
        ''' </summary>
        Private p_startPoint As Point

        Private p_symbolSize As size

        Private p_textFormatString As String

        Private p_textOverwrite As String

        Private p_textPrefix As String
        Private p_textSize As size

        Private p_textSuffix As String
        Private p_zIndex As Long
        Sub New()
            Me.pen = New pen()
            Me.pen.thickness = 0.1

            Me.offset = New size()
            With Me.offset
                .width = 10
                .Reference = Reference.contextMillimeters
            End With

            Me.p_symbolSize = New size()
            With Me.p_symbolSize
                .width = 2.5
                .Reference = Reference.contextMillimeters
            End With

            Me.p_anchorPointDistance = New size()
            With Me.p_anchorPointDistance
                .width = 3
                .Reference = Reference.contextMillimeters
            End With

            Me.p_overlength = New size()
            With Me.p_overlength
                .width = 3
                .Reference = Reference.contextMillimeters
            End With

            Me.p_textSize = New size()
            With Me.p_textSize
                .width = 3
                .Reference = Reference.contextMillimeters
            End With

            p_dimSymbol = DimSymbols.Dash

            p_textFormatString = "0.00"

            Me.startPoint = New Point
            Me.endPoint = New Point
        End Sub

        ''' <summary>
        ''' DimensionLine from start and end points in world coordinates
        ''' </summary>
        ''' <param name="startPoint"></param>
        ''' <param name="endPoint"></param>
        Sub New(startPoint As Point, endPoint As Point)
            Me.New
            Me.startPoint = startPoint
            Me.endPoint = endPoint
        End Sub

        ''' <summary>
        ''' DimensionLine from start and end coordinates in world coordinates
        ''' </summary>
        ''' <param name="startX"></param>
        ''' <param name="startY"></param>
        ''' <param name="endX"></param>
        ''' <param name="endY"></param>
        Sub New(startX As Double, startY As Double, endX As Double, endY As Double)
            Me.New
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
            Me.New
            Me.startPoint = New Point(startX, startY, coordinateReference)
            Me.endPoint = New Point(endX, endY, coordinateReference)
        End Sub

        Public Enum DimAlignement
            aligned
            horizontal
            vertical
        End Enum

        Public Enum DimSymbols
            Dash
            Arrow
        End Enum

        ''' <summary>
        ''' Alignment of dimension line
        ''' </summary>
        ''' <returns></returns>
        Public Property alignment As DimAlignement
            Get
                Return p_Alignment
            End Get
            Set(value As DimAlignement)
                p_Alignment = value
            End Set
        End Property

        ''' <summary>
        ''' Anchor Point Distance
        '''
        '''            |
        '''            |
        ''' ·  --------+--
        '''            |
        '''   ^ this distance
        ''' </summary>
        ''' <returns></returns>
        Public Property anchorPointDistance As size
            Get
                Return p_anchorPointDistance
            End Get
            Set(value As size)
                p_anchorPointDistance = value
            End Set
        End Property

        ''' <summary>
        ''' Dimension Line Symbol
        ''' </summary>
        ''' <returns></returns>
        Public Property dimSymbol As DimSymbols
            Get
                Return p_dimSymbol
            End Get
            Set(value As DimSymbols)
                p_dimSymbol = value
            End Set
        End Property

        ''' <summary>
        ''' measured distance
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property distance() As Double
            Get
                Dim v1 As New Geom.Vector
                Dim v2 As New Geom.Vector
                Dim d As Double

                Select Case alignment
                    Case DimAlignement.aligned
                        v1.x = startPoint.x
                        v1.y = startPoint.y

                        v2.x = endPoint.x
                        v2.y = endPoint.y
                    Case DimAlignement.horizontal
                        v1.x = startPoint.x
                        v1.y = startPoint.y

                        v2.x = endPoint.x
                        v2.y = startPoint.y
                    Case DimAlignement.vertical
                        v1.x = startPoint.x
                        v1.y = startPoint.y

                        v2.x = startPoint.x
                        v2.y = endPoint.y
                End Select

                d = v1.distance(v2)
                Return d
            End Get
        End Property

        ''' <summary>
        ''' End Point
        ''' </summary>
        ''' <returns></returns>
        Public Property endPoint As Point
            Get
                Return p_endPoint
            End Get
            Set(value As Point)
                p_endPoint = value
            End Set
        End Property

        ''' <summary>
        ''' Offset value
        ''' </summary>
        ''' <returns></returns>
        Public Property offset As size
            Get
                Return p_offset
            End Get
            Set(value As size)
                p_offset = value
            End Set
        End Property

        ''' <summary>
        ''' Dimension Lines overlength
        '''
        '''            |
        '''            |
        ''' ·  --------+--
        '''            | ^ this overlengths
        '''
        ''' </summary>
        ''' <returns></returns>
        Public Property overlength As size
            Get
                Return p_overlength
            End Get
            Set(value As size)
                p_overlength = value
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

        ''' <summary>
        ''' Start Point
        ''' </summary>
        ''' <returns></returns>
        Public Property startPoint As Point
            Get
                Return p_startPoint
            End Get
            Set(value As Point)
                p_startPoint = value
            End Set
        End Property

        ''' <summary>
        ''' Symbol size
        ''' ·  --------+
        '''            ^ this size
        ''' </summary>
        ''' <returns></returns>
        Public Property symbolSize As size
            Get
                Return p_symbolSize
            End Get
            Set(value As size)
                p_symbolSize = value
            End Set
        End Property

        Public Property textFormatString() As String
            Get
                Return p_textFormatString
            End Get
            Set(value As String)
                p_textFormatString = value
            End Set
        End Property

        ''' <summary>
        ''' set a Text to overwrite measured value (default is nothing)
        ''' </summary>
        ''' <returns></returns>
        Public Property textOverwrite() As String
            Get
                Return p_textOverwrite
            End Get
            Set(value As String)
                p_textOverwrite = value
            End Set
        End Property

        ''' <summary>
        '''  Text Prefix
        ''' </summary>
        ''' <returns></returns>
        Public Property textPrefix() As String
            Get
                Return p_textPrefix
            End Get
            Set(value As String)
                p_textPrefix = value
            End Set
        End Property

        ''' <summary>
        ''' Text Size (default is 6.5 mm)
        ''' </summary>
        ''' <returns></returns>
        Public Property textSize As size
            Get
                Return p_textSize
            End Get
            Set(value As size)
                p_textSize = value
            End Set
        End Property

        ''' <summary>
        '''  Text Suffix
        ''' </summary>
        ''' <returns></returns>
        Public Property textSuffix() As String
            Get
                Return p_textSuffix
            End Get
            Set(value As String)
                p_textSuffix = value
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
        ''' Calculate bounding Rectangle
        ''' </summary>
        ''' <param name="contextCoordinatesDelegate"></param>
        ''' <param name="contextSizeDelegate"></param>
        ''' <returns></returns>
        Public Function boundingRectangle(Optional getWordSize As Drawable.estimateWorldSize = Nothing) As Double() Implements Drawable.boundingRectangle

            Dim r(3) As Double

            r(0) = startPoint.x
            r(1) = startPoint.y

            r(2) = endPoint.x
            r(3) = endPoint.y

            ' sort points
            For i As Integer = 0 To 1
                If r(0 + i) > r(2 + i) Then
                    Dim t As Double = r(0 + i)
                    r(0 + i) = r(2 + i)
                    r(2 + i) = t
                End If
            Next

            ' todo: include offset value to calcultate bounding rectangle

            'getWordSize
            Dim offsetW As New size
            Dim overlengthW As New size

            offsetW = getWordSize(offset)
            overlengthW = getWordSize(overlength)

            Dim gl1 As New Geom.Line
            Dim gl2 As New Geom.Line

            Select Case alignment
                Case DimAlignement.aligned
                    gl1 = New Geom.Line(startPoint.x, startPoint.y, endPoint.x - startPoint.x, endPoint.y - startPoint.y)
                Case DimAlignement.horizontal
                    gl1 = New Geom.Line(startPoint.x, startPoint.y, endPoint.x - startPoint.x, 0)
                Case DimAlignement.vertical
                    If startPoint.x >= endPoint.x Then
                        gl1 = New Geom.Line(startPoint.x, startPoint.y, 0, endPoint.y - startPoint.y)
                    Else
                        gl1 = New Geom.Line(startPoint.x, startPoint.y, 0, endPoint.y - startPoint.y)
                    End If
            End Select

            ' normal vector
            Dim uVect As Geom.Vector
            Dim nVect As Geom.Vector
            uVect = gl1.Direction.unitVector
            nVect = uVect.normal

            gl2 = gl1
            gl2 = gl1.OffsetLine(-offsetW.width)

            Dim p As New Point
            p.x = gl2.StartPointX - uVect.x * overlengthW.width
            p.y = gl2.StartPointY - uVect.y * overlengthW.width
            myMath.includeInBoundingRectangle(r, p)

            p.x = gl2.StartPointX + gl2.dx + uVect.x * overlengthW.width
            p.y = gl2.StartPointY + gl2.dy + uVect.y * overlengthW.width
            myMath.includeInBoundingRectangle(r, p)

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

        Public Function copy() As DimensionLine
            Dim newDL As DimensionLine = DirectCast(Me.MemberwiseClone(), DimensionLine)
            Return newDL
        End Function
        ''' <summary>
        ''' place on canvas
        ''' </summary>
        ''' <param name="contextobject"></param>
        ''' <param name="contextCoordinatesDelegate"></param>
        Public Overridable Sub draw(contextobject As ContextObject, contextCoordinatesDelegate As Drawable.contextCoordinates, Optional contextSizeDelegate As Drawable.contextSize = Nothing) Implements Drawable.draw
            If startPoint.x = endPoint.x And startPoint.y = endPoint.y Then Exit Sub

            If TypeOf contextobject.Item Is System.Windows.Controls.Canvas Then
                Dim myCanvas As New System.Windows.Controls.Canvas
                myCanvas = TryCast(contextobject.Item, System.Windows.Controls.Canvas)
                With myCanvas

                    Dim p1 As New Point
                    Dim p2 As New Point

                    Dim penSize As size
                    Dim textSize As size
                    Dim symSize As size
                    Dim ancSize As size
                    Dim overLengthSize As size
                    Dim offsetSize As size

                    '
                    ' calculate coordinates in local system
                    ' by a call-back to the calling drawing object
                    '
                    p1 = contextCoordinatesDelegate(startPoint)
                    p2 = contextCoordinatesDelegate(endPoint)

                    Dim v As New Geom.Vector

                    Select Case alignment
                        Case DimAlignement.aligned
                            v = New wmg.Vector(New Geom.Point(p1.x, p1.y), New Geom.Point(p2.x, p2.y))
                        Case DimAlignement.horizontal
                            v = New wmg.Vector(New Geom.Point(p1.x, 0), New Geom.Point(p2.x, 0))
                        Case DimAlignement.vertical
                            v = New wmg.Vector(New Geom.Point(0, p1.y), New Geom.Point(0, p2.y))
                    End Select
                    If v.length = 0 Then Exit Sub

                    ' calculate sizes
                    penSize = contextSizeDelegate(Me.pen.size)
                    textSize = contextSizeDelegate(Me.textSize)
                    symSize = contextSizeDelegate(Me.symbolSize)
                    ancSize = contextSizeDelegate(Me.anchorPointDistance)
                    overLengthSize = contextSizeDelegate(Me.overlength)
                    offsetSize = contextSizeDelegate(offset)

                    If penSize.width < 0.1 Then penSize.width = 0.1

                    Dim gl1 As New Geom.Line
                    Dim gl2 As New Geom.Line

                    Select Case alignment
                        Case DimAlignement.aligned
                            gl1 = New Geom.Line(p1.x, p1.y, p2.x - p1.x, p2.y - p1.y)
                        Case DimAlignement.horizontal
                            gl1 = New Geom.Line(p1.x, p1.y, p2.x - p1.x, 0)
                        Case DimAlignement.vertical
                            If p1.x >= p2.x Then
                                gl1 = New Geom.Line(p1.x, p1.y, 0, p2.y - p1.y)
                            Else
                                gl1 = New Geom.Line(p2.x, p1.y, 0, p2.y - p1.y)
                            End If
                    End Select

                    ' normal vector
                    Dim uVect As New Geom.Vector
                    Dim nVect As New Geom.Vector
                    uVect = gl1.Direction.unitVector
                    nVect = uVect.normal

                    gl2 = gl1
                    gl2 = gl1.OffsetLine(-offsetSize.width)

                    Dim mainLine As New System.Windows.Shapes.Line
                    With mainLine
                        .X1 = gl2.StartPointX - uVect.x * overLengthSize.width
                        .Y1 = gl2.StartPointY - uVect.y * overLengthSize.width
                        .X2 = gl2.StartPointX + gl2.dx + uVect.x * overLengthSize.width
                        .Y2 = gl2.StartPointY + gl2.dy + uVect.y * overLengthSize.width
                        .Stroke = Me.pen.stroke

                        .StrokeThickness = penSize.width
                        .StrokeDashArray = Me.pen.dashArray
                    End With
                    myCanvas.Children.Add(mainLine)

                    Dim sideLine1 As New System.Windows.Shapes.Line
                    With sideLine1
                        .X1 = gl2.StartPointX + nVect.x * overLengthSize.width
                        .Y1 = gl2.StartPointY + nVect.y * overLengthSize.width
                        .X2 = p1.x + nVect.x * ancSize.width
                        .Y2 = p1.y + nVect.y * ancSize.width

                        .Stroke = Me.pen.stroke
                        .StrokeThickness = penSize.width
                        .StrokeDashArray = Me.pen.dashArray
                    End With
                    myCanvas.Children.Add(sideLine1)

                    Dim sideLine2 As New System.Windows.Shapes.Line
                    With sideLine2
                        .X1 = gl2.StartPointX + gl2.dx + nVect.x * overLengthSize.width
                        .Y1 = gl2.StartPointY + gl2.dy + nVect.y * overLengthSize.width
                        .X2 = p2.x + nVect.x * ancSize.width
                        .Y2 = p2.y + nVect.y * ancSize.width

                        .Stroke = Me.pen.stroke
                        .StrokeThickness = penSize.width
                        .StrokeDashArray = Me.pen.dashArray
                    End With
                    myCanvas.Children.Add(sideLine2)

                    Dim symbolPolyline1 As New System.Windows.Shapes.Polyline
                    Dim symbolPolyline2 As New System.Windows.Shapes.Polyline

                    Select Case dimSymbol
                        Case DimSymbols.Dash
                            With symbolPolyline1
                                .Points.Add(New System.Windows.Point(0.5, -0.5))
                                .Points.Add(New System.Windows.Point(-0.5, 0.5))
                            End With

                            With symbolPolyline2
                                .Points.Add(New System.Windows.Point(0.5, -0.5))
                                .Points.Add(New System.Windows.Point(-0.5, 0.5))
                            End With

                        Case DimSymbols.Arrow
                            With symbolPolyline1
                                .Points.Add(New System.Windows.Point(0.75, 0.5))
                                .Points.Add(New System.Windows.Point(0, 0))
                                .Points.Add(New System.Windows.Point(0.75, -0.5))
                            End With

                            With symbolPolyline2
                                .Points.Add(New System.Windows.Point(-0.75, 0.5))
                                .Points.Add(New System.Windows.Point(0, 0))
                                .Points.Add(New System.Windows.Point(-0.75, -0.5))
                            End With
                        Case Else
                    End Select

                    Dim myRotateTransform As New Windows.Media.RotateTransform
                    Dim myScaleTransform As New Windows.Media.ScaleTransform
                    Dim myTranslateTransform As New Windows.Media.TranslateTransform
                    Dim myTranslateTransform2 As New Windows.Media.TranslateTransform

                    With myRotateTransform
                        .CenterX = 0
                        .CenterY = 0
                        .Angle = gl1.Direction.angle
                    End With

                    With myScaleTransform
                        .CenterX = 0
                        .CenterY = 0
                        .ScaleX = symSize.width
                        .ScaleY = symSize.width
                    End With

                    With myTranslateTransform
                        .X = gl2.StartPointX
                        .Y = gl2.StartPointY
                    End With

                    Dim myTransformGroup As New TransformGroup
                    myTransformGroup.Children.Add(myRotateTransform)
                    myTransformGroup.Children.Add(myScaleTransform)
                    myTransformGroup.Children.Add(myTranslateTransform)

                    With symbolPolyline1
                        .RenderTransform = myTransformGroup
                        .Stroke = Me.pen.stroke
                        .StrokeDashArray = Me.pen.dashArray
                        .StrokeThickness = penSize.width / symSize.width
                    End With

                    myCanvas.Children.Add(symbolPolyline1)

                    With myTranslateTransform2
                        .X = gl2.StartPointX + gl2.dx
                        .Y = gl2.StartPointY + gl2.dy
                    End With

                    Dim myTransformGroup2 As New TransformGroup
                    myTransformGroup2.Children.Add(myRotateTransform)
                    myTransformGroup2.Children.Add(myScaleTransform)
                    myTransformGroup2.Children.Add(myTranslateTransform2)

                    With symbolPolyline2
                        .RenderTransform = myTransformGroup2
                        .Stroke = Me.pen.stroke
                        .StrokeDashArray = Me.pen.dashArray
                        .StrokeThickness = penSize.width / symSize.width
                    End With
                    myCanvas.Children.Add(symbolPolyline2)

                    '
                    ' Dimension text
                    '
                    Dim tb As New Windows.Controls.TextBlock
                    Dim myBrush As New SolidColorBrush
                    myBrush.Color = Me.pen.color
                    myBrush.Opacity = Me.pen.opacity

                    If Me.textOverwrite = Nothing Then
                        '
                        ' dimension text
                        '
                        Dim dimText As String
                        'Dim v1 As New Geom.Vector
                        'Dim v2 As New Geom.Vector
                        Dim d As Double

                        'Select Case alignment
                        '    Case DimAlignement.aligned
                        '        v1.x = startPoint.x
                        '        v1.y = startPoint.y

                        '        v2.x = endPoint.x
                        '        v2.y = endPoint.y
                        '    Case DimAlignement.horizontal
                        '        v1.x = startPoint.x
                        '        v1.y = startPoint.y

                        '        v2.x = endPoint.x
                        '        v2.y = startPoint.y
                        '    Case DimAlignement.vertical
                        '        v1.x = startPoint.x
                        '        v1.y = startPoint.y

                        '        v2.x = startPoint.x
                        '        v2.y = endPoint.y
                        'End Select

                        'd = v1.distance(v2)

                        d = distance()

                        dimText = textPrefix & d.ToString(Me.textFormatString) & textSuffix

                        With tb
                            .Text = dimText
                            .FontSize = textSize.width
                            .Foreground = myBrush
                        End With
                    Else
                        With tb
                            .Text = Me.textOverwrite
                            .FontSize = textSize.width
                            .Foreground = myBrush
                        End With
                    End If

                    Dim myRotateTransform2 As New Windows.Media.RotateTransform
                    myRotateTransform2.Angle = gl1.Direction.angle * 180 / Math.PI

                    tb.RenderTransform = myRotateTransform2

                    myCanvas.Children.Add(tb)

                    Dim myFormattedText As New FormattedText(tb.Text, CultureInfo.CurrentCulture, Windows.FlowDirection.LeftToRight,
                        New Typeface(tb.FontFamily, tb.FontStyle, tb.FontWeight, tb.FontStretch), tb.FontSize, Brushes.Black,
                        New NumberSubstitution(), 1)

                    tb.HorizontalAlignment = Windows.HorizontalAlignment.Center
                    tb.VerticalAlignment = Windows.VerticalAlignment.Center

                    myCanvas.SetLeft(tb, gl2.StartPointX + gl2.dx / 2 - myFormattedText.Width / 2 * Math.Cos(gl1.Direction.angle) + myFormattedText.Height * Math.Sin(gl1.Direction.angle))
                    myCanvas.SetTop(tb, gl2.StartPointY + gl2.dy / 2 - myFormattedText.Width / 2 * Math.Sin(gl1.Direction.angle) - myFormattedText.Height * Math.Cos(gl1.Direction.angle))

                End With
            Else
                Throw New NotImplementedException()
            End If
        End Sub

    End Class

#End Region

#Region "DimensionAngular Class"

    ''' <summary>
    ''' Line as a drawable item
    ''' </summary>
    <ComVisible(False)>
    Public Class DimensionAngular
        Implements Drawable

        ''' <summary>
        ''' Alignment
        ''' </summary>
        Private p_Alignment As DimAlignement

        Private p_anchorPointDistance As size

        Private p_dimSymbol As DimSymbols

        ''' <summary>
        ''' line 1
        ''' </summary>
        Private p_line1 As Line

        ''' <summary>
        ''' line 2
        ''' </summary>
        Private p_line2 As Line

        ''' <summary>
        ''' offset
        ''' </summary>
        '''
        Private p_offset As size

        Private p_overlength As size

        Private p_pen As pen

        Private p_symbolSize As size
        Private p_textFormatString As String
        Private p_textOverwrite As String
        Private p_textPrefix As String
        Private p_textSize As size

        Private p_textSuffix As String
        Private p_zIndex As Long

        Sub New()
            Me.pen = New pen()
            Me.pen.thickness = 0.1

            Me.offset = New size()
            With Me.offset
                .width = 10
                .Reference = Reference.contextMillimeters
            End With

            Me.p_symbolSize = New size()
            With Me.p_symbolSize
                .width = 2.5
                .Reference = Reference.contextMillimeters
            End With

            Me.p_anchorPointDistance = New size()
            With Me.p_anchorPointDistance
                .width = 3
                .Reference = Reference.contextMillimeters
            End With

            Me.p_overlength = New size()
            With Me.p_overlength
                .width = 3
                .Reference = Reference.contextMillimeters
            End With

            Me.p_textSize = New size()
            With Me.p_textSize
                .width = 3
                .Reference = Reference.contextMillimeters
            End With

            p_dimSymbol = DimSymbols.Dash

            p_textFormatString = "0°"

            Me.line1 = New Line
            Me.line2 = New Line
        End Sub

        ''' <summary>
        ''' DimensionAngle between line1 and line2
        ''' </summary>
        ''' <param name="line1"></param>
        ''' <param name="line2"></param>
        Sub New(line1 As Line, line2 As Line)
            Me.New
            Me.line1 = line1
            Me.line2 = line2
        End Sub

        ''' <summary>
        ''' DimensionAngle between line1 and line2
        ''' </summary>
        ''' <param name="line1Start"></param>
        ''' <param name="Line1End"></param>
        ''' <param name="line2Start"></param>
        ''' <param name="Line2End"></param>
        Sub New(line1Start As Point, Line1End As Point, line2Start As Point, Line2End As Point)
            Me.New
            Me.line1 = New Line(line1Start, Line1End)
            Me.line2 = New Line(line2Start, Line2End)
        End Sub

        ''' <summary>
        ''' DimensionAngle between line1 and line2
        ''' </summary>
        ''' <param name="line1StartX"></param>
        ''' <param name="line1StartY"></param>
        ''' <param name="line1EndX"></param>
        ''' <param name="line1EndY"></param>
        ''' <param name="line2StartX"></param>
        ''' <param name="line2StartY"></param>
        ''' <param name="line2EndX"></param>
        ''' <param name="line2EndY"></param>
        ''' <param name="coordinateReference"></param>
        Sub New(line1StartX As Double, line1StartY As Double,
                line1EndX As Double, line1EndY As Double,
                line2StartX As Double, line2StartY As Double,
                line2EndX As Double, line2EndY As Double,
                coordinateReference As Reference)
            Me.New
            Me.line1 = New Line(New Point(line1StartX, line1StartY), New Point(line1EndX, line1EndY))
            Me.line2 = New Line(New Point(line2StartX, line2StartY), New Point(line2EndX, line2EndY))
        End Sub

        Public Enum DimAlignement
            inside
            outside
        End Enum

        Public Enum DimSymbols
            Dash
            Arrow
        End Enum

        ''' <summary>
        ''' Alignment of dimension line
        ''' </summary>
        ''' <returns></returns>
        Public Property alignment As DimAlignement
            Get
                Return p_Alignment
            End Get
            Set(value As DimAlignement)
                p_Alignment = value
            End Set
        End Property

        ''' <summary>
        ''' Anchor Point Distance
        '''
        '''            |
        '''            |
        ''' ·  --------+--
        '''            |
        '''   ^ this distance
        ''' </summary>
        ''' <returns></returns>
        Public Property anchorPointDistance As size
            Get
                Return p_anchorPointDistance
            End Get
            Set(value As size)
                p_anchorPointDistance = value
            End Set
        End Property

        ''' <summary>
        ''' Dimension Line Symbol
        ''' </summary>
        ''' <returns></returns>
        Public Property dimSymbol As DimSymbols
            Get
                Return p_dimSymbol
            End Get
            Set(value As DimSymbols)
                p_dimSymbol = value
            End Set
        End Property
        ''' <summary>
        ''' vector of line 1
        ''' </summary>
        ''' <returns></returns>
        Private ReadOnly Property v1() As Geom.Vector
            Get
                v1 = New Geom.Vector(
                        New Geom.Point(line1.startPoint.x, line1.startPoint.y),
                        New Geom.Point(line1.endPoint.x, line1.endPoint.y))
            End Get
        End Property

        ''' <summary>
        ''' vector of line 1
        ''' </summary>
        ''' <returns></returns>
        Private ReadOnly Property v2() As Geom.Vector
            Get
                v2 = New Geom.Vector(
                        New Geom.Point(line2.startPoint.x, line2.startPoint.y),
                        New Geom.Point(line2.endPoint.x, line2.endPoint.y))
            End Get
        End Property

        ''' <summary>
        ''' measured angle
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property angle() As Double
            Get
                Dim d As Double
                d = Math.Abs(v1.angle(v2)) * 180 / Math.PI

                Return d
            End Get
        End Property

        ''' <summary>
        ''' First Line
        ''' </summary>
        ''' <returns></returns>
        Public Property line1 As Line
            Get
                Return p_line1
            End Get
            Set(value As Line)
                p_line1 = value
            End Set
        End Property

        ''' <summary>
        ''' Offset value
        ''' </summary>
        ''' <returns></returns>
        Public Property offset As size
            Get
                Return p_offset
            End Get
            Set(value As size)
                p_offset = value
            End Set
        End Property

        ''' <summary>
        ''' Dimension Lines overlength
        '''
        '''            |
        '''            |
        ''' ·  --------+--
        '''            | ^ this overlengths
        '''
        ''' </summary>
        ''' <returns></returns>
        Public Property overlength As size
            Get
                Return p_overlength
            End Get
            Set(value As size)
                p_overlength = value
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

        ''' <summary>
        ''' Second Line
        ''' </summary>
        ''' <returns></returns>
        Public Property line2 As Line
            Get
                Return p_line2
            End Get
            Set(value As Line)
                p_line2 = value
            End Set
        End Property


        ''' <summary>
        ''' Symbol size
        ''' ·  --------+
        '''            ^ this size
        ''' </summary>
        ''' <returns></returns>
        Public Property symbolSize As size
            Get
                Return p_symbolSize
            End Get
            Set(value As size)
                p_symbolSize = value
            End Set
        End Property

        Public Property textFormatString() As String
            Get
                Return p_textFormatString
            End Get
            Set(value As String)
                p_textFormatString = value
            End Set
        End Property

        ''' <summary>
        ''' set a Text to overwrite measured value (default is nothing)
        ''' </summary>
        ''' <returns></returns>
        Public Property textOverwrite() As String
            Get
                Return p_textOverwrite
            End Get
            Set(value As String)
                p_textOverwrite = value
            End Set
        End Property

        ''' <summary>
        '''  Text Prefix
        ''' </summary>
        ''' <returns></returns>
        Public Property textPrefix() As String
            Get
                Return p_textPrefix
            End Get
            Set(value As String)
                p_textPrefix = value
            End Set
        End Property

        ''' <summary>
        ''' Text Size (default is 6.5 mm)
        ''' </summary>
        ''' <returns></returns>
        Public Property textSize As size
            Get
                Return p_textSize
            End Get
            Set(value As size)
                p_textSize = value
            End Set
        End Property

        ''' <summary>
        '''  Text Suffix
        ''' </summary>
        ''' <returns></returns>
        Public Property textSuffix() As String
            Get
                Return p_textSuffix
            End Get
            Set(value As String)
                p_textSuffix = value
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
        ''' Calculate bounding Rectangle
        ''' </summary>
        ''' <param name="contextCoordinatesDelegate"></param>
        ''' <param name="contextSizeDelegate"></param>
        ''' <returns></returns>
        Public Function boundingRectangle(Optional getWordSize As Drawable.estimateWorldSize = Nothing) As Double() Implements Drawable.boundingRectangle

            Dim r(3) As Double

            r(0) = line1.startPoint.x
            r(1) = line1.startPoint.y

            r(2) = line1.endPoint.x
            r(3) = line1.endPoint.y

            ' sort points
            For i As Integer = 0 To 1
                If r(0 + i) > r(2 + i) Then
                    Dim t As Double = r(0 + i)
                    r(0 + i) = r(2 + i)
                    r(2 + i) = t
                End If
            Next

            myMath.includeInBoundingRectangle(r, line2.startPoint)
            myMath.includeInBoundingRectangle(r, line2.endPoint)

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

        Public Function copy() As DimensionAngular
            Dim newDA As DimensionAngular = DirectCast(Me.MemberwiseClone(), DimensionAngular)
            Return newDA
        End Function

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

                    Dim penSize As size
                    Dim textSize As size
                    Dim symSize As size
                    Dim ancSize As size
                    Dim overLengthSize As size
                    Dim offsetSize As size

                    '
                    ' calculate coordinates in local system
                    ' by a call-back to the calling drawing object
                    '

                    'p1 = contextCoordinatesDelegate(startPoint)
                    'p2 = contextCoordinatesDelegate(endPoint)

                    Dim g1 As New Geom.Line

                    g1.StartPoint = New Geom.Vector(line1.startPoint.x, line1.startPoint.y)
                    g1.Direction = Me.v1

                    Dim g2 As New Geom.Line

                    g2.StartPoint = New Geom.Vector(line2.startPoint.x, line2.startPoint.y)
                    g2.Direction = Me.v2

                    Dim M As New Geom.Point         ' Mid point
                    Dim ML As New Point             ' mid point in context coordinates

                    M = g1.intersect(g2)
                    ML = contextCoordinatesDelegate(New Point(M.x, M.y))

                    Dim r As Double
                    Dim rs As size      ' radius a size
                    Dim rsL As size     ' radius as context size

                    Dim vMP As New Geom.Vector(M, New Geom.Point(line1.startPoint.x, line1.startPoint.y))
                    ' todo: offset missing
                    r = vMP.length
                    rs = New size(r, Reference.world)
                    rsL = contextSizeDelegate(rs)

                    Dim a1 As Double
                    Dim a2 As Double
                    Dim a1d As Double
                    Dim a2d As Double

                    a1 = v1.angle
                    a2 = v2.angle

                    ' calculate sizes
                    penSize = contextSizeDelegate(Me.pen.size)
                    textSize = contextSizeDelegate(Me.textSize)
                    symSize = contextSizeDelegate(Me.symbolSize)
                    ancSize = contextSizeDelegate(Me.anchorPointDistance)
                    overLengthSize = contextSizeDelegate(Me.overlength)
                    offsetSize = contextSizeDelegate(offset)

                    If penSize.width < 0.1 Then penSize.width = 0.1

                    Dim polygon1 As New System.Windows.Shapes.Polyline

                    Dim PStart As New Point
                    Dim pEnd As New Point
                    Dim pMid As New Point

                    PStart.x = ML.x + (rsL.average + offsetSize.average) * Math.Cos(a1)
                    PStart.y = ML.y - (rsL.average + offsetSize.average) * Math.Sin(a1)

                    pEnd.x = ML.x + (rsL.average + offsetSize.average) * Math.Cos(a2)
                    pEnd.y = ML.y - (rsL.average + offsetSize.average) * Math.Sin(a2)

                    pMid.x = ML.x + (rsL.average + offsetSize.average) * Math.Cos((a1 + a2) / 2)
                    pMid.y = ML.y - (rsL.average + offsetSize.average) * Math.Sin((a1 + a2) / 2)

                    Dim direction As Double
                    direction = Math.Sign(a1 - a2)

                    With polygon1

                        Dim p1 As New Point
                        Dim p2 As New Point

                        Dim phi As Double
                        Dim segments As Integer = 10

                        segments = CInt(Math.Abs(a2 - a1) * 180 / 5) + 1
                        If segments < 10 Then
                            segments = 10
                        End If

                        ' overlengths start and end
                        If (rsL.average + offsetSize.average) <> 0 Then
                            a1d = a1 + direction * overLengthSize.average / (rsL.average + offsetSize.average)
                            a2d = a2 - direction * overLengthSize.average / (rsL.average + offsetSize.average)
                        Else
                            a1d = a1 - 3
                            a2d = a2 + 3
                        End If

                        ' render arc
                        For i = 0 To segments
                            phi = a1d + (a2d - a1d) * i / segments
                            p1.x = ML.x + (rsL.average + offsetSize.average) * Math.Cos(phi)
                            p1.y = ML.y - (rsL.average + offsetSize.average) * Math.Sin(phi)
                            'Debug.Print(String.Format("P{2} ({0:F3},{1:F3})", p1.x, p1.y, i))

                            'Debug.Print(String.Format("p2.{2} ({0:F3},{1:F3})", p2.x, p2.y, i))
                            .Points.Add(New Windows.Point(x:=p1.x, y:=p1.y))
                        Next

                        .Fill = Nothing
                        .Stroke = Me.pen.stroke
                        .StrokeThickness = penSize.thickness
                        .StrokeDashArray = Me.pen.dashArray
                    End With

                    myCanvas.Children.Add(polygon1)


                    Dim symbolPolyline1 As New System.Windows.Shapes.Polyline
                    Dim symbolPolyline2 As New System.Windows.Shapes.Polyline

                    Select Case dimSymbol
                        Case DimSymbols.Dash
                            With symbolPolyline1
                                .Points.Add(New System.Windows.Point(0.5, -0.5))
                                .Points.Add(New System.Windows.Point(-0.5, 0.5))
                            End With

                            With symbolPolyline2
                                .Points.Add(New System.Windows.Point(0.5, -0.5))
                                .Points.Add(New System.Windows.Point(-0.5, 0.5))
                            End With

                        Case DimSymbols.Arrow
                            With symbolPolyline1
                                .Points.Add(New System.Windows.Point(-0.75, 0.5))
                                .Points.Add(New System.Windows.Point(0, 0))
                                .Points.Add(New System.Windows.Point(-0.75, -0.5))
                            End With

                            With symbolPolyline2
                                .Points.Add(New System.Windows.Point(-0.75, 0.5))
                                .Points.Add(New System.Windows.Point(0, 0))
                                .Points.Add(New System.Windows.Point(-0.75, -0.5))
                            End With
                        Case Else
                    End Select

                    Dim myRotateTransform As New Windows.Media.RotateTransform
                    Dim myScaleTransform As New Windows.Media.ScaleTransform
                    Dim myTranslateTransform As New Windows.Media.TranslateTransform
                    Dim myRotateTransform2 As New Windows.Media.RotateTransform
                    Dim myTranslateTransform2 As New Windows.Media.TranslateTransform


                    If a1 < a2 Then

                        ' overlengths start and end
                        If (rsL.average + offsetSize.average) <> 0 Then
                            a1d = a1 - direction * symSize.width / (rsL.average + offsetSize.average) / 2
                            a2d = a2 + direction * symSize.width / (rsL.average + offsetSize.average) / 2
                        Else
                            a1d = a1 + 3
                            a2d = a2 - 3
                        End If

                        With myRotateTransform
                            .CenterX = 0
                            .CenterY = 0
                            .Angle = 1 * (-a1d * 180 / Math.PI - 90) + 180
                        End With

                        With myRotateTransform2
                            .CenterX = 0
                            .CenterY = 0
                            .Angle = 1 * (-a2d * 180 / Math.PI - 90)
                        End With
                    Else

                        ' overlengths start and end
                        If (rsL.average + offsetSize.average) <> 0 Then
                            a1d = a1 - direction * symSize.width / (rsL.average + offsetSize.average) / 2
                            a2d = a2 + direction * symSize.width / (rsL.average + offsetSize.average) / 2
                        Else
                            a1d = a1 + 3
                            a2d = a2 - 3
                        End If

                        With myRotateTransform
                            .CenterX = 0
                            .CenterY = 0
                            .Angle = 1 * (-a1d * 180 / Math.PI + 90) - 180
                        End With

                        With myRotateTransform2
                            .CenterX = 0
                            .CenterY = 0
                            .Angle = 1 * (-a2d * 180 / Math.PI + 90)
                        End With
                    End If

                    With myScaleTransform
                        .CenterX = 0
                        .CenterY = 0
                        .ScaleX = symSize.width
                        .ScaleY = symSize.width
                    End With

                    With myTranslateTransform
                        .X = PStart.x
                        .Y = PStart.y
                    End With

                    Dim myTransformGroup As New TransformGroup
                    myTransformGroup.Children.Add(myRotateTransform)
                    myTransformGroup.Children.Add(myScaleTransform)
                    myTransformGroup.Children.Add(myTranslateTransform)

                    With symbolPolyline1
                        .RenderTransform = myTransformGroup
                        .Stroke = Me.pen.stroke
                        .StrokeDashArray = Me.pen.dashArray
                        .StrokeThickness = penSize.width / symSize.width
                    End With

                    myCanvas.Children.Add(symbolPolyline1)

                    With myTranslateTransform2
                        .X = pEnd.x
                        .Y = pEnd.y
                    End With

                    'Debug.Print("Alpha1={0}°; Alpha2={1}°", a1 * 180 / Math.PI, a2 * 180 / Math.PI)

                    Dim myTransformGroup2 As New TransformGroup
                    myTransformGroup2.Children.Add(myRotateTransform2)
                    myTransformGroup2.Children.Add(myScaleTransform)
                    myTransformGroup2.Children.Add(myTranslateTransform2)

                    With symbolPolyline2
                        .RenderTransform = myTransformGroup2
                        .Stroke = Me.pen.stroke
                        .StrokeDashArray = Me.pen.dashArray
                        .StrokeThickness = penSize.width / symSize.width
                    End With

                    myCanvas.Children.Add(symbolPolyline2)

                    '
                    ' Dimension text
                    '
                    Dim tb As New Windows.Controls.TextBlock
                    Dim myBrush As New SolidColorBrush
                    myBrush.Color = Me.pen.color
                    myBrush.Opacity = Me.pen.opacity

                    If Me.textOverwrite = Nothing Then
                        '
                        ' dimension text
                        '
                        Dim dimText As String
                        'Dim v1 As New Geom.Vector
                        'Dim v2 As New Geom.Vector
                        Dim d As Double

                        d = angle()

                        dimText = textPrefix & d.ToString(Me.textFormatString) & textSuffix

                        With tb
                            .Text = dimText
                            .FontSize = textSize.width
                            .Foreground = myBrush
                        End With
                    Else
                        With tb
                            .Text = Me.textOverwrite
                            .FontSize = textSize.width
                            .Foreground = myBrush
                        End With
                    End If

                    Dim myRotateTransform3 As New Windows.Media.RotateTransform
                    myRotateTransform3.Angle = 90 - (a1 + a2) / 2 * 180 / Math.PI

                    tb.RenderTransform = myRotateTransform3

                    myCanvas.Children.Add(tb)

                    Dim myFormattedText As New FormattedText(tb.Text, CultureInfo.CurrentCulture, Windows.FlowDirection.LeftToRight,
                        New Typeface(tb.FontFamily, tb.FontStyle, tb.FontWeight, tb.FontStretch), tb.FontSize, Brushes.Black,
                        New NumberSubstitution(), 1)

                    tb.HorizontalAlignment = Windows.HorizontalAlignment.Center
                    tb.VerticalAlignment = Windows.VerticalAlignment.Center

                    myCanvas.SetLeft(tb, pMid.x - myFormattedText.Width / 2 * Math.Cos(myRotateTransform3.Angle * Math.PI / 180) + myFormattedText.Height * Math.Sin(myRotateTransform3.Angle * Math.PI / 180))
                    myCanvas.SetTop(tb, pMid.y - myFormattedText.Width / 2 * Math.Sin(myRotateTransform3.Angle * Math.PI / 180) - myFormattedText.Height * Math.Cos(myRotateTransform3.Angle * Math.PI / 180))

                End With
                    Else
                Throw New NotImplementedException()
            End If
        End Sub

    End Class

#End Region
End Namespace