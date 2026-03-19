Imports System.Runtime.InteropServices

Namespace WMDraw

#Region "CLTSection Class"

    ''' <summary>
    ''' Cross-Laminated Timber (CLT) cross-section drawable.
    ''' Renders a CLT element as a stack of laminations with optional dimensions and hatching.
    ''' 0° layers (parallel to span) are shown as individual boards with random widths/heights.
    ''' 90° layers (perpendicular) are shown as a single filled rectangle.
    ''' Multiple sections can be positioned via x0/y0 (bottom-left corner in world mm).
    ''' </summary>
    <ComVisible(False)>
    Public Class CLTSection
        Implements Drawable

        Private p_layers As Double()
        Private p_orientations As Integer()
        Private p_pen As New pen
        Private p_zIndex As Long

        ''' <summary>Thicknesses of each layer in mm.</summary>
        Public ReadOnly Property layers As Double()
            Get
                Return p_layers
            End Get
        End Property

        ''' <summary>Grain orientation per layer: 0 = parallel to span, 90 = perpendicular.</summary>
        Public ReadOnly Property orientations As Integer()
            Get
                Return p_orientations
            End Get
        End Property

        ''' <summary>True = vertical section (layers stacked downward); False = horizontal (layers left to right).</summary>
        Public Property isVertical As Boolean = True

        ''' <summary>Width of section perpendicular to stacking direction, in world mm (default 1000).</summary>
        Public Property displayWidth As Double = 1000

        ''' <summary>0 = no labels/dimensions; 1 = thickness label per layer; 2 = full dimension lines (default).</summary>
        Public Property showDimensions As Integer = 2

        ''' <summary>True = draw as defined; False = swap 0° and 90° display (cross-section view).</summary>
        Public Property drawCrossSection As Boolean = True

        ''' <summary>X coordinate of bottom-left corner in world mm (default 0).</summary>
        Public Property x0 As Double = 0

        ''' <summary>Y coordinate of bottom-left corner in world mm (default 0).</summary>
        Public Property y0 As Double = 0

        ''' <summary>False = solid colour fills (default); True = line hatching for 0° layers, transparent for 90°.</summary>
        Public Property useHatching As Boolean = False

        ''' <summary>
        ''' Creates a CLTSection with typed layer arrays.
        ''' </summary>
        ''' <param name="layers">Layer thicknesses in mm.</param>
        ''' <param name="orientations">Layer orientations: 0 = parallel, 90 = perpendicular.</param>
        Public Sub New(layers As Double(), orientations As Integer())
            p_layers = layers
            p_orientations = orientations
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

        Public Function CompareTo(other As Drawable) As Integer Implements IComparable(Of Drawable).CompareTo
            Return p_zIndex.CompareTo(other.zIndex)
        End Function

        Public Function boundingRectangle(Optional getWordSize As Drawable.estimateWorldSize = Nothing) As Double() Implements Drawable.boundingRectangle
            Dim totalThickness As Double = p_layers.Sum()
            Dim r(3) As Double
            r(0) = x0
            r(1) = y0
            r(2) = If(isVertical, x0 + displayWidth, x0 + totalThickness)
            r(3) = If(isVertical, y0 + totalThickness, y0 + displayWidth)
            Return r
        End Function

        Public Sub draw(contextobject As ContextObject, contextCoordinatesDelegate As Drawable.contextCoordinates, Optional contextSizeDelegate As Drawable.contextSize = Nothing) Implements Drawable.draw
            For Each item In buildDrawables()
                item.draw(contextobject, contextCoordinatesDelegate, contextSizeDelegate)
            Next
        End Sub

        ''' <summary>
        ''' Builds the list of sub-drawables that make up the CLT section.
        ''' </summary>
        Private Function buildDrawables() As List(Of Drawable)
            Dim items As New List(Of Drawable)

            ' --- Pens ---
            Dim penLayer As New pen             ' layer borders (0° boards + 90° rectangles)
            With penLayer
                .color = WMColors.DarkGrey
                .thickness = 0.15
                .thicknessReference = Reference.contextMillimeters
            End With

            Dim penBoard As New pen             ' individual board joints within 0° layers
            With penBoard
                .color = WMColors.DarkGrey
                .thickness = 0.15
                .thicknessReference = Reference.contextMillimeters
            End With

            Dim penCircumference As New pen     ' outer boundary of the full element
            With penCircumference
                .color = WMColors.Black
                .thickness = 0.3
                .thicknessReference = Reference.contextMillimeters
            End With

            ' --- Fills ---
            Dim fill0 As New fill               ' 0° parallel layers
            Dim fill90 As New fill              ' 90° perpendicular layers

            If useHatching Then
                ' 0° layers: line hatching; 90° layers: transparent
                Dim hatchAngle As Double = If(isVertical, 45, 135)
                fill90.color = WMColors.Transparent
                With fill0
                    .type = fill.fillType.linearHatching
                    .setLinearHatch(WMColors.DarkGrey, 0.7, 1, 3, hatchAngle)
                End With
            Else
                fill0.color = WMColors.WoodMediumYellow
                fill90.color = WMColors.WoodLightYellow
            End If

            Dim fontSize As New size(2, Reference.contextMillimeters)
            Dim n As Integer = p_layers.Length
            Dim totalThickness As Double = p_layers.Sum()

            ' xOff = x of left edge; yOff = y of top edge (WMDraw y decreases downward)
            Dim xOff As Double = x0
            Dim yOff As Double = If(isVertical, y0 + totalThickness, y0 + displayWidth)

            Dim pos As Double = 0

            Randomize()

            For i As Integer = 0 To n - 1
                Dim d_i As Double = p_layers(i)
                Dim o_i As Integer = If(Not drawCrossSection,
                                        If(p_orientations(i) = 0, 90, 0),
                                        p_orientations(i))

                If isVertical Then
                    ' Vertical section: layers stacked downward
                    If o_i = 0 Then
                        ' 0° layer: individual boards as vertical strips with random widths
                        Dim cx As Double = 0
                        Do While cx < displayWidth
                            Dim dx As Double = Int((125 - 100 + 1) * Rnd() + 100)
                            If displayWidth - (cx + dx) < 30 Then dx = displayWidth - cx
                            Dim board As New Rectangle(xOff + cx, yOff - pos, xOff + cx + dx, yOff - (pos + d_i))
                            board.pen = penBoard
                            board.fill = fill0
                            items.Add(board)
                            cx += dx
                            If cx >= displayWidth Then Exit Do
                        Loop
                    Else
                        ' 90° layer: single filled rectangle
                        Dim r As New Rectangle(xOff, yOff - pos, xOff + displayWidth, yOff - (pos + d_i))
                        r.pen = penLayer
                        r.fill = fill90
                        items.Add(r)
                    End If

                    If showDimensions = 1 AndAlso d_i >= 8 Then
                        Dim lbl As New Text(xOff + displayWidth * 0.02, yOff - (pos + d_i / 2),
                                            CInt(d_i).ToString() & If(o_i = 0, "x", "y"), 1.8)
                        lbl.horizontalAlignment = horizontalAlignment.left
                        lbl.verticalAlignment = verticalAlignment.center
                        items.Add(lbl)
                    ElseIf showDimensions = 2 Then
                        If d_i >= 8 Then
                            Dim lbl As New Text(xOff + displayWidth * 0.02, yOff - (pos + d_i / 2),
                                                o_i.ToString() & ChrW(176), 1.8)
                            lbl.horizontalAlignment = horizontalAlignment.left
                            lbl.verticalAlignment = verticalAlignment.center
                            items.Add(lbl)
                        End If
                        Dim dl As New DimensionLine()
                        dl.startPoint = New Point(xOff + displayWidth, yOff - pos)
                        dl.endPoint = New Point(xOff + displayWidth, yOff - (pos + d_i))
                        dl.offset = New size(8, Reference.contextMillimeters)
                        dl.textFormatString = "0.0"
                        dl.textSize = fontSize
                        items.Add(dl)
                    End If

                Else
                    ' Horizontal layout: layers stacked rightward
                    If o_i = 0 Then
                        ' 0° layer: individual boards as horizontal strips with random heights
                        Dim cy As Double = 0
                        Do While cy < displayWidth
                            Dim dy As Double = Int((125 - 100 + 1) * Rnd() + 100)
                            If displayWidth - (cy + dy) < 30 Then dy = displayWidth - cy
                            Dim board As New Rectangle(xOff + pos, yOff - cy, xOff + pos + d_i, yOff - (cy + dy))
                            board.pen = penBoard
                            board.fill = fill0
                            items.Add(board)
                            cy += dy
                            If cy >= displayWidth Then Exit Do
                        Loop
                    Else
                        ' 90° layer: single filled rectangle
                        Dim r As New Rectangle(xOff + pos, yOff, xOff + pos + d_i, y0)
                        r.pen = penLayer
                        r.fill = fill90
                        items.Add(r)
                    End If

                    If showDimensions = 1 AndAlso d_i >= 8 Then
                        Dim lbl As New Text(xOff + pos + d_i / 2, yOff - displayWidth / 2,
                                            CInt(d_i).ToString() & If(o_i = 0, "x", "y"), 1.8)
                        lbl.horizontalAlignment = horizontalAlignment.center
                        lbl.verticalAlignment = verticalAlignment.center
                        lbl.angle = 90
                        items.Add(lbl)
                    ElseIf showDimensions = 2 Then
                        If d_i >= 8 Then
                            Dim lbl As New Text(xOff + pos + d_i / 2, yOff - displayWidth / 2,
                                                o_i.ToString() & ChrW(176), 1.8)
                            lbl.horizontalAlignment = horizontalAlignment.center
                            lbl.verticalAlignment = verticalAlignment.center
                            lbl.angle = 90
                            items.Add(lbl)
                        End If
                        Dim dl As New DimensionLine()
                        dl.startPoint = New Point(xOff + pos, yOff)
                        dl.endPoint = New Point(xOff + pos + d_i, yOff)
                        dl.offset = New size(-8, Reference.contextMillimeters)
                        dl.textFormatString = "0.0"
                        dl.textSize = fontSize
                        items.Add(dl)
                    End If

                End If

                pos += d_i
            Next

            ' Circumference: outer boundary drawn on top of all layers
            Dim rcOuter As New Rectangle(xOff, yOff,
                                         If(isVertical, xOff + displayWidth, xOff + pos),
                                         y0)
            Dim fillNone As New fill()
            fillNone.color = WMColors.Transparent
            rcOuter.pen = penCircumference
            rcOuter.fill = fillNone
            rcOuter.zIndex = 10
            items.Add(rcOuter)

            ' Total dimension lines (mode 2 only)
            If showDimensions = 2 Then
                If isVertical Then
                    Dim dlTotal As New DimensionLine()
                    dlTotal.startPoint = New Point(xOff + displayWidth, yOff)
                    dlTotal.endPoint = New Point(xOff + displayWidth, y0)
                    dlTotal.offset = New size(16, Reference.contextMillimeters)
                    dlTotal.textFormatString = "t=0.0"
                    dlTotal.textSize = fontSize
                    items.Add(dlTotal)

                    Dim dlW As New DimensionLine()
                    dlW.startPoint = New Point(xOff, yOff)
                    dlW.endPoint = New Point(xOff + displayWidth, yOff)
                    dlW.offset = New size(-8, Reference.contextMillimeters)
                    dlW.textFormatString = "0 mm"
                    dlW.textSize = fontSize
                    items.Add(dlW)
                Else
                    Dim dlTotal As New DimensionLine()
                    dlTotal.startPoint = New Point(xOff, y0)
                    dlTotal.endPoint = New Point(xOff + pos, y0)
                    dlTotal.offset = New size(16, Reference.contextMillimeters)
                    dlTotal.textFormatString = "t=0.0"
                    dlTotal.textSize = fontSize
                    items.Add(dlTotal)

                    Dim dlH As New DimensionLine()
                    dlH.startPoint = New Point(xOff, yOff)
                    dlH.endPoint = New Point(xOff, y0)
                    dlH.offset = New size(-8, Reference.contextMillimeters)
                    dlH.textFormatString = "0 mm"
                    dlH.textSize = fontSize
                    items.Add(dlH)
                End If
            End If

            Return items
        End Function

    End Class

#End Region

End Namespace
