Imports WallnerMild
Imports WallnerMild.WMDraw
Imports WallnerMild.c_WMComDraw

Class MainWindow
    Private Sub button_Click(sender As Object, e As RoutedEventArgs) Handles button.Click

        Dim wmd As New Drawing

        '
        ' add some lines
        '

        '
        ' coordinates as fraction
        '
        'wmd.add(New Line(0, 0, 1, 1, References.contextFraction))
        'wmd.add(New Line(0, 1, 1, 0, References.contextFraction))

        ' House in world coordinates
        'wmd.add(New Line(0, 0, 3, 0))
        'wmd.add(New Line(3, 0, 3, 3))
        'wmd.add(New Line(3, 3, 0, 3))
        'wmd.add(New Line(0, 3, 0, 0))

        'wmd.add(New Line(3, 3, 1.5, 4.5))
        'wmd.add(New Line(1.5, 4.5, 0, 3))

        ' World Coordinates
        '   wmd.add(New Line(-5, -2, 3, 1))
        '   wmd.add(New Line(3, 1, 4.5, 1.5))
        '   wmd.add(New Line(4.5, 1.5, 9, 3))


        ' moment
        '
        'wmd.add(New MomentArc(3, 1, "M = 1 kNm"))
        'wmd.add(New MomentArc(4.5, 1.5, "M = 2 kNm", 30, 150))

        ' force

        'Dim FR As ForceArrow

        'For i As Integer = 0 To 90 Step 30
        '    FR = New ForceArrow(1, 1, 20, i, "{0:F2} kN", 20)
        '    wmd.add(FR)
        'Next

        'For i As Integer = 0 To 90 Step 30
        '    FR = New ForceArrow(5, 1.5, 20, i, "{0:F2} kN", 20)
        '    FR.ForceDirection = ForceArrow.forceDirections.ArrowAtBottom
        '    wmd.add(FR)
        'Next


        'point
        wmd.add(New Point(3, 3, PointDisplay.x, 2, Reference.contextMillimeters))

        wmd.add(New Line(0, 0.25, 1, 0.25, Reference.contextFraction))
        wmd.add(New Point(0.5, 0.25, Reference.contextFraction, PointDisplay.x, 2, Reference.contextMillimeters))

        Dim myBeam As New Beam(1, 1, 8, 2)
        wmd.add(myBeam)

        Dim alpha As Double
        alpha = Math.Atan2(2 - 1, 8 - 1)

        Dim el1 As New Ellipse(New Point((1 + 8) / 2, (1 + 2) / 2), New size(Math.Sqrt((8 - 1) ^ 2 + (2 - 1) ^ 2) / 2, Math.Sqrt((8 - 1) ^ 2 + (2 - 1) ^ 2) / 4))
        el1.zIndex = 2
        wmd.add(el1)

        Dim el2 As New Ellipse(New Point((1 + 8) / 2, (1 + 2) / 2), New size(Math.Sqrt((8 - 1) ^ 2 + (2 - 1) ^ 2) / 2, Math.Sqrt((8 - 1) ^ 2 + (2 - 1) ^ 2) / 4))
        el2.angle = alpha * 180 / Math.PI
        el2.fill = New fill()
        el2.fill.setLinearHatch(Color.FromRgb(255, 0, 0))
        wmd.add(el2)

        Dim rc2 As New Rectangle(New Point(1, 1), New Point(3, 6))
        wmd.add(rc2)

        Dim fa As New ForceArrow()
        fa.tipPoint = myBeam.innerPoint(0.3)
        fa.F = 10
        fa.angle = -90
        wmd.add(fa)



        Dim udl As New BeamUniformLoad(myBeam, New size(5, Reference.contextMillimeters), BeamUniformLoad.loadReferenceLength.horizontalProjection, BeamUniformLoad.loadDirectionType.localKS)
        udl.pen.color = Colors.Red
        udl.loadIntensityEnd = New size(7, Reference.contextMillimeters)
        udl.drawArrows = True
        'udl.direction = BeamUniformLoad.loadDirectionType.localKS
        udl.direction = BeamUniformLoad.loadDirectionType.globalKS
        udl.lengthReference = BeamUniformLoad.loadReferenceLength.horizontalProjection
        'udl.lengthReference = BeamUniformLoad.loadReferenceLength.realLength
        wmd.add(udl)

        'support
        wmd.add(New Support(1, 1, 30, "A"))
        wmd.add(New Support(8, 2, 0, "B"))


        wmd.add(New DimensionLine(1, 1, 8, 2))

        Dim d2 As New DimensionLine(1, 1, 8, 2)
        d2.alignment = DimensionLine.DimAlignement.vertical
        wmd.add(d2)



        Dim d3 As New DimensionLine(1, 1, 8, 2)
        d3.alignment = DimensionLine.DimAlignement.horizontal
        wmd.add(d3)

        Dim p2 As New pen
        p2.color = System.Windows.Media.Colors.Green
        p2.dashString = pen.DASHARRAY_DASHDOT
        p2.thickness = 0.1
        p2.thicknessReference = Reference.world

        wmd.pen = p2
        wmd.add(New Line(0, 0.75, 1, 0.75, Reference.contextFraction))

        '
        ' set contextObject
        '
        wmd.ContextObject.Item = Me.canvas1

        '
        ' set contexttype (should be done automatically)
        '
        wmd.Context = Contexts.WPFCanvas

        wmd.ContextObject.Margin = New WMDraw.Margin(100)
        wmd.ContextObject.Margin.Reference = Reference.contextUnits

        wmd.ContextObject.fitHeight = True
        wmd.ContextObject.fitWidth = True
        wmd.ContextObject.fitProportional = True


        '
        ' return useable height
        '
        Debug.Print(wmd.ContextObject.heightMM)
        Debug.Print(wmd.ContextObject.widthMM)

        wmd.draw()

        wmd.Context = Contexts.PNGClipboard
        ' virtual canvas with 92 dpi
        Dim c As New Canvas
        c.Height = 500
        c.Width = 500
        c.Background = New System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.Transparent)

        wmd.ContextObject.Item = c
        wmd.draw()

        MsgBox(String.Format("Size im mm: {0} / {1} mm", wmd.ContextObject.widthMM, wmd.ContextObject.heightMM))

    End Sub

    Private Sub button1_Click(sender As Object, e As RoutedEventArgs) Handles button1.Click
        Dim l1 As New Geom.Line
        l1.StartPoint = New Geom.Vector(1, 1)
        l1.Direction = New Geom.Vector(1, 1)
        'l1.offsetLine(0.5)
        Dim l2 As New Geom.Line
        l2.StartPoint = New Geom.Vector(3, 1)
        l2.Direction = New Geom.Vector(0, 1)
        'l2.offsetLine(0.5)

        Dim S As New Geom.Point
        S = l1.intersect(l2)

        MsgBox(String.Format("S ({0}/{1})", S.x, S.y))
    End Sub

    Private Sub button2_Click(sender As Object, e As RoutedEventArgs) Handles button2.Click

        Dim wmcom As New c_WMComDraw
        wmcom.test1()
        MsgBox("Tested c_WMComDraw - COM Interop class." & vbCrLf & "PNG Copied to Clipboard")
    End Sub

    Private Sub button3_Click(sender As Object, e As RoutedEventArgs) Handles button3.Click
        Dim wmd As New Drawing

        '
        ' add some lines
        '

        ' World Coordinates
        wmd.add(New Line(0, 0, 3, 1, Reference.world))
        wmd.add(New Line(3, 1, 4.5, 1.5, Reference.world))
        wmd.add(New Line(4.5, 1.5, 9, 3, Reference.world))
        wmd.setContext(Contexts.PNGFile, 50, 50, "mm")
        wmd.ContextObject.fitProportional = True
        wmd.ContextObject.Margin = New WMDraw.Margin(10, 10, 10, 10)
        MsgBox(wmd.draw())
    End Sub

    Private Sub buttonB_Click(sender As Object, e As RoutedEventArgs) Handles buttonB.Click

        Dim wmd As New Drawing

        For i As Long = 0 To 9 Step 2
            wmd.add(New Line(i / 10, 0, (i + 1) / 10, 0, References.contextFraction))
            wmd.add(New Line(0, i / 10, 0, (i + 1) / 10, References.contextFraction))

            wmd.add(New Line(i / 10, 1, (i + 1) / 10, 1, References.contextFraction))
            wmd.add(New Line(1, i / 10, 1, (i + 1) / 10, References.contextFraction))
        Next

        'wmd.add(New Line(0, 0, 0, 1, References.contextFraction))
        'wmd.add(New Line(0, 0, 1, 0, References.contextFraction))

        wmd.add(New Line(0, 0, 1, 1, References.contextFraction))
        wmd.add(New Line(0, 1, 1, 0, References.contextFraction))


        wmd.add(New Line(0, 0.5, 1, 0.5, References.contextFraction))
        wmd.add(New Line(0.5, 0, 0.5, 1, References.contextFraction))

        wmd.setContext(Contexts.PNGFile, 50, 50, "mm", "d:\temp\file1.png")

        wmd.ContextObject.Margin = New WMDraw.Margin(3)
        wmd.ContextObject.Margin.Reference = Reference.contextMillimeters

        MsgBox(wmd.draw())
    End Sub

    Private Sub button4_Click(sender As Object, e As RoutedEventArgs) Handles button4.Click
        Dim wmd As New Drawing
        '
        ' set contextObject
        '
        wmd.ContextObject.Item = Me.canvas1

        '
        ' set contexttype (should be done automatically)
        '
        wmd.Context = Contexts.WPFCanvas

        wmd.ContextObject.Margin = New WMDraw.Margin(50)
        wmd.ContextObject.Margin.Reference = Reference.contextUnits

        wmd.ContextObject.fitHeight = True
        wmd.ContextObject.fitWidth = True
        wmd.ContextObject.fitProportional = True


        '
        ' return useable height
        '
        Debug.Print(wmd.ContextObject.heightMM)
        Debug.Print(wmd.ContextObject.widthMM)

        wmd.draw()
    End Sub

    Private Sub buttonCLT_Click(sender As Object, e As RoutedEventArgs) Handles buttonCLT.Click
        ' 5-layer CLT: 40/20/20/20/40 mm, alternating 0°/90°
        Dim d(4) As Double
        d(0) = 40 : d(1) = 20 : d(2) = 20 : d(3) = 20 : d(4) = 40

        Dim o(4) As Integer
        o(0) = 0 : o(1) = 90 : o(2) = 0 : o(3) = 90 : o(4) = 0

        Dim wmcom As New c_WMComDraw
        Dim gap As Double = 50          ' mm gap between rows/columns
        Dim w As Double = 400           ' display width per section
        Dim tTotal As Double = 140      ' total CLT thickness (40+20+20+20+40)
        Dim xStep As Double = w + gap   ' x spacing between sections in a row

        ' yBase tracks the bottom-left y of each row; decremented after each row
        Dim yBase As Double = 0

        ' --- Row 1: isVertical=True,  showDimensions=2 (full dim lines) ---
        wmcom.drawCLTSection(True, d, o, True, w, 2, True, 0, yBase, True)
        wmcom.drawCLTSection(True, d, o, True, w, 2, True, xStep, yBase, False)
        wmcom.drawCLTSection(True, d, o, True, w, 2, False, 2 * xStep, yBase, False)
        yBase -= tTotal + gap

        ' --- Row 2: isVertical=False, showDimensions=2 (full dim lines) ---
        wmcom.drawCLTSection(True, d, o, False, w, 2, True, 0, yBase, False)
        wmcom.drawCLTSection(True, d, o, False, w, 2, True, xStep, yBase, False)
        wmcom.drawCLTSection(True, d, o, False, w, 2, False, 2 * xStep, yBase, False)
        yBase -= w + gap

        ' --- Row 3: isVertical=True,  showDimensions=1 (compact x/y labels) ---
        wmcom.drawCLTSection(True, d, o, True, w, 1, True, 0, yBase, False)
        wmcom.drawCLTSection(True, d, o, True, w, 1, True, xStep, yBase, False)
        wmcom.drawCLTSection(True, d, o, True, w, 1, False, 2 * xStep, yBase, False)
        yBase -= tTotal + gap

        ' --- Row 4: isVertical=False, showDimensions=1 (compact x/y labels) ---
        wmcom.drawCLTSection(True, d, o, False, w, 1, True, 0, yBase, False)
        wmcom.drawCLTSection(True, d, o, False, w, 1, True, xStep, yBase, False)
        wmcom.drawCLTSection(True, d, o, False, w, 1, False, 2 * xStep, yBase, False)
        yBase -= w + gap

        ' --- Row 5: isVertical=True,  showDimensions=0 (no labels) ---
        wmcom.drawCLTSection(True, d, o, True, w, 0, True, 0, yBase, False)
        wmcom.drawCLTSection(True, d, o, True, w, 0, True, xStep, yBase, False)
        wmcom.drawCLTSection(True, d, o, True, w, 0, False, 2 * xStep, yBase, False)
        yBase -= tTotal + gap

        ' --- Row 6: isVertical=False, showDimensions=0 (no labels) ---
        wmcom.drawCLTSection(True, d, o, False, w, 0, True, 0, yBase, False)
        wmcom.drawCLTSection(True, d, o, False, w, 0, True, xStep, yBase, False)
        wmcom.drawCLTSection(True, d, o, False, w, 0, False, 2 * xStep, yBase, False)
        yBase -= w + gap

        ' --- Row 7: isVertical=True,  showDimensions=2, useHatching=True ---
        wmcom.drawCLTSection(True, d, o, True, w, 2, True, 0, yBase, False, True)
        wmcom.drawCLTSection(True, d, o, True, w, 2, True, xStep, yBase, False, True)
        wmcom.drawCLTSection(True, d, o, True, w, 2, False, 2 * xStep, yBase, False, True)
        yBase -= tTotal + gap

        ' --- Row 8: isVertical=False, showDimensions=2, useHatching=True ---
        wmcom.drawCLTSection(True, d, o, False, w, 2, True, 0, yBase, False, True)
        wmcom.drawCLTSection(True, d, o, False, w, 2, True, xStep, yBase, False, True)
        wmcom.drawCLTSection(True, d, o, False, w, 2, False, 2 * xStep, yBase, False, True)

        ' Render all sections to canvas1 for visual inspection
        Dim wmd As Drawing = wmcom.drawing
        wmd.ContextObject.Item = Me.canvas1
        wmd.Context = Contexts.WPFCanvas
        wmd.ContextObject.Margin = New WMDraw.Margin(10)
        wmd.ContextObject.Margin.Reference = Reference.contextMillimeters
        wmd.ContextObject.fitHeight = True
        wmd.ContextObject.fitWidth = True
        wmd.ContextObject.fitProportional = True
        wmd.draw()
    End Sub

End Class
