Imports WMDraw17.WallnerMild.Draw
Imports WMDraw17.WallnerMild

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
        wmd.add(New Line(0, 0, 3, 1))
        wmd.add(New Line(3, 1, 4.5, 1.5))
        wmd.add(New Line(4.5, 1.5, 9, 3))

        'support
        wmd.add(New Support(3, 1, 0, "A"))
        wmd.add(New Support(4.5, 1.5, 0, "B"))
        wmd.add(New Support(9, 3, 30, "C"))

        'point
        wmd.add(New Point(3, 3, PointDisplay.x, 2, Reference.contextMillimeters))


        wmd.add(New Line(0, 0.25, 1, 0.25, Reference.contextFraction))
        wmd.add(New Point(0.5, 0.25, Reference.contextFraction, PointDisplay.x, 2, Reference.contextMillimeters))

        wmd.add(New Beam(3.5, 0.5, 8, 2))
        wmd.add(New BeamUniformLoad(3.5, 0.5, 8, 2, BeamUniformLoad.loadedLengthType.horizontalProjection, BeamUniformLoad.loadOrientationType.localOrientation, 1))
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
        Dim l1 As New Geom.line
        l1.P = New Geom.Vector(1, 1)
        l1.d = New Geom.Vector(1, 1)
        'l1.offsetLine(0.5)
        Dim l2 As New Geom.line
        l2.P = New Geom.Vector(3, 1)
        l2.d = New Geom.Vector(0, 1)
        'l2.offsetLine(0.5)

        Dim S As New Geom.Vector
        S = l1.intersect(l2)

        MsgBox(String.Format("S ({0}/{1})", S.x, S.y))
    End Sub

    Private Sub button2_Click(sender As Object, e As RoutedEventArgs) Handles button2.Click
        Dim wmcom As New WMDraw17COM.WMDraw
        wmcom.test1()
        MsgBox("PNG Copied to Clipboard")
    End Sub

    Private Sub button3_Click(sender As Object, e As RoutedEventArgs) Handles button3.Click
        Dim wmd As New Drawing

        '
        ' add some lines
        '

        ' World Coordinates
        wmd.add(New Line(0, 0, 3, 1))
        wmd.add(New Line(3, 1, 4.5, 1.5))
        wmd.add(New Line(4.5, 1.5, 9, 3))

        wmd.setContext(Contexts.PNGFile, 500, 500)
        MsgBox(wmd.draw())
    End Sub
End Class
