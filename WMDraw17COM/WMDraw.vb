Imports System.Runtime.InteropServices
Imports WallnerMild.WMDraw
Imports wmg = WallnerMild.Geom

<Guid("A3A1019E-474D-413F-A26F-0180C040C45B"),
        ClassInterface(ClassInterfaceType.AutoDual),
        ProgId("WMComDraw.c_WMComDraw")>
Public Class c_WMComDraw
    ''' <summary>
    ''' Core Drawing Object
    ''' </summary>
    Private p_wmd As New Drawing
    ''' <summary>
    ''' Current reference
    ''' </summary>
    Private p_ref As References
    ''' <summary>
    ''' Mirror of WallnerMild.Draw.Reference
    ''' </summary>
    Public Enum References
        world = 0
        contextMillimeters = 1
        contextUnits = 2
        contextFraction = 3
    End Enum
    ''' <summary>
    ''' Mirror of WallnerMild.Draw.Reference
    ''' </summary>
    ''' <returns></returns>
    Public Property CurRef As References
        Get
            Return p_ref
        End Get
        Set(value As References)
            p_ref = value
        End Set
    End Property

    ''' <summary>
    ''' Constructor
    ''' </summary>
    Public Sub New()
        p_wmd = New Drawing
        p_ref = References.world
        p_wmd.ContextObject.fitHeight = True
        p_wmd.ContextObject.fitWidth = True
        p_wmd.ContextObject.fitProportional = True
    End Sub
    ''' <summary>
    ''' reset drawing object
    ''' </summary>
    Public Sub clear()
        p_wmd = New Drawing
    End Sub
    ''' <summary>
    ''' Add a line
    ''' </summary>
    ''' <param name="startX"></param>
    ''' <param name="startY"></param>
    ''' <param name="endX"></param>
    ''' <param name="endY"></param>
    Public Sub addLine(startX As Double, startY As Double, endX As Double, endY As Double)
        'Dim l As New Line(startX, startY, endX, endY, WallnerMild.Draw.Reference.world)
        Dim l As New Line(startX, startY, endX, endY, CurRef)
        p_wmd.add(l)
    End Sub
    ''' <summary>
    ''' Add text
    ''' </summary>
    ''' <param name="startX"></param>
    ''' <param name="startY"></param>
    ''' <param name="Text"></param>
    Public Sub addText(startX As Double, startY As Double, Text As String)
        Dim t As New Text
        t.position = New Point(startX, startY)
        t.text = Text
        't.angle = angle
        p_wmd.add(t)
    End Sub

    Public Sub addSupport(posX As Double, posY As Double)
        Dim s As New Support(posX, posY, 0)
        s.position.coordinateReference = CurRef
        p_wmd.add(s)
    End Sub
    ''' <summary>
    ''' Draw To File
    ''' </summary>
    ''' <param name="sizeX"></param>
    ''' <param name="sizeY"></param>
    ''' <param name="unit"></param>
    ''' <returns></returns>
    Public Function drawToFile(sizeX As Double, sizeY As Double, Optional unit As String = "px") As String
        p_wmd.setContext(Contexts.PNGFile, sizeX, sizeY, unit)
        'p_wmd.setContext (Contexts.PNGFile,)
        Try
            p_wmd.draw()
            Return p_wmd.FileName
        Catch ex As Exception
            MsgBox(ex.Message)
            Return String.Empty
        End Try

    End Function
    ''' <summary>
    ''' Draw to Cliboard
    ''' </summary>
    ''' <param name="sizeX"></param>
    ''' <param name="sizeY"></param>
    ''' <param name="unit"></param>
    Public Sub drawToClipboard(sizeX As Double, sizeY As Double, Optional unit As String = "px")
        p_wmd.setContext(Contexts.PNGClipboard, sizeX, sizeY, unit)
        Try
            p_wmd.draw()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Function drawDowelGrid(toClipboard As Boolean, d As Single, alpha1 As Single, alpha2 As Single, a1 As Single, nRows As Integer, nColumns As Integer,
                                  Optional a2 As Single = -1, Optional a3 As Single = -1, Optional a4 As Single = -1)

        p_wmd = New WMDraw.Drawing
        If toClipboard Then
            p_wmd.setContext(Contexts.PNGClipboard, 90, 90, "mm")
        Else
            p_wmd.setContext(Contexts.PNGFile, 90, 90, "mm")
        End If

        Dim dAlpha As Single
        dAlpha = alpha2 - alpha1
        If dAlpha < 15 Then dAlpha = 90

        ' quer zur Faserrichtung
        If a2 = -1 Then
            a2 = 5 * d
        End If

        ' hirnholz
        If a3 = -1 Then
            a3 = 10 * d
        End If

        ' rand
        If a4 = -1 Then
            a4 = 7 * d
        End If

        Dim b As Single
        Dim h As Single

        Dim v As New wmg.Vector
        Dim n As New wmg.Vector

        ' create unit vector in D Alpha-direction
        v.angle = dAlpha * Math.PI / 180
        n = v.normal

        '
        ' skew is x' = x + y*tan(alpha)
        '
        Dim t As Single
        t = Math.Tan(Math.PI / 2 - dAlpha * Math.PI / 180)
        '
        ' horizontal lines with distance a1 and length (nColumns-1) * a2 + a3
        '
        Dim aRows As New List(Of wmg.line)
        Dim aCols As New List(Of wmg.line)

        ' Faserrichtung horizontal

        b = 2 * a4 + (nRows - 1) * a2
        h = 2 * a3 + (nColumns - 1) * a2

        ' rows (horizontal)
        '
        For iRow As Integer = 1 To nRows
            aRows.Add(New wmg.line(New wmg.Point(a3 / 2 + (a4 + (iRow - 1) * a2) * t, a4 + (iRow - 1) * a2),
                                   New wmg.Point(h - a3 / 2 + (a4 + (iRow - 1) * a2) * t, a4 + (iRow - 1) * a2)))
        Next
        '
        ' coliumns (vertical)
        '
        For iCol As Integer = 1 To nColumns
            aCols.Add(New wmg.line(New wmg.Point((a3 + (iCol - 1) * a1) + t * (a4 / 2), (a4 / 2)),
                                   New wmg.Point((a3 + (iCol - 1) * a1) + t * (b - a4 / 2), (b - a4 / 2))))
        Next


        '
        ' draw lines
        '
        With p_wmd
            .ContextObject.fitHeight = True
            .ContextObject.fitWidth = True
            .ContextObject.fitProportional = True
            .ContextObject.Margin = New Margin(5, 5, 5, 5)

            Dim f As New fill()
            f.color = WMColors.WoodLightYellow
            Dim r As New Rectangle(0, 0, h, b)
            r.fill = f

            Dim poly As New WallnerMild.WMDraw.Polygon
            poly.fill = f
            poly.points.Add(New Point(0, 0))
            poly.points.Add(New Point(h, 0))
            poly.points.Add(New Point(h + b * t, b))
            poly.points.Add(New Point(0 + b * t, b))
            poly.close()
            poly.pen = New pen()
            poly.pen.thickness = 0
            .add(poly)

            Dim p As New pen()
            p.color = WMColors.DarkGrey
            p.opacity = 0.6
            p.thickness = 0.5

            .pen = p
            For Each myRow In aRows
                .add(New Line(myRow.P.x, myRow.P.y, myRow.P_End.x, myRow.P_End.y))
            Next

            For Each myCol In aCols
                .add(New Line(myCol.P.x, myCol.P.y, myCol.P_End.x, myCol.P_End.y))
            Next

            Dim f2 As New fill()
            f2.color = WMColors.DarkWallnerMildBlue
            Dim e As New Ellipse

            ' rows (horizontal)
            '
            For iRow As Integer = 1 To nRows
                For iCol As Integer = 1 To nColumns
                    e = New Ellipse()
                    e.midPoint = New Point(a3 + (iCol - 1) * a1 + t * (a4 + (iRow - 1) * a2), a4 + (iRow - 1) * a2)
                    e.radii.height = d
                    e.radii.width = d
                    e.fill = f2
                    .add(e)
                Next
            Next

            Dim dl As New DimensionLine
            dl.startPoint = New Point(a3 + (nColumns - 1) * a1 + t * (a4 + (nRows - 2) * a2), a4 + (nRows - 2) * a2)
            dl.endPoint = New Point(a3 + (nColumns - 1) * a1 + t * (a4 + (nRows - 1) * a2), a4 + (nRows - 1) * a2)
            dl.textOverwrite = "a2"
            dl.offset.average = 5
            dl.symbolSize.average = 2
            dl.alignment = DimensionLine.DimAlignement.vertical
            .add(dl)

            Dim dl2 As New DimensionLine
            dl2.startPoint = New Point(a3 + (nColumns - 1) * a1 + t * (a4 + (nRows - 1) * a2), a4 + (nRows - 1) * a2)
            dl2.endPoint = New Point(a3 + (nColumns - 2) * a1 + t * (a4 + (nRows - 1) * a2), a4 + (nRows - 1) * a2)
            dl2.textOverwrite = "a1"
            dl2.offset.average = 5
            dl2.symbolSize.average = 2
            .add(dl2)

            '
            ' Faserrichtung
            '

            Dim frl1 As New Line()
            frl1.startPoint = New Point(5 + 5 * t, 5, Reference.contextMillimeters)
            frl1.endPoint = New Point(15 + 5 * t, 5, Reference.contextMillimeters)

            Dim frl2 As New Line()
            frl2.startPoint = New Point(5 + 5 * t + 1, 6, Reference.contextMillimeters)
            frl2.endPoint = New Point(5 + 5 * t, 5, Reference.contextMillimeters)

            Dim frl3 As New Line()
            frl3.startPoint = New Point(15 + 5 * t - 1, 4, Reference.contextMillimeters)
            frl3.endPoint = New Point(15 + 5 * t, 5, Reference.contextMillimeters)

            .add(frl1)
            .add(frl2)
            .add(frl3)

            '
            ' Kraft
            '
            Dim frc As New ForceArrow
            Dim p3 As New pen
            p3.color = WMColors.DarkRed
            p3.thickness = 2
            .pen = p3

            frc = New ForceArrow
            frc.tipPoint = New Point(a3 + t * a4, a4)
            frc.angle = alpha1 - 180
            frc.F = 0.8
            frc.ForceDirection = ForceArrow.forceDirections.ArrowAtBottom
            .add(frc)

        End With


        Return (p_wmd.draw())

    End Function


    Public Sub test3()
        Dim p As New Drawing
        MsgBox(TypeName(p))
    End Sub
    Public Sub test1()

        p_wmd.setContext(Contexts.PNGClipboard, 500, 500)
        Randomize()

        For i As Integer = 1 To 100
            Dim x1 As Double
            Dim y1 As Double
            x1 = 500 * Rnd()
            y1 = 500 * Rnd()
            Dim l As New Line(x1, y1, 500 * Rnd(), 500 * Rnd())
            p_wmd.add(l)
            Dim t As New Text
            t.text = CStr(Now())
            t.position = New Point(x1, y1)
            t.angle = 90 * Rnd()
            p_wmd.add(t)
        Next

        p_wmd.draw()

    End Sub

    Public Sub test4()
        MsgBox("Test4", vbOKOnly)
    End Sub
End Class
