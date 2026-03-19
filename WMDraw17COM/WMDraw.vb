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
    ''' Access the internal Drawing object. Not exposed to COM.
    ''' Allows test code to redirect the same Drawing to a WPF canvas after a draw method has been called.
    ''' </summary>
    <System.Runtime.InteropServices.ComVisible(False)>
    Public ReadOnly Property drawing As Drawing
        Get
            Return p_wmd
        End Get
    End Property
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

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="toClipboard"></param>
    ''' <param name="r"></param>
    ''' <param name=" hz"></param>
    ''' <param name=" bz"></param>
    ''' <param name=" lz"></param>
    ''' <param name=" tz"></param>
    ''' <param name=" gamma"></param>
    ''' <param name=" beta"></param>y
    ''' <param name=" a"></param>
    ''' <param name=" bN"></param>
    ''' <param name=" hN"></param>
    ''' <param name=" bH"></param> 
    ''' <param name=" hH"></param> 
    ''' <returns></returns>
    ''' <summary>
    ''' Draw a dovetail timber joint. Delegates all logic to the core DoveTail drawable.
    ''' </summary>
    Public Function drawDoveTail(toClipboard As Boolean, typeIndex As Integer,
                                 r As Single, hz As Single, bz As Single,
                                 lz As Single, tz As Single, gamma As Single,
                                 beta As Single, a As Single,
                                 bN As Single, hN As Single,
                                 bH As Single, hH As Single) As Object

        p_wmd = New WMDraw.Drawing
        If toClipboard Then
            p_wmd.setContext(Contexts.PNGClipboard, 160, 100, "mm")
        Else
            p_wmd.setContext(Contexts.PNGFile, 160, 100, "mm")
        End If
        With p_wmd
            .ContextObject.fitProportional = True
            .ContextObject.fitHeight = True
            .ContextObject.fitWidth = True
            .ContextObject.Margin = New WMDraw.Margin(5, 5, 5, 5)
        End With

        p_wmd.add(New DoveTail(r, hz, bz, lz, tz, gamma, beta, a, bN, hN, bH, hH))
        Return p_wmd.draw()

    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="toClipboard"></param>
    ''' <param name="typeIndex">index of dowel type fastener</param>
    ''' <param name="d">diameter</param>
    ''' <param name="b1">width of part 1</param>
    ''' <param name="h1" >byreference height of part 1</param>
    ''' <param name="b2"></param>
    ''' <param name="h2"></param>
    ''' <param name="beta"></param>
    ''' <param name="alphaF1"></param>
    ''' <param name="nRows"></param>
    ''' <param name="nColumns"></param>
    ''' <param name="a1"></param>
    ''' <param name="a2"></param>
    ''' <param name="a3c"></param>
    ''' <param name="a3t"></param>
    ''' <param name="a4c"></param>
    ''' <param name="a4t"></param>
    ''' <returns>required values for h1 or h2, if changed</returns>
    Public Function drawDowelGrid(toClipboard As Boolean, typeIndex As Integer, d As Single,
                                  b1 As Single, ByRef h1 As Single,
                                  b2 As Single, ByRef h2 As Single, beta As Single,
                                  alphaF1 As Single,
                                  nRows As Integer, nColumns As Integer,
                                  Optional a1 As Single? = Nothing, Optional a2 As Single? = Nothing,
                                  Optional a3c As Single? = Nothing, Optional a3t As Single? = Nothing,
                                  Optional a4c As Single? = Nothing, Optional a4t As Single? = Nothing, Optional member2IsSteelPlate As Boolean = False) As Object
        '
        ' Drawing
        '
        p_wmd = New WMDraw.Drawing
        If toClipboard Then
            p_wmd.setContext(Contexts.PNGClipboard, 160, 100, "mm")
        Else
            p_wmd.setContext(Contexts.PNGFile, 160, 100, "mm")
        End If

        Dim dAlpha As Single
        dAlpha = 0


        Dim dtf As New DowelTypeFastener

        '
        ' get spacing values, if not provided
        '
        dtf.dowelType = typeIndex
        dtf.d = d

        If a1 Is Nothing Then
            a1 = dtf.a1
        End If

        If a2 Is Nothing Then
            a2 = dtf.a2
        End If

        If a3c Is Nothing Then
            a3c = dtf.a3c
        End If

        If a3t Is Nothing Then
            a3t = dtf.a3t
        End If

        If a4c Is Nothing Then
            a4c = dtf.a4c
        End If

        If a4t Is Nothing Then
            a4t = dtf.a4t
        End If

        '
        ' direction vector in part 1 x-axis
        '
        Dim vPart1x As New wmg.Vector(1, 0)

        '
        ' direction vector in part 2 x-axis
        '
        Dim vPart2x As New wmg.Vector(1, 0)
        vPart2x.angle = beta * Math.PI / 180

        'Dim vPart1x As New wmg.Vector(1, 0)
        'Dim vPart2x As New wmg.Vector(1, 0)
        'vPart2x.angle = beta * Math.PI / 180

        '
        ' inner dowel type fastener 
        '

        ' spacing in part 1
        dtf.spacing1 = New spacing(1, 0, a1, a2, a1, a2)
        ' spacing in part 2
        If member2IsSteelPlate Then
            dtf.spacing2 = New spacing(1, 0, d * 2.2, d * 2.2, d * 2.2, d * 2.2)
        Else
            dtf.spacing2 = New spacing(a1, a2, a1, a2)
            dtf.spacing2.direction = vPart2x
        End If

        '
        ' intersection of the two spacing definitions
        ' intersection points to the neighbouring positions
        '
        ' v(0) is the vector to the next dtf in x-direction of part 1
        ' v(1) is the vector to the next dtf in y-direction of part 2
        '
        Dim v() As wmg.Vector
        If beta = 0 Then
            ReDim v(1)
            v(0) = New wmg.Vector(a1, 0)
            v(1) = New wmg.Vector(0, a2)
        Else
            v = dtf.spacing1.intersection(dtf.spacing2)
        End If



        '
        ' corner dowel type fastener
        '

        ' store distances for corner points (local quadrants I, II, III and IV correspond to indices 0,1,2 and3)
        Dim dq1(3) As wmg.Vector
        Dim dq2(3) As wmg.Vector
        ' store distance strings
        Dim dqi1(3, 1) As String
        Dim dqi2(3, 1) As String

        '
        ' edge distances in part 1
        '

        Dim lLeft As New wmg.line
        Dim borders(3) As wmg.line

        getEdgeDistances(alphaF1, a3t, a3c, a4t, a4c, dq1, dqi1)

        '
        'edge distances in part 2
        '
        Dim alphaF2 As Double
        alphaF2 = (180 - beta + alphaF1)

        ' edge distances
        If member2IsSteelPlate Then
            getEdgeDistances(alphaF2, d * 2, d * 2, d * 2, d * 2, dq2, dqi2)
        Else
            getEdgeDistances(alphaF2, a3t, a3c, a4t, a4c, dq2, dqi2)
        End If

        '
        ' compare required space with available space (from h1 and h2)
        ' by projecting the vectors to the local coordinate system
        '

        ' project spacing vectors to the two y-axes

        Dim vPart1y As wmg.Vector
        Dim vPart2y As wmg.Vector

        If beta = 0 Then
            vPart1y = v(1).normal(False)
            vPart2y = v(0).normal(False)
        Else
            vPart1y = v(1).project(vPart1x.normal)
            vPart2y = v(0).project(vPart2x.normal)
        End If


        Dim h1Available As Double = h1 - Math.Abs(dq1(1).y) - Math.Abs(dq1(3).y)
        Dim dh1Available As Double

        If nRows < 1 Then nRows = 1
        If nColumns < 1 Then nColumns = 1

        If nRows = 1 Then
            dh1Available = h1Available
        Else
            dh1Available = h1Available / (nRows - 1)
        End If

        Dim h2Available As Double = h2 - Math.Abs(dq2(1).y) - Math.Abs(dq2(3).y)
        Dim dh2Available As Double
        If nColumns = 1 Then
            dh2Available = h2Available
        Else
            dh2Available = h2Available / (nColumns - 1 + 1 - 1)
        End If

        Dim dimLinePen As New pen
        Dim dimLinePen1 As New pen
        Dim dimLinePen2 As New pen

        With dimLinePen
            .color = WMColors.DarkGrey
            .size = New size(0.1, Reference.contextMillimeters)
        End With

        With dimLinePen1
            .color = WMColors.DarkGrey
            .size = New size(0.1, Reference.contextMillimeters)
        End With

        With dimLinePen2
            .color = WMColors.DarkGrey
            .size = New size(0.1, Reference.contextMillimeters)
        End With

        Dim h1Overshoot As Boolean
        Dim h2Overshoot As Boolean

        If vPart1y.length > dh1Available Then
            ' ! not enough space in member 1 height !
            h1Overshoot = True
            With dimLinePen1
                .color = WMColors.Red
                .size = New size(0.15, Reference.contextMillimeters)
            End With
        Else
            h1Overshoot = False
            ' stretch vector v(1)
            Debug.Print("projected vector legth: {0}, from given h1={1} minus a3c={2} and a3t={3} divided by {4} rows: Available space is {5}", vPart1y.length, h1, a3c, a3t, nRows, dh1Available)
            Debug.Print("scale {0}", dh1Available / vPart1y.length)
            If beta = 0 Then
                v(1) = v(1).multiply(dh1Available / vPart1y.length)
            Else
                Dim a As wmg.Vector = v(1).project(vPart2x).multiply(dh1Available / vPart1y.length)
                Dim b As wmg.Vector = v(1).project(vPart2x.normal)
                v(1) = a.add(b)
            End If
        End If

        If vPart2y.length > dh2Available Then
            ' ! not enough space in member 2 height !
            h2Overshoot = True

            With dimLinePen2
                .color = WMColors.Red
                .size = New size(0.15, Reference.contextMillimeters)
            End With
        Else
            ' stretch vector v(0)
            h2Overshoot = False
            Debug.Print("projected vector legth: {0}, from given h1={1} minus a3c={2} and a3t={3} divided by {4} columns: Available space is {5}", vPart2y.length, h2, a3c, a3t, nColumns, dh2Available)
            Debug.Print("scale {0}", dh2Available / vPart2y.length)
            v(0) = v(0).multiply(dh2Available / vPart2y.length)
        End If

        '
        ' part 1 
        '
        Dim edgePolygon1 As New Polygon
        Dim edgePolygon2 As New Polygon
        Dim ixs As New wmg.Point

        If beta = 0 Then
            Dim vp As wmg.Vector
            vp = New wmg.Vector(0, 0)
            vp = vp.add(v(0).multiply(nColumns - 1))
            vp = vp.add(dq1(3))
            edgePolygon1.points.Add(New Point(vp.x, vp.y))

            vp = New wmg.Vector(0, 0)
            vp = vp.add(v(0).multiply(nColumns - 1))
            vp = vp.add(v(1).multiply(nRows - 1))
            vp = vp.add(dq1(0))
            edgePolygon1.points.Add(New Point(vp.x, vp.y))

            vp = New wmg.Vector(0, 0)
            vp = vp.add(v(1).multiply(nRows - 1))
            vp = vp.add(dq1(1))
            edgePolygon1.points.Add(New Point(vp.x, vp.y))

            vp = New wmg.Vector(0, 0)
            vp = vp.add(dq1(2))
            edgePolygon1.points.Add(New Point(vp.x, vp.y))
        Else

            ' part 1 border: right border line
            borders(0) = New wmg.Line

            With borders(0)
                .StartPoint = New wmg.Vector(0, 0)
                .StartPoint = .StartPoint.add(v(0).multiply(nColumns - 1))
                .StartPoint = .StartPoint.add(vPart1x.multiply(dq1(3).x))
                .Direction = vPart2x
            End With

            ' top border line
            borders(1) = New wmg.line
            With borders(1)
                .StartPoint = New wmg.Vector(0, 0)
                .StartPoint = .StartPoint.add(v(0).multiply(nColumns - 1))
                .StartPoint = .StartPoint.add(v(1).multiply(nRows - 1))
                .StartPoint = .StartPoint.add(vPart1x.normal.multiply(dq1(0).y))
                .Direction = vPart1x
            End With

            ' left border
            borders(2) = New wmg.line
            With borders(2)
                .StartPoint = New wmg.Vector(0, 0)
                .StartPoint = .StartPoint.add(vPart1x.multiply(dq1(2).x))
                .Direction = vPart2x
            End With

            ' bottom border line
            borders(3) = New wmg.line
            With borders(3)
                .StartPoint = New wmg.Vector(0, 0)
                .StartPoint = .StartPoint.add(vPart1x.normal.multiply(dq1(2).y))
                .Direction = vPart1x
            End With
            '
            ' intersect borders to get edge points of part 1
            '
            For i = 0 To 3
                Try
                    ixs = borders(i).intersect(borders((i + 1) Mod 4))
                Catch ex As InvalidOperationException
                    ' lines are parallel
                    ixs = borders(i).StartPoint.toPoint
                End Try
                edgePolygon1.points.Add(New Point(ixs.x, ixs.y))
            Next
        End If

        '
        ' part 2 boder
        '
        If beta = 0 Then
            Dim vp As wmg.Vector
            vp = New wmg.Vector(0, 0)
            vp = vp.add(v(0).multiply(nColumns - 1))
            vp = vp.add(dq2(3))
            edgePolygon2.points.Add(New Point(vp.x, vp.y))

            vp = New wmg.Vector(0, 0)
            vp = vp.add(v(0).multiply(nColumns - 1))
            vp = vp.add(v(1).multiply(nRows - 1))
            vp = vp.add(dq2(0))
            edgePolygon2.points.Add(New Point(vp.x, vp.y))

            vp = New wmg.Vector(0, 0)
            vp = vp.add(v(1).multiply(nRows - 1))
            vp = vp.add(dq2(1))
            edgePolygon2.points.Add(New Point(vp.x, vp.y))

            vp = New wmg.Vector(0, 0)
            vp = vp.add(dq2(2))
            edgePolygon2.points.Add(New Point(vp.x, vp.y))
        Else

            ' part 2 border: quadrant 4 - border line
            borders(0) = New wmg.line
            With borders(0)
                .StartPoint = New wmg.Vector(0, 0)
                .StartPoint = .StartPoint.add(v(0).multiply(nColumns - 1))
                .StartPoint = .StartPoint.add(v(1).multiply(nRows - 1))
                .StartPoint = .StartPoint.add(vPart2x.multiply(dq2(3).x))
                .Direction = vPart1x
            End With

            ' part 2 border: quadrant 1 - border line
            borders(1) = New wmg.line
            With borders(1)
                .StartPoint = New wmg.Vector(0, 0)
                .StartPoint = .StartPoint.add(v(1).multiply(nRows - 1))
                .StartPoint = .StartPoint.add(vPart2x.normal.multiply(dq2(0).y))
                .Direction = v(1)
            End With

            ' part 2 border: quadrant 2 - border line
            borders(2) = New wmg.line
            With borders(2)
                .StartPoint = New wmg.Vector(0, 0)
                .StartPoint = .StartPoint.add(vPart2x.multiply(dq2(2).x))
                .Direction = vPart1x
            End With

            ' part 2 border: quadrant 3 - border line
            borders(3) = New wmg.line
            With borders(3)
                .StartPoint = New wmg.Vector(0, 0)
                .StartPoint = .StartPoint.add(v(0).multiply(nColumns - 1))
                .StartPoint = .StartPoint.add(vPart2x.normal.multiply(dq2(2).y))
                .Direction = v(1)
            End With

            '
            ' now intersect to get border of part 2
            '
            ixs = New wmg.Point
            For i = 0 To 3
                Debug.Print("{0}; {1}", i, (i + 1) Mod 4)
                ixs = borders(i).intersect(borders((i + 1) Mod 4))
                edgePolygon2.points.Add(New Point(ixs.x, ixs.y))
            Next
        End If

#If DEBUG Then
        For Each ve As wmg.Vector In v
            Debug.Print("Vector ({0}/{1})", ve.x, ve.y)
        Next
#End If
        '
        ' find measure points
        '
        Dim measureLine1 As New wmg.Line
        measureLine1.StartPoint = New wmg.Vector(edgePolygon2.points(3).x, edgePolygon2.points(3).y)
        measureLine1.Direction = vPart2x

        Dim measureLine2 As New wmg.Line
        measureLine2.StartPoint = New wmg.Vector(edgePolygon2.points(0).x, edgePolygon2.points(0).y)
        measureLine2.Direction = vPart2y


        Dim measurePoint1 As wmg.Point
        Dim measurePoint2 As wmg.Point
        Try
            measurePoint1 = measureLine1.intersect(measureLine2)
        Catch ex As Exception
            measurePoint1 = measureLine1.StartPoint.toPoint
        End Try
        Try
            measurePoint2 = measureLine2.StartPoint.toPoint
        Catch ex As Exception
            measurePoint2 = measureLine2.StartPoint.toPoint
        End Try

        '
        ' horizontal lines with distance a1 and length (nColumns-1) * a2 + a3
        '


        '
        ' draw dowel net lines
        '
        With p_wmd
            .ContextObject.fitHeight = True
            .ContextObject.fitWidth = True
            .ContextObject.fitProportional = True
            .ContextObject.Margin = New WMDraw.Margin(5, 5, 5, 5)

            Dim pNetLines As New pen()

            With pNetLines
                .color = WMColors.DarkGrey
                .opacity = 0.9
                .thickness = 0.1
                .thicknessReference = Reference.contextMillimeters
            End With

            .pen = pNetLines

            '
            ' start point
            '
            Dim ps As wmg.Vector
            '
            ' end point
            '
            Dim pe As wmg.Vector

            '
            ' draw rows (x-axis of part 1)
            '
            For i = 1 To nRows - 1 + 1
                ps = New wmg.Vector(0, 0)
                ps = ps.add(v(1).multiply(i - 1))
                pe = ps.add(v(0).multiply(CDbl(nColumns - 1)))
                .add(New Line(New Point(ps.x, ps.y), New Point(pe.x, pe.y)))
            Next
            '
            ' draw columns (x-axis of part 2)
            '
            For i = 1 To nColumns - 1 + 1
                ps = New wmg.Vector(0, 0)
                ps = ps.add(v(0).multiply(i - 1))
                pe = ps.add(v(1).multiply(CDbl(nRows - 1)))
                .add(New Line(New Point(ps.x, ps.y), New Point(pe.x, pe.y)))
            Next

            '
            ' color and fill for edge polygons
            '
            Dim penPart1 As New pen
            With penPart1
                .color = WMColors.DarkRed
                .dashString = "7, 3"
            End With
            Dim part1Fill As New fill
            part1Fill.color = WMColors.CalculateTransparentColor(WMColors.WoodMediumYellow, 0.5)

            '
            ' add edge polygon of part 1
            '
            edgePolygon1.fill = part1Fill
            edgePolygon1.pen = penPart1
            .add(edgePolygon1)

            '
            ' add edge polygon of part 2
            '
            Dim part2Fill As New fill
            If member2IsSteelPlate Then
                part2Fill.color = WMColors.CalculateTransparentColor(WMColors.LightWallnerMildBlue, 0.3)
            Else
                part2Fill.color = WMColors.CalculateTransparentColor(WMColors.WoodLightYellow, 0.3)
            End If
            Dim penPart2 As New pen
            With penPart2
                .color = WMColors.DarkGreen
                .dashString = "7, 3"
            End With
            edgePolygon2.fill = part2Fill
            edgePolygon2.pen = penPart2
            .add(edgePolygon2)


            '
            ' fill and pen for space-ellipses
            '
            Dim fSpacing1 As New fill
            With fSpacing1
                .type = fill.fillType.linearHatching
                .setLinearHatch(WMColors.CalculateTransparentColor(WMColors.DarkRed, 0.5, WMColors.WoodLightYellow), 1, 0.3, 2, 0)
            End With

            Dim fSpacing2 As New fill
            With fSpacing2
                .type = fill.fillType.linearHatching
                .setLinearHatch(WMColors.CalculateTransparentColor(WMColors.DarkGreen, 0.5, WMColors.WoodLightYellow), 1, 0.3, 2, 0)
            End With

            Dim pSpacing1 As New pen
            With pSpacing1
                .dashString = "4,1"
                .color = WMColors.CalculateTransparentColor(WMColors.DarkRed, 0.7, WMColors.WoodLightYellow)
                .thickness = 0.1
                .thicknessReference = Reference.contextMillimeters
            End With

            Dim pSpacing2 As New pen
            With pSpacing2
                .dashString = "4,1"
                .color = WMColors.CalculateTransparentColor(WMColors.DarkGreen, 0.7, WMColors.WoodLightYellow)
                .thickness = 0.1
                .thicknessReference = Reference.contextMillimeters
            End With


            '
            ' fill and pen for dowel
            '
            Dim fillDowel As New fill()
            fillDowel.color = WMColors.DarkWallnerMildBlue
            Dim dowelEllipse As New Ellipse
            Dim forceArr As New ForceArrow

            '
            ' dowel spacing ellipses
            ' only drawn, if more than 3x3 dowels
            '
            If nRows > 2 And nColumns >= 2 Then
                ps = New wmg.Vector(0, 0)
                ps = ps.add(v(1).multiply(2 - 1))
                ps = ps.add(v(0).multiply(2 - 1))
                '
                ' Spacing Ellipse for part 1 and dowel in second row and second column
                '
                Dim eSpacing1 As New Ellipse
                With eSpacing1
                    .midPoint = New Point(ps.x, ps.y)
                    .radii.height = dtf.spacing1.top
                    .radii.width = dtf.spacing1.right
                    .radii.Reference = Reference.world
                    .angle = dtf.spacing1.direction.angle * 180 / Math.PI
                    .fill = fSpacing1
                    .pen = pSpacing1
                End With

                .add(eSpacing1)

                '
                ' Spacing Ellipse for part 2 and dowel in second row and second column
                '
                Dim eSpacing2 As New Ellipse
                With eSpacing2
                    .midPoint = New Point(ps.x, ps.y)
                    .radii.height = dtf.spacing2.top
                    .radii.width = dtf.spacing2.right
                    .radii.Reference = Reference.world
                    .angle = dtf.spacing2.direction.angle * 180 / Math.PI
                    .fill = fSpacing2
                    .pen = pSpacing2
                End With
                .add(eSpacing2)
            End If

            '
            ' dowels
            '
            For iRow As Integer = 1 To nRows - 1 + 1
                For iCol As Integer = 1 To nColumns - 1 + 1
                    '
                    ' mid point
                    '
                    ps = New wmg.Vector(0, 0)
                    ps = ps.add(v(1).multiply(iRow - 1))
                    ps = ps.add(v(0).multiply(iCol - 1))
                    '
                    ' ellipse
                    '
                    dowelEllipse = New Ellipse()
                    With dowelEllipse
                        .midPoint = New Point(ps.x, ps.y)
                        .radii.height = d / 2
                        .radii.width = d / 2
                        .radii.Reference = Reference.world
                        .fill = fillDowel
                    End With
                    .add(dowelEllipse)

                    '
                    ' force arrow
                    '
                    forceArr = New ForceArrow
                    With forceArr
                        .tipPoint = dowelEllipse.midPoint
                        .arrowSize = New size(1, Reference:=Reference.contextMillimeters)
                        '.MaximumSizeMM = 3
                        .MaximumSize = New size(1.5 * d, Reference.world)
                        .ForceDirection = ForceArrow.forceDirections.ArrowAtBottom
                        .pen = New pen()
                        .pen.color = WMColors.DarkRed
                        .pen.thickness = 0.25
                        .pen.thicknessReference = Reference.contextMillimeters
                        .F = 1
                        .angle = alphaF1
                    End With
                    .add(forceArr)

                    forceArr = New ForceArrow
                    With forceArr
                        .tipPoint = dowelEllipse.midPoint
                        .arrowSize = New size(1, Reference:=Reference.contextMillimeters)
                        '.MaximumSizeMM = 3
                        .MaximumSize = New size(1.5 * d, Reference.world)
                        .ForceDirection = ForceArrow.forceDirections.ArrowAtBottom
                        .pen = New pen()
                        .pen.color = WMColors.DarkGreen
                        .pen.dashString = "2,1"
                        .pen.thickness = 0.25
                        .pen.thicknessReference = Reference.contextMillimeters
                        .F = 1
                        .angle = alphaF1 - 180
                    End With
                    .add(forceArr)

                Next
            Next

            '
            '  |        Dimension lines          |
            '  /---------------------------------/
            '  |                                 |
            '

            '
            ' part 1 vertical dimension lines
            '
            Dim dl As New DimensionLine
            Dim dl3 As New DimensionLine

            ' edge 1
            dl = New DimensionLine
            dl.textFormatString = "0"
            dl.textSize = New size(1.4, Reference.contextMillimeters)
            dl.pen = dimLinePen

            dl.offset.average = 5
            dl.symbolSize.average = 2
            dl.alignment = DimensionLine.DimAlignement.vertical

            dl3 = New DimensionLine
            dl3 = dl.copy

            dl3.startPoint = New Point(0, 0)
            dl3.endPoint = New Point(edgePolygon1.points(2).x, edgePolygon1.points(2).y)
            .add(dl3)

            dl3 = New DimensionLine
            dl3 = dl.copy
            dl3.pen = dimLinePen

            ps = New wmg.Vector(0, 0)
            ps = ps.add(v(1).multiply(nRows - 1))

            dl3.startPoint = New Point(edgePolygon1.points(1).x * 0, edgePolygon1.points(1).y)
            dl3.endPoint = New Point(0, ps.y)
            .add(dl3)


            '
            ' height h1
            '
            dl3 = New DimensionLine
            dl3 = dl.copy
            dl3.pen = dimLinePen1
            dl3.startPoint = New Point(edgePolygon1.points(1).x, edgePolygon1.points(1).y)
            dl3.endPoint = New Point(edgePolygon1.points(2).x, edgePolygon1.points(2).y)

            If h1Overshoot Then
                dl3.textPrefix = "!!  "
                dl3.textSuffix = "  !!"
            End If

            .add(dl3)

            '
            ' change h1 in case of overshoot
            '
            If h1Overshoot Then h1 = dl3.distance

            For iRow As Integer = 1 To nRows - 1
                dl = New DimensionLine
                ps = New wmg.Vector(0, 0)
                ps = ps.add(v(1).multiply(iRow - 1))
                pe = ps.add(v(1))
                dl.startPoint = New Point(0, pe.y)
                dl.endPoint = New Point(0, ps.y)
                dl.textFormatString = "0"
                dl.textSize = New size(1.4, Reference.contextMillimeters)
                dl.pen = New pen()
                dl.pen.color = WMColors.DarkGrey
                dl.pen.size = New size(0.1, Reference.contextMillimeters)

                dl.offset.average = 5
                dl.symbolSize.average = 2
                dl.alignment = DimensionLine.DimAlignement.vertical
                .add(dl)
            Next

            'part 1 horizontal dimension lines
            Dim dl2 As New DimensionLine
            dl2 = dl.copy()
            dl2.alignment = DimensionLine.DimAlignement.horizontal

            dl3 = New DimensionLine
            dl3 = dl2.copy
            dl3.startPoint = New Point(edgePolygon1.points(2).x, edgePolygon1.points(2).y * 0)
            dl3.endPoint = New Point(0, 0)
            .add(dl3)

            If beta = 0 Then
                dl3 = New DimensionLine
                dl3 = dl2.copy

                Dim vp As New wmg.Vector
                vp = New wmg.Vector(0, 0)
                vp = vp.add(v(0).multiply(nColumns - 1))

                dl3.endPoint = New Point(edgePolygon1.points(0).x, edgePolygon1.points(0).y * 0)
                dl3.startPoint = New Point(vp.x, vp.y * 0)

                .add(dl3)
            End If

            dl3 = New DimensionLine
            dl3 = dl2.copy

            dl3.startPoint = New Point(edgePolygon1.points(2).x, edgePolygon1.points(2).y)
            dl3.endPoint = New Point(edgePolygon1.points(3).x, edgePolygon1.points(3).y)
            .add(dl3)


            For icol As Integer = 1 To nColumns - 1
                dl3 = New DimensionLine
                dl3 = dl2.copy
                ps = New wmg.Vector(0, 0)
                ps = ps.add(v(0).multiply(icol - 1))
                pe = ps.add(v(0))
                dl3.startPoint = New Point(ps.x, 0)
                dl3.endPoint = New Point(pe.x, 0)
                .add(dl3)
            Next

            '
            ' part2 x-dimension lines
            '
            Dim dl5 As New DimensionLine
            pe = New wmg.Vector

            For iRow As Integer = 1 To nRows - 1
                dl5 = New DimensionLine
                dl5 = dl2.copy
                dl5.alignment = DimensionLine.DimAlignement.aligned
                dl5.offset = New size(dl5.offset.thickness * 1.5, dl5.offset.Reference)
                ps = New wmg.Vector(0, 0)
                ps = ps.add(v(0).multiply(nColumns - 1))
                ps = ps.add(v(1).multiply(iRow - 1))

                pe = ps.add(v(1))

                dl5.startPoint = New Point(ps.x, ps.y)
                dl5.endPoint = New Point(pe.x, pe.y)
                .add(dl5)

            Next


            '
            ' intersection of upper border of part 2 with rightmost column
            '
            measureLine1 = New wmg.Line
            measureLine1.StartPoint = pe
            measureLine1.Direction = v(1)

            measureLine2 = New wmg.Line
            measureLine2.StartPoint = New wmg.Vector(edgePolygon2.points(3).x, edgePolygon2.points(3).y)
            measureLine2.Direction = v(0)

            Dim measurePoint3 As wmg.Point = measureLine1.intersect(measureLine2)
            Dim measurePoint4 As wmg.Point = measureLine1.StartPoint.toPoint

            '
            ' intersection of lower border of part 2 with rightmost column
            '
            measureLine1 = New wmg.line
            ps = New wmg.Vector(0, 0)
            ps = ps.add(v(0).multiply(nColumns - 1))
            measureLine1.StartPoint = ps
            measureLine1.Direction = v(1)

            measureLine2 = New wmg.Line
            measureLine2.StartPoint = New wmg.Vector(edgePolygon2.points(1).x, edgePolygon2.points(1).y)
            measureLine2.Direction = v(0)

            Dim measurePoint5 As wmg.Point = measureLine1.intersect(measureLine2)
            Dim measurePoint6 As wmg.Point = measureLine1.StartPoint.toPoint

            Dim dl4 As New DimensionLine

            dl5 = New DimensionLine
            dl5 = dl2.copy
            dl5.alignment = DimensionLine.DimAlignement.aligned
            dl5.offset = New size(dl5.offset.thickness * 1.5, dl5.offset.Reference)
            dl5.startPoint = New Point(measurePoint4.x, measurePoint4.y)
            dl5.endPoint = New Point(measurePoint3.x, measurePoint3.y)
            .add(dl5)

            dl5 = New DimensionLine
            dl5 = dl2.copy
            dl5.alignment = DimensionLine.DimAlignement.aligned
            dl5.offset = New size(dl5.offset.thickness * 1.5, dl5.offset.Reference)
            dl5.startPoint = New Point(measurePoint5.x, measurePoint5.y)
            dl5.endPoint = New Point(measurePoint6.x, measurePoint6.y)
            .add(dl5)

            '
            ' part2 y-dimension lines
            '

            dl4 = New DimensionLine
            dl4 = dl2.copy
            dl4.pen = dimLinePen2
            dl4.offset = New size(dl4.offset.thickness * 2, dl4.offset.Reference)
            dl4.alignment = DimensionLine.DimAlignement.aligned
            dl4.startPoint = New Point(measurePoint1.x, measurePoint1.y)
            dl4.endPoint = New Point(measurePoint2.x, measurePoint2.y)

            If h2Overshoot Then
                dl4.textPrefix = "!!  "
                dl4.textSuffix = "  !!"
            End If
            .add(dl4)
            '
            ' change h2 value in case of overshoot
            '
            If h2Overshoot Then h2 = dl4.distance


            Dim dimPoint As New wmg.Vector
            Dim vProj As wmg.Vector
            vProj = v(0).project(vPart2y)

            pe = New wmg.Vector

            For iCol = 1 To nColumns - 1
                dl3 = New DimensionLine
                dl3 = dl2.copy
                dl3.alignment = DimensionLine.DimAlignement.aligned

                ps = New wmg.Vector(0, 0)
                ps = ps.add(v(1).multiply(nRows - 1))
                ps = ps.add(v(0).multiply(iCol - 1))
                If iCol = 1 Then
                    pe = ps.add(vProj)
                Else
                    ps = pe
                    pe = pe.add(vProj)
                End If

                dl3.endPoint = New Point(ps.x, ps.y)
                dl3.startPoint = New Point(pe.x, pe.y)
                .add(dl3)
            Next

            '
            ' intersection of lower border of part 2 with rightmost column
            '
            measureLine1 = New wmg.line
            ps = New wmg.Vector(0, 0)
            ps = ps.add(v(1).multiply(nRows - 1))
            measureLine1.StartPoint = ps
            measureLine1.Direction = vPart2y

            measureLine2 = New wmg.Line
            measureLine2.StartPoint = New wmg.Vector(edgePolygon2.points(1).x, edgePolygon2.points(1).y)
            measureLine2.Direction = vPart2x

            Dim measurePoint7 As wmg.Point = measureLine1.intersect(measureLine2)
            Dim measurePoint8 As wmg.Point = measureLine1.StartPoint.toPoint

            dl3 = New DimensionLine
            dl3 = dl2.copy
            dl3.alignment = DimensionLine.DimAlignement.aligned
            dl3.endPoint = New Point(measurePoint7.x, measurePoint7.y)
            dl3.startPoint = New Point(measurePoint8.x, measurePoint8.y)
            .add(dl3)


            ''
            '' Faserrichtung Bauteil 1
            ''

            ps = New wmg.Vector(0, 0)
            ps = ps.add(v(1).multiply(((nRows - 1) + 0.5) - 1))
            ps = ps.add(v(0).multiply(((nColumns - 1) + 0.5) - 1))
            Dim wg1 As New WoodGrainSymbol(ps.x, ps.y, 0)
            wg1.pen.color = WMColors.DarkRed
            .add(wg1)

            ps = New wmg.Vector(0, 0)
            ps = ps.add(v(1).multiply((nRows + 0.15) - 1))
            ps = ps.add(v(0).multiply(((nColumns - 1) + 0.5) - 1))

            Dim wg2 As New WoodGrainSymbol(ps.x, ps.y, beta)
            wg2.pen.color = WMColors.DarkGreen
            .add(wg2)

        End With


        Return (p_wmd.draw())

    End Function
    Private Sub getEdgeDistances(alphaF As Double, a3t As Double, a3c As Double, a4t As Double, a4c As Double, ByRef dq() As wmg.Vector, ByRef dqi(,) As String)
        ReDim dq(3)
        ReDim dqi(3, 1)
        Dim quadrant As Integer

        quadrant = Math.Floor(alphaF / 90)
        quadrant = quadrant Mod 4
        quadrant += 1

        Select Case quadrant
            Case 1
                ' quadrant 1
                dq(0) = New wmg.Vector(a3t, a4t)
                dqi(0, 0) = "a3c" : dqi(0, 1) = "a4c"
                ' quadrant 2
                dq(1) = New wmg.Vector(-a3c, a4t)
                dqi(1, 0) = "a3t" : dqi(1, 1) = "a4c"
                ' quadrant 3
                dq(2) = New wmg.Vector(-a3c, -a4c)
                dqi(2, 0) = "a3t" : dqi(2, 1) = "a4t"
                ' quadrant 4
                dq(3) = New wmg.Vector(a3t, -a4c)
                dqi(3, 0) = "a3c" : dqi(3, 1) = "a4t"
            Case 2
                ' quadrant 1
                dq(0) = New wmg.Vector(a3c, a4t)
                dqi(0, 0) = "a3c" : dqi(0, 1) = "a4t"
                ' quadrant 2
                dq(1) = New wmg.Vector(-a3t, a4t)
                dqi(1, 0) = "a3t" : dqi(1, 1) = "a4t"
                ' quadrant 3
                dq(2) = New wmg.Vector(-a3t, -a4c)
                dqi(2, 0) = "a3t" : dqi(2, 1) = "a4c"
                ' quadrant 4
                dq(3) = New wmg.Vector(a3c, -a4c)
                dqi(3, 0) = "a3c" : dqi(3, 1) = "a4c"
            Case 3
                ' quadrant 1
                dq(0) = New wmg.Vector(a3c, a4c)
                dqi(0, 0) = "a3t" : dqi(0, 1) = "a4t"
                ' quadrant 2
                dq(1) = New wmg.Vector(-a3t, a4c)
                dqi(1, 0) = "a3c" : dqi(1, 1) = "a4t"
                ' quadrant 3
                dq(2) = New wmg.Vector(-a3t, -a4t)
                dqi(2, 0) = "a3c" : dqi(2, 1) = "a4c"
                ' quadrant 4
                dq(3) = New wmg.Vector(a3c, -a4t)
                dqi(3, 0) = "a3t" : dqi(3, 1) = "a4c"
            Case 4
                ' quadrant 1
                dq(0) = New wmg.Vector(a3t, a4c)
                dqi(0, 0) = "a3c" : dqi(0, 1) = "a4t"
                ' quadrant 2
                dq(1) = New wmg.Vector(-a3c, a4c)
                dqi(1, 0) = "a3t" : dqi(1, 1) = "a4t"
                ' quadrant 3
                dq(2) = New wmg.Vector(-a3c, -a4c)
                dqi(2, 0) = "a3t" : dqi(2, 1) = "a4c"
                ' quadrant 4
                dq(3) = New wmg.Vector(a3t, -a4t)
                dqi(3, 0) = "a3c" : dqi(3, 1) = "a4c"
        End Select

    End Sub

    ''' <summary>
    ''' Draw a CLT cross-section. Delegates all logic to the core CLTSection drawable.
    ''' </summary>
    ''' <param name="toClipboard">True = copy PNG to clipboard; False = write PNG to file</param>
    ''' <param name="d">Array of layer thicknesses in mm (Variant array from VBA)</param>
    ''' <param name="o">Array of grain orientations, same size as d: 0 = parallel, 90 = perpendicular</param>
    ''' <param name="isVertical">True = vertical section; False = horizontal</param>
    ''' <param name="displayWidth">Width perpendicular to stacking direction, in mm (default 1000)</param>
    ''' <param name="showDimensions">0 = none; 1 = thickness labels; 2 = full dimension lines (default)</param>
    ''' <param name="drawCrossSection">True = draw as defined; False = swap 0°/90° display</param>
    ''' <param name="x0">X of bottom-left corner in world mm (default 0)</param>
    ''' <param name="y0">Y of bottom-left corner in world mm (default 0)</param>
    ''' <param name="resetDrawing">True = new drawing (default); False = append to existing</param>
    ''' <param name="useHatching">True = line hatching for 0° layers, transparent for 90°</param>
    Public Function drawCLTSection(toClipboard As Boolean,
                                   d As Object,
                                   o As Object,
                                   Optional isVertical As Boolean = True,
                                   Optional displayWidth As Single = 1000,
                                   Optional showDimensions As Integer = 2,
                                   Optional drawCrossSection As Boolean = True,
                                   Optional x0 As Double = 0,
                                   Optional y0 As Double = 0,
                                   Optional resetDrawing As Boolean = True,
                                   Optional useHatching As Boolean = False) As Object

        If resetDrawing Then
            p_wmd = New WMDraw.Drawing
            If toClipboard Then
                p_wmd.setContext(Contexts.PNGClipboard, 160, 100, "mm")
            Else
                p_wmd.setContext(Contexts.PNGFile, 160, 100, "mm")
            End If
            With p_wmd
                .ContextObject.fitProportional = True
                .ContextObject.fitHeight = True
                .ContextObject.fitWidth = True
                .ContextObject.Margin = New WMDraw.Margin(10, 10, 10, 10)
            End With
        End If

        ' Convert COM Variant arrays to typed arrays
        Dim n As Integer = UBound(d) - LBound(d) + 1
        Dim layers(n - 1) As Double
        Dim orients(n - 1) As Integer
        For i As Integer = 0 To n - 1
            layers(i) = CDbl(d(LBound(d) + i))
            orients(i) = CInt(o(LBound(o) + i))
        Next

        Dim clt As New CLTSection(layers, orients)
        clt.isVertical = isVertical
        clt.displayWidth = displayWidth
        clt.showDimensions = showDimensions
        clt.drawCrossSection = drawCrossSection
        clt.x0 = x0
        clt.y0 = y0
        clt.useHatching = useHatching

        p_wmd.add(clt)
        Return p_wmd.draw()

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
