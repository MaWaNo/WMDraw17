Imports System.Runtime.InteropServices

Namespace WMDraw

#Region "DoveTail Class"

    ''' <summary>
    ''' Dovetail timber joint drawable.
    ''' Renders a cross-section view (left) and side view (right) of a dovetail notch connection
    ''' with full dimension annotations.
    ''' </summary>
    <ComVisible(False)>
    Public Class DoveTail
        Implements Drawable

        Private p_pen As New pen
        Private p_zIndex As Long

        ''' <summary>Radius of the dovetail round (mm).</summary>
        Public Property r As Double

        ''' <summary>Depth of the dovetail notch hz (mm).</summary>
        Public Property hz As Double

        ''' <summary>Width of the dovetail at the top bz (mm).</summary>
        Public Property bz As Double

        ''' <summary>Length of the dovetail lz (mm).</summary>
        Public Property lz As Double

        ''' <summary>Depth of the tapered face tz (mm).</summary>
        Public Property tz As Double

        ''' <summary>Opening angle gamma (degrees).</summary>
        Public Property gamma As Double

        ''' <summary>Taper angle beta (degrees).</summary>
        Public Property beta As Double

        ''' <summary>Notch distance a (mm).</summary>
        Public Property a As Double

        ''' <summary>Half-width of the secondary beam bN (mm).</summary>
        Public Property bN As Double

        ''' <summary>Height of the secondary beam hN (mm).</summary>
        Public Property hN As Double

        ''' <summary>Width of the primary beam bH (mm).</summary>
        Public Property bH As Double

        ''' <summary>Height of the primary beam hH (mm).</summary>
        Public Property hH As Double

        ''' <summary>
        ''' Creates a DoveTail drawable with all geometric parameters.
        ''' </summary>
        Public Sub New(r As Double, hz As Double, bz As Double, lz As Double, tz As Double,
                       gamma As Double, beta As Double, a As Double,
                       bN As Double, hN As Double, bH As Double, hH As Double)
            Me.r = r
            Me.hz = hz
            Me.bz = bz
            Me.lz = lz
            Me.tz = tz
            Me.gamma = gamma
            Me.beta = beta
            Me.a = a
            Me.bN = bN
            Me.hN = hN
            Me.bH = bH
            Me.hH = hH
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
            ' Section view: x from -bN to bN; side view: from 1.5*bN to 2*bN+bH; y from -hH to 0
            Dim r_arr(3) As Double
            r_arr(0) = -bN
            r_arr(1) = -hH
            r_arr(2) = 1.5 * bN + bH + 0.5 * bN
            r_arr(3) = 0
            Return r_arr
        End Function

        Public Sub draw(contextobject As ContextObject, contextCoordinatesDelegate As Drawable.contextCoordinates, Optional contextSizeDelegate As Drawable.contextSize = Nothing) Implements Drawable.draw
            For Each item In buildDrawables()
                item.draw(contextobject, contextCoordinatesDelegate, contextSizeDelegate)
            Next
        End Sub

        ''' <summary>
        ''' Builds all sub-drawables for the dovetail cross-section and side view.
        ''' </summary>
        Private Function buildDrawables() As List(Of Drawable)
            Dim items As New List(Of Drawable)

            ' --- Derived geometry ---
            Dim P1 As New Point(r * Math.Cos(rad(gamma) / 2), -hz + r - r * Math.Sin(rad(gamma) / 2))
            Dim P2 As New Point(-((r - hz) * Math.Sin(rad(gamma) / 2) - r) / Math.Cos(rad(gamma) / 2), 0)
            Dim P3 As New Point(0, -hz)
            Dim M As New Point(0, -hz + r)

            Dim P1m As New Point(-r * Math.Cos(rad(gamma) / 2), -hz + r - r * Math.Sin(rad(gamma) / 2))
            Dim P2m As New Point(+((r - hz) * Math.Sin(rad(gamma) / 2) - r) / Math.Cos(rad(gamma) / 2), 0)

            Dim dZ As Double = lz * Math.Tan(rad(beta))

            Dim P1_ As New Point(P1.x + dZ * Math.Cos(rad(gamma) / 2), P1.y - dZ * Math.Sin(rad(gamma) / 2))
            Dim P2_ As New Point(P2.x + dZ / Math.Cos(rad(gamma) / 2), 0)
            Dim P1m_ As New Point(-P1_.x, P1_.y)
            Dim P2m_ As New Point(-P2_.x, P2_.y)

            ' --- Pens ---
            Dim penHidden As New pen
            With penHidden
                .color = WMColors.Black
                .dashString = "7, 3"
            End With

            Dim penTransparent As New pen
            penTransparent.thickness = 0

            Dim penThin As New pen
            penThin.color = WMColors.Black

            ' --- Fills ---
            Dim fillTransp As New fill
            fillTransp.color = WMColors.CalculateTransparentColor(WMColors.LightWallnerMildBlue, 0)

            Dim fillWood1 As New fill
            fillWood1.color = WMColors.WoodLightYellow

            Dim fillWood2 As New fill
            fillWood2.color = WMColors.WoodMediumYellow

            Dim fontSize As New size(2, Reference.contextMillimeters)

            ' =============================================================
            ' Section view (cross-section, centred at x=0)
            ' =============================================================

            ' Primary beam background
            Dim RH As New Rectangle(-1 * bN, 0, 1 * bN, -hH)
            RH.pen = penTransparent
            RH.fill = fillWood2
            items.Add(RH)

            ' Primary beam top edge
            Dim okH As New Line(-1 * bN, 0, 1 * bN, 0)
            okH.pen = penThin
            items.Add(okH)

            ' Primary beam bottom edge
            Dim ukH As New Line(-1 * bN, -hH, 1 * bN, -hH)
            ukH.pen = penThin
            items.Add(ukH)

            ' Secondary beam notch rectangle
            Dim dr As New Rectangle(-bN / 2, 0, bN / 2, -hN)
            dr.pen = penThin
            dr.fill = fillWood1
            items.Add(dr)

            ' Centre point M
            Dim pM As New Point(M.x, M.y)
            pM.display = PointDisplay.x
            pM.displaySize = New size(1, Reference.contextMillimeters)
            pM.pen = penThin
            items.Add(pM)

            ' Inner dovetail polygon (hidden lines)
            Dim k As New Polygon
            k.points.Add(P2)
            k.points.Add(P1)
            For i As Integer = 0 To 20
                Dim xi As Double = gamma / 2 + (180 - gamma / 2) / 20 * i
                k.points.Add(New Point(M.x + r * Math.Cos(rad(xi)), M.y - r * Math.Sin(rad(xi))))
            Next
            k.points.Add(P1m)
            k.points.Add(P2m)
            k.pen = penHidden
            k.fill = fillTransp
            items.Add(k)

            ' Outer dovetail polygon (visible outline)
            Dim k_ As New Polygon
            k_.points.Add(P2_)
            k_.points.Add(P1_)
            For i As Integer = 0 To 20
                Dim xi As Double = gamma / 2 + (180 - gamma / 2) / 20 * i
                k_.points.Add(New Point(M.x + (r + dZ) * Math.Cos(rad(xi)), M.y - (r + dZ) * Math.Sin(rad(xi))))
            Next
            k_.points.Add(P1m_)
            k_.points.Add(P2m_)
            k_.pen = penThin
            k_.fill = fillTransp
            items.Add(k_)

            ' --- Dimension lines (section view) ---

            ' Distance from M to top surface
            Dim dl_Mpos As New DimensionLine()
            dl_Mpos.startPoint = M
            dl_Mpos.endPoint = New Point(0, 0)
            dl_Mpos.offset = New size(-2, Reference.contextMillimeters)
            dl_Mpos.textFormatString = "0.0"
            dl_Mpos.textSize = fontSize
            items.Add(dl_Mpos)

            ' Inner dovetail width bZ
            Dim dl_bZ As New DimensionLine()
            dl_bZ.startPoint = P2m
            dl_bZ.endPoint = P2
            dl_bZ.offset = New size(-5, Reference.contextMillimeters)
            dl_bZ.textFormatString = "bZ=0.0"
            dl_bZ.textSize = fontSize
            items.Add(dl_bZ)

            ' Outer dovetail width
            Dim dl_bZ2 As New DimensionLine()
            dl_bZ2.startPoint = P2m_
            dl_bZ2.endPoint = P2_
            dl_bZ2.offset = New size(5, Reference.contextMillimeters)
            dl_bZ2.textFormatString = "0.0"
            dl_bZ2.textSize = fontSize
            items.Add(dl_bZ2)

            ' Dovetail depth hZ
            Dim dl_hZ As New DimensionLine()
            dl_hZ.startPoint = P3
            dl_hZ.endPoint = P2
            dl_hZ.alignment = DimensionLine.DimAlignement.vertical
            dl_hZ.offset = New size(7, Reference.contextMillimeters)
            dl_hZ.textFormatString = "hZ=0.0"
            dl_hZ.textSize = fontSize
            items.Add(dl_hZ)

            ' Notch rest height ha
            Dim dl_ha As New DimensionLine()
            dl_ha.startPoint = New Point(0, -hN)
            dl_ha.endPoint = New Point(P2.x, P3.y)
            dl_ha.alignment = DimensionLine.DimAlignement.vertical
            dl_ha.offset = New size(7, Reference.contextMillimeters)
            dl_ha.textFormatString = "ha=0.0"
            dl_ha.textSize = fontSize
            items.Add(dl_ha)

            ' Secondary beam width bN
            Dim dl_bN As New DimensionLine()
            dl_bN.startPoint = New Point(-bN / 2, -hN)
            dl_bN.endPoint = New Point(bN / 2, -hN)
            dl_bN.offset = New size(7, Reference.contextMillimeters)
            dl_bN.textFormatString = "bN=0.0"
            dl_bN.textSize = fontSize
            items.Add(dl_bN)

            ' Secondary beam height hN
            Dim dl_hN As New DimensionLine()
            dl_hN.startPoint = New Point(-bN / 2, -hN)
            dl_hN.endPoint = New Point(-bN / 2, 0)
            dl_hN.offset = New size(-7, Reference.contextMillimeters)
            dl_hN.textFormatString = "hN=0.0"
            dl_hN.textSize = fontSize
            items.Add(dl_hN)

            ' Radius r
            Dim dl_MP1 As New DimensionLine()
            dl_MP1.startPoint = M
            dl_MP1.endPoint = P1
            dl_MP1.offset = New size(0, Reference.contextMillimeters)
            dl_MP1.textFormatString = "r=0.0"
            dl_MP1.textSize = fontSize
            items.Add(dl_MP1)

            ' Opening angle gamma
            Dim da_phi As New DimensionAngular(P1, P2, P1m, P2m)
            da_phi.dimSymbol = DimensionAngular.DimSymbols.Arrow
            da_phi.textSize = fontSize
            da_phi.textPrefix = ChrW(947) & "="
            da_phi.offset = New size(10, Reference.contextMillimeters)
            items.Add(da_phi)

            ' =============================================================
            ' Side view (shifted right by cx = 1.5 * bN)
            ' =============================================================

            Dim cx As Double = 1.5 * bN

            ' Primary beam side polygon
            Dim rH2 As New Polygon
            rH2.points.Add(New Point(cx + 0, 0))
            rH2.points.Add(New Point(cx + bH - tz, 0))
            rH2.points.Add(New Point(cx + bH - tz, -hN + a - dZ))
            rH2.points.Add(New Point(cx + bH, -hN + a))
            rH2.points.Add(New Point(cx + bH, -hH))
            rH2.points.Add(New Point(cx + 0, -hH))
            rH2.points.Add(New Point(cx + 0, 0))
            rH2.pen = penThin
            rH2.fill = fillWood2
            items.Add(rH2)

            ' Secondary beam side polygon
            Dim P40 As New Point(cx + bH - lz, 0)
            Dim P4 As New Point(cx + bH - lz, -hN + a - dZ)
            Dim P5 As New Point(cx + bH, -hN + a)
            Dim P6 As New Point(cx + bH, -hN)

            Dim rN2 As New Polygon
            rN2.points.Add(New Point(cx + bH + 0.5 * bN, 0))
            rN2.points.Add(P40)
            rN2.points.Add(P4)
            rN2.points.Add(P5)
            rN2.points.Add(P6)
            rN2.points.Add(New Point(cx + bH + 0.5 * bN, -hN))
            rN2.pen = penThin
            rN2.fill = fillWood1
            items.Add(rN2)

            ' --- Dimension lines (side view) ---

            ' Primary beam height hH
            Dim dl_hH As New DimensionLine()
            dl_hH.startPoint = New Point(cx, 0)
            dl_hH.endPoint = New Point(cx, -hH)
            dl_hH.offset = New size(7, Reference.contextMillimeters)
            dl_hH.textFormatString = "hH=0.0"
            dl_hH.textSize = fontSize
            items.Add(dl_hH)

            ' Primary beam width bH
            Dim dl_bH As New DimensionLine()
            dl_bH.startPoint = New Point(cx, -hH)
            dl_bH.endPoint = New Point(cx + bH, -hH)
            dl_bH.offset = New size(-7, Reference.contextMillimeters)
            dl_bH.textFormatString = "bH=0.0"
            dl_bH.textSize = fontSize
            items.Add(dl_bH)

            ' Dovetail length lZ
            Dim dl_lZ As New DimensionLine()
            dl_lZ.startPoint = P4
            dl_lZ.endPoint = P5
            dl_lZ.alignment = DimensionLine.DimAlignement.horizontal
            dl_lZ.offset = New size(-7, Reference.contextMillimeters)
            dl_lZ.textFormatString = "lZ=0.0"
            dl_lZ.textSize = fontSize
            items.Add(dl_lZ)

            ' Notch rest height ha (side view)
            Dim dl_ha2 As New DimensionLine()
            dl_ha2.startPoint = P5
            dl_ha2.endPoint = P6
            dl_ha2.alignment = DimensionLine.DimAlignement.vertical
            dl_ha2.offset = New size(-7, Reference.contextMillimeters)
            dl_ha2.textFormatString = "ha=0.0"
            dl_ha2.textSize = fontSize
            items.Add(dl_ha2)

            ' Taper depth dZ
            Dim dl_dz As New DimensionLine()
            dl_dz.startPoint = P5
            dl_dz.endPoint = P4
            dl_dz.alignment = DimensionLine.DimAlignement.vertical
            dl_dz.offset = New size(14, Reference.contextMillimeters)
            dl_dz.textFormatString = "0.0"
            dl_dz.textSize = fontSize
            items.Add(dl_dz)

            ' Outer height hZ (side view)
            Dim dl_hZ2 As New DimensionLine()
            dl_hZ2.startPoint = P40
            dl_hZ2.endPoint = P4
            dl_hZ2.alignment = DimensionLine.DimAlignement.vertical
            dl_hZ2.offset = New size(7, Reference.contextMillimeters)
            dl_hZ2.textFormatString = "0.0"
            dl_hZ2.textSize = fontSize
            items.Add(dl_hZ2)

            ' Taper angle beta
            Dim da_betaFW As New DimensionAngular(P5, P4, P5, P6)
            da_betaFW.dimSymbol = DimensionAngular.DimSymbols.Arrow
            da_betaFW.textSize = fontSize
            da_betaFW.textPrefix = ChrW(946) & "="
            da_betaFW.offset = New size(7, Reference.contextMillimeters)
            items.Add(da_betaFW)

            Return items
        End Function

        ''' <summary>Converts degrees to radians.</summary>
        Private Shared Function rad(angleDegrees As Double) As Double
            Return angleDegrees * Math.PI / 180
        End Function

    End Class

#End Region

End Namespace
