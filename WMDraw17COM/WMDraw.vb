Imports System.Runtime.InteropServices
Imports WMDraw17.WallnerMild.Draw

<Guid("A3A1019E-474D-413F-A26F-0180C040C45B"),
        ClassInterface(ClassInterfaceType.AutoDual),
        ProgId("WMDraw17COM.WMDraw")>
Public Class WMDraw
    Dim p_wmd As New Drawing

    Private p_ref As References
    ''' <summary>
    ''' Mirror of WMDraw17.WallnerMild.Draw.Reference
    ''' </summary>
    Public Enum References
        world = 0
        contextMillimeters = 1
        contextUnits = 2
        contextFraction = 3
    End Enum
    ''' <summary>
    ''' Mirror of WMDraw17.WallnerMild.Draw.Reference
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

    Public Sub clear()
        p_wmd = New Drawing
    End Sub
    Public Sub addLine(startX As Double, startY As Double, endX As Double, endY As Double)
        'Dim l As New Line(startX, startY, endX, endY, WMDraw17.WallnerMild.Draw.Reference.world)
        Dim l As New Line(startX, startY, endX, endY, CurRef)
        p_wmd.add(l)
    End Sub

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
    Public Sub drawToClipboard(sizeX As Double, sizeY As Double, Optional unit As String = "px")
        p_wmd.setContext(Contexts.PNGClipboard, sizeX, sizeY, unit)
        Try
            p_wmd.draw()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    Public Sub test2()
        MsgBox("test2")
    End Sub
    Public Sub test3()
        Dim p As New WMDraw17.WallnerMild.Draw.Drawing
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
End Class
