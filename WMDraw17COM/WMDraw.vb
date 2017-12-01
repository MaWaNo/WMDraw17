Imports System.Runtime.InteropServices
Imports WMDraw17.WallnerMild.Draw

<Guid("A3A1019E-474D-413F-A26F-0180C040C45B"),
        ClassInterface(ClassInterfaceType.AutoDual),
        ProgId("WMDraw17COM.WMDraw")>
Public Class WMDraw
    Dim p_wmd As New Drawing

    Public Sub New()
        p_wmd = New Drawing
        p_wmd.ContextObject.fitHeight = True
        p_wmd.ContextObject.fitWidth = True
        p_wmd.ContextObject.fitProportional = True
    End Sub

    Public Sub clear()
        p_wmd = New Drawing
    End Sub
    Public Sub addLine(startX As Double, startY As Double, endX As Double, endY As Double)
        Dim l As New Line(startX, startY, endX, endY, Reference.world)
        p_wmd.add(l)
    End Sub

    Public Sub addText(startX As Double, startY As Double, Text As String)
        Dim t As New Text
        p_wmd.add(t)
    End Sub
    Public Sub drawToClipboard()
        p_wmd.setContext(Contexts.PNGClipboard, 500, 500)
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
