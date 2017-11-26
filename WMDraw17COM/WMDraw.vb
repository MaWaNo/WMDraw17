Imports System.Runtime.InteropServices
Imports WMDraw17.WallnerMild.Draw

<Guid("A3A1019E-474D-413F-A26F-0180C040C45B"),
        ClassInterface(ClassInterfaceType.AutoDual),
        ProgId("WMDraw17COM.WMDraw")>
Public Class WMDraw
    Dim p_wmd As New Drawing
    Public Sub New()

    End Sub
    Public Sub test2()
        MsgBox("test2")
    End Sub
    Public Sub test3()
        Dim p As New WMDraw17.WallnerMild.Draw.Drawing
        MsgBox(TypeName(p))
    End Sub
    Public Sub test1()

        MsgBox("A")
        p_wmd.setContext(Contexts.PNGClipboard, 500, 500)
        MsgBox("B")
        Dim l As New Line(10, 10, 100, 100)
        MsgBox("C")
        p_wmd.add(l)
        MsgBox("D")
        p_wmd.draw()
        MsgBox("E")

    End Sub
End Class
