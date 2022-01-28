Imports System.Runtime.InteropServices

<Guid("62F8CB1D-9096-4808-B29A-5396C0FF1DEE"),
        ClassInterface(ClassInterfaceType.AutoDual),
        ProgId("WMComDraw2.c_WMComDraw2")>
Public Class testclass1
    Public Sub m(message As String)
        MsgBox(message)
    End Sub

    Public Function f() As String
        Return My.Application.Info.Version.ToString
    End Function

End Class
