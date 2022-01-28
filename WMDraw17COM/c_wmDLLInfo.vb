Imports System.Runtime.InteropServices
Imports System.IO

''' <summary>
''' Standard DLL-Information
''' </summary>
''' <remarks></remarks>
<Guid("CFCCDC8F-7741-4F05-AE62-56A1FB6D5ACF"),
ClassInterface(ClassInterfaceType.AutoDual),
ProgId("WMComDraw.c_wmDLLInfo")>
Public Class c_wmDLLInfo

    ''' <summary> 
    ''' The function determines whether the current process is a  
    ''' 64-bit process. 
    ''' </summary> 
    ''' <returns> 
    ''' The function returns true if the process is 64-bit;  
    ''' otherwise, it returns false. 
    ''' </returns> 
    Public Function is64bitProcess() As Boolean
        Return IntPtr.Size = 8
    End Function

    ''' <summary>
    ''' Name
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property name As String
        Get
            name = My.Application.Info.AssemblyName & ".dll"
        End Get
    End Property

    ''' <summary>
    ''' Path
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property path() As String
        Get
            path = My.Application.Info.DirectoryPath
        End Get
    End Property

    ''' <summary>
    ''' Version
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property version As String
        Get
            With My.Application.Info.Version
                version = .Major & "." & .Minor & "." & .Revision & " - .NET " & IIf(is64bitProcess, "64", "32") & "bit"
            End With
        End Get
    End Property

    ''' <summary>
    ''' Date
    ''' </summary>
    ''' <returns>DLL creatondate</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property d() As Date
        Get
            Dim information As FileInfo = My.Computer.FileSystem.GetFileInfo(Me.path & System.IO.Path.DirectorySeparatorChar & Me.name)
            d = information.CreationTime()
        End Get
    End Property
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>new since .Net-Version</remarks>
    Public ReadOnly Property remarks() As String
        Get
            Return (IIf(is64bitProcess, "64-bit", "32-bit"))
        End Get
    End Property
End Class

