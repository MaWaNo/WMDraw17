Imports System.Runtime.InteropServices

Imports System
Imports wpfImg = System.Windows.Media.Imaging
Imports System.Drawing
Imports System.IO
Imports System.Windows.Controls
Imports System.Windows.Media.Imaging

#Const allowOverloads = False

Namespace WallnerMild
    ''' <summary>
    ''' Drawing Class
    ''' </summary>
    ''' <remarks></remarks>
    ''' 
    <Guid("A75AB3C9-2ABA-40D4-BBA3-488CCD424252"),
        ClassInterface(ClassInterfaceType.AutoDual),
        ProgId("WMDraw17.XLDrawing")>
    Public Class DrawingOld

        Public Enum contexts
            ''' <summary>
            ''' Directly write to canvas
            ''' </summary>
            ''' <remarks></remarks>
            XAMLCanvas
            ''' <summary>
            ''' Write to PNG File
            ''' </summary>
            ''' <remarks></remarks>
            PNGFile
            ''' <summary>
            ''' Create an IN-Memory PNG
            ''' </summary>
            ''' <remarks></remarks>
            PNGBitmapInMemory
            PNG2BitmapInMemory
            ''' <summary>
            ''' Create an IN-Memory PNG for use as WPF-ImageSource
            ''' </summary>
            ''' <remarks></remarks>
            WPF_PNGImageSourceInMemory
        End Enum
        ''' <summary>
        ''' Output context
        ''' </summary>
        ''' <remarks></remarks>
        Public context As contexts

        ''' <summary>
        ''' Array List to model Com-Interop Array
        ''' </summary>
        ''' <remarks>
        ''' http://stackoverflow.com/questions/269581/what-are-alternatives-to-generic-collections-for-com-interop
        ''' </remarks>
        <Guid("A75AB3C9-2ABA-40D4-BBA3-488CCD424255"),
    ClassInterface(ClassInterfaceType.AutoDual),
    ProgId("WMDraw17.ComArrayList")>
        Public Class ComArrayList
            Inherits System.Collections.ArrayList
            Public Overridable Function GetByIndex(index As Integer) As Object
                Return MyBase.Item(index)
            End Function

            Public Overridable Sub SetByIndex(index As Integer, value As Object)
                MyBase.Item(index) = value
            End Sub

        End Class

        Public Enum lengthUnit
            model
            contextMillimeters
            contextUnits
        End Enum

        Public Class length
            Private _value As Double
            Private _unit As lengthUnit

            Property value As Double
                Get
                    Return _value
                End Get
                Set(value As Double)
                    _value = value
                End Set
            End Property

            Sub New()
                value = 0
                _unit = lengthUnit.model
            End Sub

            Friend Sub New(value As Double)
                value = Me.value
                _unit = lengthUnit.model
            End Sub
        End Class

        ''' <summary>
        ''' Lines-Array
        ''' </summary>
        Public lines As New ComArrayList

        ''' <summary>
        ''' Add a line
        ''' </summary>
        ''' <param name="newLine"></param>
        ''' <remarks></remarks>
        Public Sub addline(newLine As line)
            lines.Add(newLine)
        End Sub

        Public Class pen
            MustInherit Class pen
                ' should be derived from bas class

            End Class
        End Class

        Public Class drawingEntity
            Private _pen As pen

            Public Property pen As pen
                Get
                    Return _pen
                End Get
                Set(value As pen)
                    _pen = value
                End Set
            End Property

        End Class

        ''' <summary>
        ''' Point Entity
        ''' </summary>
        ''' <remarks></remarks>
        <Guid("A75AB3C9-2ABA-40D4-BBA3-488CCD424651"),
            ClassInterface(ClassInterfaceType.AutoDual),
            ProgId("WMDraw17.Point")> Public Class point
            Inherits drawingEntity

            Private _x As New length(0)
            Private _y As New length(0)

            Public Property x As length
                Get
                    Return _x
                End Get
                Set(value As length)
                    _x = value
                End Set
            End Property

            Public Property y As length
                Get
                    Return _y
                End Get
                Set(value As length)
                    _y = value
                End Set
            End Property

            Public Sub New()
                x = New length(0)
                y = New length(0)
            End Sub

            Friend Sub New(x As Double, y As Double)
                Me.x.value = x
                Me.y.value = y
            End Sub

        End Class

        ''' <summary>
        ''' Line Entity
        ''' </summary>
        ''' <remarks></remarks>
        <Guid("A75AB3C9-2ABA-40D4-BBA3-488CCD424251"),
            ClassInterface(ClassInterfaceType.AutoDual),
            ProgId("WMDraw17.Line")>
        Public Class line
            Inherits drawingEntity

            Private _P1 As New point
            Private _P2 As New point

            Public Property P1 As point
                Get
                    Return _P1
                End Get
                Set(value As point)
                    _P1 = value
                End Set
            End Property

            Public Property P2 As point
                Get
                    Return _P2
                End Get
                Set(value As point)
                    _P2 = value
                End Set
            End Property

            Public Sub New()
                P1 = New point
                P2 = New point
            End Sub

            Public Sub New(x1 As Double, y1 As Double, x2 As Double, y2 As Double)
                P1 = New point(x1, y1)
                P2 = New point(x2, y2)
            End Sub

            Public Sub New(P1 As point, P2 As point)
                Me.P1 = P1
                Me.P2 = P2
            End Sub
        End Class


        ''' <summary>
        ''' 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function draw() As Object
            Select Case context
                Case contexts.PNGFile, contexts.PNGBitmapInMemory, contexts.WPF_PNGImageSourceInMemory
                    Dim bmp As New Bitmap(300, 300, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
                    'bmp.SetResolution(600, 600)
                    Dim g As Graphics = Graphics.FromImage(bmp)

                    Using g
                        g.Clear(Color.Transparent)
                        g.DrawLine(Pens.Red, 0, 0, 135, 135)
                    End Using

                    Select Case context
                        Case contexts.PNGFile
                            Dim filename As String = "d:\temp\a.png"
                            bmp.Save(filename)
                            Return filename
                        Case contexts.PNGBitmapInMemory
                            Return bmp
                        Case contexts.WPF_PNGImageSourceInMemory
                            Return BitmapToImageSource(bmp)
                    End Select
                Case contexts.PNG2BitmapInMemory
                    Dim c As New Canvas
                    c.Height = 300
                    c.Width = 300
                    c.Background = New Windows.Media.SolidColorBrush(System.Windows.Media.Colors.Transparent)

                    ' if copied to clipboard transparent backgrund does not work
                    'c.Background = New Windows.Media.SolidColorBrush(System.Windows.Media.Colors.White)

                    Dim l As New System.Windows.Shapes.Line
                    Dim thickness As New Windows.Thickness(101, -11, 362, 250)
                    l.Margin = thickness
                    l.Visibility = System.Windows.Visibility.Visible
                    l.StrokeThickness = 4
                    l.Stroke = System.Windows.Media.Brushes.Black

                    l.X1 = 10
                    l.Y1 = 14
                    l.X2 = 100
                    l.Y2 = 150

                    c.Children.Add(l)

                    '
                    ' convert canvas to a renderBitmap
                    '

                    '
                    ' Save current canvas transform
                    '
                    Dim transform As Windows.Media.Transform = c.LayoutTransform
                    ' reset current transform (in case it is scaled or rotated)
                    c.LayoutTransform = Nothing

                    ' Get the size of canvas
                    Dim size As New System.Windows.Size(c.Width, c.Height)

                    ' Measure and arrange the surface
                    ' VERY IMPORTANT
                    c.Measure(size)
                    c.Arrange(New System.Windows.Rect(size))

                    '
                    ' Create a render bitmap (PNG) and push the canvas surface to it
                    '
                    Dim renderBitmap As New RenderTargetBitmap(CInt(size.Width), CInt(size.Height), 96.0, 96.0, Windows.Media.PixelFormats.Pbgra32)

                    '
                    ' canvas bounds
                    '
                    Dim bounds As System.Windows.Rect = Windows.Media.VisualTreeHelper.GetDescendantBounds(c)
                    '
                    ' drawing visual
                    '
                    Dim dv As New Windows.Media.DrawingVisual()
                    '
                    ' actually draw by using RenderOpen
                    ' (using actually closes the context at end using)
                    '
                    Using ctx As Windows.Media.DrawingContext = dv.RenderOpen()
                        '
                        ' create a visual brush from the canvas
                        '
                        Dim vb As New Windows.Media.VisualBrush(c)
                        '
                        ' draw it to the drawing context
                        '
                        ctx.DrawRectangle(vb, Nothing, New System.Windows.Rect(New System.Windows.Point(), bounds.Size))
                    End Using

                    '
                    ' Render the Bitmap
                    '
                    renderBitmap.Render(dv)


                    '
                    ' Write to png-file 
                    '
                    'Using outStream As New FileStream("d:\temp\b.png", FileMode.Create, FileAccess.Write, FileShare.None)
                    '    ' Use png encoder for our data
                    '    Dim encoder As New PngBitmapEncoder()
                    '    ' push the rendered bitmap to it
                    '    encoder.Frames.Add(BitmapFrame.Create(renderBitmap))
                    '    ' save the data to the stream
                    '    encoder.Save(outStream)
                    'End Using


                    ' Copy PNG to clipboard as BMP-Bitmap
                    ' Windows.Clipboard.SetImage(renderBitmap)

                    '
                    ' Copy PNG to clipboard as PNG-Bitmap
                    '
                    Using stream As New MemoryStream
                        Dim encoder As New PngBitmapEncoder()
                        ' push the rendered bitmap to it
                        encoder.Frames.Add(BitmapFrame.Create(renderBitmap))
                        ' save the data to the stream
                        encoder.Save(stream)
                        Dim data As New Windows.DataObject("PNG", stream)
                        Windows.Clipboard.Clear()
                        Windows.Clipboard.SetDataObject(data, True)
                    End Using

                    ' Restore canvas transformations as previously saved 
                    c.LayoutTransform = transform

            End Select

            Return Nothing

        End Function


        ''' <summary>
        ''' Convert a Bitmap to a WPF-ImageSource
        ''' </summary>
        ''' <param name="bitmap"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' source:
        ''' http://stackoverflow.com/questions/22499407/how-to-display-a-bitmap-in-a-wpf-image
        ''' </remarks>
        Private Function BitmapToImageSource(bitmap As Bitmap) As wpfImg.BitmapImage

            Using memory As New MemoryStream()
                bitmap.Save(memory, System.Drawing.Imaging.ImageFormat.Png)
                memory.Position = 0
                Dim bitmapimage As New wpfImg.BitmapImage()
                bitmapimage.BeginInit()
                bitmapimage.StreamSource = memory
                bitmapimage.CacheOption = wpfImg.BitmapCacheOption.OnLoad
                bitmapimage.EndInit()
                Return bitmapimage
            End Using

        End Function


        Public Sub New()
            lines = New ComArrayList
        End Sub
    End Class
End Namespace
