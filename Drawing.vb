Imports System.IO
Imports System.Windows.Media.Imaging

Namespace WallnerMild.Draw

    Public Enum Contexts
        PNGClipboard
        PNGFile
        DXFFile
        XAML
        WPFCanvas
    End Enum

    ''' <summary>
    ''' Wrong Context Exception
    ''' </summary>
    Public Class DrawingWrongContextObjectException
        Inherits System.ApplicationException

        Public Sub New(ByVal message As String)
            MyBase.New(message)
        End Sub

    End Class

    ''' <summary>
    ''' Drawing Class
    ''' World coordinates are x horizontal from left to right, y vertical from bottom to top
    ''' </summary>
    Public Class Drawing

        Private p_drawables As New List(Of Drawable)
        Private p_context As Contexts
        Private p_filename As String
        Private p_contextObject As New ContextObject
        Private p_boundingRectangle As Double()
        Private p_boundingRectangleUpdated As Boolean
        Private p_pen As pen

        Const DEFAULT_OUTPUT_DPI = 300 '96  ' dpi
        Const MM_PER_INCH = 25.4


        Sub New()
            flush()
            ' default context is xaml
            p_context = Contexts.WPFCanvas
            p_boundingRectangleUpdated = False
            p_pen = New pen
        End Sub
        ''' <summary>
        ''' flush drawables list
        ''' </summary>
        Sub flush()
            ' reset list
            p_drawables = New List(Of Drawable)
        End Sub
        ''' <summary>
        ''' Drawing context
        ''' </summary>
        ''' <returns></returns>
        Public Property Context As Contexts
            Get
                Return p_context
            End Get
            Set(value As Contexts)
                p_context = value
            End Set
        End Property

        Public Property FileName() As String
            Get
                Return p_filename
            End Get
            Set(value As String)
                p_filename = value
            End Set
        End Property

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="contextType"></param>
        ''' <param name="values">
        '''              | PNGClipboard |
        '''              +--------------+
        ''' fist value   | width (px)   | 
        ''' second value | height (px)  |
        ''' third value  | {mm,cm,px}
        '''              | (px is default)
        ''' 4th value    | FileName
        '''              +--------------+
        ''' </param>
        Public Sub setContext(contextType As Contexts, ParamArray values() As Object)

            Dim sc As Double = 1.0

            Me.p_context = contextType

            Select Case contextType
                Case Contexts.PNGClipboard
                    ' PNG in Clipboard
                    Dim c As New System.Windows.Controls.Canvas
                    If values.GetUpperBound(0) >= 2 Then
                        sc = scaleForUnitString(values(2))
                    End If
                    c.Width = values(0) * sc
                    c.Height = values(1) * sc
                    c.Background = New System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.Transparent)
                    Me.p_contextObject.Item = c
                Case Contexts.PNGFile
                    ' PNG File
                    Dim c As New System.Windows.Controls.Canvas
                    If values.GetUpperBound(0) >= 2 Then
                        sc = scaleForUnitString(values(2))
                    End If
                    c.Width = values(0) * sc
                    c.Height = values(1) * sc
                    c.Background = New System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.Transparent)
                    Me.p_contextObject.Item = c
                    If values.GetUpperBound(0) >= 3 Then
                        FileName = values(3)
                    Else
                        Dim fileName2 As String = ""
                        FileName = Path.GetTempFileName
                        fileName2 = FileName.Replace(Path.GetExtension(FileName), ".png")
                        File.Move(FileName, fileName2)
                        FileName = fileName2
                    End If
            End Select

        End Sub
        ''' <summary>
        ''' Scaling-factor for given unit String
        ''' 
        ''' </summary>
        ''' <param name="unitString">"px" (default), "mm", "cm"</param>
        ''' <returns></returns>
        Private Function scaleForUnitString(unitString As String) As Double
            Select Case LCase(unitString)
                Case "mm"
                    Return 96 / MM_PER_INCH
                Case "cm"
                    Return 96 / MM_PER_INCH / 10
                Case "px"
                    Return 1.0
                Case Else
                    Return 1.0
            End Select
        End Function
        Public Property pen() As pen
            Get
                Return p_pen
            End Get
            Set(value As pen)
                p_pen = value
            End Set
        End Property
        ''' <summary>
        ''' add a drawing item
        ''' </summary>
        ''' <param name="item"></param>
        Public Sub add(item As Drawable)
            If item.pen Is Nothing Then
                item.pen = Me.pen
            End If
            p_drawables.Add(item)
            p_boundingRectangleUpdated = False
        End Sub
        ''' <summary>
        ''' The Context Object to draw on
        ''' By setting the Object, the reference value is also set, if the Object Type could be found
        ''' </summary>
        ''' <returns></returns>
        Public Property ContextObject As ContextObject
            Get
                Return p_contextObject
            End Get
            Set(value As ContextObject)
                Dim contextTypeValid As Boolean = False
                '
                ' check contextItem type (valid for canvas)
                '
                If TypeOf value.Item Is System.Windows.Controls.Canvas Then
                    contextTypeValid = True
                    Me.Context = Contexts.WPFCanvas
                End If

                '
                '
                '

                If Not contextTypeValid Then
                    Throw New DrawingWrongContextObjectException("Als Objekt für den Context 'WPFCanvas' darf nur ein WPFCanvas verwendet werden.")
                    p_contextObject = Nothing
                    Exit Property
                End If

                p_contextObject = value
            End Set
        End Property

        Private Sub calculateBoundingRectangle()

            If p_boundingRectangleUpdated Then Exit Sub

            ReDim p_boundingRectangle(3)
            Dim tr(3) As Double
            Dim myItem As Drawable

            For Each myItem In p_drawables
                tr = myItem.boundingRectangle()
                For i = 0 To 3
                    If i Mod 2 = 0 Then
                        If p_boundingRectangle(0) > tr(i) Then p_boundingRectangle(0) = tr(i)
                        If p_boundingRectangle(2) < tr(i) Then p_boundingRectangle(2) = tr(i)
                    Else
                        If p_boundingRectangle(1) > tr(i) Then p_boundingRectangle(1) = tr(i)
                        If p_boundingRectangle(3) < tr(i) Then p_boundingRectangle(3) = tr(i)
                    End If
                Next
            Next

            p_boundingRectangleUpdated = True

        End Sub
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <returns>filename if WriteToFile
        ''' true if copy to clipboard</returns>
        Public Function draw() As Object

            calculateBoundingRectangle()

            For Each myItem In p_drawables
                ' transform koordinates, size and pen to context coordinates
                ' and draw on context by calling the draw routine
                myItem.draw(Me.ContextObject, AddressOf contextCoordinates, AddressOf contextSize)
            Next

            '
            ' copy to Clipboard or save to File
            '
            Select Case Me.Context
                Case Contexts.PNGClipboard, Contexts.PNGFile
                    '
                    ' Save current canvas transform
                    '
                    Dim transform As Windows.Media.Transform = Me.ContextObject.Item.LayoutTransform
                    ' reset current transform (in case it is scaled or rotated)
                    Me.ContextObject.Item.LayoutTransform = Nothing
                    '                    Me.ContextObject.Item.LayoutTransform = New Windows.Media.ScaleTransform(2, 2)

                    ' Get the size of canvas
                    ' as the output image will be bound to a given density in DPI,
                    ' the size is scaled in order to meet this density
                    '
                    Dim size As New System.Windows.Size(Me.ContextObject.Item.Width * DEFAULT_OUTPUT_DPI / 96,
                                                        Me.ContextObject.Item.Height * DEFAULT_OUTPUT_DPI / 96)

                    ' Measure and arrange the surface
                    ' VERY IMPORTANT
                    Me.ContextObject.Item.Measure(size)
                    Me.ContextObject.Item.Arrange(New System.Windows.Rect(size))

                    '
                    ' Create a render bitmap (PNG) and push the canvas surface to it
                    '
                    'Dim renderBitmap As New RenderTargetBitmap(CInt(size.Width), CInt(size.Height), 96.0, 96.0, Windows.Media.PixelFormats.Pbgra32)

                    '
                    ' Output
                    '
                    Dim renderBitmap As New RenderTargetBitmap(CInt(size.Width), CInt(size.Height), DEFAULT_OUTPUT_DPI, DEFAULT_OUTPUT_DPI, Windows.Media.PixelFormats.Pbgra32)

                    '
                    ' canvas bounds
                    '
                    Dim bounds As System.Windows.Rect = Windows.Media.VisualTreeHelper.GetDescendantBounds(Me.ContextObject.Item)
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
                        Dim vb As New Windows.Media.VisualBrush(Me.ContextObject.Item)
                        '
                        ' draw it to the drawing context
                        '
                        ctx.DrawRectangle(vb, Nothing, New System.Windows.Rect(New System.Windows.Point(), bounds.Size))
                    End Using

                    '
                    ' Render the Bitmap
                    '
                    renderBitmap.Render(dv)

                    Select Case Me.Context
                        Case Contexts.PNGFile
                            '
                            ' Write to png-file 
                            '
                            Using outStream As New FileStream(FileName, FileMode.Create, FileAccess.Write, FileShare.None)
                                ' Use png encoder for our data
                                Dim encoder As New PngBitmapEncoder()
                                ' push the rendered bitmap to it
                                encoder.Frames.Add(BitmapFrame.Create(renderBitmap))
                                ' save the data to the stream
                                encoder.Save(outStream)
                            End Using
                            draw = FileName
                        Case Contexts.PNGClipboard

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
                                ' copy to clipboard (PNG is understood by office applicatoins but not in Irfanview)
                                Dim data As New Windows.DataObject("PNG", stream)
                                Windows.Clipboard.Clear()
                                Windows.Clipboard.SetDataObject(data, True)
                            End Using
                            draw = True
                    End Select

                    ' Restore canvas transformations as previously saved 
                    Me.ContextObject.Item.LayoutTransform = transform
            End Select


        End Function

        ''' <summary>
        ''' Delegated function to calculate context-specific (transformed) size
        ''' </summary>
        ''' <param name="s">Size to be transformed</param>
        ''' <returns></returns>
        Function contextSize(ByVal s As size) As size
            '
            ' Transform point coordinates
            '
            Dim sTrans As New size

            Select Case s.Reference
                Case Reference.contextUnits
                    ' unchanged
                    sTrans.width = s.width * 96 / DEFAULT_OUTPUT_DPI
                    sTrans.height = s.height * 96 / DEFAULT_OUTPUT_DPI
                Case Reference.contextFraction
                    sTrans.width = s.width * Me.ContextObject.width '* 96 / DEFAULT_OUTPUT_DPI
                    sTrans.height = s.height * Me.ContextObject.height '* 96 / DEFAULT_OUTPUT_DPI
                Case Reference.contextMillimeters
                    sTrans.width = s.width / Me.ContextObject.widthMM * Me.ContextObject.width '* 96 / DEFAULT_OUTPUT_DPI
                    sTrans.height = s.height / Me.ContextObject.heightMM * Me.ContextObject.height '* 96 / DEFAULT_OUTPUT_DPI
                Case Reference.world

                    Dim scaleX As Double = 1
                    Dim scaleY As Double = 1

                    If Me.ContextObject.fitWidth Then
                        Dim worldWidth = p_boundingRectangle(2) - p_boundingRectangle(0)
                        If worldWidth = 0 Then
                            scaleX = 1
                        Else
                            scaleX = Me.ContextObject.width / worldWidth
                        End If
                    End If

                    If Me.ContextObject.fitHeight Then
                        Dim worldHeight = p_boundingRectangle(3) - p_boundingRectangle(1)
                        If worldHeight = 0 Then
                            scaleY = 1
                        Else
                            scaleY = Me.ContextObject.height / worldHeight
                        End If
                    End If

                    If Me.ContextObject.fitProportional Then
                        If scaleX > scaleY Then
                            scaleX = scaleY
                        Else
                            scaleY = scaleX
                        End If
                    End If

                    sTrans.width = s.width * scaleX '* 96 / DEFAULT_OUTPUT_DPI
                    sTrans.height = s.height * scaleY '* 96 / DEFAULT_OUTPUT_DPI
            End Select

            sTrans.Reference = Reference.contextUnits

            Return sTrans
        End Function

        ''' <summary>
        ''' Delegated function to calculate context-specific (transformed) coordinates
        ''' </summary>
        ''' <param name="p">Point to be transformed</param>
        ''' <returns></returns>
        Function contextCoordinates(ByVal p As Point) As Point
            '
            ' Transform point coordinates
            '
            Dim pTrans As New Point
            Select Case p.coordinateReference
                Case Reference.contextUnits
                    ' unchanged
                    pTrans.x = p.x * 96 / DEFAULT_OUTPUT_DPI
                    pTrans.y = Me.ContextObject.height - p.y * 96 / DEFAULT_OUTPUT_DPI
                Case Reference.contextFraction
                    pTrans.x = p.x * Me.ContextObject.width '* 96 / DEFAULT_OUTPUT_DPI
                    pTrans.y = Me.ContextObject.height * (1 - p.y)'* 96 / DEFAULT_OUTPUT_DPI
                Case Reference.contextMillimeters
                    pTrans.x = p.x / Me.ContextObject.widthMM * Me.ContextObject.width '* 96 / DEFAULT_OUTPUT_DPI
                    pTrans.y = Me.ContextObject.height * (1 - p.y / Me.ContextObject.heightMM) '* 96 / DEFAULT_OUTPUT_DPI
                Case Reference.world

                    Dim scaleX As Double = 1
                    Dim scaleY As Double = 1

                    If Me.ContextObject.fitWidth Then
                        Dim worldWidth = p_boundingRectangle(2) - p_boundingRectangle(0)
                        If worldWidth = 0 Then
                            scaleX = 1
                        Else
                            scaleX = Me.ContextObject.width / worldWidth
                        End If
                    End If

                    If Me.ContextObject.fitHeight Then
                        Dim worldHeight = p_boundingRectangle(3) - p_boundingRectangle(1)
                        If worldHeight = 0 Then
                            scaleY = 1
                        Else
                            scaleY = Me.ContextObject.height / worldHeight
                        End If
                    End If

                    If Me.ContextObject.fitProportional Then
                        If scaleX > scaleY Then
                            scaleX = scaleY
                        Else
                            scaleY = scaleX
                        End If
                    End If

                    pTrans.x = p.x * scaleX '* 96 / DEFAULT_OUTPUT_DPI
                    pTrans.y = Me.ContextObject.height - p.y * scaleY '* 96 / DEFAULT_OUTPUT_DPI

            End Select
            pTrans.coordinateReference = Reference.contextUnits
            Return pTrans
        End Function

    End Class
    ''' <summary>
    ''' Possible orientations in horizontal direction
    ''' </summary>
    Public Enum orientationHorizontal
        leftToRight
        rightToLeft
    End Enum

    Public Enum orientationVertical
        bottomToTop
        topToBottom
    End Enum

    Public Class ContextObject
        Const DEFAULT_DPI = 96 '96  ' dpi
        Public Const MM_PER_INCH = 25.4

        Private p_item As Object

        Private p_width As Double
        Private p_height As Double

        Private p_fitHeight As Boolean
        Private p_fitWidth As Boolean
        Private p_fitProportional As Boolean

        Private p_dpiX As Double
        Private p_dpiY As Double


        Public ReadOnly Property width As Double
            Get
                Return p_width
            End Get
        End Property

        Public ReadOnly Property height As Double
            Get
                Return p_height
            End Get
        End Property

        Public ReadOnly Property heightMM As Double
            Get
                Return height * MM_PER_INCH / p_dpiY
            End Get
        End Property

        Public ReadOnly Property widthMM As Double
            Get
                Return width * MM_PER_INCH / p_dpiX
            End Get
        End Property

        Public Property Item As Object
            Get
                Return p_item
                Return Nothing
            End Get
            Set(value As Object)
                p_item = value

                If TypeOf Me.Item Is System.Windows.Controls.Canvas Then

                    '
                    ' cast object to canvas
                    '
                    Dim myCanvas As System.Windows.Controls.Canvas
                    myCanvas = TryCast(Me.Item, System.Windows.Controls.Canvas)

                    ' this would be to obtain the resolution
                    'Dim source As Windows.PresentationSource = Windows.PresentationSource.FromVisual(myCanvas)

                    'If source IsNot Nothing Then
                    p_dpiX = DEFAULT_DPI '* source.CompositionTarget.TransformToDevice.M11
                    p_dpiY = DEFAULT_DPI '* source.CompositionTarget.TransformToDevice.M22
                    'End If

                    p_width = myCanvas.Width
                    p_height = myCanvas.Height

                End If
            End Set
        End Property

        Public ReadOnly Property dpiX As Double
            Get
                Return p_dpiX
            End Get
        End Property

        Public ReadOnly Property dpiY As Double
            Get
                Return p_dpiY
            End Get
        End Property

        Public Property fitWidth As Boolean
            Get
                Return p_fitWidth
            End Get
            Set(value As Boolean)
                p_fitWidth = value
            End Set
        End Property

        Public Property fitProportional As Boolean
            Get
                Return p_fitProportional
            End Get
            Set(value As Boolean)
                p_fitProportional = value
            End Set
        End Property
        Public Property fitHeight As Boolean
            Get
                Return p_fitHeight
            End Get
            Set(value As Boolean)
                p_fitHeight = value
            End Set
        End Property
    End Class
End Namespace
