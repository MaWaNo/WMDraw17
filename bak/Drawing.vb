Imports System.IO
Imports System.Runtime.Serialization
Imports System.Windows.Media.Imaging

#Const imgtype = "PNG"  ' other formats do not work

Namespace WMDraw
    ''' <summary>
    ''' Drawing Contexts
    ''' </summary>
    Public Enum Contexts
        PNGClipboard
        PNGFile
        DXFFile
        XAML
        WPFCanvas
    End Enum

    ''' <summary>
    ''' Geometric references for coordinates or sizes
    ''' </summary>
    Public Enum Reference
        ''' <summary>
        ''' world coordinates
        ''' </summary>
        world
        ''' <summary>
        ''' coordinates in mm
        ''' </summary>
        contextMillimeters
        ''' <summary>
        ''' context specific units (usually pixels)
        ''' </summary>
        contextUnits
        ''' <summary>
        ''' Fraction of the drawing context
        ''' </summary>
        contextFraction
    End Enum

    ''' <summary>
    ''' Wrong Context Exception
    ''' </summary>

    Public Class DrawingWrongContextObjectException
        Inherits System.ApplicationException


        Public Sub New(ByVal message As String)
            MyBase.New(message)
        End Sub

        Private Sub DrawingWrongContextObjectException_SerializeObjectState(sender As Object, e As SafeSerializationEventArgs) Handles Me.SerializeObjectState

        End Sub
    End Class

    ''' <summary>
    ''' Drawing Class
    ''' World coordinates are x horizontal from left to right, y vertical from bottom to top
    ''' </summary>
    Public Class Drawing

        Private p_drawables As New List(Of Drawable)        ' list of drawables 
        Private p_context As Contexts                       ' context type
        Private p_filename As String                        ' filename
        Private p_contextObject As New ContextObject        ' context object
        Private p_boundingRectangle As Double()             ' (0) min x; (1) min y; (2) max x; (3) max y
        Private p_boundingRectangle_outdated As Boolean     ' true if boundin rectangle is outdated
        Private p_pen As pen                                ' current drawing pen

        Const DEFAULT_OUTPUT_DPI = 300 '96  ' dpi           ' default output resolution
        Const MM_PER_INCH = 25.4                            ' unit conversion

        ''' <summary>
        ''' New drawing
        ''' </summary>
        Sub New()
            flush()
            ' default context is xaml
            p_context = Contexts.WPFCanvas
            p_boundingRectangle_outdated = True
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
        ''' Drawing context type
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
        ''' <summary>
        ''' Filename for contexts, that write to file
        ''' </summary>
        ''' <returns></returns>
        Public Property FileName() As String
            Get
                Return p_filename
            End Get
            Set(value As String)
                p_filename = value
            End Set
        End Property

        ''' <summary>
        ''' Set Context with data
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
                        Try
                            File.Move(FileName, fileName2)
                        Catch ex As Exception
                        End Try
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
        ''' <summary>
        ''' Pen
        ''' </summary>
        ''' <returns></returns>
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
            p_boundingRectangle_outdated = True
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
        ''' <summary>
        ''' Calculate the overall bounding rectangle
        ''' </summary>
        ''' <remarks>
        ''' data is stored in p_boundingRectangle
        ''' calculation is performed only, if p_boundingRectangle_outdated is true
        ''' </remarks>
        Private Sub calculateBoundingRectangle()

            If Not p_boundingRectangle_outdated Then Exit Sub

            ReDim p_boundingRectangle(3)
            Dim tr(3) As Double
            Dim myItem As Drawable
            Dim isFirst As Boolean = True

            For Each myItem In p_drawables
                tr = myItem.boundingRectangle()

                If isFirst Then
                    p_boundingRectangle(0) = tr(0)
                    p_boundingRectangle(1) = tr(1)
                    p_boundingRectangle(2) = tr(0)
                    p_boundingRectangle(3) = tr(1)

                    isFirst = False
                End If

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

            p_boundingRectangle_outdated = False

        End Sub
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <returns>filename if WriteToFile
        ''' true if copy to clipboard</returns>
        Public Function draw() As Object

            draw = Nothing

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
#If imgtype = "PNG" Then
                    'PNG
                    Dim renderBitmap As New RenderTargetBitmap(CInt(size.Width), CInt(size.Height), DEFAULT_OUTPUT_DPI, DEFAULT_OUTPUT_DPI, Windows.Media.PixelFormats.Pbgra32)
#ElseIf imgtype = "JPG" Then
                    ' JPG
                    Dim renderBitmap As New RenderTargetBitmap(CInt(size.Width), CInt(size.Height), DEFAULT_OUTPUT_DPI, DEFAULT_OUTPUT_DPI, Windows.Media.PixelFormats.Indexed1)
#Else
                    ' PNG
                    Dim renderBitmap As New RenderTargetBitmap(CInt(size.Width), CInt(size.Height), DEFAULT_OUTPUT_DPI, DEFAULT_OUTPUT_DPI, Windows.Media.PixelFormats.Pbgra32)
#End If
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
#If imgtype = "PNG" Then
                                Dim encoder As New PngBitmapEncoder()
#ElseIf imgtype = "JPG" Then
                                Dim encoder As New JpegBitmapEncoder()
#Else
                                Dim encoder As New PngBitmapEncoder()
#End If


                                'Dim encoder As New JpegBitmapEncoder()
                                ' push the rendered bitmap to it
                                encoder.Frames.Add(BitmapFrame.Create(renderBitmap))
                                ' save the data to the stream
                                encoder.Save(stream)
                                ' copy to clipboard (PNG is understood by office applicatoins but not in Irfanview)
                                'Dim data As New System.Windows.DataObject("PNG", stream)


#If imgtype = "PNG" Then
                                Dim data As New System.Windows.DataObject("PNG", stream)
#ElseIf imgtype = "JPG" Then
                                Dim data As New System.Windows.DataObject("DeviceIndependentBitmap", stream)
#Else
                                Dim data As New System.Windows.DataObject("PNG", stream)
#End If

                                System.Windows.Clipboard.Clear()
                                System.Windows.Clipboard.SetDataObject(data, True)

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
        '''  Margin in Context reference
        ''' </summary>
        Private ReadOnly Property contextMargin() As Margin
            Get
                Dim m As New Margin
                m = Me.ContextObject.Margin
                Dim mTrans As New Margin

                Select Case m.Reference
                    Case Reference.contextUnits
                        mTrans.left = m.left * 96 / DEFAULT_OUTPUT_DPI
                        mTrans.top = m.top * 96 / DEFAULT_OUTPUT_DPI
                        mTrans.right = m.right * 96 / DEFAULT_OUTPUT_DPI
                        mTrans.bottom = m.bottom * 96 / DEFAULT_OUTPUT_DPI
                    Case Reference.contextFraction
                        mTrans.left = m.left * Me.ContextObject.width
                        mTrans.top = m.top * Me.ContextObject.height
                        mTrans.right = m.right * Me.ContextObject.width
                        mTrans.bottom = m.bottom * Me.ContextObject.height
                    Case Reference.contextMillimeters
                        mTrans.left = m.left * Me.ContextObject.width / Me.ContextObject.widthMM
                        mTrans.top = m.top * Me.ContextObject.height / Me.ContextObject.heightMM
                        mTrans.right = m.right * Me.ContextObject.width / Me.ContextObject.widthMM
                        mTrans.bottom = m.bottom * Me.ContextObject.height / Me.ContextObject.heightMM
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

                        mTrans.left = m.left * scaleX
                        mTrans.top = m.top * scaleY
                        mTrans.right = m.right * scaleX
                        mTrans.bottom = m.bottom * scaleY

                        '
                        ' translate world coordinates
                        '
                        'pTrans.x = pTrans.x - p_boundingRectangle(0) * scaleX
                        'pTrans.y = pTrans.y + p_boundingRectangle(1) * scaleY

                End Select

                mTrans.Reference = Reference.contextUnits

                Return mTrans

            End Get
        End Property

        ''' <summary>
        ''' Delegated function to calculate context-specific (transformed) coordinates
        ''' </summary>
        ''' <param name="p">Point to be transformed</param>
        ''' <returns></returns>
        Function contextCoordinates(ByVal p As Point) As Point
            '
            ' Transform point coordinates
            '

            ' todo: seems to be correct, that no translation is done for context coordinates, fractions and mm

            Dim pTrans As New Point


            Select Case p.coordinateReference
                Case Reference.contextUnits
                    ' unchanged
                    pTrans.x = Me.contextMargin.left + p.x * 96 / DEFAULT_OUTPUT_DPI
                    pTrans.y = Me.ContextObject.height - Me.contextMargin.bottom - p.y * 96 / DEFAULT_OUTPUT_DPI

                    '
                    ' translate in contextUnits
                    '
                    'If (p_boundingRectangle(2) - p_boundingRectangle(0)) <> 0 Then
                    '    pTrans.x = pTrans.x - p_boundingRectangle(0) * Me.ContextObject.width / (p_boundingRectangle(2) - p_boundingRectangle(0))
                    'Else
                    '    pTrans.x = pTrans.x - p_boundingRectangle(0)
                    'End If
                    'If (p_boundingRectangle(3) - p_boundingRectangle(1)) <> 0 Then
                    '    pTrans.y = pTrans.y + p_boundingRectangle(1) * Me.ContextObject.height / (p_boundingRectangle(3) - p_boundingRectangle(1))
                    'Else
                    '    pTrans.y = pTrans.y + p_boundingRectangle(1)
                    'End If
                Case Reference.contextFraction

                    pTrans.x = Me.contextMargin.left + p.x * Me.ContextObject.width
                    pTrans.y = Me.ContextObject.height - Me.contextMargin.bottom -
                                (Me.ContextObject.height - Me.contextMargin.bottom - Me.contextMargin.top) * p.y
                Case Reference.contextMillimeters
                    pTrans.x = Me.contextMargin.left +
                                p.x / Me.ContextObject.widthMM * (Me.ContextObject.width - Me.contextMargin.left - Me.contextMargin.right)
                    pTrans.y = (Me.ContextObject.height - Me.contextMargin.bottom) -
                                (Me.ContextObject.height - Me.contextMargin.bottom - Me.contextMargin.top) * p.y / Me.ContextObject.heightMM
                Case Reference.world

                    Dim scaleX As Double = 1
                    Dim scaleY As Double = 1

                    If Me.ContextObject.fitWidth Then
                        Dim worldWidth = p_boundingRectangle(2) - p_boundingRectangle(0)
                        If worldWidth = 0 Then
                            scaleX = 1
                        Else
                            scaleX = (Me.ContextObject.width - Me.contextMargin.left - Me.contextMargin.right) / worldWidth
                        End If
                    End If

                    If Me.ContextObject.fitHeight Then
                        Dim worldHeight = p_boundingRectangle(3) - p_boundingRectangle(1)
                        If worldHeight = 0 Then
                            scaleY = 1
                        Else
                            scaleY = (Me.ContextObject.height - Me.contextMargin.bottom - Me.contextMargin.top) / worldHeight
                        End If
                    End If

                    If Me.ContextObject.fitProportional Then
                        If scaleX > scaleY Then
                            scaleX = scaleY
                        Else
                            scaleY = scaleX
                        End If
                    End If

                    pTrans.x = Me.contextMargin.left + p.x * scaleX '* 96 / DEFAULT_OUTPUT_DPI
                    pTrans.y = Me.ContextObject.height - Me.contextMargin.bottom - p.y * scaleY '* 96 / DEFAULT_OUTPUT_DPI

                    '
                    ' translate world coordinates (?)
                    '
                    pTrans.x = pTrans.x - p_boundingRectangle(0) * scaleX
                    pTrans.y = pTrans.y + p_boundingRectangle(1) * scaleY

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

    Public Class Margin
        Private p_left As Double
        Private p_top As Double
        Private p_right As Double
        Private p_bottom As Double
        Private p_all As Double

        Private p_Reference As Reference

        Public Sub New()
            p_left = 0
            p_top = 0
            p_right = 0
            p_bottom = 0
            p_all = -1
            p_Reference = Reference.contextUnits
        End Sub

        Public Sub New(all As Double)
            Me.all = all
        End Sub

        Public Sub New(left, top, right, bottom)
            Me.left = left
            Me.top = top
            Me.right = right
            Me.bottom = bottom
        End Sub

        Public Overrides Function ToString() As String
            Return String.Format("({0};{1};{2};{3})", left, top, right, bottom)
        End Function

        Public Property all() As Double
            Get
                Return p_all
            End Get
            Set(value As Double)
                p_all = value
                p_left = value
                p_top = value
                p_right = value
                p_bottom = value
            End Set
        End Property

        Public Property left() As Double
            Get
                Return IIf(p_all < 0, p_left, p_all)
            End Get
            Set(value As Double)
                p_all = -1
                p_left = value
            End Set
        End Property

        Public Property top() As Double
            Get
                Return IIf(p_all < 0, p_top, p_all)
            End Get
            Set(value As Double)
                p_all = -1
                p_top = value
            End Set
        End Property

        Public Property right() As Double
            Get
                Return IIf(p_all < 0, p_right, p_all)
            End Get
            Set(value As Double)
                p_all = -1
                p_right = value
            End Set
        End Property

        Public Property bottom() As Double
            Get
                Return IIf(p_all < 0, p_bottom, p_all)
            End Get
            Set(value As Double)
                p_all = -1
                p_bottom = value
            End Set
        End Property

        Public Property Reference() As Reference
            Get
                Return p_Reference
            End Get
            Set(value As Reference)
                p_Reference = value
            End Set
        End Property
    End Class

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

        Private p_margin As Margin

        Public Sub New()
            p_margin = New Margin
        End Sub

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
        ''' <summary>
        ''' Margin
        ''' </summary>
        ''' <returns></returns>
        Public Property Margin As Margin
            Get
                Return p_margin
            End Get
            Set(value As Margin)
                p_margin = value
            End Set
        End Property
    End Class
End Namespace
