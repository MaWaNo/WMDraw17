'Namespace WallnerMild

Module WMHelpers

    ''' <summary>
    ''' Set Significant Figures
    ''' </summary>
    ''' <param name="d">Double Value</param>
    ''' <param name="digits">Number of significant figures</param>
    ''' <returns>Rounded Value</returns>
    Public Function SetSigFigs(ByVal d As Double, ByVal digits As Integer) As Double
        Dim isNegative As Boolean
        If Double.IsNaN(d) Then Return 0
        If Double.IsInfinity(d) Then Return d

        If d < 0.0000000001 And d > -0.0000000001 Then Return 0

        If d < 0 Then
            isNegative = True
            d = Math.Abs(d)
        End If

        Dim scale As Decimal = CDec(Math.Pow(10, Math.Floor(Math.Log10(Math.Abs(d))) + 1))
        SetSigFigs = CDbl((scale * Math.Round(CDec(d) / scale, digits)))

        If isNegative Then
            SetSigFigs *= -1
        End If

        If Len(String.Format("{0:F" & CStr(digits) & "}", SetSigFigs)) > Len(CStr(SetSigFigs)) Then

        End If
    End Function

    Public Function SetSigFigsString(ByVal d As Double, ByVal digits As Integer) As String
        Dim isNegative As Boolean
        Dim isInteger As Boolean
        Dim numDigitsBeforeComma As Integer

        Dim rounded As Double

        If Double.IsNaN(d) Then Return 0
        If Double.IsInfinity(d) Then Return d

        If d < 10 ^ (-digits - 1) And d > -1 * 10 ^ (-digits - 1) Then Return 0

        If d < 0 Then
            isNegative = True
            d = Math.Abs(d)
        End If

        Dim scale As Decimal = CDec(Math.Pow(10, Math.Floor(Math.Log10(Math.Abs(d))) + 1))
        rounded = CDbl((scale * Math.Round(CDec(d) / scale, digits)))

        If isNegative Then
            rounded *= -1
        End If
        isInteger = (rounded = Math.Truncate(rounded))
        numDigitsBeforeComma = Len(CStr(Math.Truncate(rounded)))

        If isInteger Then
            If Len(CStr(rounded)) < digits Then
                Return String.Format("{0:F" & CStr(digits - Len(CStr(rounded))) & "}", rounded)
            Else
                Return CStr(rounded)
            End If
        Else
            If numDigitsBeforeComma < digits Then
                Return String.Format("{0:F" & CStr(digits - numDigitsBeforeComma) & "}", rounded)
            Else
                Return CStr(rounded)
            End If
        End If
    End Function

    ''' <summary>
    ''' Create Rainbow color
    ''' </summary>
    ''' <param name="xi"></param>
    ''' <returns></returns>
    Public Function rainbowColor(xi As Double) As Windows.Media.Color
            If xi > 1 Then xi = xi - CInt(xi)
            Return HSLtoRGB(240 * xi, 0.5, 0.5)
        End Function

        ''' <summary> 
        ''' Converts HSL to a .net Color.
        ''' </summary>
        ''' <param name="h">Hue value (must be between 0 and 360).</param>
        ''' <param name="s">Saturation value (must be between 0 and 1).</param>
        ''' <param name="l">Luminance value (must be between 0 and 1).</param>
        Public Function HSLtoRGB(ByVal h As Double, ByVal s As Double, ByVal l As Double) As Windows.Media.Color

            If (s = 0) Then

                ' achromatic color (gray scale)
                Return Windows.Media.Color.FromArgb(255,
                                CInt(Double.Parse(String.Format("{0:0.00}", l * 255.0))),
                                CInt(Double.Parse(String.Format("{0:0.00}", l * 255.0))),
                                CInt(Double.Parse(String.Format("{0:0.00}", l * 255.0))))
            Else

                Dim q As Double
                If l < 0.5 Then
                    q = l * (1.0 + s)
                Else
                    q = l + s - (l * s)
                End If

                Dim p As Double = (2.0 * l) - q
                ' modified to fit FEM-Style Scale
                If h > 240 Then h = 240
                Dim Hk As Double = (240 - h) / 360.0
                ' originally 
                ' Dim Hk As Double = h / 360

                Dim T(2) As Double
                T(0) = Hk + (1.0 / 3.0) ' Tr
                T(1) = Hk               ' Tb
                T(2) = Hk - (1.0 / 3.0) ' Tg

                For i As Integer = 0 To 2
                    If (T(i) < 0) Then T(i) += 1.0
                    If (T(i) > 1) Then T(i) -= 1.0

                    If ((T(i) * 6) < 1) Then
                        T(i) = p + ((q - p) * 6.0 * T(i))
                    ElseIf ((T(i) * 2.0) < 1) Then
                        T(i) = q
                    ElseIf ((T(i) * 3.0) < 2) Then
                        T(i) = p + (q - p) * ((2.0 / 3.0) - T(i)) * 6.0
                    Else
                        T(i) = p
                    End If
                Next

                'Return New RGB(cint(Math.Ceiling(T(0) * 255.0)), _
                '               cint(Math.Ceiling(T(1) * 255.0)), _
                '               cint(Math.Ceiling(T(2) * 255.0)))

                Return Windows.Media.Color.FromArgb(255, CInt(Double.Parse(String.Format("{0:0.00}", T(0) * 255.0))),
                               CInt(Double.Parse(String.Format("{0:0.00}", T(1) * 255.0))),
                               CInt(Double.Parse(String.Format("{0:0.00}", T(2) * 255.0))))

            End If

        End Function

    End Module

'End Namespace
