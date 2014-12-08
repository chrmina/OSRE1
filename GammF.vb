Option Explicit On 
Option Strict On

Imports System.Math

Module MathGammaF
    Private Const ITMAX As Integer = 100
    Private Const EPS As Double = 0.00000000000000022204460492503131#
    Private Const FPMIN As Double = 2.2250738585072014E-308 / EPS

    Public Function GammLn(ByVal XX As Double) As Double
        Static cof() As Double = {76.180091729471457#, -86.505320329416776#, _
                                  24.014098240830911#, -1.231739572450155#, _
                                  0.001208650973866179#, -0.000005395239384953#}
        If XX <= 0.0# Then Throw New System.ArithmeticException("Invalid Argument in GammLn")

        Dim x As Double = XX
        Dim y As Double = XX

        Dim tmp As Double = x + 5.5#
        tmp -= (x + 0.5#) * Log(tmp)

        Dim ser As Double = 1.0000000001900149#

        For j As Integer = 0 To 5
            y += 1.0#
            ser += cof(j) / y
        Next

        Return -tmp + Log(2.5066282746310007# * ser / x)
    End Function

    Public Function GammP(ByVal A As Double, ByVal X As Double) As Double
        If X < 0.0# Or A <= 0.0# Then Throw New System.ArithmeticException("Invalid Argument in GammP")

        If X < A + 1.0# Then
            Dim GamSer, Gln As Double
            GSer(GamSer, A, X, Gln)
            Return GamSer
        Else
            Dim GamMcf, Gln As Double
            Gcf(GamMcf, A, X, Gln)
            Return 1.0# - GamMcf
        End If
    End Function

    Public Function GammQ(ByVal A As Double, ByVal X As Double) As Double
        If X < 0.0# Or A <= 0.0# Then Throw New System.ArithmeticException("Invalid Argument in GammQ")

        If X < A + 1.0# Then
            Dim GamSer, Gln As Double
            GSer(GamSer, A, X, Gln)
            Return 1.0# - GamSer
        Else
            Dim GamMcf, Gln As Double
            Gcf(GamMcf, A, X, Gln)
            Return GamMcf
        End If
    End Function

    Private Sub GSer(ByRef GamSer As Double, ByVal A As Double, _
                     ByVal X As Double, ByRef Gln As Double)
        Gln = GammLn(A)

        If X <= 0.0# Then
            If X < 0.0# Then Throw New System.ArithmeticException("Invalid Argument in GSer")
            GamSer = 0.0#
            Exit Sub
        Else
            Dim ap As Double = A
            Dim del As Double = 1.0# / A
            Dim sum As Double = del

            For n As Integer = 1 To ITMAX
                ap += 1.0#
                del *= X / ap
                sum += del
                If Abs(del) < Abs(sum) * EPS Then
                    GamSer = sum * Exp(-X + A * Log(X) - Gln)
                    Exit Sub
                End If
            Next n

            Throw New System.ArithmeticException("Bad A and ITMAX in GSer")
            Exit Sub
        End If
    End Sub

    Private Sub Gcf(ByRef GamMcf As Double, ByVal A As Double, _
                    ByVal X As Double, ByRef Gln As Double)
        Gln = GammLn(A)

        Dim b As Double = X + 1.0# - A
        Dim c As Double = 1.0# / FPMIN
        Dim d As Double = 1.0# / b
        Dim h As Double = d

        Dim i As Integer

        For i = 1 To ITMAX
            Dim an As Double = -i * (i - A)
            b += 2.0#
            d = an * d + b
            If Abs(d) < FPMIN Then d = FPMIN
            c = b + an / c
            If Abs(c) < FPMIN Then c = FPMIN
            d = 1.0# / d
            Dim del As Double = d * c
            h *= del
            If Abs(del - 1.0#) <= EPS Then Exit For
        Next

        If i > ITMAX Then Throw New System.ArithmeticException("Bad A and ITMAX in Gcf")
        GamMcf = Exp(-X + A * Log(X) - Gln) * h
    End Sub

    Public Function ErrF(ByVal X As Double) As Double
        If X < 0.0# Then
            Return -GammP(0.5, X * X)
        Else
            Return GammP(0.5, X * X)
        End If

    End Function
End Module
