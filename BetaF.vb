Option Explicit On 
Option Strict On

Imports System.Math

Module MathBetaF
    'Example Usage of BetI Function
    '
    'Student's distribution A(t|Nu)
    '
    '   A(t|Nu) = 1# - BetaI(Nu / 2#, 0.5#, Nu / (Nu + t * t))
    '
    'F-Distribution Function Q(F|Nu1,Nu2)
    '
    '   Q(F|Nu1,Nu2) = BetaI(Nu2 / 2#, Nu1 / 2#, Nu2 / (Nu2 + Nu1 * F))

    Public Function BetaI(ByVal A As Double, ByVal B As Double, ByVal X As Double) As Double
        Dim Bt As Double

        If X < 0.0# Or X > 1.0# Then _
            Throw New System.ArithmeticException("Invalid Argument X in BetaI")
        If A <= 0.0# Or B <= 0.0# Then _
            Throw New System.ArithmeticException("Invalid Argument A or B in BetaI")

        If X = 0.0# Or X = 1.0# Then
            Bt = 0.0#
        Else
            Bt = Exp(GammLn(A + B) - GammLn(A) - GammLn(B) + A * Log(X) + B * Log(1.0# - X))
        End If

        If X < (A + 1.0#) / (A + B + 2.0#) Then
            Return Bt * BetaCf(A, B, X) / A
        Else
            Return 1.0# - Bt * BetaCf(B, A, 1.0# - X) / B
        End If

    End Function

    Private Function BetaCf(ByVal A As Double, ByVal B As Double, ByVal X As Double) As Double
        Const IT_MAX As Integer = 100
        Const EPS As Double = 0.00000000000000022204460492503131#
        Const FPMIN As Double = 2.2250738585072014E-308 / EPS

        Dim qab As Double = A + B
        Dim qap As Double = A + 1.0#
        Dim qam As Double = A - 1.0#

        Dim c As Double = 1.0#

        Dim d As Double = 1.0# - qab * X / qap
        If Abs(d) < FPMIN Then d = FPMIN
        d = 1.0# / d

        Dim h As Double = d

        Dim m As Integer

        For m = 1 To IT_MAX
            Dim m2 As Double = 2 * m
            Dim aa As Double = m * (B - m) * X / ((qam + m2) * (A + m2))

            d = 1.0# + aa * d
            If Abs(d) < FPMIN Then d = FPMIN
            c = 1.0# + aa / c
            If Abs(c) < FPMIN Then c = FPMIN
            d = 1.0# / d
            h *= d * c
            aa = -(A + m) * (qab + m) * X / ((A + m2) * (qap + m2))

            d = 1.0# + aa * d
            If Abs(d) < FPMIN Then d = FPMIN
            c = 1.0# + aa / c
            If Abs(c) < FPMIN Then c = FPMIN
            d = 1.0# / d
            Dim del As Double = d * c
            h *= del
            If Abs(del - 1.0#) < EPS Then Exit For
        Next

        If m > IT_MAX Then Throw New System.ArithmeticException("Bad A, B or ITMAX in BetaCf")

        Return h
    End Function
End Module

