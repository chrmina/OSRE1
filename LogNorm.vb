Option Explicit On 
Option Strict On

Imports System.Math

Module LogNorm

    Function LNCDF(ByVal lamda As Double, ByVal zeta As Double, ByVal A As Double, ByVal B As Double, ByVal S As Double, ByVal EPS As Double) As Double
        'Calculates the LogNormal Distribution CDF using the QTrap Routine of Numerical Integration

        Dim OLDS, DEL, X, SUM, LNPDFX, LNPDFA, LNPDFB As Double
        Dim j, k, JMAX, IT, TNM As Integer

        '*** LNPDFX = LogNorm PDF at point X (between A and B)
        '*** LNPDFA = LogNorm PDF at lower limit A
        '*** LNPDFB = LogNorm PDF at upper limit B

        JMAX = 20
        OLDS = -1.0E+130
        Application.DoEvents()

        For j = 1 To JMAX

            If j = 1 Then
                LNPDFA = (1 / (zeta * A * Sqrt(2 * PI)) * Exp(-0.5 * ((Log(A) - lamda) / zeta) ^ 2))
                LNPDFB = (1 / (zeta * B * Sqrt(2 * PI)) * Exp(-0.5 * ((Log(B) - lamda) / zeta) ^ 2))
                S = 0.5 * (B - A) * (LNPDFA + LNPDFB)
                IT = 1
            Else
                TNM = IT
                DEL = (B - A) / TNM
                X = A + 0.5 * DEL
                SUM = 0
                For k = 1 To IT
                    LNPDFX = (1 / (zeta * X * Sqrt(2 * PI)) * Exp(-0.5 * ((Log(X) - lamda) / zeta) ^ 2))
                    SUM = SUM + LNPDFX
                    X = X + DEL
                Next k
                S = 0.5 * (S + (B - A) * SUM / TNM)
                IT = 2 * IT
            End If

            If Abs(S - OLDS) < EPS * Abs(OLDS) Then
                Return S
                Exit Function
            End If


            OLDS = S
        Next j
        'txtOutput.Text = "Too MANY steps!"
    End Function


End Module
