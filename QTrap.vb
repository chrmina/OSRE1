Module QTrap

    Function QTRAP(ByVal A As Double, ByVal B As Double, ByVal S As Double, ByVal EPS As Double) As Double
        Dim OLDS, DEL, X, SUM As Double
        Dim j, k, JMAX, IT, TNM As Integer

        JMAX = 20
        OLDS = -1.0E+30

        For j = 1 To JMAX

            If j = 1 Then
                S = 0.5 * (B - A) * (FUNC(A) + FUNC(B))
                IT = 1
            Else
                TNM = IT
                DEL = (B - A) / TNM
                X = A + 0.5 * DEL
                SUM = 0
                For k = 1 To IT
                    SUM = SUM + FUNC(X)
                    X = X + DEL
                Next k
                S = 0.5 * (S + (B - A) * SUM / TNM)
                IT = 2 * IT
            End If

            If Math.Abs(S - OLDS) < EPS * Math.Abs(OLDS) Then
                Return S
                Exit Function
            End If


            OLDS = S
        Next j
        'txtOutput.Text = "Too MANY steps!"
    End Function

    Function FUNC(ByVal x As Double) As Double

        'Change the function
        FUNC = (x ^ 2) + 7
    End Function
End Module
