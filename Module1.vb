Option Explicit On 

Module Module1

    ' USED TO SWITCH BETWEEN THE TWO FORMS
    Public goForm1 As New frmMain
    Public goForm2 As New frmFitCurve
    Public goForm3 As New frmAbout

    

    Public Sub Main()
        Application.Run(goForm1) 'Show Form1
    End Sub

End Module
