Public Class frmFitCurve
    Inherits System.Windows.Forms.Form

    Dim dtSet As DataSet = New DataSet("test")
    Dim dtTable As DataTable
    Dim pkCol As DataColumn
    Dim pkRow As DataRow

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents DataGrid1 As System.Windows.Forms.DataGrid
    Friend WithEvents cmdPasteData As System.Windows.Forms.Button
    Friend WithEvents txtDegree As System.Windows.Forms.TextBox
    Friend WithEvents PlotSurface2D1 As NPlot.Windows.PlotSurface2D
    Friend WithEvents DataGrid2 As System.Windows.Forms.DataGrid
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cmdFitCurve As System.Windows.Forms.Button
    Friend WithEvents txtCurveEqn As System.Windows.Forms.TextBox
    Friend WithEvents cmdPlot As System.Windows.Forms.Button
    Friend WithEvents cmdClearPlot As System.Windows.Forms.Button
    Friend WithEvents cmdAcceptCurve As System.Windows.Forms.Button
    Friend WithEvents lblTitle As System.Windows.Forms.Label
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmFitCurve))
        Me.cmdFitCurve = New System.Windows.Forms.Button
        Me.txtCurveEqn = New System.Windows.Forms.TextBox
        Me.DataGrid1 = New System.Windows.Forms.DataGrid
        Me.cmdPasteData = New System.Windows.Forms.Button
        Me.txtDegree = New System.Windows.Forms.TextBox
        Me.PlotSurface2D1 = New NPlot.Windows.PlotSurface2D
        Me.DataGrid2 = New System.Windows.Forms.DataGrid
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.cmdPlot = New System.Windows.Forms.Button
        Me.cmdClearPlot = New System.Windows.Forms.Button
        Me.cmdAcceptCurve = New System.Windows.Forms.Button
        Me.lblTitle = New System.Windows.Forms.Label
        Me.cmdCancel = New System.Windows.Forms.Button
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGrid2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdFitCurve
        '
        Me.cmdFitCurve.Location = New System.Drawing.Point(488, 456)
        Me.cmdFitCurve.Name = "cmdFitCurve"
        Me.cmdFitCurve.Size = New System.Drawing.Size(88, 24)
        Me.cmdFitCurve.TabIndex = 0
        Me.cmdFitCurve.Text = "Fit Curve"
        '
        'txtCurveEqn
        '
        Me.txtCurveEqn.Location = New System.Drawing.Point(216, 424)
        Me.txtCurveEqn.Multiline = True
        Me.txtCurveEqn.Name = "txtCurveEqn"
        Me.txtCurveEqn.ReadOnly = True
        Me.txtCurveEqn.Size = New System.Drawing.Size(264, 96)
        Me.txtCurveEqn.TabIndex = 1
        Me.txtCurveEqn.Text = "TextBox1"
        '
        'DataGrid1
        '
        Me.DataGrid1.DataMember = ""
        Me.DataGrid1.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid1.Location = New System.Drawing.Point(8, 56)
        Me.DataGrid1.Name = "DataGrid1"
        Me.DataGrid1.Size = New System.Drawing.Size(200, 464)
        Me.DataGrid1.TabIndex = 2
        '
        'cmdPasteData
        '
        Me.cmdPasteData.Location = New System.Drawing.Point(8, 528)
        Me.cmdPasteData.Name = "cmdPasteData"
        Me.cmdPasteData.Size = New System.Drawing.Size(80, 24)
        Me.cmdPasteData.TabIndex = 3
        Me.cmdPasteData.Text = "Paste Data"
        '
        'txtDegree
        '
        Me.txtDegree.Location = New System.Drawing.Point(488, 424)
        Me.txtDegree.Name = "txtDegree"
        Me.txtDegree.Size = New System.Drawing.Size(88, 20)
        Me.txtDegree.TabIndex = 4
        Me.txtDegree.Text = "3"
        '
        'PlotSurface2D1
        '
        Me.PlotSurface2D1.AutoScaleAutoGeneratedAxes = False
        Me.PlotSurface2D1.AutoScaleTitle = False
        Me.PlotSurface2D1.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.PlotSurface2D1.DateTimeToolTip = False
        Me.PlotSurface2D1.Legend = Nothing
        Me.PlotSurface2D1.LegendZOrder = -1
        Me.PlotSurface2D1.Location = New System.Drawing.Point(216, 56)
        Me.PlotSurface2D1.Name = "PlotSurface2D1"
        Me.PlotSurface2D1.Padding = 10
        Me.PlotSurface2D1.RightMenu = Nothing
        Me.PlotSurface2D1.ShowCoordinates = True
        Me.PlotSurface2D1.Size = New System.Drawing.Size(360, 336)
        Me.PlotSurface2D1.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.None
        Me.PlotSurface2D1.TabIndex = 5
        Me.PlotSurface2D1.Text = "PlotSurface2D1"
        Me.PlotSurface2D1.Title = ""
        Me.PlotSurface2D1.TitleFont = New System.Drawing.Font("Arial", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel)
        Me.PlotSurface2D1.XAxis1 = Nothing
        Me.PlotSurface2D1.XAxis2 = Nothing
        Me.PlotSurface2D1.YAxis1 = Nothing
        Me.PlotSurface2D1.YAxis2 = Nothing
        '
        'DataGrid2
        '
        Me.DataGrid2.DataMember = ""
        Me.DataGrid2.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid2.Location = New System.Drawing.Point(584, 56)
        Me.DataGrid2.Name = "DataGrid2"
        Me.DataGrid2.Size = New System.Drawing.Size(200, 464)
        Me.DataGrid2.TabIndex = 6
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(216, 400)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(264, 16)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Curve Equation"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(488, 400)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(88, 16)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Curve Degree"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label3.Location = New System.Drawing.Point(16, 40)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(192, 16)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Given Data"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label4.Location = New System.Drawing.Point(584, 40)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(200, 16)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Fitted Data"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label5.Location = New System.Drawing.Point(224, 40)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(352, 16)
        Me.Label5.TabIndex = 11
        Me.Label5.Text = "Plots"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cmdPlot
        '
        Me.cmdPlot.Location = New System.Drawing.Point(488, 496)
        Me.cmdPlot.Name = "cmdPlot"
        Me.cmdPlot.Size = New System.Drawing.Size(88, 24)
        Me.cmdPlot.TabIndex = 12
        Me.cmdPlot.Text = "PLOT"
        '
        'cmdClearPlot
        '
        Me.cmdClearPlot.Location = New System.Drawing.Point(128, 528)
        Me.cmdClearPlot.Name = "cmdClearPlot"
        Me.cmdClearPlot.Size = New System.Drawing.Size(80, 24)
        Me.cmdClearPlot.TabIndex = 13
        Me.cmdClearPlot.Text = "Clear Plot"
        '
        'cmdAcceptCurve
        '
        Me.cmdAcceptCurve.Location = New System.Drawing.Point(216, 528)
        Me.cmdAcceptCurve.Name = "cmdAcceptCurve"
        Me.cmdAcceptCurve.Size = New System.Drawing.Size(128, 24)
        Me.cmdAcceptCurve.TabIndex = 14
        Me.cmdAcceptCurve.Text = "Accept Curve Equation"
        '
        'lblTitle
        '
        Me.lblTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblTitle.Location = New System.Drawing.Point(192, 8)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(408, 24)
        Me.lblTitle.TabIndex = 15
        Me.lblTitle.Text = "Curve Fit"
        Me.lblTitle.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(368, 528)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(112, 24)
        Me.cmdCancel.TabIndex = 16
        Me.cmdCancel.Text = "CANCEL"
        '
        'frmFitCurve
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(794, 568)
        Me.ControlBox = False
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.lblTitle)
        Me.Controls.Add(Me.cmdAcceptCurve)
        Me.Controls.Add(Me.cmdClearPlot)
        Me.Controls.Add(Me.cmdPlot)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.DataGrid2)
        Me.Controls.Add(Me.PlotSurface2D1)
        Me.Controls.Add(Me.txtDegree)
        Me.Controls.Add(Me.cmdPasteData)
        Me.Controls.Add(Me.DataGrid1)
        Me.Controls.Add(Me.txtCurveEqn)
        Me.Controls.Add(Me.cmdFitCurve)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmFitCurve"
        Me.Text = "Fit Curve"
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGrid2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public Structure POINT
        Friend x As Double
        Friend y As Double
    End Structure

    Dim t() As Point


    Public Function Trend(ByVal Data() As Point, ByVal Degree As Long) As Point()

        'degree 1 = straight line y=a+bx
        'degree n = polynomials!!

        Dim a(,) As Double
        Dim Ai(,) As Double
        Dim B(,) As Double
        Dim P(,) As Double
        Dim SigmaA() As Double
        Dim SigmaP() As Double
        Dim PointCount As Long
        Dim MaxTerm As Long
        Dim m As Long, n As Long
        Dim i As Long, j As Long
        Dim Ret() As Point

        Degree = Degree + 1

        MaxTerm = (2 * (Degree - 1))
        PointCount = UBound(Data) + 1

        ReDim SigmaA(MaxTerm - 1)
        ReDim SigmaP(MaxTerm - 1)

        ' Get the coefficients lists for matrices A, and P
        For m = 0 To (MaxTerm - 1)
            For n = 0 To (PointCount - 1)
                SigmaA(m) = SigmaA(m) + (Data(n).x ^ (m + 1))
                SigmaP(m) = SigmaP(m) + ((Data(n).x ^ m) * Data(n).y)
            Next
        Next

        ' Create Matrix A, and fill in the coefficients
        ReDim a(Degree - 1, Degree - 1)
        For i = 0 To (Degree - 1)
            For j = 0 To (Degree - 1)
                If i = 0 And j = 0 Then
                    a(i, j) = PointCount
                Else
                    a(i, j) = SigmaA((i + j) - 1)
                End If
            Next
        Next

        ' Create Matrix P, and fill in the coefficients
        ReDim P(Degree - 1, 0)
        For i = 0 To (Degree - 1)
            P(i, 0) = SigmaP(i)
        Next

        ' We have A, and P of AB=P, so we can solve B because B=AiP
        Ai = MxInverse(a)
        B = MxMultiplyCV(Ai, P)

        ' Now we solve the equations and generate the list of points
        PointCount = PointCount - 1
        ReDim Ret(PointCount)

        ' Work out non exponential first term
        For i = 0 To PointCount
            Ret(i).x = Data(i).x
            Ret(i).y = B(0, 0)
        Next

        ' Work out other exponential terms including exp 1
        For i = 0 To PointCount
            For j = 1 To Degree - 1
                Ret(i).y = Ret(i).y + (B(j, 0) * Ret(i).x ^ j)
            Next
        Next

        Trend = Ret

        Dim equation As String
        equation = B(0, 0)

        For j = 1 To Degree - 1
            Select Case B(j, 0)

                Case Is > 0
                    equation = equation & "+" & B(j, 0) & "*H^" & j

                Case Is < 0
                    equation = equation & B(j, 0) & "*H^" & j

            End Select

        Next

        'equation = Mid(equation, 1, Len(equation) - 3)
        txtCurveEqn.Text = equation

    End Function

    Public Function MxMultiplyCV(ByVal Matrix1(,) As Double, ByVal ColumnVector(,) As Double) As Double(,)

        Dim i As Long
        Dim j As Long
        Dim Rows As Long
        Dim Cols As Long
        Dim Ret(,) As Double

        Rows = UBound(Matrix1, 1)
        Cols = UBound(Matrix1, 2)

        ReDim Ret(UBound(ColumnVector, 1), 0) 'returns a column vector

        For i = 0 To Rows
            For j = 0 To Cols
                Ret(i, 0) = Ret(i, 0) + (Matrix1(i, j) * ColumnVector(j, 0))
            Next
        Next

        MxMultiplyCV = Ret

    End Function

    Public Function MxInverse(ByVal Matrix(,) As Double) As Double(,)

        Dim i As Long
        Dim j As Long
        Dim Rows As Long
        Dim Cols As Long
        Dim Tmp(,) As Double
        Dim Ret(,) As Double
        Dim Degree As Long

        Tmp = Matrix

        Rows = UBound(Tmp, 1)
        Cols = UBound(Tmp, 2)
        Degree = Cols + 1

        'Augment Identity matrix onto matrix M to get [M|I]
        ReDim Preserve Tmp(Rows, (Degree * 2) - 1)
        For i = Degree To (Degree * 2) - 1
            Tmp(i Mod Degree, i) = 1
        Next

        ' Now find the inverse using Gauss-Jordan Elimination which should get us [I|A-1]
        MxGaussJordan(Tmp)

        ' Copy the inverse (A-1) part to array to return
        ReDim Ret(Rows, Cols)
        For i = 0 To Rows
            For j = Degree To (Degree * 2) - 1
                Ret(i, j - Degree) = Tmp(i, j)
            Next
        Next

        MxInverse = Ret

    End Function

    Public Sub MxGaussJordan(ByVal Matrix(,) As Double)

        Dim Rows As Long
        Dim Cols As Long
        Dim P As Long
        Dim i As Long
        Dim j As Long
        Dim m As Double
        Dim d As Double
        Dim Pivot As Double

        Rows = UBound(Matrix, 1)
        Cols = UBound(Matrix, 2)

        ' Reduce so we get the leading diagonal
        For P = 0 To Rows
            Pivot = Matrix(P, P)
            For i = 0 To Rows
                If Not P = i Then
                    m = Matrix(i, P) / Pivot
                    For j = 0 To Cols
                        Matrix(i, j) = Matrix(i, j) + (Matrix(P, j) * -m)
                    Next
                End If
            Next
        Next

        'Divide through to get the identity matrix
        'Note: the identity matrix may have very small values (close to zero)
        'because of the way floating points are stored.
        For i = 0 To Rows
            d = Matrix(i, i)
            For j = 0 To Cols
                Matrix(i, j) = Matrix(i, j) / d
            Next
        Next

    End Sub

    Private Sub frmFitCurve_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtCurveEqn.Text = ""
        dtTable = dtSet.Tables.Add("t1")
        pkCol = dtTable.Columns.Add("X", Type.GetType("System.Decimal"))
        dtTable.Columns.Add("Y", Type.GetType("System.Decimal"))
        DataGrid1.SetDataBinding(dtSet, "t1")

        dtTable = dtSet.Tables.Add("t2")
        pkCol = dtTable.Columns.Add("X", Type.GetType("System.Decimal"))
        dtTable.Columns.Add("Y", Type.GetType("System.Decimal"))
        DataGrid2.SetDataBinding(dtSet, "t2")

    End Sub

    Private Sub cmdFitCurve_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFitCurve.Click
        Dim dr As DataRow

        Dim P(dtSet.Tables("t1").Rows.Count() - 1) As Point
        Dim j As Integer

        For j = 0 To (dtSet.Tables("t1").Rows.Count() - 1)
            P(j).x = dtSet.Tables("t1").Rows(j)(0)
            P(j).y = dtSet.Tables("t1").Rows(j)(1)
        Next

        t = Trend(P, Math.Round(Val(txtDegree.Text)))

        dtTable = dtSet.Tables("t2")
        dtTable.Clear()

        For j = 0 To t.Length - 1
            dr = dtTable.NewRow()
            dr("X") = t(j).x
            dr("Y") = t(j).y
            dtTable.Rows.Add(dr)
        Next

    End Sub

    Private Sub cmdPasteData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPasteData.Click
        Dim CLP As Clipboard
        dtTable = dtSet.Tables("t1")
        dtTable.Clear()

        Dim dr As DataRow
        Dim MyData As String
        Dim intCount As Integer

        MyData = CLP.GetDataObject.GetData(GetType(String))

        Try
            Dim strArr() As String = MyData.Split(ControlChars.CrLf)

            intCount = strArr.GetUpperBound(0)


            For i As Integer = 0 To intCount - 1
                dr = dtTable.NewRow()
                Dim tempStr() As String = strArr(i).Split(ControlChars.Tab)
                dr("X") = Val(tempStr(0))
                dr("Y") = Val(tempStr(1))
                dtTable.Rows.Add(dr)
            Next

            Dim lp As NPlot.LinePlot = New NPlot.LinePlot
            Dim pp As NPlot.PointPlot = New NPlot.PointPlot
            Dim yyy

            Dim XDAT As ArrayList = New ArrayList
            Dim YDAT As ArrayList = New ArrayList

            XDAT.Clear()
            YDAT.Clear()
            PlotSurface2D1.Clear()
            PlotSurface2D1.Title = "Curve Fitting"
            PlotSurface2D1.BackColor = Color.Empty

            For yyy = 0 To (dtSet.Tables("t1").Rows.Count() - 1)
                XDAT.Add(dtSet.Tables("t1").Rows(yyy)(0))
                YDAT.Add(dtSet.Tables("t1").Rows(yyy)(1))
            Next

            'lp.AbscissaData = XDAT
            'lp.DataSource = YDAT
            pp.AbscissaData = XDAT
            pp.DataSource = YDAT
            'lp.Color = Color.Blue
            pp.Marker.Color = Color.Blue
            pp.Marker.Type = NPlot.Marker.MarkerType.Cross1
            'PlotSurface2D1.Add(lp)
            PlotSurface2D1.Add(pp)
            PlotSurface2D1.XAxis1.Label = "X"
            PlotSurface2D1.YAxis1.Label = "Y"

            PlotSurface2D1.Refresh()

        Catch ex As Exception
            MessageBox.Show("Invalid Data!" & ControlChars.CrLf & "Only 2 columns of data separated by 'tabs' are accepted." _
                             & ControlChars.CrLf & "Data copied directly from Excel can be used.")
        End Try
    End Sub

    Private Sub cmdPlot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPlot.Click
        Dim lp As NPlot.LinePlot = New NPlot.LinePlot
        Dim pp As NPlot.PointPlot = New NPlot.PointPlot
        Dim yyy

        Dim XDAT As ArrayList = New ArrayList
        Dim YDAT As ArrayList = New ArrayList

        XDAT.Clear()
        YDAT.Clear()

        For yyy = 0 To t.Length - 1
            XDAT.Add(t(yyy).x)
            YDAT.Add(t(yyy).y)
        Next
        lp.AbscissaData = XDAT
        lp.DataSource = YDAT
        lp.Color = Color.FromArgb(255, Int(Rnd() * 255), Int(Rnd() * 255), Int(Rnd() * 255))
        PlotSurface2D1.Add(lp)

        PlotSurface2D1.Refresh()

    End Sub

    Private Sub txtDegree_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtDegree.Validating
        Dim value As Integer

        If Double.TryParse(Me.txtDegree.Text, Globalization.NumberStyles.Number, Nothing, value) Then
            If Me.txtDegree.Text < 1 Or Me.txtDegree.Text > 250 Then MessageBox.Show("Please enter a numerical value between 1 and 250.")
            Me.txtDegree.Text = value.ToString("n0")
        Else
            e.Cancel = True
            MessageBox.Show("Please enter a numerical value.")
        End If

    End Sub


    Private Sub cmdClearPlot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClearPlot.Click
        Dim lp As NPlot.LinePlot = New NPlot.LinePlot
        Dim pp As NPlot.PointPlot = New NPlot.PointPlot
        Dim yyy

        Dim XDAT As ArrayList = New ArrayList
        Dim YDAT As ArrayList = New ArrayList

        XDAT.Clear()
        YDAT.Clear()
        PlotSurface2D1.Clear()
        PlotSurface2D1.Title = "Curve Fitting"
        PlotSurface2D1.BackColor = Color.Empty

        For yyy = 0 To (dtSet.Tables("t1").Rows.Count() - 1)
            XDAT.Add(dtSet.Tables("t1").Rows(yyy)(0))
            YDAT.Add(dtSet.Tables("t1").Rows(yyy)(1))
        Next

        'lp.AbscissaData = XDAT
        'lp.DataSource = YDAT
        pp.AbscissaData = XDAT
        pp.DataSource = YDAT
        'lp.Color = Color.Blue
        pp.Marker.Color = Color.Blue
        pp.Marker.Type = NPlot.Marker.MarkerType.Cross1
        'PlotSurface2D1.Add(lp)
        PlotSurface2D1.Add(pp)
        PlotSurface2D1.XAxis1.Label = "X"
        PlotSurface2D1.YAxis1.Label = "Y"

        PlotSurface2D1.Refresh()
    End Sub

    Private Sub frmFitCurve_Closed(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        goForm1.Show()
        Me.Hide() 'Hide Form2

    End Sub

    Private Sub cmdAcceptCurve_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAcceptCurve.Click
        goForm1.Show()
        Me.Hide() 'Hide Form2
        Select Case goForm2.lblTitle.Text

            Case "Curve Fitting For Hazard"
                goForm1.txtHazardEqn.Text = goForm2.txtCurveEqn.Text

            Case "Curve Fitting For Vulnerability"
                goForm1.txtVulnerabilityEqn.Text = goForm2.txtCurveEqn.Text

        End Select

    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        goForm1.Show()
        Me.Hide() 'Hide Form2
    End Sub
End Class

