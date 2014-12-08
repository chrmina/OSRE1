Public Class frmMain

    Inherits System.Windows.Forms.Form
    Dim dtSet As DataSet = New DataSet("test")
    Dim dtTable As DataTable
    Dim pkCol As DataColumn
    Dim pkRow As DataRow
    Public HAZARDMATRIX(1100, 1100) As Double
    Public VULNMATRIX(1100, 1100) As Double
    Public HAZCONFINT(110, 1100) As Double
    Public HAZCONFINTCDF(110, 1100) As Double
    Public HAZCONFINTPDF(110, 1100) As Double
    Public RISKMATRIX(110, 1100) As Double

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
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents tbgHazard As System.Windows.Forms.TabPage
    Friend WithEvents tbgVulnerability As System.Windows.Forms.TabPage
    Friend WithEvents tbgDamage As System.Windows.Forms.TabPage
    Friend WithEvents DataGrid1 As System.Windows.Forms.DataGrid
    Friend WithEvents lblHazardEqn As System.Windows.Forms.Label
    Friend WithEvents txtHazardEqn As System.Windows.Forms.TextBox
    Friend WithEvents cboHazardDistrType As System.Windows.Forms.ComboBox
    Friend WithEvents lblHazardDistrType As System.Windows.Forms.Label
    Friend WithEvents lblHazardDistrParam As System.Windows.Forms.Label
    Friend WithEvents txtHazardDistrParam2 As System.Windows.Forms.TextBox
    Friend WithEvents lblHazardDistrParam2 As System.Windows.Forms.Label
    Friend WithEvents lblVulnerabilityDistrParam2 As System.Windows.Forms.Label
    Friend WithEvents lblVulnerabilityDistrParam As System.Windows.Forms.Label
    Friend WithEvents cboVulnerabilityDistrType As System.Windows.Forms.ComboBox
    Friend WithEvents lblVulnerabilityDistrType As System.Windows.Forms.Label
    Friend WithEvents txtVulnerabilityEqn As System.Windows.Forms.TextBox
    Friend WithEvents lblVulnerabilityEqn As System.Windows.Forms.Label
    Friend WithEvents DataGrid2 As System.Windows.Forms.DataGrid
    Friend WithEvents cboConfidence As System.Windows.Forms.ComboBox
    Friend WithEvents lblConfidence As System.Windows.Forms.Label
    Friend WithEvents DataGrid3 As System.Windows.Forms.DataGrid
    Friend WithEvents cmdAddGraph As System.Windows.Forms.Button
    Friend WithEvents cmdClear As System.Windows.Forms.Button
    Friend WithEvents PlotSurface2D1 As NPlot.Windows.PlotSurface2D
    Friend WithEvents PlotSurface2D2 As NPlot.Windows.PlotSurface2D
    Friend WithEvents cmdPlot As System.Windows.Forms.Button
    Friend WithEvents txtGenerateData As System.Windows.Forms.Button
    Friend WithEvents PlotSurface2D3 As NPlot.Windows.PlotSurface2D
    Friend WithEvents cmdGenVulnData As System.Windows.Forms.Button
    Friend WithEvents cmdPlotVuln As System.Windows.Forms.Button
    Friend WithEvents cmdGenDamData As System.Windows.Forms.Button
    Friend WithEvents mnuMain As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem6 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem7 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem8 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem9 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem10 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem11 As System.Windows.Forms.MenuItem
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cmdPasteHzdData As System.Windows.Forms.Button
    Friend WithEvents cmdPasteVulnData As System.Windows.Forms.Button
    Friend WithEvents prbProgress As System.Windows.Forms.ProgressBar
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents prbTotProgress As System.Windows.Forms.ProgressBar
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtVulnerabilityDistrParam2 As System.Windows.Forms.TextBox
    Friend WithEvents txtProcess As System.Windows.Forms.TextBox
    Friend WithEvents cmdFitHazCurve As System.Windows.Forms.Button
    Friend WithEvents cmdFitVulnCurve As System.Windows.Forms.Button
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents lblAboutOSRE As System.Windows.Forms.Label
    Friend WithEvents lblIntro As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmMain))
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.Label3 = New System.Windows.Forms.Label
        Me.lblAboutOSRE = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.lblIntro = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.tbgHazard = New System.Windows.Forms.TabPage
        Me.cmdFitHazCurve = New System.Windows.Forms.Button
        Me.cmdPasteHzdData = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtGenerateData = New System.Windows.Forms.Button
        Me.PlotSurface2D1 = New NPlot.Windows.PlotSurface2D
        Me.cmdPlot = New System.Windows.Forms.Button
        Me.txtHazardDistrParam2 = New System.Windows.Forms.TextBox
        Me.lblHazardDistrParam2 = New System.Windows.Forms.Label
        Me.lblHazardDistrParam = New System.Windows.Forms.Label
        Me.cboHazardDistrType = New System.Windows.Forms.ComboBox
        Me.lblHazardDistrType = New System.Windows.Forms.Label
        Me.txtHazardEqn = New System.Windows.Forms.TextBox
        Me.lblHazardEqn = New System.Windows.Forms.Label
        Me.DataGrid1 = New System.Windows.Forms.DataGrid
        Me.tbgVulnerability = New System.Windows.Forms.TabPage
        Me.cmdFitVulnCurve = New System.Windows.Forms.Button
        Me.cmdPasteVulnData = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.cmdPlotVuln = New System.Windows.Forms.Button
        Me.cmdGenVulnData = New System.Windows.Forms.Button
        Me.PlotSurface2D2 = New NPlot.Windows.PlotSurface2D
        Me.txtVulnerabilityDistrParam2 = New System.Windows.Forms.TextBox
        Me.lblVulnerabilityDistrParam2 = New System.Windows.Forms.Label
        Me.lblVulnerabilityDistrParam = New System.Windows.Forms.Label
        Me.cboVulnerabilityDistrType = New System.Windows.Forms.ComboBox
        Me.lblVulnerabilityDistrType = New System.Windows.Forms.Label
        Me.txtVulnerabilityEqn = New System.Windows.Forms.TextBox
        Me.lblVulnerabilityEqn = New System.Windows.Forms.Label
        Me.DataGrid2 = New System.Windows.Forms.DataGrid
        Me.tbgDamage = New System.Windows.Forms.TabPage
        Me.txtProcess = New System.Windows.Forms.TextBox
        Me.prbTotProgress = New System.Windows.Forms.ProgressBar
        Me.prbProgress = New System.Windows.Forms.ProgressBar
        Me.cmdGenDamData = New System.Windows.Forms.Button
        Me.PlotSurface2D3 = New NPlot.Windows.PlotSurface2D
        Me.cmdClear = New System.Windows.Forms.Button
        Me.cmdAddGraph = New System.Windows.Forms.Button
        Me.cboConfidence = New System.Windows.Forms.ComboBox
        Me.lblConfidence = New System.Windows.Forms.Label
        Me.DataGrid3 = New System.Windows.Forms.DataGrid
        Me.mnuMain = New System.Windows.Forms.MainMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.MenuItem4 = New System.Windows.Forms.MenuItem
        Me.MenuItem8 = New System.Windows.Forms.MenuItem
        Me.MenuItem9 = New System.Windows.Forms.MenuItem
        Me.MenuItem10 = New System.Windows.Forms.MenuItem
        Me.MenuItem11 = New System.Windows.Forms.MenuItem
        Me.MenuItem5 = New System.Windows.Forms.MenuItem
        Me.MenuItem6 = New System.Windows.Forms.MenuItem
        Me.MenuItem7 = New System.Windows.Forms.MenuItem
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.tbgHazard.SuspendLayout()
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tbgVulnerability.SuspendLayout()
        CType(Me.DataGrid2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tbgDamage.SuspendLayout()
        CType(Me.DataGrid3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.tbgHazard)
        Me.TabControl1.Controls.Add(Me.tbgVulnerability)
        Me.TabControl1.Controls.Add(Me.tbgDamage)
        Me.TabControl1.Location = New System.Drawing.Point(8, 8)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(776, 536)
        Me.TabControl1.TabIndex = 9
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage1.Controls.Add(Me.Label3)
        Me.TabPage1.Controls.Add(Me.lblAboutOSRE)
        Me.TabPage1.Controls.Add(Me.Label6)
        Me.TabPage1.Controls.Add(Me.lblIntro)
        Me.TabPage1.Controls.Add(Me.Label4)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Size = New System.Drawing.Size(768, 510)
        Me.TabPage1.TabIndex = 3
        Me.TabPage1.Text = "General Info"
        '
        'Label3
        '
        Me.Label3.Image = CType(resources.GetObject("Label3.Image"), System.Drawing.Image)
        Me.Label3.Location = New System.Drawing.Point(152, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(464, 70)
        Me.Label3.TabIndex = 1
        '
        'lblAboutOSRE
        '
        Me.lblAboutOSRE.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAboutOSRE.Location = New System.Drawing.Point(32, 392)
        Me.lblAboutOSRE.Name = "lblAboutOSRE"
        Me.lblAboutOSRE.Size = New System.Drawing.Size(712, 104)
        Me.lblAboutOSRE.TabIndex = 6
        Me.lblAboutOSRE.Text = "History of OSRE"
        Me.lblAboutOSRE.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(236, 352)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(296, 35)
        Me.Label6.TabIndex = 4
        Me.Label6.Text = "PROJECT INFORMATION"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblIntro
        '
        Me.lblIntro.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIntro.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblIntro.Location = New System.Drawing.Point(28, 128)
        Me.lblIntro.Name = "lblIntro"
        Me.lblIntro.Size = New System.Drawing.Size(712, 216)
        Me.lblIntro.TabIndex = 3
        Me.lblIntro.Text = "OSRE Introduction"
        Me.lblIntro.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(152, 88)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(464, 35)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "Open Source Risk Engine"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'tbgHazard
        '
        Me.tbgHazard.Controls.Add(Me.cmdFitHazCurve)
        Me.tbgHazard.Controls.Add(Me.cmdPasteHzdData)
        Me.tbgHazard.Controls.Add(Me.Label1)
        Me.tbgHazard.Controls.Add(Me.txtGenerateData)
        Me.tbgHazard.Controls.Add(Me.PlotSurface2D1)
        Me.tbgHazard.Controls.Add(Me.cmdPlot)
        Me.tbgHazard.Controls.Add(Me.txtHazardDistrParam2)
        Me.tbgHazard.Controls.Add(Me.lblHazardDistrParam2)
        Me.tbgHazard.Controls.Add(Me.lblHazardDistrParam)
        Me.tbgHazard.Controls.Add(Me.cboHazardDistrType)
        Me.tbgHazard.Controls.Add(Me.lblHazardDistrType)
        Me.tbgHazard.Controls.Add(Me.txtHazardEqn)
        Me.tbgHazard.Controls.Add(Me.lblHazardEqn)
        Me.tbgHazard.Controls.Add(Me.DataGrid1)
        Me.tbgHazard.Location = New System.Drawing.Point(4, 22)
        Me.tbgHazard.Name = "tbgHazard"
        Me.tbgHazard.Size = New System.Drawing.Size(768, 510)
        Me.tbgHazard.TabIndex = 0
        Me.tbgHazard.Text = "Hazard"
        '
        'cmdFitHazCurve
        '
        Me.cmdFitHazCurve.Location = New System.Drawing.Point(16, 472)
        Me.cmdFitHazCurve.Name = "cmdFitHazCurve"
        Me.cmdFitHazCurve.Size = New System.Drawing.Size(88, 24)
        Me.cmdFitHazCurve.TabIndex = 16
        Me.cmdFitHazCurve.Text = "Fit Curve"
        Me.ToolTip1.SetToolTip(Me.cmdFitHazCurve, "Click here if you want to do a regression on data.")
        '
        'cmdPasteHzdData
        '
        Me.cmdPasteHzdData.Location = New System.Drawing.Point(112, 352)
        Me.cmdPasteHzdData.Name = "cmdPasteHzdData"
        Me.cmdPasteHzdData.Size = New System.Drawing.Size(72, 26)
        Me.cmdPasteHzdData.TabIndex = 15
        Me.cmdPasteHzdData.Text = "Paste Data"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(16, 320)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 17)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "Pexc="
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtGenerateData
        '
        Me.txtGenerateData.Location = New System.Drawing.Point(16, 352)
        Me.txtGenerateData.Name = "txtGenerateData"
        Me.txtGenerateData.Size = New System.Drawing.Size(88, 26)
        Me.txtGenerateData.TabIndex = 13
        Me.txtGenerateData.Text = "Generate Data"
        '
        'PlotSurface2D1
        '
        Me.PlotSurface2D1.AutoScaleAutoGeneratedAxes = False
        Me.PlotSurface2D1.AutoScaleTitle = False
        Me.PlotSurface2D1.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.PlotSurface2D1.DateTimeToolTip = False
        Me.PlotSurface2D1.Legend = Nothing
        Me.PlotSurface2D1.LegendZOrder = -1
        Me.PlotSurface2D1.Location = New System.Drawing.Point(256, 8)
        Me.PlotSurface2D1.Name = "PlotSurface2D1"
        Me.PlotSurface2D1.Padding = 10
        Me.PlotSurface2D1.RightMenu = Nothing
        Me.PlotSurface2D1.ShowCoordinates = True
        Me.PlotSurface2D1.Size = New System.Drawing.Size(504, 488)
        Me.PlotSurface2D1.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.None
        Me.PlotSurface2D1.TabIndex = 12
        Me.PlotSurface2D1.Text = "PlotSurface2D1"
        Me.PlotSurface2D1.Title = ""
        Me.PlotSurface2D1.TitleFont = New System.Drawing.Font("Arial", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel)
        Me.PlotSurface2D1.XAxis1 = Nothing
        Me.PlotSurface2D1.XAxis2 = Nothing
        Me.PlotSurface2D1.YAxis1 = Nothing
        Me.PlotSurface2D1.YAxis2 = Nothing
        '
        'cmdPlot
        '
        Me.cmdPlot.Location = New System.Drawing.Point(192, 352)
        Me.cmdPlot.Name = "cmdPlot"
        Me.cmdPlot.Size = New System.Drawing.Size(48, 26)
        Me.cmdPlot.TabIndex = 11
        Me.cmdPlot.Text = "Plot"
        '
        'txtHazardDistrParam2
        '
        Me.txtHazardDistrParam2.Location = New System.Drawing.Point(176, 472)
        Me.txtHazardDistrParam2.Name = "txtHazardDistrParam2"
        Me.txtHazardDistrParam2.Size = New System.Drawing.Size(64, 20)
        Me.txtHazardDistrParam2.TabIndex = 10
        Me.txtHazardDistrParam2.Text = "0.2"
        Me.ToolTip1.SetToolTip(Me.txtHazardDistrParam2, "Input the value of the second parameter of the Hazard Distribution")
        '
        'lblHazardDistrParam2
        '
        Me.lblHazardDistrParam2.Location = New System.Drawing.Point(136, 472)
        Me.lblHazardDistrParam2.Name = "lblHazardDistrParam2"
        Me.lblHazardDistrParam2.Size = New System.Drawing.Size(40, 17)
        Me.lblHazardDistrParam2.TabIndex = 9
        Me.lblHazardDistrParam2.Text = "Zeta"
        Me.lblHazardDistrParam2.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'lblHazardDistrParam
        '
        Me.lblHazardDistrParam.Location = New System.Drawing.Point(16, 440)
        Me.lblHazardDistrParam.Name = "lblHazardDistrParam"
        Me.lblHazardDistrParam.Size = New System.Drawing.Size(224, 17)
        Me.lblHazardDistrParam.TabIndex = 6
        Me.lblHazardDistrParam.Text = "Hazard Distribution Parameters"
        Me.lblHazardDistrParam.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cboHazardDistrType
        '
        Me.cboHazardDistrType.Items.AddRange(New Object() {"Lognormal", "Beta"})
        Me.cboHazardDistrType.Location = New System.Drawing.Point(16, 408)
        Me.cboHazardDistrType.Name = "cboHazardDistrType"
        Me.cboHazardDistrType.Size = New System.Drawing.Size(224, 21)
        Me.cboHazardDistrType.TabIndex = 5
        Me.cboHazardDistrType.Text = "Lognormal"
        Me.ToolTip1.SetToolTip(Me.cboHazardDistrType, "Choose the Probability Distribution of the Hazard Equation")
        '
        'lblHazardDistrType
        '
        Me.lblHazardDistrType.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHazardDistrType.Location = New System.Drawing.Point(16, 384)
        Me.lblHazardDistrType.Name = "lblHazardDistrType"
        Me.lblHazardDistrType.Size = New System.Drawing.Size(224, 17)
        Me.lblHazardDistrType.TabIndex = 4
        Me.lblHazardDistrType.Text = "Hazard Distribution"
        Me.lblHazardDistrType.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtHazardEqn
        '
        Me.txtHazardEqn.Location = New System.Drawing.Point(56, 320)
        Me.txtHazardEqn.Name = "txtHazardEqn"
        Me.txtHazardEqn.Size = New System.Drawing.Size(184, 20)
        Me.txtHazardEqn.TabIndex = 3
        Me.txtHazardEqn.Text = ""
        Me.ToolTip1.SetToolTip(Me.txtHazardEqn, "Please enter the equation of the mean value of Hazard Vs its Probability of Exced" & _
        "ance.")
        '
        'lblHazardEqn
        '
        Me.lblHazardEqn.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHazardEqn.Location = New System.Drawing.Point(16, 288)
        Me.lblHazardEqn.Name = "lblHazardEqn"
        Me.lblHazardEqn.Size = New System.Drawing.Size(224, 17)
        Me.lblHazardEqn.TabIndex = 2
        Me.lblHazardEqn.Text = "Hazard Equation"
        Me.lblHazardEqn.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'DataGrid1
        '
        Me.DataGrid1.BackgroundColor = System.Drawing.Color.LightGray
        Me.DataGrid1.DataMember = ""
        Me.DataGrid1.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid1.Location = New System.Drawing.Point(16, 8)
        Me.DataGrid1.Name = "DataGrid1"
        Me.DataGrid1.Size = New System.Drawing.Size(224, 271)
        Me.DataGrid1.TabIndex = 0
        '
        'tbgVulnerability
        '
        Me.tbgVulnerability.Controls.Add(Me.cmdFitVulnCurve)
        Me.tbgVulnerability.Controls.Add(Me.cmdPasteVulnData)
        Me.tbgVulnerability.Controls.Add(Me.Label2)
        Me.tbgVulnerability.Controls.Add(Me.cmdPlotVuln)
        Me.tbgVulnerability.Controls.Add(Me.cmdGenVulnData)
        Me.tbgVulnerability.Controls.Add(Me.PlotSurface2D2)
        Me.tbgVulnerability.Controls.Add(Me.txtVulnerabilityDistrParam2)
        Me.tbgVulnerability.Controls.Add(Me.lblVulnerabilityDistrParam2)
        Me.tbgVulnerability.Controls.Add(Me.lblVulnerabilityDistrParam)
        Me.tbgVulnerability.Controls.Add(Me.cboVulnerabilityDistrType)
        Me.tbgVulnerability.Controls.Add(Me.lblVulnerabilityDistrType)
        Me.tbgVulnerability.Controls.Add(Me.txtVulnerabilityEqn)
        Me.tbgVulnerability.Controls.Add(Me.lblVulnerabilityEqn)
        Me.tbgVulnerability.Controls.Add(Me.DataGrid2)
        Me.tbgVulnerability.Location = New System.Drawing.Point(4, 22)
        Me.tbgVulnerability.Name = "tbgVulnerability"
        Me.tbgVulnerability.Size = New System.Drawing.Size(768, 510)
        Me.tbgVulnerability.TabIndex = 1
        Me.tbgVulnerability.Text = "Vulnerability"
        '
        'cmdFitVulnCurve
        '
        Me.cmdFitVulnCurve.Location = New System.Drawing.Point(16, 472)
        Me.cmdFitVulnCurve.Name = "cmdFitVulnCurve"
        Me.cmdFitVulnCurve.Size = New System.Drawing.Size(88, 24)
        Me.cmdFitVulnCurve.TabIndex = 27
        Me.cmdFitVulnCurve.Text = "Fit Curve"
        '
        'cmdPasteVulnData
        '
        Me.cmdPasteVulnData.Location = New System.Drawing.Point(112, 352)
        Me.cmdPasteVulnData.Name = "cmdPasteVulnData"
        Me.cmdPasteVulnData.Size = New System.Drawing.Size(72, 26)
        Me.cmdPasteVulnData.TabIndex = 26
        Me.cmdPasteVulnData.Text = "Paste Data"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(16, 320)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(24, 17)
        Me.Label2.TabIndex = 25
        Me.Label2.Text = "D="
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cmdPlotVuln
        '
        Me.cmdPlotVuln.Location = New System.Drawing.Point(192, 352)
        Me.cmdPlotVuln.Name = "cmdPlotVuln"
        Me.cmdPlotVuln.Size = New System.Drawing.Size(48, 26)
        Me.cmdPlotVuln.TabIndex = 24
        Me.cmdPlotVuln.Text = "Plot"
        '
        'cmdGenVulnData
        '
        Me.cmdGenVulnData.Location = New System.Drawing.Point(16, 352)
        Me.cmdGenVulnData.Name = "cmdGenVulnData"
        Me.cmdGenVulnData.Size = New System.Drawing.Size(88, 26)
        Me.cmdGenVulnData.TabIndex = 23
        Me.cmdGenVulnData.Text = "Generate Data"
        '
        'PlotSurface2D2
        '
        Me.PlotSurface2D2.AutoScaleAutoGeneratedAxes = False
        Me.PlotSurface2D2.AutoScaleTitle = False
        Me.PlotSurface2D2.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.PlotSurface2D2.DateTimeToolTip = False
        Me.PlotSurface2D2.Legend = Nothing
        Me.PlotSurface2D2.LegendZOrder = -1
        Me.PlotSurface2D2.Location = New System.Drawing.Point(256, 8)
        Me.PlotSurface2D2.Name = "PlotSurface2D2"
        Me.PlotSurface2D2.Padding = 10
        Me.PlotSurface2D2.RightMenu = Nothing
        Me.PlotSurface2D2.ShowCoordinates = True
        Me.PlotSurface2D2.Size = New System.Drawing.Size(504, 488)
        Me.PlotSurface2D2.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.None
        Me.PlotSurface2D2.TabIndex = 22
        Me.PlotSurface2D2.Text = "PlotSurface2D2"
        Me.PlotSurface2D2.Title = ""
        Me.PlotSurface2D2.TitleFont = New System.Drawing.Font("Arial", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel)
        Me.PlotSurface2D2.XAxis1 = Nothing
        Me.PlotSurface2D2.XAxis2 = Nothing
        Me.PlotSurface2D2.YAxis1 = Nothing
        Me.PlotSurface2D2.YAxis2 = Nothing
        '
        'txtVulnerabilityDistrParam2
        '
        Me.txtVulnerabilityDistrParam2.Location = New System.Drawing.Point(176, 472)
        Me.txtVulnerabilityDistrParam2.Name = "txtVulnerabilityDistrParam2"
        Me.txtVulnerabilityDistrParam2.Size = New System.Drawing.Size(64, 20)
        Me.txtVulnerabilityDistrParam2.TabIndex = 21
        Me.txtVulnerabilityDistrParam2.Text = "0.2"
        '
        'lblVulnerabilityDistrParam2
        '
        Me.lblVulnerabilityDistrParam2.Location = New System.Drawing.Point(136, 472)
        Me.lblVulnerabilityDistrParam2.Name = "lblVulnerabilityDistrParam2"
        Me.lblVulnerabilityDistrParam2.Size = New System.Drawing.Size(40, 17)
        Me.lblVulnerabilityDistrParam2.TabIndex = 20
        Me.lblVulnerabilityDistrParam2.Text = "Zeta"
        Me.lblVulnerabilityDistrParam2.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'lblVulnerabilityDistrParam
        '
        Me.lblVulnerabilityDistrParam.Location = New System.Drawing.Point(16, 440)
        Me.lblVulnerabilityDistrParam.Name = "lblVulnerabilityDistrParam"
        Me.lblVulnerabilityDistrParam.Size = New System.Drawing.Size(224, 17)
        Me.lblVulnerabilityDistrParam.TabIndex = 17
        Me.lblVulnerabilityDistrParam.Text = "Vulnerability Distribution Parameters"
        Me.lblVulnerabilityDistrParam.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cboVulnerabilityDistrType
        '
        Me.cboVulnerabilityDistrType.Items.AddRange(New Object() {"Lognormal", "Beta"})
        Me.cboVulnerabilityDistrType.Location = New System.Drawing.Point(16, 408)
        Me.cboVulnerabilityDistrType.Name = "cboVulnerabilityDistrType"
        Me.cboVulnerabilityDistrType.Size = New System.Drawing.Size(224, 21)
        Me.cboVulnerabilityDistrType.TabIndex = 16
        Me.cboVulnerabilityDistrType.Text = "Beta"
        '
        'lblVulnerabilityDistrType
        '
        Me.lblVulnerabilityDistrType.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVulnerabilityDistrType.Location = New System.Drawing.Point(16, 384)
        Me.lblVulnerabilityDistrType.Name = "lblVulnerabilityDistrType"
        Me.lblVulnerabilityDistrType.Size = New System.Drawing.Size(224, 17)
        Me.lblVulnerabilityDistrType.TabIndex = 15
        Me.lblVulnerabilityDistrType.Text = "Vulnerability Distribution"
        Me.lblVulnerabilityDistrType.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtVulnerabilityEqn
        '
        Me.txtVulnerabilityEqn.Location = New System.Drawing.Point(48, 320)
        Me.txtVulnerabilityEqn.Name = "txtVulnerabilityEqn"
        Me.txtVulnerabilityEqn.Size = New System.Drawing.Size(192, 20)
        Me.txtVulnerabilityEqn.TabIndex = 14
        Me.txtVulnerabilityEqn.Text = ""
        '
        'lblVulnerabilityEqn
        '
        Me.lblVulnerabilityEqn.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVulnerabilityEqn.Location = New System.Drawing.Point(16, 288)
        Me.lblVulnerabilityEqn.Name = "lblVulnerabilityEqn"
        Me.lblVulnerabilityEqn.Size = New System.Drawing.Size(224, 17)
        Me.lblVulnerabilityEqn.TabIndex = 13
        Me.lblVulnerabilityEqn.Text = "Vulnerability Equation"
        Me.lblVulnerabilityEqn.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'DataGrid2
        '
        Me.DataGrid2.BackgroundColor = System.Drawing.Color.LightGray
        Me.DataGrid2.DataMember = ""
        Me.DataGrid2.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid2.Location = New System.Drawing.Point(16, 8)
        Me.DataGrid2.Name = "DataGrid2"
        Me.DataGrid2.Size = New System.Drawing.Size(224, 272)
        Me.DataGrid2.TabIndex = 11
        '
        'tbgDamage
        '
        Me.tbgDamage.Controls.Add(Me.txtProcess)
        Me.tbgDamage.Controls.Add(Me.prbTotProgress)
        Me.tbgDamage.Controls.Add(Me.prbProgress)
        Me.tbgDamage.Controls.Add(Me.cmdGenDamData)
        Me.tbgDamage.Controls.Add(Me.PlotSurface2D3)
        Me.tbgDamage.Controls.Add(Me.cmdClear)
        Me.tbgDamage.Controls.Add(Me.cmdAddGraph)
        Me.tbgDamage.Controls.Add(Me.cboConfidence)
        Me.tbgDamage.Controls.Add(Me.lblConfidence)
        Me.tbgDamage.Controls.Add(Me.DataGrid3)
        Me.tbgDamage.Location = New System.Drawing.Point(4, 22)
        Me.tbgDamage.Name = "tbgDamage"
        Me.tbgDamage.Size = New System.Drawing.Size(768, 510)
        Me.tbgDamage.TabIndex = 2
        Me.tbgDamage.Text = "Damage"
        '
        'txtProcess
        '
        Me.txtProcess.BackColor = System.Drawing.SystemColors.WindowText
        Me.txtProcess.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtProcess.ForeColor = System.Drawing.Color.Lime
        Me.txtProcess.Location = New System.Drawing.Point(16, 320)
        Me.txtProcess.Multiline = True
        Me.txtProcess.Name = "txtProcess"
        Me.txtProcess.ReadOnly = True
        Me.txtProcess.Size = New System.Drawing.Size(224, 96)
        Me.txtProcess.TabIndex = 37
        Me.txtProcess.Text = "Current Process"
        '
        'prbTotProgress
        '
        Me.prbTotProgress.Location = New System.Drawing.Point(16, 304)
        Me.prbTotProgress.Name = "prbTotProgress"
        Me.prbTotProgress.Size = New System.Drawing.Size(224, 9)
        Me.prbTotProgress.TabIndex = 36
        '
        'prbProgress
        '
        Me.prbProgress.Location = New System.Drawing.Point(16, 280)
        Me.prbProgress.Name = "prbProgress"
        Me.prbProgress.Size = New System.Drawing.Size(224, 8)
        Me.prbProgress.TabIndex = 33
        '
        'cmdGenDamData
        '
        Me.cmdGenDamData.Location = New System.Drawing.Point(56, 248)
        Me.cmdGenDamData.Name = "cmdGenDamData"
        Me.cmdGenDamData.Size = New System.Drawing.Size(136, 26)
        Me.cmdGenDamData.TabIndex = 31
        Me.cmdGenDamData.Text = "Generate Damage Data"
        '
        'PlotSurface2D3
        '
        Me.PlotSurface2D3.AutoScaleAutoGeneratedAxes = False
        Me.PlotSurface2D3.AutoScaleTitle = False
        Me.PlotSurface2D3.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.PlotSurface2D3.DateTimeToolTip = False
        Me.PlotSurface2D3.Legend = Nothing
        Me.PlotSurface2D3.LegendZOrder = -1
        Me.PlotSurface2D3.Location = New System.Drawing.Point(256, 8)
        Me.PlotSurface2D3.Name = "PlotSurface2D3"
        Me.PlotSurface2D3.Padding = 10
        Me.PlotSurface2D3.RightMenu = Nothing
        Me.PlotSurface2D3.ShowCoordinates = True
        Me.PlotSurface2D3.Size = New System.Drawing.Size(504, 488)
        Me.PlotSurface2D3.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.None
        Me.PlotSurface2D3.TabIndex = 30
        Me.PlotSurface2D3.Text = "PlotSurface2D3"
        Me.PlotSurface2D3.Title = ""
        Me.PlotSurface2D3.TitleFont = New System.Drawing.Font("Arial", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel)
        Me.PlotSurface2D3.XAxis1 = Nothing
        Me.PlotSurface2D3.XAxis2 = Nothing
        Me.PlotSurface2D3.YAxis1 = Nothing
        Me.PlotSurface2D3.YAxis2 = Nothing
        '
        'cmdClear
        '
        Me.cmdClear.Location = New System.Drawing.Point(144, 472)
        Me.cmdClear.Name = "cmdClear"
        Me.cmdClear.Size = New System.Drawing.Size(96, 24)
        Me.cmdClear.TabIndex = 29
        Me.cmdClear.Text = "Clear All"
        '
        'cmdAddGraph
        '
        Me.cmdAddGraph.Location = New System.Drawing.Point(16, 472)
        Me.cmdAddGraph.Name = "cmdAddGraph"
        Me.cmdAddGraph.Size = New System.Drawing.Size(88, 24)
        Me.cmdAddGraph.TabIndex = 28
        Me.cmdAddGraph.Text = "Add Graph"
        '
        'cboConfidence
        '
        Me.cboConfidence.Items.AddRange(New Object() {"1%", "10%", "20%", "30%", "40%", "50%", "60%", "70%", "80%", "90%", "95%", "99%"})
        Me.cboConfidence.Location = New System.Drawing.Point(16, 440)
        Me.cboConfidence.Name = "cboConfidence"
        Me.cboConfidence.Size = New System.Drawing.Size(224, 21)
        Me.cboConfidence.TabIndex = 27
        Me.cboConfidence.Text = "50%"
        '
        'lblConfidence
        '
        Me.lblConfidence.Location = New System.Drawing.Point(16, 424)
        Me.lblConfidence.Name = "lblConfidence"
        Me.lblConfidence.Size = New System.Drawing.Size(224, 17)
        Me.lblConfidence.TabIndex = 26
        Me.lblConfidence.Text = "Confidence Interval"
        Me.lblConfidence.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'DataGrid3
        '
        Me.DataGrid3.BackgroundColor = System.Drawing.Color.LightGray
        Me.DataGrid3.DataMember = ""
        Me.DataGrid3.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid3.Location = New System.Drawing.Point(16, 8)
        Me.DataGrid3.Name = "DataGrid3"
        Me.DataGrid3.Size = New System.Drawing.Size(224, 234)
        Me.DataGrid3.TabIndex = 22
        '
        'mnuMain
        '
        Me.mnuMain.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem8, Me.MenuItem5})
        '
        'MenuItem1
        '
        Me.MenuItem1.Enabled = False
        Me.MenuItem1.Index = 0
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem2, Me.MenuItem3, Me.MenuItem4})
        Me.MenuItem1.Text = "File"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 0
        Me.MenuItem2.Text = "Open"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 1
        Me.MenuItem3.Text = "Save"
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 2
        Me.MenuItem4.Text = "Print"
        '
        'MenuItem8
        '
        Me.MenuItem8.Enabled = False
        Me.MenuItem8.Index = 1
        Me.MenuItem8.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem9, Me.MenuItem10, Me.MenuItem11})
        Me.MenuItem8.Text = "Edit"
        '
        'MenuItem9
        '
        Me.MenuItem9.Index = 0
        Me.MenuItem9.Text = "Cut"
        '
        'MenuItem10
        '
        Me.MenuItem10.Index = 1
        Me.MenuItem10.Text = "Copy"
        '
        'MenuItem11
        '
        Me.MenuItem11.Index = 2
        Me.MenuItem11.Text = "Paste"
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 2
        Me.MenuItem5.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem6, Me.MenuItem7})
        Me.MenuItem5.Text = "Help"
        '
        'MenuItem6
        '
        Me.MenuItem6.Enabled = False
        Me.MenuItem6.Index = 0
        Me.MenuItem6.Text = "Contents"
        '
        'MenuItem7
        '
        Me.MenuItem7.Index = 1
        Me.MenuItem7.Text = "About OSRE"
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(794, 547)
        Me.Controls.Add(Me.TabControl1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Menu = Me.mnuMain
        Me.Name = "frmMain"
        Me.Text = "Open Source Risk Engine"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.tbgHazard.ResumeLayout(False)
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tbgVulnerability.ResumeLayout(False)
        CType(Me.DataGrid2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tbgDamage.ResumeLayout(False)
        CType(Me.DataGrid3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub FrmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '*******************************************************
        '*** Adding the General info in the intro Tab
        '*******************************************************

        lblIntro.Text = "This project combines several emerging trends in natural hazards and " & _
        "software development. Increasingly, economic justification is required for decisionmaking " & _
        "on natural hazards mitigation projects. Therefore, natural hazards loss estimation software, " & _
        "such as HAZUS and CATS in the USA, LessLoss in the EU, and EXTREMUM in Russia, are in various " & _
        "stages of development. A common property of these programs is their closed nature, the source " & _
        "code is not open. Open Source code is a new trend in software development in general, and " & _
        "open source organizations are at the center of a broad thriving movement providing open tools " & _
        "for software development and applications. In earthquakes, OpenSHA (Open Seismic Hazard Analysis,) " & _
        "and the Open System for Earthquake Engineering Simulation (OpenSees) are projects which provide " & _
        "the hazard, and structural modeling modules for loss estimation, respectively. Currently there is " & _
        "a lack of an Open Source Risk Engine. OSRE aims to fill this void."

        lblAboutOSRE.Text = "OSRE was originally developed in Kyoto University, Graduate School of Engineering, " & _
        "Department of Urban Management, by " & ControlChars.CrLf & ControlChars.CrLf & _
        "Christakis Mina, Masaki Higuchi, Puay How Tion and Koichiro Danno, " & ControlChars.CrLf & ControlChars.CrLf & _
        "in order to meet the requirements of the `Capstone Project' for the Masters Degree in Engineering, " & _
        "under the supervision of" & ControlChars.CrLf & ControlChars.CrLf & _
        "Professors Charles Scawthorn, Junji Kiyono and Yusuke Ono."

        '*******************************************************
        '*** HAZARD TAB INITIALIZATION
        '*******************************************************

        ' * Initializing Hazard Table
        dtTable = dtSet.Tables.Add("t1")
        pkCol = dtTable.Columns.Add("Hmed", Type.GetType("System.Decimal"))
        dtTable.Columns.Add("Pexc", Type.GetType("System.Decimal"))
        DataGrid1.SetDataBinding(dtSet, "t1")
        ' * Default Hazard Equation
        txtHazardEqn.Text = "10^(-0.778-2.222*H)"
        PlotSurface2D1.Clear()
        PlotSurface2D1.Title = "HAZARD PLOT"
        PlotSurface2D1.BackColor = Color.Empty

        '*******************************************************
        '*** VULNERABILITY TAB INITIALIZATION
        '*******************************************************

        ' * Initializing Vulnerability Table
        dtTable = dtSet.Tables.Add("t2")
        pkCol = dtTable.Columns.Add("Hmed", Type.GetType("System.Decimal"))
        dtTable.Columns.Add("Dmed", Type.GetType("System.Decimal"))
        DataGrid2.SetDataBinding(dtSet, "t2")
        ' * Default Hazard Equation
        txtVulnerabilityEqn.Text = "-2*H^3+3*H^2"
        PlotSurface2D2.Clear()
        PlotSurface2D2.Title = "VULNERABILITY PLOT"
        PlotSurface2D2.BackColor = Color.Empty

        '*******************************************************
        '*** DAMAGE TAB INITIALIZATION
        '*******************************************************

        ' * Initializing Damage Table
        dtTable = dtSet.Tables.Add("t3")
        pkCol = dtTable.Columns.Add("Dmg", Type.GetType("System.Decimal"))
        dtTable.Columns.Add("P(dmg>d)", Type.GetType("System.Decimal"))
        DataGrid3.SetDataBinding(dtSet, "t3")
        DataGrid3.Enabled = False
        PlotSurface2D3.Clear()
        PlotSurface2D3.Title = "DAMAGE PLOT"
        PlotSurface2D3.BackColor = Color.Empty



    End Sub

    Private Sub FrmMain_Closed(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        '******* Terminate the program even with other forms loaded
        Application.Exit()
        'End

    End Sub


    Private Sub txtGenerateData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtGenerateData.Click
        '*******************************************************
        '*** HAZARD TAB DATA GENERATION
        '*******************************************************

        ' Store the primitives.
        m_Primatives = New Hashtable

        Dim X, yyy As Decimal
        Dim expr As String = txtHazardEqn.Text
        dtTable = dtSet.Tables("t1")

        dtTable.Clear()

        For X = 0.0 To 2.0 Step 0.002
            pkRow = dtTable.NewRow()
            pkRow("Hmed") = X
            m_Primatives.Add("H", X)

            Try
                pkRow("Pexc") = Val(EvaluateExpression(expr).ToString)
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                Exit For
            End Try
            dtTable.Rows.Add(pkRow)
            m_Primatives.Remove("H")

        Next X

    End Sub

    Private Sub cmdPlot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPlot.Click
        '*******************************************************
        '*** HAZARD PLOT
        '*******************************************************

        Dim lp As NPlot.LinePlot = New NPlot.LinePlot
        Dim pp As NPlot.PointPlot = New NPlot.PointPlot
        Dim yyy

        Dim XDAT As ArrayList = New ArrayList
        Dim YDAT As ArrayList = New ArrayList

        XDAT.Clear()
        YDAT.Clear()
        PlotSurface2D1.Clear()
        PlotSurface2D1.Title = "HAZARD PLOT"
        PlotSurface2D1.BackColor = Color.Empty

        For yyy = 0 To (dtSet.Tables("t1").Rows.Count() - 1)
            XDAT.Add(dtSet.Tables("t1").Rows(yyy)(0))
            YDAT.Add(dtSet.Tables("t1").Rows(yyy)(1))
        Next

        lp.AbscissaData = XDAT
        lp.DataSource = YDAT
        'pp.AbscissaData = XDAT     ' Plot Data Points
        'pp.DataSource = YDAT
        lp.Color = Color.Blue
        PlotSurface2D1.Add(lp)
        'PlotSurface2D1.Add(pp)
        PlotSurface2D1.XAxis1.Label = "Hazard"
        PlotSurface2D1.YAxis1.Label = "Pr. of exceed."

        PlotSurface2D1.Refresh()

    End Sub

    Private Sub cmdGenVulnData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGenVulnData.Click
        '*******************************************************
        '*** VULNERABILITY TAB DATA GENERATION
        '*******************************************************

        ' Store the primitives.
        m_Primatives = New Hashtable

        Dim X, yyy As Decimal
        Dim expr As String = txtVulnerabilityEqn.Text
        dtTable = dtSet.Tables("t2")

        dtTable.Clear()

        For X = 0.0 To 1.0 Step 0.001
            pkRow = dtTable.NewRow()
            pkRow("Hmed") = X
            m_Primatives.Add("H", X)

            Try
                pkRow("Dmed") = Val(EvaluateExpression(expr).ToString)
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                Exit For
            End Try
            dtTable.Rows.Add(pkRow)
            m_Primatives.Remove("H")

        Next X

    End Sub

    Private Sub cmdPlotVuln_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPlotVuln.Click
        '*******************************************************
        '*** VULNERABILITY PLOT
        '*******************************************************

        Dim lp As NPlot.LinePlot = New NPlot.LinePlot
        Dim pp As NPlot.PointPlot = New NPlot.PointPlot
        Dim yyy

        Dim XDAT As ArrayList = New ArrayList
        Dim YDAT As ArrayList = New ArrayList

        XDAT.Clear()
        YDAT.Clear()
        PlotSurface2D2.Clear()
        PlotSurface2D2.Title = "VULNERABILITY PLOT"
        PlotSurface2D2.BackColor = Color.Empty

        For yyy = 0 To (dtSet.Tables("t2").Rows.Count() - 1)
            XDAT.Add(dtSet.Tables("t2").Rows(yyy)(0))
            YDAT.Add(dtSet.Tables("t2").Rows(yyy)(1))
        Next

        lp.AbscissaData = XDAT
        lp.DataSource = YDAT
        'pp.AbscissaData = XDAT     ' Plot Data Points
        'pp.DataSource = YDAT
        lp.Color = Color.Blue
        PlotSurface2D2.Add(lp)
        'PlotSurface2D2.Add(pp)
        PlotSurface2D2.XAxis1.Label = "Hazard"
        PlotSurface2D2.YAxis1.Label = "Damage"

        PlotSurface2D2.Refresh()
    End Sub

    Private Sub cmdGenDamData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGenDamData.Click
        '****************************************
        '*** DAMAGE TAB DATA CALCULATION
        '****************************************

        Dim S, lamda, zeta, EPS, CDF As Double
        Dim A, B, zzz, yyy, www, tempsum, totsum As Double
        Dim i, j, k, p As Integer
        Dim AA, BB, XX As Double
        Dim Xm, Var, MAXVar, COV As Double
        Dim tmp1, tmp2, Duration As Integer
        Dim StartTime, EndTime As DateTime

        '=======================================
        '=== Hazard Calculations
        '=======================================

        Select Case cboHazardDistrType.Text

            Case "Lognormal"
                '*** LOGNORMAL Distribution for HAZARD
                Application.DoEvents()
                StartTime = Now()
                txtProcess.Text = Format(Now(), "hh:mm:ss") & " Calculating Hazard Matrix"
                txtProcess.Refresh()
                zeta = txtHazardDistrParam2.Text   'LN parameter 2
                prbTotProgress.Value = 0
                Application.DoEvents()

                For i = 0 To 1000
                    yyy = 2 * i / 1000
                    If yyy = 0 Then
                        lamda = Math.Log(1.0E-130)
                    Else
                        lamda = Math.Log(yyy)
                    End If
                    prbProgress.Value = 0 'Initializing Progress Bar
                    prbTotProgress.Value = i / 10
                    Application.DoEvents()
                    prbTotProgress.Refresh()

                    For k = 0 To 1000
                        zzz = 2 * k / 1000
                        If zzz = 0 Then
                            CDF = 0.5 + 0.5 * ErrF((Math.Log(1.0E-130) - lamda) / (zeta * Math.Sqrt(2)))
                        Else
                            CDF = 0.5 + 0.5 * ErrF((Math.Log(zzz) - lamda) / (zeta * Math.Sqrt(2)))
                        End If
                        If CDF > 1 Then CDF = 1
                        HAZARDMATRIX(i, k) = CDF
                        prbProgress.Value = k / 10
                        prbProgress.Refresh()
                    Next

                Next

                '*** Sweeping for Confidence intervals
                txtProcess.Text = txtProcess.Text & ControlChars.CrLf & Format(Now(), "hh:mm:ss") & " Extracting Confidence intervals"
                txtProcess.Refresh()

                '*** Eddie's Method
                Application.DoEvents()
                For i = 0 To 100
                    prbProgress.Value = 0 'Initializing Progress Bar
                    prbTotProgress.Value = i
                    prbTotProgress.Refresh()
                    yyy = i / 100

                    For j = 0 To 1000
                        p = 0
                        prbProgress.Value = j / 10
                        prbProgress.Refresh()

                        For k = 0 To 1000
                            If Math.Abs(HAZARDMATRIX(p, j) - yyy) > Math.Abs(HAZARDMATRIX(k, j) - yyy) Then
                                p = k
                            End If
                        Next
                        HAZCONFINT(i, j) = dtSet.Tables("t1").Rows(p)(1)
                        HAZCONFINTCDF(i, j) = 1 - HAZCONFINT(i, j)

                    Next
                Next

            Case "Beta"
                '*** BETA Distribution for HAZARD
                txtProcess.Text = Format(Now(), "hh:mm:ss") & " Calculating Hazard Matrix"
                txtProcess.Refresh()

                COV = Val(txtHazardDistrParam2.Text)

                For i = 0 To 1000
                    Xm = i / 1000

                    Select Case Xm

                        Case 0
                            AA = 0.5    'Param a
                            BB = 0.5    'Param b
                            'Var = COV * 1.0E-130
                            'MAXVar = Xm * (1 - Xm)
                            'If Var >= MAXVar Then Var = MAXVar * 0.9999999999999
                            'AA = 1.0E-130 * ((1.0E-130 / Var) * (1.0 - 1.0E-130) - 1.0)  'Param a
                            'BB = AA * (1.0 - 1.0E-130) / 1.0E-130  'Param b

                        Case 1
                            AA = 0.5    'Param a
                            BB = 0.5    'Param b

                        Case Else
                            Var = COV * Xm
                            MAXVar = Xm * (1 - Xm)
                            If Var >= MAXVar Then Var = MAXVar * 0.999
                            AA = Xm * ((Xm / Var) * (1.0 - Xm) - 1.0)     'Param a
                            BB = AA * (1.0 - Xm) / Xm                    'Param b
                    End Select

                    prbProgress.Value = 0 'Initializing Progress Bar
                    prbTotProgress.Value = i / 10
                    prbTotProgress.Refresh()

                    For k = 0 To 1000
                        XX = k / 1000
                        CDF = BetaI(AA, BB, XX)
                        If CDF > 1 Then CDF = 1
                        HAZARDMATRIX(i, k) = CDF
                        prbProgress.Value = k / 10
                        prbProgress.Refresh()
                    Next
                Next

                '*** Sweeping for Confidence intervals
                txtProcess.Text = txtProcess.Text & ControlChars.CrLf & Format(Now(), "hh:mm:ss") & " Extracting Confidence intervals"
                txtProcess.Refresh()

                '*** Eddie's Method
                Application.DoEvents()
                For i = 0 To 100
                    prbProgress.Value = 0 'Initializing Progress Bar
                    prbTotProgress.Value = i
                    prbTotProgress.Refresh()
                    yyy = i / 100

                    For j = 0 To 1000
                        p = 0
                        prbProgress.Value = j / 10
                        prbProgress.Refresh()

                        For k = 0 To 1000
                            If Math.Abs(HAZARDMATRIX(p, j) - yyy) > Math.Abs(HAZARDMATRIX(k, j) - yyy) Then
                                p = k
                            End If
                        Next
                        HAZCONFINT(i, j) = dtSet.Tables("t1").Rows(p)(1)
                        HAZCONFINTCDF(i, j) = 1 - HAZCONFINT(i, j)
                    Next
                Next
        End Select

        '----------------------------------------------------------
        '--- Calculating PDF of Confidence Intervals for HAZARD
        '----------------------------------------------------------
        txtProcess.Text = txtProcess.Text & ControlChars.CrLf & Format(Now(), "hh:mm:ss") & " Calculating PDF of Conf. Intervals"
        txtProcess.Refresh()

        For i = 0 To 100
            prbProgress.Value = 0 'Initializing Progress Bar
            prbTotProgress.Value = i
            prbTotProgress.Refresh()

            For j = 0 To 1000
                prbProgress.Value = j / 10
                prbProgress.Refresh()
                If j = 0 Then
                    'HAZCONFINTPDF(i, j) = HAZCONFINTCDF(i, j)
                    HAZCONFINTPDF(i, j) = 0 'Screws up the results!

                Else
                    HAZCONFINTPDF(i, j) = HAZCONFINTCDF(i, j) - HAZCONFINTCDF(i, j - 1)
                End If
            Next
        Next

        '=======================================
        '=== Vulnerability Calculations
        '=======================================

        Select Case cboVulnerabilityDistrType.Text

            Case "Lognormal"
                '*** LOGNORMAL Distribution for VULNERABILITY
                Application.DoEvents()
                txtProcess.Text = txtProcess.Text & ControlChars.CrLf & Format(Now(), "hh:mm:ss") & " Calculating Vulnerability Matrix"
                txtProcess.Refresh()
                Val(txtVulnerabilityDistrParam2.Text)   'LN parameter 2
                prbTotProgress.Value = 0
                Application.DoEvents()

                For i = 0 To 1000
                    yyy = i / 1000
                    If yyy = 0 Then
                        lamda = Math.Log(1.0E-130)
                    Else
                        lamda = Math.Log(yyy)
                    End If
                    prbProgress.Value = 0 'Initializing Progress Bar
                    prbTotProgress.Value = i / 10
                    Application.DoEvents()
                    prbTotProgress.Refresh()

                    For k = 0 To 1000
                        zzz = k / 1000
                        If zzz = 0 Then
                            CDF = 0.5 + 0.5 * ErrF((Math.Log(1.0E-130) - lamda) / (zeta * Math.Sqrt(2)))
                        Else
                            CDF = 0.5 + 0.5 * ErrF((Math.Log(zzz) - lamda) / (zeta * Math.Sqrt(2)))
                        End If
                        If CDF > 1 Then CDF = 1
                        VULNMATRIX(i, k) = 1 - CDF  'This is only for vulnerability!
                        prbProgress.Value = k / 10
                        prbProgress.Refresh()
                    Next

                Next


            Case "Beta"
                '*** BETA Distribution for VULNERABILITY
                Application.DoEvents()
                txtProcess.Text = txtProcess.Text & ControlChars.CrLf & Format(Now(), "hh:mm:ss") & " Calculating Vulnerability Matrix"
                txtProcess.Refresh()

                COV = Val(txtVulnerabilityDistrParam2.Text)

                For i = 0 To 1000
                    Xm = dtSet.Tables("t2").Rows(i)(1)
                    'Xm = i / 1000
                    Select Case Xm
                        Case 0
                            AA = 0.5    'Param a
                            BB = 0.5    'Param b
                            'Var = COV * 1.0E-130
                            'AA = 1.0E-130 * ((1.0E-130 / Var) * (1.0 - 1.0E-130) - 1.0)  'Param a
                            'BB = AA * (1.0 - 1.0E-130) / 1.0E-130 
                            'Param b

                        Case 1
                            AA = 0.5    'Param a
                            BB = 0.5    'Param b

                        Case Else
                            Var = COV * Xm
                            MAXVar = Xm * (1 - Xm)
                            If Var >= MAXVar Then Var = MAXVar * 0.999
                            AA = Xm * ((Xm / Var) * (1.0 - Xm) - 1.0)     'Param a
                            BB = AA * (1.0 - Xm) / Xm                    'Param b
                    End Select
                    prbProgress.Value = 0 'Initializing Progress Bar
                    prbTotProgress.Value = i / 10
                    prbTotProgress.Refresh()

                    For k = 0 To 1000
                        XX = dtSet.Tables("t2").Rows(k)(1)
                        'XX = k / 1000
                        CDF = BetaI(AA, BB, XX)
                        If CDF > 1 Then CDF = 1
                        VULNMATRIX(i, k) = 1 - CDF  'This is only for vulnerability!
                        prbProgress.Value = k / 10
                        prbProgress.Refresh()
                    Next
                Next

        End Select

        '=======================================
        '=== Multiplication of Matrices
        '=======================================
        Application.DoEvents()
        txtProcess.Text = txtProcess.Text & ControlChars.CrLf & Format(Now(), "hh:mm:ss") & " Calculating Risk Matrix"
        txtProcess.Refresh()

        For i = 0 To 100
            prbProgress.Value = 0 'Initializing Progress Bar
            prbTotProgress.Value = i
            prbTotProgress.Refresh()

            For j = 0 To 1000
                prbProgress.Value = j / 10
                prbProgress.Refresh()
                totsum = 0
                For k = 0 To 1000
                    tempsum = HAZCONFINTPDF(i, k) * VULNMATRIX(k, j)
                    totsum = totsum + tempsum
                Next
                RISKMATRIX(i, j) = totsum
            Next
        Next

        'Fill the Damage Table
        Select Case cboConfidence.Text
            Case "1%"
                FillTable(1)
            Case "10%"
                FillTable(10)
            Case "20%"
                FillTable(20)
            Case "30%"
                FillTable(30)
            Case "40%"
                FillTable(40)
            Case "50%"
                FillTable(50)
            Case "60%"
                FillTable(60)
            Case "70%"
                FillTable(70)
            Case "80%"
                FillTable(80)
            Case "90%"
                FillTable(90)
            Case "95%"
                FillTable(95)
            Case "99%"
                FillTable(99)
            Case Else
                '*** Something's wrong?
        End Select

        EndTime = Now()
        Duration = EndTime.Subtract(StartTime).Minutes()
        txtProcess.Text = txtProcess.Text & ControlChars.CrLf & Format(Now(), "hh:mm:ss") & " Completed in " & Duration & " minutes."
        txtProcess.Refresh()

    End Sub

    Private Sub cmdAddGraph_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddGraph.Click
        '****************************************
        '*** DAMAGE PLOT
        '****************************************

        Select Case cboConfidence.Text
            Case "1%"
                FillTable(1)
            Case "10%"
                FillTable(10)
            Case "20%"
                FillTable(20)
            Case "30%"
                FillTable(30)
            Case "40%"
                FillTable(40)
            Case "50%"
                FillTable(50)
            Case "60%"
                FillTable(60)
            Case "70%"
                FillTable(70)
            Case "80%"
                FillTable(80)
            Case "90%"
                FillTable(90)
            Case "95%"
                FillTable(95)
            Case "99%"
                FillTable(99)
            Case Else
                '*** Something's wrong?
        End Select

        Dim lp As NPlot.LinePlot = New NPlot.LinePlot
        Dim pp As NPlot.PointPlot = New NPlot.PointPlot
        Dim yyy

        Dim XDAT As ArrayList = New ArrayList
        Dim YDAT As ArrayList = New ArrayList

        XDAT.Clear()
        YDAT.Clear()
        'PlotSurface2D3.Clear()
        PlotSurface2D3.Title = "DAMAGE PLOT"
        PlotSurface2D3.BackColor = Color.Empty

        For yyy = 0 To (dtSet.Tables("t3").Rows.Count() - 1)
            XDAT.Add(dtSet.Tables("t3").Rows(yyy)(0))
            YDAT.Add(dtSet.Tables("t3").Rows(yyy)(1))
        Next

        lp.AbscissaData = XDAT
        lp.DataSource = YDAT
        'pp.AbscissaData = XDAT
        'pp.DataSource = YDAT
        lp.Color = Color.FromArgb(255, Int(Rnd() * 255), Int(Rnd() * 255), Int(Rnd() * 255))
        'lp.Color = Color.Blue
        PlotSurface2D3.Add(lp)
        'PlotSurface2D3.Add(pp)
        PlotSurface2D3.XAxis1.Label = "Damage"
        PlotSurface2D3.YAxis1.Label = "P(dmg>d)"

        PlotSurface2D3.Refresh()
    End Sub

    Sub FillTable(ByVal ConfInterval)
        '****************************************
        '*** FILLS THE DAMAGE TABLE
        '****************************************

        Dim k As Integer

        dtTable = dtSet.Tables("t3")
        dtTable.Clear()
        For k = 0 To 1000
            pkRow = dtTable.NewRow()
            pkRow("Dmg") = (dtSet.Tables("t2").Rows(k)(1))
            pkRow("P(dmg>d)") = RISKMATRIX(ConfInterval, k)
            dtTable.Rows.Add(pkRow)
        Next
        DataGrid3.Enabled = True

    End Sub

    Private Sub cmdPasteHzdData_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPasteHzdData.Click
        '*****************************************************
        '*** PASTE DATA FROM CLIPBOARD TO HAZARD TABLE
        '*****************************************************

        Dim CLP As Clipboard
        dtTable = dtSet.Tables("t1")
        dtTable.Clear()

        Dim dr As DataRow
        Dim MyData As String
        Dim intCount As Integer

        MyData = CLP.GetDataObject.GetData(GetType(String))

        Dim strArr() As String = MyData.Split(ControlChars.CrLf)

        intCount = strArr.GetUpperBound(0)

        Try
            For i As Integer = 0 To intCount - 1
                dr = dtTable.NewRow()
                Dim tempStr() As String = strArr(i).Split(ControlChars.Tab)
                dr("Hmed") = Val(tempStr(0))
                dr("Pexc") = Val(tempStr(1))
                dtTable.Rows.Add(dr)
            Next
        Catch ex As Exception
            MessageBox.Show("Invalid Data!" & ControlChars.CrLf & "Only 2 columns of data separated by 'tabs' are accepted." _
                             & ControlChars.CrLf & "Data copied directly from Excel can be used.")
        End Try

    End Sub

    Private Sub cmdClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClear.Click
        '*****************************************************
        '*** CLEAR THE DAMAGE PLOT
        '*****************************************************

        Dim lp As NPlot.LinePlot = New NPlot.LinePlot
        Dim pp As NPlot.PointPlot = New NPlot.PointPlot
        Dim yyy

        Dim XDAT As ArrayList = New ArrayList
        Dim YDAT As ArrayList = New ArrayList

        XDAT.Clear()
        YDAT.Clear()
        PlotSurface2D3.Clear()
        PlotSurface2D3.Title = "DAMAGE PLOT"
        PlotSurface2D3.Refresh()
    End Sub

    Private Sub cmdPasteVulnData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPasteVulnData.Click
        '*****************************************************
        '*** PASTE DATA FROM CLIPBOARD TO VULNERABILITY TABLE
        '*****************************************************

        Dim CLP As Clipboard
        dtTable = dtSet.Tables("t2")
        dtTable.Clear()

        Dim dr As DataRow
        Dim MyData As String
        Dim intCount As Integer

        MyData = CLP.GetDataObject.GetData(GetType(String))

        Dim strArr() As String = MyData.Split(ControlChars.CrLf)

        intCount = strArr.GetUpperBound(0)

        Try
            For i As Integer = 0 To intCount - 1
                dr = dtTable.NewRow()
                Dim tempStr() As String = strArr(i).Split(ControlChars.Tab)
                dr("Hmed") = Val(tempStr(0))
                dr("Dmed") = Val(tempStr(1))
                dtTable.Rows.Add(dr)
            Next
        Catch ex As Exception
            MessageBox.Show("Invalid Data!" & ControlChars.CrLf & "Only 2 columns of data separated by 'tabs' are accepted." _
                             & ControlChars.CrLf & "Data copied directly from Excel can be used.")
        End Try

    End Sub

    Private Sub cboHazardDistrType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboHazardDistrType.SelectedIndexChanged
        '*****************************************************
        '*** UPDATES THE LABEL FOR THE HAZARD DISTR. PARAM.
        '*****************************************************

        Select Case cboHazardDistrType.Text
            Case "Lognormal"
                lblHazardDistrParam2.Text = "Zeta"
            Case "Beta"
                lblHazardDistrParam2.Text = "COV"
        End Select
    End Sub

    Private Sub cboVulnerabilityDistrType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboVulnerabilityDistrType.SelectedIndexChanged
        '*****************************************************
        '*** UPDATES THE LABEL FOR THE VULNER. DISTR. PARAM.
        '*****************************************************

        Select Case cboVulnerabilityDistrType.Text
            Case "Lognormal"
                lblVulnerabilityDistrParam2.Text = "Zeta"
            Case "Beta"
                lblVulnerabilityDistrParam2.Text = "COV"
        End Select
    End Sub

    Private Sub cmdFitHazCurve_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFitHazCurve.Click
        goForm2.Show()
        Me.Hide() 'Hide Form1
        goForm2.lblTitle.Text = "Curve Fitting For Hazard"

    End Sub

    Private Sub cmdFitVulnCurve_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFitVulnCurve.Click
        goForm2.Show()
        Me.Hide() 'Hide Form1
        goForm2.lblTitle.Text = "Curve Fitting For Vulnerability"

    End Sub

    Private Sub MenuItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem7.Click
        goForm3.Show()
        
    End Sub


End Class
