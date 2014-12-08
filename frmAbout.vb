Public Class frmAbout
    Inherits System.Windows.Forms.Form

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
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents lblTitle As System.Windows.Forms.Label
    Friend WithEvents llbOSRE As System.Windows.Forms.LinkLabel
    Friend WithEvents lblLicense As System.Windows.Forms.Label
    Friend WithEvents lblVer As System.Windows.Forms.Label
    Friend WithEvents llbLicense As System.Windows.Forms.LinkLabel
    Friend WithEvents lblNplot As System.Windows.Forms.Label
    Friend WithEvents llbNplot As System.Windows.Forms.LinkLabel
    Friend WithEvents lblNplot2 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmAbout))
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.btnOK = New System.Windows.Forms.Button
        Me.lblTitle = New System.Windows.Forms.Label
        Me.llbOSRE = New System.Windows.Forms.LinkLabel
        Me.lblLicense = New System.Windows.Forms.Label
        Me.llbLicense = New System.Windows.Forms.LinkLabel
        Me.lblVer = New System.Windows.Forms.Label
        Me.lblNplot = New System.Windows.Forms.Label
        Me.llbNplot = New System.Windows.Forms.LinkLabel
        Me.lblNplot2 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(8, 8)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(136, 136)
        Me.PictureBox1.TabIndex = 0
        Me.PictureBox1.TabStop = False
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(8, 160)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(136, 24)
        Me.btnOK.TabIndex = 1
        Me.btnOK.Text = "OK, I read everything!"
        '
        'lblTitle
        '
        Me.lblTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblTitle.Location = New System.Drawing.Point(160, 8)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(256, 32)
        Me.lblTitle.TabIndex = 2
        Me.lblTitle.Text = "Open Source Risk Engine"
        Me.lblTitle.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'llbOSRE
        '
        Me.llbOSRE.Location = New System.Drawing.Point(160, 168)
        Me.llbOSRE.Name = "llbOSRE"
        Me.llbOSRE.Size = New System.Drawing.Size(256, 16)
        Me.llbOSRE.TabIndex = 3
        Me.llbOSRE.TabStop = True
        Me.llbOSRE.Text = "http://quake.kuciv.kyoto-u.ac.jp/OSRE/"
        Me.llbOSRE.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblLicense
        '
        Me.lblLicense.Location = New System.Drawing.Point(160, 88)
        Me.lblLicense.Name = "lblLicense"
        Me.lblLicense.Size = New System.Drawing.Size(256, 16)
        Me.lblLicense.TabIndex = 4
        Me.lblLicense.Text = "OSRE is licensed under the Open Source"
        Me.lblLicense.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'llbLicense
        '
        Me.llbLicense.Location = New System.Drawing.Point(160, 112)
        Me.llbLicense.Name = "llbLicense"
        Me.llbLicense.Size = New System.Drawing.Size(256, 16)
        Me.llbLicense.TabIndex = 5
        Me.llbLicense.TabStop = True
        Me.llbLicense.Text = "Academic Free License V. 2.1"
        Me.llbLicense.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblVer
        '
        Me.lblVer.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblVer.Location = New System.Drawing.Point(168, 48)
        Me.lblVer.Name = "lblVer"
        Me.lblVer.Size = New System.Drawing.Size(248, 24)
        Me.lblVer.TabIndex = 6
        Me.lblVer.Text = "Version"
        Me.lblVer.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblNplot
        '
        Me.lblNplot.Location = New System.Drawing.Point(176, 136)
        Me.lblNplot.Name = "lblNplot"
        Me.lblNplot.Size = New System.Drawing.Size(96, 16)
        Me.lblNplot.TabIndex = 7
        Me.lblNplot.Text = "OSRE is using the"
        Me.lblNplot.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'llbNplot
        '
        Me.llbNplot.Location = New System.Drawing.Point(272, 136)
        Me.llbNplot.Name = "llbNplot"
        Me.llbNplot.Size = New System.Drawing.Size(40, 16)
        Me.llbNplot.TabIndex = 8
        Me.llbNplot.TabStop = True
        Me.llbNplot.Text = "NPlot"
        Me.llbNplot.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblNplot2
        '
        Me.lblNplot2.Location = New System.Drawing.Point(304, 136)
        Me.lblNplot2.Name = "lblNplot2"
        Me.lblNplot2.Size = New System.Drawing.Size(96, 16)
        Me.lblNplot2.TabIndex = 9
        Me.lblNplot2.Text = "Graphing Library."
        Me.lblNplot2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'frmAbout
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(426, 200)
        Me.Controls.Add(Me.lblNplot2)
        Me.Controls.Add(Me.llbNplot)
        Me.Controls.Add(Me.lblNplot)
        Me.Controls.Add(Me.lblVer)
        Me.Controls.Add(Me.llbLicense)
        Me.Controls.Add(Me.lblLicense)
        Me.Controls.Add(Me.llbOSRE)
        Me.Controls.Add(Me.lblTitle)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.PictureBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmAbout"
        Me.Text = "About OSRE"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        goForm3.Hide()
    End Sub

    Private Sub llbOSRE_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbOSRE.LinkClicked
        llbOSRE.LinkVisited = True
        System.Diagnostics.Process.Start("http://quake.kuciv.kyoto-u.ac.jp/OSRE/")
    End Sub

    Private Sub frmAbout_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        lblVer.Text = "Ver: "
        With System.Reflection.Assembly.GetExecutingAssembly.GetName().Version
            lblVer.Text += .Major.ToString() & "."
            lblVer.Text += .Minor.ToString() & "."
            lblVer.Text += .Build.ToString()
        End With

    End Sub

    Private Sub llbLicense_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbLicense.LinkClicked
        llbLicense.LinkVisited = True
        System.Diagnostics.Process.Start("http://opensource.org/licenses/afl-2.1.php")
    End Sub

    Private Sub llbNplot_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbNplot.LinkClicked
        llbNplot.LinkVisited = True
        System.Diagnostics.Process.Start("http://www.nplot.com/")
    End Sub

    Private Sub llbOSRE_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles llbOSRE.MouseEnter
        llbOSRE.BackColor = Color.Yellow
    End Sub

    Private Sub llbOSRE_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles llbOSRE.MouseLeave
        llbOSRE.BackColor = Color.Empty
    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click
        llbOSRE.LinkVisited = True
        System.Diagnostics.Process.Start("http://quake.kuciv.kyoto-u.ac.jp/OSRE/")
    End Sub

    Private Sub PictureBox1_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox1.MouseEnter
        'PictureBox1.BackColor = Color.Yellow
        PictureBox1.Cursor = Cursors.Hand
    End Sub

    Private Sub PictureBox1_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox1.MouseLeave
        'PictureBox1.BackColor = Color.Empty
        PictureBox1.Cursor = Cursors.Arrow
    End Sub

End Class
