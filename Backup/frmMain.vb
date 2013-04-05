Public Class frmMain
    Inherits System.Windows.Forms.Form
    Dim DBConnection As New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source = winsteelVers5.mdb")
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
    Friend WithEvents btnJobRate As System.Windows.Forms.Button
    Friend WithEvents btnCutSheet As System.Windows.Forms.Button
    Friend WithEvents btnCompany As System.Windows.Forms.Button
    Friend WithEvents btnWeight As System.Windows.Forms.Button
    Friend WithEvents btnContractor As System.Windows.Forms.Button
    Friend WithEvents btnJob As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents lblHeading As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnReportPrint As System.Windows.Forms.Button
    Friend WithEvents btnInvoices As System.Windows.Forms.Button
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents btnCheckInv As System.Windows.Forms.Button
    Friend WithEvents btnArchive As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btnJobRate = New System.Windows.Forms.Button
        Me.btnCutSheet = New System.Windows.Forms.Button
        Me.btnCompany = New System.Windows.Forms.Button
        Me.btnWeight = New System.Windows.Forms.Button
        Me.btnContractor = New System.Windows.Forms.Button
        Me.btnJob = New System.Windows.Forms.Button
        Me.btnReportPrint = New System.Windows.Forms.Button
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.lblHeading = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.btnCheckInv = New System.Windows.Forms.Button
        Me.btnInvoices = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.btnArchive = New System.Windows.Forms.Button
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.btnExit = New System.Windows.Forms.Button
        Me.Panel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnJobRate
        '
        Me.btnJobRate.Location = New System.Drawing.Point(24, 120)
        Me.btnJobRate.Name = "btnJobRate"
        Me.btnJobRate.Size = New System.Drawing.Size(152, 40)
        Me.btnJobRate.TabIndex = 1
        Me.btnJobRate.Text = "Job Rate Maintenance"
        '
        'btnCutSheet
        '
        Me.btnCutSheet.Location = New System.Drawing.Point(32, 32)
        Me.btnCutSheet.Name = "btnCutSheet"
        Me.btnCutSheet.Size = New System.Drawing.Size(160, 40)
        Me.btnCutSheet.TabIndex = 1
        Me.btnCutSheet.Text = "Cutting Sheet Maintenance"
        '
        'btnCompany
        '
        Me.btnCompany.Location = New System.Drawing.Point(24, 24)
        Me.btnCompany.Name = "btnCompany"
        Me.btnCompany.Size = New System.Drawing.Size(152, 40)
        Me.btnCompany.TabIndex = 0
        Me.btnCompany.Text = "Company Maintenance"
        '
        'btnWeight
        '
        Me.btnWeight.Location = New System.Drawing.Point(40, 80)
        Me.btnWeight.Name = "btnWeight"
        Me.btnWeight.Size = New System.Drawing.Size(152, 40)
        Me.btnWeight.TabIndex = 3
        Me.btnWeight.Text = "Weight Maintenance"
        '
        'btnContractor
        '
        Me.btnContractor.Location = New System.Drawing.Point(24, 24)
        Me.btnContractor.Name = "btnContractor"
        Me.btnContractor.Size = New System.Drawing.Size(152, 40)
        Me.btnContractor.TabIndex = 2
        Me.btnContractor.Text = "Contractor Maintenance"
        '
        'btnJob
        '
        Me.btnJob.Location = New System.Drawing.Point(24, 168)
        Me.btnJob.Name = "btnJob"
        Me.btnJob.Size = New System.Drawing.Size(152, 40)
        Me.btnJob.TabIndex = 0
        Me.btnJob.Text = "Job Maintenance (Old)"
        '
        'btnReportPrint
        '
        Me.btnReportPrint.Location = New System.Drawing.Point(32, 200)
        Me.btnReportPrint.Name = "btnReportPrint"
        Me.btnReportPrint.Size = New System.Drawing.Size(160, 40)
        Me.btnReportPrint.TabIndex = 3
        Me.btnReportPrint.Text = "Reports && Printing"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.Controls.Add(Me.lblHeading)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(514, 88)
        Me.Panel1.TabIndex = 2
        '
        'lblHeading
        '
        Me.lblHeading.Font = New System.Drawing.Font("Times New Roman", 27.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHeading.ForeColor = System.Drawing.Color.RoyalBlue
        Me.lblHeading.Location = New System.Drawing.Point(88, 24)
        Me.lblHeading.Name = "lblHeading"
        Me.lblHeading.Size = New System.Drawing.Size(336, 40)
        Me.lblHeading.TabIndex = 0
        Me.lblHeading.Text = "Reinforcing Ability"
        Me.lblHeading.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btnCheckInv)
        Me.GroupBox1.Controls.Add(Me.btnCutSheet)
        Me.GroupBox1.Controls.Add(Me.btnReportPrint)
        Me.GroupBox1.Controls.Add(Me.btnInvoices)
        Me.GroupBox1.Location = New System.Drawing.Point(264, 104)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(232, 272)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Processing"
        '
        'btnCheckInv
        '
        Me.btnCheckInv.Location = New System.Drawing.Point(32, 144)
        Me.btnCheckInv.Name = "btnCheckInv"
        Me.btnCheckInv.Size = New System.Drawing.Size(160, 40)
        Me.btnCheckInv.TabIndex = 5
        Me.btnCheckInv.Text = "Invoicing Check"
        '
        'btnInvoices
        '
        Me.btnInvoices.Location = New System.Drawing.Point(32, 88)
        Me.btnInvoices.Name = "btnInvoices"
        Me.btnInvoices.Size = New System.Drawing.Size(160, 40)
        Me.btnInvoices.TabIndex = 4
        Me.btnInvoices.Text = "Invoicing"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.btnWeight)
        Me.GroupBox2.Controls.Add(Me.GroupBox3)
        Me.GroupBox2.Controls.Add(Me.GroupBox4)
        Me.GroupBox2.Location = New System.Drawing.Point(16, 104)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(232, 376)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.btnContractor)
        Me.GroupBox3.Controls.Add(Me.btnJobRate)
        Me.GroupBox3.Controls.Add(Me.btnArchive)
        Me.GroupBox3.Controls.Add(Me.btnJob)
        Me.GroupBox3.Location = New System.Drawing.Point(16, 144)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(200, 224)
        Me.GroupBox3.TabIndex = 4
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Job Details"
        '
        'btnArchive
        '
        Me.btnArchive.Location = New System.Drawing.Point(24, 72)
        Me.btnArchive.Name = "btnArchive"
        Me.btnArchive.Size = New System.Drawing.Size(152, 40)
        Me.btnArchive.TabIndex = 6
        Me.btnArchive.Text = "Job Maintenance"
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.btnCompany)
        Me.GroupBox4.Location = New System.Drawing.Point(16, 8)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(200, 128)
        Me.GroupBox4.TabIndex = 5
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Company Details"
        '
        'btnExit
        '
        Me.btnExit.Location = New System.Drawing.Point(296, 408)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(168, 40)
        Me.btnExit.TabIndex = 3
        Me.btnExit.Text = "Exit"
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(514, 504)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "frmMain"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Main"
        Me.Panel1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnJobRate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnJobRate.Click
        Dim Form As New frmJobRate(Me)
        Me.Hide()
        Form.Show()
    End Sub

    Private Sub btnCutSheet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCutSheet.Click
        Dim Form As New frmNewCutSheet(Me)
        Me.Hide()
        Form.Show()
    End Sub

    Private Sub btnCompany_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCompany.Click
        Dim Form As New frmCompany(Me)
        Me.Hide()
        Form.Show()
    End Sub

    Private Sub btnWeight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnWeight.Click
        Dim Form As New frmWeight(Me)
        Me.Hide()
        Form.Show()
    End Sub

    Private Sub btnContractor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnContractor.Click
        Dim Form As New frmContractor(Me)
        Me.Hide()
        Form.Show()
    End Sub

    Private Sub btnJob_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnJob.Click
        Dim Form As New frmJob(Me)
        Me.Hide()
        Form.Show()
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Close()
    End Sub

    Private Sub btnReportPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReportPrint.Click
        Dim Form As New frmReports(Me, DBConnection)
        Me.Hide()
        Form.Show()
    End Sub

    Private Sub btnInvoices_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInvoices.Click
        Dim form As frmInvoicing = New frmInvoicing(Me, DBConnection)
        Me.Hide()
        form.Show()
    End Sub

    Private Sub btnCheckInv_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCheckInv.Click
        Dim Form As New frmNotInvoiced(Me)
        Me.Hide()
        Form.Show()
    End Sub

    Private Sub Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint

    End Sub

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnArchive_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnArchive.Click

        Dim Form As frmJobArchive = New frmJobArchive(Me, DBConnection)
        Me.Hide()
        Form.Show()

    End Sub

    Private Sub GroupBox2_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox2.Enter

    End Sub
End Class
