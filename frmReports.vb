Public Class frmReports
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
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btnPrintCutInv As System.Windows.Forms.Button
    Friend WithEvents btnPrintReinforcingSummary As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents lblHeading As System.Windows.Forms.Label
    Friend WithEvents btnPrintCutSheet As System.Windows.Forms.Button
    Friend WithEvents btn_printSummaryOfBendingSchedule As System.Windows.Forms.Button
    Friend WithEvents btnInvSummary As System.Windows.Forms.Button
    Friend WithEvents btnJobSumRep As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents btnJobListing As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btnExit = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.btnJobListing = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.btnPrintCutSheet = New System.Windows.Forms.Button
        Me.btnPrintCutInv = New System.Windows.Forms.Button
        Me.btnPrintReinforcingSummary = New System.Windows.Forms.Button
        Me.btn_printSummaryOfBendingSchedule = New System.Windows.Forms.Button
        Me.btnInvSummary = New System.Windows.Forms.Button
        Me.btnJobSumRep = New System.Windows.Forms.Button
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.lblHeading = New System.Windows.Forms.Label
        Me.GroupBox1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnExit
        '
        Me.btnExit.Location = New System.Drawing.Point(168, 400)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(168, 40)
        Me.btnExit.TabIndex = 7
        Me.btnExit.Text = "Back to Main Menu"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btnJobListing)
        Me.GroupBox1.Controls.Add(Me.Button1)
        Me.GroupBox1.Controls.Add(Me.btnPrintCutSheet)
        Me.GroupBox1.Controls.Add(Me.btnPrintCutInv)
        Me.GroupBox1.Controls.Add(Me.btnPrintReinforcingSummary)
        Me.GroupBox1.Controls.Add(Me.btn_printSummaryOfBendingSchedule)
        Me.GroupBox1.Controls.Add(Me.btnInvSummary)
        Me.GroupBox1.Controls.Add(Me.btnJobSumRep)
        Me.GroupBox1.Location = New System.Drawing.Point(16, 96)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(488, 280)
        Me.GroupBox1.TabIndex = 5
        Me.GroupBox1.TabStop = False
        '
        'btnJobListing
        '
        Me.btnJobListing.Location = New System.Drawing.Point(40, 32)
        Me.btnJobListing.Name = "btnJobListing"
        Me.btnJobListing.Size = New System.Drawing.Size(168, 40)
        Me.btnJobListing.TabIndex = 5
        Me.btnJobListing.Text = "Print Job Listing"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(40, 96)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(168, 40)
        Me.Button1.TabIndex = 4
        Me.Button1.Text = "Print Job Rates"
        '
        'btnPrintCutSheet
        '
        Me.btnPrintCutSheet.Location = New System.Drawing.Point(40, 160)
        Me.btnPrintCutSheet.Name = "btnPrintCutSheet"
        Me.btnPrintCutSheet.Size = New System.Drawing.Size(168, 40)
        Me.btnPrintCutSheet.TabIndex = 2
        Me.btnPrintCutSheet.Text = "Print Cutting Sheet"
        '
        'btnPrintCutInv
        '
        Me.btnPrintCutInv.Location = New System.Drawing.Point(40, 216)
        Me.btnPrintCutInv.Name = "btnPrintCutInv"
        Me.btnPrintCutInv.Size = New System.Drawing.Size(168, 40)
        Me.btnPrintCutInv.TabIndex = 3
        Me.btnPrintCutInv.Text = "Print Invoice"
        '
        'btnPrintReinforcingSummary
        '
        Me.btnPrintReinforcingSummary.Location = New System.Drawing.Point(272, 216)
        Me.btnPrintReinforcingSummary.Name = "btnPrintReinforcingSummary"
        Me.btnPrintReinforcingSummary.Size = New System.Drawing.Size(168, 40)
        Me.btnPrintReinforcingSummary.TabIndex = 3
        Me.btnPrintReinforcingSummary.Text = "Print Job Invoice Summary"
        '
        'btn_printSummaryOfBendingSchedule
        '
        Me.btn_printSummaryOfBendingSchedule.Location = New System.Drawing.Point(272, 32)
        Me.btn_printSummaryOfBendingSchedule.Name = "btn_printSummaryOfBendingSchedule"
        Me.btn_printSummaryOfBendingSchedule.Size = New System.Drawing.Size(168, 40)
        Me.btn_printSummaryOfBendingSchedule.TabIndex = 3
        Me.btn_printSummaryOfBendingSchedule.Text = "Print Summary of Bending Schedules"
        '
        'btnInvSummary
        '
        Me.btnInvSummary.Location = New System.Drawing.Point(272, 96)
        Me.btnInvSummary.Name = "btnInvSummary"
        Me.btnInvSummary.Size = New System.Drawing.Size(168, 40)
        Me.btnInvSummary.TabIndex = 3
        Me.btnInvSummary.Text = "Print Monthly Invoices"
        '
        'btnJobSumRep
        '
        Me.btnJobSumRep.Location = New System.Drawing.Point(272, 160)
        Me.btnJobSumRep.Name = "btnJobSumRep"
        Me.btnJobSumRep.Size = New System.Drawing.Size(168, 40)
        Me.btnJobSumRep.TabIndex = 3
        Me.btnJobSumRep.Text = "Print Job Summary Report"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.Controls.Add(Me.lblHeading)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(520, 88)
        Me.Panel1.TabIndex = 6
        '
        'lblHeading
        '
        Me.lblHeading.Font = New System.Drawing.Font("Times New Roman", 27.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHeading.ForeColor = System.Drawing.Color.RoyalBlue
        Me.lblHeading.Location = New System.Drawing.Point(88, 24)
        Me.lblHeading.Name = "lblHeading"
        Me.lblHeading.Size = New System.Drawing.Size(336, 40)
        Me.lblHeading.TabIndex = 0
        Me.lblHeading.Text = "Reports"
        Me.lblHeading.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'frmReports
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(520, 462)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "frmReports"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Reports & Printing"
        Me.GroupBox1.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private CallingForm As Object
    Dim DbConnection As OleDb.OleDbConnection
    Public Sub New(ByVal caller As Object, ByRef dbc As OleDb.OleDbConnection)
        MyBase.New()
        InitializeComponent()

        CallingForm = caller
        DbConnection = dbc
    End Sub


    Private Shadows Sub FormClosing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If Not IsNothing(CallingForm) Then
            CallingForm.Show()
        End If

        CallingForm = Nothing
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Close()
    End Sub

    Private Sub btnPrintReinforcingSummary_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrintReinforcingSummary.Click
        Dim form As New frmPrintReinforcingSummary(Me)
        Me.Hide()
        form.Show()
    End Sub

    Private Sub btnPrintCutInv_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrintCutInv.Click
        Dim frm As New PrintCutInv(Me)
        Me.Hide()
        frm.Show()
    End Sub

    Private Sub btn_printSummaryOfBendingSchedule_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_printSummaryOfBendingSchedule.Click
        Dim form As New frm_printSummaryOfBendingSchedule(Me)
        Me.Hide()
        form.Show()
    End Sub

    Private Sub btnPrintJobRept_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub btnInvSummary_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInvSummary.Click
        Dim form As frmInvSummary = New frmInvSummary(Me, DbConnection)
        Me.Hide()
        form.Show()
    End Sub

    Private Sub btnJobSumRep_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnJobSumRep.Click
        Dim form As frmJobSummaryReport = New frmJobSummaryReport(Me, DbConnection)
        Me.Hide()
        form.Show()
    End Sub

    Private Sub btnPrintCutSheet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrintCutSheet.Click
        Dim frm As New PrintCutSheet(Me)
        Me.Hide()
        frm.Show()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim frm As New PrintJobRates(Me)
        Me.Hide()
        frm.Show()
    End Sub

    Private Sub btnJobListing_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnJobListing.Click
        Dim frm As New PrintJobs(Me)
        Me.Hide()
        frm.Show()
    End Sub
End Class
