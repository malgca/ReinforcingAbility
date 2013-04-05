Public Class frmPrintJobInvSumm
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
    Friend WithEvents DocumentToPrint As System.Drawing.Printing.PrintDocument
    Friend WithEvents btnPrintPreview As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents cmbJobs As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.DocumentToPrint = New System.Drawing.Printing.PrintDocument
        Me.btnPrintPreview = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.cmbJobs = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'btnPrintPreview
        '
        Me.btnPrintPreview.Location = New System.Drawing.Point(56, 152)
        Me.btnPrintPreview.Name = "btnPrintPreview"
        Me.btnPrintPreview.Size = New System.Drawing.Size(176, 40)
        Me.btnPrintPreview.TabIndex = 4
        Me.btnPrintPreview.Text = "Print Preview..."
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(56, 208)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(176, 40)
        Me.btnClose.TabIndex = 5
        Me.btnClose.Text = "Close"
        '
        'cmbJobs
        '
        Me.cmbJobs.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbJobs.Location = New System.Drawing.Point(136, 80)
        Me.cmbJobs.Name = "cmbJobs"
        Me.cmbJobs.Size = New System.Drawing.Size(104, 21)
        Me.cmbJobs.TabIndex = 6
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(24, 80)
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Job"
        '
        'frmPrintJobInvSumm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(292, 266)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmbJobs)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnPrintPreview)
        Me.Name = "frmPrintJobInvSumm"
        Me.Text = "frmPrintJobInvSumm"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub frmPrintJobInvSumm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub btnPrintPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrintPreview.Click
        If cmbJobs.Text = "" Then
            MessageBox.Show("Select a job number from the drop-down list.", "Invalid job number", MessageBoxButtons.OK)
            cmbJobs.Focus()
            Exit Sub
        End If

        Dim sql = "SELECT CompanyNo FROM Job WHERE JobNo = '" & cmbJobs.Text & "'"
        Dim DataSet = New Data.DataSet
        Dim adapter = New OleDb.OleDbDataAdapter(sql, DBConnection)
        adapter.Fill(DataSet)

        If DataSet.tables(0).rows.count = 0 Then
            MessageBox.Show("Selected job does not have an associated company number.", "Slight Error", MessageBoxButtons.OK)
            Exit Sub
        End If


        Try
            DocumentToPrint.DocumentName = "Reinforcing Summary - Job No: " + cmbJobs.Text
            Dim ppd_JCR As New PrintPreviewDialog
            ppd_JCR.WindowState = FormWindowState.Maximized
            ppd_JCR.Document = DocumentToPrint
            ppd_JCR.AutoScale = True
            ppd_JCR.AutoScroll = True
            ppd_JCR.UseAntiAlias = False
            ppd_JCR.PrintPreviewControl.Zoom = 1
            ppd_JCR.PrintPreviewControl.Columns = 1
            ppd_JCR.PrintPreviewControl.Rows = 1
            ppd_JCR.Text = "Reinforcing Summary - Job No: " + cmbJobs.Text
            curpagenum = 1
            PrintArray.Clear()

            ReinforcingSummaryPrint(DataSet.tables(0).rows(0).Item("CompanyNo").ToString, cmbJobs.Text, dtpStart.Value.ToShortDateString, dtpEnd.Value.ToShortDateString)

            curArrayPos = 0

            If All_Is_OK Then
                ppd_JCR.ShowDialog()
            Else
                Exit Sub
            End If

        Catch er As Exception
            If er.Message = "No printers installed." Then
                MessageBox.Show("There is no printer installed. Please install a printer and try again.", "Printer not found", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                MessageBox.Show(er.Message, "ERROR - PLEASE FIX ME!!")
            End If

        End Try

    End Sub
End Class
