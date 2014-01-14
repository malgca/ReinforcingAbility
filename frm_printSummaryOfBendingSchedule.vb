Imports System
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Drawing.Printing
Imports LogicTier

Public Class frm_printSummaryOfBendingSchedule
    Inherits Form

    Private Property Logic As New BendingSchedule
    Private Property CallingForm As Object

#Region " Windows Form Designer generated code "

    Public Sub New(ByRef Caller As Object)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        CallingForm = Caller
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
    Friend WithEvents btnPrintPreview As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents cmbJobs As ComboBox
    Friend WithEvents Label2 As Label
    Friend WithEvents btnClose As Button
    Friend WithEvents dtpReportDate As DateTimePicker
    Friend WithEvents DocumentToPrint As PrintDocument
    <Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btnPrintPreview = New Button
        Me.Label1 = New Label
        Me.cmbJobs = New ComboBox
        Me.dtpReportDate = New DateTimePicker
        Me.Label2 = New Label
        Me.btnClose = New Button
        Me.DocumentToPrint = New PrintDocument
        Me.SuspendLayout()
        '
        'btnPrintPreview
        '
        Me.btnPrintPreview.Location = New Point(40, 104)
        Me.btnPrintPreview.Name = "btnPrintPreview"
        Me.btnPrintPreview.Size = New Size(176, 40)
        Me.btnPrintPreview.TabIndex = 9
        Me.btnPrintPreview.Text = "Print Preview..."
        '
        'Label1
        '
        Me.Label1.Location = New Point(40, 32)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New Size(64, 23)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Job No.:"
        Me.Label1.TextAlign = ContentAlignment.MiddleLeft
        '
        'cmbJobs
        '
        Me.cmbJobs.DataSource = Logic.JobNameList
        Me.cmbJobs.Location = New Point(112, 32)
        Me.cmbJobs.Name = "cmbJobs"
        Me.cmbJobs.Size = New Size(104, 21)
        Me.cmbJobs.TabIndex = 4
        '
        'dtpReportDate
        '
        Me.dtpReportDate.Format = DateTimePickerFormat.Short
        Me.dtpReportDate.Location = New Point(112, 64)
        Me.dtpReportDate.Name = "dtpReportDate"
        Me.dtpReportDate.Size = New Size(104, 20)
        Me.dtpReportDate.TabIndex = 5
        '
        'Label2
        '
        Me.Label2.Location = New Point(40, 64)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New Size(72, 23)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Report Date:"
        Me.Label2.TextAlign = ContentAlignment.MiddleLeft
        '
        'btnClose
        '
        Me.btnClose.Location = New Point(40, 160)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New Size(176, 40)
        Me.btnClose.TabIndex = 8
        Me.btnClose.Text = "Close"
        '
        'DocumentToPrint
        '
        '
        'frm_printSummaryOfBendingSchedule
        '
        Me.AutoScaleBaseSize = New Size(5, 13)
        Me.ClientSize = New Size(264, 229)
        Me.Controls.Add(Me.btnPrintPreview)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmbJobs)
        Me.Controls.Add(Me.dtpReportDate)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnClose)
        Me.FormBorderStyle = FormBorderStyle.FixedToolWindow
        Me.Name = "frm_printSummaryOfBendingSchedule"
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.Text = "Print Summary of Bending Schedule"
        Me.ResumeLayout(False)
    End Sub
#End Region

    Private Sub frm_printSummaryOfBendingSchedule_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        Logic.InitializeProperties(0)
    End Sub

    Private Sub btnPrintPreview_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnPrintPreview.Click
        If cmbJobs.Text = String.Empty Then
            MessageBox.Show("Select a job number from the drop-down list.", "Invalid job number", MessageBoxButtons.OK)
            cmbJobs.Focus()
            Exit Sub
        End If

        Try
            DocumentToPrint.DocumentName = "Summary of Bending Schedules - Job No: " + cmbJobs.Text

            Dim printPreview As New PrintPreviewDialog

            printPreview.WindowState = FormWindowState.Maximized
            printPreview.Document = DocumentToPrint
            printPreview.AutoScale = True
            printPreview.AutoScroll = True
            printPreview.UseAntiAlias = False
            printPreview.PrintPreviewControl.Zoom = 1
            printPreview.PrintPreviewControl.Columns = 1
            printPreview.PrintPreviewControl.Rows = 1
            printPreview.Text = "Summary of Bending Schedules - Job No: " + cmbJobs.Text

            Logic.CurrentPageNumber = 1

            Logic.PrintList.Clear()

            'Put method to populate print array here
            Logic.GenerateSummaryOfBendingSchedules(cmbJobs.Text, dtpReportDate.Value)

            Logic.CurrentListPosition = 0

            Dim noErrors As Boolean = True

            If noErrors Then
                printPreview.ShowDialog()
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

    Private Sub PrintPage(ByVal sender As Object, ByVal e As PrintPageEventArgs) Handles DocumentToPrint.PrintPage
        Me.Cursor = Cursors.Arrow

        Dim topMargin As Integer = BendingSchedule.PageConstants.TopMargin
        Dim pageBound As Integer = e.PageSettings.Bounds.Height - BendingSchedule.PageConstants.BottomMargin

        If BendingSchedule.PageConstants.ReportType = "Reinforcing Summary" Then
            e.Graphics.DrawString("Date Generated : " & Today().ToShortDateString, BendingSchedule.PrintFonts.SmallItalic, Brushes.DimGray, BendingSchedule.PageConstants.LeftMargin, 1065)
            e.Graphics.DrawString("Page " & Logic.CurrentPageNumber, BendingSchedule.PrintFonts.SmallItalic, Brushes.DimGray, 700, 1065)
        End If

        While (topMargin < pageBound) And (Logic.CurrentListPosition < Logic.PrintList.Count)
            Select Case Logic.PrintList(Logic.CurrentListPosition).Text.ToString()
                Case "<SPACE>"
                    If Logic.PrintList(Logic.CurrentListPosition).includeEOL Then
                        topMargin += Logic.PrintList(Logic.CurrentListPosition).Font.Size + 30 + Logic.PrintList(Logic.CurrentListPosition).Ygap
                    End If
                Case "#LINE__"
                    e.Graphics.DrawLine(Pens.Black, Logic.PrintList(Logic.CurrentListPosition).x, topMargin, Logic.PrintList(Logic.CurrentListPosition).x2, topMargin)

                    If Logic.PrintList(Logic.CurrentListPosition).includeEOL Then
                        topMargin += Logic.PrintList(Logic.CurrentListPosition).Font.Size + 10 + Logic.PrintList(Logic.CurrentListPosition).Ygap
                    End If
                Case "#DOUBLELINE__"
                    e.Graphics.DrawLine(Pens.Black, Logic.PrintList(Logic.CurrentListPosition).x, topMargin, Logic.PrintList(Logic.CurrentListPosition).x2, topMargin)
                    e.Graphics.DrawLine(Pens.Black, Logic.PrintList(Logic.CurrentListPosition).x, topMargin + 3, Logic.PrintList(Logic.CurrentListPosition).x2, topMargin + 3)

                    If Logic.PrintList(Logic.CurrentListPosition).includeEOL Then
                        topMargin += Logic.PrintList(Logic.CurrentListPosition).Font.Size + 10 + Logic.PrintList(Logic.CurrentListPosition).Ygap
                    End If
                Case "<HR/>"
                    e.Graphics.DrawLine(Pens.LightGray, BendingSchedule.PageConstants.LeftMargin, topMargin, 800, topMargin)

                    If Logic.PrintList(Logic.CurrentListPosition).includeEOL Then
                        topMargin += Logic.PrintList(Logic.CurrentListPosition).Font.Size + 10 + Logic.PrintList(Logic.CurrentListPosition).Ygap
                    End If
                Case "<HR/BLACK>"
                    e.Graphics.DrawLine(Pens.Black, BendingSchedule.PageConstants.LeftMargin, topMargin, e.PageSettings.Bounds.Width - BendingSchedule.PageConstants.RightMargin, topMargin)

                    If Logic.PrintList(Logic.CurrentListPosition).includeEOL Then
                        topMargin += Logic.PrintList(Logic.CurrentListPosition).Font.Size + 10 + Logic.PrintList(Logic.CurrentListPosition).Ygap
                    End If
                Case "<HR/LIGHT>"
                    e.Graphics.DrawLine(Pens.WhiteSmoke, BendingSchedule.PageConstants.LeftMargin, topMargin, e.PageSettings.Bounds.Width - BendingSchedule.PageConstants.RightMargin, topMargin)

                    If Logic.PrintList(Logic.CurrentListPosition).includeEOL Then
                        topMargin += Logic.PrintList(Logic.CurrentListPosition).Font.Size + 5 + Logic.PrintList(Logic.CurrentListPosition).Ygap
                    End If
                Case Else
                    If Logic.PrintList(Logic.CurrentListPosition).center Then
                        Dim stringSize As New SizeF

                        stringSize = e.Graphics.MeasureString(Logic.PrintList(Logic.CurrentListPosition).Text, BendingSchedule.PrintFonts.Normal)
                        e.Graphics.DrawString(Logic.PrintList(Logic.CurrentListPosition).Text, Logic.PrintList(Logic.CurrentListPosition).Font, Brushes.Black, (e.PageSettings.Bounds.Width / 2) - 0.5 * stringSize.Width, topMargin)
                    ElseIf Logic.PrintList(Logic.CurrentListPosition).rAlign Then
                        Dim stringSize As New SizeF

                        stringSize = e.Graphics.MeasureString(Logic.PrintList(Logic.CurrentListPosition).Text, BendingSchedule.PrintFonts.Normal)
                        e.Graphics.DrawString(Logic.PrintList(Logic.CurrentListPosition).Text, Logic.PrintList(Logic.CurrentListPosition).Font, Brushes.Black, Logic.PrintList(Logic.CurrentListPosition).x - stringSize.Width, topMargin)
                    Else
                        e.Graphics.DrawString(Logic.PrintList(Logic.CurrentListPosition).Text, Logic.PrintList(Logic.CurrentListPosition).Font, Brushes.Black, Logic.PrintList(Logic.CurrentListPosition).x, topMargin)
                    End If

                    If Logic.PrintList(Logic.CurrentListPosition).includeEOL Then
                        topMargin += Logic.PrintList(Logic.CurrentListPosition).Font.Size + 10 + Logic.PrintList(Logic.CurrentListPosition).Ygap
                    End If
            End Select

            Logic.CurrentListPosition += 1
        End While

        If topMargin >= pageBound Then
            Logic.CurrentPageNumber += 1
            e.HasMorePages = True
        Else
            e.HasMorePages = False
            Logic.CurrentListPosition = 0
            Logic.CurrentPageNumber = 1
        End If
    End Sub

    Private Shadows Sub FormClosing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If Not IsNothing(CallingForm) Then
            CallingForm.Show()
        End If

        CallingForm = Nothing
    End Sub

    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
        Close()
    End Sub

    Private Sub cmbJobs_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles cmbJobs.SelectedIndexChanged
    End Sub
End Class
