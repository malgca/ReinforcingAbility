Imports System
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Drawing.Printing
Imports LogicTier

Public Class frm_printSummaryOfBendingSchedule
    Inherits Form

#Region " Global Variables "
    Dim DBConnection As New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source = winsteelVers5.mdb")
    Dim field As PageElement
    Dim EntryFont As New Font("Arial", 10)
    Dim Head1Font As New Font("Arial", 30, FontStyle.Bold Or FontStyle.Underline)
    Dim Head2Font As New Font("Arial", 15, FontStyle.Bold)
    Dim Head2DetFont As New Font("Arial", 15, FontStyle.Italic)
    Dim EntryFontBold As New Font("Arial", 10, FontStyle.Bold)
    Dim EntryFontUnderline As New Font("Arial", 10, FontStyle.Underline)
    Dim DetailFont As New Font("Arial", 13)
    Dim TimeCardColFont As New Font("Arial", 10, FontStyle.Italic Or FontStyle.Bold)
    Dim ColFont As New Font("Arial", 12, FontStyle.Italic)
    Dim curArrayPos As Integer = 0
    Dim curpagenum As Integer = 1
    Dim TopMargin As Integer = 60
    Dim LeftMargin As Integer = 60
    Dim RightMargin As Integer = 60
    Dim BottomMargin As Integer = 90
    Dim PageWidth As Integer = 873
    Dim ReportType As String
    Dim vatperc As String
    Dim All_Is_OK As Boolean = True

    Private Property Logic As New BendingSchedule
    Private Property CallingForm As Object

    Const d2 As Integer = 75
#End Region

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

        If cmbJobs.Text = "" Then
            MessageBox.Show("Select a job number from the drop-down list.", "Invalid job number", MessageBoxButtons.OK)
            cmbJobs.Focus()
            Exit Sub
        End If

        Try
            DocumentToPrint.DocumentName = "Summary of Bending Schedules - Job No: " + cmbJobs.Text
            Dim ppd_JCR As New PrintPreviewDialog
            ppd_JCR.WindowState = FormWindowState.Maximized
            ppd_JCR.Document = DocumentToPrint
            ppd_JCR.AutoScale = True
            ppd_JCR.AutoScroll = True
            ppd_JCR.UseAntiAlias = False
            ppd_JCR.PrintPreviewControl.Zoom = 1
            ppd_JCR.PrintPreviewControl.Columns = 1
            ppd_JCR.PrintPreviewControl.Rows = 1
            ppd_JCR.Text = "Summary of Bending Schedules - Job No: " + cmbJobs.Text
            curpagenum = 1
            PrintArray.Clear()

            'Put method to populate print array here
            GenerateSummaryOfBendingSchedules(cmbJobs.Text, dtpReportDate.Value)

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

    Private Sub PrintPage(ByVal sender As Object, ByVal e As PrintPageEventArgs) Handles DocumentToPrint.PrintPage

        Me.Cursor = Windows.Forms.Cursors.Arrow
        Dim curY As Integer = TopMargin
        Dim MaxY As Integer = e.PageSettings.Bounds.Height - BottomMargin

        If ReportType = "Reinforcing Summary" Then
            e.Graphics.DrawString("Date Generated : " & Today().ToShortDateString, New Font("Arial", 8, FontStyle.Italic), Brushes.DimGray, LeftMargin, 1065)
            e.Graphics.DrawString("Page " & curpagenum, New Font("Arial", 8, FontStyle.Italic), Brushes.DimGray, 700, 1065)
        End If

        While (curY < MaxY) And (curArrayPos < PrintArray.Count)

            Select Case PrintArray(curArrayPos).Text.ToString()
                Case "<SPACE>"
                    'e.Graphics.DrawLine(Pens.LightGray, LeftMargin, curY, 800, curY)
                    If PrintArray(curArrayPos).includeEol Then
                        curY += PrintArray(curArrayPos).Font.Size + 30 + PrintArray(curArrayPos).ygap
                    End If
                Case "#LINE__"
                    e.Graphics.DrawLine(Pens.Black, PrintArray(curArrayPos).x, curY, PrintArray(curArrayPos).x2, curY)
                    If PrintArray(curArrayPos).includeEol Then
                        curY += PrintArray(curArrayPos).Font.Size + 10 + PrintArray(curArrayPos).ygap
                    End If
                Case "#DOUBLELINE__"
                    e.Graphics.DrawLine(Pens.Black, PrintArray(curArrayPos).x, curY, PrintArray(curArrayPos).x2, curY)
                    e.Graphics.DrawLine(Pens.Black, PrintArray(curArrayPos).x, curY + 3, PrintArray(curArrayPos).x2, curY + 3)
                    If PrintArray(curArrayPos).includeEol Then
                        curY += PrintArray(curArrayPos).Font.Size + 10 + PrintArray(curArrayPos).ygap
                    End If
                Case "<HR/>"
                    e.Graphics.DrawLine(Pens.LightGray, LeftMargin, curY, 800, curY)
                    If PrintArray(curArrayPos).includeEol Then
                        curY += PrintArray(curArrayPos).Font.Size + 10 + PrintArray(curArrayPos).ygap
                    End If
                Case "<HR/BLACK>"
                    e.Graphics.DrawLine(Pens.Black, LeftMargin, curY, e.PageSettings.Bounds.Width - RightMargin, curY)
                    If PrintArray(curArrayPos).includeEol Then
                        curY += PrintArray(curArrayPos).Font.Size + 10 + PrintArray(curArrayPos).ygap
                    End If
                Case "<HR/LIGHT>"
                    e.Graphics.DrawLine(Pens.WhiteSmoke, LeftMargin, curY, e.PageSettings.Bounds.Width - RightMargin, curY)
                    If PrintArray(curArrayPos).includeEol Then
                        curY += PrintArray(curArrayPos).Font.Size + 5 + PrintArray(curArrayPos).ygap
                    End If
                    'Case "<IMG/>"
                    '   e.Graphics.DrawImage(ImageList1.Images(PrintArray(curArrayPos).imageIndex), PrintArray(curArrayPos).x, curY)
                    '  If PrintArray(curArrayPos).includeEol Then
                    ' curY += PrintArray(curArrayPos).ImageHeight + 15
                    'End If
                Case Else
                    If PrintArray(curArrayPos).center Then
                        Dim stringSize As New SizeF
                        stringSize = e.Graphics.MeasureString(PrintArray(curArrayPos).text, EntryFont)
                        e.Graphics.DrawString(PrintArray(curArrayPos).Text, PrintArray(curArrayPos).Font, Brushes.Black, (e.PageSettings.Bounds.Width / 2) - 0.5 * stringSize.Width, curY)
                    ElseIf PrintArray(curArrayPos).ralign Then
                        Dim stringSize As New SizeF
                        stringSize = e.Graphics.MeasureString(PrintArray(curArrayPos).text, EntryFont)
                        e.Graphics.DrawString(PrintArray(curArrayPos).Text, PrintArray(curArrayPos).Font, Brushes.Black, PrintArray(curArrayPos).x - stringSize.Width, curY)
                    Else
                        e.Graphics.DrawString(PrintArray(curArrayPos).Text, PrintArray(curArrayPos).Font, Brushes.Black, PrintArray(curArrayPos).x, curY)
                    End If

                    If PrintArray(curArrayPos).includeEol Then
                        curY += PrintArray(curArrayPos).Font.Size + 10 + PrintArray(curArrayPos).ygap
                    End If
            End Select

            curArrayPos += 1
        End While

        If curY >= MaxY Then
            curpagenum += 1
            e.HasMorePages = True
        Else
            e.HasMorePages = False
            curArrayPos = 0
            curpagenum = 1
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
