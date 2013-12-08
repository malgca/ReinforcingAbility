Imports System
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Drawing.Printing

Public Class frm_printSummaryOfBendingSchedule
    Inherits Form

#Region " Global Variables "
    Dim DBConnection As New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source = winsteelVers5.mdb")
    Dim field As PageElement
    Dim PrintArray As New ArrayList
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

    Dim CallingForm As Object
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
        populate_cmb_jobs()
    End Sub

    Private Sub populate_cmb_jobs()
        cmbJobs.Items.Clear()
        Dim sql As String = "SELECT JobNo FROM Job ORDER BY JobNo"
        Dim dataset As New Data.DataSet
        Dim adapter As New OleDb.OleDbDataAdapter(sql, DBConnection)
        adapter.Fill(dataset)

        Dim aunty As Integer
        For aunty = 0 To dataset.Tables(0).Rows.Count - 1
            cmbJobs.Items.Add(dataset.Tables(0).Rows(aunty).Item("JobNo").ToString())
        Next aunty
    End Sub

    Dim sql As String

    Const d2 As Integer = 75

    Private Function getColumn(ByVal tC As String) As Integer
        tC = tC.Substring(1)
        If tC = "06" Then
            Return LeftMargin + 1 * d2 - 20
        ElseIf tC = "08" Then
            Return LeftMargin + 2 * d2 - 20
        ElseIf tC = "10" Then
            Return LeftMargin + 3 * d2 - 20
        ElseIf tC = "12" Then
            Return LeftMargin + 4 * d2 - 20
        ElseIf tC = "16" Then
            Return LeftMargin + 5 * d2 - 20
        ElseIf tC = "20" Then
            Return LeftMargin + 6 * d2 - 20
        ElseIf tC = "25" Then
            Return LeftMargin + 7 * d2 - 20
        ElseIf tC = "32" Then
            Return LeftMargin + 8 * d2 - 20
        ElseIf tC = "40" Then
            Return LeftMargin + 9 * d2 - 20
        End If
    End Function

    Private Sub PrintSched(ByVal inSched As String, ByVal inType As String)
        PrintArray.Add(New PageElement(inSched, EntryFont, LeftMargin, True, False, False))
        PrintArray.Add(New PageElement(inType, EntryFont, LeftMargin + 40, False, False, False))
    End Sub
    Private Sub PrintLine()
        PrintArray.Add(New PageElement("", EntryFont, 0, True, False, False))
        PrintArray.Add(New PageElement("<HR/BLACK>", EntryFont, LeftMargin, True))
    End Sub

    Private Sub PrintYHead(ByVal inType As String)
        PrintArray.Add(New PageElement("", EntryFont, 0, True, False, False))
        PrintArray.Add(New PageElement(inType, EntryFont, LeftMargin + 40, False, False, False))
    End Sub

    Private Sub PrintTypeHeadings()
        PrintArray.Add(New PageElement("Schedule", EntryFont, LeftMargin, False))
        PrintArray.Add(New PageElement("06", EntryFont, LeftMargin + 1 * d2, False))
        PrintArray.Add(New PageElement("08", EntryFont, LeftMargin + 2 * d2, False))

        PrintArray.Add(New PageElement("10", EntryFont, LeftMargin + 3 * d2, False))
        PrintArray.Add(New PageElement("12", EntryFont, LeftMargin + 4 * d2, False))
        PrintArray.Add(New PageElement("16", EntryFont, LeftMargin + 5 * d2, False))
        PrintArray.Add(New PageElement("20", EntryFont, LeftMargin + 6 * d2, False))
        PrintArray.Add(New PageElement("25", EntryFont, LeftMargin + 7 * d2, False))
        PrintArray.Add(New PageElement("32", EntryFont, LeftMargin + 8 * d2, False))

        PrintArray.Add(New PageElement("40", EntryFont, LeftMargin + 9 * d2, True, False, False))
        PrintArray.Add(New PageElement("<HR/BLACK>", EntryFont, LeftMargin, True))
    End Sub
    Private Sub PrintRHead()
        PrintArray.Add(New PageElement("", EntryFont, 0, True, False, False))
        PrintArray.Add(New PageElement("Total", EntryFont, LeftMargin, False, False, False))
        PrintArray.Add(New PageElement("R", EntryFont, LeftMargin + 40, False, False, False))
    End Sub
    Private Sub PrintYHead()
        PrintArray.Add(New PageElement("", EntryFont, 0, True, False, False))
        PrintArray.Add(New PageElement("Y", EntryFont, LeftMargin + 40, False, False, False))
    End Sub

    Private Sub PrintVal(ByVal inVal As Double, ByVal inElem As Int16, ByVal inTKg As String)
        Dim vo As String
        If inVal <> 0 Then
            If inTKg = "T" Then
                vo = inVal.ToString("0.000")
            Else
                vo = inVal.ToString("0.0")
            End If
            PrintArray.Add(New PageElement(vo, EntryFont, PageWidth - ((8 - inElem) * d2) - 100, False, False, True))
        End If
    End Sub

    Private Sub GenerateSummaryOfBendingSchedules(ByVal jobNo As String, ByVal aDate As Date)
        PrintArray = New ArrayList
        Dim p As PageElement
        Dim TKg As String = String.Empty

        p = New PageElement("SUMMARY OF BENDING SCHEDULES", EntryFont, 0, True, True, False)
        PrintArray.Add(p)

        Dim sql4compName As String = "SELECT ContractorName, JobName,CompanyName,job.[Tons or Kilograms] AS TKG " & _
                    "FROM Job, Contractor,Company " & _
                    "WHERE Job.ContractorNo = Contractor.ContractorNo " & _
                    "AND Company.CompanyNo = Job.CompanyNo " & _
                    "AND Job.JobNo = '" & jobNo & "'"

        Dim ds As New Data.DataSet
        Dim ad As New OleDb.OleDbDataAdapter(sql4compName, DBConnection)
        Dim currJobName As String
        ad.Fill(ds)

        If ds.Tables(0).Rows.Count = 1 Then
            Const d1 As Integer = 85

            PrintArray.Add(New PageElement(ds.Tables(0).Rows(0).Item("CompanyName").ToString(), EntryFont, 0, True, True, False))
            PrintArray.Add(New PageElement("Job Number:", EntryFont, LeftMargin, False, False, False))
            PrintArray.Add(New PageElement(jobNo, EntryFont, LeftMargin + d1, True, False, False))
            currJobName = ds.Tables(0).Rows(0).Item("JobName").ToString()
            TKg = ds.Tables(0).Rows(0).Item("TKG").ToString()
            PrintArray.Add(New PageElement("Job Name:", EntryFont, LeftMargin, False, False, False))
            PrintArray.Add(New PageElement(currJobName, EntryFont, LeftMargin + d1, True, False, False))
            PrintArray.Add(New PageElement("Contractor:", EntryFont, LeftMargin, False, False, False))
            PrintArray.Add(New PageElement(ds.Tables(0).Rows(0).Item("ContractorName").ToString(), EntryFont, LeftMargin + d1, True, False, False))
            PrintArray.Add(New PageElement("Date:", EntryFont, LeftMargin, False, False, False))
            PrintArray.Add(New PageElement(aDate.ToShortDateString, EntryFont, LeftMargin + d1, True, False, False))
            PrintArray.Add(New PageElement("<SPACE>", EntryFont, LeftMargin, True, False, False))
            PrintArray.Add(New PageElement("<HR/BLACK>", EntryFont, LeftMargin, True))
            PrintTypeHeadings()
        End If

        Dim rowCnt, itemCnt As Int16
        Dim RperSched(8) As Double
        Dim YperSched(8) As Double
        Dim RTotals(8) As Double
        Dim YTotals(8) As Double
        Dim schedNo, prevSched, cutNo As String
        Dim hasItems, schedChange As Boolean
        Dim i, f, r, j As Integer
        Dim typeR, typeY, curTC As String
        typeR = "R"
        typeY = "Y"
        Dim TR, TY, curWeight, curQty, curSteel As Double
        TR = 0
        TY = 0
        For i = 0 To 8
            RTotals(i) = 0
            YTotals(i) = 0
        Next i
        schedChange = False
        prevSched = ""
        curSteel = 0
        ' GET ALL THE SCHEDULES AND CUTTING SHEETS FOR THAT JOB IN THAT DATE RANGE
        Dim sql4ScheduleNos As String = "SELECT DISTINCT ScheduleNo, CuttingSheet.CutSheetNo" & _
        " FROM CuttingSheet INNER JOIN SchedItem ON CuttingSheet.CutSheetNo = SchedItem.CutSheetNo " & _
        "WHERE CutDate <= #" & aDate.ToShortDateString & "# AND InvoiceNo <> 0 AND [Job No] = '" & jobNo & "'" & _
        "ORDER BY ScheduleNo"

        Dim DS4SchNo As Data.DataSet = New Data.DataSet
        Dim adapter As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(sql4ScheduleNos, DBConnection)
        adapter.Fill(DS4SchNo)

        rowCnt = DS4SchNo.Tables(0).Rows.Count
        ' /* FOR EACH SCHEDULE AND CUTTING SHEET*/
        For i = 0 To rowCnt - 1
            schedNo = DS4SchNo.Tables(0).Rows(i).Item("ScheduleNo").ToString().ToUpper
            cutNo = DS4SchNo.Tables(0).Rows(i).Item("CutSheetNo").ToString()

            ' see if schedule has changed
            If prevSched <> schedNo And prevSched <> "" Then
                schedChange = True
            Else
                schedChange = False
            End If

            ' see if schedule has changed '
            If schedChange Then
                'PRINT SCHEDULE NUMBER
                PrintSched(prevSched, typeR)

                '/* ROUND AND PRINT ALL Rs FOR THE SCHEDULE*/
                For f = 0 To 8
                    PrintVal(RperSched(f), f, TKg)
                Next f

                PrintYHead(typeY)
                ' PRINT ALL Ys
                For f = 0 To 8
                    PrintVal(YperSched(f), f, TKg)
                Next f

                'print line
                PrintLine()
                ' clear totals for next schedule */
                For j = 0 To 8
                    RperSched(j) = 0
                    YperSched(j) = 0
                Next j
            End If  ' end of schedule change

            '/*  GET ALL THE CUTTING SHEETS AND ITEMS FOR THE SCHEDULE */
            Dim sqlPerSchR As String = "SELECT CutItem.ScheduleNo, ProductType.TypeCode, CutItem.TypeCode, CutItem.Length, CutItem.Qty, ProductType.Weight " & _
                        "FROM CutItem, ProductType " & _
                       "WHERE CutItem.ScheduleNo = '" & schedNo & "'" & _
                       "AND CutItem.CutSheetNo = " & cutNo & _
                        "AND CutItem.TypeCode = ProductType.TypeCode "

            Dim ds4r As New Data.DataSet
            Dim ad4R As New OleDb.OleDbDataAdapter(sqlPerSchR, DBConnection)
            ad4R.Fill(ds4r)
            hasItems = False
            itemCnt = ds4r.Tables(0).Rows.Count
            '/* IF THERE ARE ITEMS IN THE SCHEDULE */
            If itemCnt <> 0 Then
                hasItems = True
            End If

            If hasItems Then
                '/* LOOP THROUGH EACH ITEM */
                For r = 0 To itemCnt - 1
                    curWeight = ds4r.Tables(0).Rows(r).Item("Weight")
                    curQty = ds4r.Tables(0).Rows(r).Item("Qty")
                    curTC = ds4r.Tables(0).Rows(r).Item("CutItem.TypeCode").ToString()
                    curSteel = ds4r.Tables(0).Rows(r).Item("Length") * curQty * curWeight

                    If TKg = "T" Then
                        curSteel = curSteel / 1000000
                    Else
                        curSteel = curSteel / 1000
                    End If

                    'MessageBox.Show("Add to Total " + "curTc " + curTC + " " + curSteel.ToString + " " + TKg)
                    ' ADD TO TOTAL FOR SCHEDULE
                    If curTC = "R06" Then
                        RperSched(0) += curSteel
                        RTotals(0) += curSteel
                    ElseIf curTC = "R08" Then
                        RperSched(1) += curSteel
                        RTotals(1) += curSteel
                    ElseIf curTC = "R10" Then
                        RperSched(2) += curSteel
                        RTotals(2) += curSteel
                    ElseIf curTC = "R12" Then
                        RperSched(3) += curSteel
                        RTotals(3) += curSteel
                    ElseIf curTC = "R16" Then
                        RperSched(4) += curSteel
                        RTotals(4) += curSteel
                    ElseIf curTC = "R20" Then
                        RperSched(5) += curSteel
                        RTotals(5) += curSteel
                    ElseIf curTC = "R25" Then
                        RperSched(6) += curSteel
                        RTotals(6) += curSteel
                    ElseIf curTC = "R32" Then
                        RperSched(7) += curSteel
                        RTotals(7) += curSteel
                    ElseIf curTC = "R40" Then
                        RperSched(8) += curSteel
                        RTotals(8) += curSteel
                    End If

                    '/* CHECK Y TYPES
                    If curTC = "Y06" Then
                        YperSched(0) += curSteel
                        YTotals(0) += curSteel
                    ElseIf curTC = "Y08" Then
                        YperSched(1) += curSteel
                        YTotals(1) += curSteel
                    ElseIf curTC = "Y10" Then
                        YperSched(2) += curSteel
                        YTotals(2) += curSteel
                    ElseIf curTC = "Y12" Then
                        YperSched(3) += curSteel
                        YTotals(3) += curSteel
                    ElseIf curTC = "Y16" Then
                        YperSched(4) += curSteel
                        YTotals(4) += curSteel
                    ElseIf curTC = "Y20" Then
                        YperSched(5) += curSteel
                        YTotals(5) += curSteel
                    ElseIf curTC = "Y25" Then
                        YperSched(6) += curSteel
                        YTotals(6) += curSteel
                    ElseIf curTC = "Y32" Then
                        YperSched(7) += curSteel
                        YTotals(7) += curSteel
                    ElseIf curTC = "Y40" Then
                        YperSched(8) += curSteel
                        YTotals(8) += curSteel
                    End If
                Next r  ' next item in schedule
            End If   ' if this cutting sheet & schedule has items

            prevSched = schedNo
        Next i
        '/* end of all schedules and cutting sheets for this job

        'PRINT SCHEDULE NUMBER
        PrintSched(prevSched, typeR)

        '/* ROUND AND PRINT ALL Rs FOR THE SCHEDULE*/
        For f = 0 To 8
            PrintVal(RperSched(f), f, TKg)
        Next f

        PrintYHead(typeY)
        ' PRINT ALL Ys
        For f = 0 To 8
            PrintVal(YperSched(f), f, TKg)
        Next f

        'print line
        PrintLine()
        PrintRHead()

        Dim ci As Integer
        For ci = 0 To 8
            PrintVal(RTotals(ci), ci, TKg)
            TR += RTotals(ci)
        Next ci

        PrintYHead()

        For ci = 0 To 8
            PrintVal(YTotals(ci), ci, TKg)
            TY += YTotals(ci)
        Next ci
        PrintArray.Add(New PageElement("", EntryFont, 0, True, False, False))
        PrintArray.Add(New PageElement("", EntryFont, 0, True, False, False))
        PrintArray.Add(New PageElement("Total Mild Steel:", EntryFont, LeftMargin, False, False, False))
        Dim v As String
        If TKg = "T" Then
            v = TR.ToString("0.000")
        Else
            v = TR.ToString("0.0")
        End If
        PrintArray.Add(New PageElement(v & " " & TKg, EntryFont, LeftMargin + 300, True, False, True))

        PrintArray.Add(New PageElement("Total High Tensile Steel:", EntryFont, LeftMargin, False, False, False))
        If TKg = "T" Then
            v = TY.ToString("0.000")
        Else
            v = TY.ToString("0.0")
        End If
        PrintArray.Add(New PageElement(v & " " & TKg, EntryFont, LeftMargin + 300, True, False, True))
        PrintArray.Add(New PageElement("", EntryFont, 0, True, False, False))
        PrintArray.Add(New PageElement("Grand Total:", EntryFont, LeftMargin, False, False, False))
        If TKg = "T" Then
            v = (TY + TR).ToString("0.000")
        Else
            v = (TY + TR).ToString("0.0")
        End If
        PrintArray.Add(New PageElement(v & " " & TKg, EntryFont, LeftMargin + 300, True, False, True))

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
