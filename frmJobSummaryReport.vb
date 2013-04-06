Public Class frmJobSummaryReport
    Inherits System.Windows.Forms.Form
#Region " Global Variables "
    Dim field As PageElement
    Dim PrintArray As New ArrayList
    Dim EntryFont As New Font("Courier", 11)
    Dim Head1Font As New Font("Courier", 30, FontStyle.Bold Or FontStyle.Underline)
    Dim Head2Font As New Font("Courier", 15, FontStyle.Bold)
    Dim Head2DetFont As New Font("Courier", 15, FontStyle.Italic)
    Dim EntryFontBold As New Font("Courier", 10, FontStyle.Bold)
    Dim EntryFontUnderline As New Font("Courier", 10, FontStyle.Underline)
    Dim DetailFont As New Font("Courier", 13)
    Dim TimeCardColFont As New Font("Courier", 10, FontStyle.Italic Or FontStyle.Bold)
    Dim ColFont As New Font("Courier", 12, FontStyle.Italic)
    Dim curArrayPos As Integer = 0
    Dim curpagenum As Integer = 1
    Dim TopMargin As Integer = 90
    Dim LeftMargin As Integer = 10
    Dim RightMargin As Integer = 60
    Dim BottomMargin As Integer = 60
    Dim PageWidth As Integer = 873
    Dim ReportType As String
    'Dim mes
    Dim vatperc As String
    Dim All_Is_OK As Boolean = True
    Dim sql As String

    Const d2 As Integer = 75
#End Region
#Region " Windows Form Designer generated code "

    Private caller As Object
    Private DBConnection As OleDb.OleDbConnection

    Public Sub New(ByRef c As Object, ByRef dbc As OleDb.OleDbConnection)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        caller = c
        DBConnection = dbc

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
    Friend WithEvents btnPrintPreview As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmbJobs As System.Windows.Forms.ComboBox
    Friend WithEvents dtpReportDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents chkFaxHeader As System.Windows.Forms.CheckBox
    Friend WithEvents DocumentToPrint As System.Drawing.Printing.PrintDocument
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtRE As System.Windows.Forms.TextBox
    Friend WithEvents txtAtt As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtFaxNo As System.Windows.Forms.TextBox
    Friend WithEvents dtpStartDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents prtInvoice As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btnPrintPreview = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.cmbJobs = New System.Windows.Forms.ComboBox
        Me.dtpReportDate = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.btnClose = New System.Windows.Forms.Button
        Me.chkFaxHeader = New System.Windows.Forms.CheckBox
        Me.DocumentToPrint = New System.Drawing.Printing.PrintDocument
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtRE = New System.Windows.Forms.TextBox
        Me.txtAtt = New System.Windows.Forms.TextBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtFaxNo = New System.Windows.Forms.TextBox
        Me.dtpStartDate = New System.Windows.Forms.DateTimePicker
        Me.Label6 = New System.Windows.Forms.Label
        Me.prtInvoice = New System.Windows.Forms.CheckBox
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnPrintPreview
        '
        Me.btnPrintPreview.Location = New System.Drawing.Point(40, 256)
        Me.btnPrintPreview.Name = "btnPrintPreview"
        Me.btnPrintPreview.Size = New System.Drawing.Size(120, 24)
        Me.btnPrintPreview.TabIndex = 15
        Me.btnPrintPreview.Text = "Print Preview..."
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 23)
        Me.Label1.TabIndex = 13
        Me.Label1.Text = "Job No.:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cmbJobs
        '
        Me.cmbJobs.Location = New System.Drawing.Point(88, 16)
        Me.cmbJobs.Name = "cmbJobs"
        Me.cmbJobs.Size = New System.Drawing.Size(104, 21)
        Me.cmbJobs.TabIndex = 10
        '
        'dtpReportDate
        '
        Me.dtpReportDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpReportDate.Location = New System.Drawing.Point(88, 80)
        Me.dtpReportDate.Name = "dtpReportDate"
        Me.dtpReportDate.Size = New System.Drawing.Size(104, 20)
        Me.dtpReportDate.TabIndex = 11
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 80)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 23)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Last Date"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(200, 256)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(120, 24)
        Me.btnClose.TabIndex = 14
        Me.btnClose.Text = "Close"
        '
        'chkFaxHeader
        '
        Me.chkFaxHeader.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkFaxHeader.Location = New System.Drawing.Point(224, 16)
        Me.chkFaxHeader.Name = "chkFaxHeader"
        Me.chkFaxHeader.Size = New System.Drawing.Size(128, 24)
        Me.chkFaxHeader.TabIndex = 16
        Me.chkFaxHeader.Text = "Print Fax Header"
        '
        'DocumentToPrint
        '
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(8, 24)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(56, 23)
        Me.Label3.TabIndex = 17
        Me.Label3.Text = "Attention:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(8, 56)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(56, 23)
        Me.Label4.TabIndex = 17
        Me.Label4.Text = "RE:"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtRE
        '
        Me.txtRE.Location = New System.Drawing.Point(72, 56)
        Me.txtRE.Name = "txtRE"
        Me.txtRE.Size = New System.Drawing.Size(240, 20)
        Me.txtRE.TabIndex = 18
        Me.txtRE.Text = ""
        '
        'txtAtt
        '
        Me.txtAtt.Location = New System.Drawing.Point(72, 24)
        Me.txtAtt.Name = "txtAtt"
        Me.txtAtt.Size = New System.Drawing.Size(240, 20)
        Me.txtAtt.TabIndex = 18
        Me.txtAtt.Text = ""
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtAtt)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.txtRE)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.txtFaxNo)
        Me.GroupBox1.Enabled = False
        Me.GroupBox1.Location = New System.Drawing.Point(16, 112)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(336, 128)
        Me.GroupBox1.TabIndex = 19
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Fax Header Details"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(8, 88)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(56, 23)
        Me.Label5.TabIndex = 17
        Me.Label5.Text = "Fax No:"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtFaxNo
        '
        Me.txtFaxNo.Location = New System.Drawing.Point(72, 88)
        Me.txtFaxNo.Name = "txtFaxNo"
        Me.txtFaxNo.Size = New System.Drawing.Size(240, 20)
        Me.txtFaxNo.TabIndex = 18
        Me.txtFaxNo.Text = ""
        '
        'dtpStartDate
        '
        Me.dtpStartDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpStartDate.Location = New System.Drawing.Point(88, 48)
        Me.dtpStartDate.Name = "dtpStartDate"
        Me.dtpStartDate.Size = New System.Drawing.Size(104, 20)
        Me.dtpStartDate.TabIndex = 20
        Me.dtpStartDate.Value = New Date(2005, 1, 1, 0, 0, 0, 0)
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(16, 56)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(56, 16)
        Me.Label6.TabIndex = 21
        Me.Label6.Text = "First Date"
        '
        'prtInvoice
        '
        Me.prtInvoice.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.prtInvoice.Location = New System.Drawing.Point(224, 48)
        Me.prtInvoice.Name = "prtInvoice"
        Me.prtInvoice.Size = New System.Drawing.Size(128, 24)
        Me.prtInvoice.TabIndex = 22
        Me.prtInvoice.Text = "Print Invoice Details"
        '
        'frmJobSummaryReport
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(402, 304)
        Me.Controls.Add(Me.prtInvoice)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.dtpStartDate)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.chkFaxHeader)
        Me.Controls.Add(Me.btnPrintPreview)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmbJobs)
        Me.Controls.Add(Me.dtpReportDate)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnClose)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "frmJobSummaryReport"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Print Job Bending Schedule Report"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub frmJobSummaryReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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

    Private Shadows Sub FormClosing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If Not IsNothing(caller) Then
            caller.Show()
        End If

        caller = Nothing
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Close()
    End Sub

    Private Sub PrintBendingSchedules(ByVal jobNo As String, ByVal begDate As Date, ByVal endDate As Date)

        PrintArray = New ArrayList
        Dim p As PageElement
        Dim TKg As String = String.Empty
        Dim sql4compName As String = "SELECT ContractorName,OrderNo,JobName,CompanyName,[Tons or Kilograms] " & _
                    "FROM Job, Contractor,Company " & _
                    "WHERE Job.ContractorNo = Contractor.ContractorNo " & _
                    "AND Company.CompanyNo = Job.CompanyNo " & _
                    "AND Job.JobNo = '" & jobNo & "'"

        Dim ds As New Data.DataSet
        Dim ad As New OleDb.OleDbDataAdapter(sql4compName, DBConnection)
        Dim currJobName As String
        Dim currType, nextY, nextR As String
        Dim currInv As String
        Dim currDate As Date
        Dim currSteel As Double
        ad.Fill(ds)

        If chkFaxHeader.Checked Then
            addFaxHeader(ds.Tables(0).Rows(0).Item("ContractorName").ToString(), ds.Tables(0).Rows(0).Item("CompanyName").ToString())
        End If
        Dim weightHead As String = String.Empty

        p = New PageElement(ds.Tables(0).Rows(0).Item("CompanyName").ToString(), EntryFont, 0, True, True, False)
        PrintArray.Add(p)
        If ds.Tables(0).Rows.Count = 1 Then
            Const d1 As Integer = 110
            TKg = ds.Tables(0).Rows(0).Item("Tons or Kilograms").ToString()

            PrintArray.Add(New PageElement("Job Number:", EntryFont, LeftMargin, False, False, False))
            PrintArray.Add(New PageElement(jobNo, EntryFont, LeftMargin + d1, True, False, False))

            PrintArray.Add(New PageElement("Order Number:", EntryFont, LeftMargin, False, False, False))
            PrintArray.Add(New PageElement(ds.Tables(0).Rows(0).Item("OrderNo").ToString(), EntryFont, LeftMargin + d1, True, False, False))
            currJobName = ds.Tables(0).Rows(0).Item("JobName").ToString()

            PrintArray.Add(New PageElement("Job Name:", EntryFont, LeftMargin, False, False, False))
            PrintArray.Add(New PageElement(currJobName, EntryFont, LeftMargin + d1, True, False, False))
            PrintArray.Add(New PageElement("Contractor:", EntryFont, LeftMargin, False, False, False))
            PrintArray.Add(New PageElement(ds.Tables(0).Rows(0).Item("ContractorName").ToString(), EntryFont, LeftMargin + d1, True, False, False))
            PrintArray.Add(New PageElement("First Date:", EntryFont, LeftMargin, False, False, False))
            PrintArray.Add(New PageElement(begDate.ToShortDateString, EntryFont, LeftMargin + d1, True, False, False))
            PrintArray.Add(New PageElement("Last Date:", EntryFont, LeftMargin, False, False, False))
            PrintArray.Add(New PageElement(endDate.ToShortDateString, EntryFont, LeftMargin + d1, True, False, False))

            PrintArray.Add(New PageElement("", EntryFont, LeftMargin, True, False, False))

            PrintArray.Add(New PageElement("Reinforcing as per schedule:", EntryFontUnderline, LeftMargin, True, False))
            PrintArray.Add(New PageElement("", EntryFont, LeftMargin, True, False, False))
        End If

        If TKg = "T" Or TKg = "t" Then
            weightHead = "TON"
        ElseIf TKg = "Kg" Or TKg = "kg" Then
            weightHead = "kg"
        End If

        Dim RTotals(8) As Double
        Dim YTotals(8) As Double

        Dim typeR(8), typeY(8) As String

        typeR(0) = "R06"
        typeR(1) = "R08"
        typeR(2) = "R10"
        typeR(3) = "R12"
        typeR(4) = "R16"
        typeR(5) = "R20"
        typeR(6) = "R25"
        typeR(7) = "R32"
        typeR(8) = "R40"
        typeY(0) = ""
        typeY(1) = ""
        typeY(2) = "Y10"
        typeY(3) = "Y12"
        typeY(4) = "Y16"
        typeY(5) = "Y20"
        typeY(6) = "Y25"
        typeY(7) = "Y32"
        typeY(8) = "Y40"

        Dim TR As Double = 0
        Dim TY As Double = 0
        Dim gTot As Double = 0

        'Initialise Totals
        Dim i As Integer
        For i = 0 To 8
            RTotals(i) = 0
            YTotals(i) = 0
        Next i

        ' GET ALL THE INVOICES FOR THE SELECTED JOB
        Dim sqlInvoice As Integer = "SELECT DISTINCT InvoiceNo, invDate " & _
                            "FROM Invoice " & _
                            "WHERE InvJobNo = '" & jobNo & "'" & _
                            "AND InvoiceType = 'Cutting Sheet'" & _
                            " AND invDate >= #" + begDate.ToLongDateString + "# " & _
                            " AND invDate <= #" + endDate.ToLongDateString + "# "


        Dim DS4Inv As Data.DataSet = New Data.DataSet
        Dim adapter As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(sqlInvoice, DBConnection)
        adapter.Fill(DS4Inv)

        Dim r, j As Integer

        Dim numInvoices, numRecs As Integer
        numInvoices = DS4Inv.tables(0).rows.count
        If numInvoices = 0 Then
            MessageBox.Show("No cutting sheet invoices found for this job and date")
        End If

        For i = 0 To numInvoices - 1
            currDate = DS4Inv.tables(0).rows(i).item("InvDate")
            currInv = DS4Inv.tables(0).rows(i).item("InvoiceNo").ToString()

            ' GET ALL INVOICE LINES
            Dim sqlInvLine As String = "Select * from InvoiceLine " & _
                "WHERE InvNo = " + currInv + " order by typeCode"

            Dim ds4r As New Data.DataSet
            Dim ad4R As New OleDb.OleDbDataAdapter(sqlInvLine, DBConnection)
            ad4R.Fill(ds4r)
            numRecs = 0
            numRecs = ds4r.Tables(0).Rows.Count()

            'Add up all R & Y types for this invoice
            If numRecs <> 0 Then

                For r = 0 To numRecs - 1
                    Dim found As Boolean = False
                    currType = ds4r.Tables(0).Rows(r).Item("TypeCode").ToString()
                    currType = currType.Substring(0, 3)
                    currSteel = Decimal.Parse(ds4r.Tables(0).Rows(r).Item("Qty"))

                    If TKg.ToLower = "kg" Then
                        currSteel = Format(currSteel, "000.0")

                    Else
                        currSteel = Format(currSteel, "0.000")
                    End If
                    If prtInvoice.Checked Then
                        MessageBox.Show("Inv " & currInv & " type" & currType & " steel " & currSteel)
                    End If

                    For j = 0 To 8
                        nextY = typeY(j)
                        nextR = typeR(j)
                        If nextR.Equals(currType.ToUpper) Then
                            RTotals(j) += currSteel
                            found = True
                            'ADD TO THE CORRECT TOTAL
                            If prtInvoice.Checked Then
                                MessageBox.Show("Add to R " & currType & " " & j & " total " & RTotals(j))
                            End If
                        End If

                        'If currType = nextY Then
                        If nextY.Equals(currType.ToUpper) Then
                            'If typeY(j).Equals(currType.ToUpper) Then
                            YTotals(j) += currSteel
                            found = True
                            If prtInvoice.Checked Then
                                MessageBox.Show("Add to Y " & currType & j & " total " & YTotals(j))
                            End If
                        End If
                    Next

                    If Not found Then
                        MessageBox.Show("Type not found for invoice " & currInv & " and type " & currType)
                    End If
                Next r  'next invoice line
            End If   ' end of if num inv lines > 0
        Next i   'next Invoice

        Dim ci As Integer
        For ci = 0 To 8

            If RTotals(ci) <> 0 Then

                If TKg.ToLower = "kg" Then
                    RTotals(ci) = Format(RTotals(ci), "000.0")
                Else
                    RTotals(ci) = Format(RTotals(ci), "0.000")
                End If
                TR += RTotals(ci)
            End If

            ' ALL Y VALUES
            If YTotals(ci) <> 0 Then

                If TKg.ToLower = "kg" Then
                    YTotals(ci) = Format(YTotals(ci), "000.0")
                Else
                    YTotals(ci) = Format(YTotals(ci), "0.000")
                End If
                TY += YTotals(ci)
            End If

        Next ci

        ci = 40
        Dim m As Integer
        For m = 0 To 8
            ' PRINT R TYPE
            If typeR(m) <> "" Then
                PrintArray.Add(New PageElement(typeR(m), EntryFont, LeftMargin, False, False))
                PrintArray.Add(New PageElement(". . . . . ", EntryFont, LeftMargin + ci, False, False))
                If RTotals(m) > 0 Then
                    PrintArray.Add(New PageElement(RTotals(m).ToString(), EntryFont, LeftMargin + 100, False, False))
                    PrintArray.Add(New PageElement(TKg, EntryFont, LeftMargin + 170, False, False))
                Else
                    PrintArray.Add(New PageElement(" ", EntryFont, LeftMargin + 170, False, False))
                End If

            ElseIf typeR(m) = "" Then
                PrintArray.Add(New PageElement("", EntryFont, LeftMargin, True, False))
            End If

            ' PRINT Y TYPE
            If typeY(m) <> "" Then
                PrintArray.Add(New PageElement(typeY(m), EntryFont, LeftMargin + 400, False, False))
                PrintArray.Add(New PageElement(". . . . . ", EntryFont, LeftMargin + 440, False, False))
                If YTotals(m) > 0 Then
                    PrintArray.Add(New PageElement(YTotals(m).ToString(), EntryFont, LeftMargin + 490, False, False))
                    PrintArray.Add(New PageElement(TKg, EntryFont, LeftMargin + 560, True, False))
                Else
                    PrintArray.Add(New PageElement("", EntryFont, LeftMargin + 560, True, False))
                End If
            ElseIf typeY(m) = "" Then
                PrintArray.Add(New PageElement("", EntryFont, LeftMargin + 400, True, False))
            End If
        Next

        If TKg.ToLower = "kg" Then
            TR = FormatNumber(TR, 1, True)
            TY = FormatNumber(TY, 1, True)
        End If
        If TKg.ToLower = "t" Then
            TR = FormatNumber(TR, 3, True)
            TY = FormatNumber(TY, 3, True)
        End If

        gTot = TR + TY
        If TKg.ToLower = "kg" Then
            gTot = FormatNumber(gTot, 1, True)
        End If
        If TKg.ToLower = "t" Then
            gTot = FormatNumber(gTot, 3, True)
        End If

        PrintArray.Add(New PageElement("", EntryFont, 0, True, False, False))

        PrintArray.Add(New PageElement("Total RMS  :", EntryFont, LeftMargin + 80, False, False, False))
        PrintArray.Add(New PageElement(TR, EntryFont, LeftMargin + 300, False, False, False))
        PrintArray.Add(New PageElement(weightHead, EntryFont, LeftMargin + 380, True, False, False))

        PrintArray.Add(New PageElement("Total HTS  :", EntryFont, LeftMargin + 80, False, False, False))
        PrintArray.Add(New PageElement(TY, EntryFont, LeftMargin + 300, False, False, False))
        PrintArray.Add(New PageElement(weightHead, EntryFont, LeftMargin + 380, True, False, False))


        PrintArray.Add(New PageElement("Grand Total:", EntryFont, LeftMargin + 80, False, False, False))
        PrintArray.Add(New PageElement(gTot, EntryFont, LeftMargin + 300, False, False, False))
        PrintArray.Add(New PageElement(weightHead, EntryFont, LeftMargin + 380, True, False, False))

        PrintArray.Add(New PageElement("<SPACE>", EntryFont, LeftMargin, True, False, False))

        ' GET ALL MESH INVOICES AND PRINT AT THE BOTTOM OF REPORT
        'sql = "SELECT * FROM Invoice INNER JOIN InvoiceLine ON Invoice.InvoiceNo = InvoiceLine.InvNo WHERE invoiceType = 'Mesh' And InvJobNo = '" & cmbJobs.Text & "' AND InvDate <= #" & dtpReportDate.Value & "#"

        Dim sql As String = "SELECT * " & _
                            "FROM Invoice " & _
                            "INNER JOIN InvoiceLine ON Invoice.InvoiceNo = InvoiceLine.InvNo " & _
                            "WHERE InvJobNo = '" & jobNo & "'" & _
                            "AND InvoiceType = 'Mesh'" & _
                            " AND invDate >= #" + begDate.ToLongDateString + "# " & _
                            " AND invDate <= #" + endDate.ToLongDateString + "# "

        ds.Clear()
        ad = New OleDb.OleDbDataAdapter(sql, DBConnection)
        ad.Fill(ds)
        Dim numMesh As Integer
        Dim invDate As Date
        Const A As Integer = 100
        Const B As Integer = 200
        Const C As Integer = 250
        Const D As Integer = 325

        numMesh = ds.Tables(0).Rows.Count

        If numMesh <> 0 Then

            PrintArray.Add(New PageElement("MESH:", EntryFont, LeftMargin, True, False, False))
            PrintArray.Add(New PageElement("--------", EntryFont, LeftMargin, True, False, False))
            PrintArray.Add(New PageElement("INV DATE:", EntryFontUnderline, LeftMargin, False, False, False))
            PrintArray.Add(New PageElement("REF", EntryFontUnderline, LeftMargin + A, False, False, False))
            PrintArray.Add(New PageElement("NO", EntryFontUnderline, LeftMargin + B, False, False, False))
            PrintArray.Add(New PageElement("MxM", EntryFontUnderline, LeftMargin + C, False, False, False))
            PrintArray.Add(New PageElement("REMARKS", EntryFontUnderline, LeftMargin + D, True, False, False))
            Dim QTY, AREA, REMARK As String
            For i = 0 To numMesh - 1
                invDate = ds.Tables(0).Rows(i).Item("InvDate").ToShortDateString()
                PrintArray.Add(New PageElement(invDate, EntryFont, LeftMargin, False, False, False))
                PrintArray.Add(New PageElement(ds.Tables(0).Rows(i).Item("InvRefNum").ToString(), EntryFont, LeftMargin + A, False, False, False))
                QTY = ds.Tables(0).Rows(i).Item("Description").ToString().Substring(0, ds.Tables(0).Rows(i).Item("Description").ToString().IndexOf(" "))
                PrintArray.Add(New PageElement(QTY, EntryFont, LeftMargin + B, False, False, False))
                AREA = ds.Tables(0).Rows(i).Item("Description").ToString().Substring(ds.Tables(0).Rows(i).Item("Description").ToString().IndexOf("=") + 2, ds.Tables(0).Rows(i).Item("Description").ToString().IndexOf("@") - ds.Tables(0).Rows(i).Item("Description").ToString().IndexOf("=") - 2)
                PrintArray.Add(New PageElement(AREA, EntryFont, LeftMargin + +C, False, False, False))
                REMARK = ds.Tables(0).Rows(i).Item("InvoiceHeading").ToString().Substring(23 + ds.Tables(0).Rows(i).Item("InvRefNum").ToString().Length)
                PrintArray.Add(New PageElement(REMARK, EntryFont, LeftMargin + D, True, False, False))
            Next i

        End If


    End Sub

    Private Sub btnPrintPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrintPreview.Click

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


            PrintBendingSchedules(cmbJobs.Text, dtpStartDate.Value.Date, dtpReportDate.Value.Date)


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

    Private Sub addFaxHeader(ByVal ContractorName As String, ByVal CompanyName As String)
        PrintArray.Add(New PageElement("TO:", EntryFont, LeftMargin, False, False, False))
        PrintArray.Add(New PageElement(ContractorName, EntryFont, LeftMargin + 55, True, False, False))
        PrintArray.Add(New PageElement("FROM:", EntryFont, LeftMargin, False, False, False))
        PrintArray.Add(New PageElement(CompanyName, EntryFont, LeftMargin + 55, True, False, False))
        PrintArray.Add(New PageElement("ATT:", EntryFont, LeftMargin, False, False, False))
        PrintArray.Add(New PageElement(txtAtt.Text, EntryFont, LeftMargin + 55, True, False, False))
        PrintArray.Add(New PageElement("RE:", EntryFont, LeftMargin, False, False, False))
        PrintArray.Add(New PageElement(txtRE.Text, EntryFont, LeftMargin + 55, True, False, False))
        PrintArray.Add(New PageElement("FAX NO:", EntryFont, LeftMargin, False, False, False))
        PrintArray.Add(New PageElement(txtFaxNo.Text, EntryFont, LeftMargin + 55, True, False, False))
        PrintArray.Add(New PageElement("", EntryFont, LeftMargin, True, False, False))

    End Sub

    Private Sub PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles DocumentToPrint.PrintPage

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

    Private Sub chkFaxHeader_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkFaxHeader.CheckedChanged
        GroupBox1.Enabled = chkFaxHeader.Checked
    End Sub

    Private Sub dtpReportDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpReportDate.ValueChanged

    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles prtInvoice.CheckedChanged

    End Sub
End Class
