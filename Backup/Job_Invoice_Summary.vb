Public Class frmPrintReinforcingSummary
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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cmbJobs As System.Windows.Forms.ComboBox
    Friend WithEvents btnPrintPreview As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents dtpStart As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpEnd As System.Windows.Forms.DateTimePicker
    Friend WithEvents DocumentToPrint As System.Drawing.Printing.PrintDocument
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents chkCS As System.Windows.Forms.CheckBox
    Friend WithEvents chkM As System.Windows.Forms.CheckBox
    Friend WithEvents chkS As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.dtpStart = New System.Windows.Forms.DateTimePicker
        Me.dtpEnd = New System.Windows.Forms.DateTimePicker
        Me.cmbJobs = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.btnPrintPreview = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.DocumentToPrint = New System.Drawing.Printing.PrintDocument
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.chkCS = New System.Windows.Forms.CheckBox
        Me.chkM = New System.Windows.Forms.CheckBox
        Me.chkS = New System.Windows.Forms.CheckBox
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'dtpStart
        '
        Me.dtpStart.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpStart.Location = New System.Drawing.Point(88, 48)
        Me.dtpStart.Name = "dtpStart"
        Me.dtpStart.Size = New System.Drawing.Size(104, 20)
        Me.dtpStart.TabIndex = 1
        Me.dtpStart.Value = New Date(2005, 1, 1, 0, 0, 0, 0)
        '
        'dtpEnd
        '
        Me.dtpEnd.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpEnd.Location = New System.Drawing.Point(88, 80)
        Me.dtpEnd.Name = "dtpEnd"
        Me.dtpEnd.Size = New System.Drawing.Size(104, 20)
        Me.dtpEnd.TabIndex = 2
        '
        'cmbJobs
        '
        Me.cmbJobs.Location = New System.Drawing.Point(88, 16)
        Me.cmbJobs.Name = "cmbJobs"
        Me.cmbJobs.Size = New System.Drawing.Size(104, 21)
        Me.cmbJobs.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 23)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Job No.:"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 48)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(64, 23)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Start Date:"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(16, 80)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(64, 23)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "End Date:"
        '
        'btnPrintPreview
        '
        Me.btnPrintPreview.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnPrintPreview.Location = New System.Drawing.Point(64, 128)
        Me.btnPrintPreview.Name = "btnPrintPreview"
        Me.btnPrintPreview.Size = New System.Drawing.Size(120, 24)
        Me.btnPrintPreview.TabIndex = 3
        Me.btnPrintPreview.Text = "Print Preview..."
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(208, 128)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(112, 24)
        Me.btnClose.TabIndex = 3
        Me.btnClose.Text = "Close"
        '
        'DocumentToPrint
        '
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.chkCS)
        Me.GroupBox1.Controls.Add(Me.chkM)
        Me.GroupBox1.Controls.Add(Me.chkS)
        Me.GroupBox1.Location = New System.Drawing.Point(208, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(184, 96)
        Me.GroupBox1.TabIndex = 4
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Summary Type"
        '
        'chkCS
        '
        Me.chkCS.Checked = True
        Me.chkCS.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkCS.Location = New System.Drawing.Point(16, 16)
        Me.chkCS.Name = "chkCS"
        Me.chkCS.Size = New System.Drawing.Size(152, 24)
        Me.chkCS.TabIndex = 0
        Me.chkCS.Text = "Reinforcing Summary"
        '
        'chkM
        '
        Me.chkM.Checked = True
        Me.chkM.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkM.Location = New System.Drawing.Point(16, 40)
        Me.chkM.Name = "chkM"
        Me.chkM.Size = New System.Drawing.Size(152, 24)
        Me.chkM.TabIndex = 0
        Me.chkM.Text = "Mesh Summary"
        '
        'chkS
        '
        Me.chkS.Checked = True
        Me.chkS.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkS.Location = New System.Drawing.Point(16, 64)
        Me.chkS.Name = "chkS"
        Me.chkS.Size = New System.Drawing.Size(152, 24)
        Me.chkS.TabIndex = 0
        Me.chkS.Text = "Sundries Summary"
        '
        'frmPrintReinforcingSummary
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(402, 168)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnPrintPreview)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmbJobs)
        Me.Controls.Add(Me.dtpStart)
        Me.Controls.Add(Me.dtpEnd)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btnClose)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "frmPrintReinforcingSummary"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Summary Detail Selection"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private CallingForm As Object

    Public Sub New(ByVal caller As Object)
        MyBase.New()
        InitializeComponent()

        CallingForm = caller
    End Sub


    Private Sub FormClosing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If Not IsNothing(CallingForm) Then
            CallingForm.Show()
        End If

        CallingForm = Nothing
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Close()
    End Sub

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
    Dim curArrayPos = 0
    Dim curpagenum = 1
    Dim TopMargin = 70
    Dim LeftMargin = 20
    Dim RightMargin = 90
    Dim BottomMargin = 90
    Dim PageWidth = 873
    Dim ReportType
    Dim mes
    Dim vatperc As String
    'Dim All_Is_OK As Boolean = True
    Dim HasMesh As Boolean = False
    Dim HasSundry As Boolean = False
    Dim HasCut As Boolean = False
    Dim currDate As Date = Today
#End Region

    Private Function toRand(ByVal input As String, ByVal r As Boolean) As String

        Dim iput As Double

        Try
            iput = Double.Parse(input)
            If r Then
                Return Format(iput, "R #,###,###,##0.00")
            Else
                Return Format(iput, "#,###,###,##0.00")
            End If

        Catch ex As Exception
            MessageBox.Show("Error with input string.", "Cannot convert to Rand format.", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
        End Try

    End Function

    Private Sub AddCompNameRegNoVatNo(ByVal compNum As String, ByVal includeTelAndAddress As Boolean)
        'Get Company Details

        Dim sql = "SELECT * FROM Company WHERE CompanyNo = '" + compNum + "'"
        Dim dataset As New Data.DataSet
        Dim adapter As New OleDb.OleDbDataAdapter(sql, DBConnection)
        adapter.Fill(dataset)



        field = New PageElement(dataset.Tables(0).Rows(0).Item("CompanyName").ToString(), EntryFont, 0, True, True)
        PrintArray.Add(field)
        field = New PageElement("Reg No. " + dataset.Tables(0).Rows(0).Item("RegNo").ToString(), EntryFont, 0, True, True)
        PrintArray.Add(field)
        field = New PageElement("Vat No. " + dataset.Tables(0).Rows(0).Item("VatNo").ToString(), EntryFont, 0, True, True)
        PrintArray.Add(field)
        If includeTelAndAddress Then
            field = New PageElement("Tel:", EntryFont, 540, False, False)
            PrintArray.Add(field)
            field = New PageElement(dataset.Tables(0).Rows(0).Item("Telephone").ToString(), EntryFont, 580, True, False)
            PrintArray.Add(field)
            field = New PageElement("Fax:", EntryFont, 540, False, False)
            PrintArray.Add(field)
            field = New PageElement(dataset.Tables(0).Rows(0).Item("Fax").ToString(), EntryFont, 580, True, False)
            PrintArray.Add(field)
            field = New PageElement(dataset.Tables(0).Rows(0).Item("Address").ToString(), EntryFont, 540, True, False)
            PrintArray.Add(field)
            field = New PageElement(dataset.Tables(0).Rows(0).Item("AddressLine2").ToString(), EntryFont, 540, True, False)

            If field.Text <> "" Then
                PrintArray.Add(field)
            End If

            field = New PageElement(dataset.Tables(0).Rows(0).Item("AddressLine3").ToString(), EntryFont, 540, True, False)
            If field.Text <> "" Then
                PrintArray.Add(field)
            End If
            field = New PageElement(dataset.Tables(0).Rows(0).Item("AddressLine4").ToString(), EntryFont, 540, True, False)
            If field.Text <> "" Then
                PrintArray.Add(field)
            End If
            field = New PageElement(dataset.Tables(0).Rows(0).Item("PostalCode").ToString(), EntryFont, 540, True, False)
            PrintArray.Add(field)
        End If

        mes = New PageElement(dataset.Tables(0).Rows(0).Item("Message").ToString(), EntryFont, LeftMargin, True, False)

        vatperc = dataset.Tables(0).Rows(0).Item("VatPerc").ToString()
        vatperc = (Decimal.Round(Decimal.Parse(vatperc) * 100, 0)).ToString + "%"



    End Sub

    Const ll = 15  'Last Column
    Const tt = 385 + 50 'T or Kg Column
    Const wv = tt + 15 ' WeightValue Column
    Const iv = 285 + 40 ' Invoice Value Column
    Const vt = 205 + 20 ' Vat Column
    Const ta = 115 + 5 ' total Column
    Const idt = 70 ' date column

    Private Sub ReinforcingSummaryPrint(ByVal compNum As String, ByVal jobNo As String, ByVal StartDate As Date, ByVal EndDate As Date)

        Try
            Dim sql = "SELECT * FROM (((invoice inner join job on Job.JobNo = Invoice.invJobNo)" & _
            "INNER JOIN Contractor on Contractor.ContractorNo = job.ContractorNo) " & _
            "INNER JOIN CuttingSheet on cuttingSheet.invoiceNo = invoice.invoiceNo)" & _
            " WHERE job.JobNo = '" + jobNo + "' " & _
            "AND invoice.invdate >= #" + StartDate.ToLongDateString + "# " & _
            "AND invoice.invdate <= #" + EndDate.ToLongDateString + "# " & _
            "ORDER BY invoice.invoiceNo"

            Dim DataSet = New Data.DataSet
            Dim adapter = New OleDb.OleDbDataAdapter(sql, DBConnection)
            adapter.Fill(DataSet)

            If DataSet.tables(0).rows.count = 0 Then
                ' All_Is_OK = False
                HasCut = False
                Exit Sub
            Else
                AddCompNameRegNoVatNo(compNum, True)
                ' All_Is_OK = True
                HasCut = True
            End If

            field = New PageElement(DataSet.Tables(0).Rows(0).Item("ContractorName").ToString(), EntryFont, LeftMargin, True, False)
            PrintArray.Add(field)
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("AddressLine1").ToString(), EntryFont, LeftMargin, True, False)
            If field.Text <> "" Then
                PrintArray.Add(field)
            End If
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("AddressLine2").ToString(), EntryFont, LeftMargin, True, False)
            If field.Text <> "" Then
                PrintArray.Add(field)
            End If
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("AddressLine3").ToString(), EntryFont, LeftMargin, True, False)
            If field.Text <> "" Then
                PrintArray.Add(field)
            End If
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("AddressLine4").ToString(), EntryFont, LeftMargin, True, False)
            If field.Text <> "" Then
                PrintArray.Add(field)
            End If
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("PostalCode").ToString(), EntryFont, LeftMargin, True, False)
            PrintArray.Add(field)
            field = New PageElement("Vat No. ", EntryFont, LeftMargin, False, False)
            PrintArray.Add(field)
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("Vat No").ToString(), EntryFont, LeftMargin + 50, 40, False)
            PrintArray.Add(field)

            field = New PageElement("JOB NO:", EntryFont, LeftMargin, False, False)
            PrintArray.Add(field)
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("JobNo").ToString(), EntryFont, 200, False, False)
            PrintArray.Add(field)
            field = New PageElement("DATE:", EntryFont, 540, False, False)
            PrintArray.Add(field)
            field = New PageElement(currDate.ToShortDateString(), EntryFont, 630, True, False)

            ' field = New PageElement(StartDate.ToShortDateString, EntryFont, 630, True, False)
            PrintArray.Add(field)
            field = New PageElement("CONTRACT:", EntryFont, LeftMargin, False, False)
            PrintArray.Add(field)
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("JobName").ToString(), EntryFont, 200, True, False)
            PrintArray.Add(field)
            field = New PageElement("ORDER NO:", EntryFont, LeftMargin, False, False)
            PrintArray.Add(field)
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("OrderNo").ToString(), EntryFont, 200, True, False)
            PrintArray.Add(field)

            field = New PageElement("REINFORCING INVOICE SUMMARY", EntryFont, 0, True, True, False)
            PrintArray.Add(field)
            field = New PageElement("<HR/BLACK>", EntryFont, LeftMargin, -15, False)
            PrintArray.Add(field)
            field = New PageElement("ACCUM", EntryFontUnderline, PageWidth - RightMargin - ll, True, False, True)
            PrintArray.Add(field)
            field = New PageElement("INV. NO", EntryFontUnderline, LeftMargin, False, False, False)
            PrintArray.Add(field)
            field = New PageElement("INV. DATE", EntryFontUnderline, LeftMargin + idt, False, False, False)
            PrintArray.Add(field)
            field = New PageElement("CUT SHEET", EntryFontUnderline, 205, False, False, False)
            PrintArray.Add(field)
            Dim WH, tKg As String
            tKg = DataSet.tables(0).rows(0).item("Tons Or Kilograms").ToString()
            If tKg = "T" Then
                WH = "TONS"
            ElseIf tKg = "Kg" Then
                WH = "KGS"
                tKg = tKg.ToLower
            End If

            field = New PageElement(WH, EntryFontUnderline, PageWidth - RightMargin - tt, False, False, True)
            PrintArray.Add(field)
            field = New PageElement("INV. VALUE", EntryFontUnderline, PageWidth - RightMargin - iv, False, False, True)
            PrintArray.Add(field)
            field = New PageElement("VAT", EntryFontUnderline, PageWidth - RightMargin - vt, False, False, True)
            PrintArray.Add(field)
            field = New PageElement("TOTAL", EntryFontUnderline, PageWidth - RightMargin - ta, False, False, True)
            PrintArray.Add(field)
            field = New PageElement("TOTAL", EntryFontUnderline, PageWidth - RightMargin - ll, True, False, True)
            PrintArray.Add(field)
            Dim t = 0
            Dim GrandWeight = 0
            Dim GrandValue = 0
            Dim GrandVat = 0
            Dim GrandNett = 0
            Dim addMonthLine As Boolean
            Dim MonthWeight = 0
            Dim MonthValue = 0
            Dim MonthVat = 0
            Dim MonthNett = 0
            Dim curMonth, nextMonth, curYear, NextYear

            Dim accum = 0


            For t = 0 To DataSet.tables(0).rows.count - 1
                'tKg = DataSet.tables(0).rows(t).item("Tons Or Kilograms").ToString()

                addMonthLine = False
                Dim currInvNo As String

                currInvNo = DataSet.tables(0).rows(t).item("invoice.InvoiceNo").ToString()

                field = New PageElement(currInvNo, EntryFont, LeftMargin, False, False, False)
                PrintArray.Add(field)

                Dim msql = "Select * from InvoiceLine " & _
                "WHERE InvNo = " + currInvNo + " order by [Line#]"

                Dim mDataSet = New Data.DataSet
                Dim madapter = New OleDb.OleDbDataAdapter(msql, DBConnection)
                madapter.Fill(mDataSet)
                Dim x
                Dim itemQty = 0
                Dim WeightTotal = 0
                Dim itemType As String
                For x = 0 To mDataSet.Tables(0).Rows.Count - 1
                    itemType = (mDataSet.Tables(0).Rows(x).Item("TypeCode"))
                    itemQty = Decimal.Parse(mDataSet.Tables(0).Rows(x).Item("Qty"))

                    If tKg = "kg" Then
                        itemQty = Format(itemQty, "000.0")
                    Else
                        itemQty = Format(itemQty, "0.000")
                    End If

                    WeightTotal += itemQty
                Next x

                If tKg = "kg" Then
                    WeightTotal = Format(WeightTotal, "000.0")
                Else
                    WeightTotal = Format(WeightTotal, "0.000")
                End If

                field = New PageElement(DataSet.tables(0).rows(t).item("invDate").ToShortDateString(), EntryFont, LeftMargin + idt, False, False, False)
                PrintArray.Add(field)

                curMonth = (Date.Parse(DataSet.tables(0).rows(t).item("invDate").ToShortDateString())).Month
                curYear = (Date.Parse(DataSet.tables(0).rows(t).item("invDate").ToShortDateString())).Year

                If (t + 1 <= DataSet.tables(0).rows.count - 1) Then 'IF there is another invoice
                    nextMonth = Date.Parse(DataSet.tables(0).rows(t + 1).item("invDate").ToShortDateString()).Month
                    NextYear = Date.Parse(DataSet.tables(0).rows(t + 1).item("invDate").ToShortDateString()).Year
                    If (nextMonth <> curMonth Or NextYear <> curYear) Then
                        addMonthLine = True
                    End If
                End If

                If t = DataSet.tables(0).rows.count - 1 Then
                    addMonthLine = True
                End If

                field = New PageElement(DataSet.tables(0).rows(t).item("CutSheetNo").ToString(), EntryFont, 230, False, False, False)
                PrintArray.Add(field)
                field = New PageElement(WeightTotal, EntryFont, PageWidth - RightMargin - wv, False, False, True)
                PrintArray.Add(field)
                MonthWeight += WeightTotal
                field = New PageElement(tKg, EntryFont, PageWidth - RightMargin - tt, False, False, True)
                PrintArray.Add(field)
                field = New PageElement(toRand(DataSet.tables(0).rows(t).item("invTotal").ToString(), False), EntryFont, PageWidth - RightMargin - iv, False, False, True)
                PrintArray.Add(field)
                MonthValue += Decimal.Parse(DataSet.tables(0).rows(t).item("invTotal").ToString())
                field = New PageElement(toRand(DataSet.tables(0).rows(t).item("invVatAmt").ToString(), False), EntryFont, PageWidth - RightMargin - vt, False, False, True)
                PrintArray.Add(field)
                MonthVat += Decimal.Parse(DataSet.tables(0).rows(t).item("invVatAmt").ToString())
                field = New PageElement(toRand(DataSet.tables(0).rows(t).item("invNett").ToString(), False), EntryFont, PageWidth - RightMargin - ta, False, False, True)
                PrintArray.Add(field)
                accum += DataSet.tables(0).rows(t).item("invNett")
                field = New PageElement(toRand(accum, False), EntryFont, PageWidth - RightMargin - ll, True, False, True)
                PrintArray.Add(field)

                MonthNett += Decimal.Parse(DataSet.tables(0).rows(t).item("invNett").ToString())
                If (addMonthLine) Then
                    field = New PageElement(MonthName(curMonth) & " Total:", EntryFont, 170, False, False, False)
                    PrintArray.Add(field)
                    'BMS ADDED 2/6/2005 FOR TOTAL FOR KG TO PRINT ONE DECIMAL
                    If tKg = "kg" Then
                        MonthWeight = Format(MonthWeight, "000.0")
                    End If

                    field = New PageElement(MonthWeight, EntryFontUnderline, PageWidth - RightMargin - wv, False, False, True)
                    PrintArray.Add(field)
                    field = New PageElement(tKg, EntryFontUnderline, PageWidth - RightMargin - tt, False, False, True)
                    PrintArray.Add(field)
                    field = New PageElement(toRand(MonthValue, False), EntryFontUnderline, PageWidth - RightMargin - iv, False, False, True)
                    PrintArray.Add(field)
                    field = New PageElement(toRand(MonthVat, False), EntryFontUnderline, PageWidth - RightMargin - vt, False, False, True)
                    PrintArray.Add(field)
                    field = New PageElement(toRand(MonthNett, False), EntryFontUnderline, PageWidth - RightMargin - ll, True, False, True)
                    PrintArray.Add(field)
                    addMonthLine = False
                    GrandWeight += MonthWeight
                    GrandValue += MonthValue
                    GrandVat += MonthVat
                    GrandNett += MonthNett
                    MonthWeight = 0
                    MonthValue = 0
                    MonthVat = 0
                    MonthNett = 0
                    ' tKg = DataSet.tables(0).rows(t).item("Tons Or Kilograms").ToString()
                End If
            Next t

            field = New PageElement("Summary Total:", EntryFont, 170, False, False, False)
            PrintArray.Add(field)
            field = New PageElement(GrandWeight, EntryFont, PageWidth - RightMargin - wv, False, False, True)
            PrintArray.Add(field)
            field = New PageElement(tKg, EntryFont, PageWidth - RightMargin - tt, False, False, True)
            PrintArray.Add(field)
            field = New PageElement(toRand(GrandValue, True), EntryFont, PageWidth - RightMargin - iv, False, False, True)
            PrintArray.Add(field)
            field = New PageElement(toRand(GrandVat, True), EntryFont, PageWidth - RightMargin - vt, False, False, True)
            PrintArray.Add(field)
            field = New PageElement(toRand(GrandNett, True), EntryFont, PageWidth - RightMargin - ll, True, False, True)
            PrintArray.Add(field)

        Catch err As Exception
            MessageBox.Show(err.StackTrace & " - " & err.Message)
        End Try

    End Sub

    Private Sub MeshSummaryPrint(ByVal compNum As String, ByVal jobNo As String, ByVal StartDate As Date, ByVal EndDate As Date)


        'Get Contractor Details

        Try

            Dim sql = "Select * from ((invoice inner join job on Job.JobNo = Invoice.invJobNo) " & _
            "INNER JOIN Contractor on Contractor.ContractorNo = job.ContractorNo) where job.JobNo = '" + jobNo + "' AND invoice.invdate BETWEEN #" + StartDate.ToLongDateString + "# AND #" + EndDate.ToLongDateString + "# AND invoice.invoiceType = 'Mesh' order by invoice.invoiceNo"
            Dim DataSet As Data.DataSet = New Data.DataSet
            Dim adapter = New OleDb.OleDbDataAdapter(sql, DBConnection)
            adapter.Fill(DataSet)

            If DataSet.Tables(0).Rows.Count = 0 Then
                ' MessageBox.Show("No invoices found matching the job no. and dates selected.", "No records found", MessageBoxButtons.OK)
                '  All_Is_OK = False
                HasMesh = False
                Exit Sub
            Else
                ' All_Is_OK = True
                If HasCut And chkCS.Checked Then
                    PrintArray.Add(New PageElement("<PAGE BREAK>", EntryFont, 0, 0))
                End If
                AddCompNameRegNoVatNo(compNum, True)
                HasMesh = True
            End If

            field = New PageElement(DataSet.Tables(0).Rows(0).Item("ContractorName").ToString(), EntryFont, LeftMargin, True, False)
            PrintArray.Add(field)
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("AddressLine1").ToString(), EntryFont, LeftMargin, True, False)
            If field.Text <> "" Then
                PrintArray.Add(field)
            End If
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("AddressLine2").ToString(), EntryFont, LeftMargin, True, False)
            If field.Text <> "" Then
                PrintArray.Add(field)
            End If
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("AddressLine3").ToString(), EntryFont, LeftMargin, True, False)
            If field.Text <> "" Then
                PrintArray.Add(field)
            End If
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("AddressLine4").ToString(), EntryFont, LeftMargin, True, False)
            If field.Text <> "" Then
                PrintArray.Add(field)
            End If
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("PostalCode").ToString(), EntryFont, LeftMargin, True, False)
            PrintArray.Add(field)
            field = New PageElement("Vat No. ", EntryFont, LeftMargin, False, False)
            PrintArray.Add(field)
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("Vat No").ToString(), EntryFont, LeftMargin + 50, 40, False)
            PrintArray.Add(field)

            field = New PageElement("JOB NO:", EntryFont, LeftMargin, False, False)
            PrintArray.Add(field)
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("JobNo").ToString(), EntryFont, 200, False, False)
            PrintArray.Add(field)
            field = New PageElement("DATE:", EntryFont, 540, False, False)
            PrintArray.Add(field)
            field = New PageElement(currDate.ToShortDateString(), EntryFont, 630, True, False)
            PrintArray.Add(field)
            field = New PageElement("CONTRACT:", EntryFont, LeftMargin, False, False)
            PrintArray.Add(field)
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("JobName").ToString(), EntryFont, 200, True, False)
            PrintArray.Add(field)
            field = New PageElement("ORDER NO:", EntryFont, LeftMargin, False, False)
            PrintArray.Add(field)
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("OrderNo").ToString(), EntryFont, 200, True, False)
            PrintArray.Add(field)

            field = New PageElement("MESH SUMMARY", EntryFont, 0, True, True, False)
            PrintArray.Add(field)
            field = New PageElement("<HR/BLACK>", EntryFont, LeftMargin, -15, False)
            PrintArray.Add(field)
            field = New PageElement("ACCUM", EntryFontUnderline, PageWidth - RightMargin - ll, True, False, True)
            PrintArray.Add(field)
            field = New PageElement("INV. NO", EntryFontUnderline, LeftMargin, False, False, False)
            PrintArray.Add(field)
            field = New PageElement("INV. DATE", EntryFontUnderline, LeftMargin + idt, False, False, False)
            PrintArray.Add(field)
            field = New PageElement("INV. REF", EntryFontUnderline, 210, False, False, False)
            PrintArray.Add(field)
            field = New PageElement("AREA m²", EntryFontUnderline, PageWidth - RightMargin - tt, False, False, True)
            PrintArray.Add(field)
            field = New PageElement("INV. VALUE", EntryFontUnderline, PageWidth - RightMargin - iv, False, False, True)
            PrintArray.Add(field)
            field = New PageElement("VAT", EntryFontUnderline, PageWidth - RightMargin - vt, False, False, True)
            PrintArray.Add(field)
            field = New PageElement("TOTAL", EntryFontUnderline, PageWidth - RightMargin - ta, False, False, True)
            PrintArray.Add(field)
            field = New PageElement("TOTAL", EntryFontUnderline, PageWidth - RightMargin - ll, True, False, True)
            PrintArray.Add(field)
            Dim t = 0
            Dim GrandArea = 0
            Dim GrandValue = 0
            Dim GrandVat = 0
            Dim GrandNett = 0
            Dim addMonthLine As Boolean
            Dim MonthArea = 0
            Dim MonthValue = 0
            Dim MonthVat = 0
            Dim MonthNett = 0
            Dim curMonth, nextMonth, curYear, NextYear
            'Dim tKg = "Tons"
            Dim accum = 0


            For t = 0 To DataSet.Tables(0).Rows.Count - 1

                addMonthLine = False

                field = New PageElement(DataSet.Tables(0).Rows(t).Item("InvoiceNo").ToString(), EntryFont, LeftMargin, False, False, False)
                PrintArray.Add(field)

                Dim msql = "Select * from InvoiceLine where InvNo = " + DataSet.Tables(0).Rows(t).Item("InvoiceNo").ToString() + " order by [Line#]"
                Dim mDataSet = New Data.DataSet
                Dim madapter = New OleDb.OleDbDataAdapter(msql, DBConnection)
                madapter.Fill(mDataSet)
                Dim x
                Dim AreaTotal = 0
                For x = 0 To mDataSet.Tables(0).Rows.Count - 1
                    Try
                        AreaTotal += Decimal.Parse(mDataSet.Tables(0).Rows(x).Item("Description").ToString().Substring(mDataSet.Tables(0).Rows(x).Item("Description").ToString().IndexOf("=") + 2, mDataSet.Tables(0).Rows(x).Item("Description").ToString().IndexOf("@") - mDataSet.Tables(0).Rows(x).Item("Description").ToString().IndexOf("=") - 6))
                    Catch edd As Exception
                        MessageBox.Show("An invalid area was encountered.", "ERROR WITH INVOICE LINE")
                    End Try
                Next x

                field = New PageElement(DataSet.Tables(0).Rows(t).Item("invDate").ToShortDateString(), EntryFont, LeftMargin + idt, False, False, False)
                PrintArray.Add(field)

                curMonth = (Date.Parse(DataSet.Tables(0).Rows(t).Item("invDate").ToShortDateString())).Month
                curYear = (Date.Parse(DataSet.Tables(0).Rows(t).Item("invDate").ToShortDateString())).Year

                If (t + 1 <= DataSet.Tables(0).Rows.Count - 1) Then 'IF there is another invoice
                    nextMonth = Date.Parse(DataSet.Tables(0).Rows(t + 1).Item("invDate").ToShortDateString()).Month
                    NextYear = Date.Parse(DataSet.Tables(0).Rows(t + 1).Item("invDate").ToShortDateString()).Year
                    If (nextMonth <> curMonth Or NextYear <> curYear) Then
                        addMonthLine = True
                    End If
                End If

                If t = DataSet.Tables(0).Rows.Count - 1 Then
                    addMonthLine = True
                End If

                PrintArray.Add(New PageElement(DataSet.Tables(0).Rows(t).Item("invRefNum").ToString(), EntryFont, 230, False, False, False))

                field = New PageElement(AreaTotal, EntryFont, PageWidth - RightMargin - wv, False, False, True)
                PrintArray.Add(field)
                MonthArea += AreaTotal
                field = New PageElement(toRand(DataSet.Tables(0).Rows(t).Item("invTotal").ToString(), False), EntryFont, PageWidth - RightMargin - iv, False, False, True)
                PrintArray.Add(field)
                MonthValue += Decimal.Parse(DataSet.Tables(0).Rows(t).Item("invTotal").ToString())
                field = New PageElement(toRand(DataSet.Tables(0).Rows(t).Item("invVatAmt").ToString(), False), EntryFont, PageWidth - RightMargin - vt, False, False, True)
                PrintArray.Add(field)
                MonthVat += Decimal.Parse(DataSet.Tables(0).Rows(t).Item("invVatAmt").ToString())
                field = New PageElement(toRand(DataSet.Tables(0).Rows(t).Item("invNett").ToString(), False), EntryFont, PageWidth - RightMargin - ta, False, False, True)
                PrintArray.Add(field)
                accum += DataSet.Tables(0).Rows(t).Item("invNett")
                field = New PageElement(toRand(accum, False), EntryFont, PageWidth - RightMargin - ll, True, False, True)
                PrintArray.Add(field)

                MonthNett += Decimal.Parse(DataSet.Tables(0).Rows(t).Item("invNett").ToString())
                If (addMonthLine) Then
                    field = New PageElement(MonthName(curMonth) & " Total:", EntryFont, 170, False, False, False)
                    PrintArray.Add(field)
                    field = New PageElement(MonthArea, EntryFontUnderline, PageWidth - RightMargin - wv, False, False, True)
                    PrintArray.Add(field)
                    field = New PageElement(toRand(MonthValue, False), EntryFontUnderline, PageWidth - RightMargin - iv, False, False, True)
                    PrintArray.Add(field)
                    field = New PageElement(toRand(MonthVat, False), EntryFontUnderline, PageWidth - RightMargin - vt, False, False, True)
                    PrintArray.Add(field)
                    field = New PageElement(toRand(MonthNett, False), EntryFontUnderline, PageWidth - RightMargin - ll, True, False, True)
                    PrintArray.Add(field)
                    addMonthLine = False
                    GrandArea += MonthArea
                    GrandValue += MonthValue
                    GrandVat += MonthVat
                    GrandNett += MonthNett
                    MonthArea = 0
                    MonthValue = 0
                    MonthVat = 0
                    MonthNett = 0
                    'tKg = DataSet.Tables(0).Rows(t).Item("Tons Or Kilograms").ToString()
                End If
            Next t

            field = New PageElement("Summary Total:", EntryFont, 170, False, False, False)
            PrintArray.Add(field)
            field = New PageElement(GrandArea, EntryFont, PageWidth - RightMargin - wv, False, False, True)
            PrintArray.Add(field)
            field = New PageElement(toRand(GrandValue, True), EntryFont, PageWidth - RightMargin - iv, False, False, True)
            PrintArray.Add(field)
            field = New PageElement(toRand(GrandVat, True), EntryFont, PageWidth - RightMargin - vt, False, False, True)
            PrintArray.Add(field)
            field = New PageElement(toRand(GrandNett, True), EntryFont, PageWidth - RightMargin - ll, True, False, True)
            PrintArray.Add(field)



        Catch err As Exception
            MessageBox.Show(err.StackTrace & " - " & err.Message)
        End Try



    End Sub

    Private Sub SundriesSummaryPrint(ByVal compNum As String, ByVal jobNo As String, ByVal StartDate As Date, ByVal EndDate As Date)


        'Get Contractor Details

        Try
            Dim currDate As Date = Today

            Dim sql = "Select * from ((invoice inner join job on Job.JobNo = Invoice.invJobNo) inner join Contractor on Contractor.ContractorNo = job.ContractorNo) where job.JobNo = '" + jobNo + "' AND invoice.invdate BETWEEN #" + StartDate.ToLongDateString + "# AND #" + EndDate.ToLongDateString + "# AND invoice.invoiceType = 'Sundry' order by invoice.invoiceNo"
            Dim DataSet = New Data.DataSet
            Dim adapter = New OleDb.OleDbDataAdapter(sql, DBConnection)
            adapter.Fill(DataSet)

            If DataSet.tables(0).rows.count = 0 Then
                'MessageBox.Show("No invoices found matching the job no. and dates selected.", "No records found", MessageBoxButtons.OK)
                ' All_Is_OK = False
                Exit Sub
            Else
                HasSundry = True
                If HasCut Or HasMesh Then
                    PrintArray.Add(New PageElement("<PAGE BREAK>", EntryFont, 0, 0))
                End If
                AddCompNameRegNoVatNo(compNum, True)
                ' All_Is_OK = True
            End If

            field = New PageElement(DataSet.Tables(0).Rows(0).Item("ContractorName").ToString(), EntryFont, LeftMargin, True, False)
            PrintArray.Add(field)
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("AddressLine1").ToString(), EntryFont, LeftMargin, True, False)
            If field.Text <> "" Then
                PrintArray.Add(field)
            End If
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("AddressLine2").ToString(), EntryFont, LeftMargin, True, False)
            If field.Text <> "" Then
                PrintArray.Add(field)
            End If
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("AddressLine3").ToString(), EntryFont, LeftMargin, True, False)
            If field.Text <> "" Then
                PrintArray.Add(field)
            End If
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("AddressLine4").ToString(), EntryFont, LeftMargin, True, False)
            If field.Text <> "" Then
                PrintArray.Add(field)
            End If
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("PostalCode").ToString(), EntryFont, LeftMargin, True, False)
            PrintArray.Add(field)
            field = New PageElement("Vat No. ", EntryFont, LeftMargin, False, False)
            PrintArray.Add(field)
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("Vat No").ToString(), EntryFont, LeftMargin + 50, 40, False)
            PrintArray.Add(field)

            field = New PageElement("JOB NO:", EntryFont, LeftMargin, False, False)
            PrintArray.Add(field)
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("JobNo").ToString(), EntryFont, 200, False, False)
            PrintArray.Add(field)
            field = New PageElement("DATE:", EntryFont, 540, False, False)
            PrintArray.Add(field)
            field = New PageElement(currDate.ToShortDateString(), EntryFont, 630, True, False)
            PrintArray.Add(field)
            field = New PageElement("CONTRACT:", EntryFont, LeftMargin, False, False)
            PrintArray.Add(field)
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("JobName").ToString(), EntryFont, 200, True, False)
            PrintArray.Add(field)
            field = New PageElement("ORDER NO:", EntryFont, LeftMargin, False, False)
            PrintArray.Add(field)
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("OrderNo").ToString(), EntryFont, 200, True, False)
            PrintArray.Add(field)

            field = New PageElement("SUNDRIES SUMMARY", EntryFont, 0, True, True, False)
            PrintArray.Add(field)
            field = New PageElement("<HR/BLACK>", EntryFont, LeftMargin, -15, False)
            PrintArray.Add(field)
            field = New PageElement("ACCUM", EntryFontUnderline, PageWidth - RightMargin - ll, True, False, True)
            PrintArray.Add(field)
            field = New PageElement("INV. NO", EntryFontUnderline, LeftMargin, False, False, False)
            PrintArray.Add(field)
            field = New PageElement("INV. DATE", EntryFontUnderline, LeftMargin + idt, False, False, False)
            PrintArray.Add(field)
            field = New PageElement("INV. VALUE", EntryFontUnderline, PageWidth - RightMargin - iv, False, False, True)
            PrintArray.Add(field)
            field = New PageElement("VAT", EntryFontUnderline, PageWidth - RightMargin - vt, False, False, True)
            PrintArray.Add(field)
            field = New PageElement("TOTAL", EntryFontUnderline, PageWidth - RightMargin - ta, False, False, True)
            PrintArray.Add(field)
            field = New PageElement("TOTAL", EntryFontUnderline, PageWidth - RightMargin - ll, True, False, True)
            PrintArray.Add(field)
            Dim t = 0
            Dim GrandValue = 0
            Dim GrandVat = 0
            Dim GrandNett = 0
            Dim addMonthLine As Boolean
            Dim MonthValue = 0
            Dim MonthVat = 0
            Dim MonthNett = 0
            Dim curMonth, nextMonth, curYear, NextYear
            Dim accum = 0
            For t = 0 To DataSet.tables(0).rows.count - 1

                addMonthLine = False

                field = New PageElement(DataSet.tables(0).rows(t).item("InvoiceNo").ToString(), EntryFont, LeftMargin, False, False, False)
                PrintArray.Add(field)

                Dim msql = "Select * from InvoiceLine where InvNo = " + DataSet.tables(0).rows(t).item("InvoiceNo").ToString() + " order by [Line#]"
                Dim mDataSet = New Data.DataSet
                Dim madapter = New OleDb.OleDbDataAdapter(msql, DBConnection)
                madapter.Fill(mDataSet)


                field = New PageElement(DataSet.tables(0).rows(t).item("invDate").ToShortDateString(), EntryFont, LeftMargin + idt, False, False, False)
                PrintArray.Add(field)

                curMonth = (Date.Parse(DataSet.tables(0).rows(t).item("invDate").ToShortDateString())).Month
                curYear = (Date.Parse(DataSet.tables(0).rows(t).item("invDate").ToShortDateString())).Year

                If (t + 1 <= DataSet.tables(0).rows.count - 1) Then 'IF there is another invoice
                    nextMonth = Date.Parse(DataSet.tables(0).rows(t + 1).item("invDate").ToShortDateString()).Month
                    NextYear = Date.Parse(DataSet.tables(0).rows(t + 1).item("invDate").ToShortDateString()).Year
                    If (nextMonth <> curMonth Or NextYear <> curYear) Then
                        addMonthLine = True
                    End If
                End If

                If t = DataSet.tables(0).rows.count - 1 Then
                    addMonthLine = True
                End If



                field = New PageElement(toRand(DataSet.tables(0).rows(t).item("invTotal").ToString(), False), EntryFont, PageWidth - RightMargin - iv, False, False, True)
                PrintArray.Add(field)
                MonthValue += Decimal.Parse(DataSet.tables(0).rows(t).item("invTotal").ToString())
                field = New PageElement(toRand(DataSet.tables(0).rows(t).item("invVatAmt").ToString(), False), EntryFont, PageWidth - RightMargin - vt, False, False, True)
                PrintArray.Add(field)
                MonthVat += Decimal.Parse(DataSet.tables(0).rows(t).item("invVatAmt").ToString())
                field = New PageElement(toRand(DataSet.tables(0).rows(t).item("invNett").ToString(), False), EntryFont, PageWidth - RightMargin - ta, False, False, True)
                PrintArray.Add(field)
                accum += DataSet.tables(0).rows(t).item("invNett")
                field = New PageElement(toRand(accum, False), EntryFont, PageWidth - RightMargin - ll, True, False, True)
                PrintArray.Add(field)

                MonthNett += Decimal.Parse(DataSet.tables(0).rows(t).item("invNett").ToString())
                If (addMonthLine) Then
                    field = New PageElement(MonthName(curMonth) & " Total:", EntryFont, 170, False, False, False)
                    PrintArray.Add(field)


                    field = New PageElement(toRand(MonthValue, False), EntryFontUnderline, PageWidth - RightMargin - iv, False, False, True)
                    PrintArray.Add(field)
                    field = New PageElement(toRand(MonthVat, False), EntryFontUnderline, PageWidth - RightMargin - vt, False, False, True)
                    PrintArray.Add(field)
                    field = New PageElement(toRand(MonthNett, False), EntryFontUnderline, PageWidth - RightMargin - ll, True, False, True)
                    PrintArray.Add(field)
                    addMonthLine = False

                    GrandValue += MonthValue
                    GrandVat += MonthVat
                    GrandNett += MonthNett
                    MonthValue = 0
                    MonthVat = 0
                    MonthNett = 0
                End If
            Next t

            field = New PageElement("Summary Total:", EntryFont, 170, False, False, False)
            PrintArray.Add(field)
            field = New PageElement(toRand(GrandValue, True), EntryFont, PageWidth - RightMargin - iv, False, False, True)
            PrintArray.Add(field)
            field = New PageElement(toRand(GrandVat, True), EntryFont, PageWidth - RightMargin - vt, False, False, True)
            PrintArray.Add(field)
            field = New PageElement(toRand(GrandNett, True), EntryFont, PageWidth - RightMargin - ll, True, False, True)
            PrintArray.Add(field)

        Catch err As Exception
            MessageBox.Show(err.StackTrace & " - " & err.Message)
        End Try

    End Sub

    Private Sub btnPrintPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrintPreview.Click

        Me.Cursor = Windows.Forms.Cursors.WaitCursor

        If cmbJobs.Text = "" Then
            MessageBox.Show("Select a job number from the drop-down list.", "Invalid job number", MessageBoxButtons.OK)
            cmbJobs.Focus()
            Exit Sub
        End If

        Dim sql = "SELECT CompanyNo FROM Job WHERE JobNo = '" & cmbJobs.Text & "' ORDER BY JobNo"
        Dim DataSet = New Data.DataSet
        Dim adapter = New OleDb.OleDbDataAdapter(sql, DBConnection)
        adapter.Fill(DataSet)

        If DataSet.tables(0).rows.count = 0 Then
            MessageBox.Show("Selected job does not have an associated company number.", "Slight Error", MessageBoxButtons.OK)
            Exit Sub
        End If


        Try
            DocumentToPrint.DocumentName = "Summary - Job No: " + cmbJobs.Text
            Dim ppd_JCR As New PrintPreviewDialog
            ppd_JCR.WindowState = FormWindowState.Maximized
            ppd_JCR.Document = DocumentToPrint
            ppd_JCR.AutoScale = True
            ppd_JCR.AutoScroll = True
            ppd_JCR.UseAntiAlias = False
            ppd_JCR.PrintPreviewControl.Zoom = 1
            ppd_JCR.PrintPreviewControl.Columns = 1
            ppd_JCR.PrintPreviewControl.Rows = 1
            ppd_JCR.Text = "Summary - Job No: " + cmbJobs.Text
            curpagenum = 1
            PrintArray.Clear()

            If chkCS.Checked Then
                ReinforcingSummaryPrint(DataSet.tables(0).rows(0).Item("CompanyNo").ToString, cmbJobs.Text, dtpStart.Value.ToShortDateString, dtpEnd.Value.ToShortDateString)

            End If



            If chkM.Checked Then
                MeshSummaryPrint(DataSet.tables(0).rows(0).Item("CompanyNo").ToString, cmbJobs.Text, dtpStart.Value.ToShortDateString, dtpEnd.Value.ToShortDateString)
            End If

            If chkS.Checked Then
                SundriesSummaryPrint(DataSet.tables(0).rows(0).Item("CompanyNo").ToString, cmbJobs.Text, dtpStart.Value.ToShortDateString, dtpEnd.Value.ToShortDateString)
            End If


            curArrayPos = 0

            If Not (HasMesh Or HasSundry Or HasCut) Then
                MessageBox.Show("No invoices found matching the job number and dates selected.", "No records found", MessageBoxButtons.OK)
                HasMesh = False
                HasSundry = False
                HasCut = False
                Me.Cursor = Windows.Forms.Cursors.Default
                Exit Sub
            End If

            'If All_Is_OK Then
            ppd_JCR.ShowDialog()
            Me.Cursor = Windows.Forms.Cursors.Default
            ' Else
            '    Me.Cursor = Windows.Forms.Cursors.Default
            '    Exit Sub
            ' End If

        Catch er As Exception
            If er.Message = "No printers installed." Then
                MessageBox.Show("There is no printer installed. Please install a printer and try again.", "Printer not found", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                MessageBox.Show(er.Message, "ERROR - PLEASE FIX ME!!")
            End If

        End Try

        '  All_Is_OK = True
        Me.Cursor = Windows.Forms.Cursors.Default

    End Sub

    Private Sub PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles DocumentToPrint.PrintPage

        Me.Cursor = Windows.Forms.Cursors.Arrow
        Dim curY = TopMargin
        Dim MaxY = e.PageSettings.Bounds.Height - BottomMargin

        If ReportType = "Reinforcing Summary" Then
            e.Graphics.DrawString("Date Generated : " & Today().ToShortDateString, New Font("Arial", 8, FontStyle.Italic), Brushes.DimGray, LeftMargin, 1065)
            e.Graphics.DrawString("Page " & curpagenum, New Font("Arial", 8, FontStyle.Italic), Brushes.DimGray, 700, 1065)
        End If

        While (curY < MaxY) And (curArrayPos < PrintArray.Count)

            Select Case PrintArray(curArrayPos).Text
                Case "<PAGE BREAK>"
                    curY = MaxY + 1000
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

    Private Sub frmPrintReinforcingSummary_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        dtpStart.Text = #1/1/2005#
        populate_cmb_jobs()
    End Sub

    Private Sub populate_cmb_jobs()
        cmbJobs.Items.Clear()
        Dim sql = "SELECT JobNo FROM Job ORDER BY JobNo"
        Dim dataset As New Data.DataSet
        Dim adapter As New OleDb.OleDbDataAdapter(sql, DBConnection)
        adapter.Fill(dataset)

        Dim aunty
        For aunty = 0 To dataset.Tables(0).Rows.Count - 1
            cmbJobs.Items.Add(dataset.Tables(0).Rows(aunty).Item("JobNo").ToString())
        Next aunty


    End Sub

    Private Function MonthName(ByVal month As Int16) As String
        Select Case month
            Case 1
                Return "January"
            Case 2
                Return "February"
            Case 3
                Return "March"
            Case 4
                Return "April"
            Case 5
                Return "May"
            Case 6
                Return "June"
            Case 7
                Return "July"
            Case 8
                Return "August"
            Case 9
                Return "September"
            Case 10
                Return "October"
            Case 11
                Return "November"
            Case 12
                Return "December"
        End Select
    End Function

    Private Sub dtpStart_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpStart.ValueChanged

    End Sub

    Private Sub cmbJobs_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbJobs.SelectedIndexChanged

    End Sub
End Class
