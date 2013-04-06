Public Class PrintCutInv
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
    Friend WithEvents btn_PrintInv As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents DocumentToPrint As System.Drawing.Printing.PrintDocument
    Friend WithEvents txt_InvNumToPrint As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtHeading As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents lblType As System.Windows.Forms.Label
    Friend WithEvents dtpInvDate As System.Windows.Forms.DateTimePicker
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btn_PrintInv = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnClose = New System.Windows.Forms.Button
        Me.DocumentToPrint = New System.Drawing.Printing.PrintDocument
        Me.txt_InvNumToPrint = New System.Windows.Forms.ComboBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.txtHeading = New System.Windows.Forms.TextBox
        Me.dtpInvDate = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.lblType = New System.Windows.Forms.Label
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btn_PrintInv
        '
        Me.btn_PrintInv.Location = New System.Drawing.Point(112, 184)
        Me.btn_PrintInv.Name = "btn_PrintInv"
        Me.btn_PrintInv.Size = New System.Drawing.Size(112, 23)
        Me.btn_PrintInv.TabIndex = 0
        Me.btn_PrintInv.Text = "Print Preview..."
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(24, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Invoice Number:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(240, 184)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(104, 23)
        Me.btnClose.TabIndex = 0
        Me.btnClose.Text = "Close"
        '
        'DocumentToPrint
        '
        '
        'txt_InvNumToPrint
        '
        Me.txt_InvNumToPrint.Location = New System.Drawing.Point(128, 24)
        Me.txt_InvNumToPrint.Name = "txt_InvNumToPrint"
        Me.txt_InvNumToPrint.Size = New System.Drawing.Size(121, 21)
        Me.txt_InvNumToPrint.TabIndex = 3
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtHeading)
        Me.GroupBox1.Controls.Add(Me.dtpInvDate)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Location = New System.Drawing.Point(24, 88)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(320, 80)
        Me.GroupBox1.TabIndex = 4
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Invoice Header Details"
        '
        'txtHeading
        '
        Me.txtHeading.Location = New System.Drawing.Point(136, 48)
        Me.txtHeading.Name = "txtHeading"
        Me.txtHeading.Size = New System.Drawing.Size(168, 20)
        Me.txtHeading.TabIndex = 2
        Me.txtHeading.Text = ""
        '
        'dtpInvDate
        '
        Me.dtpInvDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpInvDate.Location = New System.Drawing.Point(136, 24)
        Me.dtpInvDate.Name = "dtpInvDate"
        Me.dtpInvDate.Size = New System.Drawing.Size(104, 20)
        Me.dtpInvDate.TabIndex = 1
        Me.dtpInvDate.Value = New Date(2005, 1, 31, 0, 0, 0, 0)
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 24)
        Me.Label2.Name = "Label2"
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Invoice Date:"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(16, 48)
        Me.Label3.Name = "Label3"
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "Invoice Heading"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(24, 56)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(72, 23)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "Invoice Type:"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblType
        '
        Me.lblType.Location = New System.Drawing.Point(128, 56)
        Me.lblType.Name = "lblType"
        Me.lblType.Size = New System.Drawing.Size(112, 23)
        Me.lblType.TabIndex = 2
        Me.lblType.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'PrintCutInv
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(370, 224)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.txt_InvNumToPrint)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btn_PrintInv)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.lblType)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "PrintCutInv"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Select Invoice To Print"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private CallingForm As Object
    Dim command As OleDb.OleDbCommand

    Public Sub New(ByVal caller As Object)
        MyBase.New()
        InitializeComponent()

        CallingForm = caller
    End Sub


    Private Sub frmPrintCutInv_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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
    Dim EntryFont As New Font("Courier", 11)
    Dim Head1Font As New Font("Courier", 30, FontStyle.Bold Or FontStyle.Underline)
    Dim Head2Font As New Font("Courier", 15, FontStyle.Bold)
    Dim Head2DetFont As New Font("Courier", 15, FontStyle.Italic)
    Dim EntryFontBold As New Font("Courier", 11, FontStyle.Bold)
    Dim EntryFontUnderline As New Font("Courier", 11, FontStyle.Underline)
    Dim DetailFont As New Font("Courier", 14)
    Dim TimeCardColFont As New Font("Courier", 10, FontStyle.Italic Or FontStyle.Bold)
    Dim ColFont As New Font("Courier", 12, FontStyle.Italic)
    Dim curArrayPos As Integer = 0
    Dim curpagenum As Integer = 1
    'Dim TopMargin = 90
    Dim TopMargin As Integer = 50
    Dim LeftMargin As Integer = 100
    Dim RightMargin As Integer = 40
    Dim BottomMargin As Integer = 60
    ' Dim PageWidth = 893
    Dim PageWidth As Integer = 873
    Dim ReportType As String
    Dim mes As PageElement
    Dim vatperc As String
    Dim All_Is_OK As Boolean = True
#End Region

    Private Sub updateInvoiceDate()
        DBConnection.Open()
        Dim invDate As Date
        invDate = dtpInvDate.Value.Date
        Dim SQL4UPDATE As String = "UPDATE Invoice SET InvDate = #" & invDate.ToLongDateString & "# WHERE InvoiceNo = " & txt_InvNumToPrint.Text
        command = New OleDb.OleDbCommand(SQL4UPDATE, DBConnection)
        Try
            command.ExecuteNonQuery()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            DBConnection.Close()
        End Try
    End Sub

    Public Sub btn_PrintInv_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_PrintInv.Click


        If txt_InvNumToPrint.Text = "" Then
            Exit Sub
        End If



        updateInvoiceDate()

        ReportType = "Invoice"

        Try

            DocumentToPrint.DocumentName = "Invoice No: " + txt_InvNumToPrint.Text

            Dim ppd_JCR As New PrintPreviewDialog
            ppd_JCR.WindowState = FormWindowState.Maximized
            ppd_JCR.Document = DocumentToPrint
            ppd_JCR.AutoScale = True
            ppd_JCR.AutoScroll = True
            ppd_JCR.UseAntiAlias = False
            ppd_JCR.PrintPreviewControl.Zoom = 1
            ppd_JCR.PrintPreviewControl.Columns = 1
            ppd_JCR.PrintPreviewControl.Rows = 1
            ppd_JCR.Text = "Invoice No: " + txt_InvNumToPrint.Text
            curpagenum = 1
            PrintArray.Clear()
            All_Is_OK = True
            InvoicePrint(txt_InvNumToPrint.Text, ty.Item(txt_InvNumToPrint.SelectedIndex), txtHeading.Text, dtpInvDate.Value)
            curArrayPos = 0
            If All_Is_OK Then
                ppd_JCR.ShowDialog()

            End If

        Catch er As Exception
            If er.Message = "No printers installed." Then
                MessageBox.Show("There is no printer installed. Please install a printer and try again.", "Printer not found", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                MessageBox.Show(er.Message, "ERROR - PLEASE FIX ME!!")
            End If
        End Try
        populate_invoiceNumbers()
    End Sub

    Public Shared Function toRand(ByVal input As String, ByVal r As Boolean) As String

        Dim iput As Double

        Try
            iput = Double.Parse(input)
            If r Then
                Return Format(iput, "R #,###,###,##0.00")
            Else
                Return Format(iput, "#,###,###,##0.00")
            End If

        Catch ex As Exception
            Return String.Empty
            MessageBox.Show("Error with input string.", "Cannot convert to Rand format.", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
        End Try

    End Function

    Private Sub AddCompNameRegNoVatNo(ByVal compNum As String, ByVal includeTelAndAddress As Boolean)
        'Get Company Details

        Dim sql As String = "SELECT * FROM Company WHERE CompanyNo = '" + compNum + "'"
        Dim dataset As New Data.DataSet
        Dim adapter As New OleDb.OleDbDataAdapter(sql, DBConnection)
        adapter.Fill(dataset)



        field = New PageElement(dataset.Tables(0).Rows(0).Item("CompanyName").ToString(), EntryFont, 0, True, True)
        PrintArray.Add(field)
        field = New PageElement("REG NO. " + dataset.Tables(0).Rows(0).Item("RegNo").ToString(), EntryFont, 0, True, True)
        PrintArray.Add(field)
        field = New PageElement("VAT NO. " + dataset.Tables(0).Rows(0).Item("VatNo").ToString(), EntryFont, 0, True, True)
        PrintArray.Add(field)
        If includeTelAndAddress Then
            field = New PageElement("TEL:", EntryFont, 540, False, False)
            PrintArray.Add(field)
            field = New PageElement(dataset.Tables(0).Rows(0).Item("Telephone").ToString(), EntryFont, 580, True, False)
            PrintArray.Add(field)
            field = New PageElement("FAX:", EntryFont, 540, False, False)
            PrintArray.Add(field)
            field = New PageElement(dataset.Tables(0).Rows(0).Item("Fax").ToString(), EntryFont, 580, True, False)
            PrintArray.Add(field)
            field = New PageElement(dataset.Tables(0).Rows(0).Item("Address").ToString(), EntryFont, 540, True, False)
            PrintArray.Add(field)
            field = New PageElement(dataset.Tables(0).Rows(0).Item("AddressLine2").ToString(), EntryFont, 540, True, False)

            If field.Text.Trim <> "" Then
                PrintArray.Add(field)
            End If

            field = New PageElement(dataset.Tables(0).Rows(0).Item("AddressLine3").ToString(), EntryFont, 540, True, False)
            If field.Text.Trim <> "" Then
                PrintArray.Add(field)
            End If
            field = New PageElement(dataset.Tables(0).Rows(0).Item("AddressLine4").ToString(), EntryFont, 540, True, False)
            If field.Text.Trim <> "" Then
                PrintArray.Add(field)
            End If
            field = New PageElement(dataset.Tables(0).Rows(0).Item("PostalCode").ToString(), EntryFont, 540, True, False)
            PrintArray.Add(field)
        End If

        mes = New PageElement(dataset.Tables(0).Rows(0).Item("Message").ToString(), EntryFont, LeftMargin, True, False)

        vatperc = dataset.Tables(0).Rows(0).Item("VatPerc").ToString()
        vatperc = (Decimal.Round(Decimal.Parse(vatperc) * 100, 0)).ToString + "%"

    End Sub

    Private Sub InvoicePrint(ByVal invNum As String, ByVal InvType As String, ByVal InvoiceHeading As String, ByVal InvDate As Date)
        Const con1 As Integer = 520
        'Const con1 = 485
        Dim sql As String = String.Empty
        If InvType = "Cutting Sheet" Then
            sql = "Select * from ((((invoice inner join job on Job.JobNo = Invoice.InvJobNo) inner join Contractor on Contractor.ContractorNo = job.ContractorNo) inner join CuttingSheet on cuttingSheet.invoiceNo = invoice.invoiceNo) inner join company on job.companyNo = Company.companyNo) where Invoice.InvoiceNo = " + invNum
        ElseIf InvType = "Mesh" Or InvType = "Sundry" Or InvType = "Escalation" Then
            sql = "Select * from (((invoice inner join job on Job.JobNo = Invoice.InvJobNo) inner join Contractor on Contractor.ContractorNo = job.ContractorNo)inner join company on job.companyNo = Company.companyNo) where Invoice.InvoiceNo = " + invNum
        End If

        Dim DataSet As DataSet = New Data.DataSet
        Dim adapter As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(sql, DBConnection)
        adapter.Fill(DataSet)

        If DataSet.tables(0).rows.count.ToString = "0" Then
            MessageBox.Show("No invoice exists with Invoice No. " + invNum)
            All_Is_OK = False
        Else
            AddCompNameRegNoVatNo(DataSet.tables(0).rows(0).item("Company.CompanyNo").ToString(), True)

            field = New PageElement(DataSet.Tables(0).Rows(0).Item("ContractorName").ToString(), EntryFont, LeftMargin, True, False)
            PrintArray.Add(field)
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("AddressLine1").ToString(), EntryFont, LeftMargin, True, False)
            If field.Text.Trim <> "" Then
                PrintArray.Add(field)
            End If
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("Contractor.AddressLine2").ToString(), EntryFont, LeftMargin, True, False)
            If field.Text.Trim <> "" Then
                PrintArray.Add(field)
            End If
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("Contractor.AddressLine3").ToString(), EntryFont, LeftMargin, True, False)
            If field.Text.Trim <> "" Then
                PrintArray.Add(field)
            End If
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("Contractor.AddressLine4").ToString(), EntryFont, LeftMargin, True, False)
            If field.Text.Trim <> "" Then
                PrintArray.Add(field)
            End If
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("Contractor.PostalCode").ToString(), EntryFont, LeftMargin, False, False)
            PrintArray.Add(field)
            field = New PageElement("TAX INVOICE", EntryFont, 540, True, False)
            PrintArray.Add(field)
            field = New PageElement("VAT NO. ", EntryFont, LeftMargin, False, False)
            PrintArray.Add(field)
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("Vat No").ToString(), EntryFont, LeftMargin + 65, True, False)
            PrintArray.Add(field)

            field = New PageElement("INVOICE NO:", EntryFont, 540, False, False)
            PrintArray.Add(field)


            If InvType = "Cutting Sheet" Then
                field = New PageElement(DataSet.Tables(0).Rows(0).Item("Invoice.InvoiceNo").ToString(), EntryFont, 640, True, False)
            ElseIf InvType = "Mesh" Or InvType = "Sundry" Or InvType = "Escalation" Then
                field = New PageElement(DataSet.Tables(0).Rows(0).Item("InvoiceNo").ToString(), EntryFont, 640, True, False)
            End If



            PrintArray.Add(field)
            field = New PageElement("DATE:", EntryFont, 540, False, False)
            PrintArray.Add(field)
            'field = New PageElement(DataSet.Tables(0).Rows(0).Item("InvDate").ToShortDateString(), EntryFont, 630, True, False)
            field = New PageElement(InvDate.ToShortDateString, EntryFont, 640, True, False)
            PrintArray.Add(field)
            Dim tot As PageElement = New PageElement(toRand(DataSet.Tables(0).Rows(0).Item("InvTotal").ToString(), True), EntryFont, PageWidth - RightMargin - 90, True, False, True)
            Dim vat As PageElement = New PageElement(toRand(DataSet.Tables(0).Rows(0).Item("InvVatAmt").ToString(), True), EntryFont, PageWidth - RightMargin - 90, True, False, True)
            Dim net As PageElement = New PageElement(toRand(DataSet.Tables(0).Rows(0).Item("InvNett").ToString(), True), EntryFont, PageWidth - RightMargin - 90, True, False, True)

            If InvType <> "Escalation" Then
                field = New PageElement("DELIVERY NOTE NO:", EntryFont, LeftMargin, False, False)
                PrintArray.Add(field)
                field = New PageElement(DataSet.Tables(0).Rows(0).Item("InvDeliveryNoteNo").ToString(), EntryFont, 270, True, False)
                PrintArray.Add(field)
            End If


            field = New PageElement("ORDER NO:", EntryFont, LeftMargin, False, False)
            PrintArray.Add(field)
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("InvOrdNum").ToString(), EntryFont, 270, True, False)
            PrintArray.Add(field)
            field = New PageElement("JOB NO:", EntryFont, LeftMargin, False, False)
            PrintArray.Add(field)
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("JobNo").ToString(), EntryFont, 270, True, False)
            PrintArray.Add(field)
            field = New PageElement("CONTRACT:", EntryFont, LeftMargin, False, False)
            PrintArray.Add(field)
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("JobName").ToString(), EntryFont, 270, True, False)
            PrintArray.Add(field)

            If DataSet.tables(0).rows(0).item("InvoiceType").ToString() = "Cutting Sheet" Then

                field = New PageElement("CUTSHEET NO:", EntryFont, LeftMargin, False, False)
                PrintArray.Add(field)
                field = New PageElement(DataSet.Tables(0).Rows(0).Item("CutSheetNo").ToString(), EntryFont, 270, True, False)
                PrintArray.Add(field)
                field = New PageElement("DETAILS:", EntryFont, LeftMargin, False, False)
                PrintArray.Add(field)
                field = New PageElement(DataSet.Tables(0).Rows(0).Item("Details").ToString(), EntryFont, 270, True, False)
                PrintArray.Add(field)

                Dim sql4sched As String = "SELECT * FROM SchedItem WHERE CutSheetNo = " & DataSet.Tables(0).Rows(0).Item("CutSheetNo").ToString()
                Dim ds4sched As New Data.DataSet
                Dim ad4sched As New OleDb.OleDbDataAdapter(sql4sched, DBConnection)
                ad4sched.Fill(ds4sched)
                Dim firstSched As String = String.Empty
                Dim lastSched As String = String.Empty
                If ds4sched.Tables(0).Rows.Count <> 0 Then
                    firstSched = ds4sched.Tables(0).Rows(0).Item("ScheduleNo").ToString
                    lastSched = ds4sched.Tables(0).Rows(ds4sched.Tables(0).Rows.Count - 1).Item("ScheduleNo").ToString
                End If

                field = New PageElement("SCHEDULE FROM:", EntryFont, LeftMargin, False, False)
                PrintArray.Add(field)
                field = New PageElement(firstSched, EntryFont, 270, True, False)
                PrintArray.Add(field)
                field = New PageElement("SCHEDULE TO:", EntryFont, LeftMargin, False, False)
                PrintArray.Add(field)
                field = New PageElement(lastSched, EntryFont, 270, True, False)
                PrintArray.Add(field)
                field = New PageElement("<HR/BLACK>", EntryFont, LeftMargin, -15, False)
                PrintArray.Add(field)
                'field = New PageElement(DataSet.Tables(0).Rows(0).Item("InvoiceHeading").ToString(), EntryFont, 0, True, True, False)
                field = New PageElement(InvoiceHeading, EntryFont, 0, True, True, False)
                PrintArray.Add(field)

                'Get Invoice Line Details
                Dim sqlMild, sqlHigh As String
                sqlMild = "Select * from (InvoiceLine INNER JOIN ProductType ON InvoiceLine.TypeCode = ProductType.TypeCode) INNER JOIN ProdCat ON Prodcat.CatCode = ProductType.CatCode where prodCat.CatCode = 'R' and InvNo = " + invNum + " order by [Line#]"
                Dim DSM As Data.DataSet = New Data.DataSet
                Dim adM As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(sqlMild, DBConnection)
                adM.Fill(DSM)
                sqlHigh = "Select * from (InvoiceLine INNER JOIN ProductType ON InvoiceLine.TypeCode = ProductType.TypeCode) INNER JOIN ProdCat ON Prodcat.CatCode = ProductType.CatCode where prodCat.CatCode = 'Y' and InvNo = " + invNum + " order by [Line#]"
                Dim DSH As Data.DataSet = New Data.DataSet
                Dim adH As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(sqlHigh, DBConnection)
                adH.Fill(DSH)


                Dim TKg As String = ""
                Const ra As Integer = 200
                Dim recordCountMild As Integer = DSM.Tables(0).Rows.Count
                Dim recordCountHigh As Integer = DSH.Tables(0).Rows.Count
                If recordCountMild <> 0 Then
                    field = New PageElement(DSM.tables(0).rows(0).item("CatDesc"), EntryFontUnderline, LeftMargin, True, False)
                    PrintArray.Add(field)
                End If
                Dim x As Integer
                Dim WeightTotal As Double = 0
                For x = 0 To recordCountMild - 1

                    field = New PageElement(DSM.Tables(0).Rows(x).Item("InvoiceLine.TypeCode"), EntryFont, LeftMargin + 20, False, False)
                    PrintArray.Add(field)
                    TKg = DSM.tables(0).rows(x).item("TonsorKg").ToString()
                    Dim Quantity As Double

                    Quantity = DSM.Tables(0).Rows(x).Item("Qty")
                    'Round off to 3 digits after decimal if T and 1 if Kg
                    If TKg = "T" Or TKg = "t" Then
                        Quantity = Math.Round(Quantity, 3)
                        field = New PageElement(Quantity.ToString("0.000"), EntryFont, LeftMargin + ra, False, False, True)
                        PrintArray.Add(field)
                    Else
                        TKg = TKg.ToLower
                        Quantity = Math.Round(Quantity, 1)
                        field = New PageElement(Quantity.ToString("0.0"), EntryFont, LeftMargin + ra, False, False, True)
                        PrintArray.Add(field)
                    End If
                    WeightTotal += Quantity

                    field = New PageElement(TKg, EntryFont, LeftMargin + 200, False, False)
                    PrintArray.Add(field)
                    field = New PageElement("@", EntryFont, LeftMargin + 240, False, False)
                    PrintArray.Add(field)
                    field = New PageElement(toRand(DSM.Tables(0).Rows(x).Item("CostPerUnit"), True), EntryFont, LeftMargin + 370, False, False, True)
                    PrintArray.Add(field)
                    field = New PageElement(toRand(DSM.Tables(0).Rows(x).Item("Total"), True), EntryFont, PageWidth - RightMargin - 90, True, False, True)
                    PrintArray.Add(field)
                Next
                If recordCountHigh <> 0 Then
                    field = New PageElement(DSH.tables(0).rows(0).item("CatDesc"), EntryFontUnderline, LeftMargin, True, False)
                    PrintArray.Add(field)
                End If
                For x = 0 To recordCountHigh - 1

                    field = New PageElement(DSH.Tables(0).Rows(x).Item("InvoiceLine.TypeCode"), EntryFont, LeftMargin + 20, False, False)
                    PrintArray.Add(field)
                    TKg = DSH.tables(0).rows(x).item("TonsorKg").ToString()
                    Dim Quantity As Double
                    Quantity = DSH.Tables(0).Rows(x).Item("Qty")
                    If TKg = "T" Then
                        Quantity = Math.Round(Quantity, 3)
                        field = New PageElement(Quantity.ToString("0.000"), EntryFont, LeftMargin + ra, False, False, True)
                        PrintArray.Add(field)
                    Else
                        TKg = TKg.ToLower
                        Quantity = Math.Round(Quantity, 1)
                        field = New PageElement(Quantity.ToString("0.0"), EntryFont, LeftMargin + ra, False, False, True)
                        PrintArray.Add(field)
                    End If
                    WeightTotal += Quantity

                    field = New PageElement(TKg, EntryFont, LeftMargin + 200, False, False)
                    PrintArray.Add(field)
                    field = New PageElement("@", EntryFont, LeftMargin + 240, False, False)
                    PrintArray.Add(field)
                    field = New PageElement(toRand(DSH.Tables(0).Rows(x).Item("CostPerUnit"), True), EntryFont, LeftMargin + 370, False, False, True)
                    PrintArray.Add(field)
                    field = New PageElement(toRand(DSH.Tables(0).Rows(x).Item("Total"), True), EntryFont, PageWidth - RightMargin - 90, True, False, True)
                    PrintArray.Add(field)
                Next

                field = New PageElement(LeftMargin + 150, 320, False)
                PrintArray.Add(field)
                field = New PageElement(LeftMargin + con1, PageWidth - RightMargin - 90, True)
                PrintArray.Add(field)
                field = New PageElement("TOTAL MASS:", EntryFont, LeftMargin + 15, False, False)
                PrintArray.Add(field)
                If TKg.ToUpper = "T" Then
                    WeightTotal = Math.Round(WeightTotal, 3)
                    field = New PageElement(WeightTotal.ToString("0.000"), EntryFont, LeftMargin + ra, False, False, True)
                    PrintArray.Add(field)
                ElseIf TKg.ToLower = "kg" Then
                    WeightTotal = Math.Round(WeightTotal, 1)
                    field = New PageElement(WeightTotal.ToString("0.0"), EntryFont, LeftMargin + ra, False, False, True)
                    PrintArray.Add(field)
                End If

                field = New PageElement(TKg, EntryFont, LeftMargin + 200, False, False)
                PrintArray.Add(field)
                field = New PageElement("TOTAL:", EntryFont, LeftMargin + 380, False, False)
                PrintArray.Add(field)
                PrintArray.Add(tot)
                field = New PageElement("ADD VAT " + vatperc + " :", EntryFont, LeftMargin + 380, False, False)
                PrintArray.Add(field)
                PrintArray.Add(vat)
                field = New PageElement(LeftMargin + con1, PageWidth - RightMargin - 90, True)
                PrintArray.Add(field)
                field = New PageElement("NETT:", EntryFont, LeftMargin + 380, False, False)
                PrintArray.Add(field)
                PrintArray.Add(net)
                field = New PageElement(True, LeftMargin + con1, PageWidth - RightMargin - 90)
                PrintArray.Add(field)
                '======================================   MESH SPECIFIC    ================================
            ElseIf DataSet.tables(0).rows(0).item("InvoiceType").ToString() = "Mesh" Then

                PrintArray.Add(New PageElement("DETAILS:", EntryFont, LeftMargin, False, False))
                PrintArray.Add(New PageElement(DataSet.Tables(0).Rows(0).Item("invComments").ToString(), EntryFont, 270, True, False))


                field = New PageElement("<HR/BLACK>", EntryFont, LeftMargin, -15, False)
                PrintArray.Add(field)
                'field = New PageElement(DataSet.Tables(0).Rows(0).Item("InvoiceHeading").ToString(), EntryFont, LeftMargin, True, False, False)
                field = New PageElement(InvoiceHeading, EntryFont, LeftMargin, True, False, False)
                PrintArray.Add(field)
                PrintArray.Add(New PageElement("", EntryFont, 0, True, False, False))

                sql = "SELECT * FROM InvoiceLine WHERE InvNo = " & invNum
                Dim d1 As Data.DataSet = New Data.DataSet
                Dim a1 As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(sql, DBConnection)
                a1.Fill(d1)

                Dim fd As Integer
                For fd = 0 To d1.Tables(0).Rows.Count - 1
                    PrintArray.Add(New PageElement(d1.Tables(0).Rows(fd).Item("Description").ToString() & toRand(d1.Tables(0).Rows(fd).Item("CostPerUnit").ToString(), True), EntryFont, LeftMargin, False, False, False))
                    field = New PageElement(toRand(d1.Tables(0).Rows(fd).Item("Total"), True), EntryFont, PageWidth - RightMargin - 90, True, False, True)
                    PrintArray.Add(field)
                    field = New PageElement("ADD VAT " + vatperc + " :", EntryFont, LeftMargin + 380, False, False)
                    PrintArray.Add(field)
                    PrintArray.Add(New PageElement(toRand(DataSet.Tables(0).Rows(0).Item("InvVatAmt").ToString(), True), EntryFont, PageWidth - RightMargin - 90, True, False, True))
                    field = New PageElement(LeftMargin + con1, PageWidth - RightMargin - 90, True)
                    PrintArray.Add(field)
                    field = New PageElement("NETT:", EntryFont, LeftMargin + 380, False, False)
                    PrintArray.Add(field)
                    PrintArray.Add(New PageElement(toRand(DataSet.Tables(0).Rows(0).Item("InvNett").ToString(), True), EntryFont, PageWidth - RightMargin - 90, True, False, True))
                    field = New PageElement(True, LeftMargin + con1, PageWidth - RightMargin - 90)
                    PrintArray.Add(field)
                Next fd
                '======================================   SUNDRY SPECIFIC    ================================
            ElseIf DataSet.tables(0).rows(0).item("InvoiceType").ToString() = "Sundry" Then


                PrintArray.Add(New PageElement("DETAILS:", EntryFont, LeftMargin, False, False))
                PrintArray.Add(New PageElement(DataSet.Tables(0).Rows(0).Item("invRefNum").ToString(), EntryFont, 270, True, False))
                PrintArray.Add(New PageElement(DataSet.Tables(0).Rows(0).Item("invComments").ToString(), EntryFont, 270, True, False))

                field = New PageElement("<HR/BLACK>", EntryFont, LeftMargin, -15, False)
                PrintArray.Add(field)
                field = New PageElement(InvoiceHeading, EntryFont, LeftMargin, True, True)
                PrintArray.Add(field)
                PrintArray.Add(New PageElement("", EntryFont, 0, True, False, False))

                sql = "SELECT * FROM InvoiceLine WHERE InvNo = " & invNum
                Dim d1 As Data.DataSet = New Data.DataSet
                Dim a1 As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(sql, DBConnection)
                a1.Fill(d1)

                Dim fd As Integer
                For fd = 0 To d1.Tables(0).Rows.Count - 1
                    PrintArray.Add(New PageElement(d1.Tables(0).Rows(fd).Item("Description").ToString(), EntryFont, LeftMargin, False, False, False))
                    If Not (d1.Tables(0).Rows(fd).Item("Qty").ToString() = "0") Then
                        PrintArray.Add(New PageElement(d1.Tables(0).Rows(fd).Item("Qty").ToString() & "   x  ", EntryFont, 425, False, False, True))
                        PrintArray.Add(New PageElement(toRand(d1.Tables(0).Rows(fd).Item("CostPerUnit").ToString(), True), EntryFont, 525, False, False, True))

                        field = New PageElement(toRand(d1.Tables(0).Rows(fd).Item("Total"), True), EntryFont, PageWidth - RightMargin - 90, True, False, True)
                        PrintArray.Add(field)
                    Else
                        PrintArray.Add(New PageElement("", EntryFont, 0, True, False, False))
                    End If
                Next fd

                field = New PageElement(LeftMargin + con1, PageWidth - RightMargin - 90, True)
                PrintArray.Add(field)
                field = New PageElement("TOTAL:", EntryFont, LeftMargin + 380, False, False)
                PrintArray.Add(field)
                PrintArray.Add(New PageElement(toRand(DataSet.Tables(0).Rows(0).Item("InvTotal").ToString(), True), EntryFont, PageWidth - RightMargin - 90, True, False, True))
                field = New PageElement("ADD VAT " + vatperc + " :", EntryFont, LeftMargin + 380, False, False)
                PrintArray.Add(field)
                PrintArray.Add(New PageElement(toRand(DataSet.Tables(0).Rows(0).Item("InvVatAmt").ToString(), True), EntryFont, PageWidth - RightMargin - 90, True, False, True))
                field = New PageElement(LeftMargin + con1, PageWidth - RightMargin - 90, True)
                PrintArray.Add(field)
                field = New PageElement("NETT:", EntryFont, LeftMargin + 380, False, False)
                PrintArray.Add(field)
                PrintArray.Add(New PageElement(toRand(DataSet.Tables(0).Rows(0).Item("InvNett").ToString(), True), EntryFont, PageWidth - RightMargin - 90, True, False, True))
                field = New PageElement(True, LeftMargin + con1, PageWidth - RightMargin - 90)
                PrintArray.Add(field)

                '======================================   ESCALATION SPECIFIC    ================================
            ElseIf DataSet.tables(0).rows(0).item("InvoiceType").ToString() = "Escalation" Then
                field = New PageElement("<HR/BLACK>", EntryFont, LeftMargin, -15, False)
                PrintArray.Add(field)
                PrintArray.Add(New PageElement("", EntryFont, 0, True, False, False))

                sql = "SELECT * FROM InvoiceLine WHERE InvNo = " & invNum
                Dim d1 As Data.DataSet = New Data.DataSet
                Dim a1 As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(sql, DBConnection)
                a1.Fill(d1)

                PrintArray.Add(New PageElement("DETAILS :", EntryFont, LeftMargin, False, False, False))


                Const c As Integer = 70, c1 As Integer = 170

                PrintArray.Add(New PageElement(d1.Tables(0).Rows(0).Item("Description").ToString(), EntryFont, LeftMargin + c, True, False, False))
                PrintArray.Add(New PageElement(d1.Tables(0).Rows(1).Item("Description").ToString(), EntryFont, LeftMargin + c, True, False, False))
                PrintArray.Add(New PageElement("", EntryFont, 0, True, False, False))
                PrintArray.Add(New PageElement("MONTH", EntryFont, LeftMargin + c, False, False, False))
                PrintArray.Add(New PageElement("-     " & d1.Tables(0).Rows(2).Item("Description").ToString(), EntryFont, LeftMargin + c1, True, False, False))
                PrintArray.Add(New PageElement("FACTOR", EntryFont, LeftMargin + c, False, False, False))
                PrintArray.Add(New PageElement("-     " & d1.Tables(0).Rows(3).Item("Description").ToString(), EntryFont, LeftMargin + c1, True, False, False))
                PrintArray.Add(New PageElement("", EntryFont, 0, True, False, False))
                PrintArray.Add(New PageElement("WORK", EntryFont, LeftMargin + c, False, False, False))
                PrintArray.Add(New PageElement("-     " & d1.Tables(0).Rows(4).Item("Description").ToString(), EntryFont, LeftMargin + c1, False, False, False))
                field = New PageElement("TOTAL:", EntryFont, LeftMargin + 380, False, False)
                PrintArray.Add(field)
                PrintArray.Add(New PageElement(toRand(DataSet.Tables(0).Rows(0).Item("InvTotal").ToString(), True), EntryFont, PageWidth - RightMargin - 90, True, False, True))

                PrintArray.Add(New PageElement("VAT", EntryFont, LeftMargin + c, False, False, False))
                PrintArray.Add(New PageElement("-     " & d1.Tables(0).Rows(5).Item("Description").ToString(), EntryFont, LeftMargin + c1, False, False, False))
                field = New PageElement("ADD VAT " + vatperc + " :", EntryFont, LeftMargin + 380, False, False)
                PrintArray.Add(field)
                PrintArray.Add(New PageElement(toRand(DataSet.Tables(0).Rows(0).Item("InvVatAmt").ToString(), True), EntryFont, PageWidth - RightMargin - 90, True, False, True))



                field = New PageElement(LeftMargin + con1, PageWidth - RightMargin - 90, True)
                PrintArray.Add(field)
                field = New PageElement("NETT:", EntryFont, LeftMargin + 380, False, False)
                PrintArray.Add(field)
                PrintArray.Add(New PageElement(toRand(DataSet.Tables(0).Rows(0).Item("InvNett").ToString(), True), EntryFont, PageWidth - RightMargin - 90, True, False, True))
                field = New PageElement(True, LeftMargin + con1, PageWidth - RightMargin - 90)
                PrintArray.Add(field)

            End If
            '========================================================================================
            Dim cols As Integer = 40

            While mes.Text <> ""
                If mes.Text.Length >= cols Then
                    If mes.Text.IndexOf(" ") = cols Or mes.Text.IndexOf(" ") = cols + 1 Then
                        PrintArray.Add(New PageElement(mes.Text.Substring(0, cols).Trim(), EntryFont, LeftMargin, True, False))
                        mes.Text = mes.Text.Remove(0, cols)
                    Else
                        PrintArray.Add(New PageElement(mes.Text.Substring(0, mes.Text.LastIndexOf(" ", cols, cols)).Trim, EntryFont, LeftMargin, True, False))
                        mes.Text = mes.Text.Remove(0, mes.Text.LastIndexOf(" ", cols, cols))
                    End If

                Else
                    PrintArray.Add(New PageElement(mes.Text.Substring(0, mes.Text.Length).Trim, EntryFont, LeftMargin, True, False))
                    mes.Text = mes.Text.Remove(0, mes.Text.Length)
                End If

            End While
        End If



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

    Private Sub PrintCutInv_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        populate_invoiceNumbers()
    End Sub

    Dim ty, dt, hd As ArrayList

    Public Sub populate_invoiceNumbers()
        txt_InvNumToPrint.Items.Clear()
        Dim sql As String = "SELECT InvoiceNo,InvoiceType,InvDate,InvoiceHeading FROM Invoice ORDER BY InvoiceNo"
        Dim ds As New Data.DataSet
        Dim ad As New OleDb.OleDbDataAdapter(sql, DBConnection)
        ad.Fill(ds)

        ty = New ArrayList
        dt = New ArrayList
        hd = New ArrayList

        Dim f As Integer
        For f = 0 To ds.Tables(0).Rows.Count - 1
            txt_InvNumToPrint.Items.Add(ds.Tables(0).Rows(f).Item("InvoiceNo").ToString())
            ty.Add(ds.Tables(0).Rows(f).Item("InvoiceType").ToString())
            dt.Add(ds.Tables(0).Rows(f).Item("InvDate"))
            hd.Add(ds.Tables(0).Rows(f).Item("InvoiceHeading").ToString())
        Next f

    End Sub

    Private Sub txt_InvNumToPrint_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_InvNumToPrint.SelectedIndexChanged
        dtpInvDate.Value = dt(txt_InvNumToPrint.SelectedIndex)
        txtHeading.Text = hd(txt_InvNumToPrint.SelectedIndex)
        lblType.Text = ty(txt_InvNumToPrint.SelectedIndex)
    End Sub
End Class
