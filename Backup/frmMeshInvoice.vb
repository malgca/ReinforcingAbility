Public Class frmMeshInvoice
    Inherits System.Windows.Forms.Form

    Dim callingForm As Object
    Dim DbConnection As OleDb.OleDbConnection

#Region " Windows Form Designer generated code "

    Public Sub New(ByRef cf As Object, ByRef dbc As OleDb.OleDbConnection)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        callingForm = cf
        DbConnection = dbc
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
    Friend WithEvents btnCreateInvoice As System.Windows.Forms.Button
    Friend WithEvents dtpInvDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtDelNoteNum As System.Windows.Forms.TextBox
    Friend WithEvents txtOrderNum As System.Windows.Forms.TextBox
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents cmbJobNum As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtRefNo As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtAddDetails As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtLength As System.Windows.Forms.TextBox
    Friend WithEvents txtWidth As System.Windows.Forms.TextBox
    Friend WithEvents txtQty As System.Windows.Forms.TextBox
    Friend WithEvents txtRate As System.Windows.Forms.TextBox
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents lblNotify As System.Windows.Forms.Label
    Friend WithEvents cmbJobName As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtAdditional As System.Windows.Forms.TextBox
    Friend WithEvents txtJobName As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents btnNewInvoice As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cmbJobNum = New System.Windows.Forms.ComboBox
        Me.btnCreateInvoice = New System.Windows.Forms.Button
        Me.dtpInvDate = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtDelNoteNum = New System.Windows.Forms.TextBox
        Me.txtOrderNum = New System.Windows.Forms.TextBox
        Me.btnClose = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtRefNo = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtAddDetails = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtLength = New System.Windows.Forms.TextBox
        Me.txtWidth = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.txtQty = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.txtRate = New System.Windows.Forms.TextBox
        Me.btnPrint = New System.Windows.Forms.Button
        Me.lblNotify = New System.Windows.Forms.Label
        Me.cmbJobName = New System.Windows.Forms.ComboBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtAdditional = New System.Windows.Forms.TextBox
        Me.txtJobName = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.btnNewInvoice = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'cmbJobNum
        '
        Me.cmbJobNum.Location = New System.Drawing.Point(160, 16)
        Me.cmbJobNum.Name = "cmbJobNum"
        Me.cmbJobNum.Size = New System.Drawing.Size(80, 21)
        Me.cmbJobNum.TabIndex = 1
        '
        'btnCreateInvoice
        '
        Me.btnCreateInvoice.Location = New System.Drawing.Point(176, 408)
        Me.btnCreateInvoice.Name = "btnCreateInvoice"
        Me.btnCreateInvoice.Size = New System.Drawing.Size(112, 23)
        Me.btnCreateInvoice.TabIndex = 12
        Me.btnCreateInvoice.Text = "Create Invoice"
        '
        'dtpInvDate
        '
        Me.dtpInvDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpInvDate.Location = New System.Drawing.Point(160, 152)
        Me.dtpInvDate.Name = "dtpInvDate"
        Me.dtpInvDate.Size = New System.Drawing.Size(88, 20)
        Me.dtpInvDate.TabIndex = 4
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(24, 118)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(128, 23)
        Me.Label1.TabIndex = 21
        Me.Label1.Text = "Delivery Note Number"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(24, 84)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(128, 23)
        Me.Label2.TabIndex = 19
        Me.Label2.Text = "Order Number"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(24, 152)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(128, 23)
        Me.Label4.TabIndex = 15
        Me.Label4.Text = "Invoice Date"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(24, 16)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(128, 23)
        Me.Label7.TabIndex = 17
        Me.Label7.Text = "Job Number"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtDelNoteNum
        '
        Me.txtDelNoteNum.Location = New System.Drawing.Point(160, 112)
        Me.txtDelNoteNum.MaxLength = 50
        Me.txtDelNoteNum.Name = "txtDelNoteNum"
        Me.txtDelNoteNum.Size = New System.Drawing.Size(304, 20)
        Me.txtDelNoteNum.TabIndex = 3
        Me.txtDelNoteNum.Text = ""
        '
        'txtOrderNum
        '
        Me.txtOrderNum.Location = New System.Drawing.Point(160, 80)
        Me.txtOrderNum.Name = "txtOrderNum"
        Me.txtOrderNum.Size = New System.Drawing.Size(240, 20)
        Me.txtOrderNum.TabIndex = 2
        Me.txtOrderNum.Text = ""
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(464, 408)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(112, 23)
        Me.btnClose.TabIndex = 13
        Me.btnClose.Text = "Close"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(24, 200)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(128, 23)
        Me.Label3.TabIndex = 18
        Me.Label3.Text = "Reference Number"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtRefNo
        '
        Me.txtRefNo.Location = New System.Drawing.Point(160, 200)
        Me.txtRefNo.Name = "txtRefNo"
        Me.txtRefNo.Size = New System.Drawing.Size(248, 20)
        Me.txtRefNo.TabIndex = 5
        Me.txtRefNo.Text = ""
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(24, 232)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(128, 23)
        Me.Label6.TabIndex = 18
        Me.Label6.Text = "Description"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtAddDetails
        '
        Me.txtAddDetails.Location = New System.Drawing.Point(160, 232)
        Me.txtAddDetails.Name = "txtAddDetails"
        Me.txtAddDetails.Size = New System.Drawing.Size(328, 20)
        Me.txtAddDetails.TabIndex = 6
        Me.txtAddDetails.Text = ""
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(152, 320)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(48, 23)
        Me.Label8.TabIndex = 18
        Me.Label8.Text = "Length"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtLength
        '
        Me.txtLength.Location = New System.Drawing.Point(200, 320)
        Me.txtLength.Name = "txtLength"
        Me.txtLength.Size = New System.Drawing.Size(64, 20)
        Me.txtLength.TabIndex = 9
        Me.txtLength.Text = "6.0"
        '
        'txtWidth
        '
        Me.txtWidth.Location = New System.Drawing.Point(312, 320)
        Me.txtWidth.Name = "txtWidth"
        Me.txtWidth.Size = New System.Drawing.Size(64, 20)
        Me.txtWidth.TabIndex = 10
        Me.txtWidth.Text = "2.4"
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(272, 320)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(40, 23)
        Me.Label9.TabIndex = 18
        Me.Label9.Text = "Width"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtQty
        '
        Me.txtQty.Location = New System.Drawing.Point(80, 320)
        Me.txtQty.Name = "txtQty"
        Me.txtQty.Size = New System.Drawing.Size(64, 20)
        Me.txtQty.TabIndex = 8
        Me.txtQty.Text = ""
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(24, 320)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(56, 23)
        Me.Label10.TabIndex = 18
        Me.Label10.Text = "Quantity"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(384, 320)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(40, 23)
        Me.Label11.TabIndex = 18
        Me.Label11.Text = "Rate"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtRate
        '
        Me.txtRate.Location = New System.Drawing.Point(424, 320)
        Me.txtRate.Name = "txtRate"
        Me.txtRate.Size = New System.Drawing.Size(64, 20)
        Me.txtRate.TabIndex = 11
        Me.txtRate.Text = ""
        '
        'btnPrint
        '
        Me.btnPrint.Enabled = False
        Me.btnPrint.Location = New System.Drawing.Point(320, 408)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(112, 23)
        Me.btnPrint.TabIndex = 12
        Me.btnPrint.Text = "Print Preview..."
        '
        'lblNotify
        '
        Me.lblNotify.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNotify.Location = New System.Drawing.Point(32, 368)
        Me.lblNotify.Name = "lblNotify"
        Me.lblNotify.Size = New System.Drawing.Size(464, 23)
        Me.lblNotify.TabIndex = 22
        Me.lblNotify.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cmbJobName
        '
        Me.cmbJobName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbJobName.Location = New System.Drawing.Point(376, 16)
        Me.cmbJobName.Name = "cmbJobName"
        Me.cmbJobName.Size = New System.Drawing.Size(256, 21)
        Me.cmbJobName.TabIndex = 1
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(24, 280)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(128, 23)
        Me.Label5.TabIndex = 18
        Me.Label5.Text = "Additional Details"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtAdditional
        '
        Me.txtAdditional.Location = New System.Drawing.Point(160, 280)
        Me.txtAdditional.Name = "txtAdditional"
        Me.txtAdditional.Size = New System.Drawing.Size(328, 20)
        Me.txtAdditional.TabIndex = 7
        Me.txtAdditional.Text = ""
        '
        'txtJobName
        '
        Me.txtJobName.Enabled = False
        Me.txtJobName.Location = New System.Drawing.Point(160, 48)
        Me.txtJobName.Name = "txtJobName"
        Me.txtJobName.Size = New System.Drawing.Size(336, 20)
        Me.txtJobName.TabIndex = 23
        Me.txtJobName.Text = ""
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(24, 50)
        Me.Label12.Name = "Label12"
        Me.Label12.TabIndex = 24
        Me.Label12.Text = "Job Name"
        '
        'btnNewInvoice
        '
        Me.btnNewInvoice.Location = New System.Drawing.Point(32, 408)
        Me.btnNewInvoice.Name = "btnNewInvoice"
        Me.btnNewInvoice.Size = New System.Drawing.Size(112, 23)
        Me.btnNewInvoice.TabIndex = 12
        Me.btnNewInvoice.Text = "New Invoice"
        '
        'frmMeshInvoice
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(650, 448)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.txtJobName)
        Me.Controls.Add(Me.txtAdditional)
        Me.Controls.Add(Me.lblNotify)
        Me.Controls.Add(Me.cmbJobNum)
        Me.Controls.Add(Me.btnCreateInvoice)
        Me.Controls.Add(Me.dtpInvDate)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.txtDelNoteNum)
        Me.Controls.Add(Me.txtOrderNum)
        Me.Controls.Add(Me.txtRefNo)
        Me.Controls.Add(Me.txtAddDetails)
        Me.Controls.Add(Me.txtLength)
        Me.Controls.Add(Me.txtWidth)
        Me.Controls.Add(Me.txtQty)
        Me.Controls.Add(Me.txtRate)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.cmbJobName)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.btnNewInvoice)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "frmMeshInvoice"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Create a Mesh Invoice"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, ByVal keyData As Keys) As Boolean
        If keyData = Keys.Enter Then
            SendKeys.Send("{Tab}")
            Return True
        End If

        Return MyBase.ProcessCmdKey(msg, keyData)

    End Function

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Close()
        callingForm.show()
        callingForm = Nothing
    End Sub

    Private Sub frmMeshInvoice_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        populateCmbJobs()
    End Sub

    Private Sub populateCmbJobs()
        cmbJobName.Items.Clear()
        cmbJobNum.Items.Clear()
        Dim sql = "SELECT jobno,jobname FROM job ORDER BY jobno"
        Dim ds As Data.DataSet = New Data.DataSet
        Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(sql, DbConnection)
        da.Fill(ds)

        Dim i
        For i = 0 To ds.Tables(0).Rows.Count - 1
            cmbJobNum.Items.Add(ds.Tables(0).Rows(i).Item("JobNo").ToString())
            cmbJobName.Items.Add(ds.Tables(0).Rows(i).Item("JobName").ToString())
        Next i

    End Sub

    Private Sub cmbJobNum_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbJobNum.Leave
        If cmbJobNum.Text <> "" Then
            getJobOrderNum()
        End If

        lblNotify.Text = ""
        btnPrint.Enabled = False

    End Sub
    Private Sub cmbJobName_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbJobName.Leave
        cmbJobNum_Leave(sender, e)

    End Sub


    Private Sub getJobOrderNum()
        Dim sql = "SELECT jobName, orderNo FROM job WHERE jobno = '" & cmbJobNum.Text & "'"
        Dim ds As Data.DataSet = New Data.DataSet
        Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(sql, DbConnection)
        da.Fill(ds)

        Try
            txtOrderNum.Text = ds.Tables(0).Rows(0).Item("orderNo").ToString
            txtJobName.Text = ds.Tables(0).Rows(0).Item("jobName").ToString
        Catch ex As Exception
            MessageBox.Show("There are no jobs matching this job number.", "Invalid Job Number", MessageBoxButtons.OK, MessageBoxIcon.Information)
            cmbJobNum.Focus()
            cmbJobNum.SelectAll()
        End Try

    End Sub
    Dim InvoiceNumber As Long
    Private Sub btnCreateInvoice_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateInvoice.Click
        If cmbJobNum.Text.Trim = "" Then
            MessageBox.Show("Please select a valid job number.", "No Job Selected", MessageBoxButtons.OK, MessageBoxIcon.Information)
            cmbJobNum.Focus()
            Exit Sub
        End If

        If Not IsNumeric(txtQty.Text) Then
            MessageBox.Show("Invalid Value. Please enter a number.", "Invalid Number", MessageBoxButtons.OK)
            txtQty.Focus()
            txtQty.SelectAll()
            Exit Sub
        End If
        If Not IsNumeric(txtLength.Text) Then
            MessageBox.Show("Invalid Value. Please enter a number.", "Invalid Number", MessageBoxButtons.OK)
            txtLength.Focus()
            txtLength.SelectAll()
            Exit Sub
        End If
        If Not IsNumeric(txtWidth.Text) Then
            MessageBox.Show("Invalid Value. Please enter a number.", "Invalid Number", MessageBoxButtons.OK)
            txtWidth.Focus()
            txtWidth.SelectAll()
            Exit Sub
        End If
        If Not IsNumeric(txtRate.Text) Then
            MessageBox.Show("Invalid Value. Please enter a number.", "Invalid Number", MessageBoxButtons.OK)
            txtRate.Focus()
            txtRate.SelectAll()
            Exit Sub
        End If

        Dim sql = "SELECT * FROM (Job INNER JOIN Company ON Job.companyNo = Company.CompanyNo) INNER JOIN Contractor ON Job.ContractorNo = Contractor.ContractorNo WHERE Job.JobNo = '" & cmbJobNum.Text & "'"
        Dim ds As Data.DataSet = New Data.DataSet
        Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(sql, DbConnection)
        da.Fill(ds)

        If ds.Tables(0).Rows.Count = 0 Then
            MessageBox.Show("Cannot create invoice. Job does not exist.", "Invalid Job", MessageBoxButtons.OK)
            Exit Sub
        End If


        Try
            InvoiceNumber = Long.Parse(ds.Tables(0).Rows(0).Item("LastInvNum").ToString()) + 1
        Catch ex As Exception
            MessageBox.Show("Error with Invoice Number.")
        End Try

        Dim InvoiceType As String = "Mesh"
        Dim Factor As Int16 = 1
        Dim EscMoAndDa As Date = New Date(1999, 12, 11)
        Dim Work As Int16 = 75
        Dim RefNo As String = txtRefNo.Text
        Dim JobNo As String = ds.Tables(0).Rows(0).Item("JobNo").ToString()
        Dim VAT As String = ds.Tables(0).Rows(0).Item("VatPerc").ToString()
        Dim Active = "Yes"
        Dim Escalated = "No"
        Dim OnSummary = "Yes"
        Dim Comments As String = txtAdditional.Text
        Dim heading As String = "TO SUPPLY MESH REF " & txtRefNo.Text & " " & txtAddDetails.Text
        Dim invDate As Date

        invDate = dtpInvDate.Value.Date
        Dim sql4NewInvoice = "INSERT INTO Invoice(InvoiceNo,InvoiceType,InvDate,InvDeliveryNoteNo,InvFactor,Invmonthandyear,InvWork,InvOrdNum,InvRefNum,InvoiceHeading,InvTotal,InvVatAmt,InvDesign,InvNett,InvJobNo,InvActive,InvEscalated,InvOnSummary,InvComments) VALUES " & _
                "(  " & _
                    InvoiceNumber.ToString & _
                    ",'" & InvoiceType & _
                    "',#" & invDate.ToLongDateString & _
                    "#,'" & txtDelNoteNum.Text & _
                    "'," & Factor & _
                    ",#" & EscMoAndDa & _
                    "#," & Work & _
                    ",'" & txtOrderNum.Text & _
                    "','" & RefNo & _
                    "','" & heading & _
                    "'," & "-1" & _
                    "," & "-1" & _
                    "," & 0 & _
                    "," & "-1" & _
                    ",'" & JobNo & _
                    "'," & Active & _
                    "," & Escalated & _
                    "," & OnSummary & _
                    ",'" & Comments & _
                    "')"

        Dim CalcTotal = 0, CalcVat = 0, CalcNett = 0

        Dim command As New OleDb.OleDbCommand(sql4NewInvoice, DbConnection)
        Try
            DbConnection.Open()
            command.ExecuteNonQuery()

            'Create Invoice Lines

            Dim lcv
            Dim CurTypeCode = "N/A"
            Dim TotalLengthForType = 0, TypeMass = 0
            Dim qty = txtQty.Text
            Dim Total = Double.Parse(txtQty.Text) * Double.Parse(txtLength.Text) * Double.Parse(txtWidth.Text) * Double.Parse(txtRate.Text)
            Dim LineNumberCounter As Int16 = 1
            Dim DESCRIPTION As String = txtQty.Text & " Sheets x " & txtLength.Text & "m x " & txtWidth.Text & "m = " & Decimal.Round(Double.Parse(txtQty.Text) * Double.Parse(txtLength.Text) * Double.Parse(txtWidth.Text), 2) & "m²  @  "


            Dim SQL4NewInvoiceLine As String = "INSERT INTO InvoiceLine(InvNo,[Line#],TypeCode,Description,Qty,TonsorKg,CostPerUnit,Total) VALUES " & _
            "(" & InvoiceNumber & _
            "," & LineNumberCounter & _
            ",'" & CurTypeCode & _
            "','" & DESCRIPTION & _
            "'," & qty & _
            ",'" & "N/A" & _
            "'," & txtRate.Text & _
            "," & Total & _
            ")"

            CalcTotal += Total

            Dim InvLineCommand As New OleDb.OleDbCommand(SQL4NewInvoiceLine, DbConnection)
            Try
                InvLineCommand.ExecuteNonQuery()
            Catch mex As Exception
                MessageBox.Show(mex.Message)
            End Try


            CalcNett = CalcTotal * (1 + Single.Parse(VAT))
            CalcVat = CalcTotal * Single.Parse(VAT)

            Dim SQL4Totals As String = "UPDATE Invoice SET InvTotal = " & Decimal.Round(CalcTotal, 2) & ", InvVatAmt = " & Decimal.Round(CalcVat, 2) & " , InvNett = " & Decimal.Round(CalcNett, 2) & " WHERE InvoiceNo = " & InvoiceNumber
            command = New OleDb.OleDbCommand(SQL4Totals, DbConnection)

            Try
                command.ExecuteNonQuery()
            Catch ex As Exception
                MessageBox.Show(ex.Message)

            End Try

            'Update Company's Last Invoice Number
            Dim SQL4CompanyUpdate As String = "UPDATE Company SET Company.LastInvNum = " + InvoiceNumber.ToString() + " WHERE Company.CompanyNo = '" + ds.Tables(0).Rows(0).Item("Company.CompanyNo").ToString() + "'"
            Dim UpdateCommand As New OleDb.OleDbCommand(SQL4CompanyUpdate, DbConnection)
            Try
                UpdateCommand.ExecuteNonQuery()
            Catch MEEE As Exception
                MessageBox.Show(MEEE.Message)
            End Try

            lblNotify.Text = "Invoice Number: " + InvoiceNumber.ToString
            btnPrint.Enabled = True
            cmbJobNum.SelectedIndex = -1
        Catch Myerror As Exception
            MessageBox.Show(Myerror.Message)
        Finally
            DbConnection.Close()
        End Try

        btnCreateInvoice.Enabled = False

    End Sub


    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim Form As PrintCutInv = New PrintCutInv(Me)
        Form.populate_invoiceNumbers()
        Form.txt_InvNumToPrint.SelectedIndex = Form.txt_InvNumToPrint.Items.IndexOf(InvoiceNumber.ToString)
        Form.btn_PrintInv_Click(sender, e)

    End Sub

    Private Sub cmbJobNum_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbJobNum.SelectedIndexChanged
        cmbJobName.SelectedIndex = cmbJobNum.SelectedIndex
    End Sub

    Private Sub cmbJobName_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbJobName.SelectedIndexChanged
        cmbJobNum.SelectedIndex = cmbJobName.SelectedIndex
    End Sub

    Private Sub txtAddDetails_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAddDetails.TextChanged

    End Sub

    Private Sub btnNewInvoice_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNewInvoice.Click
        btnCreateInvoice.Enabled = True
        txtAddDetails.Clear()
        txtAdditional.Clear()
        txtDelNoteNum.Clear()
        txtJobName.Clear()

        txtOrderNum.Clear()
        txtQty.Clear()
        txtRate.Clear()
        txtRefNo.Clear()

        cmbJobName.Text = ""
        cmbJobNum.Text = ""

    End Sub
End Class
