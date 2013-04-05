Public Class frmSundryInv
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Dim DBConnection As OleDb.OleDbConnection
    Dim CallingForm As Object

    Public Sub New(ByRef caller As Object, ByRef dbc As OleDb.OleDbConnection)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        CallingForm = caller
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
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents lblJobNo As System.Windows.Forms.Label
    Friend WithEvents cmb_jobNo As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents DateTimePicker1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtQty As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtCost As System.Windows.Forms.TextBox
    Friend WithEvents lvwItems As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents btnAddItem As System.Windows.Forms.Button
    Friend WithEvents btnRemoveItem As System.Windows.Forms.Button
    Friend WithEvents lblGrandTotal As System.Windows.Forms.Label
    Friend WithEvents txtDelNoteNum As System.Windows.Forms.TextBox
    Friend WithEvents txtOrderNum As System.Windows.Forms.TextBox
    Friend WithEvents btnCreateInv As System.Windows.Forms.Button
    Friend WithEvents txtInvHeading As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents txtDesc As System.Windows.Forms.TextBox
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents lblNotify As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtdet1 As System.Windows.Forms.TextBox
    Friend WithEvents txtdet2 As System.Windows.Forms.TextBox
    Friend WithEvents cmbJobName As System.Windows.Forms.ComboBox
    Friend WithEvents btn_newinvoice As System.Windows.Forms.Button
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtJobName As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btnCreateInv = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.lblJobNo = New System.Windows.Forms.Label
        Me.cmb_jobNo = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtDelNoteNum = New System.Windows.Forms.TextBox
        Me.txtOrderNum = New System.Windows.Forms.TextBox
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.lblGrandTotal = New System.Windows.Forms.Label
        Me.lvwItems = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.btnAddItem = New System.Windows.Forms.Button
        Me.txtQty = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtCost = New System.Windows.Forms.TextBox
        Me.btnRemoveItem = New System.Windows.Forms.Button
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtDesc = New System.Windows.Forms.TextBox
        Me.txtInvHeading = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.btnPrint = New System.Windows.Forms.Button
        Me.lblNotify = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.txtdet1 = New System.Windows.Forms.TextBox
        Me.txtdet2 = New System.Windows.Forms.TextBox
        Me.cmbJobName = New System.Windows.Forms.ComboBox
        Me.btn_newinvoice = New System.Windows.Forms.Button
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.txtJobName = New System.Windows.Forms.TextBox
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnCreateInv
        '
        Me.btnCreateInv.Location = New System.Drawing.Point(248, 512)
        Me.btnCreateInv.Name = "btnCreateInv"
        Me.btnCreateInv.Size = New System.Drawing.Size(128, 23)
        Me.btnCreateInv.TabIndex = 0
        Me.btnCreateInv.Text = "Create Invoice"
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(552, 512)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(128, 23)
        Me.btnClose.TabIndex = 1
        Me.btnClose.Text = "Close"
        '
        'lblJobNo
        '
        Me.lblJobNo.Location = New System.Drawing.Point(16, 16)
        Me.lblJobNo.Name = "lblJobNo"
        Me.lblJobNo.Size = New System.Drawing.Size(72, 23)
        Me.lblJobNo.TabIndex = 4
        Me.lblJobNo.Text = "Job No"
        Me.lblJobNo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cmb_jobNo
        '
        Me.cmb_jobNo.Location = New System.Drawing.Point(104, 16)
        Me.cmb_jobNo.Name = "cmb_jobNo"
        Me.cmb_jobNo.Size = New System.Drawing.Size(80, 21)
        Me.cmb_jobNo.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 88)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(72, 23)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Delivery Note"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(448, 56)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 23)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Order No"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(448, 120)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(72, 23)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "Invoice Date"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtDelNoteNum
        '
        Me.txtDelNoteNum.Location = New System.Drawing.Point(104, 88)
        Me.txtDelNoteNum.Name = "txtDelNoteNum"
        Me.txtDelNoteNum.Size = New System.Drawing.Size(584, 20)
        Me.txtDelNoteNum.TabIndex = 3
        Me.txtDelNoteNum.Text = ""
        '
        'txtOrderNum
        '
        Me.txtOrderNum.Location = New System.Drawing.Point(528, 56)
        Me.txtOrderNum.Name = "txtOrderNum"
        Me.txtOrderNum.Size = New System.Drawing.Size(160, 20)
        Me.txtOrderNum.TabIndex = 2
        Me.txtOrderNum.Text = ""
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.CustomFormat = ""
        Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.DateTimePicker1.Location = New System.Drawing.Point(528, 120)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(120, 20)
        Me.DateTimePicker1.TabIndex = 6
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.lblGrandTotal)
        Me.GroupBox1.Controls.Add(Me.lvwItems)
        Me.GroupBox1.Controls.Add(Me.btnAddItem)
        Me.GroupBox1.Controls.Add(Me.txtQty)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.txtCost)
        Me.GroupBox1.Controls.Add(Me.btnRemoveItem)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.txtDesc)
        Me.GroupBox1.Location = New System.Drawing.Point(96, 192)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(584, 304)
        Me.GroupBox1.TabIndex = 10
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Items"
        '
        'lblGrandTotal
        '
        Me.lblGrandTotal.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGrandTotal.Location = New System.Drawing.Point(312, 240)
        Me.lblGrandTotal.Name = "lblGrandTotal"
        Me.lblGrandTotal.Size = New System.Drawing.Size(208, 23)
        Me.lblGrandTotal.TabIndex = 10
        Me.lblGrandTotal.Text = "Total Ex. VAT: R0.00"
        Me.lblGrandTotal.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lvwItems
        '
        Me.lvwItems.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4})
        Me.lvwItems.FullRowSelect = True
        Me.lvwItems.Location = New System.Drawing.Point(8, 104)
        Me.lvwItems.MultiSelect = False
        Me.lvwItems.Name = "lvwItems"
        Me.lvwItems.Size = New System.Drawing.Size(512, 120)
        Me.lvwItems.TabIndex = 9
        Me.lvwItems.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Description"
        Me.ColumnHeader1.Width = 263
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Num of Units"
        Me.ColumnHeader2.Width = 76
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Cost per Unit"
        Me.ColumnHeader3.Width = 78
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "Total"
        Me.ColumnHeader4.Width = 86
        '
        'btnAddItem
        '
        Me.btnAddItem.Location = New System.Drawing.Point(432, 56)
        Me.btnAddItem.Name = "btnAddItem"
        Me.btnAddItem.Size = New System.Drawing.Size(88, 23)
        Me.btnAddItem.TabIndex = 4
        Me.btnAddItem.Text = "Add"
        '
        'txtQty
        '
        Me.txtQty.Location = New System.Drawing.Point(120, 56)
        Me.txtQty.Name = "txtQty"
        Me.txtQty.TabIndex = 2
        Me.txtQty.Text = ""
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(16, 56)
        Me.Label6.Name = "Label6"
        Me.Label6.TabIndex = 4
        Me.Label6.Text = "Number of Units"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(232, 56)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(80, 23)
        Me.Label7.TabIndex = 4
        Me.Label7.Text = "Cost per Unit"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtCost
        '
        Me.txtCost.Location = New System.Drawing.Point(312, 56)
        Me.txtCost.Name = "txtCost"
        Me.txtCost.TabIndex = 3
        Me.txtCost.Text = ""
        '
        'btnRemoveItem
        '
        Me.btnRemoveItem.Location = New System.Drawing.Point(16, 240)
        Me.btnRemoveItem.Name = "btnRemoveItem"
        Me.btnRemoveItem.Size = New System.Drawing.Size(144, 23)
        Me.btnRemoveItem.TabIndex = 8
        Me.btnRemoveItem.Text = "Remove"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(16, 24)
        Me.Label5.Name = "Label5"
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "Description"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtDesc
        '
        Me.txtDesc.Location = New System.Drawing.Point(120, 24)
        Me.txtDesc.Name = "txtDesc"
        Me.txtDesc.Size = New System.Drawing.Size(400, 20)
        Me.txtDesc.TabIndex = 1
        Me.txtDesc.Text = ""
        '
        'txtInvHeading
        '
        Me.txtInvHeading.Location = New System.Drawing.Point(528, 152)
        Me.txtInvHeading.Name = "txtInvHeading"
        Me.txtInvHeading.Size = New System.Drawing.Size(280, 20)
        Me.txtInvHeading.TabIndex = 7
        Me.txtInvHeading.Text = ""
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(448, 152)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(48, 23)
        Me.Label8.TabIndex = 4
        Me.Label8.Text = "Heading"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnPrint
        '
        Me.btnPrint.Enabled = False
        Me.btnPrint.Location = New System.Drawing.Point(400, 512)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(128, 23)
        Me.btnPrint.TabIndex = 0
        Me.btnPrint.Text = "Print Preview..."
        '
        'lblNotify
        '
        Me.lblNotify.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNotify.Location = New System.Drawing.Point(120, 464)
        Me.lblNotify.Name = "lblNotify"
        Me.lblNotify.Size = New System.Drawing.Size(488, 23)
        Me.lblNotify.TabIndex = 9
        Me.lblNotify.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(16, 120)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(56, 23)
        Me.Label9.TabIndex = 4
        Me.Label9.Text = "Details"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtdet1
        '
        Me.txtdet1.Location = New System.Drawing.Point(104, 120)
        Me.txtdet1.Name = "txtdet1"
        Me.txtdet1.Size = New System.Drawing.Size(288, 20)
        Me.txtdet1.TabIndex = 4
        Me.txtdet1.Text = ""
        '
        'txtdet2
        '
        Me.txtdet2.Location = New System.Drawing.Point(104, 152)
        Me.txtdet2.Name = "txtdet2"
        Me.txtdet2.Size = New System.Drawing.Size(288, 20)
        Me.txtdet2.TabIndex = 5
        Me.txtdet2.Text = ""
        '
        'cmbJobName
        '
        Me.cmbJobName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbJobName.Location = New System.Drawing.Point(528, 16)
        Me.cmbJobName.Name = "cmbJobName"
        Me.cmbJobName.Size = New System.Drawing.Size(200, 21)
        Me.cmbJobName.TabIndex = 14
        '
        'btn_newinvoice
        '
        Me.btn_newinvoice.Location = New System.Drawing.Point(96, 512)
        Me.btn_newinvoice.Name = "btn_newinvoice"
        Me.btn_newinvoice.Size = New System.Drawing.Size(128, 23)
        Me.btn_newinvoice.TabIndex = 1
        Me.btn_newinvoice.Text = "New Invoice"
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(424, 16)
        Me.Label10.Name = "Label10"
        Me.Label10.TabIndex = 12
        Me.Label10.Text = "Job Name Search"
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(16, 56)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(64, 23)
        Me.Label11.TabIndex = 13
        Me.Label11.Text = "Job Name"
        '
        'txtJobName
        '
        Me.txtJobName.Enabled = False
        Me.txtJobName.Location = New System.Drawing.Point(104, 56)
        Me.txtJobName.Name = "txtJobName"
        Me.txtJobName.Size = New System.Drawing.Size(296, 20)
        Me.txtJobName.TabIndex = 17
        Me.txtJobName.Text = ""
        '
        'frmSundryInv
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(806, 550)
        Me.Controls.Add(Me.txtJobName)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.lblNotify)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.DateTimePicker1)
        Me.Controls.Add(Me.txtDelNoteNum)
        Me.Controls.Add(Me.txtOrderNum)
        Me.Controls.Add(Me.txtInvHeading)
        Me.Controls.Add(Me.txtdet1)
        Me.Controls.Add(Me.txtdet2)
        Me.Controls.Add(Me.cmb_jobNo)
        Me.Controls.Add(Me.lblJobNo)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnCreateInv)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.cmbJobName)
        Me.Controls.Add(Me.btn_newinvoice)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "frmSundryInv"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Sundry Invoices"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim ItemCount = 0

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        'Close the form
        Me.Close()
        CallingForm.show()
        CallingForm = Nothing
    End Sub

    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, ByVal keyData As Keys) As Boolean
        If keyData = Keys.Enter Then
            SendKeys.Send("{Tab}")
            Return True
        End If

        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function

    Private Sub frmSundryInv_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        populateCmbJobs()
    End Sub

    Private Sub populateCmbJobs()
        cmbJobName.Items.Clear()
        cmb_jobNo.Items.Clear()
        Dim sql = "SELECT jobno,jobname FROM job ORDER BY jobno"
        Dim ds As Data.DataSet = New Data.DataSet
        Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(sql, DBConnection)
        da.Fill(ds)

        Dim i
        For i = 0 To ds.Tables(0).Rows.Count - 1
            cmb_jobNo.Items.Add(ds.Tables(0).Rows(i).Item("JobNo").ToString())
            cmbJobName.Items.Add(ds.Tables(0).Rows(i).Item("JobName").ToString())
        Next i

    End Sub

    Private Sub btnAddItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddItem.Click
        If txtQty.Text.Trim = "" Then
            txtQty.Text = 0
        End If
        If txtCost.Text.Trim = "" Then
            txtCost.Text = 0
        End If
        ItemCount += 1
        Dim item As ListViewItem
        item = New ListViewItem(txtDesc.Text)
        item.SubItems.Add(txtQty.Text)
        item.SubItems.Add(txtCost.Text)

        Try
            item.SubItems.Add(Format((Double.Parse(txtQty.Text) * Double.Parse(txtCost.Text)), "######0.00"))
        Catch Ex As Exception
            MessageBox.Show("Invalid numeric value.")
            Exit Sub
        End Try


        lvwItems.Items.Add(item)
        updateTotal()
        txtCost.Clear()
        txtQty.Clear()
        txtDesc.Clear()
        txtDesc.Focus()
    End Sub

    Private Sub updateTotal()
        Dim i
        Dim tot As Double = 0
        For i = 0 To lvwItems.Items.Count - 1
            tot += Double.Parse(lvwItems.Items(i).SubItems(3).Text)
        Next

        lblGrandTotal.Text = "Total Ex. VAT: " & PrintCutInv.toRand(tot.ToString, True)

    End Sub

    Private Sub btnRemoveItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemoveItem.Click
        If lvwItems.SelectedItems.Count <> 0 Then
            lvwItems.Items.Remove(lvwItems.SelectedItems.Item(0))
            ItemCount -= 1
            updateTotal()
        End If



    End Sub

    Private Sub cmb_jobNo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb_jobNo.Leave
        If cmb_jobNo.Text <> "" Then
            getJobDetails()
        End If

        lblNotify.Text = ""
        btnPrint.Enabled = False
    End Sub
    Private Sub cmbJobName_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbJobName.Leave
        cmb_jobNo_SelectedIndexChanged(sender, e)
    End Sub


    Private Sub getJobDetails()
        Dim sql = "SELECT jobName, orderNo,ContractorName FROM (job INNER JOIN Contractor On Job.ContractorNo = Contractor.ContractorNo) WHERE jobno = '" & cmb_jobNo.Text & "'"
        Dim ds As Data.DataSet = New Data.DataSet
        Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(sql, DBConnection)
        da.Fill(ds)

        Try
            txtOrderNum.Text = ds.Tables(0).Rows(0).Item("orderNo").ToString
            txtJobName.Text = ds.Tables(0).Rows(0).Item("jobName").ToString
            'txtContractor.Text = ds.Tables(0).Rows(0).Item("ContractorName").ToString
        Catch ex As Exception
            MessageBox.Show("There are no jobs matching this job number.", "Invalid Job Number", MessageBoxButtons.OK, MessageBoxIcon.Information)
            cmb_jobNo.Focus()
            cmb_jobNo.SelectAll()
        End Try

    End Sub
    Dim InvoiceNumber As Long
    Private Sub btnCreateInv_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateInv.Click
        If lvwItems.Items.Count = 0 Then
            MessageBox.Show("Please enter at least one item before creating the invoice.", "No Items", MessageBoxButtons.OK)
            txtQty.Focus()
            Exit Sub
        End If

        If cmb_jobNo.Text.Trim = "" Then
            MessageBox.Show("Please enter a valid job number.", "No Job Selected", MessageBoxButtons.OK, MessageBoxIcon.Information)
            cmb_jobNo.Focus()
            Exit Sub
        End If

        Dim sql = "SELECT * FROM (Job INNER JOIN Company ON Job.companyNo = Company.CompanyNo) INNER JOIN Contractor ON Job.ContractorNo = Contractor.ContractorNo WHERE Job.JobNo = '" & cmb_jobNo.Text & "'"
        Dim ds As Data.DataSet = New Data.DataSet
        Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(sql, DBConnection)
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

        Dim InvoiceType As String = "Sundry"
        Dim Factor As Int16 = 1
        Dim EscMoAndDa As Date = New Date(1999, 12, 11)
        Dim Work As Int16 = 75
        Dim RefNo As String = txtdet1.Text
        Dim JobNo As String = ds.Tables(0).Rows(0).Item("JobNo").ToString()
        Dim VAT As String = ds.Tables(0).Rows(0).Item("VatPerc").ToString()
        Dim Active = "Yes"
        Dim Escalated = "No"
        Dim OnSummary = "Yes"
        Dim Comments As String = txtdet2.Text
        Dim invDate As Date
        invDate = DateTimePicker1.Value.Date
        '"',#" & Format(invDate, "dd/MM/yyyy") & _'
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
                    "','" & txtInvHeading.Text & _
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

        Dim command As New OleDb.OleDbCommand(sql4NewInvoice, DBConnection)
        Try
            DBConnection.Open()
            command.ExecuteNonQuery()

            'Create Invoice Lines

            Dim lcv
            Dim CurTypeCode = "N/A"
            Dim TotalLengthForType = 0, TypeMass = 0
            Dim qty = txtQty.Text
            Dim Total = 0, rate = 0
            Dim LineNumberCounter As Int16 = 1
            Dim DESCRIPTION As String = " "

            For lcv = 0 To lvwItems.Items.Count - 1
                DESCRIPTION = lvwItems.Items(lcv).Text
                qty = lvwItems.Items(lcv).SubItems(1).Text
                rate = lvwItems.Items(lcv).SubItems(2).Text
                Total = Double.Parse(Format((Double.Parse(qty) * Double.Parse(rate)), "#####0.00"))
                Dim SQL4NewInvoiceLine As String = "INSERT INTO InvoiceLine(InvNo,[Line#],TypeCode,Description,Qty,TonsorKg,CostPerUnit,Total) VALUES " & _
                "(" & InvoiceNumber & _
                "," & LineNumberCounter & _
                ",'" & CurTypeCode & _
                "','" & DESCRIPTION & _
                "'," & qty & _
                ",'" & "N/A" & _
                "'," & rate & _
                "," & Total & _
                    ")"

                CalcTotal += Total

                Dim InvLineCommand As New OleDb.OleDbCommand(SQL4NewInvoiceLine, DBConnection)
                Try
                    InvLineCommand.ExecuteNonQuery()
                Catch mex As Exception
                    MessageBox.Show(mex.Message)
                End Try

                LineNumberCounter += 1

            Next lcv


            CalcNett = CalcTotal * (1 + Single.Parse(VAT))
            CalcVat = CalcTotal * Single.Parse(VAT)

            Dim SQL4Totals As String = "UPDATE Invoice SET InvTotal = " & Decimal.Round(CalcTotal, 2) & ", InvVatAmt = " & Decimal.Round(CalcVat, 2) & " , InvNett = " & Decimal.Round(CalcNett, 2) & " WHERE InvoiceNo = " & InvoiceNumber
            command = New OleDb.OleDbCommand(SQL4Totals, DBConnection)

            Try
                command.ExecuteNonQuery()
            Catch ex As Exception
                MessageBox.Show(ex.Message)

            End Try

            'Update Company's Last Invoice Number
            Dim SQL4CompanyUpdate As String = "UPDATE Company SET Company.LastInvNum = " + InvoiceNumber.ToString() + " WHERE Company.CompanyNo = '" + ds.Tables(0).Rows(0).Item("Company.CompanyNo").ToString() + "'"
            Dim UpdateCommand As New OleDb.OleDbCommand(SQL4CompanyUpdate, DBConnection)
            Try
                UpdateCommand.ExecuteNonQuery()
            Catch MEEE As Exception
                MessageBox.Show(MEEE.Message)
            End Try


            lblNotify.Text = "Invoice Number: " + InvoiceNumber.ToString
            btnPrint.Enabled = True
            cmb_jobNo.SelectedIndex = -1

        Catch Myerror As Exception
            MessageBox.Show(Myerror.Message)
        Finally
            DBConnection.Close()
        End Try

        btnCreateInv.Enabled = False

    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim Form As PrintCutInv = New PrintCutInv(Me)
        Form.populate_invoiceNumbers()
        Form.txt_InvNumToPrint.SelectedIndex = Form.txt_InvNumToPrint.Items.IndexOf(InvoiceNumber.ToString)
        Form.btn_PrintInv_Click(sender, e)

    End Sub

    Private Sub txtCost_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCost.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(32) Then
            btnAddItem.PerformClick()
        End If
    End Sub


    Private Sub cmbJobNum_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmb_jobNo.SelectedIndexChanged
        cmbJobName.SelectedIndex = cmb_jobNo.SelectedIndex
    End Sub

    Private Sub cmbJobName_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbJobName.SelectedIndexChanged
        cmb_jobNo.SelectedIndex = cmbJobName.SelectedIndex
    End Sub


    Private Sub btn_newinvoice_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_newinvoice.Click
        txtJobName.Clear()
        txtCost.Clear()
        txtDelNoteNum.Clear()
        txtDesc.Clear()
        txtdet1.Clear()
        txtdet2.Clear()
        txtInvHeading.Clear()
        txtOrderNum.Clear()
        txtJobName.Clear()
        txtQty.Clear()
        lvwItems.Items.Clear()
        lblNotify.Text = ""
        lblGrandTotal.Text = "Total Ex. VAT: R0.00"
        btnCreateInv.Enabled = True
        btnPrint.Enabled = False
    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub
End Class
