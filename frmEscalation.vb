Public Class frmEscalation
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Dim callingForm As Object
    Dim DBConnection As OleDb.OleDbConnection


    Public Sub New(ByVal c As Object, ByVal dbcon As OleDb.OleDbConnection)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        DBConnection = dbcon
        callingForm = c
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
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cmbMonth As System.Windows.Forms.ComboBox
    Friend WithEvents nudYear As System.Windows.Forms.NumericUpDown
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents cmbJobNum As System.Windows.Forms.ComboBox
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents txtOrderNum As System.Windows.Forms.TextBox
    Friend WithEvents txtEscFactor As System.Windows.Forms.TextBox
    Friend WithEvents btnCreateEscInv As System.Windows.Forms.Button
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtDet1 As System.Windows.Forms.TextBox
    Friend WithEvents txtDet2 As System.Windows.Forms.TextBox
    Friend WithEvents dtpInvdate As System.Windows.Forms.DateTimePicker
    Friend WithEvents cut As System.Windows.Forms.RadioButton
    Friend WithEvents mesh As System.Windows.Forms.RadioButton
    Friend WithEvents sundry As System.Windows.Forms.RadioButton
    Friend WithEvents all As System.Windows.Forms.RadioButton
    Friend WithEvents lblNotify As System.Windows.Forms.Label
    Friend WithEvents cmbJobName As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.cut = New System.Windows.Forms.RadioButton
        Me.mesh = New System.Windows.Forms.RadioButton
        Me.sundry = New System.Windows.Forms.RadioButton
        Me.all = New System.Windows.Forms.RadioButton
        Me.txtEscFactor = New System.Windows.Forms.TextBox
        Me.txtOrderNum = New System.Windows.Forms.TextBox
        Me.cmbMonth = New System.Windows.Forms.ComboBox
        Me.nudYear = New System.Windows.Forms.NumericUpDown
        Me.dtpInvdate = New System.Windows.Forms.DateTimePicker
        Me.btnCreateEscInv = New System.Windows.Forms.Button
        Me.btnPrint = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.cmbJobNum = New System.Windows.Forms.ComboBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtDet1 = New System.Windows.Forms.TextBox
        Me.txtDet2 = New System.Windows.Forms.TextBox
        Me.lblNotify = New System.Windows.Forms.Label
        Me.cmbJobName = New System.Windows.Forms.ComboBox
        Me.GroupBox1.SuspendLayout()
        CType(Me.nudYear, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Job Number:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(256, 48)
        Me.Label2.Name = "Label2"
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Month:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(256, 80)
        Me.Label3.Name = "Label3"
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "Year:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(256, 112)
        Me.Label4.Name = "Label4"
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "Escalation Factor:"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(16, 80)
        Me.Label5.Name = "Label5"
        Me.Label5.TabIndex = 0
        Me.Label5.Text = "Date:"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(16, 48)
        Me.Label6.Name = "Label6"
        Me.Label6.TabIndex = 0
        Me.Label6.Text = "Order Number:"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cut)
        Me.GroupBox1.Controls.Add(Me.mesh)
        Me.GroupBox1.Controls.Add(Me.sundry)
        Me.GroupBox1.Controls.Add(Me.all)
        Me.GroupBox1.Location = New System.Drawing.Point(16, 232)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(224, 136)
        Me.GroupBox1.TabIndex = 9
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Invoice Types to Escalate"
        '
        'cut
        '
        Me.cut.Location = New System.Drawing.Point(24, 24)
        Me.cut.Name = "cut"
        Me.cut.TabIndex = 0
        Me.cut.Text = "Cutting Sheet"
        '
        'mesh
        '
        Me.mesh.Location = New System.Drawing.Point(24, 48)
        Me.mesh.Name = "mesh"
        Me.mesh.TabIndex = 0
        Me.mesh.Text = "Mesh"
        '
        'sundry
        '
        Me.sundry.Location = New System.Drawing.Point(24, 72)
        Me.sundry.Name = "sundry"
        Me.sundry.TabIndex = 0
        Me.sundry.Text = "Sundry"
        '
        'all
        '
        Me.all.Checked = True
        Me.all.Location = New System.Drawing.Point(24, 96)
        Me.all.Name = "all"
        Me.all.TabIndex = 0
        Me.all.TabStop = True
        Me.all.Text = "All"
        '
        'txtEscFactor
        '
        Me.txtEscFactor.Location = New System.Drawing.Point(360, 112)
        Me.txtEscFactor.Name = "txtEscFactor"
        Me.txtEscFactor.Size = New System.Drawing.Size(120, 20)
        Me.txtEscFactor.TabIndex = 6
        Me.txtEscFactor.Text = ""
        '
        'txtOrderNum
        '
        Me.txtOrderNum.Location = New System.Drawing.Point(120, 48)
        Me.txtOrderNum.Name = "txtOrderNum"
        Me.txtOrderNum.Size = New System.Drawing.Size(120, 20)
        Me.txtOrderNum.TabIndex = 3
        Me.txtOrderNum.Text = ""
        '
        'cmbMonth
        '
        Me.cmbMonth.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbMonth.Items.AddRange(New Object() {"January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"})
        Me.cmbMonth.Location = New System.Drawing.Point(360, 48)
        Me.cmbMonth.MaxDropDownItems = 12
        Me.cmbMonth.Name = "cmbMonth"
        Me.cmbMonth.Size = New System.Drawing.Size(120, 21)
        Me.cmbMonth.TabIndex = 2
        '
        'nudYear
        '
        Me.nudYear.Location = New System.Drawing.Point(360, 80)
        Me.nudYear.Maximum = New Decimal(New Integer() {2999, 0, 0, 0})
        Me.nudYear.Minimum = New Decimal(New Integer() {1969, 0, 0, 0})
        Me.nudYear.Name = "nudYear"
        Me.nudYear.TabIndex = 4
        Me.nudYear.Value = New Decimal(New Integer() {2003, 0, 0, 0})
        '
        'dtpInvdate
        '
        Me.dtpInvdate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpInvdate.Location = New System.Drawing.Point(120, 80)
        Me.dtpInvdate.Name = "dtpInvdate"
        Me.dtpInvdate.Size = New System.Drawing.Size(120, 20)
        Me.dtpInvdate.TabIndex = 5
        '
        'btnCreateEscInv
        '
        Me.btnCreateEscInv.Location = New System.Drawing.Point(272, 264)
        Me.btnCreateEscInv.Name = "btnCreateEscInv"
        Me.btnCreateEscInv.Size = New System.Drawing.Size(192, 24)
        Me.btnCreateEscInv.TabIndex = 10
        Me.btnCreateEscInv.Text = "Create Invoice"
        '
        'btnPrint
        '
        Me.btnPrint.Enabled = False
        Me.btnPrint.Location = New System.Drawing.Point(272, 304)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(192, 24)
        Me.btnPrint.TabIndex = 11
        Me.btnPrint.Text = "Print Preview..."
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(272, 344)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(192, 24)
        Me.btnClose.TabIndex = 12
        Me.btnClose.Text = "Close"
        '
        'cmbJobNum
        '
        Me.cmbJobNum.Location = New System.Drawing.Point(120, 16)
        Me.cmbJobNum.Name = "cmbJobNum"
        Me.cmbJobNum.Size = New System.Drawing.Size(88, 21)
        Me.cmbJobNum.TabIndex = 1
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(16, 144)
        Me.Label7.Name = "Label7"
        Me.Label7.TabIndex = 11
        Me.Label7.Text = "Details:"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtDet1
        '
        Me.txtDet1.Location = New System.Drawing.Point(120, 144)
        Me.txtDet1.Name = "txtDet1"
        Me.txtDet1.Size = New System.Drawing.Size(360, 20)
        Me.txtDet1.TabIndex = 7
        Me.txtDet1.Text = "ESCALATION"
        '
        'txtDet2
        '
        Me.txtDet2.Location = New System.Drawing.Point(120, 176)
        Me.txtDet2.Name = "txtDet2"
        Me.txtDet2.Size = New System.Drawing.Size(360, 20)
        Me.txtDet2.TabIndex = 8
        Me.txtDet2.Text = ""
        '
        'lblNotify
        '
        Me.lblNotify.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNotify.Location = New System.Drawing.Point(272, 232)
        Me.lblNotify.Name = "lblNotify"
        Me.lblNotify.Size = New System.Drawing.Size(192, 23)
        Me.lblNotify.TabIndex = 13
        Me.lblNotify.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cmbJobName
        '
        Me.cmbJobName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbJobName.Location = New System.Drawing.Point(208, 16)
        Me.cmbJobName.Name = "cmbJobName"
        Me.cmbJobName.Size = New System.Drawing.Size(272, 21)
        Me.cmbJobName.TabIndex = 1
        '
        'frmEscalation
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(490, 384)
        Me.Controls.Add(Me.lblNotify)
        Me.Controls.Add(Me.txtDet1)
        Me.Controls.Add(Me.txtEscFactor)
        Me.Controls.Add(Me.txtOrderNum)
        Me.Controls.Add(Me.txtDet2)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.cmbJobNum)
        Me.Controls.Add(Me.btnCreateEscInv)
        Me.Controls.Add(Me.dtpInvdate)
        Me.Controls.Add(Me.cmbMonth)
        Me.Controls.Add(Me.nudYear)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.cmbJobName)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "frmEscalation"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Create an Escalation Invoice"
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.nudYear, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

  
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
        callingForm.show()
        callingForm = Nothing
    End Sub

    Private Sub frmEscalation_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim currentDate As DateTime = Today
        cmbMonth.SelectedIndex = currentDate.Month - 1
        nudYear.Value = currentDate.Year
        populateCmbJobs()
    End Sub

    Private Sub populateCmbJobs()
        cmbJobNum.Items.Clear()
        cmbJobName.Items.Clear()
        Dim sql As String = "SELECT jobno,jobname FROM job ORDER BY jobno"
        Dim ds As Data.DataSet = New Data.DataSet
        Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(sql, DBConnection)
        da.Fill(ds)

        Dim i As Integer
        For i = 0 To ds.Tables(0).Rows.Count - 1
            cmbJobNum.Items.Add(ds.Tables(0).Rows(i).Item("JobNo").ToString())
            cmbJobName.Items.Add(ds.Tables(0).Rows(i).Item("JobName").ToString())
        Next i

    End Sub

    Private Sub cmbJobNum_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbJobNum.Leave
        If cmbJobNum.Text <> "" Then
            getJobOrderNum()
        End If

        btnPrint.Enabled = False

    End Sub

    Private Sub cmbJobName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbJobName.Leave
        cmbJobNum_SelectedIndexChanged(sender, e)
    End Sub


    Private Sub getJobOrderNum()
        Dim sql As String = "SELECT orderNo FROM job WHERE jobno = '" & cmbJobNum.Text & "'"
        Dim ds As Data.DataSet = New Data.DataSet
        Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(sql, DBConnection)
        da.Fill(ds)

        Try
            txtOrderNum.Text = ds.Tables(0).Rows(0).Item("orderNo").ToString
        Catch ex As Exception
            MessageBox.Show("There are no jobs matching this job number.", "Invalid Job Number", MessageBoxButtons.OK, MessageBoxIcon.Information)
            cmbJobNum.Focus()
            cmbJobNum.SelectAll()
        End Try

    End Sub

    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, ByVal keyData As Keys) As Boolean
        If keyData = Keys.Enter Then
            SendKeys.Send("{Tab}")
            Return True
        End If

        Return MyBase.ProcessCmdKey(msg, keyData)

    End Function

    Private Sub txtEscFactor_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEscFactor.Leave
        If Not IsNumeric(txtEscFactor.Text) Then
            MessageBox.Show("Invalid Value. Please enter a number.", "Invalid Number", MessageBoxButtons.OK)
            txtEscFactor.Focus()
            txtEscFactor.SelectAll()
        End If
    End Sub

    Dim InvoiceNumber As Long

    Private Function getWork(ByVal job As String, ByVal month As Int16, ByVal year As Int16, ByVal mesh As Boolean, ByVal cuttingsheet As Boolean, ByVal sundry As Boolean, ByVal all As Boolean) As Double
        Dim sql As String = String.Empty
        If all Then
            sql = "SELECT InvTotal FROM Invoice WHERE InvJobNo = '" & cmbJobNum.Text & "' AND NOT InvoiceType = 'Escalation' AND YEAR(InvDate) = " & year & " AND MONTH(InvDate) = " & month
        End If
        If mesh Then
            sql = "SELECT InvTotal FROM Invoice WHERE InvJobNo = '" & cmbJobNum.Text & "' AND NOT InvoiceType = 'Escalation' AND YEAR(InvDate) = " & year & " AND MONTH(InvDate) = " & month & " AND InvoiceType = 'Mesh'"
        End If
        If sundry Then
            sql = "SELECT InvTotal FROM Invoice WHERE InvJobNo = '" & cmbJobNum.Text & "' AND NOT InvoiceType = 'Escalation' AND YEAR(InvDate) = " & year & " AND MONTH(InvDate) = " & month & " AND InvoiceType = 'Sundry'"
        End If
        If cuttingsheet Then
            sql = "SELECT InvTotal FROM Invoice WHERE InvJobNo = '" & cmbJobNum.Text & "' AND NOT InvoiceType = 'Escalation' AND YEAR(InvDate) = " & year & " AND MONTH(InvDate) = " & month & " AND InvoiceType = 'Cutting Sheet'"
        End If


        Dim ds As Data.DataSet = New Data.DataSet
        Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(sql, DBConnection)
        da.Fill(ds)

        If ds.Tables(0).Rows.Count = 0 Then
            Return -1
        End If

        Dim TotExVat As Double = 0

        Dim i As Integer
        For i = 0 To ds.Tables(0).Rows.Count - 1
            TotExVat += ds.Tables(0).Rows(i).Item("InvTotal")
        Next i

        Return TotExVat

    End Function

    Private Sub btnCreateEscInv_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateEscInv.Click
        If cmbJobNum.Text.Trim = "" Then
            MessageBox.Show("Please enter a valid job number.", "No Job Selected", MessageBoxButtons.OK, MessageBoxIcon.Information)
            cmbJobNum.Focus()
            Exit Sub
        End If

        If txtEscFactor.Text.Trim = "" Then
            MessageBox.Show("Please enter a valid escalation factor.", "Invalid Factor", MessageBoxButtons.OK, MessageBoxIcon.Information)
            txtEscFactor.Focus()
            txtEscFactor.SelectAll()
            Exit Sub
        End If

        Dim sql As String = "SELECT * FROM (Job INNER JOIN Company ON Job.companyNo = Company.CompanyNo) INNER JOIN Contractor ON Job.ContractorNo = Contractor.ContractorNo WHERE Job.JobNo = '" & cmbJobNum.Text & "'"
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


        Dim InvoiceType As String = "Escalation"
        Dim Factor As Double = Double.Parse(txtEscFactor.Text)
        Dim EscMoAndDa As Date = New Date(nudYear.Value, cmbMonth.SelectedIndex + 1, 1)
        Dim Work As Double = getWork(cmbJobNum.Text, cmbMonth.SelectedIndex + 1, nudYear.Value, mesh.Checked, cut.Checked, sundry.Checked, all.Checked)
        Dim RefNo As String = "N/A"
        Dim JobNo As String = ds.Tables(0).Rows(0).Item("JobNo").ToString()
        Dim VAT As String = ds.Tables(0).Rows(0).Item("VatPerc").ToString()
        Dim vatAmt As Double = Math.Round(Work * Double.Parse(VAT), 2)
        Dim escWork As Double = Math.Round(Work * Factor, 2)
        Dim EscVatAmt As Double = Math.Round(vatAmt * Factor, 2)
        Dim invNett As Double = escWork + EscVatAmt
        Dim Active As String = "Yes"
        Dim Escalated As String = "No"
        Dim OnSummary As String = "Yes"
        Dim Comments As String = "Comments"
        Dim invDate As Date
        invDate = dtpInvdate.Value.Date

        Dim sql4NewInvoice As String = "INSERT INTO Invoice(InvoiceNo,InvoiceType,InvDate,InvDeliveryNoteNo,InvFactor,Invmonthandyear,InvWork,InvOrdNum,InvRefNum,InvoiceHeading,InvTotal,InvVatAmt,InvDesign,InvNett,InvJobNo,InvActive,InvEscalated,InvOnSummary,InvComments) VALUES " & _
                "(  " & _
                    InvoiceNumber.ToString & _
                    ",'" & InvoiceType & _
                    "',#" & invDate.ToLongDateString & _
                    "#,'" & "N/A" & _
                    "'," & Factor & _
                    ",#" & EscMoAndDa & _
                    "#," & Work & _
                    ",'" & txtOrderNum.Text & _
                    "','" & RefNo & _
                    "','" & "N/A" & _
                    "'," & escWork & _
                    "," & EscVatAmt & _
                    "," & 0 & _
                    "," & invNett & _
                    ",'" & JobNo & _
                    "'," & Active & _
                    "," & Escalated & _
                    "," & OnSummary & _
                    ",'" & Comments & _
                    "')"

        Dim command As New OleDb.OleDbCommand(sql4NewInvoice, DBConnection)
        Try
            DBConnection.Open()
            command.ExecuteNonQuery()

            'Create Invoice Lines

            Dim SQL4NewInvoiceLine As String = "INSERT INTO InvoiceLine(InvNo,[Line#],TypeCode,Description,Qty,TonsorKg,CostPerUnit,Total) VALUES " & _
"(" & InvoiceNumber & _
"," & 1 & _
",'" & "N/A" & _
"','" & txtDet1.Text & _
"'," & -1 & _
",'" & "N/A" & _
"'," & -1 & _
"," & -1 & _
")"

            Dim InvLineCommand As New OleDb.OleDbCommand(SQL4NewInvoiceLine, DBConnection)
            Try
                InvLineCommand.ExecuteNonQuery()
            Catch mex As Exception
                MessageBox.Show(mex.Message)
            End Try

            SQL4NewInvoiceLine = "INSERT INTO InvoiceLine(InvNo,[Line#],TypeCode,Description,Qty,TonsorKg,CostPerUnit,Total) VALUES " & _
"(" & InvoiceNumber & _
"," & 2 & _
",'" & "N/A" & _
"','" & txtDet2.Text & _
"'," & -1 & _
",'" & "N/A" & _
"'," & -1 & _
"," & -1 & _
")"

            InvLineCommand = New OleDb.OleDbCommand(SQL4NewInvoiceLine, DBConnection)
            Try
                InvLineCommand.ExecuteNonQuery()
            Catch mex As Exception
                MessageBox.Show(mex.Message)
            End Try

            SQL4NewInvoiceLine = "INSERT INTO InvoiceLine(InvNo,[Line#],TypeCode,Description,Qty,TonsorKg,CostPerUnit,Total) VALUES " & _
"(" & InvoiceNumber & _
"," & 3 & _
",'" & "N/A" & _
"','" & cmbMonth.Text & " " & nudYear.Text & _
"'," & -1 & _
",'" & "N/A" & _
"'," & -1 & _
"," & -1 & _
")"

            InvLineCommand = New OleDb.OleDbCommand(SQL4NewInvoiceLine, DBConnection)
            Try
                InvLineCommand.ExecuteNonQuery()
            Catch mex As Exception
                MessageBox.Show(mex.Message)
            End Try

            SQL4NewInvoiceLine = "INSERT INTO InvoiceLine(InvNo,[Line#],TypeCode,Description,Qty,TonsorKg,CostPerUnit,Total) VALUES " & _
"(" & InvoiceNumber & _
"," & 4 & _
",'" & "N/A" & _
"','" & Factor.ToString & _
"'," & -1 & _
",'" & "N/A" & _
"'," & -1 & _
"," & -1 & _
")"

            InvLineCommand = New OleDb.OleDbCommand(SQL4NewInvoiceLine, DBConnection)
            Try
                InvLineCommand.ExecuteNonQuery()
            Catch mex As Exception
                MessageBox.Show(mex.Message)
            End Try

            SQL4NewInvoiceLine = "INSERT INTO InvoiceLine(InvNo,[Line#],TypeCode,Description,Qty,TonsorKg,CostPerUnit,Total) VALUES " & _
"(" & InvoiceNumber & _
"," & 5 & _
",'" & "N/A" & _
"','" & toRand(Work, True) & _
"'," & -1 & _
",'" & "N/A" & _
"'," & -1 & _
"," & -1 & _
")"

            InvLineCommand = New OleDb.OleDbCommand(SQL4NewInvoiceLine, DBConnection)
            Try
                InvLineCommand.ExecuteNonQuery()
            Catch mex As Exception
                MessageBox.Show(mex.Message)
            End Try

            SQL4NewInvoiceLine = "INSERT INTO InvoiceLine(InvNo,[Line#],TypeCode,Description,Qty,TonsorKg,CostPerUnit,Total) VALUES " & _
"(" & InvoiceNumber & _
"," & 6 & _
",'" & "N/A" & _
"','" & toRand(vatAmt, True) & _
"'," & -1 & _
",'" & "N/A" & _
"'," & -1 & _
"," & -1 & _
")"

            InvLineCommand = New OleDb.OleDbCommand(SQL4NewInvoiceLine, DBConnection)
            Try
                InvLineCommand.ExecuteNonQuery()
            Catch mex As Exception
                MessageBox.Show(mex.Message)
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
            cmbJobNum.SelectedIndex = -1
        Catch Myerror As Exception
            MessageBox.Show(Myerror.Message)
        Finally
            DBConnection.Close()
        End Try

    End Sub

    Private Function toRand(ByVal input As String, ByVal r As Boolean) As String

        Dim iput As Double

        Try
            iput = Double.Parse(input)
            If r Then
                Return Format(iput, "R #,###,##0.00")
            Else
                Return Format(iput, "#,###,##0.00")
            End If

        Catch ex As Exception
            Return String.Empty
            MessageBox.Show("Error with input string.", "Cannot convert to Rand format.", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
        End Try

    End Function

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim Form As PrintCutInv = New PrintCutInv(Me)
        Form.populate_invoiceNumbers()
        Form.txt_InvNumToPrint.SelectedIndex = Form.txt_InvNumToPrint.Items.IndexOf(InvoiceNumber.ToString)
        Form.btn_PrintInv_Click(sender, e)
    End Sub


    Private Sub cmbJobNum_SelectedIndexChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbJobNum.SelectedIndexChanged
        cmbJobName.SelectedIndex = cmbJobNum.SelectedIndex
    End Sub


    Private Sub cmbJobName_SelectedIndexChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbJobName.SelectedIndexChanged
        cmbJobNum.SelectedIndex = cmbJobName.SelectedIndex
    End Sub
End Class
