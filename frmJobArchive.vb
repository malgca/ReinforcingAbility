Public Class frmJobArchive
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Dim dbconnection As OleDb.OleDbConnection
    Dim caller As Object

    Public Sub New(ByRef callingForm As Object, ByRef dbcon As OleDb.OleDbConnection)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        dbconnection = dbcon
        caller = callingForm
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
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents lblType As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents btn_Archive As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents txtOrderNum As System.Windows.Forms.TextBox
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents cbxTonsKilograms As System.Windows.Forms.ComboBox
    Friend WithEvents cmbJobs As System.Windows.Forms.ComboBox
    Friend WithEvents cmbJobName As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cmb_contName As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cmbCont As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents cmbJobNum As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.cmbJobNum = New System.Windows.Forms.ComboBox
        Me.cmbCont = New System.Windows.Forms.ComboBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.cmb_contName = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.cmbJobName = New System.Windows.Forms.ComboBox
        Me.cmbJobs = New System.Windows.Forms.ComboBox
        Me.cbxTonsKilograms = New System.Windows.Forms.ComboBox
        Me.txtOrderNum = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.btn_Archive = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.lblType = New System.Windows.Forms.Label
        Me.btnAdd = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.cmbJobNum)
        Me.GroupBox1.Controls.Add(Me.cmbCont)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.cmb_contName)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.cmbJobName)
        Me.GroupBox1.Controls.Add(Me.cmbJobs)
        Me.GroupBox1.Controls.Add(Me.cbxTonsKilograms)
        Me.GroupBox1.Controls.Add(Me.txtOrderNum)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Location = New System.Drawing.Point(24, 32)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(736, 256)
        Me.GroupBox1.TabIndex = 11
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Job Details"
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(16, 24)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(88, 23)
        Me.Label7.TabIndex = 23
        Me.Label7.Text = "Job Number"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cmbJobNum
        '
        Me.cmbJobNum.ItemHeight = 13
        Me.cmbJobNum.Location = New System.Drawing.Point(136, 24)
        Me.cmbJobNum.Name = "cmbJobNum"
        Me.cmbJobNum.Size = New System.Drawing.Size(80, 21)
        Me.cmbJobNum.TabIndex = 22
        '
        'cmbCont
        '
        Me.cmbCont.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbCont.Enabled = False
        Me.cmbCont.ItemHeight = 13
        Me.cmbCont.Location = New System.Drawing.Point(624, 144)
        Me.cmbCont.Name = "cmbCont"
        Me.cmbCont.Size = New System.Drawing.Size(56, 21)
        Me.cmbCont.TabIndex = 21
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(528, 144)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(88, 23)
        Me.Label6.TabIndex = 20
        Me.Label6.Text = "Contractor Num"
        '
        'cmb_contName
        '
        Me.cmb_contName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmb_contName.ItemHeight = 13
        Me.cmb_contName.Location = New System.Drawing.Point(136, 144)
        Me.cmb_contName.Name = "cmb_contName"
        Me.cmb_contName.Size = New System.Drawing.Size(376, 21)
        Me.cmb_contName.TabIndex = 18
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(16, 64)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(64, 23)
        Me.Label4.TabIndex = 17
        Me.Label4.Text = "Job Name"
        '
        'cmbJobName
        '
        Me.cmbJobName.ItemHeight = 13
        Me.cmbJobName.Location = New System.Drawing.Point(136, 64)
        Me.cmbJobName.Name = "cmbJobName"
        Me.cmbJobName.Size = New System.Drawing.Size(376, 21)
        Me.cmbJobName.TabIndex = 16
        '
        'cmbJobs
        '
        Me.cmbJobs.ItemHeight = 13
        Me.cmbJobs.Location = New System.Drawing.Point(496, 24)
        Me.cmbJobs.Name = "cmbJobs"
        Me.cmbJobs.Size = New System.Drawing.Size(192, 21)
        Me.cmbJobs.TabIndex = 14
        '
        'cbxTonsKilograms
        '
        Me.cbxTonsKilograms.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxTonsKilograms.Location = New System.Drawing.Point(136, 104)
        Me.cbxTonsKilograms.Name = "cbxTonsKilograms"
        Me.cbxTonsKilograms.Size = New System.Drawing.Size(56, 21)
        Me.cbxTonsKilograms.TabIndex = 12
        '
        'txtOrderNum
        '
        Me.txtOrderNum.Location = New System.Drawing.Point(136, 184)
        Me.txtOrderNum.Name = "txtOrderNum"
        Me.txtOrderNum.TabIndex = 10
        Me.txtOrderNum.Text = ""
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(16, 184)
        Me.Label3.Name = "Label3"
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Order Num"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 144)
        Me.Label2.Name = "Label2"
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Contractor Name"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 104)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(96, 24)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Tons or Kilograms"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(392, 24)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(88, 23)
        Me.Label5.TabIndex = 0
        Me.Label5.Text = "Job Search"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btn_Archive
        '
        Me.btn_Archive.Location = New System.Drawing.Point(352, 320)
        Me.btn_Archive.Name = "btn_Archive"
        Me.btn_Archive.Size = New System.Drawing.Size(88, 23)
        Me.btn_Archive.TabIndex = 5
        Me.btn_Archive.Text = "Archive "
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(488, 320)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(80, 23)
        Me.btnClose.TabIndex = 6
        Me.btnClose.Text = "Close"
        '
        'lblType
        '
        Me.lblType.Location = New System.Drawing.Point(128, 56)
        Me.lblType.Name = "lblType"
        Me.lblType.Size = New System.Drawing.Size(120, 23)
        Me.lblType.TabIndex = 7
        Me.lblType.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnAdd
        '
        Me.btnAdd.Location = New System.Drawing.Point(48, 320)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.TabIndex = 12
        Me.btnAdd.Text = "Add"
        '
        'btnEdit
        '
        Me.btnEdit.Location = New System.Drawing.Point(144, 320)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.TabIndex = 13
        Me.btnEdit.Text = "Edit"
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(248, 320)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.TabIndex = 14
        Me.btnSave.Text = "Save"
        '
        'frmJobArchive
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(794, 400)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnAdd)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btn_Archive)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.lblType)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "frmJobArchive"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Job Maintenance"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim ty, dt, hd, jn As ArrayList


    Private state As String
    Private Sub JobCancel_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If Not IsNothing(caller) Then
            caller.Show()
        End If

        caller = Nothing
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        state = ""
        Me.Close()
    End Sub

    Private Sub btn_Archive_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Archive.Click
        ' Dim jobNum As String = cmbJobs.Text -- Equivalent to BS's Code
        Dim jobNum As String = cmbJobNum.Text ' -- Equivalent to MG's Code

        If jobNum = "" Then
            MsgBox("A job must be selected first ", MsgBoxStyle.Critical, "Error")
            cmbJobName.Focus()
        Else
            If MessageBox.Show("Are you sure you want to delete job " & jobNum & " ? ", "WARNING!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.Yes Then
                copyJob(jobNum)
                copyJobCuts(jobNum)
                copyJobInv(jobNum)
                deleteInvoice(jobNum)
                delJobCuts(jobNum)
                deleteJob(jobNum)
                populate_cmb_jobs()
                populate_cmb_jobNames()
                lblType.Text = ""
                MessageBox.Show("Job successfully archived")
            End If
        End If
    End Sub

    Private Sub copyJobCuts(ByVal inJob As String)
        Dim command As OleDb.OleDbCommand = New OleDb.OleDbCommand
        Dim OleDbCmdObj1 As OleDb.OleDbCommand = New OleDb.OleDbCommand
        Dim OleDbCmdObj2 As OleDb.OleDbCommand = New OleDb.OleDbCommand
        Dim OleDbCmdObj3 As OleDb.OleDbCommand = New OleDb.OleDbCommand
        Dim rowCnt, i As Int16
        Dim cutNo As String

        dbconnection.Open()

        Dim sqlIns As String = "INSERT INTO ACuttingSheet " & _
            "SELECT CuttingSheet.CutsheetNo, invoiceNo, details, CSHeading, [Job No], CutDate " & _
            "FROM CuttingSheet WHERE cuttingSheet.[job No] = '" & inJob & "'"
        Try
            Dim MyDataAdapter As New System.Data.OleDb.OleDbDataAdapter
            MyDataAdapter.InsertCommand = OleDbCmdObj1
            MyDataAdapter.InsertCommand.CommandText = sqlIns
            MyDataAdapter.InsertCommand.Connection = dbconnection

            MyDataAdapter.InsertCommand.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try

        ' GET ALL THE CUTTING SHEETS FOR THAT JOB
        Dim sql4 As String = "SELECT CuttingSheet.CutSheetNo" & _
        " FROM CuttingSheet " & _
        "WHERE [Job No] = '" & inJob & "'"

        Dim DS4SchNo As Data.DataSet = New Data.DataSet
        Dim adapter As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(sql4, dbconnection)

        adapter.Fill(DS4SchNo)

        rowCnt = DS4SchNo.tables(0).rows.count
    
        ' /* FOR EACH CUTTING SHEET*/
        For i = 0 To rowCnt - 1

            cutNo = DS4SchNo.Tables(0).rows(i).item("CutSheetNo").ToString()
            'INSERT CUTTING SHEET SCHEDULES
            Dim sqlIns2 As String = "INSERT INTO ASchedItem " & _
                    "SELECT SchedItem.CutsheetNo, ScheduleNo FROM SchedItem " & _
            "WHERE SchedItem.CutSheetNo = " & cutNo
            Try
                Dim MyDataAdapter2 As New System.Data.OleDb.OleDbDataAdapter
                MyDataAdapter2.InsertCommand = OleDbCmdObj2
                MyDataAdapter2.InsertCommand.CommandText = sqlIns2
                MyDataAdapter2.InsertCommand.Connection = dbconnection

                MyDataAdapter2.InsertCommand.ExecuteNonQuery()

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            'INSERT CUTTING SHEET ITEMS
            Dim sqlIns3 As String = "INSERT INTO ACutItem " & _
                    "SELECT CutItem.CutsheetNo, ScheduleNo, Item, TypeCode, Length, Qty FROM CutItem " & _
            "WHERE CutItem.CutSheetNo = " & cutNo
            Try
                Dim MyDataAdapter3 As New System.Data.OleDb.OleDbDataAdapter
                MyDataAdapter3.InsertCommand = OleDbCmdObj3
                MyDataAdapter3.InsertCommand.CommandText = sqlIns3
                MyDataAdapter3.InsertCommand.Connection = dbconnection

                MyDataAdapter3.InsertCommand.ExecuteNonQuery()

            Catch ex As Exception
                MsgBox(ex.Message)

            End Try
        Next i
        MessageBox.Show("Copied Cutting Sheet records")
        dbconnection.Close()
    End Sub

    Private Sub copyJobInv(ByVal inJob As String)
        Dim command As OleDb.OleDbCommand = New OleDb.OleDbCommand
        Dim OleDbCmdObj1 As OleDb.OleDbCommand = New OleDb.OleDbCommand
        Dim OleDbCmdObj2 As OleDb.OleDbCommand = New OleDb.OleDbCommand
        Dim OleDbCmdObj3 As OleDb.OleDbCommand = New OleDb.OleDbCommand
        Dim rowCnt, i As Int16
        Dim invNo As String

        dbconnection.Open()

        Dim sqlIns As String = "INSERT INTO AInvoice " & _
            "SELECT Invoice.invoiceNo, invoiceType, invDate, invoiceHeading, invRefNum, invDeliveryNoteNo, invFactor, invMonthAndYear, invOrdNum , invTotal, invVatAmt, invNett, invJobNo " & _
            "FROM Invoice WHERE Invoice.invJobNo = '" & inJob & "'"
        Try
            Dim MyDataAdapter As New System.Data.OleDb.OleDbDataAdapter
            MyDataAdapter.InsertCommand = OleDbCmdObj1
            MyDataAdapter.InsertCommand.CommandText = sqlIns
            MyDataAdapter.InsertCommand.Connection = dbconnection

            MyDataAdapter.InsertCommand.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try

        ' GET ALL THE INVOICE LINES FOR THAT INVOICE
        Dim sqlInv As String = "SELECT Invoice.InvoiceNo" & _
        " FROM Invoice " & _
        "WHERE invJobNo = '" & inJob & "'"

        Dim DSInv As Data.DataSet = New Data.DataSet
        Dim adapter As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(sqlInv, dbconnection)

        adapter.Fill(DSInv)

        rowCnt = DSInv.tables(0).rows.count

        ' /* FOR EACH INVOICE*/
        For i = 0 To rowCnt - 1

            invNo = DSInv.Tables(0).rows(i).item("invoiceNo").ToString()

            'INSERT INVOICE LINES
            Dim sqlIns2 As String = "INSERT INTO AInvoiceLine " & _
                    "SELECT InvoiceLine.InvNo, [line#], TypeCode, Description, Qty, TonsOrKg, costPerUnit, total FROM InvoiceLine " & _
            "WHERE InvoiceLine.InvNo = " & invNo
            Try
                Dim MyDataAdapter2 As New System.Data.OleDb.OleDbDataAdapter
                MyDataAdapter2.InsertCommand = OleDbCmdObj2
                MyDataAdapter2.InsertCommand.CommandText = sqlIns2
                MyDataAdapter2.InsertCommand.Connection = dbconnection

                MyDataAdapter2.InsertCommand.ExecuteNonQuery()

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            ' NOW DELETE ALL INVOICE LINES FOR THAT INVOICE
            Dim sqlDel As String = "DELETE * FROM InvoiceLine WHERE InvNo = " & invNo

            Try
                command.CommandText = sqlDel
                command.Connection = dbconnection

                command.ExecuteNonQuery()

            Catch ex As Exception
                MsgBox(ex.Message)
            Finally

            End Try
        Next i
        dbconnection.Close()
    End Sub

    Private Sub copyJob(ByVal inJob As String)
        Dim command As OleDb.OleDbCommand = New OleDb.OleDbCommand
        Dim OleDbCmdObj1 As OleDb.OleDbCommand = New OleDb.OleDbCommand
        Dim OleDbCmdObj2 As OleDb.OleDbCommand = New OleDb.OleDbCommand
        Dim OleDbCmdObj3 As OleDb.OleDbCommand = New OleDb.OleDbCommand

        dbconnection.Open()

        Dim sqlIns As String = "INSERT INTO AJob " & _
            "SELECT Job.jobNo, companyNo, jobName, contractorNo " & _
            "FROM Job WHERE Job.JobNo = '" & inJob & "'"
        Try
            Dim MyDataAdapter As New System.Data.OleDb.OleDbDataAdapter
            MyDataAdapter.InsertCommand = OleDbCmdObj1
            MyDataAdapter.InsertCommand.CommandText = sqlIns
            MyDataAdapter.InsertCommand.Connection = dbconnection
            MyDataAdapter.InsertCommand.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        ' ADD JOB RATES
        Dim sqlIns2 As String = "INSERT INTO AJobRate " & _
                "SELECT JobRate.jobNo,TypeCode, rate FROM JobRate " & _
         "WHERE JobRate.JobNo = '" & inJob & "'"

        Try
            Dim MyDataAdapter2 As New System.Data.OleDb.OleDbDataAdapter
            MyDataAdapter2.InsertCommand = OleDbCmdObj2
            MyDataAdapter2.InsertCommand.CommandText = sqlIns2
            MyDataAdapter2.InsertCommand.Connection = dbconnection

            MyDataAdapter2.InsertCommand.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try

        dbconnection.Close()
    End Sub


       
    Private Sub delJobCuts(ByVal inJob As String)

        Dim command As OleDb.OleDbCommand = New OleDb.OleDbCommand
        Dim rowCnt, i As Int16
        Dim cutNo As String

        dbconnection.Open()

        ' GET ALL THE CUTTING SHEETS FOR THAT JOB
        Dim sqlCut As String = "SELECT CuttingSheet.cutSheetNo" & _
        " FROM CuttingSheet " & _
        "WHERE [Job No] = '" & inJob & "'"

        Dim DSCut As Data.DataSet = New Data.DataSet
        Dim adapter As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(sqlCut, dbconnection)

        adapter.Fill(DSCut)

        rowCnt = DSCut.tables(0).rows.count

        For i = 0 To rowCnt - 1
            cutNo = DSCut.Tables(0).rows(i).item("CutSheetNo").ToString()
            deleteCutItems(cutNo)
            DelCutSched(cutNo)
        Next i
        DeleteCut(inJob)

        dbconnection.Close()
    End Sub

    Private Sub deleteJob(ByVal inJob As String)
        Dim command As OleDb.OleDbCommand = New OleDb.OleDbCommand
        dbconnection.Open()

        ' DELETE ALL JOB RATES
        Dim sqlDel As String = "DELETE FROM JobRate " & _
                "WHERE JobNo = '" & inJob & "'"
        Try
            command.CommandText = sqlDel
            command.Connection = dbconnection
            command.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        ' DELETE ALL JOB RATES
        Dim sqlDel2 As String = "DELETE FROM Job " & _
                "WHERE JobNo = '" & inJob & "'"
        Try
            command.CommandText = sqlDel2
            command.Connection = dbconnection
            command.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        MessageBox.Show("Job & rates deleted  ")
        dbconnection.Close()
    End Sub
    Private Sub DeleteCut(ByVal inJob As String)
        Dim command As OleDb.OleDbCommand = New OleDb.OleDbCommand
        ' DELETE ALL CUTTING SHEETS
        Dim sqlDel2 As String = "DELETE FROM CuttingSheet " & _
                "WHERE [Job No] = '" & inJob & "'"
        Try
            command.CommandText = sqlDel2
            command.Connection = dbconnection
            command.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    Private Sub deleteCutItems(ByVal inCut As String)
        ' NOW DELETE ALL CUTTING SHEET ITEMS FOR THAT CUTTING SHEET
        Dim sqlDel1 As String = "DELETE FROM CutItem WHERE CutSheetNo = " & inCut
        Dim command As OleDb.OleDbCommand = New OleDb.OleDbCommand

        Try
            command.CommandText = sqlDel1
            command.Connection = dbconnection
            command.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub DelCutSched(ByVal inCut As String)
        ' DELETE ALL SCHEDULES FOR CUTTING SHEET
        Dim command As OleDb.OleDbCommand = New OleDb.OleDbCommand
        Dim sqlDel As String = "DELETE * FROM SchedItem WHERE CutSheetNo = " & inCut
        Try
            command.CommandText = sqlDel
            command.Connection = dbconnection
            command.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub deleteInvoice(ByVal inJob As String)
        Dim rowCnt As Integer
        Dim command As OleDb.OleDbCommand = New OleDb.OleDbCommand
        Dim sqlDel As String = "DELETE * FROM Invoice WHERE InvJobNo = '" & inJob & "'"
        dbconnection.Open()
        Try
            command.CommandText = sqlDel
            command.Connection = dbconnection
            rowCnt = command.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
        dbconnection.Close()
        MessageBox.Show("Invoices deleted: " + rowCnt.ToString)
    End Sub

    Private Sub frmJobArchive_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        state = ""
        populate_cmb_cont()
        load_tonsKg()
        ' initialise state to be add mode
    End Sub

    Private Sub load_tonsKg()
        cbxTonsKilograms.Items.Clear()
        cbxTonsKilograms.Items.Add("T")
        cbxTonsKilograms.Items.Add("Kg")
        cbxTonsKilograms.SelectedIndex = 0
    End Sub
    Private Sub clear_jobs()
        cmbJobs.Items.Clear()
        cmbJobName.Items.Clear()
        cmbJobNum.Items.Clear()
    End Sub

    Private Sub populate_cmb_jobs()
        Dim jobNum, jobName, jobs As String

        cmbJobs.Items.Clear()
        cmbJobs.Items.Add("---Select a Job---")
        cmbJobNum.Items.Clear()
        cmbJobNum.Items.Add("----Select a Job ----")
        Dim sql As String = "SELECT JobNo, jobname FROM Job ORDER BY JobName"
        Dim dataset As New Data.DataSet
        Dim adapter As New OleDb.OleDbDataAdapter(sql, dbconnection)
        adapter.Fill(dataset)

        Dim aunty As Integer
        For aunty = 0 To dataset.Tables(0).Rows.Count - 1
            jobName = dataset.Tables(0).Rows(aunty).Item("jobname").ToString()
            jobNum = dataset.Tables(0).Rows(aunty).Item("JobNo").ToString()
            jobs = jobNum + " - " + jobName
            cmbJobs.Items.Add(jobs)
            cmbJobNum.Items.Add(jobNum)
        Next aunty
    End Sub

    Private Sub populate_cmb_jobNames()

        Dim jobName As String
        cmbJobName.Items.Clear()
        cmbJobName.Items.Add("---Select a Job Name ---")
        Dim sql As String = "SELECT jobname, jobNo FROM Job ORDER BY JobName"
        Dim dataset As New Data.DataSet
        Dim adapter As New OleDb.OleDbDataAdapter(sql, dbconnection)
        adapter.Fill(dataset)

        Dim aunty As Integer
        For aunty = 0 To dataset.Tables(0).Rows.Count - 1
            jobName = dataset.Tables(0).Rows(aunty).Item("JobName").ToString()
            cmbJobName.Items.Add(jobName)
        Next aunty
    End Sub

    Private Sub populate_cmb_cont()
        Dim i As Integer
        Dim cnt As Double
        Dim contNum, contName As String

        cmb_contName.Items.Clear()
        cmb_contName.Items.Add("---Select a Contractor ---")
        Dim sql2 As String = "SELECT contractorNo, contractorName FROM contractor ORDER BY contractorName"
        Dim dsCont2 As New Data.DataSet
        Dim adapter2 As New OleDb.OleDbDataAdapter(sql2, dbconnection)
        adapter2.Fill(dsCont2)

        Try
            cnt = dsCont2.Tables(0).Rows.Count
            For i = 0 To cnt - 1
                contName = dsCont2.Tables(0).Rows(i).Item("ContractorName").ToString()
                cmb_contName.Items.Add(contName)
            Next i
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        cmbCont.Items.Clear()
        cmbCont.Items.Add("---Select a Contractor Num---")
        Dim sqlName As String = "SELECT contractorNo FROM contractor ORDER BY contractorName"
        Dim dsContName As New Data.DataSet
        Dim adapterName As New OleDb.OleDbDataAdapter(sqlName, dbconnection)
        adapterName.Fill(dsContName)

        Try
            cnt = dsContName.Tables(0).Rows.Count
            For i = 0 To cnt - 1
                contNum = dsCont2.Tables(0).Rows(i).Item("ContractorNo").ToString()
                cmbCont.Items.Add(contNum)
            Next i
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub


    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        cmbJobs.Enabled = False
        cmbJobNum.Focus()
        state = "add"
        clearAll()
        clear_jobs()
        populate_cmb_cont()
    End Sub

    Private Sub clearAll()
        cmbJobNum.Text = ""
        cmbJobName.Text = ""
        cmbJobs.Enabled = True
        txtOrderNum.Clear()
        cmbJobName.Text = ""
        cmbJobs.Text = ""
        cmb_contName.Text = ""
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim command As OleDb.OleDbCommand = New OleDb.OleDbCommand
        Dim dsJob As Data.DataSet = New Data.DataSet
        Dim discount As Integer
        Dim addDisc, design As Double
        Dim jobError As Boolean
        Dim compNo, unitMeas, jobNum, jobName, contNo, tonsOrKg, orderNum As String
        discount = 0
        addDisc = 0
        design = 0
        'assign company = 1 for all jobs
        compNo = "1"
        unitMeas = " "
        jobNum = cmbJobNum.Text
        jobName = cmbJobName.Text
        'contNo = cmbCont.Text
        tonsOrKg = cbxTonsKilograms.Text
        orderNum = txtOrderNum.Text
        contNo = cmbCont.Text
        jobError = False
        If cmbJobName.Text = "" Then
            MsgBox("A Job Name is required", MsgBoxStyle.Critical, "Error")
            jobError = True
            cmbJobName.Focus()
        Else
            If cmb_contName.Text = "" Then
                MsgBox("A Contractor Number is required", MsgBoxStyle.Critical, "Error")
                cmb_contName.Focus()
                jobError = True
            End If
        End If

        If Not jobError Then
            If state = "add" Then

                Dim sqlNewJob As String = "INSERT INTO Job VALUES (?,?,?,?,?,?,?,?,?,?)"

                Dim daJob As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter
                dbconnection.Open()
                Try
                    daJob.InsertCommand = command
                    daJob.InsertCommand.CommandText = sqlNewJob
                    ' assign values to all the fields
                    daJob.InsertCommand.Parameters.AddWithValue("jobNo", jobNum)
                    daJob.InsertCommand.Parameters.AddWithValue("companyNo", compNo)
                    daJob.InsertCommand.Parameters.AddWithValue("jobName", jobName)
                    daJob.InsertCommand.Parameters.AddWithValue("discount", discount)
                    daJob.InsertCommand.Parameters.AddWithValue("addedDiscount", addDisc)
                    daJob.InsertCommand.Parameters.AddWithValue("design", design)
                    daJob.InsertCommand.Parameters.AddWithValue("[Tons or Kilograms]", tonsOrKg)
                    daJob.InsertCommand.Parameters.AddWithValue("contractorNo", contNo)
                    daJob.InsertCommand.Parameters.AddWithValue("orderNo", orderNum)
                    daJob.InsertCommand.Parameters.AddWithValue("unitOfMeas", unitMeas)

                    daJob.InsertCommand.Connection = dbconnection
                    daJob.InsertCommand.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
                ' GET ALL THE RATES FOR THAT JOB
                PopulateJobRate(jobNum)
                dbconnection.Close()
            Else
                Dim sqlChgJob As String

                sqlChgJob = "UPDATE Job SET jobName = '" & jobName & "'" & _
                    ", contractorNo = '" & contNo & "'" & _
                   ", orderNo = '" & orderNum & "'" & _
                   ",[tons or Kilograms]= '" & tonsOrKg & "'" & _
                   " WHERE Job.jobNo = '" & jobNum & "'"

                Dim daJob As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter
                dbconnection.Open()
                Try
                    daJob.UpdateCommand = command
                    daJob.UpdateCommand.CommandText = sqlChgJob
                    daJob.UpdateCommand.Connection = dbconnection
                    daJob.UpdateCommand.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
                dbconnection.Close()
            End If
            MsgBox("Record was successfully saved", MsgBoxStyle.Information, "Information")
        End If
    End Sub

    Private Sub GetCont(ByVal inCont As String)
        Dim rowCnt As Integer

        Dim sql As String = "SELECT Contractor.contractorNo, contractorName" & _
        " FROM Contractor " & _
        " WHERE Contractor.contractorNo = '" & inCont & "'"

        Dim dsCont As Data.DataSet = New Data.DataSet
        Dim adpCont As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(sql, dbconnection)
        Try
            dsCont.Clear()
            adpCont.Fill(dsCont)
            rowCnt = dsCont.tables(0).rows.count
            If rowCnt = 0 Then
                MsgBox("Invalid Contractor Number", MsgBoxStyle.Critical, "Error")
                cmb_contName.Focus()
            Else
                cmb_contName.Text = dsCont.Tables(0).rows(rowCnt - 1).item("contractorName").ToString()
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub GetJob(ByVal inJob As String)
        Dim rowCnt, recId As Integer
        Dim NewJobName, contNum As String

        If inJob = "" And state <> "" Then
            MsgBox("A Job Number is required", MsgBoxStyle.Critical, "Error")
            cmbJobs.Focus()
        Else

            Dim sql As String = "SELECT jobNo, jobName, [Tons or Kilograms], contractorNo, OrderNo " & _
             " FROM Job " & _
             " WHERE Job.jobNo = '" & inJob & "'"

            Dim dsJob As Data.DataSet = New Data.DataSet
            Dim adpJob As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(sql, dbconnection)
            Try
                dsJob.Clear()
                adpJob.Fill(dsJob)
                rowCnt = dsJob.tables(0).rows.count
                If rowCnt = 0 And state = "edit" Then
                    MsgBox("Invalid Job Number", MsgBoxStyle.Critical, "Error")
                    cmbJobs.Focus()
                End If

                If rowCnt <> 0 Then
                    recId = rowCnt - 1
                    If state = "add" Then
                        MsgBox("Job number already exists", MsgBoxStyle.Critical, "Error")
                        'clearAll()
                        cmbJobName.Focus()
                    Else
                        cmbJobName.Focus()

                        NewJobName = dsJob.Tables(0).rows(recId).item("jobName").ToString()
                        If NewJobName <> cmbJobName.Text() Then
                            cmbJobName.Text() = NewJobName
                        End If
                        cbxTonsKilograms.Text = dsJob.Tables(0).rows(recId).item("Tons Or Kilograms").ToString()

                        contNum = dsJob.Tables(0).rows(recId).item("ContractorNo").ToString()
                        cmbCont.Text = contNum
                        txtOrderNum.Text = dsJob.Tables(0).rows(recId).item("OrderNo").ToString()
                    End If
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
    End Sub
    Private Sub PopulateJobRate(ByVal inJob As String)
        Dim rowCnt, i As Integer
        Dim typeCode As String

        ' GET ALL THE RATES FOR THAT JOB
        Dim sql As String = "SELECT ProductType.*" & _
        " FROM ProductType "

        Dim dsProd As Data.DataSet = New Data.DataSet
        Dim dsJobRate As Data.DataSet = New Data.DataSet
        Dim adpProd As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(sql, dbconnection)

        dsProd.Clear()
        adpProd.Fill(dsProd)
        rowCnt = dsProd.tables(0).rows.count
        For i = 0 To rowCnt - 1
            typeCode = dsProd.Tables(0).rows(i).item("typeCode").ToString()
            createRateRec(inJob, typeCode)
        Next i
    End Sub
    Private Sub createRateRec(ByVal inJob As String, ByVal inType As String)
        Dim command As OleDb.OleDbCommand = New OleDb.OleDbCommand
        Dim curRate As Double
        Dim daRate As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter

        Dim sqlIns As String = "INSERT INTO JobRate VALUES (?,?,?)"
        Try
            daRate.InsertCommand = command
            daRate.InsertCommand.CommandText = sqlIns
            ' assign values to all the fields
            daRate.InsertCommand.Parameters.AddWithValue("jobNo", inJob)
            daRate.InsertCommand.Parameters.AddWithValue("typeCode", inType)
            daRate.InsertCommand.Parameters.AddWithValue("rate", curRate)
            daRate.InsertCommand.Connection = dbconnection
            daRate.InsertCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox1.Enter
    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        state = "edit"
        cmbJobs.Enabled = True
        populate_cmb_jobs()
        populate_cmb_jobNames()
    End Sub

    Private Sub cmbJobs_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbJobs.SelectedIndexChanged
        cmbJobName.SelectedIndex = cmbJobs.SelectedIndex
    End Sub

    Private Sub txtCont_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs)
        GetCont(cmbCont.Text)
    End Sub

    Private Sub cmbJobs_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbJobs.Leave

        If state = "add" Then
            GetJob(cmbJobNum.Text)
        End If
    End Sub

    Private Sub cmbJobName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbJobName.SelectedIndexChanged
        cmbJobNum.SelectedIndex = cmbJobName.SelectedIndex
        cmbJobs.SelectedIndex = cmbJobName.SelectedIndex
        If state <> "add" Then
            GetJob(cmbJobNum.Text)
            GetCont(cmbCont.Text)

        End If
    End Sub

    Private Sub cmbCont_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'GetCont(cmbCont.Text)
    End Sub

    Private Sub cmb_contName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb_contName.SelectedIndexChanged
        cmbCont.SelectedIndex = cmb_contName.SelectedIndex
        'contRec = cmb_contName.SelectedIndex.ToString()
        'GetContId(contRec)
    End Sub
    Private Sub GetContId(ByVal inCont As String)

        Dim rowCnt As Integer

        Dim sql As String = "SELECT Contractor.contractorNo, contractorName" & _
        " FROM Contractor " & _
        " WHERE Contractor.contractorNo = '" & inCont & "'"

        Dim dsCont As Data.DataSet = New Data.DataSet
        Dim adpCont As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(sql, dbconnection)
        Try
            dsCont.Clear()
            adpCont.Fill(dsCont)
            rowCnt = dsCont.tables(0).rows.count
            If rowCnt = 0 Then
                MsgBox("Invalid Contractor Number", MsgBoxStyle.Critical, "Error")
                cmb_contName.Focus()
            Else
                cmb_contName.Text = dsCont.Tables(0).rows(rowCnt - 1).item("contractorName").ToString()
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    Private Sub Label4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label4.Click

    End Sub

    Private Sub Label5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label5.Click

    End Sub

    Private Sub cmbJobNum_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbJobNum.SelectedIndexChanged
        cmbJobName.SelectedIndex = cmbJobNum.SelectedIndex
    End Sub
End Class
