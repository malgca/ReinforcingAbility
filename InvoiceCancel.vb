Public Class InvoiceCancel
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
    Friend WithEvents txtHeading As System.Windows.Forms.TextBox
    Friend WithEvents dtpInvDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txt_InvNumToPrint As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btn_PrintInv As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents lblType As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtJobNum As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.txtHeading = New System.Windows.Forms.TextBox
        Me.dtpInvDate = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtJobNum = New System.Windows.Forms.TextBox
        Me.txt_InvNumToPrint = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.btn_PrintInv = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.lblType = New System.Windows.Forms.Label
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtHeading)
        Me.GroupBox1.Controls.Add(Me.dtpInvDate)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.txtJobNum)
        Me.GroupBox1.Location = New System.Drawing.Point(24, 88)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(320, 112)
        Me.GroupBox1.TabIndex = 11
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Invoice Details"
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
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(16, 48)
        Me.Label3.Name = "Label3"
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "Invoice Heading"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(16, 72)
        Me.Label5.Name = "Label5"
        Me.Label5.TabIndex = 0
        Me.Label5.Text = "Job Number"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtJobNum
        '
        Me.txtJobNum.Location = New System.Drawing.Point(136, 72)
        Me.txtJobNum.Name = "txtJobNum"
        Me.txtJobNum.Size = New System.Drawing.Size(168, 20)
        Me.txtJobNum.TabIndex = 2
        Me.txtJobNum.Text = ""
        '
        'txt_InvNumToPrint
        '
        Me.txt_InvNumToPrint.Location = New System.Drawing.Point(128, 24)
        Me.txt_InvNumToPrint.Name = "txt_InvNumToPrint"
        Me.txt_InvNumToPrint.Size = New System.Drawing.Size(121, 21)
        Me.txt_InvNumToPrint.TabIndex = 10
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(24, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Invoice Number:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btn_PrintInv
        '
        Me.btn_PrintInv.Location = New System.Drawing.Point(104, 216)
        Me.btn_PrintInv.Name = "btn_PrintInv"
        Me.btn_PrintInv.Size = New System.Drawing.Size(112, 23)
        Me.btn_PrintInv.TabIndex = 5
        Me.btn_PrintInv.Text = "Cancel Invoice"
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(232, 216)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(104, 23)
        Me.btnClose.TabIndex = 6
        Me.btnClose.Text = "Close"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(24, 56)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(72, 23)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "Invoice Type:"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblType
        '
        Me.lblType.Location = New System.Drawing.Point(128, 56)
        Me.lblType.Name = "lblType"
        Me.lblType.Size = New System.Drawing.Size(120, 23)
        Me.lblType.TabIndex = 7
        Me.lblType.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'InvoiceCancel
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(362, 264)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.txt_InvNumToPrint)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btn_PrintInv)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.lblType)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "InvoiceCancel"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Invoice Cancellation"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

   

    Private Sub InvoiceCancel_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        populate_invoiceNumbers()
    End Sub

    Dim ty, dt, hd, jn As ArrayList

    Private Sub delCutSheet(ByVal inInvNo As String)

        Dim command As OleDb.OleDbCommand = New OleDb.OleDbCommand
        Dim OleDbCmdObj1 As OleDb.OleDbCommand = New OleDb.OleDbCommand
        Dim OleDbCmdObj2 As OleDb.OleDbCommand = New OleDb.OleDbCommand
        Dim OleDbCmdObj3 As OleDb.OleDbCommand = New OleDb.OleDbCommand
        Dim rowCnt, i As Int16
        Dim cutNo As String = String.Empty

        dbconnection.Open()
        ' GET ALL THE CUTTING SHEETS FOR THAT INVOICE

        Dim sqlCut As String = "SELECT CuttingSheet.CutSheetNo" & _
        " FROM CuttingSheet " & _
        "WHERE InvoiceNo = " & inInvNo

        Dim DSCut As Data.DataSet = New Data.DataSet
        Dim adapter As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(sqlCut, dbconnection)

        adapter.Fill(DSCut)

        rowCnt = DSCut.tables(0).rows.count

        ' /* FOR EACH CUTTING SHEET*/
        For i = 0 To rowCnt - 1

            cutNo = DSCut.Tables(0).rows(i).item("CutSheetNo").ToString()
            'DELETE CUTTING SHEET ITEMS
            Dim sqlDel1 As String = "DELETE * from CutItem " & _
                          "WHERE CutItem.CutSheetNo = " & cutNo
            Try
                Dim MyDataAdapter2 As New System.Data.OleDb.OleDbDataAdapter
                MyDataAdapter2.DeleteCommand = OleDbCmdObj2
                MyDataAdapter2.DeleteCommand.CommandText = sqlDel1
                MyDataAdapter2.DeleteCommand.Connection = dbconnection

                MyDataAdapter2.DeleteCommand.ExecuteNonQuery()

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            'DELETE CUTTING SHEET SCHEDULES
            Dim sqlDel2 As String = "DELETE * FROM SchedItem " & _
                     "WHERE SchedItem.CutSheetNo = " & cutNo
            Try
                Dim MyDataAdapter3 As New System.Data.OleDb.OleDbDataAdapter
                MyDataAdapter3.DeleteCommand = OleDbCmdObj3
                MyDataAdapter3.DeleteCommand.CommandText = sqlDel2
                MyDataAdapter3.DeleteCommand.Connection = dbconnection

                MyDataAdapter3.DeleteCommand.ExecuteNonQuery()

            Catch ex As Exception
                MsgBox(ex.Message)

            End Try
        Next i

        Dim sqlDel3 As String = "DELETE * from CuttingSheet " & _
                                 "WHERE CuttingSheet.CutSheetNo = " & cutNo
        Try
            Dim MyDataAdapter2 As New System.Data.OleDb.OleDbDataAdapter
            MyDataAdapter2.DeleteCommand = OleDbCmdObj2
            MyDataAdapter2.DeleteCommand.CommandText = sqlDel3
            MyDataAdapter2.DeleteCommand.Connection = dbconnection

            MyDataAdapter2.DeleteCommand.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        MessageBox.Show("Successfully deleted Cutting Sheet records")
        dbconnection.Close()
    End Sub

    Private Sub txt_InvNumToPrint_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_InvNumToPrint.SelectedIndexChanged
        dtpInvDate.Value = dt(txt_InvNumToPrint.SelectedIndex)
        txtHeading.Text = hd(txt_InvNumToPrint.SelectedIndex)
        lblType.Text = ty(txt_InvNumToPrint.SelectedIndex)
        txtJobNum.Text = jn(txt_InvNumToPrint.SelectedIndex)
    End Sub

    Public Sub populate_invoiceNumbers()
        txt_InvNumToPrint.Items.Clear()
        Dim sql As String = "SELECT InvoiceNo,InvoiceType,InvDate,InvoiceHeading,InvJobNo FROM Invoice ORDER BY InvoiceNo"
        Dim ds As New Data.DataSet
        Dim ad As New OleDb.OleDbDataAdapter(sql, dbconnection)
        ad.Fill(ds)

        ty = New ArrayList
        dt = New ArrayList
        hd = New ArrayList
        jn = New ArrayList

        Dim f As Integer
        For f = 0 To ds.Tables(0).Rows.Count - 1
            txt_InvNumToPrint.Items.Add(ds.Tables(0).Rows(f).Item("InvoiceNo").ToString())
            ty.Add(ds.Tables(0).Rows(f).Item("InvoiceType").ToString())
            dt.Add(ds.Tables(0).Rows(f).Item("InvDate"))
            hd.Add(ds.Tables(0).Rows(f).Item("InvoiceHeading").ToString())
            jn.Add(ds.Tables(0).Rows(f).Item("InvJobNo").ToString())
        Next f

    End Sub

    Private Sub InvoiceCancel_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If Not IsNothing(caller) Then
            caller.Show()
        End If

        caller = Nothing
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btn_PrintInv_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_PrintInv.Click
        If MessageBox.Show("Are you sure you want to delete invoice " & txt_InvNumToPrint.Text & " ? ", "WARNING!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.Yes Then
            deleteInvoice(txt_InvNumToPrint.Text, lblType.Text)
            populate_invoiceNumbers()
            lblType.Text = ""
            txtHeading.Clear()
            txtJobNum.Clear()

        End If
    End Sub

    Private Sub deleteInvoice(ByVal num As String, ByVal type As String)
        Dim command As OleDb.OleDbCommand = New OleDb.OleDbCommand
        Dim sql As String = "DELETE * FROM InvoiceLine WHERE InvNo = " & num

        command.CommandText = sql
        command.Connection = dbconnection

        Try
            dbconnection.Open()
            command.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            dbconnection.Close()
        End Try

        sql = "DELETE * FROM Invoice WHERE InvoiceNo = " & num

        command.CommandText = sql
        Try
            dbconnection.Open()
            command.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            dbconnection.Close()
        End Try

        If type = "Cutting Sheet" Then
            If MessageBox.Show("Do you want to delete the cutting sheet for invoice " & num & " ? ", "WARNING!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.Yes Then
                delCutSheet(num)
            Else
                sql = "UPDATE CuttingSheet SET InvoiceNo = 0 WHERE InvoiceNo = " & num
                command.CommandText = sql
                Try
                    dbconnection.Open()
                    command.ExecuteNonQuery()

                Catch ex As Exception
                    MsgBox(ex.Message)
                Finally
                    dbconnection.Close()
                End Try
            End If
        End If
    End Sub


    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox1.Enter

    End Sub
End Class
