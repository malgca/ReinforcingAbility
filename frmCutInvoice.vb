Public Class GenCutSheetInvoice
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
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents txtDelNoteNum As System.Windows.Forms.TextBox
    Friend WithEvents txtOrderNum As System.Windows.Forms.TextBox
    Friend WithEvents txtDesign As System.Windows.Forms.TextBox
    Friend WithEvents txtInvHeading As System.Windows.Forms.TextBox
    Friend WithEvents dtpInvDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnCreateInvoice As System.Windows.Forms.Button
    Friend WithEvents cmb_AllCutSheets As System.Windows.Forms.ComboBox
    Friend WithEvents lblNotify As System.Windows.Forms.Label
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtDelNoteNum = New System.Windows.Forms.TextBox
        Me.txtOrderNum = New System.Windows.Forms.TextBox
        Me.txtDesign = New System.Windows.Forms.TextBox
        Me.txtInvHeading = New System.Windows.Forms.TextBox
        Me.dtpInvDate = New System.Windows.Forms.DateTimePicker
        Me.btnCreateInvoice = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.cmb_AllCutSheets = New System.Windows.Forms.ComboBox
        Me.lblNotify = New System.Windows.Forms.Label
        Me.btnPrint = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 72)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(128, 23)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Delivery Note Number"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 112)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(128, 23)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Order Number"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(16, 152)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(128, 23)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "Design"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(16, 192)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(128, 23)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "Invoice Date"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(16, 232)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(128, 23)
        Me.Label5.TabIndex = 0
        Me.Label5.Text = "Invoice Heading"
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(16, 32)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(128, 23)
        Me.Label7.TabIndex = 0
        Me.Label7.Text = "Cutting Sheet Number "
        '
        'txtDelNoteNum
        '
        Me.txtDelNoteNum.Location = New System.Drawing.Point(152, 72)
        Me.txtDelNoteNum.MaxLength = 50
        Me.txtDelNoteNum.Name = "txtDelNoteNum"
        Me.txtDelNoteNum.Size = New System.Drawing.Size(304, 20)
        Me.txtDelNoteNum.TabIndex = 2
        Me.txtDelNoteNum.Text = ""
        '
        'txtOrderNum
        '
        Me.txtOrderNum.Location = New System.Drawing.Point(152, 112)
        Me.txtOrderNum.Name = "txtOrderNum"
        Me.txtOrderNum.Size = New System.Drawing.Size(240, 20)
        Me.txtOrderNum.TabIndex = 4
        Me.txtOrderNum.Text = ""
        '
        'txtDesign
        '
        Me.txtDesign.Location = New System.Drawing.Point(152, 152)
        Me.txtDesign.Name = "txtDesign"
        Me.txtDesign.Size = New System.Drawing.Size(144, 20)
        Me.txtDesign.TabIndex = 6
        Me.txtDesign.Text = ""
        '
        'txtInvHeading
        '
        Me.txtInvHeading.Location = New System.Drawing.Point(152, 232)
        Me.txtInvHeading.Name = "txtInvHeading"
        Me.txtInvHeading.Size = New System.Drawing.Size(312, 20)
        Me.txtInvHeading.TabIndex = 10
        Me.txtInvHeading.Text = ""
        '
        'dtpInvDate
        '
        Me.dtpInvDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpInvDate.Location = New System.Drawing.Point(152, 192)
        Me.dtpInvDate.Name = "dtpInvDate"
        Me.dtpInvDate.Size = New System.Drawing.Size(144, 20)
        Me.dtpInvDate.TabIndex = 8
        '
        'btnCreateInvoice
        '
        Me.btnCreateInvoice.Location = New System.Drawing.Point(64, 304)
        Me.btnCreateInvoice.Name = "btnCreateInvoice"
        Me.btnCreateInvoice.Size = New System.Drawing.Size(112, 23)
        Me.btnCreateInvoice.TabIndex = 12
        Me.btnCreateInvoice.Text = "Create Invoice"
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(336, 304)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(112, 23)
        Me.btnClose.TabIndex = 14
        Me.btnClose.Text = "Close"
        '
        'cmb_AllCutSheets
        '
        Me.cmb_AllCutSheets.Location = New System.Drawing.Point(152, 32)
        Me.cmb_AllCutSheets.MaxDropDownItems = 15
        Me.cmb_AllCutSheets.Name = "cmb_AllCutSheets"
        Me.cmb_AllCutSheets.Size = New System.Drawing.Size(144, 21)
        Me.cmb_AllCutSheets.TabIndex = 0
        '
        'lblNotify
        '
        Me.lblNotify.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNotify.Location = New System.Drawing.Point(16, 264)
        Me.lblNotify.Name = "lblNotify"
        Me.lblNotify.Size = New System.Drawing.Size(488, 32)
        Me.lblNotify.TabIndex = 15
        Me.lblNotify.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnPrint
        '
        Me.btnPrint.Enabled = False
        Me.btnPrint.Location = New System.Drawing.Point(200, 304)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(112, 23)
        Me.btnPrint.TabIndex = 12
        Me.btnPrint.Text = "Print Preview..."
        '
        'GenCutSheetInvoice
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(538, 351)
        Me.Controls.Add(Me.lblNotify)
        Me.Controls.Add(Me.cmb_AllCutSheets)
        Me.Controls.Add(Me.btnCreateInvoice)
        Me.Controls.Add(Me.dtpInvDate)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.txtDelNoteNum)
        Me.Controls.Add(Me.txtOrderNum)
        Me.Controls.Add(Me.txtDesign)
        Me.Controls.Add(Me.txtInvHeading)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnPrint)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "GenCutSheetInvoice"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Create a Cutting Sheet Invoice"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private CallingForm As Object
    Dim DbConnection As OleDb.OleDbConnection

    Public Sub New(ByRef caller As Object, ByRef dbc As OleDb.OleDbConnection)
        MyBase.New()
        InitializeComponent()
        CallingForm = caller
        DbConnection = dbc
    End Sub

    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, ByVal keyData As Keys) As Boolean
        If keyData = Keys.Enter Then
            SendKeys.Send("{Tab}")
            Return True
        End If

        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function

    Private Sub frmGenCutSheet_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If Not IsNothing(CallingForm) Then
            CallingForm.Show()
        End If

        CallingForm = Nothing
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Close()
    End Sub

    Private Sub btnPrintInvoice_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateInvoice.Click
        generateCuttingSheetInvoice(cmb_AllCutSheets.Text, txtDelNoteNum.Text, txtOrderNum.Text, txtInvHeading.Text, txtDesign.Text, dtpInvDate.Value.ToLongDateString)
    End Sub

    Dim InvoiceNumber As Long

    Private Sub generateCuttingSheetInvoice(ByVal CuttingSheetNumber As String, ByVal DeliveryNoteNumber As String, ByVal OrderNo As String, ByVal InvoiceHeader As String, ByVal Design As String, ByVal InvoiceDate As Date)

        If CuttingSheetNumber = "" Or CuttingSheetNumber = "Please Select..." Then
            CuttingSheetNumber = "-1"
        End If
        If Design = "" Then
            Design = "0"
        End If

        Dim SQL As String = "SELECT * FROM ((((((CuttingSheet INNER JOIN Job ON [CuttingSheet].[Job No] = Job.JobNo ) INNER JOIN Company ON Job.CompanyNo = Company.CompanyNo) INNER JOIN SchedItem ON CuttingSheet.CutSheetNo = SchedItem.CutSheetNo) INNER JOIN CutItem ON SchedItem.CutSheetNo = CutItem.CutSheetNo AND SchedItem.ScheduleNo = CutItem.ScheduleNo) INNER JOIN ProductType ON CutItem.TypeCode = ProductType.TypeCode) INNER JOIN JobRate ON Job.JobNo = JobRate.JobNo AND JobRate.TypeCode = ProductType.TypeCode ) WHERE CuttingSheet.CutSheetNo = " + CuttingSheetNumber + " ORDER BY ProductType.TypeCode"
        Dim DataSet As New Data.DataSet
        Dim Adapter As New OleDb.OleDbDataAdapter(SQL, DbConnection)
        Adapter.Fill(DataSet)

        If DataSet.Tables(0).Rows.Count = 0 Then
            If CuttingSheetNumber = "-1" Then
                MessageBox.Show("Please enter a valid Cutting Sheet Number.", "Invalid Cutting Sheet Number.")
            Else
                MessageBox.Show("Cannot generate Invoice. Possible reason :" & Microsoft.VisualBasic.Chr(13) & "1. Rates are not specified for the job." & Microsoft.VisualBasic.Chr(13) & "2. The cutting sheet has no Schedules and/or Items.")
            End If

        Else
            If DataSet.Tables(0).Rows(0).Item("InvoiceNo").ToString = "0" Or DataSet.Tables(0).Rows(0).Item("InvoiceNo").ToString = "" Then
                'This IF STATEMENT checks if the cutting sheet has not been invoiced.

                Try
                    InvoiceNumber = Long.Parse(DataSet.Tables(0).Rows(0).Item("LastInvNum").ToString()) + 1
                Catch ex As Exception
                    MessageBox.Show("Gotcha!")
                End Try

                Dim InvoiceType As String = "Cutting Sheet"
                Dim Factor As Int16 = 1.5
                Dim EscMoAndDa As Date = New Date(1999, 12, 11)
                Dim Work As Int16 = 75
                Dim RefNo As String = "Ref"
                Dim JobNo As String = DataSet.Tables(0).Rows(0).Item("Job.JobNo").ToString()
                Dim VAT As String = DataSet.Tables(0).Rows(0).Item("VatPerc").ToString()
                Dim Active = "Yes"
                Dim Escalated = "No"
                Dim OnSummary = "Yes"
                Dim Comments As String = "Comments"


                Dim CalcTotal = 0, CalcVat = 0, CalcNett = 0

                Dim SQL4NewInvoice As String = "INSERT INTO Invoice(InvoiceNo,InvoiceType,InvDate,InvDeliveryNoteNo,InvFactor,Invmonthandyear,InvWork,InvOrdNum,InvRefNum,InvoiceHeading,InvTotal,InvVatAmt,InvDesign,InvNett,InvJobNo,InvActive,InvEscalated,InvOnSummary,InvComments) VALUES " & _
                "(  " & _
                    InvoiceNumber.ToString & _
                    ",'" & InvoiceType & _
                    "',#" & InvoiceDate.ToLongDateString & _
                    "#,'" & DeliveryNoteNumber & _
                    "'," & Factor & _
                    ",#" & EscMoAndDa & _
                    "#," & Work & _
                    ",'" & OrderNo & _
                    "','" & RefNo & _
                    "','" & InvoiceHeader & _
                    "'," & "-1" & _
                    "," & "-1" & _
                    "," & Design & _
                    "," & "-1" & _
                    ",'" & JobNo & _
                    "'," & Active & _
                    "," & Escalated & _
                    "," & OnSummary & _
                    ",'" & Comments & _
                    "')"


                Dim command As New OleDb.OleDbCommand(SQL4NewInvoice, DbConnection)
                Try
                    DbConnection.Open()
                    command.ExecuteNonQuery()

                    'Create Invoice Lines

                    Dim lcv
                    Dim CurTypeCode, NextTypeCode
                    Dim TotalLengthForType = 0, TypeMass = 0
                    Dim TotalMassForType = 0
                    Dim TotalCostForType = 0
                    Dim LineNumberCounter As Int16 = 1
                    Dim DESCRIPTION As String = "Description"

                    For lcv = 0 To DataSet.Tables(0).Rows.Count - 1

                        CurTypeCode = DataSet.Tables(0).Rows(lcv).Item("CutItem.TypeCode").ToString()
                        'MessageBox.Show("Current Type = " + CurTypeCode)

                        If lcv + 1 <= DataSet.Tables(0).Rows.Count - 1 Then
                            NextTypeCode = DataSet.Tables(0).Rows(lcv + 1).Item("CutItem.TypeCode").ToString()
                        Else
                            NextTypeCode = "NO NEXT"
                        End If

                        TotalLengthForType += Double.Parse(DataSet.Tables(0).Rows(lcv).Item("Qty").ToString) * Double.Parse(DataSet.Tables(0).Rows(lcv).Item("Length").ToString)

                        'If NextTypeCode <> CurTypeCode Or lcv = DataSet.Tables(0).Rows.Count - 1 Then
                        If NextTypeCode <> CurTypeCode Or NextTypeCode = "NO NEXT" Then
                            TotalLengthForType /= 1000
                            TypeMass = Double.Parse(DataSet.Tables(0).Rows(lcv).Item("Weight").ToString())
                            If DataSet.Tables(0).Rows(lcv).Item("Tons Or Kilograms").ToString() = "T" Then
                                TotalMassForType = TotalLengthForType * TypeMass / 1000
                                TotalMassForType = Decimal.Parse(TotalMassForType.ToString()).Round(TotalMassForType, 3)
                                TotalCostForType = TotalMassForType * DataSet.Tables(0).Rows(lcv).Item("Rate")
                            Else ' KG
                                TotalMassForType = TotalLengthForType * TypeMass
                                TotalMassForType = Decimal.Parse(TotalMassForType.ToString()).Round(TotalMassForType, 1)
                                TotalCostForType = Math.Round(TotalMassForType, 1) * DataSet.Tables(0).Rows(lcv).Item("Rate")
                            End If

                            Dim SQL4NewInvoiceLine As String = "INSERT INTO InvoiceLine(InvNo,[Line#],TypeCode,Description,Qty,TonsorKg,CostPerUnit,Total) VALUES " & _
                            "(" & InvoiceNumber & _
                            "," & LineNumberCounter & _
                            ",'" & CurTypeCode & _
                            "','" & DESCRIPTION & _
                            "'," & TotalMassForType & _
                            ",'" & DataSet.Tables(0).Rows(lcv).Item("Tons Or Kilograms").ToString() & _
                            "'," & DataSet.Tables(0).Rows(lcv).Item("Rate").ToString() & _
                            "," & TotalCostForType & _
                            ")"

                            LineNumberCounter += 1
                            CalcTotal += TotalCostForType

                            TotalLengthForType = 0
                            TypeMass = 0
                            TotalMassForType = 0
                            TotalCostForType = 0

                            Dim InvLineCommand As New OleDb.OleDbCommand(SQL4NewInvoiceLine, DbConnection)
                            Try
                                InvLineCommand.ExecuteNonQuery()
                                'MessageBox.Show("Executed : " + SQL4NewInvoiceLine)
                            Catch mex As Exception
                                MessageBox.Show(mex.Message)
                            End Try


                        End If
                    Next lcv

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
                    Dim SQL4CompanyUpdate As String = "UPDATE Company SET Company.LastInvNum = " + InvoiceNumber.ToString() + " WHERE Company.CompanyNo = '" + DataSet.Tables(0).Rows(0).Item("Company.CompanyNo").ToString() + "'"
                    Dim UpdateCommand As New OleDb.OleDbCommand(SQL4CompanyUpdate, DbConnection)
                    Try
                        UpdateCommand.ExecuteNonQuery()
                    Catch MEEE As Exception
                        MessageBox.Show(MEEE.Message)
                    End Try

                    'Update Cutting Sheet Invoice Number
                    Dim SQL4CuttingSheetUpdate As String = "UPDATE CuttingSheet SET CuttingSheet.InvoiceNo = " + InvoiceNumber.ToString() + " WHERE CuttingSheet.CutSheetNo = " + DataSet.Tables(0).Rows(0).Item("CuttingSheet.CutSheetNo").ToString()
                    Dim UpdateInvCommand As New OleDb.OleDbCommand(SQL4CuttingSheetUpdate, DbConnection)
                    Try
                        UpdateInvCommand.ExecuteNonQuery()

                    Catch MEE As Exception
                        MessageBox.Show(MEE.Message)
                    End Try

                    'MessageBox.Show("Cutting Sheet Number: " + CuttingSheetNumber.ToString + " has been successfully invoiced - Invoice Number: " + InvoiceNumber.ToString())
                    popCMB()
                    lblNotify.Text = "Invoice Number: " + InvoiceNumber.ToString()
                    btnPrint.Enabled = True
                Catch Myerror As Exception
                    MessageBox.Show(Myerror.Message)
                Finally
                    DbConnection.Close()
                End Try
            Else
                MessageBox.Show("Cutting sheet " + CuttingSheetNumber + " has already been invoiced. See invoice no. " + DataSet.Tables(0).Rows(0).Item("InvoiceNo").ToString)
            End If

        End If
    End Sub

    Private Sub GenCutSheetInvoice_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        popCMB()
    End Sub

    Private Sub popCMB()
        cmb_AllCutSheets.Items.Clear()
        Dim sql = "SELECT CutSheetNo FROM CuttingSheet WHERE InvoiceNo = 0 ORDER BY CutSheetNo ASC"
        Dim DataSet As New Data.DataSet
        Dim Adapter As New OleDb.OleDbDataAdapter(sql, DbConnection)
        Adapter.Fill(DataSet)
        Dim d
        For d = 0 To DataSet.Tables(0).Rows.Count - 1
            cmb_AllCutSheets.Items.Add(DataSet.Tables(0).Rows(d).Item("CutSheetNo").ToString())
        Next d
    End Sub

    Private Sub cmb_AllCutSheets_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb_AllCutSheets.SelectedIndexChanged
        Dim sql = "SELECT OrderNo FROM (CuttingSheet INNER JOIN Job ON CuttingSheet.[Job No] = Job.JobNo) WHERE CuttingSheet.CutSheetNo = " & cmb_AllCutSheets.Text
        Dim ds As New Data.DataSet
        Dim Adapter As New OleDb.OleDbDataAdapter(sql, DbConnection)
        Adapter.Fill(ds)

        If ds.Tables(0).Rows.Count <> 0 Then
            txtOrderNum.Text = ds.Tables(0).Rows(0).Item("OrderNo").ToString()
        End If

        lblNotify.Text = ""
        btnPrint.Enabled = False

    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim Form As PrintCutInv = New PrintCutInv(Me)
        Form.populate_invoiceNumbers()
        Form.txt_InvNumToPrint.SelectedIndex = Form.txt_InvNumToPrint.Items.IndexOf(InvoiceNumber.ToString)
        Form.btn_PrintInv_Click(sender, e)

    End Sub

    Private Sub cmb_AllCutSheets_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmb_AllCutSheets.Leave
        If Not cmb_AllCutSheets.SelectedIndex >= 0 Then
            If Not IsNumeric(cmb_AllCutSheets.Text) Then
                cmb_AllCutSheets.Text = "Please Select..."
            Else
                cmb_AllCutSheets_SelectedIndexChanged(sender, e)
            End If
        End If
    End Sub
End Class
