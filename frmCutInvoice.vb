Imports System
Imports System.Windows.Forms
Imports System.Drawing
Imports LogicTier

Public Class GenCutSheetInvoice
    Inherits Form

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
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents btnClose As Button
    Friend WithEvents txtDelNoteNum As TextBox
    Friend WithEvents txtOrderNum As TextBox
    Friend WithEvents txtDesign As TextBox
    Friend WithEvents txtInvHeading As TextBox
    Friend WithEvents dtpInvDate As DateTimePicker
    Friend WithEvents btnCreateInvoice As Button
    Friend WithEvents cmb_AllCutSheets As ComboBox
    Friend WithEvents lblNotify As Label
    Friend WithEvents btnPrint As Button
    <Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label1 = New Label
        Me.Label2 = New Label
        Me.Label3 = New Label
        Me.Label4 = New Label
        Me.Label5 = New Label
        Me.Label7 = New Label
        Me.txtDelNoteNum = New TextBox
        Me.txtOrderNum = New TextBox
        Me.txtDesign = New TextBox
        Me.txtInvHeading = New TextBox
        Me.dtpInvDate = New DateTimePicker
        Me.btnCreateInvoice = New Button
        Me.btnClose = New Button
        Me.cmb_AllCutSheets = New ComboBox
        Me.lblNotify = New Label
        Me.btnPrint = New Button
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New Point(16, 72)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New Size(128, 23)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Delivery Note Number"
        '
        'Label2
        '
        Me.Label2.Location = New Point(16, 112)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New Size(128, 23)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Order Number"
        '
        'Label3
        '
        Me.Label3.Location = New Point(16, 152)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New Size(128, 23)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "Design"
        '
        'Label4
        '
        Me.Label4.Location = New Point(16, 192)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New Size(128, 23)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "Invoice Date"
        '
        'Label5
        '
        Me.Label5.Location = New Point(16, 232)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New Size(128, 23)
        Me.Label5.TabIndex = 0
        Me.Label5.Text = "Invoice Heading"
        '
        'Label7
        '
        Me.Label7.Location = New Point(16, 32)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New Size(128, 23)
        Me.Label7.TabIndex = 0
        Me.Label7.Text = "Cutting Sheet Number "
        '
        'txtDelNoteNum
        '
        Me.txtDelNoteNum.Location = New Point(152, 72)
        Me.txtDelNoteNum.MaxLength = 50
        Me.txtDelNoteNum.Name = "txtDelNoteNum"
        Me.txtDelNoteNum.Size = New Size(304, 20)
        Me.txtDelNoteNum.TabIndex = 2
        Me.txtDelNoteNum.Text = ""
        '
        'txtOrderNum
        '
        Me.txtOrderNum.Location = New Point(152, 112)
        Me.txtOrderNum.Name = "txtOrderNum"
        Me.txtOrderNum.Size = New Size(240, 20)
        Me.txtOrderNum.TabIndex = 4
        Me.txtOrderNum.Text = ""
        '
        'txtDesign
        '
        Me.txtDesign.Location = New Point(152, 152)
        Me.txtDesign.Name = "txtDesign"
        Me.txtDesign.Size = New Size(144, 20)
        Me.txtDesign.TabIndex = 6
        Me.txtDesign.Text = ""
        '
        'txtInvHeading
        '
        Me.txtInvHeading.Location = New Point(152, 232)
        Me.txtInvHeading.Name = "txtInvHeading"
        Me.txtInvHeading.Size = New Size(312, 20)
        Me.txtInvHeading.TabIndex = 10
        Me.txtInvHeading.Text = ""
        '
        'dtpInvDate
        '
        Me.dtpInvDate.Format = DateTimePickerFormat.Short
        Me.dtpInvDate.Location = New Point(152, 192)
        Me.dtpInvDate.Name = "dtpInvDate"
        Me.dtpInvDate.Size = New Size(144, 20)
        Me.dtpInvDate.TabIndex = 8
        '
        'btnCreateInvoice
        '
        Me.btnCreateInvoice.Location = New Point(64, 304)
        Me.btnCreateInvoice.Name = "btnCreateInvoice"
        Me.btnCreateInvoice.Size = New Size(112, 23)
        Me.btnCreateInvoice.TabIndex = 12
        Me.btnCreateInvoice.Text = "Create Invoice"
        '
        'btnClose
        '
        Me.btnClose.Location = New Point(336, 304)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New Size(112, 23)
        Me.btnClose.TabIndex = 14
        Me.btnClose.Text = "Close"
        '
        'cmb_AllCutSheets
        '
        Me.cmb_AllCutSheets.Location = New Point(152, 32)
        Me.cmb_AllCutSheets.MaxDropDownItems = 15
        Me.cmb_AllCutSheets.Name = "cmb_AllCutSheets"
        Me.cmb_AllCutSheets.Size = New Size(144, 21)
        Me.cmb_AllCutSheets.TabIndex = 0
        '
        'lblNotify
        '
        Me.lblNotify.Font = New Font("Microsoft Sans Serif", 10.0!, FontStyle.Bold, GraphicsUnit.Point, CType(0, Byte))
        Me.lblNotify.Location = New Point(16, 264)
        Me.lblNotify.Name = "lblNotify"
        Me.lblNotify.Size = New Size(488, 32)
        Me.lblNotify.TabIndex = 15
        Me.lblNotify.TextAlign = ContentAlignment.MiddleCenter
        '
        'btnPrint
        '
        Me.btnPrint.Enabled = False
        Me.btnPrint.Location = New Point(200, 304)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New Size(112, 23)
        Me.btnPrint.TabIndex = 12
        Me.btnPrint.Text = "Print Preview..."
        '
        'GenCutSheetInvoice
        '
        Me.AutoScaleBaseSize = New Size(5, 13)
        Me.ClientSize = New Size(538, 351)
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
        Me.FormBorderStyle = FormBorderStyle.FixedToolWindow
        Me.Name = "GenCutSheetInvoice"
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.Text = "Create a Cutting Sheet Invoice"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Property Logic As CutInvoice
    Private Property CallingForm As Object
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

    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
        Close()
    End Sub

    Private Sub btnPrintInvoice_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnCreateInvoice.Click
        generateCuttingSheetInvoice(cmb_AllCutSheets.Text, txtDelNoteNum.Text, txtOrderNum.Text, txtInvHeading.Text, txtDesign.Text, dtpInvDate.Value.ToLongDateString)
    End Sub

    Dim InvoiceNumber As Long

    Private Sub generateCuttingSheetInvoice(ByVal CuttingSheetNumber As String, ByVal DeliveryNoteNumber As String, ByVal OrderNo As String, ByVal InvoiceHeader As String, ByVal Design As String, ByVal InvoiceDate As Date)
        If Logic.InvoiceCount = 0 Then
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
                Dim Active As String = "Yes"
                Dim Escalated As String = "No"
                Dim OnSummary As String = "Yes"
                Dim Comments As String = "Comments"


                Dim CalcTotal As Integer = 0
                Dim CalcVat As Integer = 0
                Dim CalcNett As Integer = 0

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

                    Dim lcv As Integer
                    Dim CurTypeCode As String
                    Dim NextTypeCode As String
                    Dim TotalLengthForType As Integer = 0
                    Dim TypeMass As Integer = 0
                    Dim TotalMassForType As Integer = 0
                    Dim TotalCostForType As Integer = 0
                    Dim LineNumberCounter As Integer = 1
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
                                TotalMassForType = Decimal.Round(TotalMassForType, 3)
                                TotalCostForType = TotalMassForType * DataSet.Tables(0).Rows(lcv).Item("Rate")
                            Else ' KG
                                TotalMassForType = TotalLengthForType * TypeMass
                                TotalMassForType = Decimal.Round(TotalMassForType, 1)
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

    Private Sub GenCutSheetInvoice_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        popCMB()
    End Sub

    Private Sub popCMB()
        cmb_AllCutSheets.Items.Clear()
        Dim sql As String = "SELECT CutSheetNo FROM CuttingSheet WHERE InvoiceNo = 0 ORDER BY CutSheetNo ASC"
        Dim DataSet As New Data.DataSet
        Dim Adapter As New OleDb.OleDbDataAdapter(sql, DbConnection)
        Adapter.Fill(DataSet)
        Dim d As Integer
        For d = 0 To DataSet.Tables(0).Rows.Count - 1
            cmb_AllCutSheets.Items.Add(DataSet.Tables(0).Rows(d).Item("CutSheetNo").ToString())
        Next d
    End Sub

    Private Sub cmb_AllCutSheets_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles cmb_AllCutSheets.SelectedIndexChanged
        Dim sql As String = "SELECT OrderNo FROM (CuttingSheet INNER JOIN Job ON CuttingSheet.[Job No] = Job.JobNo) WHERE CuttingSheet.CutSheetNo = " & cmb_AllCutSheets.Text
        Dim ds As New Data.DataSet
        Dim Adapter As New OleDb.OleDbDataAdapter(sql, DbConnection)
        Adapter.Fill(ds)

        If ds.Tables(0).Rows.Count <> 0 Then
            txtOrderNum.Text = ds.Tables(0).Rows(0).Item("OrderNo").ToString()
        End If

        lblNotify.Text = ""
        btnPrint.Enabled = False

    End Sub

    Private Sub btnPrint_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnPrint.Click
        Dim Form As PrintCutInv = New PrintCutInv(Me)
        Form.populate_invoiceNumbers()
        Form.txt_InvNumToPrint.SelectedIndex = Form.txt_InvNumToPrint.Items.IndexOf(InvoiceNumber.ToString)
        Form.btn_PrintInv_Click(sender, e)

    End Sub

    Private Sub cmb_AllCutSheets_Leave(ByVal sender As Object, ByVal e As EventArgs) Handles cmb_AllCutSheets.Leave
        If Not cmb_AllCutSheets.SelectedIndex >= 0 Then
            If Not IsNumeric(cmb_AllCutSheets.Text) Then
                cmb_AllCutSheets.Text = "Please Select..."
            Else
                cmb_AllCutSheets_SelectedIndexChanged(sender, e)
            End If
        End If
    End Sub
End Class
