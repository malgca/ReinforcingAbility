Public Class frmInvSummary
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Dim CallingForm As Object
    Dim DBConnection As OleDb.OleDbConnection

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
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents chkS As System.Windows.Forms.CheckBox
    Friend WithEvents chkM As System.Windows.Forms.CheckBox
    Friend WithEvents chkCS As System.Windows.Forms.CheckBox
    Friend WithEvents nudYear As System.Windows.Forms.NumericUpDown
    Friend WithEvents cmbMonth As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents cmbJobNumber As System.Windows.Forms.ComboBox
    Friend WithEvents rdbAll As System.Windows.Forms.RadioButton
    Friend WithEvents rdbSingle As System.Windows.Forms.RadioButton
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents DocumentToPrint As System.Drawing.Printing.PrintDocument
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.chkS = New System.Windows.Forms.CheckBox
        Me.chkM = New System.Windows.Forms.CheckBox
        Me.chkCS = New System.Windows.Forms.CheckBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.nudYear = New System.Windows.Forms.NumericUpDown
        Me.cmbMonth = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.rdbAll = New System.Windows.Forms.RadioButton
        Me.cmbJobNumber = New System.Windows.Forms.ComboBox
        Me.rdbSingle = New System.Windows.Forms.RadioButton
        Me.Label2 = New System.Windows.Forms.Label
        Me.btnPrint = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.DocumentToPrint = New System.Drawing.Printing.PrintDocument
        Me.GroupBox1.SuspendLayout()
        CType(Me.nudYear, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'chkS
        '
        Me.chkS.Checked = True
        Me.chkS.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkS.Location = New System.Drawing.Point(40, 64)
        Me.chkS.Name = "chkS"
        Me.chkS.Size = New System.Drawing.Size(152, 24)
        Me.chkS.TabIndex = 0
        Me.chkS.Text = "Sundry Invoices"
        '
        'chkM
        '
        Me.chkM.Checked = True
        Me.chkM.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkM.Location = New System.Drawing.Point(40, 40)
        Me.chkM.Name = "chkM"
        Me.chkM.Size = New System.Drawing.Size(152, 24)
        Me.chkM.TabIndex = 0
        Me.chkM.Text = "Mesh Invoices"
        '
        'chkCS
        '
        Me.chkCS.Checked = True
        Me.chkCS.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkCS.Location = New System.Drawing.Point(40, 16)
        Me.chkCS.Name = "chkCS"
        Me.chkCS.Size = New System.Drawing.Size(152, 24)
        Me.chkCS.TabIndex = 0
        Me.chkCS.Text = "Cutting Sheet Invoices"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.chkCS)
        Me.GroupBox1.Controls.Add(Me.chkM)
        Me.GroupBox1.Controls.Add(Me.chkS)
        Me.GroupBox1.Location = New System.Drawing.Point(16, 232)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(224, 96)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Invoice Types for Summary"
        '
        'nudYear
        '
        Me.nudYear.Location = New System.Drawing.Point(72, 56)
        Me.nudYear.Maximum = New Decimal(New Integer() {2999, 0, 0, 0})
        Me.nudYear.Minimum = New Decimal(New Integer() {1969, 0, 0, 0})
        Me.nudYear.Name = "nudYear"
        Me.nudYear.TabIndex = 2
        Me.nudYear.Value = New Decimal(New Integer() {2003, 0, 0, 0})
        '
        'cmbMonth
        '
        Me.cmbMonth.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbMonth.Items.AddRange(New Object() {"January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"})
        Me.cmbMonth.Location = New System.Drawing.Point(72, 24)
        Me.cmbMonth.MaxDropDownItems = 12
        Me.cmbMonth.Name = "cmbMonth"
        Me.cmbMonth.Size = New System.Drawing.Size(121, 21)
        Me.cmbMonth.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 23)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Month:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.rdbAll)
        Me.GroupBox2.Controls.Add(Me.cmbJobNumber)
        Me.GroupBox2.Controls.Add(Me.rdbSingle)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Location = New System.Drawing.Point(16, 16)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(224, 112)
        Me.GroupBox2.TabIndex = 5
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Job Selection"
        '
        'rdbAll
        '
        Me.rdbAll.Checked = True
        Me.rdbAll.Location = New System.Drawing.Point(16, 24)
        Me.rdbAll.Name = "rdbAll"
        Me.rdbAll.Size = New System.Drawing.Size(72, 16)
        Me.rdbAll.TabIndex = 1
        Me.rdbAll.TabStop = True
        Me.rdbAll.Text = "All Jobs"
        '
        'cmbJobNumber
        '
        Me.cmbJobNumber.Enabled = False
        Me.cmbJobNumber.Location = New System.Drawing.Point(88, 72)
        Me.cmbJobNumber.Name = "cmbJobNumber"
        Me.cmbJobNumber.Size = New System.Drawing.Size(104, 21)
        Me.cmbJobNumber.TabIndex = 0
        '
        'rdbSingle
        '
        Me.rdbSingle.Location = New System.Drawing.Point(16, 48)
        Me.rdbSingle.Name = "rdbSingle"
        Me.rdbSingle.Size = New System.Drawing.Size(88, 16)
        Me.rdbSingle.TabIndex = 1
        Me.rdbSingle.Text = "Single Job"
        '
        'Label2
        '
        Me.Label2.Enabled = False
        Me.Label2.Location = New System.Drawing.Point(16, 72)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(72, 23)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Job Number:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnPrint
        '
        Me.btnPrint.Location = New System.Drawing.Point(16, 344)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(104, 23)
        Me.btnPrint.TabIndex = 6
        Me.btnPrint.Text = "Print Preview..."
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(136, 344)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(104, 23)
        Me.Button2.TabIndex = 6
        Me.Button2.Text = "Close"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.cmbMonth)
        Me.GroupBox3.Controls.Add(Me.nudYear)
        Me.GroupBox3.Controls.Add(Me.Label1)
        Me.GroupBox3.Controls.Add(Me.Label3)
        Me.GroupBox3.Location = New System.Drawing.Point(16, 136)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(224, 88)
        Me.GroupBox3.TabIndex = 7
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Month && Year Selection"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(16, 56)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(40, 23)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Year:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'frmInvSummary
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(250, 385)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Button2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "frmInvSummary"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Invoice Summary - Print Options"
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.nudYear, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Global Variables "
    Dim field As PageElement
    Dim PrintArray As New ArrayList
    Dim EntryFont As New Font("Arial", 10)
    Dim Head1Font As New Font("Arial", 30, FontStyle.Bold Or FontStyle.Underline)
    Dim Head2Font As New Font("Arial", 15, FontStyle.Bold)
    Dim Head2DetFont As New Font("Arial", 15, FontStyle.Italic)
    Dim EntryFontBold As New Font("Arial", 10, FontStyle.Bold)
    Dim EntryFontUnderline As New Font("Arial", 10, FontStyle.Underline)
    Dim DetailFont As New Font("Arial", 13)
    Dim ColFont As New Font("Arial", 12, FontStyle.Italic)
    Dim curArrayPos = 0
    Dim curpagenum = 1
    Dim TopMargin = 70
    Dim LeftMargin = 100
    Dim RightMargin = 90
    Dim BottomMargin = 80
    Dim PageWidth = 873
    Dim ReportType
    Dim mes
    Dim vatperc As String
    Dim All_Is_OK As Boolean = True
    Const col1 = 100
    Const col2 = 250
    Const col3 = 400
    Const col4 = 550
    Const col5 = 700

#End Region

    Private Sub frmInvSummary_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cmbMonth.SelectedIndex = Now.Month - 1
        nudYear.Value = Now.Year
        populateCmbJobs()
    End Sub




    Private Sub populateCmbJobs()
        Dim sql = "SELECT jobno FROM job ORDER BY jobno"
        Dim ds As Data.DataSet = New Data.DataSet
        Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(sql, DBConnection)
        da.Fill(ds)

        Dim i
        For i = 0 To ds.Tables(0).Rows.Count - 1
            cmbJobNumber.Items.Add(ds.Tables(0).Rows(i).Item("JobNo").ToString())
        Next i

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
        CallingForm.show()
    End Sub

    Private Sub rdbSingle_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdbSingle.CheckedChanged
        Label2.Enabled = rdbSingle.Checked
        cmbJobNumber.Enabled = rdbSingle.Checked
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Try
            DocumentToPrint.DocumentName = "Summary of Invoices"
            Dim ppd_JCR As New PrintPreviewDialog
            ppd_JCR.WindowState = FormWindowState.Maximized
            ppd_JCR.Document = DocumentToPrint
            ppd_JCR.AutoScale = True
            ppd_JCR.AutoScroll = True
            ppd_JCR.UseAntiAlias = False
            ppd_JCR.PrintPreviewControl.Zoom = 1
            ppd_JCR.PrintPreviewControl.Columns = 1
            ppd_JCR.PrintPreviewControl.Rows = 1
            ppd_JCR.Text = "Summary of Invoices"
            curpagenum = 1
            PrintArray.Clear()
            All_Is_OK = True
            If chkCS.Checked Then
                addSummary("Cutting Sheet")
                PrintArray.Add(New PageElement("<PAGE BREAK>", EntryFont, 0, 0, False))
            End If
            If chkM.Checked Then
                addSummary("Mesh")
                PrintArray.Add(New PageElement("<PAGE BREAK>", EntryFont, 0, 0, False))
            End If
            If chkS.Checked Then
                addSummary("Sundry")
            End If

            curArrayPos = 0
            If All_Is_OK Then
                ppd_JCR.ShowDialog()
            End If

        Catch er As Exception
            If er.Message = "No printers installed." Then
                MessageBox.Show("There is no printer installed. Please install a printer and try again.", "Printer not found", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                MessageBox.Show(er.Message, "ERROR 1!")
            End If
        End Try
    End Sub

    Private Sub addSummary(ByVal type As String)

        If rdbAll.Checked Then
            PrintArray.Add(New PageElement("MONTHLY INVOICES - ALL JOBS", EntryFont, 0, 0, True))
        ElseIf rdbSingle.Checked Then
            PrintArray.Add(New PageElement("MONTHLY INVOICES - JOB NUMBER: " & cmbJobNumber.Text, EntryFont, 0, 0, True))
        End If

        PrintArray.Add(New PageElement(" ", EntryFont, 0, 0, True))
        PrintArray.Add(New PageElement(type & " Totals for the Month of " & cmbMonth.Text & " " & nudYear.Text, EntryFont, 0, 0, True))
        PrintArray.Add(New PageElement("Invoice Number", EntryFont, col1, False, False, False))
        PrintArray.Add(New PageElement("Invoice Date", EntryFont, col2, False, False, False))
        PrintArray.Add(New PageElement("Invoice Value", EntryFont, col3, False, False, False))
        PrintArray.Add(New PageElement("Vat", EntryFont, col4, False, False, False))
        PrintArray.Add(New PageElement("Nett", EntryFont, col5, True, False, False))
        PrintArray.Add(New PageElement(col1, col5 + 75, False))

        Dim TOTALNETT = 0
        Dim TOTALVAT = 0
        Dim TOTALVALUE = 0

        Dim sql
        If rdbAll.Checked Then
            sql = "SELECT InvoiceNo,invDate,InvTotal,InvVatAmt,InvNett FROM Invoice WHERE InvoiceType = '" & type & "'"
        Else
            sql = "SELECT InvoiceNo,invDate,InvTotal,InvVatAmt,InvNett FROM Invoice WHERE InvoiceType = '" & type & "' AND InvJobNo = '" & cmbJobNumber.Text & "'"
        End If

        Dim ds As Data.DataSet = New Data.DataSet
        Dim ad As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(sql, DBConnection)
        ad.Fill(ds)

        Dim lcv = 0, count = 0
        For lcv = 0 To ds.Tables(0).Rows().Count - 1
            Dim d As Date = ds.Tables(0).Rows(lcv).Item("invDate")

            If d.Year = nudYear.Text And d.Month = cmbMonth.SelectedIndex + 1 Then
                PrintArray.Add(New PageElement(ds.Tables(0).Rows(lcv).Item("invoiceNo").ToString, EntryFont, col1, False, False, False))
                PrintArray.Add(New PageElement(d.ToShortDateString, EntryFont, col2, False, False, False))
                PrintArray.Add(New PageElement(PrintCutInv.toRand(ds.Tables(0).Rows(lcv).Item("InvTotal").ToString, True), EntryFont, col3 + 75, False, False, True))
                PrintArray.Add(New PageElement(PrintCutInv.toRand(ds.Tables(0).Rows(lcv).Item("invVatAmt").ToString, True), EntryFont, col4 + 75, False, False, True))
                PrintArray.Add(New PageElement(PrintCutInv.toRand(ds.Tables(0).Rows(lcv).Item("invNett").ToString, True), EntryFont, col5 + 75, True, False, True))
                TOTALNETT += ds.Tables(0).Rows(lcv).Item("invNett")
                TOTALVALUE += ds.Tables(0).Rows(lcv).Item("invTotal")
                TOTALVAT += ds.Tables(0).Rows(lcv).Item("invVatAmt")
                count += 1

            End If
        Next lcv

        If count = 0 Then
            PrintArray.Add(New PageElement(" No invoices to display", EntryFont, 0, 0, True))
        End If

        PrintArray.Add(New PageElement(col1, col5 + 75, False))
        PrintArray.Add(New PageElement("Totals:", EntryFont, col1, False, False, False))
        PrintArray.Add(New PageElement(PrintCutInv.toRand(TOTALVALUE, True), EntryFont, col3 + 75, False, False, True))
        PrintArray.Add(New PageElement(PrintCutInv.toRand(TOTALVAT, True), EntryFont, col4 + 75, False, False, True))
        PrintArray.Add(New PageElement(PrintCutInv.toRand(TOTALNETT, True), EntryFont, col5 + 75, True, False, True))
        PrintArray.Add(New PageElement(True, col1, col5 + 75))



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

End Class
