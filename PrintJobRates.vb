Public Class PrintJobRates
    Inherits System.Windows.Forms.Form




#Region " Windows Form Designer generated code "

    Public Sub New(ByRef dbc As OleDb.OleDbConnection)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
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
    Friend WithEvents btn_Close As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btn_Print As System.Windows.Forms.Button
    Friend WithEvents DocumentToPrint As System.Drawing.Printing.PrintDocument
    Friend WithEvents txtJobNum As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.DocumentToPrint = New System.Drawing.Printing.PrintDocument
        Me.btn_Print = New System.Windows.Forms.Button
        Me.btn_Close = New System.Windows.Forms.Button
        Me.txtJobNum = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'DocumentToPrint
        '
        '
        'btn_Print
        '
        Me.btn_Print.Location = New System.Drawing.Point(32, 80)
        Me.btn_Print.Name = "btn_Print"
        Me.btn_Print.Size = New System.Drawing.Size(96, 23)
        Me.btn_Print.TabIndex = 0
        Me.btn_Print.Text = "Print Preview..."
        '
        'btn_Close
        '
        Me.btn_Close.Location = New System.Drawing.Point(152, 80)
        Me.btn_Close.Name = "btn_Close"
        Me.btn_Close.Size = New System.Drawing.Size(96, 23)
        Me.btn_Close.TabIndex = 1
        Me.btn_Close.Text = "Close"
        '
        'txtJobNum
        '
        Me.txtJobNum.Location = New System.Drawing.Point(152, 24)
        Me.txtJobNum.Name = "txtJobNum"
        Me.txtJobNum.Size = New System.Drawing.Size(96, 21)
        Me.txtJobNum.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(32, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Job No"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'PrintJobRates
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(292, 142)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtJobNum)
        Me.Controls.Add(Me.btn_Close)
        Me.Controls.Add(Me.btn_Print)
        Me.Name = "PrintJobRates"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Print Job Rates"
        Me.ResumeLayout(False)

    End Sub

#End Region

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
    Dim TopMargin = 90
    Dim LeftMargin = 60
    Dim RightMargin = 90
    Dim BottomMargin = 60
    Dim PageWidth = 873
    Dim ReportType

    Dim All_Is_OK As Boolean = True
#End Region



    Private CallingForm As Object
    Public Sub New(ByVal caller As Object)
        MyBase.New()
        InitializeComponent()
        CallingForm = caller
    End Sub

    Private Sub frmPrintCut_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If Not IsNothing(CallingForm) Then
            CallingForm.Show()
        End If

        CallingForm = Nothing
    End Sub

    Private Sub PrintJob_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        populate_jobNumbers()
    End Sub

    Private Sub populate_jobNumbers()
        txtJobNum.Items.Clear()
        Dim sql = "SELECT Job.JobNo FROM Job"
        Dim ds As New Data.DataSet
        Dim ad As New OleDb.OleDbDataAdapter(sql, DBConnection)
        ad.Fill(ds)

        Dim f
        For f = 0 To ds.Tables(0).Rows.Count - 1
            txtJobNum.Items.Add(ds.Tables(0).Rows(f).Item("JobNo").ToString())
        Next f

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
            MessageBox.Show("Error with input string.", "Cannot convert to Rand format.", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
        End Try

    End Function
    Private Sub btn_Close_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Close.Click
        Close()
    End Sub

    Private Sub btn_Print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Print.Click
        If txtJobNum.Text = "" Then
            Exit Sub
        End If
        ReportType = "Job Rates"

        Try
            DocumentToPrint.DocumentName = "Job Rates : " + txtJobNum.Text
            Dim ppd_JCR As New PrintPreviewDialog
            ppd_JCR.WindowState = FormWindowState.Maximized
            ppd_JCR.Document = DocumentToPrint
            ppd_JCR.AutoScale = True
            ppd_JCR.AutoScroll = True
            ppd_JCR.UseAntiAlias = False
            ppd_JCR.PrintPreviewControl.Zoom = 1
            ppd_JCR.PrintPreviewControl.Columns = 1
            ppd_JCR.PrintPreviewControl.Rows = 1
            ppd_JCR.Text = "JOB RATES " + txtJobNum.Text
            curpagenum = 1
            PrintArray.Clear()
            All_Is_OK = True
            Dim selJob As String
            selJob = txtJobNum.Text
            JobPrint(selJob)
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
    End Sub

    Private Sub JobPrint(ByVal thisJob As String)

        Dim sql
        Dim x

        sql = "SELECT JobNo, TypeCode, Rate " & _
            "FROM JobRate " & _
            "WHERE JobNo = '" + thisJob + "'" & _
            "ORDER BY TypeCode"

        '"WHERE JobRate.JobNo = '290'"
        Dim DataSet = New Data.DataSet
        Dim curType As String
        Dim curRate As Double
        Dim adapter = New OleDb.OleDbDataAdapter(sql, DBConnection)
        adapter.Fill(DataSet)

        Dim recordCount = DataSet.Tables(0).Rows.Count
        If recordCount = "0" Then
            MessageBox.Show("No Job rates exist for Job No. " + thisJob)
            All_Is_OK = False
        Else
            field = New PageElement("JOB RATES FOR JOB ", EntryFont, 50, False, False)
            PrintArray.Add(field)

            field = New PageElement(thisJob, EntryFont, 200, True, False)
            PrintArray.Add(field)
            field = New PageElement("======================== ", EntryFont, 50, True, False, False)
            PrintArray.Add(field)
            field = New PageElement("Type             Rate", EntryFont, 50, True, False, False)
            PrintArray.Add(field)
            field = New PageElement("------", EntryFont, 50, False, False, False)
            PrintArray.Add(field)
            field = New PageElement("--------", EntryFont, 130, True, False, False)
            PrintArray.Add(field)
            ' LOOP THROUGH ITEM RECORDS
            ' HIGH TENSILE Y
            For x = 0 To recordCount - 1
                curType = DataSet.Tables(0).Rows(x).Item("TypeCode").ToString()
                field = New PageElement(curType, EntryFont, 50, False, False)
                PrintArray.Add(field)
                curRate = DataSet.Tables(0).Rows(x).Item("Rate")
                field = New PageElement(toRand(DataSet.Tables(0).Rows(x).Item("Rate").ToString(), True), EntryFont, 200, True, False, True)
                'field = New PageElement(toRand(curRate.ToString), EntryFont, 200, True, False, True)
                PrintArray.Add(field)
            Next

            '/* END OF PRINTING */


        End If

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
