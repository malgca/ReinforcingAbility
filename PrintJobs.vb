Public Class PrintJobs
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
    Friend WithEvents txtCoNum As System.Windows.Forms.ComboBox
    Friend WithEvents txtCoName As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.DocumentToPrint = New System.Drawing.Printing.PrintDocument
        Me.btn_Print = New System.Windows.Forms.Button
        Me.btn_Close = New System.Windows.Forms.Button
        Me.txtCoNum = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtCoName = New System.Windows.Forms.ComboBox
        Me.SuspendLayout()
        '
        'DocumentToPrint
        '
        '
        'btn_Print
        '
        Me.btn_Print.Location = New System.Drawing.Point(72, 80)
        Me.btn_Print.Name = "btn_Print"
        Me.btn_Print.Size = New System.Drawing.Size(96, 23)
        Me.btn_Print.TabIndex = 0
        Me.btn_Print.Text = "Print Preview..."
        '
        'btn_Close
        '
        Me.btn_Close.Location = New System.Drawing.Point(200, 80)
        Me.btn_Close.Name = "btn_Close"
        Me.btn_Close.Size = New System.Drawing.Size(96, 23)
        Me.btn_Close.TabIndex = 1
        Me.btn_Close.Text = "Close"
        '
        'txtCoNum
        '
        Me.txtCoNum.Location = New System.Drawing.Point(120, 24)
        Me.txtCoNum.Name = "txtCoNum"
        Me.txtCoNum.Size = New System.Drawing.Size(64, 21)
        Me.txtCoNum.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(32, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(72, 23)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Company"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtCoName
        '
        Me.txtCoName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.txtCoName.Location = New System.Drawing.Point(200, 24)
        Me.txtCoName.Name = "txtCoName"
        Me.txtCoName.Size = New System.Drawing.Size(232, 21)
        Me.txtCoName.TabIndex = 4
        '
        'PrintJobs
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(536, 142)
        Me.Controls.Add(Me.txtCoName)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtCoNum)
        Me.Controls.Add(Me.btn_Close)
        Me.Controls.Add(Me.btn_Print)
        Me.Name = "PrintJobs"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Print Job Listing"
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
    Dim curArrayPos As Integer = 0
    Dim curpagenum As Integer = 1
    Dim TopMargin As Integer = 60
    Dim LeftMargin As Integer = 60
    Dim RightMargin As Integer = 90
    Dim BottomMargin As Integer = 90
    Dim PageWidth As Integer = 873
    Dim ReportType As String

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
        populate_CoNumbers()
        populate_CoNames()
    End Sub

    Private Sub btn_Close_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Close.Click
        Close()
    End Sub

    Private Sub populate_CoNumbers()
        txtCoNum.Items.Clear()
        Dim sql As String = "SELECT Company.CompanyNo FROM Company"
        Dim ds As New Data.DataSet
        Dim ad As New OleDb.OleDbDataAdapter(sql, DBConnection)
        ad.Fill(ds)

        Dim f As Integer
        For f = 0 To ds.Tables(0).Rows.Count - 1
            txtCoNum.Items.Add(ds.Tables(0).Rows(f).Item("CompanyNo").ToString())
        Next f
    End Sub

    Private Sub populate_CoNames()
        txtCoName.Items.Clear()
        Dim sql2 As String = "SELECT Company.CompanyName FROM Company"
        Dim ds2 As New Data.DataSet
        Dim ad2 As New OleDb.OleDbDataAdapter(sql2, DBConnection)
        ad2.Fill(ds2)

        Dim f As Integer
        For f = 0 To ds2.Tables(0).Rows.Count - 1
            txtCoName.Items.Add(ds2.Tables(0).Rows(f).Item("CompanyName").ToString())
        Next f
    End Sub
    Private Sub btn_Print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Print.Click
        If txtCoNum.Text = "" Then
            Exit Sub
        End If
        ReportType = "Job Listing"

        Try
            DocumentToPrint.DocumentName = "Job Listing "
            Dim ppd_JCR As New PrintPreviewDialog
            ppd_JCR.WindowState = FormWindowState.Maximized
            ppd_JCR.Document = DocumentToPrint
            ppd_JCR.AutoScale = True
            ppd_JCR.AutoScroll = True
            ppd_JCR.UseAntiAlias = False
            ppd_JCR.PrintPreviewControl.Zoom = 1
            ppd_JCR.PrintPreviewControl.Columns = 1
            ppd_JCR.PrintPreviewControl.Rows = 1
            ppd_JCR.Text = "JOB LISTING " + txtCoNum.Text
            curpagenum = 1
            PrintArray.Clear()
            All_Is_OK = True
            Dim selCo As Integer
            Dim SelCoName As String
            selCo = txtCoNum.Text
            SelCoName = txtCoName.Text

            JobPrint(selCo, selCoName)
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

    Private Sub JobPrint(ByVal thisCo As String, ByVal thisCoName As String)
        Dim currDate As Date = Today
        Dim sql As String
        Dim x As Integer

        Dim jobName, jobnum, tonsOrKg, contName As String
        
        sql = "SELECT Job.JobNo, Job.JobName, Job.ContractorNo, " & _
              "Job.[Tons or Kilograms], Contractor.ContractorName " & _
            "FROM Job, Contractor " & _
            "WHERE Job.CompanyNo = '" + thisCo + "'" & _
            "AND Contractor.ContractorNo = Job.ContractorNo " & _
            "ORDER BY JobNo"

        Dim DataSet As Data.DataSet = New Data.DataSet

        Dim adapter As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(sql, DBConnection)
        Dim recordCount As Integer

        Try
            adapter.Fill(DataSet)

            recordCount = DataSet.Tables(0).Rows.Count
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        If recordCount = "0" Then
            MessageBox.Show("No Jobs exist for Company No. " + thisCo)
            All_Is_OK = False
        Else
            field = New PageElement("JOB LISTING FOR ", EntryFont, 50, False, False)
            PrintArray.Add(field)
            field = New PageElement(thisCoName, EntryFont, 170, False, False)

            PrintArray.Add(field)
            field = New PageElement(currDate.ToLongDateString, EntryFont, 580, True, False)
            PrintArray.Add(field)
            field = New PageElement("================================================ ", EntryFont, 50, True, False, False)
            PrintArray.Add(field)
            field = New PageElement(" ", EntryFont, 50, True, False, False)
            PrintArray.Add(field)


            field = New PageElement("Job No    Job Name", EntryFont, 50, False, False, False)
            PrintArray.Add(field)

            field = New PageElement("Tons/Kg   Contractor Name", EntryFont, 360, True, False, False)
            PrintArray.Add(field)
            field = New PageElement("-------------------------------------------------", EntryFont, 50, False, False, False)
            PrintArray.Add(field)
            field = New PageElement("------  ------------------------------------------------------", EntryFont, 360, True, False, False)
            PrintArray.Add(field)

            ' LOOP THROUGH RECORDS
            For x = 0 To recordCount - 1
                jobnum = DataSet.Tables(0).Rows(x).Item("JobNo").ToString()
                field = New PageElement(jobNum, EntryFont, 50, False, False, False)
                PrintArray.Add(field)
                jobName = DataSet.Tables(0).Rows(x).Item("JobName")

                field = New PageElement(jobName, EntryFont, 100, False, False, False)
                PrintArray.Add(field)
                tonsOrKg = DataSet.Tables(0).Rows(x).Item("Tons or Kilograms")
                field = New PageElement(tonsOrKg, EntryFont, 380, False, False, False)
                PrintArray.Add(field)
                contName = DataSet.Tables(0).Rows(x).Item("ContractorName")
                field = New PageElement(contName, EntryFont, 410, True, False, False)
                PrintArray.Add(field)
                'MessageBox.Show("in loop " + x.ToString + jobnum)
            Next

            '/* END OF PRINTING */


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


    Private Sub txtCoNum_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCoNum.SelectedIndexChanged
        txtCoName.SelectedIndex = txtCoNum.SelectedIndex
    End Sub

    Private Sub txtCoName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCoName.SelectedIndexChanged
        txtCoNum.SelectedIndex = txtCoName.SelectedIndex
    End Sub
End Class
