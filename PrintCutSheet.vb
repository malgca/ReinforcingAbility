Public Class PrintCutSheet
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
    Friend WithEvents txtCutNum As System.Windows.Forms.ComboBox
    Friend WithEvents btn_Close As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btn_Print As System.Windows.Forms.Button
    Friend WithEvents DocumentToPrint As System.Drawing.Printing.PrintDocument
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.DocumentToPrint = New System.Drawing.Printing.PrintDocument
        Me.btn_Print = New System.Windows.Forms.Button
        Me.btn_Close = New System.Windows.Forms.Button
        Me.txtCutNum = New System.Windows.Forms.ComboBox
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
        'txtCutNum
        '
        Me.txtCutNum.Location = New System.Drawing.Point(152, 24)
        Me.txtCutNum.Name = "txtCutNum"
        Me.txtCutNum.Size = New System.Drawing.Size(96, 21)
        Me.txtCutNum.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(32, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Cutting Sheet No:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'PrintCutSheet
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(292, 142)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtCutNum)
        Me.Controls.Add(Me.btn_Close)
        Me.Controls.Add(Me.btn_Print)
        Me.Name = "PrintCutSheet"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Print Cutting Sheet"
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

    Private Sub PrintCutSheet_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        populate_cutNumbers()
    End Sub

    Private Sub populate_cutNumbers()
        txtCutNum.Items.Clear()
        Dim sql = "SELECT CutSheetNo FROM CuttingSheet ORDER BY CutSheetNo"
        Dim ds As New Data.DataSet
        Dim ad As New OleDb.OleDbDataAdapter(sql, DBConnection)
        ad.Fill(ds)

        Dim f
        For f = 0 To ds.Tables(0).Rows.Count - 1
            txtCutNum.Items.Add(ds.Tables(0).Rows(f).Item("CutSheetNo").ToString())
        Next f

    End Sub

    Private Sub btn_Close_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Close.Click
        Close()
    End Sub

    Private Sub btn_Print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Print.Click
        If txtCutNum.Text = "" Then
            Exit Sub
        End If
        ReportType = "Cutting Sheet"

        Try
            DocumentToPrint.DocumentName = "Cutting Sheet No: " + txtCutNum.Text
            Dim ppd_JCR As New PrintPreviewDialog
            ppd_JCR.WindowState = FormWindowState.Maximized
            ppd_JCR.Document = DocumentToPrint
            ppd_JCR.AutoScale = True
            ppd_JCR.AutoScroll = True
            ppd_JCR.UseAntiAlias = False
            ppd_JCR.PrintPreviewControl.Zoom = 1
            ppd_JCR.PrintPreviewControl.Columns = 1
            ppd_JCR.PrintPreviewControl.Rows = 1
            ppd_JCR.Text = "CUTTING SHEET " + txtCutNum.Text
            curpagenum = 1
            PrintArray.Clear()
            All_Is_OK = True
            CutPrint(txtCutNum.Text)
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

    Private Sub CutPrint(ByVal cutNum As String)
        Const con1 = 495
        Const con3 = 70
        Dim sql, schedSql, itemHighSql, itemMildSql
        Dim x
        Dim kgTons, kgTonsDesc As String
        Dim nothingPrinted As Boolean
        Dim totalTensile, curQty, totalQty, curLength, nextLength, gTotTensile, lengthQty As Double
        Dim totalMild As Double = 0
        Dim metresMild As Double = 0
        Dim metresHigh As Double = 0

        Dim weightHigh, tonsMild, curWeight, totalMetres As Double
        gTotTensile = 0

        Dim firstSched, lastSched, curType, nextType, curSize, curTensile, nextTensile As String

        sql = "SELECT CSHeading, CutDate, CutSheetNo, Details, InvoiceNo, CuttingSheet.[Job No]," & _
            "Job.JobNo, Job.JobName, Job.OrderNo, Job.[Tons or Kilograms],Company.CompanyName, Contractor.ContractorName " & _
            "FROM CuttingSheet, Job, Company, Contractor   " & _
            "WHERE CuttingSheet.[Job No] = Job.JobNo " & _
            "AND Job.CompanyNo = Company.CompanyNo " & _
            "AND Job.ContractorNo = Contractor.ContractorNo " & _
           "AND CuttingSheet.CutSheetNo = " & cutNum

        Dim DataSet = New Data.DataSet
        Dim adapter = New OleDb.OleDbDataAdapter(sql, DBConnection)
        adapter.Fill(DataSet)

        If DataSet.tables(0).rows.count.ToString = "0" Then
            MessageBox.Show("No cutting sheet exists with Cutting Sheet No. " + cutNum)
            All_Is_OK = False
        Else
            kgTons = DataSet.Tables(0).Rows(0).Item("Tons or Kilograms").ToString()
            If kgTons = "T" Then
                kgTons = "Tons"
                kgTonsDesc = "T"
            Else
                kgTonsDesc = "Kg"
            End If
            field = New PageElement("CUTTING SHEET  ", EntryFont, 340, False, False)
            PrintArray.Add(field)
            field = New PageElement(cutNum, EntryFont, 455, True, False)
            PrintArray.Add(field)
            field = New PageElement(LeftMargin + 242, PageWidth - RightMargin - 275, True)
            PrintArray.Add(field)
            field = New PageElement(" ", EntryFont, 340, True, False)
            PrintArray.Add(field)
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("CompanyName").ToString(), EntryFont, 350, True, True)
            PrintArray.Add(field)
            field = New PageElement(" ", EntryFont, 340, True, False)
            PrintArray.Add(field)
            field = New PageElement("Job Name :", EntryFont, LeftMargin, False, False)
            PrintArray.Add(field)
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("JobName").ToString(), EntryFont, 185, False, False)
            PrintArray.Add(field)
            'GET FIRST SCHEDULE NUMBER OF CUTTING SHEET

            schedSql = "SELECT * FROM SchedItem " & _
                              "WHERE SchedItem.CutSheetNo = " & cutNum


            Dim schedAdapter As New OleDb.OleDbDataAdapter(schedSql, DBConnection)
            Dim schedDataSet = New Data.DataSet
            schedAdapter.Fill(schedDataSet)
            lastSched = ""
            firstSched = ""

            Dim schedRecCount = schedDataSet.Tables(0).Rows.Count
            If schedRecCount = 0 Then
                MessageBox.Show("No schedules found for this cutting sheet", "Warning")
            Else
                lastSched = schedDataSet.Tables(0).Rows(schedDataSet.Tables(0).Rows.Count - 1).Item("ScheduleNo").ToString
                firstSched = schedDataSet.Tables(0).Rows(0).Item("ScheduleNo").ToString()
            End If

            field = New PageElement("Schedule :", EntryFont, 600, False, False)
            PrintArray.Add(field)
            field = New PageElement(firstSched, EntryFont, 690, True, False)
            PrintArray.Add(field)
            field = New PageElement("Contractor :", EntryFont, LeftMargin, False, False)
            PrintArray.Add(field)
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("ContractorName").ToString(), EntryFont, 185, False, False)
            PrintArray.Add(field)
            field = New PageElement("To Schedule :", EntryFont, 600, False, False)
            PrintArray.Add(field)
            field = New PageElement(lastSched, EntryFont, 690, True, False)
            PrintArray.Add(field)
            field = New PageElement("Details :", EntryFont, LeftMargin, False, False)
            PrintArray.Add(field)
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("Details").ToString(), EntryFont, 185, False, False)
            PrintArray.Add(field)
            field = New PageElement("Job Number :", EntryFont, 600, False, False)
            PrintArray.Add(field)
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("JobNo").ToString(), EntryFont, 690, True, False)
            PrintArray.Add(field)
            field = New PageElement("Order Number :", EntryFont, LeftMargin, False, False)
            PrintArray.Add(field)
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("OrderNo").ToString(), EntryFont, 185, False, False)
            PrintArray.Add(field)
            field = New PageElement("Date :", EntryFont, 600, False, False)
            PrintArray.Add(field)
            field = New PageElement(DataSet.Tables(0).Rows(0).Item("CutDate").ToShortDateString(), EntryFont, 690, True, False)
            PrintArray.Add(field)
            field = New PageElement("<HR/BLACK>", EntryFont, LeftMargin, -15, False)
            PrintArray.Add(field)

            ' GET DETAIL LINES FOR CUTTING SHEET
            itemHighSql = "SELECT CutItem.TypeCode, Length, Qty, Weight FROM CutItem, ProductType, ProdCat " & _
                      "WHERE CutItem.CutSheetNo = " & cutNum & _
                      " AND ProductType.TypeCode = CutItem.TypeCode " & _
                        " AND ProdCat.CatCode = ProductType.CatCode" & _
                        " ORDER BY CutItem.TypeCode, CutItem.Length"


            Dim itemAdapter As New OleDb.OleDbDataAdapter(itemHighSql, DBConnection)
            Dim itemDSet = New Data.DataSet
            itemAdapter.Fill(itemDSet)
            Dim recordCount = itemDSet.Tables(0).Rows.Count
            If recordCount = 0 Then
                MessageBox.Show("There are no items for this cutting sheet", "Warning")
            End If

            field = New PageElement("Quantity       Type          Length", EntryFont, LeftMargin, False, False)
            PrintArray.Add(field)
            field = New PageElement("Totals                                    Weight Summary", EntryFont, 395, True, False)
            PrintArray.Add(field)

            field = New PageElement("Metres       Kg/M", EntryFont, 320, False, False)
            PrintArray.Add(field)
            field = New PageElement(kgTons, EntryFont, 470, True, False)
            PrintArray.Add(field)
            field = New PageElement("Type     Weight", EntryFont, 585, True, False)
            PrintArray.Add(field)
            'PRINT LINE
            field = New PageElement("<HR/BLACK>", EntryFont, LeftMargin, -15, False)
            PrintArray.Add(field)

            nextTensile = ""
            curTensile = ""
            nextLength = 0

            ' LOOP THROUGH ITEM RECORDS
            ' HIGH TENSILE Y
            For x = 0 To recordCount - 1

                curType = itemDSet.Tables(0).Rows(x).Item("TypeCode")
                curSize = curType.Substring(1)
                curTensile = curType.Substring(0, 1)
                curQty = itemDSet.Tables(0).Rows(x).Item("Qty")
                curLength = itemDSet.Tables(0).Rows(x).Item("Length")
                curWeight = itemDSet.Tables(0).Rows(x).Item("Weight")
                lengthQty += curQty
                nothingPrinted = True

                'CHANGE IN TENSILE IE Y TO R
                If nextTensile <> curTensile Then
                    field = New PageElement(curTensile, EntryFont, 595, True, False)
                    PrintArray.Add(field)
                    field = New PageElement(LeftMargin + con1, PageWidth - RightMargin - 160, True)
                    PrintArray.Add(field)
                    nothingPrinted = False
                End If

                If x + 1 < recordCount Then
                    nextType = itemDSet.Tables(0).Rows(x + 1).Item("TypeCode")
                    nextTensile = nextType.Substring(0, 1)
                    nextLength = itemDSet.Tables(0).Rows(x + 1).Item("Length")
                Else
                    nextTensile = ""
                    nextType = ""
                    nextLength = 0
                End If

                If (nextLength <> curLength) Or curType <> nextType Then
                    ' PRINT TOTALS FOR THIS TYPE AND LENGTH
                    field = New PageElement(lengthQty.ToString(), EntryFont, 135, False, False, True)
                    PrintArray.Add(field)
                    field = New PageElement(curType.ToString(), EntryFont, 160, False, False)
                    PrintArray.Add(field)
                    field = New PageElement(curLength.ToString(), EntryFont, 250, False, False, True)
                    PrintArray.Add(field)
                    metresHigh = lengthQty * curLength
                    metresHigh = Math.Round(metresHigh / 1000, 3)
                    'ADD TO TOTAL
                    totalMetres += Double.Parse(metresHigh.ToString)
                    lengthQty = 0
                    nothingPrinted = False

                End If


                'PRODUCT TYPE CHANGE
                If nextType <> curType Or x = recordCount - 1 Then
                    field = New PageElement(totalMetres.ToString(), EntryFont, 315, False, False, True)
                    PrintArray.Add(field)
                    field = New PageElement("@", EntryFont, 340, False, False)
                    PrintArray.Add(field)
                    'If kgTons = "Tons" Then
                    field = New PageElement(curWeight.ToString("0.000"), EntryFont, 400, False, False, True)
                    'Else
                    'field = New PageElement(curWeight.ToString("0.0"), EntryFont, 400, False, False, True)
                    'End If
                PrintArray.Add(field)
                field = New PageElement("=", EntryFont, 420, False, False)
                PrintArray.Add(field)

                weightHigh = (totalMetres * curWeight)


                If kgTons = "Tons" Then
                    weightHigh = Math.Round(weightHigh / 1000, 3)
                    field = New PageElement(weightHigh.ToString("0.000"), EntryFont, 490, False, False, True)
                Else
                    'If Kilograms then 1 decimal 
                    weightHigh = Math.Round(weightHigh, 1)
                    field = New PageElement(weightHigh.ToString("0.0"), EntryFont, 490, False, False, True)
                End If

                PrintArray.Add(field)
                field = New PageElement(curSize.ToString(), EntryFont, 590, False, False)
                PrintArray.Add(field)
                If kgTons = "Tons" Then
                    field = New PageElement(weightHigh.ToString("0.000"), EntryFont, 680, True, False, True)
                Else
                    field = New PageElement(weightHigh.ToString("0.0"), EntryFont, 680, True, False, True)
                End If
                PrintArray.Add(field)
                totalTensile += weightHigh

                totalMetres = 0
                totalQty = 0
                nothingPrinted = False
                End If

                    'CHANGE IN TENSILE IE Y TO R
                    If nextTensile <> curTensile Or x = recordCount - 1 Then
                        field = New PageElement(LeftMargin + con1, PageWidth - RightMargin - con3, True)
                        PrintArray.Add(field)
                        field = New PageElement("TOTAL " & curTensile + ":", EntryFont, 530, False, False)
                    PrintArray.Add(field)
                    If kgTons = "Tons" Then
                        field = New PageElement(totalTensile.ToString("0.000"), EntryFont, 680, False, False, True)
                    Else
                        field = New PageElement(totalTensile.ToString("0.0"), EntryFont, 680, False, False, True)
                    End If
                    PrintArray.Add(field)

                    field = New PageElement(kgTonsDesc, EntryFont, 710, True, False)
                    PrintArray.Add(field)
                    field = New PageElement(LeftMargin + con1, PageWidth - RightMargin - con3, True)
                    PrintArray.Add(field)
                    gTotTensile += totalTensile
                    totalTensile = 0
                    'PRINT NEXT TENSILE HEADING
                    If nextTensile <> "" Then
                        field = New PageElement("", EntryFont, 710, True, False)
                        PrintArray.Add(field)
                        field = New PageElement(nextTensile, EntryFont, 595, True, False)
                        PrintArray.Add(field)
                        field = New PageElement(LeftMargin + con1, PageWidth - RightMargin - 160, True)
                        PrintArray.Add(field)
                    End If
                    nothingPrinted = False
                End If

                    If Not nothingPrinted Then
                        ' PRINT BLANK LINE
                        field = New PageElement("", EntryFont, 700, True, False)
                        PrintArray.Add(field)
                    End If
            Next

            field = New PageElement(LeftMargin + con1, PageWidth - RightMargin - con3, True)
            PrintArray.Add(field)
            field = New PageElement("GRAND TOTAL :", EntryFont, 485, False, False)
            PrintArray.Add(field)
            If kgTons = "Tons" Then
                field = New PageElement(gTotTensile.ToString("0.000"), EntryFont, 680, False, False, True)
            Else
                field = New PageElement(gTotTensile.ToString("0.0"), EntryFont, 680, False, False, True)
            End If
            PrintArray.Add(field)
            field = New PageElement(kgTonsDesc, EntryFont, 710, True, False)
            PrintArray.Add(field)
            field = New PageElement(True, LeftMargin + con1, PageWidth - RightMargin - con3)
            PrintArray.Add(field)

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


    Private Sub txtCutNum_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCutNum.SelectedIndexChanged

    End Sub
End Class
