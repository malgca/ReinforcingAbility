Imports DataTier
Imports System.Data
Imports System.Drawing
Imports System.ComponentModel

Public Class BendingSchedule
    Implements INotifyPropertyChanged

    Private Class PageConstants
        Public Const Six As String = "06"
        Public Const Eight As String = "08"
        Public Const Ten As String = "10"
        Public Const Twelve As String = "12"
        Public Const Sixteen As String = "16"
        Public Const Twenty As String = "20"
        Public Const TwentyFive As String = "25"
        Public Const ThirtyTwo As String = "32"
        Public Const Forty As String = "40"

        Public Const LeftMargin As Integer = 60
        Public Const PageWidth As Integer = 873
        Public Const d2 As Integer = 75
    End Class

    Dim DBConnection As New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source = winsteelVers5.mdb")
    Dim Font As New Font("Arial", 10)

    Public Property JobNumber As String
    Public Property JobNameList As New List(Of String)
    Dim PrintList As New List(Of PageElement)

    Private Property BendingScheduleData As BendingScheduleData
    Private Property BendingScheduleSet As New DataSet

    Public Sub New()
        BendingScheduleData = New BendingScheduleData(JobNumber)

        InitializeProperties(0)
    End Sub

    Public Sub InitializeProperties(ByVal index As Integer)

        PopulateJobList()
    End Sub

    Public Sub PopulateJobList()
        JobNameList.Clear()
        BendingScheduleData.PopulateJobsSet()

        BendingScheduleSet.Clear()
        BendingScheduleData.Adapter.Fill(BendingScheduleSet)

        For Each newRow As DataRow In BendingScheduleSet.Tables.Item(0).Rows
            If (IsNotNull(newRow("JobNo"))) Then
                JobNameList.Add(newRow("JobNo"))
            End If
        Next
    End Sub

    Private Sub PrintSchedule(ByVal inSched As String, ByVal inType As String)
        PrintList.Add(New PageElement(inSched, Font, PageConstants.LeftMargin, True, False, False))
        PrintList.Add(New PageElement(inType, Font, PageConstants.LeftMargin + 40, False, False, False))
    End Sub

    Private Sub PrintLine()
        PrintList.Add(New PageElement("", Font, 0, True, False, False))
        PrintList.Add(New PageElement("<HR/BLACK>", Font, PageConstants.LeftMargin, True))
    End Sub

    Private Sub PrintYHead(ByVal inType As String)
        PrintList.Add(New PageElement("", Font, 0, True, False, False))
        PrintList.Add(New PageElement(inType, Font, PageConstants.LeftMargin + 40, False, False, False))
    End Sub

    Private Sub PrintTypeHeadings()
        PrintList.Add(New PageElement("Schedule", Font, PageConstants.LeftMargin, False))

        PrintList.Add(New PageElement(PageConstants.Six, Font, PageConstants.LeftMargin + 1 * PageConstants.d2, False))
        PrintList.Add(New PageElement(PageConstants.Eight, Font, PageConstants.LeftMargin + 2 * PageConstants.d2, False))
        PrintList.Add(New PageElement(PageConstants.Ten, Font, PageConstants.LeftMargin + 3 * PageConstants.d2, False))
        PrintList.Add(New PageElement(PageConstants.Twelve, Font, PageConstants.LeftMargin + 4 * PageConstants.d2, False))
        PrintList.Add(New PageElement(PageConstants.Sixteen, Font, PageConstants.LeftMargin + 5 * PageConstants.d2, False))
        PrintList.Add(New PageElement(PageConstants.Twenty, Font, PageConstants.LeftMargin + 6 * PageConstants.d2, False))
        PrintList.Add(New PageElement(PageConstants.TwentyFive, Font, PageConstants.LeftMargin + 7 * PageConstants.d2, False))
        PrintList.Add(New PageElement(PageConstants.ThirtyTwo, Font, PageConstants.LeftMargin + 8 * PageConstants.d2, False))
        PrintList.Add(New PageElement(PageConstants.Forty, Font, PageConstants.LeftMargin + 9 * PageConstants.d2, True, False, False))

        PrintList.Add(New PageElement("<HR/BLACK>", Font, PageConstants.LeftMargin, True))
    End Sub

    Private Sub PrintRHead()
        PrintList.Add(New PageElement("", Font, 0, True, False, False))
        PrintList.Add(New PageElement("Total", Font, PageConstants.LeftMargin, False, False, False))
        PrintList.Add(New PageElement("R", Font, PageConstants.LeftMargin + 40, False, False, False))
    End Sub

    Private Sub PrintYHead()
        PrintList.Add(New PageElement("", Font, 0, True, False, False))
        PrintList.Add(New PageElement("Y", Font, PageConstants.LeftMargin + 40, False, False, False))
    End Sub

    Private Sub PrintValue(ByVal inVal As Double, ByVal inElem As Int16, ByVal inTKg As String)
        Dim vo As String
        If inVal <> 0 Then
            If inTKg = "T" Then
                vo = inVal.ToString("0.000")
            Else
                vo = inVal.ToString("0.0")
            End If
            PrintList.Add(New PageElement(vo, Font, PageConstants.PageWidth - ((8 - inElem) * PageConstants.d2) - 100, False, False, True))
        End If
    End Sub

    Private Sub GenerateSummaryOfBendingSchedules(ByVal jobNo As String, ByVal aDate As Date)
        PrintList = New List(Of PageElement)
        Dim newPage As PageElement
        Dim TKg As String = String.Empty

        newPage = New PageElement("SUMMARY OF BENDING SCHEDULES", Font, 0, True, True, False)
        PrintList.Add(newPage)

        BendingScheduleData.PopulateScheduleSummary(JobNumber)

        BendingScheduleSet.Clear()
        BendingScheduleData.Adapter.Fill(BendingScheduleSet)

        If BendingScheduleSet.Tables(0).Rows.Count = 1 Then
            Const d1 As Integer = 85

            PrintList.Add(New PageElement(ds.Tables(0).Rows(0).Item("CompanyName").ToString(), EntryFont, 0, True, True, False))
            PrintList.Add(New PageElement("Job Number:", EntryFont, LeftMargin, False, False, False))
            PrintList.Add(New PageElement(jobNo, EntryFont, LeftMargin + d1, True, False, False))
            currJobName = BendingScheduleSet.Tables(0).Rows(0).Item("JobName").ToString()
            TKg = BendingScheduleSet.Tables(0).Rows(0).Item("TKG").ToString()
            PrintList.Add(New PageElement("Job Name:", EntryFont, LeftMargin, False, False, False))
            PrintList.Add(New PageElement(currJobName, EntryFont, LeftMargin + d1, True, False, False))
            PrintList.Add(New PageElement("Contractor:", EntryFont, LeftMargin, False, False, False))
            PrintList.Add(New PageElement(ds.Tables(0).Rows(0).Item("ContractorName").ToString(), EntryFont, LeftMargin + d1, True, False, False))
            PrintList.Add(New PageElement("Date:", EntryFont, LeftMargin, False, False, False))
            PrintList.Add(New PageElement(aDate.ToShortDateString, EntryFont, LeftMargin + d1, True, False, False))
            PrintList.Add(New PageElement("<SPACE>", EntryFont, LeftMargin, True, False, False))
            PrintList.Add(New PageElement("<HR/BLACK>", EntryFont, LeftMargin, True))
            PrintTypeHeadings()
        End If

        Dim rowCnt, itemCnt As Int16
        Dim RperSched(8) As Double
        Dim YperSched(8) As Double
        Dim RTotals(8) As Double
        Dim YTotals(8) As Double
        Dim schedNo, prevSched, cutNo As String
        Dim hasItems, schedChange As Boolean
        Dim i, f, r, j As Integer
        Dim typeR, typeY, curTC As String
        typeR = "R"
        typeY = "Y"
        Dim TR, TY, curWeight, curQty, curSteel As Double
        TR = 0
        TY = 0
        For i = 0 To 8
            RTotals(i) = 0
            YTotals(i) = 0
        Next i
        schedChange = False
        prevSched = ""
        curSteel = 0
        ' GET ALL THE SCHEDULES AND CUTTING SHEETS FOR THAT JOB IN THAT DATE RANGE
        Dim sql4ScheduleNos As String = "SELECT DISTINCT ScheduleNo, CuttingSheet.CutSheetNo" & _
        " FROM CuttingSheet INNER JOIN SchedItem ON CuttingSheet.CutSheetNo = SchedItem.CutSheetNo " & _
        "WHERE CutDate <= #" & aDate.ToShortDateString & "# AND InvoiceNo <> 0 AND [Job No] = '" & jobNo & "'" & _
        "ORDER BY ScheduleNo"

        Dim DS4SchNo As Data.DataSet = New Data.DataSet
        Dim adapter As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(sql4ScheduleNos, DBConnection)
        adapter.Fill(DS4SchNo)

        rowCnt = DS4SchNo.Tables(0).Rows.Count
        ' /* FOR EACH SCHEDULE AND CUTTING SHEET*/
        For i = 0 To rowCnt - 1
            schedNo = DS4SchNo.Tables(0).Rows(i).Item("ScheduleNo").ToString().ToUpper
            cutNo = DS4SchNo.Tables(0).Rows(i).Item("CutSheetNo").ToString()

            ' see if schedule has changed
            If prevSched <> schedNo And prevSched <> "" Then
                schedChange = True
            Else
                schedChange = False
            End If

            ' see if schedule has changed '
            If schedChange Then
                'PRINT SCHEDULE NUMBER
                PrintSchedule(prevSched, typeR)

                '/* ROUND AND PRINT ALL Rs FOR THE SCHEDULE*/
                For f = 0 To 8
                    PrintValue(RperSched(f), f, TKg)
                Next f

                PrintYHead(typeY)
                ' PRINT ALL Ys
                For f = 0 To 8
                    PrintValue(YperSched(f), f, TKg)
                Next f

                'print line
                PrintLine()
                ' clear totals for next schedule */
                For j = 0 To 8
                    RperSched(j) = 0
                    YperSched(j) = 0
                Next j
            End If  ' end of schedule change

            '/*  GET ALL THE CUTTING SHEETS AND ITEMS FOR THE SCHEDULE */
            Dim sqlPerSchR As String = "SELECT CutItem.ScheduleNo, ProductType.TypeCode, CutItem.TypeCode, CutItem.Length, CutItem.Qty, ProductType.Weight " & _
                        "FROM CutItem, ProductType " & _
                       "WHERE CutItem.ScheduleNo = '" & schedNo & "'" & _
                       "AND CutItem.CutSheetNo = " & cutNo & _
                        "AND CutItem.TypeCode = ProductType.TypeCode "

            Dim ds4r As New Data.DataSet
            Dim ad4R As New OleDb.OleDbDataAdapter(sqlPerSchR, DBConnection)
            ad4R.Fill(ds4r)
            hasItems = False
            itemCnt = ds4r.Tables(0).Rows.Count
            '/* IF THERE ARE ITEMS IN THE SCHEDULE */
            If itemCnt <> 0 Then
                hasItems = True
            End If

            If hasItems Then
                '/* LOOP THROUGH EACH ITEM */
                For r = 0 To itemCnt - 1
                    curWeight = ds4r.Tables(0).Rows(r).Item("Weight")
                    curQty = ds4r.Tables(0).Rows(r).Item("Qty")
                    curTC = ds4r.Tables(0).Rows(r).Item("CutItem.TypeCode").ToString()
                    curSteel = ds4r.Tables(0).Rows(r).Item("Length") * curQty * curWeight

                    If TKg = "T" Then
                        curSteel = curSteel / 1000000
                    Else
                        curSteel = curSteel / 1000
                    End If

                    'MessageBox.Show("Add to Total " + "curTc " + curTC + " " + curSteel.ToString + " " + TKg)
                    ' ADD TO TOTAL FOR SCHEDULE
                    If curTC = "R06" Then
                        RperSched(0) += curSteel
                        RTotals(0) += curSteel
                    ElseIf curTC = "R08" Then
                        RperSched(1) += curSteel
                        RTotals(1) += curSteel
                    ElseIf curTC = "R10" Then
                        RperSched(2) += curSteel
                        RTotals(2) += curSteel
                    ElseIf curTC = "R12" Then
                        RperSched(3) += curSteel
                        RTotals(3) += curSteel
                    ElseIf curTC = "R16" Then
                        RperSched(4) += curSteel
                        RTotals(4) += curSteel
                    ElseIf curTC = "R20" Then
                        RperSched(5) += curSteel
                        RTotals(5) += curSteel
                    ElseIf curTC = "R25" Then
                        RperSched(6) += curSteel
                        RTotals(6) += curSteel
                    ElseIf curTC = "R32" Then
                        RperSched(7) += curSteel
                        RTotals(7) += curSteel
                    ElseIf curTC = "R40" Then
                        RperSched(8) += curSteel
                        RTotals(8) += curSteel
                    End If

                    '/* CHECK Y TYPES
                    If curTC = "Y06" Then
                        YperSched(0) += curSteel
                        YTotals(0) += curSteel
                    ElseIf curTC = "Y08" Then
                        YperSched(1) += curSteel
                        YTotals(1) += curSteel
                    ElseIf curTC = "Y10" Then
                        YperSched(2) += curSteel
                        YTotals(2) += curSteel
                    ElseIf curTC = "Y12" Then
                        YperSched(3) += curSteel
                        YTotals(3) += curSteel
                    ElseIf curTC = "Y16" Then
                        YperSched(4) += curSteel
                        YTotals(4) += curSteel
                    ElseIf curTC = "Y20" Then
                        YperSched(5) += curSteel
                        YTotals(5) += curSteel
                    ElseIf curTC = "Y25" Then
                        YperSched(6) += curSteel
                        YTotals(6) += curSteel
                    ElseIf curTC = "Y32" Then
                        YperSched(7) += curSteel
                        YTotals(7) += curSteel
                    ElseIf curTC = "Y40" Then
                        YperSched(8) += curSteel
                        YTotals(8) += curSteel
                    End If
                Next r  ' next item in schedule
            End If   ' if this cutting sheet & schedule has items

            prevSched = schedNo
        Next i
        '/* end of all schedules and cutting sheets for this job

        'PRINT SCHEDULE NUMBER
        PrintSchedule(prevSched, typeR)

        '/* ROUND AND PRINT ALL Rs FOR THE SCHEDULE*/
        For f = 0 To 8
            PrintValue(RperSched(f), f, TKg)
        Next f

        PrintYHead(typeY)
        ' PRINT ALL Ys
        For f = 0 To 8
            PrintValue(YperSched(f), f, TKg)
        Next f

        'print line
        PrintLine()
        PrintRHead()

        Dim ci As Integer
        For ci = 0 To 8
            PrintValue(RTotals(ci), ci, TKg)
            TR += RTotals(ci)
        Next ci

        PrintYHead()

        For ci = 0 To 8
            PrintValue(YTotals(ci), ci, TKg)
            TY += YTotals(ci)
        Next ci
        PrintList.Add(New PageElement("", EntryFont, 0, True, False, False))
        PrintList.Add(New PageElement("", EntryFont, 0, True, False, False))
        PrintList.Add(New PageElement("Total Mild Steel:", EntryFont, LeftMargin, False, False, False))
        Dim v As String
        If TKg = "T" Then
            v = TR.ToString("0.000")
        Else
            v = TR.ToString("0.0")
        End If
        PrintList.Add(New PageElement(v & " " & TKg, EntryFont, LeftMargin + 300, True, False, True))

        PrintList.Add(New PageElement("Total High Tensile Steel:", EntryFont, LeftMargin, False, False, False))
        If TKg = "T" Then
            v = TY.ToString("0.000")
        Else
            v = TY.ToString("0.0")
        End If
        PrintList.Add(New PageElement(v & " " & TKg, EntryFont, LeftMargin + 300, True, False, True))
        PrintList.Add(New PageElement("", EntryFont, 0, True, False, False))
        PrintList.Add(New PageElement("Grand Total:", EntryFont, LeftMargin, False, False, False))
        If TKg = "T" Then
            v = (TY + TR).ToString("0.000")
        Else
            v = (TY + TR).ToString("0.0")
        End If
        PrintList.Add(New PageElement(v & " " & TKg, EntryFont, LeftMargin + 300, True, False, True))

    End Sub

    Public Event PropertyChanged(sender As Object, e As PropertyChangedEventArgs) Implements INotifyPropertyChanged.PropertyChanged

    ''' <summary>
    ''' Notifies listener of a change in a property
    ''' </summary>
    Private Sub NotifyPropertyChanged(ByVal ParamArray Properties() As String)
        For Each Prop As String In Properties
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(Prop))
        Next
    End Sub
    ' Checks if a DB field is not equal to null
    Private Function IsNotNull(ByRef param As Object) As Boolean
        Return Not param.Equals(DBNull.Value)
    End Function
End Class
