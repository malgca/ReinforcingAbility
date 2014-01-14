Imports DataTier
Imports System.Data
Imports System.Drawing
Imports System.ComponentModel

Public Class BendingSchedule
    Implements INotifyPropertyChanged

    ' settings and constants used to print pages
    Public Class PageConstants
        Public Const TopMargin As Integer = 60
        Public Const LeftMargin As Integer = 60
        Public Const RightMargin As Integer = 60
        Public Const BottomMargin As Integer = 90
        Public Const PageWidth As Integer = 873

        Public Const d1 As Integer = 85
        Public Const d2 As Integer = 75

        Public Shared ReportType As String
    End Class

    ' fonts used in printing pages
    Public Class PrintFonts
        ''' <summary>
        ''' Arial 10, Normal
        ''' </summary>
        Public Shared Normal As New Font("Arial", 10)

        ''' <summary>
        ''' Arial 8, Italic
        ''' </summary>
        Public Shared SmallItalic As New Font("Arial", 8, FontStyle.Italic)
    End Class

    ' number codes given to different beams
    Private Class BeamNumber
        Public Const Six As String = "06"
        Public Const Eight As String = "08"
        Public Const Ten As String = "10"
        Public Const Twelve As String = "12"
        Public Const Sixteen As String = "16"
        Public Const Twenty As String = "20"
        Public Const TwentyFive As String = "25"
        Public Const ThirtyTwo As String = "32"
        Public Const Forty As String = "40"
    End Class

    Public Property TKG As String = String.Empty
    Public Property CurrentListPosition As Integer
    Public Property CurrentPageNumber As Integer = 1

    Public Property PrintList As New List(Of PageElement)

    Public Property JobNameList As New List(Of String)

    Private Property BendingScheduleData As BendingScheduleData

    Public Sub New()
        BendingScheduleData = New BendingScheduleData()

        InitializeProperties(0)
    End Sub

    Public Sub InitializeProperties(ByVal index As Integer)
        PopulateJobList()
    End Sub

    ''' <summary>
    ''' Populates the existing list of jobs from the database.
    ''' </summary>
    Public Sub PopulateJobList()
        JobNameList.Clear()
        JobNameList.Add(String.Empty)
        BendingScheduleData.PopulateJobsSet()

        Dim JobNameSet As New DataSet
        BendingScheduleData.Adapter.Fill(JobNameSet)

        For Each newRow As DataRow In JobNameSet.Tables.Item(0).Rows
            If (IsNotNull(newRow("JobNo"))) Then
                JobNameList.Add(newRow("JobNo"))
            End If
        Next
    End Sub

    ' adds a schedule to the print queue
    Private Sub PrintSchedule(ByVal inSched As String, ByVal inType As String)
        PrintList.Add(New PageElement(inSched, PrintFonts.Normal, PageConstants.LeftMargin, True, False, False))
        PrintList.Add(New PageElement(inType, PrintFonts.Normal, PageConstants.LeftMargin + 40, False, False, False))
    End Sub

    ' adds a line to the print queue
    Private Sub PrintLine()
        PrintList.Add(New PageElement(String.Empty, PrintFonts.Normal, 0, True, False, False))
        PrintList.Add(New PageElement("<HR/BLACK>", PrintFonts.Normal, PageConstants.LeftMargin, True))
    End Sub

    ' adds a y head beam of a specific type to the print queue
    Private Sub PrintYHead(ByVal inType As String)
        PrintList.Add(New PageElement(String.Empty, PrintFonts.Normal, 0, True, False, False))
        PrintList.Add(New PageElement(inType, PrintFonts.Normal, PageConstants.LeftMargin + 40, False, False, False))
    End Sub

    ' adds an r head to the print queue
    Private Sub PrintRHead()
        PrintList.Add(New PageElement(String.Empty, PrintFonts.Normal, 0, True, False, False))
        PrintList.Add(New PageElement("Total", PrintFonts.Normal, PageConstants.LeftMargin, False, False, False))
        PrintList.Add(New PageElement(BeamTypes.R.ToString(), PrintFonts.Normal, PageConstants.LeftMargin + 40, False, False, False))
    End Sub

    ' adds the default y head beam to the print queue
    Private Sub PrintYHead()
        PrintList.Add(New PageElement(String.Empty, PrintFonts.Normal, 0, True, False, False))
        PrintList.Add(New PageElement(BeamTypes.Y.ToString(), PrintFonts.Normal, PageConstants.LeftMargin + 40, False, False, False))
    End Sub

    ' adds the value to be printed to the print queue
    Private Sub PrintValue(ByVal value As Double, ByVal element As Int16, ByVal TKG As String)
        Dim valueString As String

        If value <> 0 Then
            If TKG = "T" Then
                valueString = value.ToString("0.000")
            Else
                valueString = value.ToString("0.0")
            End If

            PrintList.Add(New PageElement(valueString, PrintFonts.Normal, PageConstants.PageWidth - ((8 - element) * PageConstants.d2) - 100, False, False, True))
        End If
    End Sub

    ' generates a summary of bending schedules for a given job and date
    Public Sub GenerateSummaryOfBendingSchedules(ByVal jobNo As String, ByVal aDate As Date)
        CreateBendingScheduleSummary(jobNo, aDate)

        Dim prevScheduleNumber = String.Empty

        ' totals
        Dim RTotals(8) As Double
        Dim YTotals(8) As Double

        ' schedules
        Dim RBeamsPerSchedule(8) As Double
        Dim YBeamsPerSchedule(8) As Double

        GetScheduleItems(jobNo, aDate, prevScheduleNumber, RBeamsPerSchedule, YBeamsPerSchedule, RTotals, YTotals)

        'PRINT SCHEDULE NUMBER
        PrintScheduleByNumber(prevScheduleNumber, RBeamsPerSchedule, YBeamsPerSchedule, RTotals, YTotals)
    End Sub

    ' creates a summary of the Bending schedules
    Private Sub CreateBendingScheduleSummary(ByVal jobNo As String, ByVal aDate As Date)
        PrintList.Clear()

        PrintList.Add(New PageElement("SUMMARY OF BENDING SCHEDULES", PrintFonts.Normal, 0, True, True, False))

        BendingScheduleData.PopulateScheduleSummary(jobNo)

        Dim ScheduleSummarySet As New DataSet
        BendingScheduleData.Adapter.Fill(ScheduleSummarySet)

        If ScheduleSummarySet.Tables(0).Rows.Count = 1 Then
            TKG = ScheduleSummarySet.Tables(0).Rows(0).Item("TKG").ToString()

            PrintList.Add(New PageElement(ScheduleSummarySet.Tables(0).Rows(0).Item("CompanyName").ToString(), PrintFonts.Normal, 0, True, True, False))

            PrintList.Add(New PageElement("Job Number:", PrintFonts.Normal, PageConstants.LeftMargin, False, False, False))
            PrintList.Add(New PageElement(jobNo, PrintFonts.Normal, PageConstants.LeftMargin + PageConstants.d1, True, False, False))

            PrintList.Add(New PageElement("Job Name:", PrintFonts.Normal, PageConstants.LeftMargin, False, False, False))
            PrintList.Add(New PageElement(ScheduleSummarySet.Tables(0).Rows(0).Item("JobName").ToString(), PrintFonts.Normal, PageConstants.LeftMargin + PageConstants.d1, True, False, False))

            PrintList.Add(New PageElement("Contractor:", PrintFonts.Normal, PageConstants.LeftMargin, False, False, False))
            PrintList.Add(New PageElement(ScheduleSummarySet.Tables(0).Rows(0).Item("ContractorName").ToString(), PrintFonts.Normal, PageConstants.LeftMargin + PageConstants.d1, True, False, False))

            PrintList.Add(New PageElement("Date:", PrintFonts.Normal, PageConstants.LeftMargin, False, False, False))
            PrintList.Add(New PageElement(aDate.ToShortDateString(), PrintFonts.Normal, PageConstants.LeftMargin + PageConstants.d1, True, False, False))

            PrintList.Add(New PageElement("<SPACE>", PrintFonts.Normal, PageConstants.LeftMargin, True, False, False))
            PrintList.Add(New PageElement("<HR/BLACK>", PrintFonts.Normal, PageConstants.LeftMargin, True))

            PrintList.Add(New PageElement("Schedule", PrintFonts.Normal, PageConstants.LeftMargin, False))

            PrintList.Add(New PageElement(BeamNumber.Six, PrintFonts.Normal, PageConstants.LeftMargin + 1 * PageConstants.d2, False))
            PrintList.Add(New PageElement(BeamNumber.Eight, PrintFonts.Normal, PageConstants.LeftMargin + 2 * PageConstants.d2, False))
            PrintList.Add(New PageElement(BeamNumber.Ten, PrintFonts.Normal, PageConstants.LeftMargin + 3 * PageConstants.d2, False))
            PrintList.Add(New PageElement(BeamNumber.Twelve, PrintFonts.Normal, PageConstants.LeftMargin + 4 * PageConstants.d2, False))
            PrintList.Add(New PageElement(BeamNumber.Sixteen, PrintFonts.Normal, PageConstants.LeftMargin + 5 * PageConstants.d2, False))
            PrintList.Add(New PageElement(BeamNumber.Twenty, PrintFonts.Normal, PageConstants.LeftMargin + 6 * PageConstants.d2, False))
            PrintList.Add(New PageElement(BeamNumber.TwentyFive, PrintFonts.Normal, PageConstants.LeftMargin + 7 * PageConstants.d2, False))
            PrintList.Add(New PageElement(BeamNumber.ThirtyTwo, PrintFonts.Normal, PageConstants.LeftMargin + 8 * PageConstants.d2, False))
            PrintList.Add(New PageElement(BeamNumber.Forty, PrintFonts.Normal, PageConstants.LeftMargin + 9 * PageConstants.d2, True, False, False))

            PrintList.Add(New PageElement("<HR/BLACK>", PrintFonts.Normal, PageConstants.LeftMargin, True))
        End If
    End Sub

    ' get the cutting sheets and the items for the schedule\
    Private Sub GetScheduleItems(ByVal jobNo As String, ByVal aDate As Date, ByRef prevScheduleNumber As String, ByRef RBeamsPerSchedule() As Double, ByRef YBeamsPerSchedule() As Double, ByRef RTotals() As Double, ByRef YTotals() As Double)
        ' GET ALL THE SCHEDULES AND CUTTING SHEETS FOR THAT JOB IN THAT DATE RANGE

        BendingScheduleData.PopulateJobSchedule(jobNo, aDate)

        Dim JobScheduleSet As New DataSet
        BendingScheduleData.Adapter.Fill(JobScheduleSet)

        ' /* FOR EACH SCHEDULE AND CUTTING SHEET*/
        For i As Integer = 0 To JobScheduleSet.Tables(0).Rows.Count - 1
            Dim scheduleNumber = JobScheduleSet.Tables(0).Rows(i).Item("ScheduleNo").ToString().ToUpper

            ChangeSchedules(scheduleNumber, prevScheduleNumber, RBeamsPerSchedule, YBeamsPerSchedule)

            ' ******************************
            '/*  GET ALL THE CUTTING SHEETS AND ITEMS FOR THE SCHEDULE */

            Dim cuttingSheetNumber = JobScheduleSet.Tables(0).Rows(i).Item("CutSheetNo")

            BendingScheduleData.PopulateCuttingSheetPerSchedule(scheduleNumber, cuttingSheetNumber)

            Dim CuttingSheetSet As New DataSet
            BendingScheduleData.Adapter.Fill(CuttingSheetSet)

            '/* IF THERE ARE ITEMS IN THE SCHEDULE */
            If CuttingSheetSet.Tables(0).Rows.Count <> 0 Then

                '/* LOOP THROUGH EACH ITEM */
                For r As Integer = 0 To CuttingSheetSet.Tables(0).Rows.Count - 1
                    Dim currentWeight As Double = CuttingSheetSet.Tables(0).Rows(r).Item("Weight")
                    Dim currentQuantity As Integer = CuttingSheetSet.Tables(0).Rows(r).Item("Qty")
                    Dim currentTypeCode As String = CuttingSheetSet.Tables(0).Rows(r).Item("CutItem.TypeCode").ToString()
                    Dim currentSteel As Double = CuttingSheetSet.Tables(0).Rows(r).Item("Length") * currentQuantity * currentWeight

                    If TKG = "T" Then
                        currentSteel = currentSteel / 1000000
                    Else
                        currentSteel = currentSteel / 1000
                    End If

                    ' ADD TO TOTAL FOR SCHEDULE
                    If (currentTypeCode.StartsWith("R")) Then
                        FillSchedules(RBeamsPerSchedule, RTotals, currentSteel, currentTypeCode)
                    Else
                        FillSchedules(YBeamsPerSchedule, YTotals, currentSteel, currentTypeCode)
                    End If
                Next ' next item in schedule
            End If   ' if this cutting sheet & schedule has items

            prevScheduleNumber = scheduleNumber
        Next
        '/* end of all schedules and cutting sheets for this job
    End Sub

    ' reprints schedule if any changes have happened.
    Private Sub ChangeSchedules(ByRef scheduleNumber As String, ByRef prevScheduleNumber As String, ByRef RBeamsPerSchedule() As Double, ByRef YBeamsPerSchedule() As Double)
        Dim scheduleChange = False

        ' see if schedule has changed
        If prevScheduleNumber <> scheduleNumber And prevScheduleNumber <> String.Empty Then
            scheduleChange = True
        Else
            scheduleChange = False
        End If

        ' see if schedule has changed '
        If scheduleChange Then
            'PRINT SCHEDULE NUMBER
            PrintSchedule(prevScheduleNumber, BeamTypes.R.ToString())

            '/* ROUND AND PRINT ALL Rs FOR THE SCHEDULE*/
            For f As Integer = 0 To 8
                PrintValue(RBeamsPerSchedule(f), f, TKG)
            Next

            PrintYHead(BeamTypes.Y.ToString())
            ' PRINT ALL Ys
            For f As Integer = 0 To 8
                PrintValue(YBeamsPerSchedule(f), f, TKG)
            Next

            'print line
            PrintLine()
            ' clear totals for next schedule */
            For j As Integer = 0 To 8
                RBeamsPerSchedule(j) = 0
                YBeamsPerSchedule(j) = 0
            Next
        End If  ' end of schedule change
    End Sub

    ' fills totals and schedules for given beam types
    Private Sub FillSchedules(ByRef BeamSchedules() As Double, ByRef BeamTotals() As Double, ByVal currentSteel As Double, ByVal currentTypeCode As String)
        Dim codeNumber As String = currentTypeCode.Substring(1)

        Select Case codeNumber
            Case BeamNumber.Six
                BeamSchedules(0) += currentSteel
                BeamTotals(0) += currentSteel
            Case BeamNumber.Eight
                BeamSchedules(1) += currentSteel
                BeamTotals(1) += currentSteel
            Case BeamNumber.Ten
                BeamSchedules(2) += currentSteel
                BeamTotals(2) += currentSteel
            Case BeamNumber.Twelve
                BeamSchedules(3) += currentSteel
                BeamTotals(3) += currentSteel
            Case BeamNumber.Sixteen
                BeamSchedules(4) += currentSteel
                BeamTotals(4) += currentSteel
            Case BeamNumber.Twenty
                BeamSchedules(5) += currentSteel
                BeamTotals(5) += currentSteel
            Case BeamNumber.TwentyFive
                BeamSchedules(6) += currentSteel
                BeamTotals(6) += currentSteel
            Case BeamNumber.ThirtyTwo
                BeamSchedules(7) += currentSteel
                BeamTotals(7) += currentSteel
            Case BeamNumber.Forty
                BeamSchedules(8) += currentSteel
                BeamTotals(8) += currentSteel
            Case Else
        End Select
    End Sub

    ' print a schedule according to it's schedule number
    Private Sub PrintScheduleByNumber(ByRef prevScheduleNumber As String, ByRef RBeamsPerSchedule() As Double, ByRef YBeamsPerSchedule() As Double, ByRef RTotals() As Double, ByRef YTotals() As Double)
        PrintSchedule(prevScheduleNumber, BeamTypes.R.ToString())

        '/* ROUND AND PRINT ALL Rs FOR THE SCHEDULE*/
        For i As Integer = 0 To 8
            PrintValue(RBeamsPerSchedule(i), i, TKG)
        Next

        PrintYHead(BeamTypes.Y.ToString())
        ' PRINT ALL Ys
        For i As Integer = 0 To 8
            PrintValue(YBeamsPerSchedule(i), i, TKG)
        Next

        PrintLine()
        PrintRHead()

        Dim totalR As Double = 0

        For i As Integer = 0 To 8
            PrintValue(RTotals(i), i, TKG)
            totalR += RTotals(i)
        Next i

        PrintYHead()

        Dim totalY As Double = 0

        For i As Integer = 0 To 8
            PrintValue(YTotals(i), i, TKG)
            totalY += YTotals(i)
        Next

        PrintList.Add(New PageElement(String.Empty, PrintFonts.Normal, 0, True, False, False))
        PrintList.Add(New PageElement(String.Empty, PrintFonts.Normal, 0, True, False, False))
        PrintList.Add(New PageElement("Total Mild Steel:", PrintFonts.Normal, PageConstants.LeftMargin, False, False, False))

        Dim totalValue As String

        If TKG = "T" Then
            totalValue = totalR.ToString("0.000")
        Else
            totalValue = totalR.ToString("0.0")
        End If

        PrintList.Add(New PageElement(totalValue & " " & TKG, PrintFonts.Normal, PageConstants.LeftMargin + 300, True, False, True))
        PrintList.Add(New PageElement("Total High Tensile Steel:", PrintFonts.Normal, PageConstants.LeftMargin, False, False, False))

        If TKG = "T" Then
            totalValue = totalY.ToString("0.000")
        Else
            totalValue = totalY.ToString("0.0")
        End If

        PrintList.Add(New PageElement(totalValue & " " & TKG, PrintFonts.Normal, PageConstants.LeftMargin + 300, True, False, True))
        PrintList.Add(New PageElement(String.Empty, PrintFonts.Normal, 0, True, False, False))
        PrintList.Add(New PageElement("Grand Total:", PrintFonts.Normal, PageConstants.LeftMargin, False, False, False))

        If TKG = "T" Then
            totalValue = (totalY + totalR).ToString("0.000")
        Else
            totalValue = (totalY + totalR).ToString("0.0")
        End If

        PrintList.Add(New PageElement(totalValue & " " & TKG, PrintFonts.Normal, PageConstants.LeftMargin + 300, True, False, True))
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
