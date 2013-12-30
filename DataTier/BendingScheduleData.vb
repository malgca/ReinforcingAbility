Imports System
Imports System.Globalization
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Common

''' <summary>
''' Bending Schedule database operations
''' </summary>
Public Class BendingScheduleData
    Public Property Adapter As New OleDbDataAdapter

    Private Property SelectJobsCommand As New OleDbCommand 'used
    Private Property ScheduleSummaryCommand As New OleDbCommand 'used
    Private Property CuttingSheetScheduleCommand As New OleDbCommand 'used
    Private Property JobScheduleCommand As New OleDbCommand 'used

    ''' <summary>
    ''' Populates a dataset with data from the Job table.
    ''' </summary>
    Public Sub PopulateJobsSet()
        Me.Adapter.SelectCommand = Me.SelectJobsCommand

        Me.SelectJobsCommand.CommandText = "SELECT JobNo FROM Job ORDER BY JobNo"

        Me.SelectJobsCommand.Connection = DBOperations.GetInstance.Connection
    End Sub

    ''' <summary>
    ''' Populates a schedule for a given job number.
    ''' </summary>
    Public Sub PopulateScheduleSummary(ByRef jobNumber As String)
        Me.Adapter.SelectCommand = Me.ScheduleSummaryCommand

        Me.ScheduleSummaryCommand.CommandText = "SELECT ContractorName, " & _
            "JobName, " & _
            "CompanyName, " & _
            "Job.[Tons or Kilograms] AS TKG " & _
            "FROM Job, Contractor, Company " & _
            "WHERE Job.ContractorNo = Contractor.ContractorNo " & _
            "AND Company.CompanyNo = Job.CompanyNo " & _
            "AND Job.JobNo = ?"

        Me.ScheduleSummaryCommand.Connection = DBOperations.GetInstance.Connection

        Dim newParam As New OleDbParameter()
        newParam.SourceColumn = "Job.JobNo"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 10
        newParam.Value = jobNumber
        Me.ScheduleSummaryCommand.Parameters.Add(newParam)
    End Sub

    ''' <summary>
    ''' Populate a complete schedule for a given schedule and date.
    ''' </summary>
    Public Sub PopulateJobSchedule(ByRef jobNumber As String, ByRef thisDate As Date)
        Me.Adapter.SelectCommand = Me.JobScheduleCommand

        Me.JobScheduleCommand.CommandText = "SELECT DISTINCT ScheduleNo, " & _
            "CuttingSheet.CutSheetNo " & _
            "FROM CuttingSheet INNER JOIN SchedItem ON CuttingSheet.CutSheetNo = SchedItem.CutSheetNo " & _
            "WHERE CutDate <= #" & thisDate.ToShortDateString() & "# " & _
            "AND InvoiceNo <> 0 " & _
            "AND [Job No] = '?'" & _
            "ORDER BY ScheduleNo"

        Me.JobScheduleCommand.Connection = DBOperations.GetInstance.Connection

        Dim newParam As New OleDbParameter()
        newParam.SourceColumn = "Job No"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 50
        newParam.Value = jobNumber
        Me.JobScheduleCommand.Parameters.Add(newParam)
    End Sub

    ''' <summary>
    ''' Populate summary of all cutting sheets and items per schedule
    ''' </summary>
    Public Sub PopulateCuttingSheetPerSchedule(ByRef scheduleNumber As String, ByRef cuttingSheet As String)
        Me.Adapter.SelectCommand = Me.CuttingSheetScheduleCommand

        Me.CuttingSheetScheduleCommand.CommandText = "SELECT CutItem.ScheduleNo, " & _
            "ProductType.TypeCode, " & _
            "CutItem.TypeCode, " & _
            "CutItem.Length, " & _
            "CutItem.Qty, " & _
            "ProductType.Weight " & _
            "FROM CutItem, ProductType " & _
            "WHERE CutItem.ScheduleNo = ? " & _
            "AND CutItem.CutSheetNo = ? " & _
            "AND CutItem.TypeCode = ProductType.TypeCode "

        Me.CuttingSheetScheduleCommand.Connection = DBOperations.GetInstance.Connection

        Dim newParam As New OleDbParameter()
        newParam.SourceColumn = "ScheduleNo"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 12
        newParam.Value = scheduleNumber
        Me.CuttingSheetScheduleCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "CutSheetNo"
        newParam.OleDbType = OleDbType.Numeric
        newParam.Value = cuttingSheet
        Me.CuttingSheetScheduleCommand.Parameters.Add(newParam)
    End Sub
End Class
