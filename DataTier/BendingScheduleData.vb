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

    Private Property SelectJobsCommand As New OleDbCommand
    Private Property ScheduleSummaryCommand As New OleDbCommand
    Private Property DateRangeCommand As New OleDbCommand
    Private Property FullScheduleCommand As New OleDbCommand

    Public Sub New(ByRef jobNumber As String)
    
    End Sub

    ''' <summary>
    ''' Populates a dataset with data from the Job table.
    ''' </summary>
    Public Sub PopulateJobsSet()
        Me.Adapter.SelectCommand = Me.SelectJobsCommand

        Me.SelectJobsCommand.CommandText = "SELECT JobNo FROM Job ORDER BY JobNo"

        Me.SelectJobsCommand.Connection = DBOperations.GetInstance.Connection
    End Sub

    ''' <summary>
    ''' Populates a schedule for a given 
    ''' </summary>
    Public Sub PopulateScheduleSummary(ByRef jobNumber As String)
        Me.Adapter.SelectCommand = Me.ScheduleSummaryCommand

        Me.ScheduleSummaryCommand.CommandText = "SELECT ContractorName, " & _
            "JobName, " & _
            "CompanyName, " & _
            "job.[Tons or Kilograms] AS TKG " & _
            "FROM Job, Contractor,Company " & _
            "WHERE Job.ContractorNo = Contractor.ContractorNo " & _
            "AND Company.CompanyNo = Job.CompanyNo " & _
            "AND Job.JobNo = ?"

        Me.SelectJobsCommand.Connection = DBOperations.GetInstance.Connection

        Dim newParam As New OleDbParameter()
        newParam.SourceColumn = "JobNo"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 10
        newParam.Value = jobNumber
        Me.ScheduleSummaryCommand.Parameters.Add(newParam)
    End Sub
End Class
