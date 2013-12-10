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
    Public Sub PopulateJobsSet(ByRef dataSet As DataSet)
        Me.Adapter.SelectCommand = Me.SelectJobsCommand

        ' SelectCommand
        Me.SelectJobsCommand.CommandText = "SELECT JobNo FROM Job ORDER BY JobNo"

        Me.SelectJobsCommand.Connection = DBOperations.GetInstance.Connection

        Me.Adapter.Fill(dataSet)
    End Sub
End Class
