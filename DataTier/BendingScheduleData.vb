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

    Private Property InsertCommand As New OleDbCommand
    Private Property SelectCommand As New OleDbCommand
    Private Property UpdateCommand As New OleDbCommand
    Private Property DeleteCommand As New OleDbCommand

    Private Property CountCommand As New OleDbCommand

    Private Property ScheduleSet As New DataSet

    Public Sub New(ByRef jobNumber As String)
        ScheduleSet.Locale = New CultureInfo("en-ZA")
        ScheduleSet.SchemaSerializationMode = SchemaSerializationMode.IncludeSchema

        MapTable()

        PrepareDeleteCommand(jobNumber)

        PrepareInsertCommand()

        PrepareSelectJobsCommand()

        PrepareUpdateCommand()

        PrepareCountCommand(jobNumber)
    End Sub

    ' prepares select query for adapter
    Private Sub PrepareSelectJobsCommand()
        Me.Adapter.SelectCommand = Me.SelectCommand

        ' SelectCommand
        Me.SelectCommand.CommandText = "SELECT JobNo FROM Job ORDER BY JobNo"

        Me.SelectCommand.Connection = DBOperations.GetInstance.Connection
    End Sub
End Class
