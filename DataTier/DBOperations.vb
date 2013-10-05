Imports System
Imports System.Data
Imports System.Data.SqlClient

Public Class DBOperations
    ' Private connection variable
    Private _connection As OleDbConnection

    Public Property Connection As OleDbConnection
        Get
            Return _connection
        End Get
        Private Set(value As OleDbConnection)
            If (Not (value.Equals(_connection))) Then
                _connection = value
            End If
        End Set
    End Property

    Private Sub New()
        Me._connection = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source = winsteelVers5.mdb")

        ' Make sure only a single instance of this class may exist
    End Sub
    ''' <summary>
    ''' Get the running instance of the DBOperations class
    ''' </summary>
    ''' <returns>Singleton instance of DBOperations class</returns>
    Public Shared ReadOnly Property GetInstance As DBOperations
        Get
            Static Instance As DBOperations = New DBOperations
            Return Instance
        End Get
    End Property

    ''' <summary>
    ''' Execute a given query against the database.
    ''' </summary>
    ''' <param name="query">The query to be executed.</param>
    ''' <returns>Results of executed query.</returns>
    Public Function ExecuteQuery(query)
        '' execute given query
        Return vbNull
    End Function
End Class