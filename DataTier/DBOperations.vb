Imports System
Imports System.Data
Imports System.Data.SqlClient

Namespace DataTier
    Public Class DBOperations
        Dim connection As OleDbConnection
        Dim objConnection As SqlConnection
        Dim connectionString As String
        Private Sub New()
            Me.connection = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source = winsteelVers5.mdb")
            connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = winsteelVers5.mdb"

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

        Public Function getCompanyDataSet()

            objConnection = New SqlConnection(connectionString)
            objConnection.Open()

            Dim daCompany As New SqlDataAdapter("Select * From Company", objConnection)

            Dim dsCompany As New DataSet("Companys")

            daCompany.FillSchema(dsCompany, SchemaType.Source, "Company")
            daCompany.Fill(dsCompany, "Company")

            'Dim tblCompany As DataTable
            'tblCompany = dsCompany.Tables("Authors")

            Return dsCompany
        End Function
    End Class
End Namespace
