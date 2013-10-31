Imports System
Imports System.Data
Imports DataTier
Imports System.Data.OleDb

Public Class Company
    ''' <summary>
    ''' Company DataSet. Used for Databindings
    ''' </summary>
    Public Property DataSetCompany As DataTable

    ''' <summary>
    ''' Value Added Tax
    ''' </summary>
    Public ReadOnly Property VAT As Double
        Get
            Return 0.14
        End Get
    End Property

    Public Sub getCompanyDataSet(ByRef daCompany As OleDbDataAdapter, ByRef dsCompany As DataTable)
        ' Open the Ms Access connection.
        Dim dbConnection As OleDbConnection = DBOperations.GetInstance().Connection

        daCompany.SelectCommand = New OleDbCommand("Select * From Company",
                                              dbConnection)

        ' Should move any Dataset operations to the DataTier (malusi moved it here)
        Dim companySet As New DataSet("Companies")

        daCompany.FillSchema(companySet, SchemaType.Source, "Company")
        daCompany.Fill(companySet.Tables.Item(0))

        Console.WriteLine(companySet)
        dsCompany = companySet.Tables.Item(0)
    End Sub
End Class

