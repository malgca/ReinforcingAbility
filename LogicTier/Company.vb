Imports System
Imports System.Data
Imports DataTier
Imports System.Data.OleDb

Public Class Company
    Public Property DataSetCompany As DataTable
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

