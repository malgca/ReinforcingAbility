Imports System
Imports System.Data
Imports DataTier
Imports System.Data.OleDb

Namespace LogicTier
    Public Class Company
        Public Property CompanyNo As String
        Public Property CompanyName As String
        Public Property Address As String
        Public Property AddressLine2 As String
        Public Property AddressLine3 As String
        Public Property AddressLine4 As String
        Public Property Email As String
        Public Property Fax As String
        Public Property LastCutNum As Integer
        Public Property LastInvNum As Integer
        Public Property Message As String
        Public Property PostalCode As String
        Public Property RegNo As String
        Public Property TelNo As String
        Public Property UnitofMeas As String
        Public Property VatNo As String
        Public Property VatPerc As Decimal
        Public Property Website As String
        Public Property NoAndName As String

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
End Namespace

