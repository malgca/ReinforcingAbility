Imports System
Imports System.Globalization
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Common
''' <summary>
''' Company Data Operations
''' </summary>
Public Class CompanyData

    Public Property Adapter As New OleDbDataAdapter

    Private Property InsertCommand As New OleDbCommand
    Private Property SelectCommand As New OleDbCommand
    Private Property UpdateCommand As New OleDbCommand
    Private Property DeleteCommand As New OleDbCommand

    Private Property CountCommand As New OleDbCommand

    Private Property CompanySet As New DataSet

    Public Sub New(ByRef companyNumber As String)
        CompanySet.Locale = New CultureInfo("en-ZA")
        CompanySet.SchemaSerializationMode = SchemaSerializationMode.IncludeSchema

        MapTable()

        PrepareDeleteCommand(companyNumber)

        PrepareInsertCommand()

        PrepareSelectCommand()

        PrepareUpdateCommand()

        PrepareCountCommand(companyNumber)
    End Sub

    ''' <summary>
    ''' Checks if a particular company number key exists in the company table
    ''' </summary>
    ''' <param name="companyNumber">ID to identify a given company</param>
    ''' <param name="count">Count of existing companies</param>
    Public Sub GetNumberOfCompanies(ByRef companyNumber As String, ByRef count As Integer)
        DBOperations.GetInstance.Connection.Open()
        CountCommand.Parameters("companyNumber").Value = companyNumber

        Dim dataReader = CountCommand.ExecuteReader(CommandBehavior.CloseConnection)

        While dataReader.Read()
            count += 1
        End While

        dataReader.Close()
        DBOperations.GetInstance.Connection.Close()
    End Sub

    ''' <summary>
    ''' Adds a new row to the company table
    ''' </summary>
    Public Sub AddRow(ByRef companyNumber As String, ByRef companyName As String, ByRef regNumber As String, ByRef vatNumber As String, ByRef addressLine1 As String, ByRef addressLine2 As String, ByRef addressLine3 As String, ByRef addressLine4 As String, ByRef postalCode As Integer, ByRef telephone As String, ByRef email As String, ByRef fax As String, ByRef website As String, ByRef message As String, ByRef vatPercentage As Double, ByRef lastInvoiceNumber As Integer, ByRef lastCuttingSheetNumber As Integer)
        CompanySet.Clear()
        Adapter.Fill(CompanySet)

        Dim row = CompanySet.Tables.Item(0).NewRow()

        UpdateRowValues(row, companyNumber, companyName, regNumber, vatNumber, addressLine1, addressLine2, addressLine3, addressLine4, postalCode, telephone, email, fax, website, message, vatPercentage, lastInvoiceNumber, lastCuttingSheetNumber)

        CompanySet.Tables.Item(0).Rows.InsertAt(row, companyNumber)
        Adapter.Update(CompanySet.Tables.Item(0))
    End Sub

    ''' <summary>
    ''' Saves an edit to a row in the company table
    ''' </summary>
    Public Sub SaveRowEdit(ByRef companyNumber As String, ByRef companyName As String, ByRef regNumber As String, ByRef vatNumber As String, ByRef addressLine1 As String, ByRef addressLine2 As String, ByRef addressLine3 As String, ByRef addressLine4 As String, ByRef postalCode As Integer, ByRef telephone As String, ByRef email As String, ByRef fax As String, ByRef website As String, ByRef message As String, ByRef vatPercentage As Double, ByRef lastInvoiceNumber As Integer, ByRef lastCuttingSheetNumber As Integer)
        CompanySet.Clear()
        Adapter.Fill(CompanySet)

        Dim row = CompanySet.Tables.Item(0).Rows.Item(0)

        For counter As Integer = 0 To CompanySet.Tables.Item(0).Rows.Count - 1
            Dim thisRow = CompanySet.Tables.Item(0).Rows.Item(counter)

            If thisRow("CompanyNo") = companyNumber Then
                row = thisRow
                Exit For
            End If
        Next

        UpdateRowValues(row, companyNumber, companyName, regNumber, vatNumber, addressLine1, addressLine2, addressLine3, addressLine4, postalCode, telephone, email, fax, website, message, vatPercentage, lastInvoiceNumber, lastCuttingSheetNumber)

        Adapter.Update(CompanySet.Tables.Item(0))
    End Sub

    ' prepares insert command for adapter
    Private Sub PrepareInsertCommand()
        ' set insert command in adapter
        Adapter.InsertCommand = Me.InsertCommand

        'InsertCommand
        InsertCommand.CommandText =
        "INSERT INTO Company" & _
        "(Address, AddressLine2, AddressLine3, AddressLine4, CompanyName, CompanyNo, Email, Fax, LastCutNum, LastInvNum, Message, PostalCode, RegNo, Telephone, UnitOfMeas, VatNo, VatPerc, Website) " & _
        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"

        InsertCommand.Connection = DBOperations.GetInstance.Connection

        Dim newParam As New OleDbParameter()
        newParam.SourceColumn = "Address"
        newParam.OleDbType = OleDbType.VarWChar
        Me.InsertCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "AddressLine2"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 40
        Me.InsertCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "AddressLine3"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 40
        newParam.IsNullable = True
        Me.InsertCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "AddressLine4"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 40
        newParam.IsNullable = True
        Me.InsertCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "CompanyName"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 40
        Me.InsertCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "CompanyNo"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 10
        Me.InsertCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "Email"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 40
        newParam.IsNullable = True
        Me.InsertCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "Fax"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 15
        Me.InsertCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "LastCutNum"
        newParam.OleDbType = OleDbType.Integer
        Me.InsertCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "LastInvNum"
        newParam.OleDbType = OleDbType.Integer
        Me.InsertCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "Message"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 200
        Me.InsertCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "PostalCode"
        newParam.OleDbType = OleDbType.Integer
        Me.InsertCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "RegNo"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 20
        Me.InsertCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "Telephone"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 15
        Me.InsertCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "UnitOfMeas"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 50
        newParam.IsNullable = True
        Me.InsertCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "VatNo"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 20
        Me.InsertCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "VatPerc"
        newParam.OleDbType = OleDbType.Double
        newParam.Direction = ParameterDirection.Input
        newParam.IsNullable = False
        newParam.Precision = CByte(2)
        newParam.Scale = CByte(2)
        newParam.SourceVersion = DataRowVersion.Current
        Me.InsertCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "Website"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 30
        newParam.IsNullable = True
        Me.InsertCommand.Parameters.Add(newParam)
    End Sub

    ' prepares select query for adapter
    Private Sub PrepareSelectCommand()
        Me.Adapter.SelectCommand = Me.SelectCommand

        ' SelectCommand
        Me.SelectCommand.CommandText = "SELECT * FROM Company"

        Me.SelectCommand.Connection = DBOperations.GetInstance.Connection
    End Sub

    ' prepares update query for adapter
    Private Sub PrepareUpdateCommand()
        Me.Adapter.UpdateCommand = Me.UpdateCommand

        CompanySet.Clear()

        Adapter.Fill(CompanySet)

        ' UpdateCommand
        Me.UpdateCommand.CommandText =
            "UPDATE Company " & _
            "SET Address = ?, " & _
            "AddressLine2 = ?, " & _
            "AddressLine3 = ?, " & _
            "AddressLine4 = ?, " & _
            "CompanyName = ?, " & _
            "CompanyNo = ?, " & _
            "Email = ?, " & _
            "Fax = ?, " & _
            "LastCutNum = ?, " & _
            "LastInvNum = ?, " & _
            "Message = ?, " & _
            "PostalCode = ?, " & _
            "RegNo = ?, " & _
            "Telephone = ?, " & _
            "UnitOfMeas = ?, " & _
            "VatNo = ?, " & _
            "VatPerc = ?, " & _
            "Website = ? " & _
            "WHERE (CompanyNo = ?) " & _
        "AND (AddressLine2 = ? OR ? IS NULL AND AddressLine2 IS NULL) " & _
        "AND (AddressLine3 = ? OR ? IS NULL AND AddressLine3 IS NULL) " & _
        "AND (AddressLine4 = ? OR ? IS NULL AND AddressLine4 IS NULL) " & _
        "AND (CompanyName = ? OR ? IS NULL AND CompanyName IS NULL) " & _
        "AND (Email = ? OR ? IS NULL AND Email IS NULL) " & _
        "AND (Fax = ? OR ? IS NULL AND Fax IS NULL) " & _
        "AND (LastCutNum = ? OR ? IS NULL AND LastCutNum IS NULL) " & _
        "AND (LastInvNum = ? OR ? IS NULL AND LastInvNum IS NULL) " & _
        "AND (Message = ? OR ? IS NULL AND Message IS NULL) " & _
        "AND (PostalCode = ? OR ? IS NULL AND PostalCode IS NULL) " & _
        "AND (RegNo = ? OR ? IS NULL AND RegNo IS NULL) " & _
        "AND (Telephone = ? OR ? IS NULL AND Telephone IS NULL) " & _
        "AND (UnitOfMeas = ? OR ? IS NULL AND UnitOfMeas IS NULL) " & _
        "AND (VatNo = ? OR ? IS NULL AND VatNo IS NULL) " & _
        "AND (VatPerc = ? OR ? IS NULL AND VatPerc IS NULL) " & _
        "AND (Website = ? OR ? IS NULL AND Website IS NULL)"

        Me.UpdateCommand.Connection = DBOperations.GetInstance.Connection

        Dim newParam As New OleDbParameter()
        newParam.SourceColumn = "Address"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 0
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "AddressLine2"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 40
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "AddressLine3"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 40
        newParam.Value = DBNull.Value
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "AddressLine4"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 40
        newParam.Value = DBNull.Value
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "CompanyName"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 40
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "CompanyNo"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 10
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "Email"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 40
        newParam.Value = DBNull.Value
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "Fax"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 40
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "LastCutNum"
        newParam.OleDbType = OleDbType.Integer
        newParam.Size = 0
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "LastInvNum"
        newParam.OleDbType = OleDbType.Integer
        newParam.Size = 0
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "Message"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 200
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "PostalCode"
        newParam.OleDbType = OleDbType.Integer
        newParam.Size = 0
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "RegNo"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 20
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "Telephone"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 15
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "UnitOfMeas"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 50
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "VatNo"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 20
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "VatPerc"
        newParam.OleDbType = OleDbType.Double
        newParam.Size = 0
        newParam.Direction = ParameterDirection.Input
        newParam.IsNullable = False
        newParam.Precision = CByte(2)
        newParam.Scale = CByte(2)
        newParam.SourceVersion = DataRowVersion.Current
        newParam.Value = DBNull.Value
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "Website"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 30
        newParam.Value = DBNull.Value
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "CompanyNo"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 10
        newParam.Direction = ParameterDirection.Input
        newParam.IsNullable = False
        newParam.Precision = CType(0, Byte)
        newParam.Scale = CType(0, Byte)
        newParam.SourceVersion = DataRowVersion.Original
        newParam.Value = Nothing
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "AddressLine2"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 40
        newParam.Direction = ParameterDirection.Input
        newParam.IsNullable = False
        newParam.Precision = CType(0, Byte)
        newParam.Scale = CType(0, Byte)
        newParam.SourceVersion = DataRowVersion.Original
        newParam.Value = Nothing
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "AddressLine2"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 40
        newParam.Direction = ParameterDirection.Input
        newParam.IsNullable = False
        newParam.Precision = CType(0, Byte)
        newParam.Scale = CType(0, Byte)
        newParam.SourceVersion = DataRowVersion.Original
        newParam.Value = Nothing
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "AddressLine3"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 40
        newParam.Direction = ParameterDirection.Input
        newParam.IsNullable = False
        newParam.Precision = CType(0, Byte)
        newParam.Scale = CType(0, Byte)
        newParam.SourceVersion = DataRowVersion.Original
        newParam.Value = Nothing
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "AddressLine3"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 40
        newParam.Direction = ParameterDirection.Input
        newParam.IsNullable = False
        newParam.Precision = CType(0, Byte)
        newParam.Scale = CType(0, Byte)
        newParam.SourceVersion = DataRowVersion.Original
        newParam.Value = Nothing
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "AddressLine4"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 40
        newParam.Direction = ParameterDirection.Input
        newParam.IsNullable = False
        newParam.Precision = CType(0, Byte)
        newParam.Scale = CType(0, Byte)
        newParam.SourceVersion = DataRowVersion.Original
        newParam.Value = Nothing
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "AddressLine4"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 40
        newParam.Direction = ParameterDirection.Input
        newParam.IsNullable = False
        newParam.Precision = CType(0, Byte)
        newParam.Scale = CType(0, Byte)
        newParam.SourceVersion = DataRowVersion.Original
        newParam.Value = Nothing
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "CompanyName"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 40
        newParam.Direction = ParameterDirection.Input
        newParam.IsNullable = False
        newParam.Precision = CType(0, Byte)
        newParam.Scale = CType(0, Byte)
        newParam.SourceVersion = DataRowVersion.Original
        newParam.Value = Nothing
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "CompanyName"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 40
        newParam.Direction = ParameterDirection.Input
        newParam.IsNullable = False
        newParam.Precision = CType(0, Byte)
        newParam.Scale = CType(0, Byte)
        newParam.SourceVersion = DataRowVersion.Original
        newParam.Value = Nothing
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "Email"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 40
        newParam.Direction = ParameterDirection.Input
        newParam.IsNullable = False
        newParam.Precision = CType(0, Byte)
        newParam.Scale = CType(0, Byte)
        newParam.SourceVersion = DataRowVersion.Original
        newParam.Value = Nothing
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "Email"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 40
        newParam.Direction = ParameterDirection.Input
        newParam.IsNullable = False
        newParam.Precision = CType(0, Byte)
        newParam.Scale = CType(0, Byte)
        newParam.SourceVersion = DataRowVersion.Original
        newParam.Value = Nothing
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "Fax"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 15
        newParam.Direction = ParameterDirection.Input
        newParam.IsNullable = False
        newParam.Precision = CType(0, Byte)
        newParam.Scale = CType(0, Byte)
        newParam.SourceVersion = DataRowVersion.Original
        newParam.Value = Nothing
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "Fax"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 15
        newParam.Direction = ParameterDirection.Input
        newParam.IsNullable = False
        newParam.Precision = CType(0, Byte)
        newParam.Scale = CType(0, Byte)
        newParam.SourceVersion = DataRowVersion.Original
        newParam.Value = Nothing
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "LastCutNum"
        newParam.OleDbType = OleDbType.Integer
        newParam.Size = 0
        newParam.Direction = ParameterDirection.Input
        newParam.IsNullable = False
        newParam.Precision = CType(0, Byte)
        newParam.Scale = CType(0, Byte)
        newParam.SourceVersion = DataRowVersion.Original
        newParam.Value = Nothing
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "LastCutNum"
        newParam.OleDbType = OleDbType.Integer
        newParam.Size = 0
        newParam.Direction = ParameterDirection.Input
        newParam.IsNullable = False
        newParam.Precision = CType(0, Byte)
        newParam.Scale = CType(0, Byte)
        newParam.SourceVersion = DataRowVersion.Original
        newParam.Value = Nothing
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "LastInvNum"
        newParam.OleDbType = OleDbType.Integer
        newParam.Size = 0
        newParam.Direction = ParameterDirection.Input
        newParam.IsNullable = False
        newParam.Precision = CType(0, Byte)
        newParam.Scale = CType(0, Byte)
        newParam.SourceVersion = DataRowVersion.Original
        newParam.Value = Nothing
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "LastInvNum"
        newParam.OleDbType = OleDbType.Integer
        newParam.Size = 0
        newParam.Direction = ParameterDirection.Input
        newParam.IsNullable = False
        newParam.Precision = CType(0, Byte)
        newParam.Scale = CType(0, Byte)
        newParam.SourceVersion = DataRowVersion.Original
        newParam.Value = Nothing
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "Message"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 200
        newParam.Direction = ParameterDirection.Input
        newParam.IsNullable = False
        newParam.Precision = CType(0, Byte)
        newParam.Scale = CType(0, Byte)
        newParam.SourceVersion = DataRowVersion.Original
        newParam.Value = Nothing
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "Message"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 200
        newParam.Direction = ParameterDirection.Input
        newParam.IsNullable = False
        newParam.Precision = CType(0, Byte)
        newParam.Scale = CType(0, Byte)
        newParam.SourceVersion = DataRowVersion.Original
        newParam.Value = Nothing
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "PostalCode"
        newParam.OleDbType = OleDbType.Integer
        newParam.Size = 0
        newParam.Direction = ParameterDirection.Input
        newParam.IsNullable = False
        newParam.Precision = CType(0, Byte)
        newParam.Scale = CType(0, Byte)
        newParam.SourceVersion = DataRowVersion.Original
        newParam.Value = Nothing
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "PostalCode"
        newParam.OleDbType = OleDbType.Integer
        newParam.Size = 0
        newParam.Direction = ParameterDirection.Input
        newParam.IsNullable = False
        newParam.Precision = CType(0, Byte)
        newParam.Scale = CType(0, Byte)
        newParam.SourceVersion = DataRowVersion.Original
        newParam.Value = Nothing
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "RegNo"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 20
        newParam.Direction = ParameterDirection.Input
        newParam.IsNullable = False
        newParam.Precision = CType(0, Byte)
        newParam.Scale = CType(0, Byte)
        newParam.SourceVersion = DataRowVersion.Original
        newParam.Value = Nothing
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "RegNo"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 20
        newParam.Direction = ParameterDirection.Input
        newParam.IsNullable = False
        newParam.Precision = CType(0, Byte)
        newParam.Scale = CType(0, Byte)
        newParam.SourceVersion = DataRowVersion.Original
        newParam.Value = Nothing
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "Telephone"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 15
        newParam.Direction = ParameterDirection.Input
        newParam.IsNullable = False
        newParam.Precision = CType(0, Byte)
        newParam.Scale = CType(0, Byte)
        newParam.SourceVersion = DataRowVersion.Original
        newParam.Value = Nothing
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "Telephone"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 15
        newParam.Direction = ParameterDirection.Input
        newParam.IsNullable = False
        newParam.Precision = CType(0, Byte)
        newParam.Scale = CType(0, Byte)
        newParam.SourceVersion = DataRowVersion.Original
        newParam.Value = Nothing
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "UnitOfMeas"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 50
        newParam.Direction = ParameterDirection.Input
        newParam.IsNullable = False
        newParam.Precision = CType(0, Byte)
        newParam.Scale = CType(0, Byte)
        newParam.SourceVersion = DataRowVersion.Original
        newParam.Value = Nothing
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "UnitOfMeas"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 50
        newParam.Direction = ParameterDirection.Input
        newParam.IsNullable = False
        newParam.Precision = CType(0, Byte)
        newParam.Scale = CType(0, Byte)
        newParam.SourceVersion = DataRowVersion.Original
        newParam.Value = Nothing
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "VatNo"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 20
        newParam.Direction = ParameterDirection.Input
        newParam.IsNullable = False
        newParam.Precision = CType(0, Byte)
        newParam.Scale = CType(0, Byte)
        newParam.SourceVersion = DataRowVersion.Original
        newParam.Value = Nothing
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "VatNo"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 20
        newParam.Direction = ParameterDirection.Input
        newParam.IsNullable = False
        newParam.Precision = CType(0, Byte)
        newParam.Scale = CType(0, Byte)
        newParam.SourceVersion = DataRowVersion.Original
        newParam.Value = Nothing
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "VatPerc"
        newParam.OleDbType = OleDbType.Double
        newParam.Size = 0
        newParam.Direction = ParameterDirection.Input
        newParam.IsNullable = False
        newParam.Precision = CByte(2)
        newParam.Scale = CByte(2)
        newParam.SourceVersion = DataRowVersion.Original
        newParam.Value = Nothing
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "VatPerc"
        newParam.OleDbType = OleDbType.Double
        newParam.Size = 0
        newParam.Direction = ParameterDirection.Input
        newParam.IsNullable = False
        newParam.Precision = CByte(2)
        newParam.Scale = CByte(2)
        newParam.SourceVersion = DataRowVersion.Original
        newParam.Value = Nothing
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "Website"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 30
        newParam.Direction = ParameterDirection.Input
        newParam.IsNullable = False
        newParam.Precision = CType(0, Byte)
        newParam.Scale = CType(0, Byte)
        newParam.SourceVersion = DataRowVersion.Original
        newParam.Value = Nothing
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "Website"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 30
        newParam.Direction = ParameterDirection.Input
        newParam.IsNullable = False
        newParam.Precision = CType(0, Byte)
        newParam.Scale = CType(0, Byte)
        newParam.SourceVersion = DataRowVersion.Original
        newParam.Value = Nothing
        Me.UpdateCommand.Parameters.Add(newParam)
    End Sub

    ' prepares delete query for adapter
    Private Sub PrepareDeleteCommand(ByRef companyNumber As String)
        Me.Adapter.DeleteCommand = Me.DeleteCommand

        'Delete Command
        Me.DeleteCommand.CommandText =
            "DELETE FROM Company WHERE (CompanyNo = ?)"

        Me.DeleteCommand.Connection = DBOperations.GetInstance.Connection

        Dim newParam As New OleDbParameter("originalCompanyNo", companyNumber)
        newParam.OleDbType = OleDbType.WChar
        newParam.Direction = ParameterDirection.Input
        newParam.IsNullable = False
        newParam.Precision = CType(0, Byte)
        newParam.Scale = CType(0, Byte)
        newParam.SourceColumn = "CompanyNo"
        newParam.SourceVersion = DataRowVersion.Original
        newParam.Value = DBNull.Value
        Me.DeleteCommand.Parameters.Add(newParam)
    End Sub

    ' prepares count companies query for adapter
    Private Sub PrepareCountCommand(ByRef companyNumber As String)
        Me.CountCommand.CommandText = "SELECT * FROM Company WHERE (CompanyNo = ?)"

        Me.CountCommand.Connection = DBOperations.GetInstance.Connection

        Dim newParam As New OleDbParameter("companyNumber", companyNumber)
        newParam.SourceColumn = "CompanyNo"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 10
        Me.CountCommand.Parameters.Add(newParam)
    End Sub

    ' map table to values
    Private Sub MapTable()
        Me.Adapter.TableMappings.AddRange(New DataTableMapping() {
                                  New DataTableMapping("Table", "Company",
                                                       New DataColumnMapping() {
                                                           New DataColumnMapping("Address", "Address"),
                                                           New DataColumnMapping("AddressLine2", "AddressLine2"),
                                                           New DataColumnMapping("AddressLine3", "AddressLine3"),
                                                           New DataColumnMapping("AddressLine4", "AddressLine4"),
                                                           New DataColumnMapping("CompanyName", "CompanyName"),
                                                           New DataColumnMapping("CompanyNo", "CompanyNo"),
                                                           New DataColumnMapping("Email", "Email"),
                                                           New DataColumnMapping("Fax", "Fax"),
                                                           New DataColumnMapping("LastCutNum", "LastCutNum"),
                                                           New DataColumnMapping("LastInvNum", "LastInvNum"),
                                                           New DataColumnMapping("Message", "Message"),
                                                           New DataColumnMapping("PostalCode", "PostalCode"),
                                                           New DataColumnMapping("RegNo", "RegNo"),
                                                           New DataColumnMapping("Telephone", "Telephone"),
                                                           New DataColumnMapping("UnitOfMeas", "UnitOfMeas"),
                                                           New DataColumnMapping("VatNo", "VatNo"),
                                                           New DataColumnMapping("VatPerc", "VatPerc"),
                                                           New DataColumnMapping("Website", "Website")})})
    End Sub

    ' updates a single row in the database
    Private Sub UpdateRowValues(ByRef row As DataRow, ByRef companyNumber As String, ByRef companyName As String, ByRef regNumber As String, ByRef vatNumber As String, ByRef addressLine1 As String, ByRef addressLine2 As String, ByRef addressLine3 As String, ByRef addressLine4 As String, ByRef postalCode As Integer, ByRef telephone As String, ByRef email As String, ByRef fax As String, ByRef website As String, ByRef message As String, ByRef vatPercentage As Double, ByRef lastInvoiceNumber As Integer, ByRef lastCuttingSheetNumber As Integer)
        If IsNotEmpty(companyNumber) Then
            row("CompanyNo") = companyNumber
        End If

        If IsNotEmpty(companyName) Then
            row("CompanyName") = companyName
        End If

        If IsNotEmpty(regNumber) Then
            row("RegNo") = regNumber
        End If

        If IsNotEmpty(vatNumber) Then
            row("VatNo") = vatNumber
        End If

        row("Address") = addressLine1
        row("AddressLine2") = addressLine2
        row("AddressLine3") = addressLine3
        row("AddressLine4") = addressLine4

        If IsNotEmpty(postalCode.ToString()) Then
            row("PostalCode") = postalCode
        Else
            row("PostalCode") = 0
        End If

        If IsNotEmpty(telephone) Then
            row("Telephone") = telephone
        End If

        If IsNotEmpty(email) Then
            row("Email") = email
        End If

        row("Fax") = fax

        If IsNotEmpty(website) Then
            row("Website") = website
        End If

        row("Message") = message

        If IsNotEmpty(vatPercentage) Then
            row("VatPerc") = vatPercentage
        Else
            row("VatPerc") = 0
        End If

        If IsNotEmpty(lastInvoiceNumber) Then
            row("LastInvNum") = lastInvoiceNumber
        Else
            row("LastInvNum") = 0
        End If

        If IsNotEmpty(lastCuttingSheetNumber) Then
            row("LastCutNum") = lastCuttingSheetNumber
        Else
            row("LastCutNum") = 0
        End If

        row("UnitOfMeas") = DBNull.Value
    End Sub
    'Determines whether a parameter string is empty or not
    Private Function IsNotEmpty(ByVal parameter As String) As Boolean
        If parameter = Nothing Then
            Return False
        End If

        Return Not parameter.Equals(String.Empty)
    End Function
End Class