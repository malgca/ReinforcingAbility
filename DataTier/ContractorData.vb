Imports System
Imports System.Globalization
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Common
''' <summary>
''' Contractor Data Operations
''' </summary>
Public Class ContractorData

    Public Property Adapter As New OleDbDataAdapter

    Private Property InsertCommand As New OleDbCommand
    Private Property SelectCommand As New OleDbCommand
    Private Property UpdateCommand As New OleDbCommand
    Private Property DeleteCommand As New OleDbCommand

    Private Property CountCompaniesCommand As New OleDbCommand

    Private Property CompanySet As New DataSet

    Public Sub New(ByRef companyNumber As String, ByRef companyName As String, ByRef regNumber As String, ByRef vatNumber As String, ByRef addressLine1 As String, ByRef addressLine2 As String, ByRef addressLine3 As String, ByRef addressLine4 As String, ByRef postalCode As Integer, ByRef telephone As String, ByRef email As String, ByRef fax As String, ByRef website As String, ByRef message As String, ByRef vatPercentage As Double, ByRef lastInvoiceNumber As Integer, ByRef lastCuttingSheetNumber As Integer)
        CompanySet.Locale = New CultureInfo("en-ZA")
        CompanySet.SchemaSerializationMode = SchemaSerializationMode.IncludeSchema

        MapTable()

        PrepareDeleteCommand(companyNumber)

        PrepareInsertCommand(companyNumber, companyName, regNumber, vatNumber, addressLine1, addressLine2, addressLine3, addressLine4, postalCode, telephone, email, fax, website, message, vatPercentage, lastInvoiceNumber, lastCuttingSheetNumber)

        PrepareSelectCommand()

        PrepareUpdateCommand(companyNumber, companyName, regNumber, vatNumber, addressLine1, addressLine2, addressLine3, addressLine4, postalCode, telephone, email, fax, website, message, vatPercentage, lastInvoiceNumber, lastCuttingSheetNumber)

        PrepareCountCompaniesCommand(companyNumber)
    End Sub

    ''' <summary>
    ''' Gets the number of companies currently in the company table
    ''' </summary>
    ''' <param name="CompanyNumber">ID to identify a given company</param>
    ''' <param name="count">Count of existing companies</param>
    Public Sub GetNumberOfCompanies(ByRef CompanyNumber As String, ByRef count As Integer)
        DBOperations.GetInstance.Connection.Open()
        CountCompaniesCommand.Parameters("CompanyNo").Value = CompanyNumber

        Dim dataReader = CountCompaniesCommand.ExecuteReader(CommandBehavior.CloseConnection)

        While dataReader.Read()
            count += 1
        End While

        dataReader.Close()
        DBOperations.GetInstance.Connection.Close()
    End Sub

    ''' <summary>
    ''' adds a new row to the company table
    ''' </summary>
    Public Sub AddCompanyRow(ByRef companyNumber As String, ByRef companyName As String, ByRef regNumber As String, ByRef vatNumber As String, ByRef addressLine1 As String, ByRef addressLine2 As String, ByRef addressLine3 As String, ByRef addressLine4 As String, ByRef postalCode As Integer, ByRef telephone As String, ByRef email As String, ByRef fax As String, ByRef website As String, ByRef message As String, ByRef vatPercentage As Double, ByRef lastInvoiceNumber As Integer, ByRef lastCuttingSheetNumber As Integer)
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
    Public Sub SaveCompanyRowEdit(ByRef companyNumber As String, ByRef companyName As String, ByRef regNumber As String, ByRef vatNumber As String, ByRef addressLine1 As String, ByRef addressLine2 As String, ByRef addressLine3 As String, ByRef addressLine4 As String, ByRef postalCode As Integer, ByRef telephone As String, ByRef email As String, ByRef fax As String, ByRef website As String, ByRef message As String, ByRef vatPercentage As Double, ByRef lastInvoiceNumber As Integer, ByRef lastCuttingSheetNumber As Integer)
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
    Private Sub PrepareInsertCommand(ByRef companyNumber As String, ByRef companyName As String, ByRef regNumber As String, ByRef vatNumber As String, ByRef addressLine1 As String, ByRef addressLine2 As String, ByRef addressLine3 As String, ByRef addressLine4 As String, ByRef postalCode As Integer, ByRef telephone As String, ByRef email As String, ByRef fax As String, ByRef website As String, ByRef message As String, ByRef vatPercentage As Double, ByRef lastInvoiceNumber As Integer, ByRef lastCuttingSheetNumber As Integer)
        ' set insert command in adapter
        Adapter.InsertCommand = Me.InsertCommand

        'InsertCommand
        InsertCommand.CommandText =
        "INSERT INTO Company" & _
        "(Address, AddressLine2, AddressLine3, AddressLine4, CompanyName, CompanyNo, Email, Fax, LastCutNum, LastInvNum, Message, PostalCode, RegNo, Telephone, UnitOfMeas, VatNo, VatPerc, Website) " & _
        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"

        InsertCommand.Connection = DBOperations.GetInstance.Connection

        Dim newParam As New OleDbParameter("addressLine1", addressLine1)
        newParam.SourceColumn = "Address"
        newParam.OleDbType = OleDbType.VarWChar
        Me.InsertCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter("addressLine2", addressLine2)
        newParam.SourceColumn = "AddressLine2"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 40
        Me.InsertCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter("addressLine3", addressLine3)
        newParam.SourceColumn = "AddressLine3"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 40
        newParam.IsNullable = True
        newParam.Value = DBNull.Value
        Me.InsertCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter("addressLine4", addressLine4)
        newParam.SourceColumn = "AddressLine4"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 40
        newParam.IsNullable = True
        newParam.Value = DBNull.Value
        Me.InsertCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter("companyName", companyName)
        newParam.SourceColumn = "CompanyName"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 40
        Me.InsertCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter("companyNo", companyNumber)
        newParam.SourceColumn = "CompanyNo"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 255
        Me.InsertCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter("email", email)
        newParam.SourceColumn = "Email"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 40
        newParam.IsNullable = True
        newParam.Value = DBNull.Value
        Me.InsertCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter("fax", fax)
        newParam.SourceColumn = "Fax"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 15
        Me.InsertCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter("lastCuttingSheetNumber", lastCuttingSheetNumber)
        newParam.SourceColumn = "LastCutNum"
        newParam.OleDbType = OleDbType.Integer
        Me.InsertCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter("lastInvoiceNumber", lastInvoiceNumber)
        newParam.SourceColumn = "LastInvNum"
        newParam.OleDbType = OleDbType.Integer
        Me.InsertCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter("message", message)
        newParam.SourceColumn = "Message"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 200
        Me.InsertCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter("postalCode", postalCode)
        newParam.SourceColumn = "PostalCode"
        newParam.OleDbType = OleDbType.Integer
        Me.InsertCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter("regNumber", regNumber)
        newParam.SourceColumn = "RegNo"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 20
        Me.InsertCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter("telephone", telephone)
        newParam.SourceColumn = "Telephone"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 15
        Me.InsertCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter("unitOfMeasurement", DBNull.Value)
        newParam.SourceColumn = "UnitOfMeas"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 50
        newParam.IsNullable = True
        Me.InsertCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter("vatNumber", vatNumber)
        newParam.SourceColumn = "VatNo"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 20
        Me.InsertCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter("vatPercentage", vatPercentage)
        newParam.SourceColumn = "VatPerc"
        newParam.OleDbType = OleDbType.Double
        newParam.Direction = ParameterDirection.Input
        newParam.IsNullable = False
        newParam.Precision = CByte(2)
        newParam.Scale = CByte(2)
        newParam.SourceVersion = DataRowVersion.Current
        Me.InsertCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter("website", website)
        newParam.SourceColumn = "Website"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 30
        newParam.IsNullable = True
        newParam.Value = DBNull.Value
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
    Private Sub PrepareUpdateCommand(ByRef companyNumber As String, ByRef companyName As String, ByRef regNumber As String, ByRef vatNumber As String, ByRef addressLine1 As String, ByRef addressLine2 As String, ByRef addressLine3 As String, ByRef addressLine4 As String, ByRef postalCode As Integer, ByRef telephone As String, ByRef email As String, ByRef fax As String, ByRef website As String, ByRef message As String, ByRef vatPercentage As Double, ByRef lastInvoiceNumber As Integer, ByRef lastCuttingSheetNumber As Integer)
        Me.Adapter.UpdateCommand = Me.UpdateCommand

        CompanySet.Clear()

        Adapter.Fill(CompanySet)

        Dim col = CompanySet.Tables.Item(0).Columns
        Dim row = CompanySet.Tables.Item(0).Rows.Item(0)

        For counter As Integer = 0 To CompanySet.Tables.Item(0).Rows.Count - 1
            Dim thisRow = CompanySet.Tables.Item(0).Rows.Item(counter)

            If thisRow("CompanyNo") = companyNumber Then
                row = thisRow
                Exit For
            End If
        Next

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

        Dim newParam As New OleDbParameter("addressLine1", addressLine1)
        newParam.SourceColumn = "Address"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 0
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter("addressLine2", addressLine2)
        newParam.SourceColumn = "AddressLine2"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 40
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter("addressLine3", addressLine3)
        newParam.SourceColumn = "AddressLine3"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 40
        newParam.Value = DBNull.Value
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter("addressLine4", addressLine4)
        newParam.SourceColumn = "AddressLine4"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 40
        newParam.Value = DBNull.Value
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter("companyName", companyName)
        newParam.SourceColumn = "CompanyName"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 40
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter("companyNumber", companyNumber)
        newParam.SourceColumn = "CompanyNo"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 10
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter("email", email)
        newParam.SourceColumn = "Email"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 40
        newParam.Value = DBNull.Value
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter("fax", fax)
        newParam.SourceColumn = "Fax"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 40
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter("lastCuttingSheetNumber", lastCuttingSheetNumber)
        newParam.SourceColumn = "LastCutNum"
        newParam.OleDbType = OleDbType.Integer
        newParam.Size = 0
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter("lastInvoiceNumber", lastInvoiceNumber)
        newParam.SourceColumn = "LastInvNum"
        newParam.OleDbType = OleDbType.Integer
        newParam.Size = 0
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter("message", message)
        newParam.SourceColumn = "Message"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 200
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter("postalCode", postalCode)
        newParam.SourceColumn = "PostalCode"
        newParam.OleDbType = OleDbType.Integer
        newParam.Size = 0
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter("regNumber", regNumber)
        newParam.SourceColumn = "RegNo"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 20
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter("telephone", telephone)
        newParam.SourceColumn = "Telephone"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 15
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter("unitOfMeasurement", DBNull.Value)
        newParam.SourceColumn = "UnitOfMeas"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 50
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter("vatNumber", vatNumber)
        newParam.SourceColumn = "VatNo"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 20
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter("vatPercentage", vatPercentage)
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

        newParam = New OleDbParameter("website", website)
        newParam.SourceColumn = "Website"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 30
        newParam.Value = DBNull.Value
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter("originalCompayNo", companyNumber)
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

        newParam = New OleDbParameter("originalAddressLine2", row("AddressLine2"))
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

        newParam = New OleDbParameter("originalAddressLine21", row("AddressLine2"))
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

        newParam = New OleDbParameter("originalAddressLine3", row("AddressLine3"))
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

        newParam = New OleDbParameter("originalAddressLine31", row("AddressLine3"))
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

        newParam = New OleDbParameter("originalAddressLine4", row("AddressLine4"))
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

        newParam = New OleDbParameter("originalAddressLine41", row("AddressLine4"))
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

        newParam = New OleDbParameter("originalCompanyName", row("CompanyName"))
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

        newParam = New OleDbParameter("originalCompanyName1", row("CompanyName"))
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

        newParam = New OleDbParameter("originalEmail", row("Email"))
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

        newParam = New OleDbParameter("originalEmail1", row("Email"))
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

        newParam = New OleDbParameter("originalFax", row("Fax"))
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

        newParam = New OleDbParameter("originalFax1", row("Fax"))
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

        newParam = New OleDbParameter("originalLastCuttingSheetNumber", row("LastCutNum"))
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

        newParam = New OleDbParameter("originalLastCuttingSheetNumber1", row("LastCutNum"))
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

        newParam = New OleDbParameter("originalLastInvNum", row("LastInvNum"))
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

        newParam = New OleDbParameter("originalLastInvoiceNumber", row("LastInvNum"))
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

        newParam = New OleDbParameter("originalMessage", row("Message"))
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

        newParam = New OleDbParameter("originalMessage1", row("Message"))
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

        newParam = New OleDbParameter("originalPostalCode", row("PostalCode"))
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

        newParam = New OleDbParameter("originalPostalCode1", row("PostalCode"))
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

        newParam = New OleDbParameter("originalRegNumber", row("RegNo"))
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

        newParam = New OleDbParameter("originalRegNumber1", row("RegNo"))
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

        newParam = New OleDbParameter("originalTelephone", row("Telephone"))
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

        newParam = New OleDbParameter("originalTelephone1", row("Telephone"))
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

        newParam = New OleDbParameter("originalUnitOfMeasurement", row("UnitOfMeas"))
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

        newParam = New OleDbParameter("originalUnitOfMeasurement1", row("UnitOfMeas"))
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

        newParam = New OleDbParameter("originalVatNumber", row("VatNo"))
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

        newParam = New OleDbParameter("originalVatNumber1", row("VatNo"))
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

        newParam = New OleDbParameter("originalVatPercentage", row("VatPerc"))
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

        newParam = New OleDbParameter("originalVatPercentage1", row("VatPerc"))
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

        newParam = New OleDbParameter("originalWebsite", row("Website"))
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

        newParam = New OleDbParameter("originalWebsite1", row("Website"))
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
    Private Sub PrepareCountCompaniesCommand(ByRef companyNumber As String)
        Me.CountCompaniesCommand.CommandText = "SELECT * FROM Company WHERE (CompanyNo = ?)"

        Me.CountCompaniesCommand.Connection = DBOperations.GetInstance.Connection

        Dim newParam As New OleDbParameter("CompanyNo", companyNumber)
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 10
        Me.CountCompaniesCommand.Parameters.Add(newParam)
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

        If IsNotEmpty(addressLine1) Then
            row("Address") = addressLine1
        End If

        If IsNotEmpty(addressLine2) Then
            row("AddressLine2") = addressLine2
        End If

        If IsNotEmpty(addressLine3) Then
            row("AddressLine3") = addressLine3
        End If

        If IsNotEmpty(addressLine4) Then
            row("AddressLine4") = addressLine4
        End If

        If IsNotEmpty(postalCode.ToString()) Then
            row("PostalCode") = postalCode
        End If

        If IsNotEmpty(telephone) Then
            row("Telephone") = telephone
        End If

        If IsNotEmpty(email) Then
            row("Email") = email
        End If

        If IsNotEmpty(fax) Then
            row("Fax") = fax
        End If

        If IsNotEmpty(website) Then
            row("Website") = website
        End If

        If IsNotEmpty(message) Then
            row("Message") = message
        End If

        If IsNotEmpty(vatPercentage) Then
            row("VatPerc") = vatPercentage
        End If

        If IsNotEmpty(lastInvoiceNumber) Then
            row("LastInvNum") = lastInvoiceNumber
        End If

        If IsNotEmpty(lastCuttingSheetNumber) Then
            row("LastCutNum") = lastCuttingSheetNumber
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