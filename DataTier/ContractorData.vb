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

    Private Property CountCommand As New OleDbCommand

    Private Property ContracterSet As New DataSet

    Public Sub New(ByRef contractorNumber As String)
        ContracterSet.Locale = New CultureInfo("en-ZA")
        ContracterSet.SchemaSerializationMode = SchemaSerializationMode.IncludeSchema

        MapTable()

        PrepareDeleteCommand(contractorNumber)

        PrepareInsertCommand()

        PrepareSelectCommand()

        PrepareUpdateCommand()

        PrepareCountCommand(contractorNumber)
    End Sub

    ''' <summary>
    ''' Checks if a particular contractor number key exists in the contractor table
    ''' </summary>
    ''' <param name="contractorNumber">ID to identify a given company</param>
    ''' <param name="count">Count of existing companies</param>
    Public Sub GetNumberOfContractors(ByRef contractorNumber As String, ByRef count As Integer)
        DBOperations.GetInstance.Connection.Open()
        CountCommand.Parameters("contractorNumber").Value = contractorNumber

        Dim dataReader = CountCommand.ExecuteReader(CommandBehavior.CloseConnection)

        While dataReader.Read()
            count += 1
        End While

        dataReader.Close()
        DBOperations.GetInstance.Connection.Close()
    End Sub

    ''' <summary>
    ''' Adds a new row to the contractor table
    ''' </summary>
    Public Sub AddContractorRow(ByRef contractorNumber As String, ByRef contractorName As String, ByRef addressLine1 As String, ByRef addressLine2 As String, ByRef addressLine3 As String, ByRef addressLine4 As String, ByVal postalCode As Integer, ByRef telephone As String, ByVal isActive As Boolean, ByRef vatNumber As String, ByRef regNumber As String)
        ContracterSet.Clear()
        Adapter.Fill(ContracterSet)

        Dim row = ContracterSet.Tables.Item(0).NewRow()

        UpdateRowValues(row, contractorNumber, contractorName, addressLine1, addressLine2, addressLine3, addressLine4, postalCode, telephone, isActive, vatNumber, regNumber)

        ContracterSet.Tables.Item(0).Rows.InsertAt(row, contractorNumber)
        Adapter.Update(ContracterSet.Tables.Item(0))
    End Sub

    ''' <summary>
    ''' Saves an edit to a row in the company table
    ''' </summary>
    Public Sub SaveCompanyRowEdit(ByRef contractorNumber As String, ByRef contractorName As String, ByRef addressLine1 As String, ByRef addressLine2 As String, ByRef addressLine3 As String, ByRef addressLine4 As String, ByVal postalCode As Integer, ByRef telephone As String, ByVal isActive As Boolean, ByRef vatNumber As String, ByRef regNumber As String)
        ContracterSet.Clear()
        Adapter.Fill(ContracterSet)

        Dim row = ContracterSet.Tables.Item(0).Rows.Item(0)

        For counter As Integer = 0 To ContracterSet.Tables.Item(0).Rows.Count - 1
            Dim thisRow = ContracterSet.Tables.Item(0).Rows.Item(counter)

            If thisRow("ContractorNo") = contractorNumber Then
                row = thisRow
                Exit For
            End If
        Next

        UpdateRowValues(row, contractorNumber, contractorName, addressLine1, addressLine2, addressLine3, addressLine4, postalCode, telephone, isActive, vatNumber, regNumber)

        Adapter.Update(ContracterSet.Tables.Item(0))
    End Sub

    ' prepares insert command for adapter
    Private Sub PrepareInsertCommand()
        ' set insert command in adapter
        Adapter.InsertCommand = Me.InsertCommand

        'InsertCommand
        InsertCommand.CommandText =
        "INSERT INTO Contractor" & _
        "([ActiveY/N], AddressLine1, AddressLine2, AddressLine3, AddressLine4, ContractorName, ContractorNo, PostalCode, [Reg No], Telephone, [VAT No]) " & _
        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"

        InsertCommand.Connection = DBOperations.GetInstance.Connection

        Dim newParam As New OleDbParameter()
        newParam.SourceColumn = "ActiveY/N"
        newParam.OleDbType = OleDbType.Boolean
        newParam.Size = 2
        Me.InsertCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "AddressLine1"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 40
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
        Me.InsertCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "AddressLine4"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 40
        Me.InsertCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "ContractorName"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 70
        Me.InsertCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "ContractorNo"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 10
        Me.InsertCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "PostalCode"
        newParam.OleDbType = OleDbType.Integer
        Me.InsertCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "Reg No"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 20
        Me.InsertCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "Telephone"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 15
        Me.InsertCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "Vat No"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 20
        Me.InsertCommand.Parameters.Add(newParam)
    End Sub

    ' prepares select query for adapter
    Private Sub PrepareSelectCommand()
        Me.Adapter.SelectCommand = Me.SelectCommand

        ' SelectCommand
        Me.SelectCommand.CommandText = "SELECT * FROM Contractor ORDER BY ContractorNo"

        Me.SelectCommand.Connection = DBOperations.GetInstance.Connection
    End Sub

    ' prepares update query for adapter
    Private Sub PrepareUpdateCommand()
        Me.Adapter.UpdateCommand = Me.UpdateCommand

        ContracterSet.Clear()

        Adapter.Fill(ContracterSet)

        ' UpdateCommand
        Me.UpdateCommand.CommandText =
            "UPDATE Contractor " & _
            "SET [ActiveY/N] = ?, " & _
            "AddressLine1 = ?, " & _
            "AddressLine2 = ?, " & _
            "AddressLine3 = ?, " & _
            "AddressLine4 = ?, " & _
            "ContractorName = ?, " & _
            "ContractorNo = ?, " & _
            "PostalCode = ?, " & _
            "[Reg No] = ?, " & _
            "Telephone = ?, " & _
            "[VAT No] = ? " & _
            "WHERE (ContractorNo = ?) " & _
            "AND ([ActiveY/N] = ?) " & _
            "AND (AddressLine1 = ? OR ? IS NULL AND AddressLine1 IS NULL) " & _
            "AND (AddressLine2 = ? OR ? IS NULL AND AddressLine2 IS NULL) " & _
            "AND (AddressLine3 = ? OR ? IS NULL AND AddressLine3 IS NULL) " & _
            "AND (AddressLine4 = ? OR ? IS NULL AND AddressLine4 IS NULL) " & _
            "AND (ContractorName = ? OR ? IS NULL AND ContractorName IS NULL)" & _
            "AND (PostalCode = ? OR ? IS NULL AND PostalCode IS NULL) " & _
            "AND ([Reg No] = ? OR ? IS NULL AND [Reg No] IS NULL) " & _
            "AND (Telephone = ? OR ? IS NULL AND Telephone IS NULL) " & _
            "AND ([VAT No] = ? OR ? IS NULL AND [VAT No] IS NULL)"

        Me.UpdateCommand.Connection = DBOperations.GetInstance.Connection

        Dim newParam As New OleDbParameter()
        newParam.SourceColumn = "ActiveY/N"
        newParam.OleDbType = OleDbType.Boolean
        newParam.Size = 2
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "AddressLine1"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 40
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
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "AddressLine4"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 40
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "ContractorName"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 70
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "ContractorNo"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 10
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "PostalCode"
        newParam.OleDbType = OleDbType.Integer
        newParam.Size = 0
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "Reg No"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 20
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "Telephone"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 15
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "VAT No"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 20
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "ContractorNo"
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
        newParam.SourceColumn = "ActiveY/N"
        newParam.OleDbType = OleDbType.Boolean
        newParam.Size = 2
        newParam.Direction = ParameterDirection.Input
        newParam.IsNullable = False
        newParam.Precision = CType(0, Byte)
        newParam.Scale = CType(0, Byte)
        newParam.SourceVersion = DataRowVersion.Original
        newParam.Value = Nothing
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "AddressLine1"
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
        newParam.SourceColumn = "AddressLine1"
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
        newParam.SourceColumn = "ContractorName"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 70
        newParam.Direction = ParameterDirection.Input
        newParam.IsNullable = False
        newParam.Precision = CType(0, Byte)
        newParam.Scale = CType(0, Byte)
        newParam.SourceVersion = DataRowVersion.Original
        newParam.Value = Nothing
        Me.UpdateCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "ContractorName"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 70
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
        newParam.SourceColumn = "Reg No"
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
        newParam.SourceColumn = "Reg No"
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
        newParam.SourceColumn = "VAT No"
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
        newParam.SourceColumn = "VAT No"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 20
        newParam.Direction = ParameterDirection.Input
        newParam.IsNullable = False
        newParam.Precision = CType(0, Byte)
        newParam.Scale = CType(0, Byte)
        newParam.SourceVersion = DataRowVersion.Original
        newParam.Value = Nothing
        Me.UpdateCommand.Parameters.Add(newParam)
    End Sub

    ' prepares count companies query for adapter
    Private Sub PrepareCountCommand(ByRef contractorNumber As String)
        Me.CountCommand.CommandText = "SELECT * FROM Contractor WHERE (ContractorNo = ?) ORDER BY ContractorNo"

        Me.CountCommand.Connection = DBOperations.GetInstance.Connection

        Dim newParam As New OleDbParameter("contractorNumber", contractorNumber)
        newParam.SourceColumn = "ContractorNo"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 10
        Me.CountCommand.Parameters.Add(newParam)
    End Sub

    ' map table to values
    Private Sub MapTable()
        Me.Adapter.TableMappings.AddRange(New DataTableMapping() {
                                  New DataTableMapping("Table", "Contractor",
                                                       New DataColumnMapping() {
                                                           New DataColumnMapping("ActiveY/N", "ActiveY/N"),
                                                           New DataColumnMapping("AddressLine1", "AddressLine1"),
                                                           New DataColumnMapping("AddressLine2", "AddressLine2"),
                                                           New DataColumnMapping("AddressLine3", "AddressLine3"),
                                                           New DataColumnMapping("AddressLine4", "AddressLine4"),
                                                           New DataColumnMapping("ContractorName", "ContractorName"),
                                                           New DataColumnMapping("ContractorNo", "ContractorNo"),
                                                           New DataColumnMapping("PostalCode", "PostalCode"),
                                                           New DataColumnMapping("Reg No", "Reg No"),
                                                           New DataColumnMapping("Telephone", "Telephone"),
                                                           New DataColumnMapping("VAT No", "VAT No")})})
    End Sub

    ' updates a single row in the database
    Private Sub UpdateRowValues(ByRef row As DataRow, ByRef contractorNumber As String, ByRef contractorName As String, ByRef addressLine1 As String, ByRef addressLine2 As String, ByRef addressLine3 As String, ByRef addressLine4 As String, ByVal postalCode As Integer, ByRef telephone As String, ByVal isActive As Boolean, ByRef vatNumber As String, ByRef regNumber As String)
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