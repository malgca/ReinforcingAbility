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

    Private Property cmdCountCompNo As New OleDbCommand

    Public Sub New()
        ' initialize adapter
        Me.Adapter.InsertCommand = Me.InsertCommand
        Me.Adapter.SelectCommand = Me.SelectCommand
        Me.Adapter.UpdateCommand = Me.UpdateCommand
        Me.Adapter.DeleteCommand = Me.DeleteCommand

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

        'Delete Command
        Me.DeleteCommand.CommandText =
            "DELETE FROM Company " & _
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

        Me.DeleteCommand.Connection = DBOperations.GetInstance.Connection

        Me.DeleteCommand.Parameters.Add(New OleDbParameter("Original_CompanyNo", OleDbType.VarWChar, 10, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CompanyNo", DataRowVersion.Original, Nothing))

        Me.DeleteCommand.Parameters.Add(New OleDbParameter("Original_AddressLine2", OleDbType.VarWChar, 40, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine2", DataRowVersion.Original, Nothing))

        Me.DeleteCommand.Parameters.Add(New OleDbParameter("Original_AddressLine21", OleDbType.VarWChar, 40, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine2", DataRowVersion.Original, Nothing))

        Me.DeleteCommand.Parameters.Add(New OleDbParameter("Original_AddressLine3", OleDbType.VarWChar, 40, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine3", DataRowVersion.Original, Nothing))

        Me.DeleteCommand.Parameters.Add(New OleDbParameter("Original_AddressLine31", OleDbType.VarWChar, 40, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine3", DataRowVersion.Original, Nothing))

        Me.DeleteCommand.Parameters.Add(New OleDbParameter("Original_AddressLine4", OleDbType.VarWChar, 40, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine4", DataRowVersion.Original, Nothing))

        Me.DeleteCommand.Parameters.Add(New OleDbParameter("Original_AddressLine41", OleDbType.VarWChar, 40, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine4", DataRowVersion.Original, Nothing))

        Me.DeleteCommand.Parameters.Add(New OleDbParameter("Original_CompanyName", OleDbType.VarWChar, 40, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CompanyName", DataRowVersion.Original, Nothing))

        Me.DeleteCommand.Parameters.Add(New OleDbParameter("Original_CompanyName1", OleDbType.VarWChar, 40, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CompanyName", DataRowVersion.Original, Nothing))

        Me.DeleteCommand.Parameters.Add(New OleDbParameter("Original_Email", OleDbType.VarWChar, 40, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Email", DataRowVersion.Original, Nothing))

        Me.DeleteCommand.Parameters.Add(New OleDbParameter("Original_Email1", OleDbType.VarWChar, 40, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Email", DataRowVersion.Original, Nothing))

        Me.DeleteCommand.Parameters.Add(New OleDbParameter("Original_Fax", OleDbType.VarWChar, 15, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Fax", DataRowVersion.Original, Nothing))

        Me.DeleteCommand.Parameters.Add(New OleDbParameter("Original_Fax1", OleDbType.VarWChar, 15, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Fax", DataRowVersion.Original, Nothing))

        Me.DeleteCommand.Parameters.Add(New OleDbParameter("Original_LastCutNum", OleDbType.Integer, 0, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "LastCutNum", DataRowVersion.Original, Nothing))

        Me.DeleteCommand.Parameters.Add(New OleDbParameter("Original_LastCutNum1", OleDbType.Integer, 0, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "LastCutNum", DataRowVersion.Original, Nothing))

        Me.DeleteCommand.Parameters.Add(New OleDbParameter("Original_LastInvNum", OleDbType.Integer, 0, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "LastInvNum", DataRowVersion.Original, Nothing))

        Me.DeleteCommand.Parameters.Add(New OleDbParameter("Original_LastInvNum1", OleDbType.Integer, 0, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "LastInvNum", DataRowVersion.Original, Nothing))

        Me.DeleteCommand.Parameters.Add(New OleDbParameter("Original_Message", OleDbType.VarWChar, 200, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Message", DataRowVersion.Original, Nothing))

        Me.DeleteCommand.Parameters.Add(New OleDbParameter("Original_Message1", OleDbType.VarWChar, 200, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Message", DataRowVersion.Original, Nothing))

        Me.DeleteCommand.Parameters.Add(New OleDbParameter("Original_PostalCode", OleDbType.Integer, 0, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "PostalCode", DataRowVersion.Original, Nothing))

        Me.DeleteCommand.Parameters.Add(New OleDbParameter("Original_PostalCode1", OleDbType.Integer, 0, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "PostalCode", DataRowVersion.Original, Nothing))

        Me.DeleteCommand.Parameters.Add(New OleDbParameter("Original_RegNo", OleDbType.VarWChar, 20, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "RegNo", DataRowVersion.Original, Nothing))

        Me.DeleteCommand.Parameters.Add(New OleDbParameter("Original_RegNo1", OleDbType.VarWChar, 20, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "RegNo", DataRowVersion.Original, Nothing))

        Me.DeleteCommand.Parameters.Add(New OleDbParameter("Original_Telephone", OleDbType.VarWChar, 15, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Telephone", DataRowVersion.Original, Nothing))

        Me.DeleteCommand.Parameters.Add(New OleDbParameter("Original_Telephone1", OleDbType.VarWChar, 15, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Telephone", DataRowVersion.Original, Nothing))

        Me.DeleteCommand.Parameters.Add(New OleDbParameter("Original_UnitOfMeas", OleDbType.VarWChar, 50, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "UnitOfMeas", DataRowVersion.Original, Nothing))

        Me.DeleteCommand.Parameters.Add(New OleDbParameter("Original_UnitOfMeas1", OleDbType.VarWChar, 50, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "UnitOfMeas", DataRowVersion.Original, Nothing))

        Me.DeleteCommand.Parameters.Add(New OleDbParameter("Original_VatNo", OleDbType.VarWChar, 20, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "VatNo", DataRowVersion.Original, Nothing))

        Me.DeleteCommand.Parameters.Add(New OleDbParameter("Original_VatNo1", OleDbType.VarWChar, 20, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "VatNo", DataRowVersion.Original, Nothing))

        Me.DeleteCommand.Parameters.Add(New OleDbParameter("Original_VatPerc", OleDbType.Decimal, 0, ParameterDirection.Input, False, CType(2, Byte), CType(2, Byte), "VatPerc", DataRowVersion.Original, Nothing))

        Me.DeleteCommand.Parameters.Add(New OleDbParameter("Original_VatPerc1", OleDbType.Decimal, 0, ParameterDirection.Input, False, CType(2, Byte), CType(2, Byte), "VatPerc", DataRowVersion.Original, Nothing))

        Me.DeleteCommand.Parameters.Add(New OleDbParameter("Original_Website", OleDbType.VarWChar, 30, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Website", DataRowVersion.Original, Nothing))

        Me.DeleteCommand.Parameters.Add(New OleDbParameter("Original_Website1", OleDbType.VarWChar, 30, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Website", DataRowVersion.Original, Nothing))

        'InsertCommand
        Me.InsertCommand.CommandText =
            "INSERT INTO Company" & _
            "(Address, AddressLine2, AddressLine3, AddressLine4, CompanyName, CompanyNo, Email, Fax, LastCutNum, LastInvNum, Message, PostalCode, RegNo, Telephone, UnitOfMeas, VatNo, VatPerc, Website) " & _
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"

        Me.InsertCommand.Connection = DBOperations.GetInstance.Connection

        Me.InsertCommand.Parameters.Add(New OleDbParameter("Address", OleDbType.VarWChar, 0, "Address"))

        Me.InsertCommand.Parameters.Add(New OleDbParameter("AddressLine2", OleDbType.VarWChar, 40, "AddressLine2"))

        Me.InsertCommand.Parameters.Add(New OleDbParameter("AddressLine3", OleDbType.VarWChar, 40, "AddressLine3"))

        Me.InsertCommand.Parameters.Add(New OleDbParameter("AddressLine4", OleDbType.VarWChar, 40, "AddressLine4"))

        Me.InsertCommand.Parameters.Add(New OleDbParameter("CompanyName", OleDbType.VarWChar, 40, "CompanyName"))

        Me.InsertCommand.Parameters.Add(New OleDbParameter("CompanyNo", OleDbType.VarWChar, 10, "CompanyNo"))

        Me.InsertCommand.Parameters.Add(New OleDbParameter("Email", OleDbType.VarWChar, 40, "Email"))

        Me.InsertCommand.Parameters.Add(New OleDbParameter("Fax", OleDbType.VarWChar, 15, "Fax"))

        Me.InsertCommand.Parameters.Add(New OleDbParameter("LastCutNum", OleDbType.Integer, 0, "LastCutNum"))

        Me.InsertCommand.Parameters.Add(New OleDbParameter("LastInvNum", OleDbType.Integer, 0, "LastInvNum"))

        Me.InsertCommand.Parameters.Add(New OleDbParameter("Message", OleDbType.VarWChar, 200, "Message"))

        Me.InsertCommand.Parameters.Add(New OleDbParameter("PostalCode", OleDbType.Integer, 0, "PostalCode"))

        Me.InsertCommand.Parameters.Add(New OleDbParameter("RegNo", OleDbType.VarWChar, 20, "RegNo"))

        Me.InsertCommand.Parameters.Add(New OleDbParameter("Telephone", OleDbType.VarWChar, 15, "Telephone"))

        Me.InsertCommand.Parameters.Add(New OleDbParameter("UnitOfMeas", OleDbType.VarWChar, 50, "UnitOfMeas"))

        Me.InsertCommand.Parameters.Add(New OleDbParameter("VatNo", OleDbType.VarWChar, 20, "VatNo"))

        Me.InsertCommand.Parameters.Add(New OleDbParameter("VatPerc", OleDbType.Decimal, 0, ParameterDirection.Input, False, CType(2, Byte), CType(2, Byte), "VatPerc", DataRowVersion.Current, Nothing))

        Me.InsertCommand.Parameters.Add(New OleDbParameter("Website", OleDbType.VarWChar, 30, "Website"))

        'SelectCommand
        Me.SelectCommand.CommandText = "SELECT Address, AddressLine2, AddressLine3, AddressLine4, CompanyName, CompanyNo, Email, Fax, LastCutNum, LastInvNum, Message, PostalCode, RegNo, Telephone, UnitOfMeas, VatNo, VatPerc, Website FROM Company"

        Me.SelectCommand.Connection = DBOperations.GetInstance.Connection

        'UpdateCommand
        Me.UpdateCommand.CommandText =
            "UPDATE Company " & _
            "SET Address = ?, " & _
            "AddressLine2 = ?, " & _
            "AddressLine3 = ?, " & _
            "AddressLine4 = ?, " & _
            "CompanyName = ?, " & _
            "CompanyNo = ?, " & _
            "Email = ?, Fax = ?, " & _
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

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("Address", OleDbType.VarWChar, 0, "Address"))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("AddressLine2", OleDbType.VarWChar, 40, "AddressLine2"))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("AddressLine3", OleDbType.VarWChar, 40, "AddressLine3"))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("AddressLine4", OleDbType.VarWChar, 40, "AddressLine4"))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("CompanyName", OleDbType.VarWChar, 40, "CompanyName"))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("CompanyNo", OleDbType.VarWChar, 10, "CompanyNo"))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("Email", OleDbType.VarWChar, 40, "Email"))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("Fax", OleDbType.VarWChar, 15, "Fax"))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("LastCutNum", OleDbType.Integer, 0, "LastCutNum"))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("LastInvNum", OleDbType.Integer, 0, "LastInvNum"))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("Message", OleDbType.VarWChar, 200, "Message"))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("PostalCode", OleDbType.Integer, 0, "PostalCode"))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("RegNo", OleDbType.VarWChar, 20, "RegNo"))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("Telephone", OleDbType.VarWChar, 15, "Telephone"))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("UnitOfMeas", OleDbType.VarWChar, 50, "UnitOfMeas"))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("VatNo", OleDbType.VarWChar, 20, "VatNo"))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("VatPerc", OleDbType.Decimal, 0, ParameterDirection.Input, False, CType(2, Byte), CType(2, Byte), "VatPerc", DataRowVersion.Current, Nothing))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("Website", OleDbType.VarWChar, 30, "Website"))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("Original_CompanyNo", OleDbType.VarWChar, 10, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CompanyNo", DataRowVersion.Original, Nothing))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("Original_AddressLine2", OleDbType.VarWChar, 40, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine2", DataRowVersion.Original, Nothing))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("Original_AddressLine21", OleDbType.VarWChar, 40, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine2", DataRowVersion.Original, Nothing))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("Original_AddressLine3", OleDbType.VarWChar, 40, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine3", DataRowVersion.Original, Nothing))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("Original_AddressLine31", OleDbType.VarWChar, 40, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine3", DataRowVersion.Original, Nothing))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("Original_AddressLine4", OleDbType.VarWChar, 40, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine4", DataRowVersion.Original, Nothing))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("Original_AddressLine41", OleDbType.VarWChar, 40, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine4", DataRowVersion.Original, Nothing))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("Original_CompanyName", OleDbType.VarWChar, 40, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CompanyName", DataRowVersion.Original, Nothing))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("Original_CompanyName1", OleDbType.VarWChar, 40, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CompanyName", DataRowVersion.Original, Nothing))
        Me.UpdateCommand.Parameters.Add(New OleDbParameter("Original_Email", OleDbType.VarWChar, 40, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Email", DataRowVersion.Original, Nothing))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("Original_Email1", OleDbType.VarWChar, 40, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Email", DataRowVersion.Original, Nothing))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("Original_Fax", OleDbType.VarWChar, 15, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Fax", DataRowVersion.Original, Nothing))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("Original_Fax1", OleDbType.VarWChar, 15, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Fax", DataRowVersion.Original, Nothing))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("Original_LastCutNum", OleDbType.Integer, 0, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "LastCutNum", DataRowVersion.Original, Nothing))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("Original_LastCutNum1", OleDbType.Integer, 0, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "LastCutNum", DataRowVersion.Original, Nothing))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("Original_LastInvNum", OleDbType.Integer, 0, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "LastInvNum", DataRowVersion.Original, Nothing))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("Original_LastInvNum1", OleDbType.Integer, 0, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "LastInvNum", DataRowVersion.Original, Nothing))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("Original_Message", OleDbType.VarWChar, 200, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Message", DataRowVersion.Original, Nothing))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("Original_Message1", OleDbType.VarWChar, 200, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Message", DataRowVersion.Original, Nothing))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("Original_PostalCode", OleDbType.Integer, 0, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "PostalCode", DataRowVersion.Original, Nothing))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("Original_PostalCode1", OleDbType.Integer, 0, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "PostalCode", DataRowVersion.Original, Nothing))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("Original_RegNo", OleDbType.VarWChar, 20, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "RegNo", DataRowVersion.Original, Nothing))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("Original_RegNo1", OleDbType.VarWChar, 20, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "RegNo", DataRowVersion.Original, Nothing))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("Original_Telephone", OleDbType.VarWChar, 15, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Telephone", DataRowVersion.Original, Nothing))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("Original_Telephone1", OleDbType.VarWChar, 15, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Telephone", DataRowVersion.Original, Nothing))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("Original_UnitOfMeas", OleDbType.VarWChar, 50, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "UnitOfMeas", DataRowVersion.Original, Nothing))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("Original_UnitOfMeas1", OleDbType.VarWChar, 50, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "UnitOfMeas", DataRowVersion.Original, Nothing))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("Original_VatNo", OleDbType.VarWChar, 20, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "VatNo", DataRowVersion.Original, Nothing))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("Original_VatNo1", OleDbType.VarWChar, 20, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "VatNo", DataRowVersion.Original, Nothing))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("Original_VatPerc", OleDbType.Decimal, 0, ParameterDirection.Input, False, CType(2, Byte), CType(2, Byte), "VatPerc", DataRowVersion.Original, Nothing))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("Original_VatPerc1", OleDbType.Decimal, 0, ParameterDirection.Input, False, CType(2, Byte), CType(2, Byte), "VatPerc", DataRowVersion.Original, Nothing))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("Original_Website", OleDbType.VarWChar, 30, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Website", DataRowVersion.Original, Nothing))

        Me.UpdateCommand.Parameters.Add(New OleDbParameter("Original_Website1", OleDbType.VarWChar, 30, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Website", DataRowVersion.Original, Nothing))


        Me.cmdCountCompNo.CommandText = "SELECT Company.* FROM Company WHERE (CompanyNo = ?)"

        Me.cmdCountCompNo.Connection = DBOperations.GetInstance.Connection

        Me.cmdCountCompNo.Parameters.Add(New OleDbParameter("CompanyNo", OleDbType.VarWChar, 10, "CompanyNo"))
    End Sub

    ''' <summary>
    ''' Gets the number of companies currently in the company table
    ''' </summary>
    ''' <param name="CompanyNumber">ID to identify a given company</param>
    ''' <param name="count">Count of existing companies</param>
    Public Sub GetNumberOfCompanies(ByRef CompanyNumber As String, ByRef count As Integer)
        DBOperations.GetInstance.Connection.Open()
        cmdCountCompNo.Parameters("CompanyNo").Value = CompanyNumber

        Dim dataReader = cmdCountCompNo.ExecuteReader(CommandBehavior.CloseConnection)

        While dataReader.Read()
            count += 1
        End While

        dataReader.Close()
        DBOperations.GetInstance.Connection.Close()
    End Sub
    ''' adds a new row to the company table
    Public Sub AddCompanyRow(ByRef companyNumber As String, ByRef companyName As String, ByRef regNumber As String, ByRef vatNumber As String, ByRef addressLine1 As String, ByRef addressLine2 As String, ByRef addressLine3 As String, ByRef postalCode As Integer, ByRef telephone As String, ByRef email As String, ByRef fax As String, ByRef website As String, ByRef message As String, ByRef vatPercentage As Double, ByRef lastInvoiceNumber As Integer)
        Dim companySet As New DataSet

        Adapter.Fill(companySet)

        Dim row = companySet.Tables.Item(0).NewRow()

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

        companySet.Tables.Item(0).Rows.Add(row)
        Adapter.Update(companySet.Tables.Item(0))
    End Sub

    ''' <summary>
    ''' Determines whether a parameter string is empty or not
    ''' </summary>
    Private Function IsNotEmpty(ByVal parameter As String) As Boolean
        Return Not parameter.Equals(String.Empty)
    End Function
End Class