Public Class frmContractor
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents grpContractorDetails As System.Windows.Forms.GroupBox
    Friend WithEvents lblContractorNo As System.Windows.Forms.Label
    Friend WithEvents txtContractorNo As System.Windows.Forms.TextBox
    Friend WithEvents lblContractorName As System.Windows.Forms.Label
    Friend WithEvents txtContractorName As System.Windows.Forms.TextBox
    Friend WithEvents lblAddress1 As System.Windows.Forms.Label
    Friend WithEvents txtAddress2 As System.Windows.Forms.TextBox
    Friend WithEvents lblAddress2 As System.Windows.Forms.Label
    Friend WithEvents txtAddress3 As System.Windows.Forms.TextBox
    Friend WithEvents lblAddress3 As System.Windows.Forms.Label
    Friend WithEvents txtAddress4 As System.Windows.Forms.TextBox
    Friend WithEvents lblAddress4 As System.Windows.Forms.Label
    Friend WithEvents lblPostalCode As System.Windows.Forms.Label
    Friend WithEvents lblTelNo As System.Windows.Forms.Label
    Friend WithEvents txtTelNo As System.Windows.Forms.TextBox
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents txtAddress1 As System.Windows.Forms.TextBox
    Friend WithEvents txtPostalCode As System.Windows.Forms.TextBox
    Friend WithEvents cmdCountContractorNo As System.Data.OleDb.OleDbCommand
    Friend WithEvents txtVATNo As System.Windows.Forms.TextBox
    Friend WithEvents lblVATNo As System.Windows.Forms.Label
    Friend WithEvents OleDbSelectCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbInsertCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbUpdateCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbDeleteCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents conContractor As System.Data.OleDb.OleDbConnection
    Friend WithEvents adpContractor As System.Data.OleDb.OleDbDataAdapter
    Friend WithEvents dsContractor As Reinforcing_Ability.dsReinforcingAbility
    Friend WithEvents cbxCompNo As System.Windows.Forms.ComboBox
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.grpContractorDetails = New System.Windows.Forms.GroupBox
        Me.txtVATNo = New System.Windows.Forms.TextBox
        Me.lblVATNo = New System.Windows.Forms.Label
        Me.txtTelNo = New System.Windows.Forms.TextBox
        Me.lblTelNo = New System.Windows.Forms.Label
        Me.txtPostalCode = New System.Windows.Forms.TextBox
        Me.lblPostalCode = New System.Windows.Forms.Label
        Me.txtAddress4 = New System.Windows.Forms.TextBox
        Me.lblAddress4 = New System.Windows.Forms.Label
        Me.txtAddress3 = New System.Windows.Forms.TextBox
        Me.lblAddress3 = New System.Windows.Forms.Label
        Me.txtAddress2 = New System.Windows.Forms.TextBox
        Me.lblAddress2 = New System.Windows.Forms.Label
        Me.txtAddress1 = New System.Windows.Forms.TextBox
        Me.lblAddress1 = New System.Windows.Forms.Label
        Me.txtContractorName = New System.Windows.Forms.TextBox
        Me.lblContractorName = New System.Windows.Forms.Label
        Me.txtContractorNo = New System.Windows.Forms.TextBox
        Me.lblContractorNo = New System.Windows.Forms.Label
        Me.btnAdd = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.cmdCountContractorNo = New System.Data.OleDb.OleDbCommand
        Me.conContractor = New System.Data.OleDb.OleDbConnection
        Me.adpContractor = New System.Data.OleDb.OleDbDataAdapter
        Me.OleDbDeleteCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbInsertCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbSelectCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbUpdateCommand1 = New System.Data.OleDb.OleDbCommand
        Me.dsContractor = New Reinforcing_Ability.dsReinforcingAbility
        Me.cbxCompNo = New System.Windows.Forms.ComboBox
        Me.btnEdit = New System.Windows.Forms.Button
        Me.grpContractorDetails.SuspendLayout()
        CType(Me.dsContractor, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'grpContractorDetails
        '
        Me.grpContractorDetails.Controls.Add(Me.txtVATNo)
        Me.grpContractorDetails.Controls.Add(Me.lblVATNo)
        Me.grpContractorDetails.Controls.Add(Me.txtTelNo)
        Me.grpContractorDetails.Controls.Add(Me.lblTelNo)
        Me.grpContractorDetails.Controls.Add(Me.txtPostalCode)
        Me.grpContractorDetails.Controls.Add(Me.lblPostalCode)
        Me.grpContractorDetails.Controls.Add(Me.txtAddress4)
        Me.grpContractorDetails.Controls.Add(Me.lblAddress4)
        Me.grpContractorDetails.Controls.Add(Me.txtAddress3)
        Me.grpContractorDetails.Controls.Add(Me.lblAddress3)
        Me.grpContractorDetails.Controls.Add(Me.txtAddress2)
        Me.grpContractorDetails.Controls.Add(Me.lblAddress2)
        Me.grpContractorDetails.Controls.Add(Me.txtAddress1)
        Me.grpContractorDetails.Controls.Add(Me.lblAddress1)
        Me.grpContractorDetails.Controls.Add(Me.txtContractorName)
        Me.grpContractorDetails.Controls.Add(Me.lblContractorName)
        Me.grpContractorDetails.Controls.Add(Me.txtContractorNo)
        Me.grpContractorDetails.Controls.Add(Me.lblContractorNo)
        Me.grpContractorDetails.Location = New System.Drawing.Point(16, 16)
        Me.grpContractorDetails.Name = "grpContractorDetails"
        Me.grpContractorDetails.Size = New System.Drawing.Size(496, 320)
        Me.grpContractorDetails.TabIndex = 0
        Me.grpContractorDetails.TabStop = False
        Me.grpContractorDetails.Text = "Contractor Details"
        '
        'txtVATNo
        '
        Me.txtVATNo.Location = New System.Drawing.Point(128, 96)
        Me.txtVATNo.MaxLength = 20
        Me.txtVATNo.Name = "txtVATNo"
        Me.txtVATNo.Size = New System.Drawing.Size(136, 20)
        Me.txtVATNo.TabIndex = 5
        Me.txtVATNo.Text = ""
        '
        'lblVATNo
        '
        Me.lblVATNo.Location = New System.Drawing.Point(72, 96)
        Me.lblVATNo.Name = "lblVATNo"
        Me.lblVATNo.Size = New System.Drawing.Size(48, 16)
        Me.lblVATNo.TabIndex = 4
        Me.lblVATNo.Text = "VAT No."
        Me.lblVATNo.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtTelNo
        '
        Me.txtTelNo.Location = New System.Drawing.Point(128, 288)
        Me.txtTelNo.MaxLength = 15
        Me.txtTelNo.Name = "txtTelNo"
        Me.txtTelNo.Size = New System.Drawing.Size(120, 20)
        Me.txtTelNo.TabIndex = 17
        Me.txtTelNo.Text = ""
        '
        'lblTelNo
        '
        Me.lblTelNo.Location = New System.Drawing.Point(24, 288)
        Me.lblTelNo.Name = "lblTelNo"
        Me.lblTelNo.Size = New System.Drawing.Size(100, 16)
        Me.lblTelNo.TabIndex = 16
        Me.lblTelNo.Text = "Tel. No."
        Me.lblTelNo.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtPostalCode
        '
        Me.txtPostalCode.Location = New System.Drawing.Point(128, 256)
        Me.txtPostalCode.MaxLength = 5
        Me.txtPostalCode.Name = "txtPostalCode"
        Me.txtPostalCode.Size = New System.Drawing.Size(64, 20)
        Me.txtPostalCode.TabIndex = 15
        Me.txtPostalCode.Text = ""
        '
        'lblPostalCode
        '
        Me.lblPostalCode.Location = New System.Drawing.Point(24, 256)
        Me.lblPostalCode.Name = "lblPostalCode"
        Me.lblPostalCode.Size = New System.Drawing.Size(100, 16)
        Me.lblPostalCode.TabIndex = 14
        Me.lblPostalCode.Text = "Postal Code"
        Me.lblPostalCode.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtAddress4
        '
        Me.txtAddress4.Location = New System.Drawing.Point(128, 224)
        Me.txtAddress4.MaxLength = 40
        Me.txtAddress4.Name = "txtAddress4"
        Me.txtAddress4.Size = New System.Drawing.Size(216, 20)
        Me.txtAddress4.TabIndex = 13
        Me.txtAddress4.Text = ""
        '
        'lblAddress4
        '
        Me.lblAddress4.Location = New System.Drawing.Point(24, 224)
        Me.lblAddress4.Name = "lblAddress4"
        Me.lblAddress4.Size = New System.Drawing.Size(100, 16)
        Me.lblAddress4.TabIndex = 12
        Me.lblAddress4.Text = "Address 4"
        Me.lblAddress4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtAddress3
        '
        Me.txtAddress3.Location = New System.Drawing.Point(128, 192)
        Me.txtAddress3.MaxLength = 40
        Me.txtAddress3.Name = "txtAddress3"
        Me.txtAddress3.Size = New System.Drawing.Size(216, 20)
        Me.txtAddress3.TabIndex = 11
        Me.txtAddress3.Text = ""
        '
        'lblAddress3
        '
        Me.lblAddress3.Location = New System.Drawing.Point(24, 192)
        Me.lblAddress3.Name = "lblAddress3"
        Me.lblAddress3.Size = New System.Drawing.Size(100, 16)
        Me.lblAddress3.TabIndex = 10
        Me.lblAddress3.Text = "Address 3"
        Me.lblAddress3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtAddress2
        '
        Me.txtAddress2.Location = New System.Drawing.Point(128, 160)
        Me.txtAddress2.MaxLength = 40
        Me.txtAddress2.Name = "txtAddress2"
        Me.txtAddress2.Size = New System.Drawing.Size(216, 20)
        Me.txtAddress2.TabIndex = 9
        Me.txtAddress2.Text = ""
        '
        'lblAddress2
        '
        Me.lblAddress2.Location = New System.Drawing.Point(24, 160)
        Me.lblAddress2.Name = "lblAddress2"
        Me.lblAddress2.Size = New System.Drawing.Size(100, 16)
        Me.lblAddress2.TabIndex = 8
        Me.lblAddress2.Text = "Address 2"
        Me.lblAddress2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtAddress1
        '
        Me.txtAddress1.Location = New System.Drawing.Point(128, 128)
        Me.txtAddress1.MaxLength = 40
        Me.txtAddress1.Name = "txtAddress1"
        Me.txtAddress1.Size = New System.Drawing.Size(216, 20)
        Me.txtAddress1.TabIndex = 7
        Me.txtAddress1.Text = ""
        '
        'lblAddress1
        '
        Me.lblAddress1.Location = New System.Drawing.Point(24, 128)
        Me.lblAddress1.Name = "lblAddress1"
        Me.lblAddress1.Size = New System.Drawing.Size(100, 16)
        Me.lblAddress1.TabIndex = 6
        Me.lblAddress1.Text = "Address 1"
        Me.lblAddress1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtContractorName
        '
        Me.txtContractorName.Location = New System.Drawing.Point(128, 64)
        Me.txtContractorName.MaxLength = 70
        Me.txtContractorName.Name = "txtContractorName"
        Me.txtContractorName.Size = New System.Drawing.Size(344, 20)
        Me.txtContractorName.TabIndex = 3
        Me.txtContractorName.Text = ""
        '
        'lblContractorName
        '
        Me.lblContractorName.Location = New System.Drawing.Point(24, 64)
        Me.lblContractorName.Name = "lblContractorName"
        Me.lblContractorName.Size = New System.Drawing.Size(100, 16)
        Me.lblContractorName.TabIndex = 2
        Me.lblContractorName.Text = "Contractor Name"
        Me.lblContractorName.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtContractorNo
        '
        Me.txtContractorNo.Location = New System.Drawing.Point(128, 32)
        Me.txtContractorNo.MaxLength = 10
        Me.txtContractorNo.Name = "txtContractorNo"
        Me.txtContractorNo.Size = New System.Drawing.Size(88, 20)
        Me.txtContractorNo.TabIndex = 1
        Me.txtContractorNo.Text = ""
        '
        'lblContractorNo
        '
        Me.lblContractorNo.Location = New System.Drawing.Point(24, 32)
        Me.lblContractorNo.Name = "lblContractorNo"
        Me.lblContractorNo.Size = New System.Drawing.Size(100, 16)
        Me.lblContractorNo.TabIndex = 0
        Me.lblContractorNo.Text = "Contractor No."
        Me.lblContractorNo.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnAdd
        '
        Me.btnAdd.Location = New System.Drawing.Point(16, 352)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.TabIndex = 1
        Me.btnAdd.Text = "Add"
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(352, 352)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.TabIndex = 3
        Me.btnSave.Text = "Save"
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(440, 352)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.TabIndex = 4
        Me.btnClose.Text = "Close"
        '
        'cmdCountContractorNo
        '
        Me.cmdCountContractorNo.CommandText = "SELECT Contractor.* FROM Contractor WHERE (ContractorNo = ?) ORDER BY ContractorNo"
        Me.cmdCountContractorNo.Connection = Me.conContractor
        Me.cmdCountContractorNo.Parameters.Add(New System.Data.OleDb.OleDbParameter("ContractorNo", System.Data.OleDb.OleDbType.VarWChar, 10, "ContractorNo"))
        '
        'conContractor
        '
        Me.conContractor.ConnectionString = "Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Registry Path=;Jet OLEDB:Database L" & _
        "ocking Mode=1;Data Source=""winsteelVers5.mdb"";Mode=Share Deny None;Jet OLEDB:Eng" & _
        "ine Type=5;Provider=""Microsoft.Jet.OLEDB.4.0"";Jet OLEDB:System database=;Jet OLE" & _
        "DB:SFP=False;persist security info=False;Extended Properties=;Jet OLEDB:Compact " & _
        "Without Replica Repair=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Create S" & _
        "ystem Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;User ID=Admin;" & _
        "Jet OLEDB:Global Bulk Transactions=1"
        '
        'adpContractor
        '
        Me.adpContractor.DeleteCommand = Me.OleDbDeleteCommand1
        Me.adpContractor.InsertCommand = Me.OleDbInsertCommand1
        Me.adpContractor.SelectCommand = Me.OleDbSelectCommand1
        Me.adpContractor.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "Contractor", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("ActiveY/N", "ActiveY/N"), New System.Data.Common.DataColumnMapping("AddressLine1", "AddressLine1"), New System.Data.Common.DataColumnMapping("AddressLine2", "AddressLine2"), New System.Data.Common.DataColumnMapping("AddressLine3", "AddressLine3"), New System.Data.Common.DataColumnMapping("AddressLine4", "AddressLine4"), New System.Data.Common.DataColumnMapping("ContractorName", "ContractorName"), New System.Data.Common.DataColumnMapping("ContractorNo", "ContractorNo"), New System.Data.Common.DataColumnMapping("PostalCode", "PostalCode"), New System.Data.Common.DataColumnMapping("Reg No", "Reg No"), New System.Data.Common.DataColumnMapping("Telephone", "Telephone"), New System.Data.Common.DataColumnMapping("VAT No", "VAT No")})})
        Me.adpContractor.UpdateCommand = Me.OleDbUpdateCommand1
        '
        'OleDbDeleteCommand1
        '
        Me.OleDbDeleteCommand1.CommandText = "DELETE FROM Contractor WHERE (ContractorNo = ?) AND ([ActiveY/N] = ?) AND (Addres" & _
        "sLine1 = ? OR ? IS NULL AND AddressLine1 IS NULL) AND (AddressLine2 = ? OR ? IS " & _
        "NULL AND AddressLine2 IS NULL) AND (AddressLine3 = ? OR ? IS NULL AND AddressLin" & _
        "e3 IS NULL) AND (AddressLine4 = ? OR ? IS NULL AND AddressLine4 IS NULL) AND (Co" & _
        "ntractorName = ? OR ? IS NULL AND ContractorName IS NULL) AND (PostalCode = ? OR" & _
        " ? IS NULL AND PostalCode IS NULL) AND ([Reg No] = ? OR ? IS NULL AND [Reg No] I" & _
        "S NULL) AND (Telephone = ? OR ? IS NULL AND Telephone IS NULL) AND ([VAT No] = ?" & _
        " OR ? IS NULL AND [VAT No] IS NULL)"
        Me.OleDbDeleteCommand1.Connection = Me.conContractor
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ContractorNo", System.Data.OleDb.OleDbType.VarWChar, 10, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ContractorNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ActiveY_N", System.Data.OleDb.OleDbType.Boolean, 2, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ActiveY/N", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine1", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine1", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine11", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine1", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine2", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine2", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine21", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine2", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine3", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine3", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine31", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine3", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine4", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine4", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine41", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine4", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ContractorName", System.Data.OleDb.OleDbType.VarWChar, 70, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ContractorName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ContractorName1", System.Data.OleDb.OleDbType.VarWChar, 70, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ContractorName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_PostalCode", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "PostalCode", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_PostalCode1", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "PostalCode", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Reg_No", System.Data.OleDb.OleDbType.VarWChar, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Reg No", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Reg_No1", System.Data.OleDb.OleDbType.VarWChar, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Reg No", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Telephone", System.Data.OleDb.OleDbType.VarWChar, 15, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Telephone", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Telephone1", System.Data.OleDb.OleDbType.VarWChar, 15, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Telephone", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_VAT_No", System.Data.OleDb.OleDbType.VarWChar, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "VAT No", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_VAT_No1", System.Data.OleDb.OleDbType.VarWChar, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "VAT No", System.Data.DataRowVersion.Original, Nothing))
        '
        'OleDbInsertCommand1
        '
        Me.OleDbInsertCommand1.CommandText = "INSERT INTO Contractor([ActiveY/N], AddressLine1, AddressLine2, AddressLine3, Add" & _
        "ressLine4, ContractorName, ContractorNo, PostalCode, [Reg No], Telephone, [VAT N" & _
        "o]) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
        Me.OleDbInsertCommand1.Connection = Me.conContractor
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("ActiveY_N", System.Data.OleDb.OleDbType.Boolean, 2, "ActiveY/N"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("AddressLine1", System.Data.OleDb.OleDbType.VarWChar, 40, "AddressLine1"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("AddressLine2", System.Data.OleDb.OleDbType.VarWChar, 40, "AddressLine2"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("AddressLine3", System.Data.OleDb.OleDbType.VarWChar, 40, "AddressLine3"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("AddressLine4", System.Data.OleDb.OleDbType.VarWChar, 40, "AddressLine4"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("ContractorName", System.Data.OleDb.OleDbType.VarWChar, 70, "ContractorName"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("ContractorNo", System.Data.OleDb.OleDbType.VarWChar, 10, "ContractorNo"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("PostalCode", System.Data.OleDb.OleDbType.Integer, 0, "PostalCode"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Reg_No", System.Data.OleDb.OleDbType.VarWChar, 20, "Reg No"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Telephone", System.Data.OleDb.OleDbType.VarWChar, 15, "Telephone"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("VAT_No", System.Data.OleDb.OleDbType.VarWChar, 20, "VAT No"))
        '
        'OleDbSelectCommand1
        '
        Me.OleDbSelectCommand1.CommandText = "SELECT [ActiveY/N], AddressLine1, AddressLine2, AddressLine3, AddressLine4, Contr" & _
        "actorName, ContractorNo, PostalCode, [Reg No], Telephone, [VAT No] FROM Contract" & _
        "or ORDER BY ContractorNo"
        Me.OleDbSelectCommand1.Connection = Me.conContractor
        '
        'OleDbUpdateCommand1
        '
        Me.OleDbUpdateCommand1.CommandText = "UPDATE Contractor SET [ActiveY/N] = ?, AddressLine1 = ?, AddressLine2 = ?, Addres" & _
        "sLine3 = ?, AddressLine4 = ?, ContractorName = ?, ContractorNo = ?, PostalCode =" & _
        " ?, [Reg No] = ?, Telephone = ?, [VAT No] = ? WHERE (ContractorNo = ?) AND ([Act" & _
        "iveY/N] = ?) AND (AddressLine1 = ? OR ? IS NULL AND AddressLine1 IS NULL) AND (A" & _
        "ddressLine2 = ? OR ? IS NULL AND AddressLine2 IS NULL) AND (AddressLine3 = ? OR " & _
        "? IS NULL AND AddressLine3 IS NULL) AND (AddressLine4 = ? OR ? IS NULL AND Addre" & _
        "ssLine4 IS NULL) AND (ContractorName = ? OR ? IS NULL AND ContractorName IS NULL" & _
        ") AND (PostalCode = ? OR ? IS NULL AND PostalCode IS NULL) AND ([Reg No] = ? OR " & _
        "? IS NULL AND [Reg No] IS NULL) AND (Telephone = ? OR ? IS NULL AND Telephone IS" & _
        " NULL) AND ([VAT No] = ? OR ? IS NULL AND [VAT No] IS NULL)"
        Me.OleDbUpdateCommand1.Connection = Me.conContractor
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("ActiveY_N", System.Data.OleDb.OleDbType.Boolean, 2, "ActiveY/N"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("AddressLine1", System.Data.OleDb.OleDbType.VarWChar, 40, "AddressLine1"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("AddressLine2", System.Data.OleDb.OleDbType.VarWChar, 40, "AddressLine2"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("AddressLine3", System.Data.OleDb.OleDbType.VarWChar, 40, "AddressLine3"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("AddressLine4", System.Data.OleDb.OleDbType.VarWChar, 40, "AddressLine4"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("ContractorName", System.Data.OleDb.OleDbType.VarWChar, 70, "ContractorName"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("ContractorNo", System.Data.OleDb.OleDbType.VarWChar, 10, "ContractorNo"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("PostalCode", System.Data.OleDb.OleDbType.Integer, 0, "PostalCode"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Reg_No", System.Data.OleDb.OleDbType.VarWChar, 20, "Reg No"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Telephone", System.Data.OleDb.OleDbType.VarWChar, 15, "Telephone"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("VAT_No", System.Data.OleDb.OleDbType.VarWChar, 20, "VAT No"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ContractorNo", System.Data.OleDb.OleDbType.VarWChar, 10, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ContractorNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ActiveY_N", System.Data.OleDb.OleDbType.Boolean, 2, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ActiveY/N", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine1", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine1", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine11", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine1", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine2", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine2", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine21", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine2", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine3", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine3", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine31", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine3", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine4", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine4", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine41", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine4", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ContractorName", System.Data.OleDb.OleDbType.VarWChar, 70, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ContractorName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ContractorName1", System.Data.OleDb.OleDbType.VarWChar, 70, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ContractorName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_PostalCode", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "PostalCode", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_PostalCode1", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "PostalCode", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Reg_No", System.Data.OleDb.OleDbType.VarWChar, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Reg No", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Reg_No1", System.Data.OleDb.OleDbType.VarWChar, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Reg No", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Telephone", System.Data.OleDb.OleDbType.VarWChar, 15, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Telephone", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Telephone1", System.Data.OleDb.OleDbType.VarWChar, 15, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Telephone", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_VAT_No", System.Data.OleDb.OleDbType.VarWChar, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "VAT No", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_VAT_No1", System.Data.OleDb.OleDbType.VarWChar, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "VAT No", System.Data.DataRowVersion.Original, Nothing))
        '
        'dsContractor
        '
        Me.dsContractor.DataSetName = "dsReinforcingAbility"
        Me.dsContractor.Locale = New System.Globalization.CultureInfo("en-ZA")
        '
        'cbxCompNo
        '
        Me.cbxCompNo.DataSource = Me.dsContractor
        Me.cbxCompNo.DisplayMember = "Contractor.No&Name"
        Me.cbxCompNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxCompNo.Location = New System.Drawing.Point(144, 48)
        Me.cbxCompNo.Name = "cbxCompNo"
        Me.cbxCompNo.Size = New System.Drawing.Size(344, 21)
        Me.cbxCompNo.TabIndex = 5
        Me.cbxCompNo.ValueMember = "Contractor.ContractorNo"
        '
        'btnEdit
        '
        Me.btnEdit.Location = New System.Drawing.Point(104, 352)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "Edit"
        '
        'frmContractor
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(530, 392)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.cbxCompNo)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnAdd)
        Me.Controls.Add(Me.grpContractorDetails)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "frmContractor"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Contractor Maintenance"
        Me.grpContractorDetails.ResumeLayout(False)
        CType(Me.dsContractor, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private state As String

    Private CallingForm As Object

    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, ByVal keyData As Keys) As Boolean
        If keyData = Keys.Enter Then
            SendKeys.Send("{Tab}")
            Return True
        End If

        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function

    Public Sub New(ByVal caller As Object)
        MyBase.New()
        InitializeComponent()

        CallingForm = caller
    End Sub

    Private Sub frmContractor_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If Not IsNothing(CallingForm) Then
            CallingForm.Show()
        End If

        CallingForm = Nothing
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Close()
    End Sub

    Private Sub frmContractor_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        dsContractor.Clear()
        adpContractor.Fill(dsContractor.Contractor)

        grpContractorDetails.Enabled = False
        enablebinding()
    End Sub

    Private Sub disablebinding()
        txtContractorNo.DataBindings.Clear()
        txtContractorNo.Clear()

        txtContractorName.DataBindings.Clear()
        txtContractorName.Clear()

        txtVATNo.DataBindings.Clear()
        txtVATNo.Clear()

        txtAddress1.DataBindings.Clear()
        txtAddress1.Clear()

        txtAddress2.DataBindings.Clear()
        txtAddress2.Clear()

        txtAddress3.DataBindings.Clear()
        txtAddress3.Clear()

        txtAddress4.DataBindings.Clear()
        txtAddress4.Clear()

        txtPostalCode.DataBindings.Clear()
        txtPostalCode.Clear()

        txtTelNo.DataBindings.Clear()
        txtTelNo.Clear()

        'chbActive.DataBindings.Clear()
        'chbActive.CheckState = CheckState.Unchecked
    End Sub

    Private Sub enablebinding()
        txtContractorNo.DataBindings.Add("Text", dsContractor, "Contractor.ContractorNo")
        txtContractorName.DataBindings.Add("Text", dsContractor, "Contractor.ContractorName")
        txtVATNo.DataBindings.Add("Text", dsContractor, "Contractor.VAT No")
        txtAddress1.DataBindings.Add("Text", dsContractor, "Contractor.AddressLine1")
        txtAddress2.DataBindings.Add("Text", dsContractor, "Contractor.AddressLine2")
        txtAddress3.DataBindings.Add("Text", dsContractor, "Contractor.AddressLine3")
        txtAddress4.DataBindings.Add("Text", dsContractor, "Contractor.AddressLine4")
        txtPostalCode.DataBindings.Add("Text", dsContractor, "Contractor.PostalCode")
        txtTelNo.DataBindings.Add("Text", dsContractor, "Contractor.Telephone")
        'chbActive.DataBindings.Add("Checked", dsContractor, "Contractor.ActiveY/N")
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        If state = "" Then
            cbxCompNo.SendToBack()
            cbxCompNo.Enabled = False

            grpContractorDetails.Enabled = True
            disablebinding()

            state = "add"
            txtContractorNo.Focus()
        End If
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If state = "add" Then
            If txtContractorNo.Text = "" Then
                MsgBox("A Contractor Number is required", MsgBoxStyle.Critical, "Error")
                txtContractorNo.Focus()
            Else
                Dim DataReader As System.Data.OleDb.OleDbDataReader
                Dim count As Integer

                conContractor.Open()
                cmdCountContractorNo.Parameters("ContractorNo").Value = txtContractorNo.Text
                DataReader = cmdCountContractorNo.ExecuteReader(CommandBehavior.CloseConnection)
                While DataReader.Read()
                    count += 1
                End While
                DataReader.Close()
                conContractor.Close()

                If count > 0 Then
                    MsgBox("Contractor Number entered is already used", MsgBoxStyle.Critical, "Error")
                    txtContractorNo.Focus()
                Else
                    Dim row As DataRow = dsContractor.Contractor.NewContractorRow
                    row("ContractorNo") = txtContractorNo.Text

                    If txtContractorName.Text <> "" Then
                        row("ContractorName") = txtContractorName.Text
                    End If
                    If txtVATNo.Text <> "" Then
                        row("VAT No") = txtVATNo.Text
                    End If
                    If txtAddress1.Text <> "" Then
                        row("AddressLine1") = txtAddress1.Text
                    End If
                    If txtAddress2.Text <> "" Then
                        row("AddressLine2") = txtAddress2.Text
                    End If
                    If txtAddress3.Text <> "" Then
                        row("AddressLine3") = txtAddress3.Text
                    End If
                    If txtAddress4.Text <> "" Then
                        row("AddressLine4") = txtAddress4.Text
                    End If
                    If txtPostalCode.Text <> "" Then
                        row("PostalCode") = txtPostalCode.Text
                    End If
                    If txtTelNo.Text <> "" Then
                        row("Telephone") = txtTelNo.Text
                    End If
                    'row("ActiveY/N") = chbActive.Checked

                    dsContractor.Contractor.AddContractorRow(row)

                    adpContractor.Update(dsContractor.Contractor)
                    MsgBox("Record was successfully saved", MsgBoxStyle.Information, "Information")

                    enablebinding()

                    cbxCompNo.BringToFront()
                    cbxCompNo.Enabled = True

                    grpContractorDetails.Enabled = False
                    state = ""
                End If
            End If
        End If

        If state = "edit" Then
            dsContractor.Contractor.FindByContractorNo(txtContractorNo.Text).EndEdit()

            adpContractor.Update(dsContractor.Contractor)
            MsgBox("Record was successfully saved", MsgBoxStyle.Information, "Information")

            cbxCompNo.BringToFront()
            cbxCompNo.Enabled = True

            grpContractorDetails.Enabled = False
            txtContractorNo.Enabled = True

            state = ""
        End If
    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        If state = "" Then
            cbxCompNo.SendToBack()
            cbxCompNo.Enabled = False

            grpContractorDetails.Enabled = True
            txtContractorNo.Enabled = False

            state = "edit"
            txtContractorName.Focus()
        End If
    End Sub
End Class
