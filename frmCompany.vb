Public Class frmCompany
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
    Friend WithEvents lblCompNo As System.Windows.Forms.Label
    Friend WithEvents grpCompDetails As System.Windows.Forms.GroupBox
    Friend WithEvents txtCompNo As System.Windows.Forms.TextBox
    Friend WithEvents lblRegNo As System.Windows.Forms.Label
    Friend WithEvents lblAddress As System.Windows.Forms.Label
    Friend WithEvents txtAddress As System.Windows.Forms.TextBox
    Friend WithEvents lblPostalCode As System.Windows.Forms.Label
    Friend WithEvents lblCompName As System.Windows.Forms.Label
    Friend WithEvents lblVATNo As System.Windows.Forms.Label
    Friend WithEvents grpContactDetails As System.Windows.Forms.GroupBox
    Friend WithEvents lblTelNo As System.Windows.Forms.Label
    Friend WithEvents lblFaxNo As System.Windows.Forms.Label
    Friend WithEvents lblEmail As System.Windows.Forms.Label
    Friend WithEvents lblWebsite As System.Windows.Forms.Label
    Friend WithEvents grpMiscDetails As System.Windows.Forms.GroupBox
    Friend WithEvents lblMessage As System.Windows.Forms.Label
    Friend WithEvents lblLastInvNo As System.Windows.Forms.Label
    Friend WithEvents lblVAT As System.Windows.Forms.Label
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents conCompany As System.Data.OleDb.OleDbConnection
    Friend WithEvents adpCompany As System.Data.OleDb.OleDbDataAdapter
    Friend WithEvents OleDbSelectCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbInsertCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbUpdateCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbDeleteCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents dsCompany As Reinforcing_Ability.dsReinforcingAbility
    Friend WithEvents txtCompName As System.Windows.Forms.TextBox
    Friend WithEvents txtVATNo As System.Windows.Forms.TextBox
    Friend WithEvents txtPostalCode As System.Windows.Forms.TextBox
    Friend WithEvents txtWebsite As System.Windows.Forms.TextBox
    Friend WithEvents txtEmail As System.Windows.Forms.TextBox
    Friend WithEvents txtFaxNo As System.Windows.Forms.TextBox
    Friend WithEvents txtTelNo As System.Windows.Forms.TextBox
    Friend WithEvents txtMessage As System.Windows.Forms.TextBox
    Friend WithEvents txtVAT As System.Windows.Forms.TextBox
    Friend WithEvents txtLastInvNo As System.Windows.Forms.TextBox
    Friend WithEvents txtRegNo As System.Windows.Forms.TextBox
    Friend WithEvents cmdCountCompNo As System.Data.OleDb.OleDbCommand
    Friend WithEvents cbxCompNo As System.Windows.Forms.ComboBox
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents txtAddress2 As System.Windows.Forms.TextBox
    Friend WithEvents txtAddress3 As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.grpCompDetails = New System.Windows.Forms.GroupBox
        Me.txtVATNo = New System.Windows.Forms.TextBox
        Me.lblVATNo = New System.Windows.Forms.Label
        Me.txtCompName = New System.Windows.Forms.TextBox
        Me.lblCompName = New System.Windows.Forms.Label
        Me.txtPostalCode = New System.Windows.Forms.TextBox
        Me.lblPostalCode = New System.Windows.Forms.Label
        Me.txtAddress = New System.Windows.Forms.TextBox
        Me.lblAddress = New System.Windows.Forms.Label
        Me.txtRegNo = New System.Windows.Forms.TextBox
        Me.lblRegNo = New System.Windows.Forms.Label
        Me.txtCompNo = New System.Windows.Forms.TextBox
        Me.lblCompNo = New System.Windows.Forms.Label
        Me.dsCompany = New Reinforcing_Ability.dsReinforcingAbility
        Me.grpContactDetails = New System.Windows.Forms.GroupBox
        Me.txtWebsite = New System.Windows.Forms.TextBox
        Me.lblWebsite = New System.Windows.Forms.Label
        Me.txtEmail = New System.Windows.Forms.TextBox
        Me.lblEmail = New System.Windows.Forms.Label
        Me.txtFaxNo = New System.Windows.Forms.TextBox
        Me.lblFaxNo = New System.Windows.Forms.Label
        Me.txtTelNo = New System.Windows.Forms.TextBox
        Me.lblTelNo = New System.Windows.Forms.Label
        Me.grpMiscDetails = New System.Windows.Forms.GroupBox
        Me.txtVAT = New System.Windows.Forms.TextBox
        Me.lblVAT = New System.Windows.Forms.Label
        Me.txtLastInvNo = New System.Windows.Forms.TextBox
        Me.lblLastInvNo = New System.Windows.Forms.Label
        Me.txtMessage = New System.Windows.Forms.TextBox
        Me.lblMessage = New System.Windows.Forms.Label
        Me.btnAdd = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.conCompany = New System.Data.OleDb.OleDbConnection
        Me.adpCompany = New System.Data.OleDb.OleDbDataAdapter
        Me.OleDbDeleteCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbInsertCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbSelectCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbUpdateCommand1 = New System.Data.OleDb.OleDbCommand
        Me.cmdCountCompNo = New System.Data.OleDb.OleDbCommand
        Me.cbxCompNo = New System.Windows.Forms.ComboBox
        Me.btnEdit = New System.Windows.Forms.Button
        Me.txtAddress2 = New System.Windows.Forms.TextBox
        Me.txtAddress3 = New System.Windows.Forms.TextBox
        Me.grpCompDetails.SuspendLayout()
        CType(Me.dsCompany, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpContactDetails.SuspendLayout()
        Me.grpMiscDetails.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpCompDetails
        '
        Me.grpCompDetails.Controls.Add(Me.txtAddress3)
        Me.grpCompDetails.Controls.Add(Me.txtAddress2)
        Me.grpCompDetails.Controls.Add(Me.txtVATNo)
        Me.grpCompDetails.Controls.Add(Me.lblVATNo)
        Me.grpCompDetails.Controls.Add(Me.txtCompName)
        Me.grpCompDetails.Controls.Add(Me.lblCompName)
        Me.grpCompDetails.Controls.Add(Me.txtPostalCode)
        Me.grpCompDetails.Controls.Add(Me.lblPostalCode)
        Me.grpCompDetails.Controls.Add(Me.txtAddress)
        Me.grpCompDetails.Controls.Add(Me.lblAddress)
        Me.grpCompDetails.Controls.Add(Me.txtRegNo)
        Me.grpCompDetails.Controls.Add(Me.lblRegNo)
        Me.grpCompDetails.Controls.Add(Me.txtCompNo)
        Me.grpCompDetails.Controls.Add(Me.lblCompNo)
        Me.grpCompDetails.Enabled = False
        Me.grpCompDetails.Location = New System.Drawing.Point(16, 16)
        Me.grpCompDetails.Name = "grpCompDetails"
        Me.grpCompDetails.Size = New System.Drawing.Size(632, 216)
        Me.grpCompDetails.TabIndex = 0
        Me.grpCompDetails.TabStop = False
        Me.grpCompDetails.Text = "Company Details"
        '
        'txtVATNo
        '
        Me.txtVATNo.Location = New System.Drawing.Point(352, 64)
        Me.txtVATNo.MaxLength = 20
        Me.txtVATNo.Name = "txtVATNo"
        Me.txtVATNo.Size = New System.Drawing.Size(136, 20)
        Me.txtVATNo.TabIndex = 7
        Me.txtVATNo.Text = ""
        '
        'lblVATNo
        '
        Me.lblVATNo.Location = New System.Drawing.Point(296, 64)
        Me.lblVATNo.Name = "lblVATNo"
        Me.lblVATNo.Size = New System.Drawing.Size(48, 16)
        Me.lblVATNo.TabIndex = 6
        Me.lblVATNo.Text = "VAT No."
        Me.lblVATNo.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtCompName
        '
        Me.txtCompName.Location = New System.Drawing.Point(352, 32)
        Me.txtCompName.MaxLength = 40
        Me.txtCompName.Name = "txtCompName"
        Me.txtCompName.Size = New System.Drawing.Size(256, 20)
        Me.txtCompName.TabIndex = 3
        Me.txtCompName.Text = ""
        '
        'lblCompName
        '
        Me.lblCompName.Location = New System.Drawing.Point(248, 32)
        Me.lblCompName.Name = "lblCompName"
        Me.lblCompName.Size = New System.Drawing.Size(96, 16)
        Me.lblCompName.TabIndex = 2
        Me.lblCompName.Text = "Company Name"
        Me.lblCompName.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtPostalCode
        '
        Me.txtPostalCode.Location = New System.Drawing.Point(104, 184)
        Me.txtPostalCode.MaxLength = 5
        Me.txtPostalCode.Name = "txtPostalCode"
        Me.txtPostalCode.Size = New System.Drawing.Size(80, 20)
        Me.txtPostalCode.TabIndex = 11
        Me.txtPostalCode.Text = ""
        '
        'lblPostalCode
        '
        Me.lblPostalCode.Location = New System.Drawing.Point(24, 184)
        Me.lblPostalCode.Name = "lblPostalCode"
        Me.lblPostalCode.Size = New System.Drawing.Size(72, 16)
        Me.lblPostalCode.TabIndex = 10
        Me.lblPostalCode.Text = "Postal Code"
        Me.lblPostalCode.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtAddress
        '
        Me.txtAddress.AcceptsReturn = True
        Me.txtAddress.Location = New System.Drawing.Point(104, 96)
        Me.txtAddress.Multiline = True
        Me.txtAddress.Name = "txtAddress"
        Me.txtAddress.Size = New System.Drawing.Size(184, 24)
        Me.txtAddress.TabIndex = 9
        Me.txtAddress.Text = ""
        '
        'lblAddress
        '
        Me.lblAddress.Location = New System.Drawing.Point(48, 96)
        Me.lblAddress.Name = "lblAddress"
        Me.lblAddress.Size = New System.Drawing.Size(48, 16)
        Me.lblAddress.TabIndex = 8
        Me.lblAddress.Text = "Address"
        Me.lblAddress.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtRegNo
        '
        Me.txtRegNo.Location = New System.Drawing.Point(104, 64)
        Me.txtRegNo.MaxLength = 20
        Me.txtRegNo.Name = "txtRegNo"
        Me.txtRegNo.Size = New System.Drawing.Size(128, 20)
        Me.txtRegNo.TabIndex = 5
        Me.txtRegNo.Text = ""
        '
        'lblRegNo
        '
        Me.lblRegNo.Location = New System.Drawing.Point(40, 64)
        Me.lblRegNo.Name = "lblRegNo"
        Me.lblRegNo.Size = New System.Drawing.Size(56, 16)
        Me.lblRegNo.TabIndex = 4
        Me.lblRegNo.Text = "Reg. No."
        Me.lblRegNo.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtCompNo
        '
        Me.txtCompNo.Location = New System.Drawing.Point(104, 32)
        Me.txtCompNo.MaxLength = 10
        Me.txtCompNo.Name = "txtCompNo"
        Me.txtCompNo.Size = New System.Drawing.Size(80, 20)
        Me.txtCompNo.TabIndex = 1
        Me.txtCompNo.Text = ""
        '
        'lblCompNo
        '
        Me.lblCompNo.Location = New System.Drawing.Point(16, 32)
        Me.lblCompNo.Name = "lblCompNo"
        Me.lblCompNo.Size = New System.Drawing.Size(80, 16)
        Me.lblCompNo.TabIndex = 0
        Me.lblCompNo.Text = "Company No."
        Me.lblCompNo.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dsCompany
        '
        Me.dsCompany.DataSetName = "dsReinforcingAbility"
        Me.dsCompany.Locale = New System.Globalization.CultureInfo("en-ZA")
        '
        'grpContactDetails
        '
        Me.grpContactDetails.Controls.Add(Me.txtWebsite)
        Me.grpContactDetails.Controls.Add(Me.lblWebsite)
        Me.grpContactDetails.Controls.Add(Me.txtEmail)
        Me.grpContactDetails.Controls.Add(Me.lblEmail)
        Me.grpContactDetails.Controls.Add(Me.txtFaxNo)
        Me.grpContactDetails.Controls.Add(Me.lblFaxNo)
        Me.grpContactDetails.Controls.Add(Me.txtTelNo)
        Me.grpContactDetails.Controls.Add(Me.lblTelNo)
        Me.grpContactDetails.Location = New System.Drawing.Point(16, 240)
        Me.grpContactDetails.Name = "grpContactDetails"
        Me.grpContactDetails.Size = New System.Drawing.Size(632, 104)
        Me.grpContactDetails.TabIndex = 1
        Me.grpContactDetails.TabStop = False
        Me.grpContactDetails.Text = "Contact Details"
        '
        'txtWebsite
        '
        Me.txtWebsite.Location = New System.Drawing.Point(352, 64)
        Me.txtWebsite.MaxLength = 30
        Me.txtWebsite.Name = "txtWebsite"
        Me.txtWebsite.Size = New System.Drawing.Size(256, 20)
        Me.txtWebsite.TabIndex = 7
        Me.txtWebsite.Text = ""
        '
        'lblWebsite
        '
        Me.lblWebsite.Location = New System.Drawing.Point(296, 64)
        Me.lblWebsite.Name = "lblWebsite"
        Me.lblWebsite.Size = New System.Drawing.Size(48, 16)
        Me.lblWebsite.TabIndex = 6
        Me.lblWebsite.Text = "Website"
        Me.lblWebsite.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtEmail
        '
        Me.txtEmail.Location = New System.Drawing.Point(352, 32)
        Me.txtEmail.MaxLength = 40
        Me.txtEmail.Name = "txtEmail"
        Me.txtEmail.Size = New System.Drawing.Size(256, 20)
        Me.txtEmail.TabIndex = 3
        Me.txtEmail.Text = ""
        '
        'lblEmail
        '
        Me.lblEmail.Location = New System.Drawing.Point(296, 32)
        Me.lblEmail.Name = "lblEmail"
        Me.lblEmail.Size = New System.Drawing.Size(48, 16)
        Me.lblEmail.TabIndex = 2
        Me.lblEmail.Text = "Email"
        Me.lblEmail.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtFaxNo
        '
        Me.txtFaxNo.Location = New System.Drawing.Point(104, 64)
        Me.txtFaxNo.MaxLength = 15
        Me.txtFaxNo.Name = "txtFaxNo"
        Me.txtFaxNo.Size = New System.Drawing.Size(160, 20)
        Me.txtFaxNo.TabIndex = 5
        Me.txtFaxNo.Text = ""
        '
        'lblFaxNo
        '
        Me.lblFaxNo.Location = New System.Drawing.Point(48, 64)
        Me.lblFaxNo.Name = "lblFaxNo"
        Me.lblFaxNo.Size = New System.Drawing.Size(48, 16)
        Me.lblFaxNo.TabIndex = 4
        Me.lblFaxNo.Text = "Fax No."
        Me.lblFaxNo.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtTelNo
        '
        Me.txtTelNo.Location = New System.Drawing.Point(104, 32)
        Me.txtTelNo.MaxLength = 15
        Me.txtTelNo.Name = "txtTelNo"
        Me.txtTelNo.Size = New System.Drawing.Size(160, 20)
        Me.txtTelNo.TabIndex = 1
        Me.txtTelNo.Text = ""
        '
        'lblTelNo
        '
        Me.lblTelNo.Location = New System.Drawing.Point(48, 32)
        Me.lblTelNo.Name = "lblTelNo"
        Me.lblTelNo.Size = New System.Drawing.Size(48, 16)
        Me.lblTelNo.TabIndex = 0
        Me.lblTelNo.Text = "Tel. No."
        Me.lblTelNo.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'grpMiscDetails
        '
        Me.grpMiscDetails.Controls.Add(Me.txtVAT)
        Me.grpMiscDetails.Controls.Add(Me.lblVAT)
        Me.grpMiscDetails.Controls.Add(Me.txtLastInvNo)
        Me.grpMiscDetails.Controls.Add(Me.lblLastInvNo)
        Me.grpMiscDetails.Controls.Add(Me.txtMessage)
        Me.grpMiscDetails.Controls.Add(Me.lblMessage)
        Me.grpMiscDetails.Location = New System.Drawing.Point(16, 360)
        Me.grpMiscDetails.Name = "grpMiscDetails"
        Me.grpMiscDetails.Size = New System.Drawing.Size(632, 104)
        Me.grpMiscDetails.TabIndex = 2
        Me.grpMiscDetails.TabStop = False
        Me.grpMiscDetails.Text = "Misc. Details"
        '
        'txtVAT
        '
        Me.txtVAT.Location = New System.Drawing.Point(448, 24)
        Me.txtVAT.MaxLength = 3
        Me.txtVAT.Name = "txtVAT"
        Me.txtVAT.Size = New System.Drawing.Size(64, 20)
        Me.txtVAT.TabIndex = 3
        Me.txtVAT.Text = ""
        '
        'lblVAT
        '
        Me.lblVAT.Location = New System.Drawing.Point(400, 24)
        Me.lblVAT.Name = "lblVAT"
        Me.lblVAT.Size = New System.Drawing.Size(40, 16)
        Me.lblVAT.TabIndex = 2
        Me.lblVAT.Text = "VAT %"
        Me.lblVAT.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtLastInvNo
        '
        Me.txtLastInvNo.Location = New System.Drawing.Point(448, 64)
        Me.txtLastInvNo.MaxLength = 10
        Me.txtLastInvNo.Name = "txtLastInvNo"
        Me.txtLastInvNo.Size = New System.Drawing.Size(120, 20)
        Me.txtLastInvNo.TabIndex = 5
        Me.txtLastInvNo.Text = ""
        '
        'lblLastInvNo
        '
        Me.lblLastInvNo.Location = New System.Drawing.Point(352, 64)
        Me.lblLastInvNo.Name = "lblLastInvNo"
        Me.lblLastInvNo.Size = New System.Drawing.Size(88, 16)
        Me.lblLastInvNo.TabIndex = 4
        Me.lblLastInvNo.Text = "Last Invoice No."
        Me.lblLastInvNo.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtMessage
        '
        Me.txtMessage.AcceptsReturn = True
        Me.txtMessage.Location = New System.Drawing.Point(104, 24)
        Me.txtMessage.Multiline = True
        Me.txtMessage.Name = "txtMessage"
        Me.txtMessage.Size = New System.Drawing.Size(232, 64)
        Me.txtMessage.TabIndex = 1
        Me.txtMessage.Text = ""
        '
        'lblMessage
        '
        Me.lblMessage.Location = New System.Drawing.Point(40, 48)
        Me.lblMessage.Name = "lblMessage"
        Me.lblMessage.Size = New System.Drawing.Size(56, 16)
        Me.lblMessage.TabIndex = 0
        Me.lblMessage.Text = "Message"
        Me.lblMessage.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnAdd
        '
        Me.btnAdd.Location = New System.Drawing.Point(16, 480)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.TabIndex = 3
        Me.btnAdd.Text = "Add"
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(488, 480)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.TabIndex = 5
        Me.btnSave.Text = "Save"
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(576, 480)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.TabIndex = 6
        Me.btnClose.Text = "Close"
        '
        'conCompany
        '
        Me.conCompany.ConnectionString = "Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Registry Path=;Jet OLEDB:Database L" & _
        "ocking Mode=1;Data Source=""winsteelVers5.mdb"";Mode=Share Deny None;Jet OLEDB:Eng" & _
        "ine Type=5;Provider=""Microsoft.Jet.OLEDB.4.0"";Jet OLEDB:System database=;Jet OLE" & _
        "DB:SFP=False;persist security info=False;Extended Properties=;Jet OLEDB:Compact " & _
        "Without Replica Repair=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Create S" & _
        "ystem Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;User ID=Admin;" & _
        "Jet OLEDB:Global Bulk Transactions=1"
        '
        'adpCompany
        '
        Me.adpCompany.DeleteCommand = Me.OleDbDeleteCommand1
        Me.adpCompany.InsertCommand = Me.OleDbInsertCommand1
        Me.adpCompany.SelectCommand = Me.OleDbSelectCommand1
        Me.adpCompany.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "Company", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("Address", "Address"), New System.Data.Common.DataColumnMapping("AddressLine2", "AddressLine2"), New System.Data.Common.DataColumnMapping("AddressLine3", "AddressLine3"), New System.Data.Common.DataColumnMapping("AddressLine4", "AddressLine4"), New System.Data.Common.DataColumnMapping("CompanyName", "CompanyName"), New System.Data.Common.DataColumnMapping("CompanyNo", "CompanyNo"), New System.Data.Common.DataColumnMapping("Email", "Email"), New System.Data.Common.DataColumnMapping("Fax", "Fax"), New System.Data.Common.DataColumnMapping("LastCutNum", "LastCutNum"), New System.Data.Common.DataColumnMapping("LastInvNum", "LastInvNum"), New System.Data.Common.DataColumnMapping("Message", "Message"), New System.Data.Common.DataColumnMapping("PostalCode", "PostalCode"), New System.Data.Common.DataColumnMapping("RegNo", "RegNo"), New System.Data.Common.DataColumnMapping("Telephone", "Telephone"), New System.Data.Common.DataColumnMapping("UnitOfMeas", "UnitOfMeas"), New System.Data.Common.DataColumnMapping("VatNo", "VatNo"), New System.Data.Common.DataColumnMapping("VatPerc", "VatPerc"), New System.Data.Common.DataColumnMapping("Website", "Website")})})
        Me.adpCompany.UpdateCommand = Me.OleDbUpdateCommand1
        '
        'OleDbDeleteCommand1
        '
        Me.OleDbDeleteCommand1.CommandText = "DELETE FROM Company WHERE (CompanyNo = ?) AND (AddressLine2 = ? OR ? IS NULL AND " & _
        "AddressLine2 IS NULL) AND (AddressLine3 = ? OR ? IS NULL AND AddressLine3 IS NUL" & _
        "L) AND (AddressLine4 = ? OR ? IS NULL AND AddressLine4 IS NULL) AND (CompanyName" & _
        " = ? OR ? IS NULL AND CompanyName IS NULL) AND (Email = ? OR ? IS NULL AND Email" & _
        " IS NULL) AND (Fax = ? OR ? IS NULL AND Fax IS NULL) AND (LastCutNum = ? OR ? IS" & _
        " NULL AND LastCutNum IS NULL) AND (LastInvNum = ? OR ? IS NULL AND LastInvNum IS" & _
        " NULL) AND (Message = ? OR ? IS NULL AND Message IS NULL) AND (PostalCode = ? OR" & _
        " ? IS NULL AND PostalCode IS NULL) AND (RegNo = ? OR ? IS NULL AND RegNo IS NULL" & _
        ") AND (Telephone = ? OR ? IS NULL AND Telephone IS NULL) AND (UnitOfMeas = ? OR " & _
        "? IS NULL AND UnitOfMeas IS NULL) AND (VatNo = ? OR ? IS NULL AND VatNo IS NULL)" & _
        " AND (VatPerc = ? OR ? IS NULL AND VatPerc IS NULL) AND (Website = ? OR ? IS NUL" & _
        "L AND Website IS NULL)"
        Me.OleDbDeleteCommand1.Connection = Me.conCompany
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_CompanyNo", System.Data.OleDb.OleDbType.VarWChar, 10, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CompanyNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine2", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine2", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine21", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine2", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine3", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine3", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine31", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine3", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine4", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine4", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine41", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine4", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_CompanyName", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CompanyName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_CompanyName1", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CompanyName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Email", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Email", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Email1", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Email", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Fax", System.Data.OleDb.OleDbType.VarWChar, 15, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Fax", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Fax1", System.Data.OleDb.OleDbType.VarWChar, 15, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Fax", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_LastCutNum", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "LastCutNum", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_LastCutNum1", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "LastCutNum", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_LastInvNum", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "LastInvNum", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_LastInvNum1", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "LastInvNum", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Message", System.Data.OleDb.OleDbType.VarWChar, 200, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Message", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Message1", System.Data.OleDb.OleDbType.VarWChar, 200, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Message", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_PostalCode", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "PostalCode", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_PostalCode1", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "PostalCode", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_RegNo", System.Data.OleDb.OleDbType.VarWChar, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "RegNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_RegNo1", System.Data.OleDb.OleDbType.VarWChar, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "RegNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Telephone", System.Data.OleDb.OleDbType.VarWChar, 15, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Telephone", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Telephone1", System.Data.OleDb.OleDbType.VarWChar, 15, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Telephone", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_UnitOfMeas", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "UnitOfMeas", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_UnitOfMeas1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "UnitOfMeas", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_VatNo", System.Data.OleDb.OleDbType.VarWChar, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "VatNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_VatNo1", System.Data.OleDb.OleDbType.VarWChar, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "VatNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_VatPerc", System.Data.OleDb.OleDbType.Decimal, 0, System.Data.ParameterDirection.Input, False, CType(2, Byte), CType(2, Byte), "VatPerc", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_VatPerc1", System.Data.OleDb.OleDbType.Decimal, 0, System.Data.ParameterDirection.Input, False, CType(2, Byte), CType(2, Byte), "VatPerc", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Website", System.Data.OleDb.OleDbType.VarWChar, 30, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Website", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Website1", System.Data.OleDb.OleDbType.VarWChar, 30, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Website", System.Data.DataRowVersion.Original, Nothing))
        '
        'OleDbInsertCommand1
        '
        Me.OleDbInsertCommand1.CommandText = "INSERT INTO Company(Address, AddressLine2, AddressLine3, AddressLine4, CompanyNam" & _
        "e, CompanyNo, Email, Fax, LastCutNum, LastInvNum, Message, PostalCode, RegNo, Te" & _
        "lephone, UnitOfMeas, VatNo, VatPerc, Website) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?," & _
        " ?, ?, ?, ?, ?, ?, ?, ?, ?)"
        Me.OleDbInsertCommand1.Connection = Me.conCompany
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Address", System.Data.OleDb.OleDbType.VarWChar, 0, "Address"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("AddressLine2", System.Data.OleDb.OleDbType.VarWChar, 40, "AddressLine2"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("AddressLine3", System.Data.OleDb.OleDbType.VarWChar, 40, "AddressLine3"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("AddressLine4", System.Data.OleDb.OleDbType.VarWChar, 40, "AddressLine4"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("CompanyName", System.Data.OleDb.OleDbType.VarWChar, 40, "CompanyName"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("CompanyNo", System.Data.OleDb.OleDbType.VarWChar, 10, "CompanyNo"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Email", System.Data.OleDb.OleDbType.VarWChar, 40, "Email"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Fax", System.Data.OleDb.OleDbType.VarWChar, 15, "Fax"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("LastCutNum", System.Data.OleDb.OleDbType.Integer, 0, "LastCutNum"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("LastInvNum", System.Data.OleDb.OleDbType.Integer, 0, "LastInvNum"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Message", System.Data.OleDb.OleDbType.VarWChar, 200, "Message"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("PostalCode", System.Data.OleDb.OleDbType.Integer, 0, "PostalCode"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("RegNo", System.Data.OleDb.OleDbType.VarWChar, 20, "RegNo"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Telephone", System.Data.OleDb.OleDbType.VarWChar, 15, "Telephone"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("UnitOfMeas", System.Data.OleDb.OleDbType.VarWChar, 50, "UnitOfMeas"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("VatNo", System.Data.OleDb.OleDbType.VarWChar, 20, "VatNo"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("VatPerc", System.Data.OleDb.OleDbType.Decimal, 0, System.Data.ParameterDirection.Input, False, CType(2, Byte), CType(2, Byte), "VatPerc", System.Data.DataRowVersion.Current, Nothing))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Website", System.Data.OleDb.OleDbType.VarWChar, 30, "Website"))
        '
        'OleDbSelectCommand1
        '
        Me.OleDbSelectCommand1.CommandText = "SELECT Address, AddressLine2, AddressLine3, AddressLine4, CompanyName, CompanyNo," & _
        " Email, Fax, LastCutNum, LastInvNum, Message, PostalCode, RegNo, Telephone, Unit" & _
        "OfMeas, VatNo, VatPerc, Website FROM Company"
        Me.OleDbSelectCommand1.Connection = Me.conCompany
        '
        'OleDbUpdateCommand1
        '
        Me.OleDbUpdateCommand1.CommandText = "UPDATE Company SET Address = ?, AddressLine2 = ?, AddressLine3 = ?, AddressLine4 " & _
        "= ?, CompanyName = ?, CompanyNo = ?, Email = ?, Fax = ?, LastCutNum = ?, LastInv" & _
        "Num = ?, Message = ?, PostalCode = ?, RegNo = ?, Telephone = ?, UnitOfMeas = ?, " & _
        "VatNo = ?, VatPerc = ?, Website = ? WHERE (CompanyNo = ?) AND (AddressLine2 = ? " & _
        "OR ? IS NULL AND AddressLine2 IS NULL) AND (AddressLine3 = ? OR ? IS NULL AND Ad" & _
        "dressLine3 IS NULL) AND (AddressLine4 = ? OR ? IS NULL AND AddressLine4 IS NULL)" & _
        " AND (CompanyName = ? OR ? IS NULL AND CompanyName IS NULL) AND (Email = ? OR ? " & _
        "IS NULL AND Email IS NULL) AND (Fax = ? OR ? IS NULL AND Fax IS NULL) AND (LastC" & _
        "utNum = ? OR ? IS NULL AND LastCutNum IS NULL) AND (LastInvNum = ? OR ? IS NULL " & _
        "AND LastInvNum IS NULL) AND (Message = ? OR ? IS NULL AND Message IS NULL) AND (" & _
        "PostalCode = ? OR ? IS NULL AND PostalCode IS NULL) AND (RegNo = ? OR ? IS NULL " & _
        "AND RegNo IS NULL) AND (Telephone = ? OR ? IS NULL AND Telephone IS NULL) AND (U" & _
        "nitOfMeas = ? OR ? IS NULL AND UnitOfMeas IS NULL) AND (VatNo = ? OR ? IS NULL A" & _
        "ND VatNo IS NULL) AND (VatPerc = ? OR ? IS NULL AND VatPerc IS NULL) AND (Websit" & _
        "e = ? OR ? IS NULL AND Website IS NULL)"
        Me.OleDbUpdateCommand1.Connection = Me.conCompany
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Address", System.Data.OleDb.OleDbType.VarWChar, 0, "Address"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("AddressLine2", System.Data.OleDb.OleDbType.VarWChar, 40, "AddressLine2"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("AddressLine3", System.Data.OleDb.OleDbType.VarWChar, 40, "AddressLine3"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("AddressLine4", System.Data.OleDb.OleDbType.VarWChar, 40, "AddressLine4"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("CompanyName", System.Data.OleDb.OleDbType.VarWChar, 40, "CompanyName"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("CompanyNo", System.Data.OleDb.OleDbType.VarWChar, 10, "CompanyNo"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Email", System.Data.OleDb.OleDbType.VarWChar, 40, "Email"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Fax", System.Data.OleDb.OleDbType.VarWChar, 15, "Fax"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("LastCutNum", System.Data.OleDb.OleDbType.Integer, 0, "LastCutNum"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("LastInvNum", System.Data.OleDb.OleDbType.Integer, 0, "LastInvNum"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Message", System.Data.OleDb.OleDbType.VarWChar, 200, "Message"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("PostalCode", System.Data.OleDb.OleDbType.Integer, 0, "PostalCode"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("RegNo", System.Data.OleDb.OleDbType.VarWChar, 20, "RegNo"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Telephone", System.Data.OleDb.OleDbType.VarWChar, 15, "Telephone"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("UnitOfMeas", System.Data.OleDb.OleDbType.VarWChar, 50, "UnitOfMeas"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("VatNo", System.Data.OleDb.OleDbType.VarWChar, 20, "VatNo"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("VatPerc", System.Data.OleDb.OleDbType.Decimal, 0, System.Data.ParameterDirection.Input, False, CType(2, Byte), CType(2, Byte), "VatPerc", System.Data.DataRowVersion.Current, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Website", System.Data.OleDb.OleDbType.VarWChar, 30, "Website"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_CompanyNo", System.Data.OleDb.OleDbType.VarWChar, 10, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CompanyNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine2", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine2", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine21", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine2", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine3", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine3", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine31", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine3", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine4", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine4", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine41", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine4", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_CompanyName", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CompanyName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_CompanyName1", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CompanyName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Email", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Email", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Email1", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Email", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Fax", System.Data.OleDb.OleDbType.VarWChar, 15, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Fax", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Fax1", System.Data.OleDb.OleDbType.VarWChar, 15, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Fax", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_LastCutNum", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "LastCutNum", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_LastCutNum1", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "LastCutNum", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_LastInvNum", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "LastInvNum", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_LastInvNum1", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "LastInvNum", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Message", System.Data.OleDb.OleDbType.VarWChar, 200, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Message", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Message1", System.Data.OleDb.OleDbType.VarWChar, 200, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Message", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_PostalCode", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "PostalCode", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_PostalCode1", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "PostalCode", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_RegNo", System.Data.OleDb.OleDbType.VarWChar, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "RegNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_RegNo1", System.Data.OleDb.OleDbType.VarWChar, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "RegNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Telephone", System.Data.OleDb.OleDbType.VarWChar, 15, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Telephone", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Telephone1", System.Data.OleDb.OleDbType.VarWChar, 15, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Telephone", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_UnitOfMeas", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "UnitOfMeas", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_UnitOfMeas1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "UnitOfMeas", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_VatNo", System.Data.OleDb.OleDbType.VarWChar, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "VatNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_VatNo1", System.Data.OleDb.OleDbType.VarWChar, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "VatNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_VatPerc", System.Data.OleDb.OleDbType.Decimal, 0, System.Data.ParameterDirection.Input, False, CType(2, Byte), CType(2, Byte), "VatPerc", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_VatPerc1", System.Data.OleDb.OleDbType.Decimal, 0, System.Data.ParameterDirection.Input, False, CType(2, Byte), CType(2, Byte), "VatPerc", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Website", System.Data.OleDb.OleDbType.VarWChar, 30, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Website", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Website1", System.Data.OleDb.OleDbType.VarWChar, 30, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Website", System.Data.DataRowVersion.Original, Nothing))
        '
        'cmdCountCompNo
        '
        Me.cmdCountCompNo.CommandText = "SELECT Company.* FROM Company WHERE (CompanyNo = ?)"
        Me.cmdCountCompNo.Connection = Me.conCompany
        Me.cmdCountCompNo.Parameters.Add(New System.Data.OleDb.OleDbParameter("CompanyNo", System.Data.OleDb.OleDbType.VarWChar, 10, "CompanyNo"))
        '
        'cbxCompNo
        '
        Me.cbxCompNo.DataSource = Me.dsCompany
        Me.cbxCompNo.DisplayMember = "Company.No&Name"
        Me.cbxCompNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxCompNo.Location = New System.Drawing.Point(120, 48)
        Me.cbxCompNo.Name = "cbxCompNo"
        Me.cbxCompNo.Size = New System.Drawing.Size(504, 21)
        Me.cbxCompNo.TabIndex = 7
        Me.cbxCompNo.ValueMember = "Company.CompanyNo"
        '
        'btnEdit
        '
        Me.btnEdit.Location = New System.Drawing.Point(104, 480)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.TabIndex = 4
        Me.btnEdit.Text = "Edit"
        '
        'txtAddress2
        '
        Me.txtAddress2.Location = New System.Drawing.Point(104, 128)
        Me.txtAddress2.Name = "txtAddress2"
        Me.txtAddress2.Size = New System.Drawing.Size(184, 20)
        Me.txtAddress2.TabIndex = 12
        Me.txtAddress2.Text = ""
        '
        'txtAddress3
        '
        Me.txtAddress3.Location = New System.Drawing.Point(104, 160)
        Me.txtAddress3.Name = "txtAddress3"
        Me.txtAddress3.Size = New System.Drawing.Size(184, 20)
        Me.txtAddress3.TabIndex = 13
        Me.txtAddress3.Text = ""
        '
        'frmCompany
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(664, 518)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.cbxCompNo)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnAdd)
        Me.Controls.Add(Me.grpMiscDetails)
        Me.Controls.Add(Me.grpContactDetails)
        Me.Controls.Add(Me.grpCompDetails)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "frmCompany"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Company Maintenance"
        Me.grpCompDetails.ResumeLayout(False)
        CType(Me.dsCompany, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpContactDetails.ResumeLayout(False)
        Me.grpMiscDetails.ResumeLayout(False)
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

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Close()
    End Sub

    Private Sub frmCompany_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        dsCompany.Clear()
        adpCompany.Fill(dsCompany.Company)

        enablebinding()
        disablefields()
    End Sub

    Private Sub disablefields()
        grpCompDetails.Enabled = False
        grpContactDetails.Enabled = False
        grpMiscDetails.Enabled = False
    End Sub

    Private Sub enablefields()
        grpCompDetails.Enabled = True
        grpContactDetails.Enabled = True
        grpMiscDetails.Enabled = True
    End Sub

    Private Sub disablebinding()
        txtCompNo.DataBindings.Clear()
        txtCompNo.Clear()

        txtCompName.DataBindings.Clear()
        txtCompName.Clear()

        txtRegNo.DataBindings.Clear()
        txtRegNo.Clear()

        txtVATNo.DataBindings.Clear()
        txtVATNo.Clear()

        txtAddress.DataBindings.Clear()
        txtAddress.Clear()

        txtAddress2.DataBindings.Clear()
        txtAddress2.Clear()

        txtAddress3.DataBindings.Clear()
        txtAddress3.Clear()

        txtPostalCode.DataBindings.Clear()
        txtPostalCode.Clear()

        txtTelNo.DataBindings.Clear()
        txtTelNo.Clear()

        txtEmail.DataBindings.Clear()
        txtEmail.Clear()

        txtFaxNo.DataBindings.Clear()
        txtFaxNo.Clear()

        txtWebsite.DataBindings.Clear()
        txtWebsite.Clear()

        txtMessage.DataBindings.Clear()
        txtMessage.Clear()

        txtVAT.DataBindings.Clear()
        txtVAT.Clear()

        txtLastInvNo.DataBindings.Clear()
        txtLastInvNo.Clear()
    End Sub

    Private Sub enablebinding()
        txtCompNo.DataBindings.Add("Text", dsCompany, "Company.CompanyNo")
        txtCompName.DataBindings.Add("Text", dsCompany, "Company.CompanyName")
        txtRegNo.DataBindings.Add("Text", dsCompany, "Company.RegNo")
        txtVATNo.DataBindings.Add("Text", dsCompany, "Company.VatNo")
        txtAddress.DataBindings.Add("Text", dsCompany, "Company.Address")
        txtAddress2.DataBindings.Add("Text", dsCompany, "Company.AddressLine2")
        txtAddress3.DataBindings.Add("Text", dsCompany, "Company.AddressLine3")
        txtPostalCode.DataBindings.Add("Text", dsCompany, "Company.PostalCode")
        txtTelNo.DataBindings.Add("Text", dsCompany, "Company.Telephone")
        txtEmail.DataBindings.Add("Text", dsCompany, "Company.Email")
        txtFaxNo.DataBindings.Add("Text", dsCompany, "Company.Fax")
        txtWebsite.DataBindings.Add("Text", dsCompany, "Company.Website")
        txtMessage.DataBindings.Add("Text", dsCompany, "Company.Message")
        txtVAT.DataBindings.Add("Text", dsCompany, "Company.VatPerc")
        txtLastInvNo.DataBindings.Add("Text", dsCompany, "Company.LastInvNum")
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        If state = "" Then
            cbxCompNo.SendToBack()
            cbxCompNo.Enabled = False

            enablefields()
            disablebinding()

            txtVAT.Text = 0.14
            state = "add"
            txtCompNo.Focus()
        End If
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If state = "add" Then
            If txtCompNo.Text = "" Then
                MsgBox("A Company Number is required", MsgBoxStyle.Critical, "Error")
                txtCompNo.Focus()
            Else
                Dim DataReader As System.Data.OleDb.OleDbDataReader
                Dim count As Integer

                conCompany.Open()
                cmdCountCompNo.Parameters("CompanyNo").Value = txtCompNo.Text
                DataReader = cmdCountCompNo.ExecuteReader(CommandBehavior.CloseConnection)
                While DataReader.Read()
                    count += 1
                End While
                DataReader.Close()
                conCompany.Close()

                If count > 0 Then
                    MsgBox("Company Number entered is already used", MsgBoxStyle.Critical, "Error")
                    txtCompNo.Focus()
                Else
                    Dim row As DataRow = dsCompany.Company.NewCompanyRow
                    row("CompanyNo") = txtCompNo.Text

                    If txtCompName.Text <> "" Then
                        row("CompanyName") = txtCompName.Text
                    End If
                    If txtRegNo.Text <> "" Then
                        row("RegNo") = txtRegNo.Text
                    End If
                    If txtVATNo.Text <> "" Then
                        row("VatNo") = txtVATNo.Text
                    End If
                    If txtAddress.Text <> "" Then
                        row("Address") = txtAddress.Text
                    End If
                    If txtAddress2.Text <> "" Then
                        row("AddressLine2") = txtAddress2.Text
                    End If
                    If txtAddress3.Text <> "" Then
                        row("AddressLine3") = txtAddress3.Text
                    End If
                    If txtPostalCode.Text <> "" Then
                        row("PostalCode") = txtPostalCode.Text
                    End If
                    If txtTelNo.Text <> "" Then
                        row("Telephone") = txtTelNo.Text
                    End If
                    If txtEmail.Text <> "" Then
                        row("Email") = txtEmail.Text
                    End If
                    If txtFaxNo.Text <> "" Then
                        row("Fax") = txtFaxNo.Text
                    End If
                    If txtWebsite.Text <> "" Then
                        row("Website") = txtWebsite.Text
                    End If
                    If txtMessage.Text <> "" Then
                        row("Message") = txtMessage.Text
                    End If
                    If txtVAT.Text <> "" Then
                        row("VatPerc") = txtVAT.Text
                    End If
                    If txtLastInvNo.Text <> "" Then
                        row("LastInvNum") = txtLastInvNo.Text
                    End If

                    dsCompany.Company.AddCompanyRow(row)

                    adpCompany.Update(dsCompany.Company)
                    MsgBox("Record was successfully saved", MsgBoxStyle.Information, "Information")

                    enablebinding()
                    cbxCompNo.BringToFront()
                    disablefields()
                    state = ""
                End If
            End If
        End If

        If state = "edit" Then
            dsCompany.Company.FindByCompanyNo(txtCompNo.Text).EndEdit()

            adpCompany.Update(dsCompany.Company)
            MsgBox("Record was successfully saved", MsgBoxStyle.Information, "Information")

            cbxCompNo.BringToFront()
            cbxCompNo.Enabled = True

            disablefields()
            txtCompNo.Enabled = True

            state = ""
        End If
    End Sub

    Private Sub frmCompany_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If Not IsNothing(CallingForm) Then
            CallingForm.Show()
        End If

        CallingForm = Nothing
    End Sub

    Public Sub New(ByVal caller As Object)
        MyBase.New()
        InitializeComponent()

        CallingForm = caller
    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        If state = "" Then
            cbxCompNo.SendToBack()
            txtCompNo.Enabled = False

            enablefields()

            state = "edit"
            txtCompName.Focus()
        End If
    End Sub
End Class
