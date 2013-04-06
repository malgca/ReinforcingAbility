Public Class frmJob
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
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents lblContractorName As System.Windows.Forms.Label
    Friend WithEvents conJob As System.Data.OleDb.OleDbConnection
    Friend WithEvents adpJob As System.Data.OleDb.OleDbDataAdapter
    Friend WithEvents OleDbSelectCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbInsertCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbUpdateCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbDeleteCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents dsJob As PresentationTier.dsReinforcingAbility
    Friend WithEvents grpJobDetails As System.Windows.Forms.GroupBox
    Friend WithEvents lblJobNo As System.Windows.Forms.Label
    Friend WithEvents txtJobNo As System.Windows.Forms.TextBox
    Friend WithEvents cbxCompanyNo As System.Windows.Forms.ComboBox
    Friend WithEvents txtJobName As System.Windows.Forms.TextBox
    Friend WithEvents lblJobName As System.Windows.Forms.Label
    Friend WithEvents lblDiscount As System.Windows.Forms.Label
    Friend WithEvents txtDiscount As System.Windows.Forms.TextBox
    Friend WithEvents lblAddDiscount As System.Windows.Forms.Label
    Friend WithEvents txtAddDiscount As System.Windows.Forms.TextBox
    Friend WithEvents lblDesignCost As System.Windows.Forms.Label
    Friend WithEvents txtDesignCost As System.Windows.Forms.TextBox
    Friend WithEvents lblTonsKilograms As System.Windows.Forms.Label
    Friend WithEvents cbxTonsKilograms As System.Windows.Forms.ComboBox
    Friend WithEvents lblContractor As System.Windows.Forms.Label
    Friend WithEvents lblOrderNo As System.Windows.Forms.Label
    Friend WithEvents adpCompany As System.Data.OleDb.OleDbDataAdapter
    Friend WithEvents OleDbSelectCommand2 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbInsertCommand2 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbUpdateCommand2 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbDeleteCommand2 As System.Data.OleDb.OleDbCommand
    Friend WithEvents adpContractor As System.Data.OleDb.OleDbDataAdapter
    Friend WithEvents OleDbSelectCommand3 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbInsertCommand3 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbUpdateCommand3 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbDeleteCommand3 As System.Data.OleDb.OleDbCommand
    Friend WithEvents txtOrderNo As System.Windows.Forms.TextBox
    Friend WithEvents cbxContractor As System.Windows.Forms.ComboBox
    Friend WithEvents cmdCountJobNo As System.Data.OleDb.OleDbCommand
    Friend WithEvents cbxJobNo As System.Windows.Forms.ComboBox
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents adpJobRate As System.Data.OleDb.OleDbDataAdapter
    Friend WithEvents OleDbSelectCommand4 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbInsertCommand4 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbUpdateCommand4 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbDeleteCommand4 As System.Data.OleDb.OleDbCommand
    Friend WithEvents adpProductType As System.Data.OleDb.OleDbDataAdapter
    Friend WithEvents OleDbSelectCommand5 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbInsertCommand5 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbUpdateCommand5 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbDeleteCommand5 As System.Data.OleDb.OleDbCommand
    Friend WithEvents btnArchive As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnAdd = New System.Windows.Forms.Button
        Me.grpJobDetails = New System.Windows.Forms.GroupBox
        Me.lblOrderNo = New System.Windows.Forms.Label
        Me.cbxContractor = New System.Windows.Forms.ComboBox
        Me.dsJob = New PresentationTier.dsReinforcingAbility
        Me.cbxTonsKilograms = New System.Windows.Forms.ComboBox
        Me.cbxCompanyNo = New System.Windows.Forms.ComboBox
        Me.txtOrderNo = New System.Windows.Forms.TextBox
        Me.lblContractor = New System.Windows.Forms.Label
        Me.lblTonsKilograms = New System.Windows.Forms.Label
        Me.txtDesignCost = New System.Windows.Forms.TextBox
        Me.lblDesignCost = New System.Windows.Forms.Label
        Me.txtAddDiscount = New System.Windows.Forms.TextBox
        Me.lblAddDiscount = New System.Windows.Forms.Label
        Me.txtDiscount = New System.Windows.Forms.TextBox
        Me.lblDiscount = New System.Windows.Forms.Label
        Me.txtJobName = New System.Windows.Forms.TextBox
        Me.lblJobName = New System.Windows.Forms.Label
        Me.lblContractorName = New System.Windows.Forms.Label
        Me.txtJobNo = New System.Windows.Forms.TextBox
        Me.lblJobNo = New System.Windows.Forms.Label
        Me.conJob = New System.Data.OleDb.OleDbConnection
        Me.adpJob = New System.Data.OleDb.OleDbDataAdapter
        Me.OleDbDeleteCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbInsertCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbSelectCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbUpdateCommand1 = New System.Data.OleDb.OleDbCommand
        Me.adpCompany = New System.Data.OleDb.OleDbDataAdapter
        Me.OleDbDeleteCommand2 = New System.Data.OleDb.OleDbCommand
        Me.OleDbInsertCommand2 = New System.Data.OleDb.OleDbCommand
        Me.OleDbSelectCommand2 = New System.Data.OleDb.OleDbCommand
        Me.OleDbUpdateCommand2 = New System.Data.OleDb.OleDbCommand
        Me.adpContractor = New System.Data.OleDb.OleDbDataAdapter
        Me.OleDbDeleteCommand3 = New System.Data.OleDb.OleDbCommand
        Me.OleDbInsertCommand3 = New System.Data.OleDb.OleDbCommand
        Me.OleDbSelectCommand3 = New System.Data.OleDb.OleDbCommand
        Me.OleDbUpdateCommand3 = New System.Data.OleDb.OleDbCommand
        Me.cmdCountJobNo = New System.Data.OleDb.OleDbCommand
        Me.cbxJobNo = New System.Windows.Forms.ComboBox
        Me.btnEdit = New System.Windows.Forms.Button
        Me.adpJobRate = New System.Data.OleDb.OleDbDataAdapter
        Me.OleDbDeleteCommand4 = New System.Data.OleDb.OleDbCommand
        Me.OleDbInsertCommand4 = New System.Data.OleDb.OleDbCommand
        Me.OleDbSelectCommand4 = New System.Data.OleDb.OleDbCommand
        Me.OleDbUpdateCommand4 = New System.Data.OleDb.OleDbCommand
        Me.adpProductType = New System.Data.OleDb.OleDbDataAdapter
        Me.OleDbDeleteCommand5 = New System.Data.OleDb.OleDbCommand
        Me.OleDbInsertCommand5 = New System.Data.OleDb.OleDbCommand
        Me.OleDbSelectCommand5 = New System.Data.OleDb.OleDbCommand
        Me.OleDbUpdateCommand5 = New System.Data.OleDb.OleDbCommand
        Me.btnArchive = New System.Windows.Forms.Button
        Me.grpJobDetails.SuspendLayout()
        CType(Me.dsJob, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(440, 296)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.TabIndex = 4
        Me.btnClose.Text = "Close"
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(352, 296)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.TabIndex = 3
        Me.btnSave.Text = "Save"
        '
        'btnAdd
        '
        Me.btnAdd.Location = New System.Drawing.Point(16, 296)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.TabIndex = 1
        Me.btnAdd.Text = "Add"
        '
        'grpJobDetails
        '
        Me.grpJobDetails.Controls.Add(Me.lblOrderNo)
        Me.grpJobDetails.Controls.Add(Me.cbxContractor)
        Me.grpJobDetails.Controls.Add(Me.cbxTonsKilograms)
        Me.grpJobDetails.Controls.Add(Me.cbxCompanyNo)
        Me.grpJobDetails.Controls.Add(Me.txtOrderNo)
        Me.grpJobDetails.Controls.Add(Me.lblContractor)
        Me.grpJobDetails.Controls.Add(Me.lblTonsKilograms)
        Me.grpJobDetails.Controls.Add(Me.txtDesignCost)
        Me.grpJobDetails.Controls.Add(Me.lblDesignCost)
        Me.grpJobDetails.Controls.Add(Me.txtAddDiscount)
        Me.grpJobDetails.Controls.Add(Me.lblAddDiscount)
        Me.grpJobDetails.Controls.Add(Me.txtDiscount)
        Me.grpJobDetails.Controls.Add(Me.lblDiscount)
        Me.grpJobDetails.Controls.Add(Me.txtJobName)
        Me.grpJobDetails.Controls.Add(Me.lblJobName)
        Me.grpJobDetails.Controls.Add(Me.lblContractorName)
        Me.grpJobDetails.Controls.Add(Me.txtJobNo)
        Me.grpJobDetails.Controls.Add(Me.lblJobNo)
        Me.grpJobDetails.Location = New System.Drawing.Point(16, 16)
        Me.grpJobDetails.Name = "grpJobDetails"
        Me.grpJobDetails.Size = New System.Drawing.Size(496, 264)
        Me.grpJobDetails.TabIndex = 0
        Me.grpJobDetails.TabStop = False
        Me.grpJobDetails.Text = "Job Details"
        '
        'lblOrderNo
        '
        Me.lblOrderNo.Location = New System.Drawing.Point(24, 224)
        Me.lblOrderNo.Name = "lblOrderNo"
        Me.lblOrderNo.Size = New System.Drawing.Size(96, 16)
        Me.lblOrderNo.TabIndex = 16
        Me.lblOrderNo.Text = "Order No."
        Me.lblOrderNo.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbxContractor
        '
        Me.cbxContractor.DataBindings.Add(New System.Windows.Forms.Binding("SelectedValue", Me.dsJob, "Job.ContractorNo"))
        Me.cbxContractor.DataSource = Me.dsJob
        Me.cbxContractor.DisplayMember = "Contractor.ContractorName"
        Me.cbxContractor.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxContractor.Location = New System.Drawing.Point(128, 192)
        Me.cbxContractor.Name = "cbxContractor"
        Me.cbxContractor.Size = New System.Drawing.Size(344, 21)
        Me.cbxContractor.TabIndex = 15
        Me.cbxContractor.ValueMember = "Contractor.ContractorNo"
        '
        'dsJob
        '
        Me.dsJob.DataSetName = "dsReinforcingAbility"
        Me.dsJob.Locale = New System.Globalization.CultureInfo("en-ZA")
        '
        'cbxTonsKilograms
        '
        Me.cbxTonsKilograms.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.dsJob, "Job.Tons Or Kilograms"))
        Me.cbxTonsKilograms.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxTonsKilograms.Location = New System.Drawing.Point(128, 160)
        Me.cbxTonsKilograms.Name = "cbxTonsKilograms"
        Me.cbxTonsKilograms.Size = New System.Drawing.Size(104, 21)
        Me.cbxTonsKilograms.TabIndex = 11
        '
        'cbxCompanyNo
        '
        Me.cbxCompanyNo.DataBindings.Add(New System.Windows.Forms.Binding("SelectedValue", Me.dsJob, "Job.CompanyNo"))
        Me.cbxCompanyNo.DataSource = Me.dsJob
        Me.cbxCompanyNo.DisplayMember = "Company.CompanyName"
        Me.cbxCompanyNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxCompanyNo.Location = New System.Drawing.Point(128, 64)
        Me.cbxCompanyNo.Name = "cbxCompanyNo"
        Me.cbxCompanyNo.Size = New System.Drawing.Size(344, 21)
        Me.cbxCompanyNo.TabIndex = 3
        Me.cbxCompanyNo.ValueMember = "Company.CompanyNo"
        '
        'txtOrderNo
        '
        Me.txtOrderNo.Location = New System.Drawing.Point(128, 224)
        Me.txtOrderNo.MaxLength = 25
        Me.txtOrderNo.Name = "txtOrderNo"
        Me.txtOrderNo.Size = New System.Drawing.Size(104, 20)
        Me.txtOrderNo.TabIndex = 17
        Me.txtOrderNo.Text = ""
        '
        'lblContractor
        '
        Me.lblContractor.Location = New System.Drawing.Point(24, 192)
        Me.lblContractor.Name = "lblContractor"
        Me.lblContractor.Size = New System.Drawing.Size(96, 16)
        Me.lblContractor.TabIndex = 14
        Me.lblContractor.Text = "Contractor"
        Me.lblContractor.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblTonsKilograms
        '
        Me.lblTonsKilograms.Location = New System.Drawing.Point(24, 160)
        Me.lblTonsKilograms.Name = "lblTonsKilograms"
        Me.lblTonsKilograms.Size = New System.Drawing.Size(96, 16)
        Me.lblTonsKilograms.TabIndex = 10
        Me.lblTonsKilograms.Text = "Tons / Kilograms"
        Me.lblTonsKilograms.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtDesignCost
        '
        Me.txtDesignCost.Location = New System.Drawing.Point(344, 160)
        Me.txtDesignCost.MaxLength = 10
        Me.txtDesignCost.Name = "txtDesignCost"
        Me.txtDesignCost.Size = New System.Drawing.Size(128, 20)
        Me.txtDesignCost.TabIndex = 13
        Me.txtDesignCost.Text = ""
        '
        'lblDesignCost
        '
        Me.lblDesignCost.Location = New System.Drawing.Point(264, 160)
        Me.lblDesignCost.Name = "lblDesignCost"
        Me.lblDesignCost.Size = New System.Drawing.Size(72, 16)
        Me.lblDesignCost.TabIndex = 12
        Me.lblDesignCost.Text = "Design Cost"
        Me.lblDesignCost.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtAddDiscount
        '
        Me.txtAddDiscount.Location = New System.Drawing.Point(344, 128)
        Me.txtAddDiscount.MaxLength = 10
        Me.txtAddDiscount.Name = "txtAddDiscount"
        Me.txtAddDiscount.Size = New System.Drawing.Size(128, 20)
        Me.txtAddDiscount.TabIndex = 9
        Me.txtAddDiscount.Text = ""
        '
        'lblAddDiscount
        '
        Me.lblAddDiscount.Location = New System.Drawing.Point(264, 128)
        Me.lblAddDiscount.Name = "lblAddDiscount"
        Me.lblAddDiscount.Size = New System.Drawing.Size(72, 16)
        Me.lblAddDiscount.TabIndex = 8
        Me.lblAddDiscount.Text = "Add Discount"
        Me.lblAddDiscount.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtDiscount
        '
        Me.txtDiscount.Location = New System.Drawing.Point(128, 128)
        Me.txtDiscount.MaxLength = 5
        Me.txtDiscount.Name = "txtDiscount"
        Me.txtDiscount.Size = New System.Drawing.Size(104, 20)
        Me.txtDiscount.TabIndex = 7
        Me.txtDiscount.Text = ""
        '
        'lblDiscount
        '
        Me.lblDiscount.Location = New System.Drawing.Point(24, 128)
        Me.lblDiscount.Name = "lblDiscount"
        Me.lblDiscount.Size = New System.Drawing.Size(96, 16)
        Me.lblDiscount.TabIndex = 6
        Me.lblDiscount.Text = "Discount"
        Me.lblDiscount.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtJobName
        '
        Me.txtJobName.Location = New System.Drawing.Point(128, 96)
        Me.txtJobName.MaxLength = 70
        Me.txtJobName.Name = "txtJobName"
        Me.txtJobName.Size = New System.Drawing.Size(344, 20)
        Me.txtJobName.TabIndex = 5
        Me.txtJobName.Text = ""
        '
        'lblJobName
        '
        Me.lblJobName.Location = New System.Drawing.Point(24, 96)
        Me.lblJobName.Name = "lblJobName"
        Me.lblJobName.Size = New System.Drawing.Size(96, 16)
        Me.lblJobName.TabIndex = 4
        Me.lblJobName.Text = "Job Name"
        Me.lblJobName.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblContractorName
        '
        Me.lblContractorName.Location = New System.Drawing.Point(24, 64)
        Me.lblContractorName.Name = "lblContractorName"
        Me.lblContractorName.Size = New System.Drawing.Size(96, 16)
        Me.lblContractorName.TabIndex = 2
        Me.lblContractorName.Text = "Company Name"
        Me.lblContractorName.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtJobNo
        '
        Me.txtJobNo.Location = New System.Drawing.Point(128, 32)
        Me.txtJobNo.MaxLength = 50
        Me.txtJobNo.Name = "txtJobNo"
        Me.txtJobNo.Size = New System.Drawing.Size(104, 20)
        Me.txtJobNo.TabIndex = 1
        Me.txtJobNo.Text = ""
        '
        'lblJobNo
        '
        Me.lblJobNo.Location = New System.Drawing.Point(24, 32)
        Me.lblJobNo.Name = "lblJobNo"
        Me.lblJobNo.Size = New System.Drawing.Size(96, 16)
        Me.lblJobNo.TabIndex = 0
        Me.lblJobNo.Text = "Job No."
        Me.lblJobNo.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'conJob
        '
        Me.conJob.ConnectionString = "Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Registry Path=;Jet OLEDB:Database L" & _
        "ocking Mode=1;Data Source=""winsteelVers5.mdb"";Mode=Share Deny None;Jet OLEDB:Eng" & _
        "ine Type=5;Provider=""Microsoft.Jet.OLEDB.4.0"";Jet OLEDB:System database=;Jet OLE" & _
        "DB:SFP=False;persist security info=False;Extended Properties=;Jet OLEDB:Compact " & _
        "Without Replica Repair=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Create S" & _
        "ystem Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;User ID=Admin;" & _
        "Jet OLEDB:Global Bulk Transactions=1"
        '
        'adpJob
        '
        Me.adpJob.DeleteCommand = Me.OleDbDeleteCommand1
        Me.adpJob.InsertCommand = Me.OleDbInsertCommand1
        Me.adpJob.SelectCommand = Me.OleDbSelectCommand1
        Me.adpJob.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "Job", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("AddedDiscount", "AddedDiscount"), New System.Data.Common.DataColumnMapping("CompanyNo", "CompanyNo"), New System.Data.Common.DataColumnMapping("ContractorNo", "ContractorNo"), New System.Data.Common.DataColumnMapping("Design", "Design"), New System.Data.Common.DataColumnMapping("Discount", "Discount"), New System.Data.Common.DataColumnMapping("JobName", "JobName"), New System.Data.Common.DataColumnMapping("JobNo", "JobNo"), New System.Data.Common.DataColumnMapping("OrderNo", "OrderNo"), New System.Data.Common.DataColumnMapping("Tons Or Kilograms", "Tons Or Kilograms"), New System.Data.Common.DataColumnMapping("UnitOfMeas", "UnitOfMeas")})})
        Me.adpJob.UpdateCommand = Me.OleDbUpdateCommand1
        '
        'OleDbDeleteCommand1
        '
        Me.OleDbDeleteCommand1.CommandText = "DELETE FROM Job WHERE (JobNo = ?) AND (AddedDiscount = ? OR ? IS NULL AND AddedDi" & _
        "scount IS NULL) AND (CompanyNo = ? OR ? IS NULL AND CompanyNo IS NULL) AND (Cont" & _
        "ractorNo = ? OR ? IS NULL AND ContractorNo IS NULL) AND (Design = ? OR ? IS NULL" & _
        " AND Design IS NULL) AND (Discount = ? OR ? IS NULL AND Discount IS NULL) AND (J" & _
        "obName = ? OR ? IS NULL AND JobName IS NULL) AND (OrderNo = ? OR ? IS NULL AND O" & _
        "rderNo IS NULL) AND ([Tons Or Kilograms] = ? OR ? IS NULL AND [Tons Or Kilograms" & _
        "] IS NULL) AND (UnitOfMeas = ? OR ? IS NULL AND UnitOfMeas IS NULL)"
        Me.OleDbDeleteCommand1.Connection = Me.conJob
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_JobNo", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "JobNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddedDiscount", System.Data.OleDb.OleDbType.Currency, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddedDiscount", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddedDiscount1", System.Data.OleDb.OleDbType.Currency, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddedDiscount", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_CompanyNo", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CompanyNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_CompanyNo1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CompanyNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ContractorNo", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ContractorNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ContractorNo1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ContractorNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Design", System.Data.OleDb.OleDbType.Currency, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Design", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Design1", System.Data.OleDb.OleDbType.Currency, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Design", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Discount", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Discount", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Discount1", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Discount", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_JobName", System.Data.OleDb.OleDbType.VarWChar, 70, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "JobName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_JobName1", System.Data.OleDb.OleDbType.VarWChar, 70, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "JobName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_OrderNo", System.Data.OleDb.OleDbType.VarWChar, 25, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "OrderNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_OrderNo1", System.Data.OleDb.OleDbType.VarWChar, 25, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "OrderNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Tons_Or_Kilograms", System.Data.OleDb.OleDbType.VarWChar, 10, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Tons Or Kilograms", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Tons_Or_Kilograms1", System.Data.OleDb.OleDbType.VarWChar, 10, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Tons Or Kilograms", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_UnitOfMeas", System.Data.OleDb.OleDbType.VarWChar, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "UnitOfMeas", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_UnitOfMeas1", System.Data.OleDb.OleDbType.VarWChar, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "UnitOfMeas", System.Data.DataRowVersion.Original, Nothing))
        '
        'OleDbInsertCommand1
        '
        Me.OleDbInsertCommand1.CommandText = "INSERT INTO Job(AddedDiscount, CompanyNo, ContractorNo, Design, Discount, JobName" & _
        ", JobNo, OrderNo, [Tons Or Kilograms], UnitOfMeas) VALUES (?, ?, ?, ?, ?, ?, ?, " & _
        "?, ?, ?)"
        Me.OleDbInsertCommand1.Connection = Me.conJob
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("AddedDiscount", System.Data.OleDb.OleDbType.Currency, 0, "AddedDiscount"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("CompanyNo", System.Data.OleDb.OleDbType.VarWChar, 50, "CompanyNo"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("ContractorNo", System.Data.OleDb.OleDbType.VarWChar, 50, "ContractorNo"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Design", System.Data.OleDb.OleDbType.Currency, 0, "Design"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Discount", System.Data.OleDb.OleDbType.Integer, 0, "Discount"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("JobName", System.Data.OleDb.OleDbType.VarWChar, 70, "JobName"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("JobNo", System.Data.OleDb.OleDbType.VarWChar, 50, "JobNo"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("OrderNo", System.Data.OleDb.OleDbType.VarWChar, 25, "OrderNo"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Tons_Or_Kilograms", System.Data.OleDb.OleDbType.VarWChar, 10, "Tons Or Kilograms"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("UnitOfMeas", System.Data.OleDb.OleDbType.VarWChar, 8, "UnitOfMeas"))
        '
        'OleDbSelectCommand1
        '
        Me.OleDbSelectCommand1.CommandText = "SELECT AddedDiscount, CompanyNo, ContractorNo, Design, Discount, JobName, JobNo, " & _
        "OrderNo, [Tons Or Kilograms], UnitOfMeas FROM Job ORDER BY JobNo"
        Me.OleDbSelectCommand1.Connection = Me.conJob
        '
        'OleDbUpdateCommand1
        '
        Me.OleDbUpdateCommand1.CommandText = "UPDATE Job SET AddedDiscount = ?, CompanyNo = ?, ContractorNo = ?, Design = ?, Di" & _
        "scount = ?, JobName = ?, JobNo = ?, OrderNo = ?, [Tons Or Kilograms] = ?, UnitOf" & _
        "Meas = ? WHERE (JobNo = ?) AND (AddedDiscount = ? OR ? IS NULL AND AddedDiscount" & _
        " IS NULL) AND (CompanyNo = ? OR ? IS NULL AND CompanyNo IS NULL) AND (Contractor" & _
        "No = ? OR ? IS NULL AND ContractorNo IS NULL) AND (Design = ? OR ? IS NULL AND D" & _
        "esign IS NULL) AND (Discount = ? OR ? IS NULL AND Discount IS NULL) AND (JobName" & _
        " = ? OR ? IS NULL AND JobName IS NULL) AND (OrderNo = ? OR ? IS NULL AND OrderNo" & _
        " IS NULL) AND ([Tons Or Kilograms] = ? OR ? IS NULL AND [Tons Or Kilograms] IS N" & _
        "ULL) AND (UnitOfMeas = ? OR ? IS NULL AND UnitOfMeas IS NULL)"
        Me.OleDbUpdateCommand1.Connection = Me.conJob
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("AddedDiscount", System.Data.OleDb.OleDbType.Currency, 0, "AddedDiscount"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("CompanyNo", System.Data.OleDb.OleDbType.VarWChar, 50, "CompanyNo"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("ContractorNo", System.Data.OleDb.OleDbType.VarWChar, 50, "ContractorNo"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Design", System.Data.OleDb.OleDbType.Currency, 0, "Design"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Discount", System.Data.OleDb.OleDbType.Integer, 0, "Discount"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("JobName", System.Data.OleDb.OleDbType.VarWChar, 70, "JobName"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("JobNo", System.Data.OleDb.OleDbType.VarWChar, 50, "JobNo"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("OrderNo", System.Data.OleDb.OleDbType.VarWChar, 25, "OrderNo"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Tons_Or_Kilograms", System.Data.OleDb.OleDbType.VarWChar, 10, "Tons Or Kilograms"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("UnitOfMeas", System.Data.OleDb.OleDbType.VarWChar, 8, "UnitOfMeas"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_JobNo", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "JobNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddedDiscount", System.Data.OleDb.OleDbType.Currency, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddedDiscount", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddedDiscount1", System.Data.OleDb.OleDbType.Currency, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddedDiscount", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_CompanyNo", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CompanyNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_CompanyNo1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CompanyNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ContractorNo", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ContractorNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ContractorNo1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ContractorNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Design", System.Data.OleDb.OleDbType.Currency, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Design", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Design1", System.Data.OleDb.OleDbType.Currency, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Design", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Discount", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Discount", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Discount1", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Discount", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_JobName", System.Data.OleDb.OleDbType.VarWChar, 70, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "JobName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_JobName1", System.Data.OleDb.OleDbType.VarWChar, 70, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "JobName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_OrderNo", System.Data.OleDb.OleDbType.VarWChar, 25, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "OrderNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_OrderNo1", System.Data.OleDb.OleDbType.VarWChar, 25, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "OrderNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Tons_Or_Kilograms", System.Data.OleDb.OleDbType.VarWChar, 10, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Tons Or Kilograms", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Tons_Or_Kilograms1", System.Data.OleDb.OleDbType.VarWChar, 10, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Tons Or Kilograms", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_UnitOfMeas", System.Data.OleDb.OleDbType.VarWChar, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "UnitOfMeas", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_UnitOfMeas1", System.Data.OleDb.OleDbType.VarWChar, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "UnitOfMeas", System.Data.DataRowVersion.Original, Nothing))
        '
        'adpCompany
        '
        Me.adpCompany.DeleteCommand = Me.OleDbDeleteCommand2
        Me.adpCompany.InsertCommand = Me.OleDbInsertCommand2
        Me.adpCompany.SelectCommand = Me.OleDbSelectCommand2
        Me.adpCompany.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "Company", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("Address", "Address"), New System.Data.Common.DataColumnMapping("AddressLine2", "AddressLine2"), New System.Data.Common.DataColumnMapping("AddressLine3", "AddressLine3"), New System.Data.Common.DataColumnMapping("AddressLine4", "AddressLine4"), New System.Data.Common.DataColumnMapping("CompanyName", "CompanyName"), New System.Data.Common.DataColumnMapping("CompanyNo", "CompanyNo"), New System.Data.Common.DataColumnMapping("Email", "Email"), New System.Data.Common.DataColumnMapping("Fax", "Fax"), New System.Data.Common.DataColumnMapping("LastCutNum", "LastCutNum"), New System.Data.Common.DataColumnMapping("LastInvNum", "LastInvNum"), New System.Data.Common.DataColumnMapping("Message", "Message"), New System.Data.Common.DataColumnMapping("PostalCode", "PostalCode"), New System.Data.Common.DataColumnMapping("RegNo", "RegNo"), New System.Data.Common.DataColumnMapping("Telephone", "Telephone"), New System.Data.Common.DataColumnMapping("UnitOfMeas", "UnitOfMeas"), New System.Data.Common.DataColumnMapping("VatNo", "VatNo"), New System.Data.Common.DataColumnMapping("VatPerc", "VatPerc"), New System.Data.Common.DataColumnMapping("Website", "Website")})})
        Me.adpCompany.UpdateCommand = Me.OleDbUpdateCommand2
        '
        'OleDbDeleteCommand2
        '
        Me.OleDbDeleteCommand2.CommandText = "DELETE FROM Company WHERE (CompanyNo = ?) AND (AddressLine2 = ? OR ? IS NULL AND " & _
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
        Me.OleDbDeleteCommand2.Connection = Me.conJob
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_CompanyNo", System.Data.OleDb.OleDbType.VarWChar, 10, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CompanyNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine2", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine2", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine21", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine2", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine3", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine3", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine31", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine3", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine4", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine4", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine41", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine4", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_CompanyName", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CompanyName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_CompanyName1", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CompanyName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Email", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Email", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Email1", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Email", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Fax", System.Data.OleDb.OleDbType.VarWChar, 15, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Fax", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Fax1", System.Data.OleDb.OleDbType.VarWChar, 15, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Fax", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_LastCutNum", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "LastCutNum", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_LastCutNum1", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "LastCutNum", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_LastInvNum", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "LastInvNum", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_LastInvNum1", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "LastInvNum", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Message", System.Data.OleDb.OleDbType.VarWChar, 200, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Message", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Message1", System.Data.OleDb.OleDbType.VarWChar, 200, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Message", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_PostalCode", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "PostalCode", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_PostalCode1", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "PostalCode", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_RegNo", System.Data.OleDb.OleDbType.VarWChar, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "RegNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_RegNo1", System.Data.OleDb.OleDbType.VarWChar, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "RegNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Telephone", System.Data.OleDb.OleDbType.VarWChar, 15, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Telephone", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Telephone1", System.Data.OleDb.OleDbType.VarWChar, 15, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Telephone", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_UnitOfMeas", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "UnitOfMeas", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_UnitOfMeas1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "UnitOfMeas", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_VatNo", System.Data.OleDb.OleDbType.VarWChar, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "VatNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_VatNo1", System.Data.OleDb.OleDbType.VarWChar, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "VatNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_VatPerc", System.Data.OleDb.OleDbType.Decimal, 0, System.Data.ParameterDirection.Input, False, CType(2, Byte), CType(2, Byte), "VatPerc", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_VatPerc1", System.Data.OleDb.OleDbType.Decimal, 0, System.Data.ParameterDirection.Input, False, CType(2, Byte), CType(2, Byte), "VatPerc", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Website", System.Data.OleDb.OleDbType.VarWChar, 30, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Website", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Website1", System.Data.OleDb.OleDbType.VarWChar, 30, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Website", System.Data.DataRowVersion.Original, Nothing))
        '
        'OleDbInsertCommand2
        '
        Me.OleDbInsertCommand2.CommandText = "INSERT INTO Company(Address, AddressLine2, AddressLine3, AddressLine4, CompanyNam" & _
        "e, CompanyNo, Email, Fax, LastCutNum, LastInvNum, Message, PostalCode, RegNo, Te" & _
        "lephone, UnitOfMeas, VatNo, VatPerc, Website) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?," & _
        " ?, ?, ?, ?, ?, ?, ?, ?, ?)"
        Me.OleDbInsertCommand2.Connection = Me.conJob
        Me.OleDbInsertCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Address", System.Data.OleDb.OleDbType.VarWChar, 0, "Address"))
        Me.OleDbInsertCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("AddressLine2", System.Data.OleDb.OleDbType.VarWChar, 40, "AddressLine2"))
        Me.OleDbInsertCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("AddressLine3", System.Data.OleDb.OleDbType.VarWChar, 40, "AddressLine3"))
        Me.OleDbInsertCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("AddressLine4", System.Data.OleDb.OleDbType.VarWChar, 40, "AddressLine4"))
        Me.OleDbInsertCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("CompanyName", System.Data.OleDb.OleDbType.VarWChar, 40, "CompanyName"))
        Me.OleDbInsertCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("CompanyNo", System.Data.OleDb.OleDbType.VarWChar, 10, "CompanyNo"))
        Me.OleDbInsertCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Email", System.Data.OleDb.OleDbType.VarWChar, 40, "Email"))
        Me.OleDbInsertCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Fax", System.Data.OleDb.OleDbType.VarWChar, 15, "Fax"))
        Me.OleDbInsertCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("LastCutNum", System.Data.OleDb.OleDbType.Integer, 0, "LastCutNum"))
        Me.OleDbInsertCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("LastInvNum", System.Data.OleDb.OleDbType.Integer, 0, "LastInvNum"))
        Me.OleDbInsertCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Message", System.Data.OleDb.OleDbType.VarWChar, 200, "Message"))
        Me.OleDbInsertCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("PostalCode", System.Data.OleDb.OleDbType.Integer, 0, "PostalCode"))
        Me.OleDbInsertCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("RegNo", System.Data.OleDb.OleDbType.VarWChar, 20, "RegNo"))
        Me.OleDbInsertCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Telephone", System.Data.OleDb.OleDbType.VarWChar, 15, "Telephone"))
        Me.OleDbInsertCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("UnitOfMeas", System.Data.OleDb.OleDbType.VarWChar, 50, "UnitOfMeas"))
        Me.OleDbInsertCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("VatNo", System.Data.OleDb.OleDbType.VarWChar, 20, "VatNo"))
        Me.OleDbInsertCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("VatPerc", System.Data.OleDb.OleDbType.Decimal, 0, System.Data.ParameterDirection.Input, False, CType(2, Byte), CType(2, Byte), "VatPerc", System.Data.DataRowVersion.Current, Nothing))
        Me.OleDbInsertCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Website", System.Data.OleDb.OleDbType.VarWChar, 30, "Website"))
        '
        'OleDbSelectCommand2
        '
        Me.OleDbSelectCommand2.CommandText = "SELECT Address, AddressLine2, AddressLine3, AddressLine4, CompanyName, CompanyNo," & _
        " Email, Fax, LastCutNum, LastInvNum, Message, PostalCode, RegNo, Telephone, Unit" & _
        "OfMeas, VatNo, VatPerc, Website FROM Company"
        Me.OleDbSelectCommand2.Connection = Me.conJob
        '
        'OleDbUpdateCommand2
        '
        Me.OleDbUpdateCommand2.CommandText = "UPDATE Company SET Address = ?, AddressLine2 = ?, AddressLine3 = ?, AddressLine4 " & _
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
        Me.OleDbUpdateCommand2.Connection = Me.conJob
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Address", System.Data.OleDb.OleDbType.VarWChar, 0, "Address"))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("AddressLine2", System.Data.OleDb.OleDbType.VarWChar, 40, "AddressLine2"))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("AddressLine3", System.Data.OleDb.OleDbType.VarWChar, 40, "AddressLine3"))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("AddressLine4", System.Data.OleDb.OleDbType.VarWChar, 40, "AddressLine4"))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("CompanyName", System.Data.OleDb.OleDbType.VarWChar, 40, "CompanyName"))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("CompanyNo", System.Data.OleDb.OleDbType.VarWChar, 10, "CompanyNo"))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Email", System.Data.OleDb.OleDbType.VarWChar, 40, "Email"))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Fax", System.Data.OleDb.OleDbType.VarWChar, 15, "Fax"))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("LastCutNum", System.Data.OleDb.OleDbType.Integer, 0, "LastCutNum"))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("LastInvNum", System.Data.OleDb.OleDbType.Integer, 0, "LastInvNum"))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Message", System.Data.OleDb.OleDbType.VarWChar, 200, "Message"))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("PostalCode", System.Data.OleDb.OleDbType.Integer, 0, "PostalCode"))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("RegNo", System.Data.OleDb.OleDbType.VarWChar, 20, "RegNo"))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Telephone", System.Data.OleDb.OleDbType.VarWChar, 15, "Telephone"))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("UnitOfMeas", System.Data.OleDb.OleDbType.VarWChar, 50, "UnitOfMeas"))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("VatNo", System.Data.OleDb.OleDbType.VarWChar, 20, "VatNo"))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("VatPerc", System.Data.OleDb.OleDbType.Decimal, 0, System.Data.ParameterDirection.Input, False, CType(2, Byte), CType(2, Byte), "VatPerc", System.Data.DataRowVersion.Current, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Website", System.Data.OleDb.OleDbType.VarWChar, 30, "Website"))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_CompanyNo", System.Data.OleDb.OleDbType.VarWChar, 10, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CompanyNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine2", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine2", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine21", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine2", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine3", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine3", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine31", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine3", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine4", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine4", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine41", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine4", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_CompanyName", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CompanyName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_CompanyName1", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CompanyName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Email", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Email", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Email1", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Email", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Fax", System.Data.OleDb.OleDbType.VarWChar, 15, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Fax", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Fax1", System.Data.OleDb.OleDbType.VarWChar, 15, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Fax", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_LastCutNum", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "LastCutNum", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_LastCutNum1", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "LastCutNum", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_LastInvNum", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "LastInvNum", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_LastInvNum1", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "LastInvNum", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Message", System.Data.OleDb.OleDbType.VarWChar, 200, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Message", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Message1", System.Data.OleDb.OleDbType.VarWChar, 200, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Message", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_PostalCode", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "PostalCode", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_PostalCode1", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "PostalCode", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_RegNo", System.Data.OleDb.OleDbType.VarWChar, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "RegNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_RegNo1", System.Data.OleDb.OleDbType.VarWChar, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "RegNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Telephone", System.Data.OleDb.OleDbType.VarWChar, 15, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Telephone", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Telephone1", System.Data.OleDb.OleDbType.VarWChar, 15, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Telephone", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_UnitOfMeas", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "UnitOfMeas", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_UnitOfMeas1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "UnitOfMeas", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_VatNo", System.Data.OleDb.OleDbType.VarWChar, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "VatNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_VatNo1", System.Data.OleDb.OleDbType.VarWChar, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "VatNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_VatPerc", System.Data.OleDb.OleDbType.Decimal, 0, System.Data.ParameterDirection.Input, False, CType(2, Byte), CType(2, Byte), "VatPerc", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_VatPerc1", System.Data.OleDb.OleDbType.Decimal, 0, System.Data.ParameterDirection.Input, False, CType(2, Byte), CType(2, Byte), "VatPerc", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Website", System.Data.OleDb.OleDbType.VarWChar, 30, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Website", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Website1", System.Data.OleDb.OleDbType.VarWChar, 30, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Website", System.Data.DataRowVersion.Original, Nothing))
        '
        'adpContractor
        '
        Me.adpContractor.DeleteCommand = Me.OleDbDeleteCommand3
        Me.adpContractor.InsertCommand = Me.OleDbInsertCommand3
        Me.adpContractor.SelectCommand = Me.OleDbSelectCommand3
        Me.adpContractor.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "Contractor", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("ActiveY/N", "ActiveY/N"), New System.Data.Common.DataColumnMapping("AddressLine1", "AddressLine1"), New System.Data.Common.DataColumnMapping("AddressLine2", "AddressLine2"), New System.Data.Common.DataColumnMapping("AddressLine3", "AddressLine3"), New System.Data.Common.DataColumnMapping("AddressLine4", "AddressLine4"), New System.Data.Common.DataColumnMapping("ContractorName", "ContractorName"), New System.Data.Common.DataColumnMapping("ContractorNo", "ContractorNo"), New System.Data.Common.DataColumnMapping("PostalCode", "PostalCode"), New System.Data.Common.DataColumnMapping("Telephone", "Telephone")})})
        Me.adpContractor.UpdateCommand = Me.OleDbUpdateCommand3
        '
        'OleDbDeleteCommand3
        '
        Me.OleDbDeleteCommand3.CommandText = "DELETE FROM Contractor WHERE (ContractorNo = ?) AND ([ActiveY/N] = ?) AND (Addres" & _
        "sLine1 = ? OR ? IS NULL AND AddressLine1 IS NULL) AND (AddressLine2 = ? OR ? IS " & _
        "NULL AND AddressLine2 IS NULL) AND (AddressLine3 = ? OR ? IS NULL AND AddressLin" & _
        "e3 IS NULL) AND (AddressLine4 = ? OR ? IS NULL AND AddressLine4 IS NULL) AND (Co" & _
        "ntractorName = ? OR ? IS NULL AND ContractorName IS NULL) AND (PostalCode = ? OR" & _
        " ? IS NULL AND PostalCode IS NULL) AND (Telephone = ? OR ? IS NULL AND Telephone" & _
        " IS NULL)"
        Me.OleDbDeleteCommand3.Connection = Me.conJob
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ContractorNo", System.Data.OleDb.OleDbType.VarWChar, 10, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ContractorNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ActiveY_N", System.Data.OleDb.OleDbType.Boolean, 2, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ActiveY/N", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine1", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine1", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine11", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine1", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine2", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine2", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine21", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine2", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine3", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine3", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine31", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine3", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine4", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine4", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine41", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine4", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ContractorName", System.Data.OleDb.OleDbType.VarWChar, 70, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ContractorName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ContractorName1", System.Data.OleDb.OleDbType.VarWChar, 70, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ContractorName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_PostalCode", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "PostalCode", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_PostalCode1", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "PostalCode", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Telephone", System.Data.OleDb.OleDbType.VarWChar, 15, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Telephone", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Telephone1", System.Data.OleDb.OleDbType.VarWChar, 15, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Telephone", System.Data.DataRowVersion.Original, Nothing))
        '
        'OleDbInsertCommand3
        '
        Me.OleDbInsertCommand3.CommandText = "INSERT INTO Contractor([ActiveY/N], AddressLine1, AddressLine2, AddressLine3, Add" & _
        "ressLine4, ContractorName, ContractorNo, PostalCode, Telephone) VALUES (?, ?, ?," & _
        " ?, ?, ?, ?, ?, ?)"
        Me.OleDbInsertCommand3.Connection = Me.conJob
        Me.OleDbInsertCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("ActiveY_N", System.Data.OleDb.OleDbType.Boolean, 2, "ActiveY/N"))
        Me.OleDbInsertCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("AddressLine1", System.Data.OleDb.OleDbType.VarWChar, 40, "AddressLine1"))
        Me.OleDbInsertCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("AddressLine2", System.Data.OleDb.OleDbType.VarWChar, 40, "AddressLine2"))
        Me.OleDbInsertCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("AddressLine3", System.Data.OleDb.OleDbType.VarWChar, 40, "AddressLine3"))
        Me.OleDbInsertCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("AddressLine4", System.Data.OleDb.OleDbType.VarWChar, 40, "AddressLine4"))
        Me.OleDbInsertCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("ContractorName", System.Data.OleDb.OleDbType.VarWChar, 70, "ContractorName"))
        Me.OleDbInsertCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("ContractorNo", System.Data.OleDb.OleDbType.VarWChar, 10, "ContractorNo"))
        Me.OleDbInsertCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("PostalCode", System.Data.OleDb.OleDbType.Integer, 0, "PostalCode"))
        Me.OleDbInsertCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Telephone", System.Data.OleDb.OleDbType.VarWChar, 15, "Telephone"))
        '
        'OleDbSelectCommand3
        '
        Me.OleDbSelectCommand3.CommandText = "SELECT [ActiveY/N], AddressLine1, AddressLine2, AddressLine3, AddressLine4, Contr" & _
        "actorName, ContractorNo, PostalCode, Telephone FROM Contractor"
        Me.OleDbSelectCommand3.Connection = Me.conJob
        '
        'OleDbUpdateCommand3
        '
        Me.OleDbUpdateCommand3.CommandText = "UPDATE Contractor SET [ActiveY/N] = ?, AddressLine1 = ?, AddressLine2 = ?, Addres" & _
        "sLine3 = ?, AddressLine4 = ?, ContractorName = ?, ContractorNo = ?, PostalCode =" & _
        " ?, Telephone = ? WHERE (ContractorNo = ?) AND ([ActiveY/N] = ?) AND (AddressLin" & _
        "e1 = ? OR ? IS NULL AND AddressLine1 IS NULL) AND (AddressLine2 = ? OR ? IS NULL" & _
        " AND AddressLine2 IS NULL) AND (AddressLine3 = ? OR ? IS NULL AND AddressLine3 I" & _
        "S NULL) AND (AddressLine4 = ? OR ? IS NULL AND AddressLine4 IS NULL) AND (Contra" & _
        "ctorName = ? OR ? IS NULL AND ContractorName IS NULL) AND (PostalCode = ? OR ? I" & _
        "S NULL AND PostalCode IS NULL) AND (Telephone = ? OR ? IS NULL AND Telephone IS " & _
        "NULL)"
        Me.OleDbUpdateCommand3.Connection = Me.conJob
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("ActiveY_N", System.Data.OleDb.OleDbType.Boolean, 2, "ActiveY/N"))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("AddressLine1", System.Data.OleDb.OleDbType.VarWChar, 40, "AddressLine1"))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("AddressLine2", System.Data.OleDb.OleDbType.VarWChar, 40, "AddressLine2"))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("AddressLine3", System.Data.OleDb.OleDbType.VarWChar, 40, "AddressLine3"))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("AddressLine4", System.Data.OleDb.OleDbType.VarWChar, 40, "AddressLine4"))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("ContractorName", System.Data.OleDb.OleDbType.VarWChar, 70, "ContractorName"))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("ContractorNo", System.Data.OleDb.OleDbType.VarWChar, 10, "ContractorNo"))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("PostalCode", System.Data.OleDb.OleDbType.Integer, 0, "PostalCode"))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Telephone", System.Data.OleDb.OleDbType.VarWChar, 15, "Telephone"))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ContractorNo", System.Data.OleDb.OleDbType.VarWChar, 10, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ContractorNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ActiveY_N", System.Data.OleDb.OleDbType.Boolean, 2, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ActiveY/N", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine1", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine1", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine11", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine1", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine2", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine2", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine21", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine2", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine3", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine3", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine31", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine3", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine4", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine4", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddressLine41", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddressLine4", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ContractorName", System.Data.OleDb.OleDbType.VarWChar, 70, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ContractorName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ContractorName1", System.Data.OleDb.OleDbType.VarWChar, 70, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ContractorName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_PostalCode", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "PostalCode", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_PostalCode1", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "PostalCode", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Telephone", System.Data.OleDb.OleDbType.VarWChar, 15, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Telephone", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Telephone1", System.Data.OleDb.OleDbType.VarWChar, 15, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Telephone", System.Data.DataRowVersion.Original, Nothing))
        '
        'cmdCountJobNo
        '
        Me.cmdCountJobNo.CommandText = "SELECT Job.* FROM Job WHERE (JobNo = ?)ORDER BY JobNo"
        Me.cmdCountJobNo.Connection = Me.conJob
        Me.cmdCountJobNo.Parameters.Add(New System.Data.OleDb.OleDbParameter("JobNo", System.Data.OleDb.OleDbType.VarWChar, 50, "JobNo"))
        '
        'cbxJobNo
        '
        Me.cbxJobNo.DataSource = Me.dsJob
        Me.cbxJobNo.DisplayMember = "Job.No&Name"
        Me.cbxJobNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxJobNo.Location = New System.Drawing.Point(144, 48)
        Me.cbxJobNo.Name = "cbxJobNo"
        Me.cbxJobNo.Size = New System.Drawing.Size(344, 21)
        Me.cbxJobNo.TabIndex = 5
        Me.cbxJobNo.ValueMember = "Job.JobNo"
        '
        'btnEdit
        '
        Me.btnEdit.Location = New System.Drawing.Point(104, 296)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "Edit"
        '
        'adpJobRate
        '
        Me.adpJobRate.DeleteCommand = Me.OleDbDeleteCommand4
        Me.adpJobRate.InsertCommand = Me.OleDbInsertCommand4
        Me.adpJobRate.SelectCommand = Me.OleDbSelectCommand4
        Me.adpJobRate.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "JobRate", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("JobNo", "JobNo"), New System.Data.Common.DataColumnMapping("Rate", "Rate"), New System.Data.Common.DataColumnMapping("TypeCode", "TypeCode")})})
        Me.adpJobRate.UpdateCommand = Me.OleDbUpdateCommand4
        '
        'OleDbDeleteCommand4
        '
        Me.OleDbDeleteCommand4.CommandText = "DELETE FROM JobRate WHERE (JobNo = ?) AND (TypeCode = ?) AND (Rate = ? OR ? IS NU" & _
        "LL AND Rate IS NULL)"
        Me.OleDbDeleteCommand4.Connection = Me.conJob
        Me.OleDbDeleteCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_JobNo", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "JobNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_TypeCode", System.Data.OleDb.OleDbType.VarWChar, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "TypeCode", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Rate", System.Data.OleDb.OleDbType.Currency, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Rate", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Rate1", System.Data.OleDb.OleDbType.Currency, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Rate", System.Data.DataRowVersion.Original, Nothing))
        '
        'OleDbInsertCommand4
        '
        Me.OleDbInsertCommand4.CommandText = "INSERT INTO JobRate(JobNo, Rate, TypeCode) VALUES (?, ?, ?)"
        Me.OleDbInsertCommand4.Connection = Me.conJob
        Me.OleDbInsertCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("JobNo", System.Data.OleDb.OleDbType.VarWChar, 50, "JobNo"))
        Me.OleDbInsertCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Rate", System.Data.OleDb.OleDbType.Currency, 0, "Rate"))
        Me.OleDbInsertCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("TypeCode", System.Data.OleDb.OleDbType.VarWChar, 8, "TypeCode"))
        '
        'OleDbSelectCommand4
        '
        Me.OleDbSelectCommand4.CommandText = "SELECT JobNo, Rate, TypeCode FROM JobRate"
        Me.OleDbSelectCommand4.Connection = Me.conJob
        '
        'OleDbUpdateCommand4
        '
        Me.OleDbUpdateCommand4.CommandText = "UPDATE JobRate SET JobNo = ?, Rate = ?, TypeCode = ? WHERE (JobNo = ?) AND (TypeC" & _
        "ode = ?) AND (Rate = ? OR ? IS NULL AND Rate IS NULL)"
        Me.OleDbUpdateCommand4.Connection = Me.conJob
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("JobNo", System.Data.OleDb.OleDbType.VarWChar, 50, "JobNo"))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Rate", System.Data.OleDb.OleDbType.Currency, 0, "Rate"))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("TypeCode", System.Data.OleDb.OleDbType.VarWChar, 8, "TypeCode"))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_JobNo", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "JobNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_TypeCode", System.Data.OleDb.OleDbType.VarWChar, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "TypeCode", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Rate", System.Data.OleDb.OleDbType.Currency, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Rate", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Rate1", System.Data.OleDb.OleDbType.Currency, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Rate", System.Data.DataRowVersion.Original, Nothing))
        '
        'adpProductType
        '
        Me.adpProductType.DeleteCommand = Me.OleDbDeleteCommand5
        Me.adpProductType.InsertCommand = Me.OleDbInsertCommand5
        Me.adpProductType.SelectCommand = Me.OleDbSelectCommand5
        Me.adpProductType.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "ProductType", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("TypeCode", "TypeCode"), New System.Data.Common.DataColumnMapping("Weight", "Weight")})})
        Me.adpProductType.UpdateCommand = Me.OleDbUpdateCommand5
        '
        'OleDbDeleteCommand5
        '
        Me.OleDbDeleteCommand5.CommandText = "DELETE FROM ProductType WHERE (TypeCode = ?) AND (Weight = ? OR ? IS NULL AND Wei" & _
        "ght IS NULL)"
        Me.OleDbDeleteCommand5.Connection = Me.conJob
        Me.OleDbDeleteCommand5.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_TypeCode", System.Data.OleDb.OleDbType.VarWChar, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "TypeCode", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand5.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Weight", System.Data.OleDb.OleDbType.Double, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Weight", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand5.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Weight1", System.Data.OleDb.OleDbType.Double, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Weight", System.Data.DataRowVersion.Original, Nothing))
        '
        'OleDbInsertCommand5
        '
        Me.OleDbInsertCommand5.CommandText = "INSERT INTO ProductType(TypeCode, Weight) VALUES (?, ?)"
        Me.OleDbInsertCommand5.Connection = Me.conJob
        Me.OleDbInsertCommand5.Parameters.Add(New System.Data.OleDb.OleDbParameter("TypeCode", System.Data.OleDb.OleDbType.VarWChar, 8, "TypeCode"))
        Me.OleDbInsertCommand5.Parameters.Add(New System.Data.OleDb.OleDbParameter("Weight", System.Data.OleDb.OleDbType.Double, 0, "Weight"))
        '
        'OleDbSelectCommand5
        '
        Me.OleDbSelectCommand5.CommandText = "SELECT TypeCode, Weight FROM ProductType"
        Me.OleDbSelectCommand5.Connection = Me.conJob
        '
        'OleDbUpdateCommand5
        '
        Me.OleDbUpdateCommand5.CommandText = "UPDATE ProductType SET TypeCode = ?, Weight = ? WHERE (TypeCode = ?) AND (Weight " & _
        "= ? OR ? IS NULL AND Weight IS NULL)"
        Me.OleDbUpdateCommand5.Connection = Me.conJob
        Me.OleDbUpdateCommand5.Parameters.Add(New System.Data.OleDb.OleDbParameter("TypeCode", System.Data.OleDb.OleDbType.VarWChar, 8, "TypeCode"))
        Me.OleDbUpdateCommand5.Parameters.Add(New System.Data.OleDb.OleDbParameter("Weight", System.Data.OleDb.OleDbType.Double, 0, "Weight"))
        Me.OleDbUpdateCommand5.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_TypeCode", System.Data.OleDb.OleDbType.VarWChar, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "TypeCode", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand5.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Weight", System.Data.OleDb.OleDbType.Double, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Weight", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand5.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Weight1", System.Data.OleDb.OleDbType.Double, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Weight", System.Data.DataRowVersion.Original, Nothing))
        '
        'btnArchive
        '
        Me.btnArchive.Location = New System.Drawing.Point(232, 296)
        Me.btnArchive.Name = "btnArchive"
        Me.btnArchive.TabIndex = 6
        Me.btnArchive.Text = "Archive"
        '
        'frmJob
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(530, 336)
        Me.Controls.Add(Me.btnArchive)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.cbxJobNo)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnAdd)
        Me.Controls.Add(Me.grpJobDetails)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "frmJob"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Job Maintenance"
        Me.grpJobDetails.ResumeLayout(False)
        CType(Me.dsJob, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
    Dim DBConnection As New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source = winsteelVers5.mdb")

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

    Private Sub frmJob_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If Not IsNothing(CallingForm) Then
            CallingForm.Show()
        End If

        CallingForm = Nothing
    End Sub

    Private Sub frmJob_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        enablebinding()

        dsJob.Clear()
        adpContractor.Fill(dsJob.Contractor)
        adpCompany.Fill(dsJob.Company)
        adpJob.Fill(dsJob.Job)

        grpJobDetails.Enabled = False
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Close()
    End Sub

    Private Sub disablebinding()
        Dim temp As Object = Nothing

        txtJobNo.DataBindings.Clear()
        txtJobNo.Clear()

        txtJobName.DataBindings.Clear()
        txtJobName.Clear()

        txtDiscount.DataBindings.Clear()
        txtDiscount.Clear()

        txtAddDiscount.DataBindings.Clear()
        txtAddDiscount.Clear()

        txtDesignCost.DataBindings.Clear()
        txtDesignCost.Clear()

        txtOrderNo.DataBindings.Clear()
        txtOrderNo.Clear()

        cbxTonsKilograms.DataSource = temp
        cbxTonsKilograms.Items.Add("T")
        cbxTonsKilograms.Items.Add("Kg")
        cbxTonsKilograms.SelectedIndex = 0
    End Sub

    Private Sub enablebinding()
        txtJobNo.DataBindings.Add("Text", dsJob, "Job.JobNo")
        txtJobName.DataBindings.Add("Text", dsJob, "Job.JobName")
        txtDiscount.DataBindings.Add("Text", dsJob, "Job.Discount")
        txtAddDiscount.DataBindings.Add("Text", dsJob, "Job.AddedDiscount")
        txtDesignCost.DataBindings.Add("Text", dsJob, "Job.Design")
        txtOrderNo.DataBindings.Add("Text", dsJob, "Job.OrderNo")
        cbxTonsKilograms.DataSource = dsJob
        cbxTonsKilograms.DisplayMember = "Job.Tons Or Kilograms"
        cbxTonsKilograms.ValueMember = "Job.Tons Or Kilograms"
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        If state = "" Then
            cbxJobNo.SendToBack()
            cbxJobNo.Enabled = False

            grpJobDetails.Enabled = True
            disablebinding()

            state = "add"
            txtJobNo.Focus()
        End If
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If state = "add" Then
            If txtJobNo.Text = "" Then
                MsgBox("A Job Number is required", MsgBoxStyle.Critical, "Error")
                txtJobNo.Focus()
            Else
                Dim DataReader As System.Data.OleDb.OleDbDataReader
                Dim count As Integer

                conJob.Open()
                cmdCountJobNo.Parameters("JobNo").Value = txtJobNo.Text
                DataReader = cmdCountJobNo.ExecuteReader(CommandBehavior.CloseConnection)
                While DataReader.Read()
                    count += 1
                End While
                DataReader.Close()
                conJob.Close()

                If count > 0 Then
                    MsgBox("Job Number entered is already used", MsgBoxStyle.Critical, "Error")
                    txtJobNo.Focus()
                ElseIf cbxTonsKilograms.SelectedIndex = -1 Then
                    MsgBox("Tons or Kilograms must be selected", MsgBoxStyle.Critical, "Error")
                Else
                    Dim row As DataRow = dsJob.Job.NewJobRow
                    row("JobNo") = txtJobNo.Text
                    row("CompanyNo") = cbxCompanyNo.SelectedValue
                    row("ContractorNo") = cbxContractor.SelectedValue

                    If txtJobName.Text <> "" Then
                        row("JobName") = txtJobName.Text
                    End If
                    If txtDiscount.Text <> "" Then
                        row("Discount") = txtDiscount.Text
                    End If
                    If txtAddDiscount.Text <> "" Then
                        row("AddedDiscount") = txtAddDiscount.Text
                    End If
                    If txtDesignCost.Text <> "" Then
                        row("Design") = txtDesignCost.Text
                    End If
                    If txtOrderNo.Text <> "" Then
                        row("OrderNo") = txtOrderNo.Text
                    End If

                    row("Tons Or Kilograms") = cbxTonsKilograms.Text

                    dsJob.Job.AddJobRow(row)

                    Try
                        adpJob.Update(dsJob.Job)
                    Catch ex As Exception
                        MsgBox(ex.Message)
                    End Try
                    MsgBox("Record was successfully saved", MsgBoxStyle.Information, "Information")

                    PopulateJobRate()

                    enablebinding()

                    cbxJobNo.BringToFront()
                    cbxJobNo.Enabled = True

                    grpJobDetails.Enabled = False

                    state = ""
                End If
            End If
        End If

        If state = "edit" Then
            dsJob.Job.FindByJobNo(txtJobNo.Text).EndEdit()

            adpJob.Update(dsJob.Job)
            MsgBox("Record was successfully saved", MsgBoxStyle.Information, "Information")

            cbxJobNo.BringToFront()
            cbxJobNo.Enabled = True

            cbxTonsKilograms.DataSource = dsJob
            cbxTonsKilograms.DisplayMember = "Job.Tons Or Kilograms"
            cbxTonsKilograms.ValueMember = "Job.Tons Or Kilograms"

            grpJobDetails.Enabled = False
            txtJobNo.Enabled = True

            state = ""
        End If
    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Dim temp As Object = Nothing
        Dim s As String

        If state = "" Then
            cbxJobNo.SendToBack()
            cbxJobNo.Enabled = False

            grpJobDetails.Enabled = True
            txtJobNo.Enabled = False

            s = cbxTonsKilograms.SelectedValue

            cbxTonsKilograms.DataSource = temp
            cbxTonsKilograms.Items.Add("T")
            cbxTonsKilograms.Items.Add("Kg")
            cbxTonsKilograms.SelectedItem = s

            state = "edit"
            cbxCompanyNo.Focus()
        End If
    End Sub

    Private Sub PopulateJobRate()
        adpJobRate.Fill(dsJob.JobRate)
        adpProductType.Fill(dsJob.ProductType)

        Dim reader As System.Data.OleDb.OleDbDataReader
        Dim cmd As New System.Data.OleDb.OleDbCommand("SELECT * FROM ProductType", conJob)

        conJob.Open()
        reader = cmd.ExecuteReader(CommandBehavior.CloseConnection)
        While reader.Read()
            Dim row As DataRow = dsJob.JobRate.NewJobRateRow
            row("JobNo") = txtJobNo.Text
            row("TypeCode") = reader("TypeCode")
            row("Rate") = 0
            dsJob.JobRate.AddJobRateRow(row)
        End While
        reader.Close()
        conJob.Close()

        adpJobRate.Update(dsJob.JobRate)
    End Sub

    Private Sub btnArchive_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnArchive.Click
        Dim row As DataRow = dsJob.Job.FindByJobNo(txtJobNo.Text)
        Dim Form As frmJobArchive = New frmJobArchive(Me, DBConnection)
        Me.Hide()
        Form.Show()
        MsgBox("Record was successfully deleted", MsgBoxStyle.Information, "Information")
    End Sub

    Private Sub grpJobDetails_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles grpJobDetails.Enter

    End Sub

    Private Sub cbxJobNo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxJobNo.SelectedIndexChanged

    End Sub

    Private Sub txtJobName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtJobName.TextChanged

    End Sub

    Private Sub adpJob_RowUpdated(ByVal sender As System.Object, ByVal e As System.Data.OleDb.OleDbRowUpdatedEventArgs) Handles adpJob.RowUpdated

    End Sub

    Private Sub cbxTonsKilograms_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxTonsKilograms.SelectedIndexChanged

    End Sub
End Class
