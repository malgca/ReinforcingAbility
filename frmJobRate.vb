Public Class frmJobRate
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
    Friend WithEvents lblJobNo As System.Windows.Forms.Label
    Friend WithEvents OleDbSelectCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbInsertCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbUpdateCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbDeleteCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents conJobRate As System.Data.OleDb.OleDbConnection
    Friend WithEvents adpJobRate As System.Data.OleDb.OleDbDataAdapter
    Friend WithEvents dsJobRate As PresentationTier.dsReinforcingAbility
    Friend WithEvents dvJobRate As System.Data.DataView
    Friend WithEvents grdJobRate As System.Windows.Forms.DataGrid
    Friend WithEvents adpJob As System.Data.OleDb.OleDbDataAdapter
    Friend WithEvents OleDbSelectCommand2 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbInsertCommand2 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbUpdateCommand2 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbDeleteCommand2 As System.Data.OleDb.OleDbCommand
    Friend WithEvents dvJob As System.Data.DataView
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents DataGridTableStyle1 As System.Windows.Forms.DataGridTableStyle
    Friend WithEvents colType As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents colRate As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents cmdTypeExist As System.Data.OleDb.OleDbCommand
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents cbxJobNo As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.lblJobNo = New System.Windows.Forms.Label
        Me.dvJob = New System.Data.DataView
        Me.dsJobRate = New PresentationTier.dsReinforcingAbility
        Me.dvJobRate = New System.Data.DataView
        Me.adpJobRate = New System.Data.OleDb.OleDbDataAdapter
        Me.OleDbDeleteCommand1 = New System.Data.OleDb.OleDbCommand
        Me.conJobRate = New System.Data.OleDb.OleDbConnection
        Me.OleDbInsertCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbSelectCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbUpdateCommand1 = New System.Data.OleDb.OleDbCommand
        Me.grdJobRate = New System.Windows.Forms.DataGrid
        Me.DataGridTableStyle1 = New System.Windows.Forms.DataGridTableStyle
        Me.colType = New System.Windows.Forms.DataGridTextBoxColumn
        Me.colRate = New System.Windows.Forms.DataGridTextBoxColumn
        Me.adpJob = New System.Data.OleDb.OleDbDataAdapter
        Me.OleDbDeleteCommand2 = New System.Data.OleDb.OleDbCommand
        Me.OleDbInsertCommand2 = New System.Data.OleDb.OleDbCommand
        Me.OleDbSelectCommand2 = New System.Data.OleDb.OleDbCommand
        Me.OleDbUpdateCommand2 = New System.Data.OleDb.OleDbCommand
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.cmdTypeExist = New System.Data.OleDb.OleDbCommand
        Me.cbxJobNo = New System.Windows.Forms.ComboBox
        CType(Me.dvJob, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dsJobRate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dvJobRate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdJobRate, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblJobNo
        '
        Me.lblJobNo.Location = New System.Drawing.Point(16, 16)
        Me.lblJobNo.Name = "lblJobNo"
        Me.lblJobNo.Size = New System.Drawing.Size(56, 23)
        Me.lblJobNo.TabIndex = 0
        Me.lblJobNo.Text = "Job No."
        Me.lblJobNo.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dvJob
        '
        Me.dvJob.Table = Me.dsJobRate.Job
        '
        'dsJobRate
        '
        Me.dsJobRate.DataSetName = "dsReinforcingAbility"
        Me.dsJobRate.Locale = New System.Globalization.CultureInfo("en-ZA")
        '
        'dvJobRate
        '
        Me.dvJobRate.AllowNew = False
        Me.dvJobRate.Table = Me.dsJobRate.JobRate
        '
        'adpJobRate
        '
        Me.adpJobRate.DeleteCommand = Me.OleDbDeleteCommand1
        Me.adpJobRate.InsertCommand = Me.OleDbInsertCommand1
        Me.adpJobRate.SelectCommand = Me.OleDbSelectCommand1
        Me.adpJobRate.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "JobRate", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("JobNo", "JobNo"), New System.Data.Common.DataColumnMapping("Rate", "Rate"), New System.Data.Common.DataColumnMapping("TypeCode", "TypeCode")})})
        Me.adpJobRate.UpdateCommand = Me.OleDbUpdateCommand1
        '
        'OleDbDeleteCommand1
        '
        Me.OleDbDeleteCommand1.CommandText = "DELETE FROM JobRate WHERE (JobNo = ?) AND (TypeCode = ?) AND (Rate = ? OR ? IS NU" & _
        "LL AND Rate IS NULL)"
        Me.OleDbDeleteCommand1.Connection = Me.conJobRate
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_JobNo", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "JobNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_TypeCode", System.Data.OleDb.OleDbType.VarWChar, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "TypeCode", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Rate", System.Data.OleDb.OleDbType.Currency, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Rate", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Rate1", System.Data.OleDb.OleDbType.Currency, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Rate", System.Data.DataRowVersion.Original, Nothing))
        '
        'conJobRate
        '
        Me.conJobRate.ConnectionString = "Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Registry Path=;Jet OLEDB:Database L" & _
        "ocking Mode=1;Data Source=""winsteelVers5.mdb"";Mode=Share Deny None;Jet OLEDB:Eng" & _
        "ine Type=5;Provider=""Microsoft.Jet.OLEDB.4.0"";Jet OLEDB:System database=;Jet OLE" & _
        "DB:SFP=False;persist security info=False;Extended Properties=;Jet OLEDB:Compact " & _
        "Without Replica Repair=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Create S" & _
        "ystem Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;User ID=Admin;" & _
        "Jet OLEDB:Global Bulk Transactions=1"
        '
        'OleDbInsertCommand1
        '
        Me.OleDbInsertCommand1.CommandText = "INSERT INTO JobRate(JobNo, Rate, TypeCode) VALUES (?, ?, ?)"
        Me.OleDbInsertCommand1.Connection = Me.conJobRate
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("JobNo", System.Data.OleDb.OleDbType.VarWChar, 50, "JobNo"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Rate", System.Data.OleDb.OleDbType.Currency, 0, "Rate"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("TypeCode", System.Data.OleDb.OleDbType.VarWChar, 8, "TypeCode"))
        '
        'OleDbSelectCommand1
        '
        Me.OleDbSelectCommand1.CommandText = "SELECT JobNo, Rate, TypeCode FROM JobRate"
        Me.OleDbSelectCommand1.Connection = Me.conJobRate
        '
        'OleDbUpdateCommand1
        '
        Me.OleDbUpdateCommand1.CommandText = "UPDATE JobRate SET JobNo = ?, Rate = ?, TypeCode = ? WHERE (JobNo = ?) AND (TypeC" & _
        "ode = ?) AND (Rate = ? OR ? IS NULL AND Rate IS NULL)"
        Me.OleDbUpdateCommand1.Connection = Me.conJobRate
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("JobNo", System.Data.OleDb.OleDbType.VarWChar, 50, "JobNo"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Rate", System.Data.OleDb.OleDbType.Currency, 0, "Rate"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("TypeCode", System.Data.OleDb.OleDbType.VarWChar, 8, "TypeCode"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_JobNo", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "JobNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_TypeCode", System.Data.OleDb.OleDbType.VarWChar, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "TypeCode", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Rate", System.Data.OleDb.OleDbType.Currency, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Rate", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Rate1", System.Data.OleDb.OleDbType.Currency, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Rate", System.Data.DataRowVersion.Original, Nothing))
        '
        'grdJobRate
        '
        Me.grdJobRate.CaptionVisible = False
        Me.grdJobRate.DataMember = ""
        Me.grdJobRate.DataSource = Me.dvJobRate
        Me.grdJobRate.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.grdJobRate.Location = New System.Drawing.Point(32, 56)
        Me.grdJobRate.Name = "grdJobRate"
        Me.grdJobRate.Size = New System.Drawing.Size(360, 432)
        Me.grdJobRate.TabIndex = 4
        Me.grdJobRate.TableStyles.AddRange(New System.Windows.Forms.DataGridTableStyle() {Me.DataGridTableStyle1})
        '
        'DataGridTableStyle1
        '
        Me.DataGridTableStyle1.DataGrid = Me.grdJobRate
        Me.DataGridTableStyle1.GridColumnStyles.AddRange(New System.Windows.Forms.DataGridColumnStyle() {Me.colType, Me.colRate})
        Me.DataGridTableStyle1.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGridTableStyle1.MappingName = "JobRate"
        '
        'colType
        '
        Me.colType.Format = ""
        Me.colType.FormatInfo = Nothing
        Me.colType.HeaderText = "Type"
        Me.colType.MappingName = "TypeCode"
        Me.colType.NullText = ""
        Me.colType.ReadOnly = True
        Me.colType.Width = 150
        '
        'colRate
        '
        Me.colRate.Format = "#####0.00"
        Me.colRate.FormatInfo = Nothing
        Me.colRate.HeaderText = "Rate"
        Me.colRate.MappingName = "Rate"
        Me.colRate.NullText = ""
        Me.colRate.Width = 150
        '
        'adpJob
        '
        Me.adpJob.DeleteCommand = Me.OleDbDeleteCommand2
        Me.adpJob.InsertCommand = Me.OleDbInsertCommand2
        Me.adpJob.SelectCommand = Me.OleDbSelectCommand2
        Me.adpJob.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "Job", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("AddedDiscount", "AddedDiscount"), New System.Data.Common.DataColumnMapping("CompanyNo", "CompanyNo"), New System.Data.Common.DataColumnMapping("ContractorNo", "ContractorNo"), New System.Data.Common.DataColumnMapping("Design", "Design"), New System.Data.Common.DataColumnMapping("Discount", "Discount"), New System.Data.Common.DataColumnMapping("JobName", "JobName"), New System.Data.Common.DataColumnMapping("JobNo", "JobNo"), New System.Data.Common.DataColumnMapping("OrderNo", "OrderNo"), New System.Data.Common.DataColumnMapping("Tons Or Kilograms", "Tons Or Kilograms"), New System.Data.Common.DataColumnMapping("UnitOfMeas", "UnitOfMeas")})})
        Me.adpJob.UpdateCommand = Me.OleDbUpdateCommand2
        '
        'OleDbDeleteCommand2
        '
        Me.OleDbDeleteCommand2.CommandText = "DELETE FROM Job WHERE (JobNo = ?) AND (AddedDiscount = ? OR ? IS NULL AND AddedDi" & _
        "scount IS NULL) AND (CompanyNo = ? OR ? IS NULL AND CompanyNo IS NULL) AND (Cont" & _
        "ractorNo = ? OR ? IS NULL AND ContractorNo IS NULL) AND (Design = ? OR ? IS NULL" & _
        " AND Design IS NULL) AND (Discount = ? OR ? IS NULL AND Discount IS NULL) AND (J" & _
        "obName = ? OR ? IS NULL AND JobName IS NULL) AND (OrderNo = ? OR ? IS NULL AND O" & _
        "rderNo IS NULL) AND ([Tons Or Kilograms] = ? OR ? IS NULL AND [Tons Or Kilograms" & _
        "] IS NULL) AND (UnitOfMeas = ? OR ? IS NULL AND UnitOfMeas IS NULL)"
        Me.OleDbDeleteCommand2.Connection = Me.conJobRate
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_JobNo", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "JobNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddedDiscount", System.Data.OleDb.OleDbType.Currency, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddedDiscount", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddedDiscount1", System.Data.OleDb.OleDbType.Currency, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddedDiscount", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_CompanyNo", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CompanyNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_CompanyNo1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CompanyNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ContractorNo", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ContractorNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ContractorNo1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ContractorNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Design", System.Data.OleDb.OleDbType.Currency, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Design", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Design1", System.Data.OleDb.OleDbType.Currency, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Design", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Discount", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Discount", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Discount1", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Discount", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_JobName", System.Data.OleDb.OleDbType.VarWChar, 70, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "JobName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_JobName1", System.Data.OleDb.OleDbType.VarWChar, 70, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "JobName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_OrderNo", System.Data.OleDb.OleDbType.VarWChar, 25, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "OrderNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_OrderNo1", System.Data.OleDb.OleDbType.VarWChar, 25, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "OrderNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Tons_Or_Kilograms", System.Data.OleDb.OleDbType.VarWChar, 10, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Tons Or Kilograms", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Tons_Or_Kilograms1", System.Data.OleDb.OleDbType.VarWChar, 10, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Tons Or Kilograms", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_UnitOfMeas", System.Data.OleDb.OleDbType.VarWChar, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "UnitOfMeas", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_UnitOfMeas1", System.Data.OleDb.OleDbType.VarWChar, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "UnitOfMeas", System.Data.DataRowVersion.Original, Nothing))
        '
        'OleDbInsertCommand2
        '
        Me.OleDbInsertCommand2.CommandText = "INSERT INTO Job(AddedDiscount, CompanyNo, ContractorNo, Design, Discount, JobName" & _
        ", JobNo, OrderNo, [Tons Or Kilograms], UnitOfMeas) VALUES (?, ?, ?, ?, ?, ?, ?, " & _
        "?, ?, ?)"
        Me.OleDbInsertCommand2.Connection = Me.conJobRate
        Me.OleDbInsertCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("AddedDiscount", System.Data.OleDb.OleDbType.Currency, 0, "AddedDiscount"))
        Me.OleDbInsertCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("CompanyNo", System.Data.OleDb.OleDbType.VarWChar, 50, "CompanyNo"))
        Me.OleDbInsertCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("ContractorNo", System.Data.OleDb.OleDbType.VarWChar, 50, "ContractorNo"))
        Me.OleDbInsertCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Design", System.Data.OleDb.OleDbType.Currency, 0, "Design"))
        Me.OleDbInsertCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Discount", System.Data.OleDb.OleDbType.Integer, 0, "Discount"))
        Me.OleDbInsertCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("JobName", System.Data.OleDb.OleDbType.VarWChar, 70, "JobName"))
        Me.OleDbInsertCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("JobNo", System.Data.OleDb.OleDbType.VarWChar, 50, "JobNo"))
        Me.OleDbInsertCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("OrderNo", System.Data.OleDb.OleDbType.VarWChar, 25, "OrderNo"))
        Me.OleDbInsertCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Tons_Or_Kilograms", System.Data.OleDb.OleDbType.VarWChar, 10, "Tons Or Kilograms"))
        Me.OleDbInsertCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("UnitOfMeas", System.Data.OleDb.OleDbType.VarWChar, 8, "UnitOfMeas"))
        '
        'OleDbSelectCommand2
        '
        Me.OleDbSelectCommand2.CommandText = "SELECT AddedDiscount, CompanyNo, ContractorNo, Design, Discount, JobName, JobNo, " & _
        "OrderNo, [Tons Or Kilograms], UnitOfMeas FROM Job"
        Me.OleDbSelectCommand2.Connection = Me.conJobRate
        '
        'OleDbUpdateCommand2
        '
        Me.OleDbUpdateCommand2.CommandText = "UPDATE Job SET AddedDiscount = ?, CompanyNo = ?, ContractorNo = ?, Design = ?, Di" & _
        "scount = ?, JobName = ?, JobNo = ?, OrderNo = ?, [Tons Or Kilograms] = ?, UnitOf" & _
        "Meas = ? WHERE (JobNo = ?) AND (AddedDiscount = ? OR ? IS NULL AND AddedDiscount" & _
        " IS NULL) AND (CompanyNo = ? OR ? IS NULL AND CompanyNo IS NULL) AND (Contractor" & _
        "No = ? OR ? IS NULL AND ContractorNo IS NULL) AND (Design = ? OR ? IS NULL AND D" & _
        "esign IS NULL) AND (Discount = ? OR ? IS NULL AND Discount IS NULL) AND (JobName" & _
        " = ? OR ? IS NULL AND JobName IS NULL) AND (OrderNo = ? OR ? IS NULL AND OrderNo" & _
        " IS NULL) AND ([Tons Or Kilograms] = ? OR ? IS NULL AND [Tons Or Kilograms] IS N" & _
        "ULL) AND (UnitOfMeas = ? OR ? IS NULL AND UnitOfMeas IS NULL)"
        Me.OleDbUpdateCommand2.Connection = Me.conJobRate
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("AddedDiscount", System.Data.OleDb.OleDbType.Currency, 0, "AddedDiscount"))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("CompanyNo", System.Data.OleDb.OleDbType.VarWChar, 50, "CompanyNo"))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("ContractorNo", System.Data.OleDb.OleDbType.VarWChar, 50, "ContractorNo"))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Design", System.Data.OleDb.OleDbType.Currency, 0, "Design"))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Discount", System.Data.OleDb.OleDbType.Integer, 0, "Discount"))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("JobName", System.Data.OleDb.OleDbType.VarWChar, 70, "JobName"))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("JobNo", System.Data.OleDb.OleDbType.VarWChar, 50, "JobNo"))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("OrderNo", System.Data.OleDb.OleDbType.VarWChar, 25, "OrderNo"))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Tons_Or_Kilograms", System.Data.OleDb.OleDbType.VarWChar, 10, "Tons Or Kilograms"))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("UnitOfMeas", System.Data.OleDb.OleDbType.VarWChar, 8, "UnitOfMeas"))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_JobNo", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "JobNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddedDiscount", System.Data.OleDb.OleDbType.Currency, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddedDiscount", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddedDiscount1", System.Data.OleDb.OleDbType.Currency, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddedDiscount", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_CompanyNo", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CompanyNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_CompanyNo1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CompanyNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ContractorNo", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ContractorNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ContractorNo1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ContractorNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Design", System.Data.OleDb.OleDbType.Currency, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Design", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Design1", System.Data.OleDb.OleDbType.Currency, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Design", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Discount", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Discount", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Discount1", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Discount", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_JobName", System.Data.OleDb.OleDbType.VarWChar, 70, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "JobName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_JobName1", System.Data.OleDb.OleDbType.VarWChar, 70, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "JobName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_OrderNo", System.Data.OleDb.OleDbType.VarWChar, 25, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "OrderNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_OrderNo1", System.Data.OleDb.OleDbType.VarWChar, 25, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "OrderNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Tons_Or_Kilograms", System.Data.OleDb.OleDbType.VarWChar, 10, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Tons Or Kilograms", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Tons_Or_Kilograms1", System.Data.OleDb.OleDbType.VarWChar, 10, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Tons Or Kilograms", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_UnitOfMeas", System.Data.OleDb.OleDbType.VarWChar, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "UnitOfMeas", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_UnitOfMeas1", System.Data.OleDb.OleDbType.VarWChar, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "UnitOfMeas", System.Data.DataRowVersion.Original, Nothing))
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(320, 504)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.TabIndex = 6
        Me.btnClose.Text = "Close"
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(232, 504)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.TabIndex = 5
        Me.btnSave.Text = "Save"
        '
        'cmdTypeExist
        '
        Me.cmdTypeExist.CommandText = "SELECT * FROM ProductType WHERE (TypeCode = ?)"
        Me.cmdTypeExist.Connection = Me.conJobRate
        Me.cmdTypeExist.Parameters.Add(New System.Data.OleDb.OleDbParameter("TypeCode", System.Data.OleDb.OleDbType.VarWChar, 8, "TypeCode"))
        '
        'cbxJobNo
        '
        Me.cbxJobNo.DataSource = Me.dvJob
        Me.cbxJobNo.DisplayMember = "No&Name"
        Me.cbxJobNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxJobNo.Location = New System.Drawing.Point(80, 16)
        Me.cbxJobNo.Name = "cbxJobNo"
        Me.cbxJobNo.Size = New System.Drawing.Size(312, 21)
        Me.cbxJobNo.TabIndex = 1
        Me.cbxJobNo.ValueMember = "JobNo"
        '
        'frmJobRate
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(418, 542)
        Me.Controls.Add(Me.cbxJobNo)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.grdJobRate)
        Me.Controls.Add(Me.lblJobNo)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "frmJobRate"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Job Rate Maintenance"
        CType(Me.dvJob, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dsJobRate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dvJobRate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdJobRate, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub frmJobRate_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        dsJobRate.Clear()
        adpJob.Fill(dsJobRate.Job)
        adpJobRate.Fill(dsJobRate.JobRate)

        dvJobRate.RowFilter = "JobNo='" + cbxJobNo.SelectedValue + "'"
        dsJobRate.JobRate.Columns("JobNo").DefaultValue = cbxJobNo.SelectedValue

        table = dsJobRate.JobRate
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Close()
    End Sub

    Dim WithEvents table As DataTable

    Private Sub table_ColumnChanging(ByVal sender As Object, ByVal e As System.Data.DataColumnChangeEventArgs) Handles table.ColumnChanging
        If e.Column.ColumnName = "TypeCode" Then
            If Not IsDBNull(e.ProposedValue) Then

                Dim DataReader As System.Data.OleDb.OleDbDataReader
                Dim count As Integer = 0

                conJobRate.Open()
                cmdTypeExist.Parameters("TypeCode").Value = e.ProposedValue
                DataReader = cmdTypeExist.ExecuteReader(CommandBehavior.CloseConnection)
                While DataReader.Read()
                    count += 1
                End While
                DataReader.Close()
                conJobRate.Close()

                If count < 1 Then
                    MsgBox("Type does not exist", MsgBoxStyle.Critical, "Error")
                    Throw New ArgumentException("Type does no exist")

                    'grdJobRate.CurrentCell = e.Row.Item("TypeCode")
                End If
            End If
        End If
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        adpJobRate.Update(dsJobRate.JobRate)
        MsgBox("Changes were successfully saved", MsgBoxStyle.Information, "Information")
    End Sub

    Private CallingForm As Object

    Public Sub New(ByVal caller As Object)
        MyBase.New()
        InitializeComponent()

        CallingForm = caller
    End Sub

    Private Sub frmAddClient_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If Not IsNothing(CallingForm) Then
            CallingForm.Show()
        End If

        CallingForm = Nothing
    End Sub

    Private Sub cbxJobNo_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxJobNo.SelectedValueChanged
        dvJobRate.RowFilter = "JobNo='" + cbxJobNo.SelectedValue + "'"
        dsJobRate.JobRate.Columns("JobNo").DefaultValue = cbxJobNo.SelectedValue
    End Sub

    Private Sub adpJobRate_RowUpdated(ByVal sender As System.Object, ByVal e As System.Data.OleDb.OleDbRowUpdatedEventArgs) Handles adpJobRate.RowUpdated

    End Sub
End Class
