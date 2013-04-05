Imports System.Data.OleDb
Public Class frmNewCutSheet
    Inherits System.Windows.Forms.Form
    Shared cnnJob As New _
            OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; " & _
            "Data Source=winsteelVers5.mdb")
    
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
    Friend WithEvents conCutSheet As System.Data.OleDb.OleDbConnection
    Friend WithEvents lblCutSheetNo As System.Windows.Forms.Label
    Friend WithEvents adpCutSheet As System.Data.OleDb.OleDbDataAdapter
    Friend WithEvents lblDate As System.Windows.Forms.Label
    Friend WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblDetails As System.Windows.Forms.Label
    Friend WithEvents grpCutSheetDetails As System.Windows.Forms.GroupBox
    Friend WithEvents grpSchedules As System.Windows.Forms.GroupBox
    Friend WithEvents dvSchedules As System.Data.DataView
    Friend WithEvents OleDbSelectCommand3 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbInsertCommand3 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbUpdateCommand3 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbDeleteCommand3 As System.Data.OleDb.OleDbCommand
    Friend WithEvents adpItems As System.Data.OleDb.OleDbDataAdapter
    Friend WithEvents grpItems As System.Windows.Forms.GroupBox
    Friend WithEvents grdItems As System.Windows.Forms.DataGrid
    Friend WithEvents dvItems As System.Data.DataView
    Friend WithEvents colType As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents colLength As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents colQty As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents adpSchedules As System.Data.OleDb.OleDbDataAdapter
    Friend WithEvents OleDbSelectCommand2 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbInsertCommand2 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbUpdateCommand2 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbDeleteCommand2 As System.Data.OleDb.OleDbCommand
    Friend WithEvents styleItems As System.Windows.Forms.DataGridTableStyle
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents cbxCutSheetNo As System.Windows.Forms.ComboBox
    Friend WithEvents txtDetails As System.Windows.Forms.TextBox
    Friend WithEvents dvJob As System.Data.DataView
    Friend WithEvents OleDbInsertCommand4 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbUpdateCommand4 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbDeleteCommand4 As System.Data.OleDb.OleDbCommand
    Friend WithEvents cmdTypeExist As System.Data.OleDb.OleDbCommand
    Friend WithEvents dvCutSheet As System.Data.DataView
    Friend WithEvents txtCutSheetNo As System.Windows.Forms.TextBox
    Friend WithEvents OleDbSelectCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbInsertCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbUpdateCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbDeleteCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents dsCutSheet As Reinforcing_Ability.dsReinforcingAbility
    Friend WithEvents btnSaveAll As System.Windows.Forms.Button
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents lblSchedNo As System.Windows.Forms.Label
    Friend WithEvents cbxSchedNo As System.Windows.Forms.ComboBox
    Friend WithEvents txtSchedNo As System.Windows.Forms.TextBox
    Friend WithEvents btnAddSched As System.Windows.Forms.Button
    Friend WithEvents btnDelSched As System.Windows.Forms.Button
    Friend WithEvents btnSaveSched As System.Windows.Forms.Button
    Friend WithEvents txtJobName As System.Windows.Forms.TextBox
    Friend WithEvents lblJobName As System.Windows.Forms.Label
    Friend WithEvents txtJob As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.lblCutSheetNo = New System.Windows.Forms.Label
        Me.lblJobNo = New System.Windows.Forms.Label
        Me.dvJob = New System.Data.DataView
        Me.dsCutSheet = New Reinforcing_Ability.dsReinforcingAbility
        Me.conCutSheet = New System.Data.OleDb.OleDbConnection
        Me.adpCutSheet = New System.Data.OleDb.OleDbDataAdapter
        Me.OleDbDeleteCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbInsertCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbSelectCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbUpdateCommand1 = New System.Data.OleDb.OleDbCommand
        Me.lblDate = New System.Windows.Forms.Label
        Me.dtpDate = New System.Windows.Forms.DateTimePicker
        Me.dvCutSheet = New System.Data.DataView
        Me.lblDetails = New System.Windows.Forms.Label
        Me.grpCutSheetDetails = New System.Windows.Forms.GroupBox
        Me.txtJob = New System.Windows.Forms.TextBox
        Me.lblJobName = New System.Windows.Forms.Label
        Me.txtJobName = New System.Windows.Forms.TextBox
        Me.txtDetails = New System.Windows.Forms.TextBox
        Me.btnAdd = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.txtCutSheetNo = New System.Windows.Forms.TextBox
        Me.cbxCutSheetNo = New System.Windows.Forms.ComboBox
        Me.grpSchedules = New System.Windows.Forms.GroupBox
        Me.txtSchedNo = New System.Windows.Forms.TextBox
        Me.dvSchedules = New System.Data.DataView
        Me.lblSchedNo = New System.Windows.Forms.Label
        Me.btnAddSched = New System.Windows.Forms.Button
        Me.btnSaveSched = New System.Windows.Forms.Button
        Me.btnDelSched = New System.Windows.Forms.Button
        Me.adpItems = New System.Data.OleDb.OleDbDataAdapter
        Me.OleDbDeleteCommand3 = New System.Data.OleDb.OleDbCommand
        Me.OleDbInsertCommand3 = New System.Data.OleDb.OleDbCommand
        Me.OleDbSelectCommand3 = New System.Data.OleDb.OleDbCommand
        Me.OleDbUpdateCommand3 = New System.Data.OleDb.OleDbCommand
        Me.grpItems = New System.Windows.Forms.GroupBox
        Me.grdItems = New System.Windows.Forms.DataGrid
        Me.dvItems = New System.Data.DataView
        Me.styleItems = New System.Windows.Forms.DataGridTableStyle
        Me.colType = New System.Windows.Forms.DataGridTextBoxColumn
        Me.colLength = New System.Windows.Forms.DataGridTextBoxColumn
        Me.colQty = New System.Windows.Forms.DataGridTextBoxColumn
        Me.btnSaveAll = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.adpSchedules = New System.Data.OleDb.OleDbDataAdapter
        Me.OleDbDeleteCommand2 = New System.Data.OleDb.OleDbCommand
        Me.OleDbInsertCommand2 = New System.Data.OleDb.OleDbCommand
        Me.OleDbSelectCommand2 = New System.Data.OleDb.OleDbCommand
        Me.OleDbUpdateCommand2 = New System.Data.OleDb.OleDbCommand
        Me.OleDbDeleteCommand4 = New System.Data.OleDb.OleDbCommand
        Me.OleDbInsertCommand4 = New System.Data.OleDb.OleDbCommand
        Me.OleDbUpdateCommand4 = New System.Data.OleDb.OleDbCommand
        Me.cmdTypeExist = New System.Data.OleDb.OleDbCommand
        Me.btnPrint = New System.Windows.Forms.Button
        Me.cbxSchedNo = New System.Windows.Forms.ComboBox
        CType(Me.dvJob, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dsCutSheet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dvCutSheet, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpCutSheetDetails.SuspendLayout()
        Me.grpSchedules.SuspendLayout()
        CType(Me.dvSchedules, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpItems.SuspendLayout()
        CType(Me.grdItems, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dvItems, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblCutSheetNo
        '
        Me.lblCutSheetNo.Location = New System.Drawing.Point(408, 32)
        Me.lblCutSheetNo.Name = "lblCutSheetNo"
        Me.lblCutSheetNo.Size = New System.Drawing.Size(96, 23)
        Me.lblCutSheetNo.TabIndex = 2
        Me.lblCutSheetNo.Text = "Cutting Sheet No."
        Me.lblCutSheetNo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblJobNo
        '
        Me.lblJobNo.Location = New System.Drawing.Point(16, 32)
        Me.lblJobNo.Name = "lblJobNo"
        Me.lblJobNo.Size = New System.Drawing.Size(48, 23)
        Me.lblJobNo.TabIndex = 0
        Me.lblJobNo.Text = "Job No."
        Me.lblJobNo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'dvJob
        '
        Me.dvJob.Table = Me.dsCutSheet.Job
        '
        'dsCutSheet
        '
        Me.dsCutSheet.DataSetName = "dsReinforcingAbility"
        Me.dsCutSheet.Locale = New System.Globalization.CultureInfo("en-ZA")
        '
        'conCutSheet
        '
        Me.conCutSheet.ConnectionString = "Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Registry Path=;Jet OLEDB:Database L" & _
        "ocking Mode=1;Data Source=""winsteelVers5.mdb"";Mode=Share Deny None;Jet OLEDB:Eng" & _
        "ine Type=5;Provider=""Microsoft.Jet.OLEDB.4.0"";Jet OLEDB:System database=;Jet OLE" & _
        "DB:SFP=False;persist security info=False;Extended Properties=;Jet OLEDB:Compact " & _
        "Without Replica Repair=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Create S" & _
        "ystem Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;User ID=Admin;" & _
        "Jet OLEDB:Global Bulk Transactions=1"
        '
        'adpCutSheet
        '
        Me.adpCutSheet.DeleteCommand = Me.OleDbDeleteCommand1
        Me.adpCutSheet.InsertCommand = Me.OleDbInsertCommand1
        Me.adpCutSheet.SelectCommand = Me.OleDbSelectCommand1
        Me.adpCutSheet.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "CuttingSheet", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("CSHeading", "CSHeading"), New System.Data.Common.DataColumnMapping("CutDate", "CutDate"), New System.Data.Common.DataColumnMapping("CutSheetNo", "CutSheetNo"), New System.Data.Common.DataColumnMapping("Details", "Details"), New System.Data.Common.DataColumnMapping("InvoiceNo", "InvoiceNo"), New System.Data.Common.DataColumnMapping("Job No", "Job No")})})
        Me.adpCutSheet.UpdateCommand = Me.OleDbUpdateCommand1
        '
        'OleDbDeleteCommand1
        '
        Me.OleDbDeleteCommand1.CommandText = "DELETE FROM CuttingSheet WHERE (CutSheetNo = ?) AND (CSHeading = ? OR ? IS NULL A" & _
        "ND CSHeading IS NULL) AND (CutDate = ? OR ? IS NULL AND CutDate IS NULL) AND (De" & _
        "tails = ? OR ? IS NULL AND Details IS NULL) AND (InvoiceNo = ? OR ? IS NULL AND " & _
        "InvoiceNo IS NULL) AND ([Job No] = ? OR ? IS NULL AND [Job No] IS NULL)"
        Me.OleDbDeleteCommand1.Connection = Me.conCutSheet
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_CutSheetNo", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CutSheetNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_CSHeading", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CSHeading", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_CSHeading1", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CSHeading", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_CutDate", System.Data.OleDb.OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CutDate", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_CutDate1", System.Data.OleDb.OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CutDate", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Details", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Details", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Details1", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Details", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_InvoiceNo", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "InvoiceNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_InvoiceNo1", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "InvoiceNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Job_No", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Job No", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Job_No1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Job No", System.Data.DataRowVersion.Original, Nothing))
        '
        'OleDbInsertCommand1
        '
        Me.OleDbInsertCommand1.CommandText = "INSERT INTO CuttingSheet(CSHeading, CutDate, CutSheetNo, Details, InvoiceNo, [Job" & _
        " No]) VALUES (?, ?, ?, ?, ?, ?)"
        Me.OleDbInsertCommand1.Connection = Me.conCutSheet
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("CSHeading", System.Data.OleDb.OleDbType.VarWChar, 40, "CSHeading"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("CutDate", System.Data.OleDb.OleDbType.DBDate, 0, "CutDate"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("CutSheetNo", System.Data.OleDb.OleDbType.Integer, 0, "CutSheetNo"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Details", System.Data.OleDb.OleDbType.VarWChar, 40, "Details"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("InvoiceNo", System.Data.OleDb.OleDbType.Integer, 0, "InvoiceNo"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Job_No", System.Data.OleDb.OleDbType.VarWChar, 50, "Job No"))
        '
        'OleDbSelectCommand1
        '
        Me.OleDbSelectCommand1.CommandText = "SELECT CSHeading, CutDate, CutSheetNo, Details, InvoiceNo, [Job No] FROM CuttingS" & _
        "heet"
        Me.OleDbSelectCommand1.Connection = Me.conCutSheet
        '
        'OleDbUpdateCommand1
        '
        Me.OleDbUpdateCommand1.CommandText = "UPDATE CuttingSheet SET CSHeading = ?, CutDate = ?, CutSheetNo = ?, Details = ?, " & _
        "InvoiceNo = ?, [Job No] = ? WHERE (CutSheetNo = ?) AND (CSHeading = ? OR ? IS NU" & _
        "LL AND CSHeading IS NULL) AND (CutDate = ? OR ? IS NULL AND CutDate IS NULL) AND" & _
        " (Details = ? OR ? IS NULL AND Details IS NULL) AND (InvoiceNo = ? OR ? IS NULL " & _
        "AND InvoiceNo IS NULL) AND ([Job No] = ? OR ? IS NULL AND [Job No] IS NULL)"
        Me.OleDbUpdateCommand1.Connection = Me.conCutSheet
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("CSHeading", System.Data.OleDb.OleDbType.VarWChar, 40, "CSHeading"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("CutDate", System.Data.OleDb.OleDbType.DBDate, 0, "CutDate"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("CutSheetNo", System.Data.OleDb.OleDbType.Integer, 0, "CutSheetNo"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Details", System.Data.OleDb.OleDbType.VarWChar, 40, "Details"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("InvoiceNo", System.Data.OleDb.OleDbType.Integer, 0, "InvoiceNo"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Job_No", System.Data.OleDb.OleDbType.VarWChar, 50, "Job No"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_CutSheetNo", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CutSheetNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_CSHeading", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CSHeading", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_CSHeading1", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CSHeading", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_CutDate", System.Data.OleDb.OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CutDate", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_CutDate1", System.Data.OleDb.OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CutDate", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Details", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Details", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Details1", System.Data.OleDb.OleDbType.VarWChar, 40, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Details", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_InvoiceNo", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "InvoiceNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_InvoiceNo1", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "InvoiceNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Job_No", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Job No", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Job_No1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Job No", System.Data.DataRowVersion.Original, Nothing))
        '
        'lblDate
        '
        Me.lblDate.Location = New System.Drawing.Point(464, 64)
        Me.lblDate.Name = "lblDate"
        Me.lblDate.Size = New System.Drawing.Size(32, 23)
        Me.lblDate.TabIndex = 5
        Me.lblDate.Text = "Date"
        Me.lblDate.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'dtpDate
        '
        Me.dtpDate.CustomFormat = ""
        Me.dtpDate.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.dvCutSheet, "CutDate"))
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpDate.Location = New System.Drawing.Point(512, 64)
        Me.dtpDate.MinDate = New Date(2003, 1, 1, 0, 0, 0, 0)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(104, 20)
        Me.dtpDate.TabIndex = 6
        Me.dtpDate.Value = New Date(2004, 3, 9, 2, 3, 4, 525)
        '
        'dvCutSheet
        '
        Me.dvCutSheet.Table = Me.dsCutSheet.CuttingSheet
        '
        'lblDetails
        '
        Me.lblDetails.Location = New System.Drawing.Point(16, 93)
        Me.lblDetails.Name = "lblDetails"
        Me.lblDetails.Size = New System.Drawing.Size(40, 23)
        Me.lblDetails.TabIndex = 7
        Me.lblDetails.Text = "Details"
        Me.lblDetails.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'grpCutSheetDetails
        '
        Me.grpCutSheetDetails.Controls.Add(Me.txtJob)
        Me.grpCutSheetDetails.Controls.Add(Me.lblJobName)
        Me.grpCutSheetDetails.Controls.Add(Me.txtJobName)
        Me.grpCutSheetDetails.Controls.Add(Me.txtDetails)
        Me.grpCutSheetDetails.Controls.Add(Me.lblJobNo)
        Me.grpCutSheetDetails.Controls.Add(Me.dtpDate)
        Me.grpCutSheetDetails.Controls.Add(Me.lblCutSheetNo)
        Me.grpCutSheetDetails.Controls.Add(Me.lblDetails)
        Me.grpCutSheetDetails.Controls.Add(Me.lblDate)
        Me.grpCutSheetDetails.Controls.Add(Me.btnAdd)
        Me.grpCutSheetDetails.Controls.Add(Me.btnSave)
        Me.grpCutSheetDetails.Controls.Add(Me.txtCutSheetNo)
        Me.grpCutSheetDetails.Location = New System.Drawing.Point(16, 8)
        Me.grpCutSheetDetails.Name = "grpCutSheetDetails"
        Me.grpCutSheetDetails.Size = New System.Drawing.Size(640, 168)
        Me.grpCutSheetDetails.TabIndex = 0
        Me.grpCutSheetDetails.TabStop = False
        Me.grpCutSheetDetails.Text = "Cutting Sheet Details"
        '
        'txtJob
        '
        Me.txtJob.Location = New System.Drawing.Point(80, 32)
        Me.txtJob.Name = "txtJob"
        Me.txtJob.TabIndex = 14
        Me.txtJob.Text = ""
        '
        'lblJobName
        '
        Me.lblJobName.Location = New System.Drawing.Point(16, 64)
        Me.lblJobName.Name = "lblJobName"
        Me.lblJobName.Size = New System.Drawing.Size(56, 23)
        Me.lblJobName.TabIndex = 13
        Me.lblJobName.Text = "Job Name"
        Me.lblJobName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtJobName
        '
        Me.txtJobName.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.dvJob, "JobName"))
        Me.txtJobName.Enabled = False
        Me.txtJobName.Location = New System.Drawing.Point(80, 64)
        Me.txtJobName.Name = "txtJobName"
        Me.txtJobName.Size = New System.Drawing.Size(312, 20)
        Me.txtJobName.TabIndex = 12
        Me.txtJobName.Text = ""
        '
        'txtDetails
        '
        Me.txtDetails.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.dvCutSheet, "Details"))
        Me.txtDetails.Location = New System.Drawing.Point(80, 96)
        Me.txtDetails.Name = "txtDetails"
        Me.txtDetails.Size = New System.Drawing.Size(536, 20)
        Me.txtDetails.TabIndex = 8
        Me.txtDetails.Text = ""
        '
        'btnAdd
        '
        Me.btnAdd.Location = New System.Drawing.Point(16, 128)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.TabIndex = 9
        Me.btnAdd.Text = "Add"
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(112, 128)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.TabIndex = 11
        Me.btnSave.Text = "Save"
        '
        'txtCutSheetNo
        '
        Me.txtCutSheetNo.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.dvCutSheet, "CutSheetNo"))
        Me.txtCutSheetNo.Enabled = False
        Me.txtCutSheetNo.Location = New System.Drawing.Point(512, 32)
        Me.txtCutSheetNo.Name = "txtCutSheetNo"
        Me.txtCutSheetNo.Size = New System.Drawing.Size(104, 20)
        Me.txtCutSheetNo.TabIndex = 10
        Me.txtCutSheetNo.Text = ""
        '
        'cbxCutSheetNo
        '
        Me.cbxCutSheetNo.DataSource = Me.dvCutSheet
        Me.cbxCutSheetNo.DisplayMember = "CutSheetNo"
        Me.cbxCutSheetNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxCutSheetNo.Location = New System.Drawing.Point(528, 40)
        Me.cbxCutSheetNo.Name = "cbxCutSheetNo"
        Me.cbxCutSheetNo.Size = New System.Drawing.Size(104, 21)
        Me.cbxCutSheetNo.TabIndex = 1
        Me.cbxCutSheetNo.ValueMember = "CutSheetNo"
        '
        'grpSchedules
        '
        Me.grpSchedules.Controls.Add(Me.txtSchedNo)
        Me.grpSchedules.Controls.Add(Me.lblSchedNo)
        Me.grpSchedules.Controls.Add(Me.btnAddSched)
        Me.grpSchedules.Controls.Add(Me.btnSaveSched)
        Me.grpSchedules.Controls.Add(Me.btnDelSched)
        Me.grpSchedules.Location = New System.Drawing.Point(16, 184)
        Me.grpSchedules.Name = "grpSchedules"
        Me.grpSchedules.Size = New System.Drawing.Size(208, 160)
        Me.grpSchedules.TabIndex = 2
        Me.grpSchedules.TabStop = False
        Me.grpSchedules.Text = "Schedules"
        '
        'txtSchedNo
        '
        Me.txtSchedNo.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.dvSchedules, "ScheduleNo"))
        Me.txtSchedNo.Enabled = False
        Me.txtSchedNo.Location = New System.Drawing.Point(104, 32)
        Me.txtSchedNo.MaxLength = 12
        Me.txtSchedNo.Name = "txtSchedNo"
        Me.txtSchedNo.Size = New System.Drawing.Size(88, 20)
        Me.txtSchedNo.TabIndex = 5
        Me.txtSchedNo.Text = ""
        '
        'dvSchedules
        '
        Me.dvSchedules.Table = Me.dsCutSheet.SchedItem
        '
        'lblSchedNo
        '
        Me.lblSchedNo.Location = New System.Drawing.Point(16, 32)
        Me.lblSchedNo.Name = "lblSchedNo"
        Me.lblSchedNo.Size = New System.Drawing.Size(72, 23)
        Me.lblSchedNo.TabIndex = 3
        Me.lblSchedNo.Text = "Schedule No."
        Me.lblSchedNo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnAddSched
        '
        Me.btnAddSched.Location = New System.Drawing.Point(16, 72)
        Me.btnAddSched.Name = "btnAddSched"
        Me.btnAddSched.TabIndex = 10
        Me.btnAddSched.Text = "Add"
        '
        'btnSaveSched
        '
        Me.btnSaveSched.Location = New System.Drawing.Point(112, 72)
        Me.btnSaveSched.Name = "btnSaveSched"
        Me.btnSaveSched.TabIndex = 14
        Me.btnSaveSched.Text = "Save"
        '
        'btnDelSched
        '
        Me.btnDelSched.Location = New System.Drawing.Point(16, 120)
        Me.btnDelSched.Name = "btnDelSched"
        Me.btnDelSched.TabIndex = 15
        Me.btnDelSched.Text = "Delete"
        '
        'adpItems
        '
        Me.adpItems.DeleteCommand = Me.OleDbDeleteCommand3
        Me.adpItems.InsertCommand = Me.OleDbInsertCommand3
        Me.adpItems.SelectCommand = Me.OleDbSelectCommand3
        Me.adpItems.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "CutItem", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("CutSheetNo", "CutSheetNo"), New System.Data.Common.DataColumnMapping("Item", "Item"), New System.Data.Common.DataColumnMapping("Length", "Length"), New System.Data.Common.DataColumnMapping("Qty", "Qty"), New System.Data.Common.DataColumnMapping("ScheduleNo", "ScheduleNo"), New System.Data.Common.DataColumnMapping("TypeCode", "TypeCode")})})
        Me.adpItems.UpdateCommand = Me.OleDbUpdateCommand3
        '
        'OleDbDeleteCommand3
        '
        Me.OleDbDeleteCommand3.CommandText = "DELETE FROM CutItem WHERE (CutSheetNo = ?) AND (Item = ?) AND (ScheduleNo = ?) AN" & _
        "D (Length = ? OR ? IS NULL AND Length IS NULL) AND (Qty = ? OR ? IS NULL AND Qty" & _
        " IS NULL) AND (TypeCode = ? OR ? IS NULL AND TypeCode IS NULL)"
        Me.OleDbDeleteCommand3.Connection = Me.conCutSheet
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_CutSheetNo", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CutSheetNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Item", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Item", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ScheduleNo", System.Data.OleDb.OleDbType.VarWChar, 12, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ScheduleNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Length", System.Data.OleDb.OleDbType.Double, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Length", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Length1", System.Data.OleDb.OleDbType.Double, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Length", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Qty", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Qty", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Qty1", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Qty", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_TypeCode", System.Data.OleDb.OleDbType.VarWChar, 5, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "TypeCode", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_TypeCode1", System.Data.OleDb.OleDbType.VarWChar, 5, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "TypeCode", System.Data.DataRowVersion.Original, Nothing))
        '
        'OleDbInsertCommand3
        '
        Me.OleDbInsertCommand3.CommandText = "INSERT INTO CutItem(CutSheetNo, Item, Length, Qty, ScheduleNo, TypeCode) VALUES (" & _
        "?, ?, ?, ?, ?, ?)"
        Me.OleDbInsertCommand3.Connection = Me.conCutSheet
        Me.OleDbInsertCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("CutSheetNo", System.Data.OleDb.OleDbType.Integer, 0, "CutSheetNo"))
        Me.OleDbInsertCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Item", System.Data.OleDb.OleDbType.Integer, 0, "Item"))
        Me.OleDbInsertCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Length", System.Data.OleDb.OleDbType.Double, 0, "Length"))
        Me.OleDbInsertCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Qty", System.Data.OleDb.OleDbType.Integer, 0, "Qty"))
        Me.OleDbInsertCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("ScheduleNo", System.Data.OleDb.OleDbType.VarWChar, 12, "ScheduleNo"))
        Me.OleDbInsertCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("TypeCode", System.Data.OleDb.OleDbType.VarWChar, 5, "TypeCode"))
        '
        'OleDbSelectCommand3
        '
        Me.OleDbSelectCommand3.CommandText = "SELECT CutSheetNo, Item, Length, Qty, ScheduleNo, TypeCode FROM CutItem"
        Me.OleDbSelectCommand3.Connection = Me.conCutSheet
        '
        'OleDbUpdateCommand3
        '
        Me.OleDbUpdateCommand3.CommandText = "UPDATE CutItem SET CutSheetNo = ?, Item = ?, Length = ?, Qty = ?, ScheduleNo = ?," & _
        " TypeCode = ? WHERE (CutSheetNo = ?) AND (Item = ?) AND (ScheduleNo = ?) AND (Le" & _
        "ngth = ? OR ? IS NULL AND Length IS NULL) AND (Qty = ? OR ? IS NULL AND Qty IS N" & _
        "ULL) AND (TypeCode = ? OR ? IS NULL AND TypeCode IS NULL)"
        Me.OleDbUpdateCommand3.Connection = Me.conCutSheet
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("CutSheetNo", System.Data.OleDb.OleDbType.Integer, 0, "CutSheetNo"))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Item", System.Data.OleDb.OleDbType.Integer, 0, "Item"))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Length", System.Data.OleDb.OleDbType.Double, 0, "Length"))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Qty", System.Data.OleDb.OleDbType.Integer, 0, "Qty"))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("ScheduleNo", System.Data.OleDb.OleDbType.VarWChar, 12, "ScheduleNo"))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("TypeCode", System.Data.OleDb.OleDbType.VarWChar, 5, "TypeCode"))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_CutSheetNo", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CutSheetNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Item", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Item", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ScheduleNo", System.Data.OleDb.OleDbType.VarWChar, 12, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ScheduleNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Length", System.Data.OleDb.OleDbType.Double, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Length", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Length1", System.Data.OleDb.OleDbType.Double, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Length", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Qty", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Qty", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Qty1", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Qty", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_TypeCode", System.Data.OleDb.OleDbType.VarWChar, 5, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "TypeCode", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand3.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_TypeCode1", System.Data.OleDb.OleDbType.VarWChar, 5, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "TypeCode", System.Data.DataRowVersion.Original, Nothing))
        '
        'grpItems
        '
        Me.grpItems.Controls.Add(Me.grdItems)
        Me.grpItems.Controls.Add(Me.btnSaveAll)
        Me.grpItems.Location = New System.Drawing.Point(240, 184)
        Me.grpItems.Name = "grpItems"
        Me.grpItems.Size = New System.Drawing.Size(416, 296)
        Me.grpItems.TabIndex = 3
        Me.grpItems.TabStop = False
        Me.grpItems.Text = "Items"
        '
        'grdItems
        '
        Me.grdItems.CaptionVisible = False
        Me.grdItems.DataMember = ""
        Me.grdItems.DataSource = Me.dvItems
        Me.grdItems.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.grdItems.Location = New System.Drawing.Point(16, 24)
        Me.grdItems.Name = "grdItems"
        Me.grdItems.Size = New System.Drawing.Size(384, 216)
        Me.grdItems.TabIndex = 0
        Me.grdItems.TableStyles.AddRange(New System.Windows.Forms.DataGridTableStyle() {Me.styleItems})
        '
        'dvItems
        '
        Me.dvItems.Table = Me.dsCutSheet.CutItem
        '
        'styleItems
        '
        Me.styleItems.DataGrid = Me.grdItems
        Me.styleItems.GridColumnStyles.AddRange(New System.Windows.Forms.DataGridColumnStyle() {Me.colType, Me.colLength, Me.colQty})
        Me.styleItems.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.styleItems.MappingName = "CutItem"
        '
        'colType
        '
        Me.colType.Format = ""
        Me.colType.FormatInfo = Nothing
        Me.colType.HeaderText = "Type"
        Me.colType.MappingName = "TypeCode"
        Me.colType.NullText = ""
        Me.colType.Width = 108
        '
        'colLength
        '
        Me.colLength.Format = ""
        Me.colLength.FormatInfo = Nothing
        Me.colLength.HeaderText = "Length"
        Me.colLength.MappingName = "Length"
        Me.colLength.NullText = ""
        Me.colLength.Width = 110
        '
        'colQty
        '
        Me.colQty.Format = ""
        Me.colQty.FormatInfo = Nothing
        Me.colQty.HeaderText = "Qty"
        Me.colQty.MappingName = "Qty"
        Me.colQty.NullText = ""
        Me.colQty.Width = 105
        '
        'btnSaveAll
        '
        Me.btnSaveAll.Location = New System.Drawing.Point(320, 256)
        Me.btnSaveAll.Name = "btnSaveAll"
        Me.btnSaveAll.TabIndex = 12
        Me.btnSaveAll.Text = "Save"
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(576, 496)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.TabIndex = 5
        Me.btnClose.Text = "Close"
        '
        'adpSchedules
        '
        Me.adpSchedules.DeleteCommand = Me.OleDbDeleteCommand2
        Me.adpSchedules.InsertCommand = Me.OleDbInsertCommand2
        Me.adpSchedules.SelectCommand = Me.OleDbSelectCommand2
        Me.adpSchedules.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "SchedItem", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("CutSheetNo", "CutSheetNo"), New System.Data.Common.DataColumnMapping("ScheduleNo", "ScheduleNo")})})
        Me.adpSchedules.UpdateCommand = Me.OleDbUpdateCommand2
        '
        'OleDbDeleteCommand2
        '
        Me.OleDbDeleteCommand2.CommandText = "DELETE FROM SchedItem WHERE (CutSheetNo = ?) AND (ScheduleNo = ?)"
        Me.OleDbDeleteCommand2.Connection = Me.conCutSheet
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Delete2_Original_CutSheetNo", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CutSheetNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Delete2_Original_ScheduleNo", System.Data.OleDb.OleDbType.VarWChar, 12, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ScheduleNo", System.Data.DataRowVersion.Original, Nothing))
        '
        'OleDbInsertCommand2
        '
        Me.OleDbInsertCommand2.CommandText = "INSERT INTO SchedItem(CutSheetNo, ScheduleNo) VALUES (?, ?)"
        Me.OleDbInsertCommand2.Connection = Me.conCutSheet
        Me.OleDbInsertCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Insert2_CutSheetNo", System.Data.OleDb.OleDbType.Integer, 0, "CutSheetNo"))
        Me.OleDbInsertCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Insert2_ScheduleNo", System.Data.OleDb.OleDbType.VarWChar, 12, "ScheduleNo"))
        '
        'OleDbSelectCommand2
        '
        Me.OleDbSelectCommand2.CommandText = "SELECT SchedItem.CutSheetNo, SchedItem.ScheduleNo FROM (CuttingSheet INNER JOIN S" & _
        "chedItem ON CuttingSheet.CutSheetNo = SchedItem.CutSheetNo)"
        Me.OleDbSelectCommand2.Connection = Me.conCutSheet
        '
        'OleDbUpdateCommand2
        '
        Me.OleDbUpdateCommand2.CommandText = "UPDATE SchedItem SET CutSheetNo = ?, ScheduleNo = ? WHERE (CutSheetNo = ?) AND (S" & _
        "cheduleNo = ?)"
        Me.OleDbUpdateCommand2.Connection = Me.conCutSheet
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Update2_CutSheetNo", System.Data.OleDb.OleDbType.Integer, 0, "CutSheetNo"))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Update2_ScheduleNo", System.Data.OleDb.OleDbType.VarWChar, 12, "ScheduleNo"))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Update2_Original_CutSheetNo", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CutSheetNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Update2_Original_ScheduleNo", System.Data.OleDb.OleDbType.VarWChar, 12, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ScheduleNo", System.Data.DataRowVersion.Original, Nothing))
        '
        'OleDbDeleteCommand4
        '
        Me.OleDbDeleteCommand4.CommandText = "DELETE FROM Job WHERE (JobNo = ?) AND (AddedDiscount = ? OR ? IS NULL AND AddedDi" & _
        "scount IS NULL) AND (CompanyNo = ? OR ? IS NULL AND CompanyNo IS NULL) AND (Cont" & _
        "ractorNo = ? OR ? IS NULL AND ContractorNo IS NULL) AND (Design = ? OR ? IS NULL" & _
        " AND Design IS NULL) AND (Discount = ? OR ? IS NULL AND Discount IS NULL) AND (J" & _
        "obName = ? OR ? IS NULL AND JobName IS NULL) AND (OrderNo = ? OR ? IS NULL AND O" & _
        "rderNo IS NULL) AND ([Tons Or Kilograms] = ? OR ? IS NULL AND [Tons Or Kilograms" & _
        "] IS NULL) AND (UnitOfMeas = ? OR ? IS NULL AND UnitOfMeas IS NULL)"
        Me.OleDbDeleteCommand4.Connection = Me.conCutSheet
        Me.OleDbDeleteCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_JobNo", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "JobNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddedDiscount", System.Data.OleDb.OleDbType.Currency, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddedDiscount", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddedDiscount1", System.Data.OleDb.OleDbType.Currency, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddedDiscount", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_CompanyNo", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CompanyNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_CompanyNo1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CompanyNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ContractorNo", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ContractorNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ContractorNo1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ContractorNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Design", System.Data.OleDb.OleDbType.Currency, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Design", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Design1", System.Data.OleDb.OleDbType.Currency, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Design", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Discount", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Discount", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Discount1", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Discount", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_JobName", System.Data.OleDb.OleDbType.VarWChar, 70, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "JobName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_JobName1", System.Data.OleDb.OleDbType.VarWChar, 70, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "JobName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_OrderNo", System.Data.OleDb.OleDbType.VarWChar, 25, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "OrderNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_OrderNo1", System.Data.OleDb.OleDbType.VarWChar, 25, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "OrderNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Tons_Or_Kilograms", System.Data.OleDb.OleDbType.VarWChar, 10, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Tons Or Kilograms", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Tons_Or_Kilograms1", System.Data.OleDb.OleDbType.VarWChar, 10, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Tons Or Kilograms", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_UnitOfMeas", System.Data.OleDb.OleDbType.VarWChar, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "UnitOfMeas", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_UnitOfMeas1", System.Data.OleDb.OleDbType.VarWChar, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "UnitOfMeas", System.Data.DataRowVersion.Original, Nothing))
        '
        'OleDbInsertCommand4
        '
        Me.OleDbInsertCommand4.CommandText = "INSERT INTO Job(AddedDiscount, CompanyNo, ContractorNo, Design, Discount, JobName" & _
        ", JobNo, OrderNo, [Tons Or Kilograms], UnitOfMeas) VALUES (?, ?, ?, ?, ?, ?, ?, " & _
        "?, ?, ?)"
        Me.OleDbInsertCommand4.Connection = Me.conCutSheet
        Me.OleDbInsertCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("AddedDiscount", System.Data.OleDb.OleDbType.Currency, 0, "AddedDiscount"))
        Me.OleDbInsertCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("CompanyNo", System.Data.OleDb.OleDbType.VarWChar, 50, "CompanyNo"))
        Me.OleDbInsertCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("ContractorNo", System.Data.OleDb.OleDbType.VarWChar, 50, "ContractorNo"))
        Me.OleDbInsertCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Design", System.Data.OleDb.OleDbType.Currency, 0, "Design"))
        Me.OleDbInsertCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Discount", System.Data.OleDb.OleDbType.Integer, 0, "Discount"))
        Me.OleDbInsertCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("JobName", System.Data.OleDb.OleDbType.VarWChar, 70, "JobName"))
        Me.OleDbInsertCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("JobNo", System.Data.OleDb.OleDbType.VarWChar, 50, "JobNo"))
        Me.OleDbInsertCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("OrderNo", System.Data.OleDb.OleDbType.VarWChar, 25, "OrderNo"))
        Me.OleDbInsertCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Tons_Or_Kilograms", System.Data.OleDb.OleDbType.VarWChar, 10, "Tons Or Kilograms"))
        Me.OleDbInsertCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("UnitOfMeas", System.Data.OleDb.OleDbType.VarWChar, 8, "UnitOfMeas"))
        '
        'OleDbUpdateCommand4
        '
        Me.OleDbUpdateCommand4.CommandText = "UPDATE Job SET AddedDiscount = ?, CompanyNo = ?, ContractorNo = ?, Design = ?, Di" & _
        "scount = ?, JobName = ?, JobNo = ?, OrderNo = ?, [Tons Or Kilograms] = ?, UnitOf" & _
        "Meas = ? WHERE (JobNo = ?) AND (AddedDiscount = ? OR ? IS NULL AND AddedDiscount" & _
        " IS NULL) AND (CompanyNo = ? OR ? IS NULL AND CompanyNo IS NULL) AND (Contractor" & _
        "No = ? OR ? IS NULL AND ContractorNo IS NULL) AND (Design = ? OR ? IS NULL AND D" & _
        "esign IS NULL) AND (Discount = ? OR ? IS NULL AND Discount IS NULL) AND (JobName" & _
        " = ? OR ? IS NULL AND JobName IS NULL) AND (OrderNo = ? OR ? IS NULL AND OrderNo" & _
        " IS NULL) AND ([Tons Or Kilograms] = ? OR ? IS NULL AND [Tons Or Kilograms] IS N" & _
        "ULL) AND (UnitOfMeas = ? OR ? IS NULL AND UnitOfMeas IS NULL)"
        Me.OleDbUpdateCommand4.Connection = Me.conCutSheet
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("AddedDiscount", System.Data.OleDb.OleDbType.Currency, 0, "AddedDiscount"))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("CompanyNo", System.Data.OleDb.OleDbType.VarWChar, 50, "CompanyNo"))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("ContractorNo", System.Data.OleDb.OleDbType.VarWChar, 50, "ContractorNo"))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Design", System.Data.OleDb.OleDbType.Currency, 0, "Design"))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Discount", System.Data.OleDb.OleDbType.Integer, 0, "Discount"))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("JobName", System.Data.OleDb.OleDbType.VarWChar, 70, "JobName"))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("JobNo", System.Data.OleDb.OleDbType.VarWChar, 50, "JobNo"))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("OrderNo", System.Data.OleDb.OleDbType.VarWChar, 25, "OrderNo"))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Tons_Or_Kilograms", System.Data.OleDb.OleDbType.VarWChar, 10, "Tons Or Kilograms"))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("UnitOfMeas", System.Data.OleDb.OleDbType.VarWChar, 8, "UnitOfMeas"))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_JobNo", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "JobNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddedDiscount", System.Data.OleDb.OleDbType.Currency, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddedDiscount", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_AddedDiscount1", System.Data.OleDb.OleDbType.Currency, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "AddedDiscount", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_CompanyNo", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CompanyNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_CompanyNo1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CompanyNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ContractorNo", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ContractorNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ContractorNo1", System.Data.OleDb.OleDbType.VarWChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ContractorNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Design", System.Data.OleDb.OleDbType.Currency, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Design", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Design1", System.Data.OleDb.OleDbType.Currency, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Design", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Discount", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Discount", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Discount1", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Discount", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_JobName", System.Data.OleDb.OleDbType.VarWChar, 70, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "JobName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_JobName1", System.Data.OleDb.OleDbType.VarWChar, 70, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "JobName", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_OrderNo", System.Data.OleDb.OleDbType.VarWChar, 25, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "OrderNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_OrderNo1", System.Data.OleDb.OleDbType.VarWChar, 25, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "OrderNo", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Tons_Or_Kilograms", System.Data.OleDb.OleDbType.VarWChar, 10, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Tons Or Kilograms", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Tons_Or_Kilograms1", System.Data.OleDb.OleDbType.VarWChar, 10, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Tons Or Kilograms", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_UnitOfMeas", System.Data.OleDb.OleDbType.VarWChar, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "UnitOfMeas", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand4.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_UnitOfMeas1", System.Data.OleDb.OleDbType.VarWChar, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "UnitOfMeas", System.Data.DataRowVersion.Original, Nothing))
        '
        'cmdTypeExist
        '
        Me.cmdTypeExist.CommandText = "SELECT * FROM ProductType WHERE (TypeCode = ?)"
        Me.cmdTypeExist.Connection = Me.conCutSheet
        Me.cmdTypeExist.Parameters.Add(New System.Data.OleDb.OleDbParameter("TypeCode", System.Data.OleDb.OleDbType.VarWChar, 8, "TypeCode"))
        '
        'btnPrint
        '
        Me.btnPrint.Location = New System.Drawing.Point(456, 496)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(96, 23)
        Me.btnPrint.TabIndex = 13
        Me.btnPrint.Text = "Print Preview..."
        '
        'cbxSchedNo
        '
        Me.cbxSchedNo.DataSource = Me.dvSchedules
        Me.cbxSchedNo.DisplayMember = "ScheduleNo"
        Me.cbxSchedNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxSchedNo.Location = New System.Drawing.Point(120, 216)
        Me.cbxSchedNo.Name = "cbxSchedNo"
        Me.cbxSchedNo.Size = New System.Drawing.Size(88, 21)
        Me.cbxSchedNo.TabIndex = 4
        Me.cbxSchedNo.ValueMember = "CutSheetNo"
        '
        'frmNewCutSheet
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(674, 536)
        Me.Controls.Add(Me.cbxSchedNo)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.cbxCutSheetNo)
        Me.Controls.Add(Me.grpCutSheetDetails)
        Me.Controls.Add(Me.grpItems)
        Me.Controls.Add(Me.grpSchedules)
        Me.Controls.Add(Me.btnClose)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "frmNewCutSheet"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Cutting Sheet Maintenance"
        CType(Me.dvJob, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dsCutSheet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dvCutSheet, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpCutSheetDetails.ResumeLayout(False)
        Me.grpSchedules.ResumeLayout(False)
        CType(Me.dvSchedules, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpItems.ResumeLayout(False)
        CType(Me.grdItems, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dvItems, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private state As String
    Private stateSched As String

    Private CallingForm As Object

    Private SchedNo As String = ""

    Dim WithEvents tableItems As DataTable

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

    Private Sub disablefields()
        'txtHeading.Enabled = False
        dtpDate.Enabled = False
        txtDetails.Enabled = False
    End Sub

    Private Sub enablefields()
        'txtHeading.Enabled = True
        dtpDate.Enabled = True
        txtDetails.Enabled = True
    End Sub

    Private Sub disablebinding()
        txtCutSheetNo.DataBindings.Clear()
        dtpDate.DataBindings.Clear()
        dtpDate.Value = Today

        txtDetails.DataBindings.Clear()
        txtDetails.Clear()
    End Sub

    Private Sub enablebinding()
        txtCutSheetNo.DataBindings.Add("Text", dvCutSheet, "CutSheetNo")
        'txtHeading.DataBindings.Add("Text", dvCutSheet, "CSHeading")
        dtpDate.DataBindings.Add("Text", dvCutSheet, "CutDate")
        txtDetails.DataBindings.Add("Text", dvCutSheet, "Details")
    End Sub

    Private Sub FilterCuttingSheets(ByVal inJob As String)
        dvCutSheet.RowFilter = "[Job No]='" + inJob + "'"
    End Sub

    Private Sub FilterSchedules()
        dvSchedules.RowFilter = "CutSheetNo=" + cbxCutSheetNo.Text
        dsCutSheet.SchedItem.Columns("CutSheetNo").DefaultValue = cbxCutSheetNo.Text
    End Sub

    Private Sub FilterCutItems()
        dvItems.RowFilter = "CutSheetNo=" + cbxCutSheetNo.Text + " and ScheduleNo='" + cbxSchedNo.Text + "'"
        dsCutSheet.CutItem.Columns("CutSheetNo").DefaultValue = cbxCutSheetNo.Text
        dsCutSheet.CutItem.Columns("ScheduleNo").DefaultValue = cbxSchedNo.Text
    End Sub

    Private Sub frmCutSheet_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        dsCutSheet.Clear()
        disablefields()
    End Sub
    Private Sub txtJob_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtJob.Leave

        Dim Key As String
        Key = txtJob.Text()
        Dim dsJob As New DataSet

        Try
            Dim found As Boolean
            Dim sql As String = "SELECT * FROM Job WHERE JobNo ='" & Key & "'"
            Dim adpJob As New _
                OleDbDataAdapter(sql, cnnJob)
            adpJob.Fill(dsJob, "Job")
            found = False
            Dim job As String
            If dsJob.Tables("Job").Rows.Count > 0 Then
                found = True
                txtJobName.Text = dsJob.Tables("Job").Rows(0).Item("JobName")
                ' IF FOUND
                If found Then
                    adpCutSheet.Fill(dsCutSheet.CuttingSheet)
                    adpSchedules.Fill(dsCutSheet.SchedItem)
                    adpItems.Fill(dsCutSheet.CutItem)

                    FilterCuttingSheets(txtJob.Text)
                    If dvSchedules.Count > 0 Then
                        FilterSchedules()
                    End If

                    If dvItems.Count > 0 Then
                        FilterCutItems()
                    End If

                    tableItems = dsCutSheet.CutItem
                    If cbxCutSheetNo.Text <> "" Then
                        FilterSchedules()

                        If dvSchedules.Count > 0 Then
                            FilterCutItems()
                            grpItems.Enabled = True
                        Else
                            grpItems.Enabled = False
                        End If
                    End If
                End If
            Else
                MessageBox.Show("No Job Found", "Error")
                txtJobName.Text = ""
                grpItems.Enabled = False
                dsCutSheet.Clear()
                disablefields()
            End If
        Catch ee As OleDb.OleDbException
            MessageBox.Show(ee.ToString)
        End Try

    End Sub

    Private Sub cbxCutSheetNo_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxCutSheetNo.SelectedValueChanged
        If cbxCutSheetNo.Text <> "" Then
            FilterSchedules()

            If dvSchedules.Count > 0 Then
                FilterCutItems()
                grpItems.Enabled = True
            Else
                grpItems.Enabled = False
            End If
        End If
    End Sub

    Private Sub grdItems_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdItems.Enter
        If dvSchedules.Count > 0 Then
            FilterCutItems()
            grpItems.Enabled = True
        Else
            grpItems.Enabled = False
        End If
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        Dim nextCut, rowCnt As Integer
        Dim compNum As String
        If state = "" Then
            If txtJob.Text = "" Then
                MsgBox("A Job Number is required", MsgBoxStyle.Critical, "Error")
                txtJob.Focus()
            Else
                txtJob.Enabled = False
                cbxCutSheetNo.Enabled = False
                cbxCutSheetNo.SendToBack()

                enablefields()
                disablebinding()

                grpSchedules.Enabled = False
                compNum = "1"
                getCutSheet(compNum)
                'rowCnt = dsCutSheet.CuttingSheet.Rows.Count - 1
                'nextCut = dsCutSheet.CuttingSheet(rowCnt)("CutSheetNo") + 1
                'txtCutSheetNo.Text = nextCut
                state = "add"
            End If
        End If
    End Sub
    Private Sub getCutSheet(ByVal inCompany As String)

        Dim rowCnt, i As Integer
        Dim lastCut As Integer
        Try
            cnnJob.Open()

            ' GET ALL THE CUTTING SHEETS FOR THAT JOB
            Dim sql4 = "SELECT Company.lastCutNum" & _
            " FROM Company " & _
            "WHERE companyNo = '" & inCompany & "'"

            Dim compDS = New Data.DataSet
            Dim adapter = New OleDb.OleDbDataAdapter(sql4, cnnJob)

            adapter.Fill(compDS)
            rowCnt = compDS.tables(0).rows.count

            If rowCnt <> 1 Then
                MessageBox.Show("Error in company record ", rowCnt)
            End If
            ' /* FOR EACH CUTTING SHEET*/
            For i = 0 To rowCnt - 1
                lastCut = compDS.tables(0).rows(i).item("lastCutNum").ToString()
            Next i
            txtCutSheetNo.Text = lastCut + 1
            cnnJob.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
    Private Sub updateCut(ByVal inCompany As String, ByVal lastCut As Integer)
        Dim sqlChgJob As String
        Dim command As OleDb.OleDbCommand = New OleDb.OleDbCommand

        sqlChgJob = "UPDATE Company SET lastCutNum = '" & lastCut & "'" & _
                  " WHERE companyNo = '" & inCompany & "'"

        Dim daJob = New OleDb.OleDbDataAdapter

        Try
            cnnJob.Open()
            daJob.UpdateCommand = command
            daJob.UpdateCommand.CommandText = sqlChgJob
            daJob.UpdateCommand.Connection = cnnJob
            daJob.UpdateCommand.ExecuteNonQuery()
            cnnJob.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If state = "add" Then
            Dim row As DataRow = dsCutSheet.CuttingSheet.NewCuttingSheetRow
            Dim compNum As String = "1"
            Dim lastCut As Integer

            row("Job No") = txtJob.Text
            row("CutSheetNo") = txtCutSheetNo.Text
            row("CutDate") = dtpDate.Value.Date
            row("InvoiceNo") = 0

            If txtDetails.Text <> "" Then
                row("Details") = txtDetails.Text
            End If

            dsCutSheet.CuttingSheet.AddCuttingSheetRow(row)

            adpCutSheet.Update(dsCutSheet.CuttingSheet)
            MsgBox("Changes were successfully saved to cutting sheet", MsgBoxStyle.Information, "Information")

            cbxCutSheetNo.SelectedIndex = cbxCutSheetNo.Items.Count() - 1

            txtJob.Enabled = True
            cbxCutSheetNo.Enabled = True
            cbxCutSheetNo.BringToFront()

            disablefields()
            enablebinding()

            grpSchedules.Enabled = True
            lastCut = txtCutSheetNo.Text
            cbxCutSheetNo.Focus()
            updateCut(compNum, lastCut)

            state = ""
        End If
    End Sub

    Private Sub btnSaveAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveAll.Click
        adpItems.Update(dsCutSheet.CutItem)

        MsgBox("Changes were successfully saved", MsgBoxStyle.Information, "Information")
    End Sub

    Private Sub grdItems_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdItems.Click
        If dvSchedules.Count > 0 Then
            FilterCutItems()
            grpItems.Enabled = True
        Else
            grpItems.Enabled = False
        End If
    End Sub

    Private Sub tableItems_ColumnChanging(ByVal sender As Object, ByVal e As System.Data.DataColumnChangeEventArgs) Handles tableItems.ColumnChanging
        If e.Column.ColumnName = "TypeCode" Then
            If Not IsDBNull(e.ProposedValue) Then
                e.ProposedValue = e.ProposedValue.ToString.ToUpper

                Dim DataReader As System.Data.OleDb.OleDbDataReader
                Dim count As Integer = 0

                conCutSheet.Open()
                cmdTypeExist.Parameters("TypeCode").Value = e.ProposedValue
                DataReader = cmdTypeExist.ExecuteReader(CommandBehavior.CloseConnection)
                While DataReader.Read()
                    count += 1
                End While
                DataReader.Close()
                conCutSheet.Close()

                If count < 1 Then
                    MsgBox("Type does not exist", MsgBoxStyle.Critical, "Error")
                    Throw New ArgumentException("Type does no exist")

                    e.Row.Item("Item") = ""
                    'grdJobRate.CurrentCell = e.Row.Item("TypeCode")
                Else
                    e.Row.Item("Item") = grdItems.CurrentRowIndex + 1
                End If
            End If
        End If
    End Sub

    Private Sub cbxCutSheetNo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxCutSheetNo.SelectedIndexChanged
        If cbxCutSheetNo.Text = Nothing Then
            dvSchedules.RowFilter = "CutSheetNo=-1"
            dvItems.RowFilter = "CutSheetNo=-1 and ScheduleNo='-1'"
        End If
    End Sub

    Private Sub btnPrint_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim form As PrintCutSheet
        form = New PrintCutSheet(Me)
        Me.Hide()
        form.Show()
        form.txtCutNum.Text = cbxCutSheetNo.Text
        form.btn_Print.PerformClick()
        form.btn_Close.PerformClick()
    End Sub

    Private Sub txtJob_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If txtCutSheetNo.Text = Nothing Then
            grpSchedules.Enabled = False
        Else
            grpSchedules.Enabled = True
        End If
    End Sub

    Private Sub btnAddSched_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddSched.Click
        If stateSched = "" Then
            grpCutSheetDetails.Enabled = False

            cbxSchedNo.Enabled = False
            cbxSchedNo.SendToBack()

            txtSchedNo.Enabled = True
            txtSchedNo.DataBindings.Clear()
            txtSchedNo.Text = Nothing
            txtSchedNo.Focus()

            grpItems.Enabled = False

            stateSched = "add"
        End If
    End Sub

    Private Sub btnSaveSched_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveSched.Click
        If stateSched = "add" Then
            Dim row As DataRow = dsCutSheet.SchedItem.NewSchedItemRow
            row("ScheduleNo") = txtSchedNo.Text
            row("CutSheetNo") = txtCutSheetNo.Text

            dsCutSheet.SchedItem.AddSchedItemRow(row)

            adpSchedules.Update(dsCutSheet.SchedItem)
            MsgBox("Changes were successfully saved", MsgBoxStyle.Information, "Information")

            cbxSchedNo.SelectedIndex = cbxSchedNo.Items.Count() - 1

            grpCutSheetDetails.Enabled = True

            cbxSchedNo.Enabled = True
            cbxSchedNo.BringToFront()

            txtSchedNo.Enabled = False
            txtSchedNo.DataBindings.Add("Text", dvSchedules, "ScheduleNo")

            grpItems.Enabled = True

            stateSched = ""
        End If
    End Sub

    Private Sub cbxSchedNo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxSchedNo.SelectedIndexChanged
        If cbxSchedNo.SelectedIndex = -1 Then
            grpItems.Enabled = False
        Else
            grpItems.Enabled = True
        End If

        FilterCutItems()
    End Sub

    Private Sub btnDelSched_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelSched.Click
        If stateSched = "" Then
            If txtSchedNo.Text <> Nothing Then
                If MessageBox.Show("Delete selected schedule?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then

                    Dim cmd As New OleDb.OleDbCommand

                    conCutSheet.Open()
                    cmd.Connection = conCutSheet
                    cmd.CommandText = "DELETE FROM CutItem WHERE CutSheetNo=" + txtCutSheetNo.Text + " AND ScheduleNo='" + txtSchedNo.Text + "'"
                    cmd.ExecuteNonQuery()
                    conCutSheet.Close()

                    dsCutSheet.SchedItem.FindByCutSheetNoScheduleNo(txtCutSheetNo.Text, txtSchedNo.Text).Delete()
                    adpSchedules.Update(dsCutSheet.SchedItem)

                    btnDelSched.Focus()

                    FilterCutItems()
                End If
            End If
        End If

    End Sub

    Private Sub txtJob_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtJob.TextChanged

    End Sub

    Private Sub grpCutSheetDetails_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles grpCutSheetDetails.Enter

    End Sub
End Class
