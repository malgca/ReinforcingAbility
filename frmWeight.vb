Public Class frmWeight
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
    Friend WithEvents conWeight As System.Data.OleDb.OleDbConnection
    Friend WithEvents adpWeight As System.Data.OleDb.OleDbDataAdapter
    Friend WithEvents OleDbSelectCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbInsertCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbUpdateCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbDeleteCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents dsWeight As Reinforcing_Ability.dsReinforcingAbility
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents dvWeight As System.Data.DataView
    Friend WithEvents grdWeight As System.Windows.Forms.DataGrid
    Friend WithEvents styleWeight As System.Windows.Forms.DataGridTableStyle
    Friend WithEvents colType As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents colWeight As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents btnSave As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.grdWeight = New System.Windows.Forms.DataGrid
        Me.dvWeight = New System.Data.DataView
        Me.dsWeight = New Reinforcing_Ability.dsReinforcingAbility
        Me.styleWeight = New System.Windows.Forms.DataGridTableStyle
        Me.colType = New System.Windows.Forms.DataGridTextBoxColumn
        Me.colWeight = New System.Windows.Forms.DataGridTextBoxColumn
        Me.conWeight = New System.Data.OleDb.OleDbConnection
        Me.adpWeight = New System.Data.OleDb.OleDbDataAdapter
        Me.OleDbDeleteCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbInsertCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbSelectCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbUpdateCommand1 = New System.Data.OleDb.OleDbCommand
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        CType(Me.grdWeight, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dvWeight, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dsWeight, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'grdWeight
        '
        Me.grdWeight.CaptionVisible = False
        Me.grdWeight.DataMember = ""
        Me.grdWeight.DataSource = Me.dvWeight
        Me.grdWeight.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.grdWeight.Location = New System.Drawing.Point(16, 16)
        Me.grdWeight.Name = "grdWeight"
        Me.grdWeight.Size = New System.Drawing.Size(216, 456)
        Me.grdWeight.TabIndex = 0
        Me.grdWeight.TableStyles.AddRange(New System.Windows.Forms.DataGridTableStyle() {Me.styleWeight})
        '
        'dvWeight
        '
        Me.dvWeight.Table = Me.dsWeight.ProductType
        '
        'dsWeight
        '
        Me.dsWeight.DataSetName = "dsReinforcingAbility"
        Me.dsWeight.Locale = New System.Globalization.CultureInfo("en-ZA")
        '
        'styleWeight
        '
        Me.styleWeight.DataGrid = Me.grdWeight
        Me.styleWeight.GridColumnStyles.AddRange(New System.Windows.Forms.DataGridColumnStyle() {Me.colType, Me.colWeight})
        Me.styleWeight.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.styleWeight.MappingName = "ProductType"
        '
        'colType
        '
        Me.colType.Alignment = System.Windows.Forms.HorizontalAlignment.Center
        Me.colType.Format = ""
        Me.colType.FormatInfo = Nothing
        Me.colType.HeaderText = "Type"
        Me.colType.MappingName = "TypeCode"
        Me.colType.NullText = ""
        Me.colType.Width = 80
        '
        'colWeight
        '
        Me.colWeight.Alignment = System.Windows.Forms.HorizontalAlignment.Center
        Me.colWeight.Format = ""
        Me.colWeight.FormatInfo = Nothing
        Me.colWeight.HeaderText = "Weight"
        Me.colWeight.MappingName = "Weight"
        Me.colWeight.NullText = ""
        Me.colWeight.Width = 80
        '
        'conWeight
        '
        Me.conWeight.ConnectionString = "Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Registry Path=;Jet OLEDB:Database L" & _
        "ocking Mode=1;Data Source=""winsteelVers5.mdb"";Mode=Share Deny None;Jet OLEDB:Eng" & _
        "ine Type=5;Provider=""Microsoft.Jet.OLEDB.4.0"";Jet OLEDB:System database=;Jet OLE" & _
        "DB:SFP=False;persist security info=False;Extended Properties=;Jet OLEDB:Compact " & _
        "Without Replica Repair=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Create S" & _
        "ystem Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;User ID=Admin;" & _
        "Jet OLEDB:Global Bulk Transactions=1"
        '
        'adpWeight
        '
        Me.adpWeight.DeleteCommand = Me.OleDbDeleteCommand1
        Me.adpWeight.InsertCommand = Me.OleDbInsertCommand1
        Me.adpWeight.SelectCommand = Me.OleDbSelectCommand1
        Me.adpWeight.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "ProductType", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("TypeCode", "TypeCode"), New System.Data.Common.DataColumnMapping("Weight", "Weight")})})
        Me.adpWeight.UpdateCommand = Me.OleDbUpdateCommand1
        '
        'OleDbDeleteCommand1
        '
        Me.OleDbDeleteCommand1.CommandText = "DELETE FROM ProductType WHERE (TypeCode = ?) AND (Weight = ? OR ? IS NULL AND Wei" & _
        "ght IS NULL)"
        Me.OleDbDeleteCommand1.Connection = Me.conWeight
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_TypeCode", System.Data.OleDb.OleDbType.VarWChar, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "TypeCode", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Weight", System.Data.OleDb.OleDbType.Double, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Weight", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Weight1", System.Data.OleDb.OleDbType.Double, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Weight", System.Data.DataRowVersion.Original, Nothing))
        '
        'OleDbInsertCommand1
        '
        Me.OleDbInsertCommand1.CommandText = "INSERT INTO ProductType(TypeCode, Weight) VALUES (?, ?)"
        Me.OleDbInsertCommand1.Connection = Me.conWeight
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("TypeCode", System.Data.OleDb.OleDbType.VarWChar, 8, "TypeCode"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Weight", System.Data.OleDb.OleDbType.Double, 0, "Weight"))
        '
        'OleDbSelectCommand1
        '
        Me.OleDbSelectCommand1.CommandText = "SELECT TypeCode, Weight FROM ProductType"
        Me.OleDbSelectCommand1.Connection = Me.conWeight
        '
        'OleDbUpdateCommand1
        '
        Me.OleDbUpdateCommand1.CommandText = "UPDATE ProductType SET TypeCode = ?, Weight = ? WHERE (TypeCode = ?) AND (Weight " & _
        "= ? OR ? IS NULL AND Weight IS NULL)"
        Me.OleDbUpdateCommand1.Connection = Me.conWeight
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("TypeCode", System.Data.OleDb.OleDbType.VarWChar, 8, "TypeCode"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Weight", System.Data.OleDb.OleDbType.Double, 0, "Weight"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_TypeCode", System.Data.OleDb.OleDbType.VarWChar, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "TypeCode", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Weight", System.Data.OleDb.OleDbType.Double, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Weight", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Weight1", System.Data.OleDb.OleDbType.Double, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Weight", System.Data.DataRowVersion.Original, Nothing))
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(160, 480)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.TabIndex = 2
        Me.btnClose.Text = "Close"
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(16, 480)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.TabIndex = 1
        Me.btnSave.Text = "Save"
        '
        'frmWeight
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(248, 518)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.grdWeight)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "frmWeight"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Weight Maintenance"
        CType(Me.grdWeight, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dvWeight, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dsWeight, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

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

    Private Sub frmWeight_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If Not IsNothing(CallingForm) Then
            CallingForm.Show()
        End If

        CallingForm = Nothing
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Close()
    End Sub

    Private Sub frmWeight_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        dsWeight.Clear()
        adpWeight.Fill(dsWeight.ProductType)
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        adpWeight.Update(dsWeight.ProductType)
        MsgBox("Changes were successfully saved", MsgBoxStyle.Information, "Information")
    End Sub
End Class
