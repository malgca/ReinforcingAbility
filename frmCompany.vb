Imports System
Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Data.OleDb
Imports LogicTier

Public Class frmCompany
    Inherits Form

    Private Property FormState As FormStates
    Private Property Logic As New Company
    Private Property CallingForm As Object

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()
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
    Private components As IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents lblCompNo As Label
    Friend WithEvents grpCompDetails As GroupBox
    Friend WithEvents txtCompNo As TextBox
    Friend WithEvents lblRegNo As Label
    Friend WithEvents lblAddress As Label
    Friend WithEvents txtAddress As TextBox
    Friend WithEvents lblPostalCode As Label
    Friend WithEvents lblCompName As Label
    Friend WithEvents lblVATNo As Label
    Friend WithEvents grpContactDetails As GroupBox
    Friend WithEvents lblTelNo As Label
    Friend WithEvents lblFaxNo As Label
    Friend WithEvents lblEmail As Label
    Friend WithEvents lblWebsite As Label
    Friend WithEvents grpMiscDetails As GroupBox
    Friend WithEvents lblMessage As Label
    Friend WithEvents lblLastInvNo As Label
    Friend WithEvents lblVAT As Label
    Friend WithEvents btnAdd As Button
    Friend WithEvents btnSave As Button
    Friend WithEvents btnClose As Button
    Friend WithEvents txtCompName As TextBox
    Friend WithEvents txtVATNo As TextBox
    Friend WithEvents txtPostalCode As TextBox
    Friend WithEvents txtWebsite As TextBox
    Friend WithEvents txtEmail As TextBox
    Friend WithEvents txtFaxNo As TextBox
    Friend WithEvents txtTelNo As TextBox
    Friend WithEvents txtMessage As TextBox
    Friend WithEvents txtVAT As TextBox
    Friend WithEvents txtLastInvNo As TextBox
    Friend WithEvents txtRegNo As TextBox
    Friend WithEvents cbxCompNo As ComboBox
    Friend WithEvents btnEdit As Button
    Friend WithEvents txtAddress2 As TextBox
    Friend WithEvents txtAddress3 As TextBox
    <Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.grpCompDetails = New System.Windows.Forms.GroupBox()
        Me.txtAddress3 = New System.Windows.Forms.TextBox()
        Me.txtAddress2 = New System.Windows.Forms.TextBox()
        Me.txtVATNo = New System.Windows.Forms.TextBox()
        Me.lblVATNo = New System.Windows.Forms.Label()
        Me.txtCompName = New System.Windows.Forms.TextBox()
        Me.lblCompName = New System.Windows.Forms.Label()
        Me.txtPostalCode = New System.Windows.Forms.TextBox()
        Me.lblPostalCode = New System.Windows.Forms.Label()
        Me.txtAddress = New System.Windows.Forms.TextBox()
        Me.lblAddress = New System.Windows.Forms.Label()
        Me.txtRegNo = New System.Windows.Forms.TextBox()
        Me.lblRegNo = New System.Windows.Forms.Label()
        Me.txtCompNo = New System.Windows.Forms.TextBox()
        Me.lblCompNo = New System.Windows.Forms.Label()
        Me.grpContactDetails = New System.Windows.Forms.GroupBox()
        Me.txtWebsite = New System.Windows.Forms.TextBox()
        Me.lblWebsite = New System.Windows.Forms.Label()
        Me.txtEmail = New System.Windows.Forms.TextBox()
        Me.lblEmail = New System.Windows.Forms.Label()
        Me.txtFaxNo = New System.Windows.Forms.TextBox()
        Me.lblFaxNo = New System.Windows.Forms.Label()
        Me.txtTelNo = New System.Windows.Forms.TextBox()
        Me.lblTelNo = New System.Windows.Forms.Label()
        Me.grpMiscDetails = New System.Windows.Forms.GroupBox()
        Me.txtVAT = New System.Windows.Forms.TextBox()
        Me.lblVAT = New System.Windows.Forms.Label()
        Me.txtLastInvNo = New System.Windows.Forms.TextBox()
        Me.lblLastInvNo = New System.Windows.Forms.Label()
        Me.txtMessage = New System.Windows.Forms.TextBox()
        Me.lblMessage = New System.Windows.Forms.Label()
        Me.btnAdd = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.cbxCompNo = New System.Windows.Forms.ComboBox()
        Me.btnEdit = New System.Windows.Forms.Button()
        Me.grpCompDetails.SuspendLayout()
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
        'txtAddress3
        '
        Me.txtAddress3.Location = New System.Drawing.Point(104, 160)
        Me.txtAddress3.Name = "txtAddress3"
        Me.txtAddress3.Size = New System.Drawing.Size(184, 20)
        Me.txtAddress3.TabIndex = 13
        '
        'txtAddress2
        '
        Me.txtAddress2.Location = New System.Drawing.Point(104, 128)
        Me.txtAddress2.Name = "txtAddress2"
        Me.txtAddress2.Size = New System.Drawing.Size(184, 20)
        Me.txtAddress2.TabIndex = 12
        '
        'txtVATNo
        '
        Me.txtVATNo.Location = New System.Drawing.Point(352, 64)
        Me.txtVATNo.MaxLength = 20
        Me.txtVATNo.Name = "txtVATNo"
        Me.txtVATNo.Size = New System.Drawing.Size(136, 20)
        Me.txtVATNo.TabIndex = 7
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
        Me.btnAdd.Size = New System.Drawing.Size(75, 23)
        Me.btnAdd.TabIndex = 3
        Me.btnAdd.Text = "Add"
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(488, 480)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 23)
        Me.btnSave.TabIndex = 5
        Me.btnSave.Text = "Save"
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(576, 480)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 6
        Me.btnClose.Text = "Close"
        '
        'cbxCompNo
        '
        Me.cbxCompNo.DataSource = Logic.CompanyNameList
        Me.cbxCompNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxCompNo.Location = New System.Drawing.Point(120, 48)
        Me.cbxCompNo.Name = "cbxCompNo"
        Me.cbxCompNo.Size = New System.Drawing.Size(504, 21)
        Me.cbxCompNo.TabIndex = 7
        '
        'btnEdit
        '
        Me.btnEdit.Location = New System.Drawing.Point(104, 480)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(75, 23)
        Me.btnEdit.TabIndex = 4
        Me.btnEdit.Text = "Edit"
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
        Me.MaximizeBox = False
        Me.Name = "frmCompany"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Company Maintenance"
        Me.grpCompDetails.ResumeLayout(False)
        Me.grpCompDetails.PerformLayout()
        Me.grpContactDetails.ResumeLayout(False)
        Me.grpContactDetails.PerformLayout()
        Me.grpMiscDetails.ResumeLayout(False)
        Me.grpMiscDetails.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public Sub New(ByVal caller As Object)
        MyBase.New()
        InitializeComponent()

        ' save the calling form
        CallingForm = caller
    End Sub

    ''' <summary>
    ''' Converts Enter key to Tab key when pressed.
    ''' </summary>
    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, ByVal key As Keys) As Boolean
        If key = Keys.Enter Then
            SendKeys.Send("{Tab}")
            Return True
        End If

        Return MyBase.ProcessCmdKey(msg, key)
    End Function

    Private Sub frmCompany_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        AddDataBindings()
        DisableForm()
    End Sub
    ' disables all elements on the company form
    Private Sub DisableForm()
        grpCompDetails.Enabled = False
        grpContactDetails.Enabled = False
        grpMiscDetails.Enabled = False
    End Sub

    ' enable all elements on the company form
    Private Sub EnableForm()
        grpCompDetails.Enabled = True
        grpContactDetails.Enabled = True
        grpMiscDetails.Enabled = True
    End Sub

    ' clears databindings on all text fields
    Private Sub ClearDataBindings()
        txtCompNo.DataBindings.Clear()
        txtCompName.DataBindings.Clear()
        txtRegNo.DataBindings.Clear()
        txtVATNo.DataBindings.Clear()
        txtAddress.DataBindings.Clear()
        txtAddress2.DataBindings.Clear()
        txtAddress3.DataBindings.Clear()
        txtPostalCode.DataBindings.Clear()
        txtTelNo.DataBindings.Clear()
        txtEmail.DataBindings.Clear()
        txtFaxNo.DataBindings.Clear()
        txtWebsite.DataBindings.Clear()
        txtMessage.DataBindings.Clear()
        txtVAT.DataBindings.Clear()
        txtLastInvNo.DataBindings.Clear()
    End Sub

    ' clears text from all text fields
    Private Sub ClearTextFields()
        txtCompNo.Clear()
        txtCompName.Clear()
        txtRegNo.Clear()
        txtVATNo.Clear()
        txtAddress.Clear()
        txtAddress2.Clear()
        txtAddress3.Clear()
        txtPostalCode.Clear()
        txtTelNo.Clear()
        txtEmail.Clear()
        txtFaxNo.Clear()
        txtWebsite.Clear()
        txtMessage.Clear()
        txtVAT.Clear()
        txtLastInvNo.Clear()
    End Sub

    ' binds database data to text fields
    Private Sub AddDataBindings()
        txtCompNo.DataBindings.Add("Text", Logic, "CompanyNumber", False, DataSourceUpdateMode.OnPropertyChanged)
        txtCompName.DataBindings.Add("Text", Logic, "CompanyName", False, DataSourceUpdateMode.OnPropertyChanged)
        txtRegNo.DataBindings.Add("Text", Logic, "RegNumber", False, DataSourceUpdateMode.OnPropertyChanged)
        txtVATNo.DataBindings.Add("Text", Logic, "VatNumber", False, DataSourceUpdateMode.OnPropertyChanged)
        txtAddress.DataBindings.Add("Text", Logic, "AddressLine1", False, DataSourceUpdateMode.OnPropertyChanged)
        txtAddress2.DataBindings.Add("Text", Logic, "AddressLine2", False, DataSourceUpdateMode.OnPropertyChanged)
        txtAddress3.DataBindings.Add("Text", Logic, "AddressLine3", False, DataSourceUpdateMode.OnPropertyChanged)
        txtPostalCode.DataBindings.Add("Text", Logic, "PostalCode", False, DataSourceUpdateMode.OnPropertyChanged)
        txtTelNo.DataBindings.Add("Text", Logic, "Telephone", False, DataSourceUpdateMode.OnPropertyChanged)
        txtEmail.DataBindings.Add("Text", Logic, "Email", False, DataSourceUpdateMode.OnPropertyChanged)
        txtFaxNo.DataBindings.Add("Text", Logic, "Fax", False, DataSourceUpdateMode.OnPropertyChanged)
        txtWebsite.DataBindings.Add("Text", Logic, "Website", False, DataSourceUpdateMode.OnPropertyChanged)
        txtMessage.DataBindings.Add("Text", Logic, "Message", False, DataSourceUpdateMode.OnPropertyChanged)
        txtVAT.DataBindings.Add("Text", Logic, "VAT", False, DataSourceUpdateMode.OnPropertyChanged)
        txtLastInvNo.DataBindings.Add("Text", Logic, "LastInvoiceNumber", False, DataSourceUpdateMode.OnPropertyChanged)
    End Sub

    ' handles form add logic
    Private Sub btnAdd_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnAdd.Click
        If FormState = FormStates.Empty Then
            cbxCompNo.SendToBack()
            cbxCompNo.Enabled = False

            EnableForm()
            'ClearDataBindings()
            ClearTextFields()

            txtVAT.Text = Logic.VAT
            txtCompNo.Focus()
            FormState = FormStates.Add
        End If
    End Sub
    ' handles form save logic
    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnSave.Click
        If FormState = FormStates.Add Then
            If txtCompNo.Text = "" Then
                MsgBox("A Company Number is required", MsgBoxStyle.Critical, "Error")
                txtCompNo.Focus()
            Else
                Dim count As New Integer

                Logic.GetCompanyCount(count)

                If count > 0 Then
                    MsgBox("Company Number entered is already used", MsgBoxStyle.Critical, "Error")
                    txtCompNo.Focus()
                Else
                    Logic.AddRowToCompanyTable()
                    MsgBox("Record was successfully saved", MsgBoxStyle.Information, "Information")

                    'DataBindTextFields()
                    cbxCompNo.BringToFront()
                    DisableForm()
                    FormState = FormStates.Empty
                End If
            End If
        End If

        If FormState = FormStates.Edit Then
            Logic.SaveEditToCompanyTable()
            MsgBox("Record was successfully saved", MsgBoxStyle.Information, "Information")

            DisableForm()
            ' enable required fields
            cbxCompNo.BringToFront()
            cbxCompNo.Enabled = True
            txtCompNo.Enabled = True

            FormState = FormStates.Empty
        End If
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnEdit.Click
        If FormState = FormStates.Empty Then
            cbxCompNo.SendToBack()
            txtCompNo.Enabled = False

            EnableForm()

            txtCompName.Focus()
            FormState = FormStates.Edit
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmCompany_Closing(ByVal sender As Object, ByVal e As CancelEventArgs) Handles MyBase.Closing
        If Not IsNothing(CallingForm) Then
            CallingForm.Show()
        End If

        CallingForm = Nothing
    End Sub

    Private Sub cbxCompNo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxCompNo.SelectedIndexChanged
        Logic.InitializeContractorProperties(cbxCompNo.SelectedIndex)
    End Sub
End Class
