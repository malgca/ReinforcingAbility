Imports System
Imports ComponentModel
Imports Drawing
Imports Windows.Forms
Imports Data.OleDb
Imports LogicTier

Public Class frmContractor
    Inherits Form

#Region " Windows Form Designer generated code "

    Private Property FormState As FormStates
    Private Property CallingForm As Object
    Private Property Logic As New Contractor

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
    Private components As ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents grpContractorDetails As GroupBox
    Friend WithEvents lblContractorNo As Label
    Friend WithEvents txtContractorNo As TextBox
    Friend WithEvents lblContractorName As Label
    Friend WithEvents txtContractorName As TextBox
    Friend WithEvents lblAddress1 As Label
    Friend WithEvents txtAddress2 As TextBox
    Friend WithEvents lblAddress2 As Label
    Friend WithEvents txtAddress3 As TextBox
    Friend WithEvents lblAddress3 As Label
    Friend WithEvents txtAddress4 As TextBox
    Friend WithEvents lblAddress4 As Label
    Friend WithEvents lblPostalCode As Label
    Friend WithEvents lblTelNo As Label
    Friend WithEvents txtTelNo As TextBox
    Friend WithEvents btnAdd As Button
    Friend WithEvents btnSave As Button
    Friend WithEvents btnClose As Button
    Friend WithEvents txtAddress1 As TextBox
    Friend WithEvents txtPostalCode As TextBox
    Friend WithEvents txtVATNo As TextBox
    Friend WithEvents lblVATNo As Label
    Friend WithEvents cbxCompNo As ComboBox
    Friend WithEvents btnEdit As Button
    <Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.grpContractorDetails = New GroupBox
        Me.txtVATNo = New TextBox
        Me.lblVATNo = New Label
        Me.txtTelNo = New TextBox
        Me.lblTelNo = New Label
        Me.txtPostalCode = New TextBox
        Me.lblPostalCode = New Label
        Me.txtAddress4 = New TextBox
        Me.lblAddress4 = New Label
        Me.txtAddress3 = New TextBox
        Me.lblAddress3 = New Label
        Me.txtAddress2 = New TextBox
        Me.lblAddress2 = New Label
        Me.txtAddress1 = New TextBox
        Me.lblAddress1 = New Label
        Me.txtContractorName = New TextBox
        Me.lblContractorName = New Label
        Me.txtContractorNo = New TextBox
        Me.lblContractorNo = New Label
        Me.btnAdd = New Button
        Me.btnSave = New Button
        Me.btnClose = New Button
        Me.cbxCompNo = New ComboBox
        Me.btnEdit = New Button
        Me.grpContractorDetails.SuspendLayout()
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
        Me.grpContractorDetails.Location = New Point(16, 16)
        Me.grpContractorDetails.Name = "grpContractorDetails"
        Me.grpContractorDetails.Size = New Size(496, 320)
        Me.grpContractorDetails.TabIndex = 0
        Me.grpContractorDetails.TabStop = False
        Me.grpContractorDetails.Text = "Contractor Details"
        '
        'txtVATNo
        '
        Me.txtVATNo.Location = New Point(128, 96)
        Me.txtVATNo.MaxLength = 20
        Me.txtVATNo.Name = "txtVATNo"
        Me.txtVATNo.Size = New Size(136, 20)
        Me.txtVATNo.TabIndex = 5
        Me.txtVATNo.Text = ""
        '
        'lblVATNo
        '
        Me.lblVATNo.Location = New Point(72, 96)
        Me.lblVATNo.Name = "lblVATNo"
        Me.lblVATNo.Size = New Size(48, 16)
        Me.lblVATNo.TabIndex = 4
        Me.lblVATNo.Text = "VAT No."
        Me.lblVATNo.TextAlign = ContentAlignment.MiddleRight
        '
        'txtTelNo
        '
        Me.txtTelNo.Location = New Point(128, 288)
        Me.txtTelNo.MaxLength = 15
        Me.txtTelNo.Name = "txtTelNo"
        Me.txtTelNo.Size = New Size(120, 20)
        Me.txtTelNo.TabIndex = 17
        Me.txtTelNo.Text = ""
        '
        'lblTelNo
        '
        Me.lblTelNo.Location = New Point(24, 288)
        Me.lblTelNo.Name = "lblTelNo"
        Me.lblTelNo.Size = New Size(100, 16)
        Me.lblTelNo.TabIndex = 16
        Me.lblTelNo.Text = "Tel. No."
        Me.lblTelNo.TextAlign = ContentAlignment.MiddleRight
        '
        'txtPostalCode
        '
        Me.txtPostalCode.Location = New Point(128, 256)
        Me.txtPostalCode.MaxLength = 5
        Me.txtPostalCode.Name = "txtPostalCode"
        Me.txtPostalCode.Size = New Size(64, 20)
        Me.txtPostalCode.TabIndex = 15
        Me.txtPostalCode.Text = ""
        '
        'lblPostalCode
        '
        Me.lblPostalCode.Location = New Point(24, 256)
        Me.lblPostalCode.Name = "lblPostalCode"
        Me.lblPostalCode.Size = New Size(100, 16)
        Me.lblPostalCode.TabIndex = 14
        Me.lblPostalCode.Text = "Postal Code"
        Me.lblPostalCode.TextAlign = ContentAlignment.MiddleRight
        '
        'txtAddress4
        '
        Me.txtAddress4.Location = New Point(128, 224)
        Me.txtAddress4.MaxLength = 40
        Me.txtAddress4.Name = "txtAddress4"
        Me.txtAddress4.Size = New Size(216, 20)
        Me.txtAddress4.TabIndex = 13
        Me.txtAddress4.Text = ""
        '
        'lblAddress4
        '
        Me.lblAddress4.Location = New Point(24, 224)
        Me.lblAddress4.Name = "lblAddress4"
        Me.lblAddress4.Size = New Size(100, 16)
        Me.lblAddress4.TabIndex = 12
        Me.lblAddress4.Text = "Address 4"
        Me.lblAddress4.TextAlign = ContentAlignment.MiddleRight
        '
        'txtAddress3
        '
        Me.txtAddress3.Location = New Point(128, 192)
        Me.txtAddress3.MaxLength = 40
        Me.txtAddress3.Name = "txtAddress3"
        Me.txtAddress3.Size = New Size(216, 20)
        Me.txtAddress3.TabIndex = 11
        Me.txtAddress3.Text = ""
        '
        'lblAddress3
        '
        Me.lblAddress3.Location = New Point(24, 192)
        Me.lblAddress3.Name = "lblAddress3"
        Me.lblAddress3.Size = New Size(100, 16)
        Me.lblAddress3.TabIndex = 10
        Me.lblAddress3.Text = "Address 3"
        Me.lblAddress3.TextAlign = ContentAlignment.MiddleRight
        '
        'txtAddress2
        '
        Me.txtAddress2.Location = New Point(128, 160)
        Me.txtAddress2.MaxLength = 40
        Me.txtAddress2.Name = "txtAddress2"
        Me.txtAddress2.Size = New Size(216, 20)
        Me.txtAddress2.TabIndex = 9
        Me.txtAddress2.Text = ""
        '
        'lblAddress2
        '
        Me.lblAddress2.Location = New Point(24, 160)
        Me.lblAddress2.Name = "lblAddress2"
        Me.lblAddress2.Size = New Size(100, 16)
        Me.lblAddress2.TabIndex = 8
        Me.lblAddress2.Text = "Address 2"
        Me.lblAddress2.TextAlign = ContentAlignment.MiddleRight
        '
        'txtAddress1
        '
        Me.txtAddress1.Location = New Point(128, 128)
        Me.txtAddress1.MaxLength = 40
        Me.txtAddress1.Name = "txtAddress1"
        Me.txtAddress1.Size = New Size(216, 20)
        Me.txtAddress1.TabIndex = 7
        Me.txtAddress1.Text = ""
        '
        'lblAddress1
        '
        Me.lblAddress1.Location = New Point(24, 128)
        Me.lblAddress1.Name = "lblAddress1"
        Me.lblAddress1.Size = New Size(100, 16)
        Me.lblAddress1.TabIndex = 6
        Me.lblAddress1.Text = "Address 1"
        Me.lblAddress1.TextAlign = ContentAlignment.MiddleRight
        '
        'txtContractorName
        '
        Me.txtContractorName.Location = New Point(128, 64)
        Me.txtContractorName.MaxLength = 70
        Me.txtContractorName.Name = "txtContractorName"
        Me.txtContractorName.Size = New Size(344, 20)
        Me.txtContractorName.TabIndex = 3
        Me.txtContractorName.Text = ""
        '
        'lblContractorName
        '
        Me.lblContractorName.Location = New Point(24, 64)
        Me.lblContractorName.Name = "lblContractorName"
        Me.lblContractorName.Size = New Size(100, 16)
        Me.lblContractorName.TabIndex = 2
        Me.lblContractorName.Text = "Contractor Name"
        Me.lblContractorName.TextAlign = ContentAlignment.MiddleRight
        '
        'txtContractorNo
        '
        Me.txtContractorNo.Location = New Point(128, 32)
        Me.txtContractorNo.MaxLength = 10
        Me.txtContractorNo.Name = "txtContractorNo"
        Me.txtContractorNo.Size = New Size(88, 20)
        Me.txtContractorNo.TabIndex = 1
        Me.txtContractorNo.Text = ""
        '
        'lblContractorNo
        '
        Me.lblContractorNo.Location = New Point(24, 32)
        Me.lblContractorNo.Name = "lblContractorNo"
        Me.lblContractorNo.Size = New Size(100, 16)
        Me.lblContractorNo.TabIndex = 0
        Me.lblContractorNo.Text = "Contractor No."
        Me.lblContractorNo.TextAlign = ContentAlignment.MiddleRight
        '
        'btnAdd
        '
        Me.btnAdd.Location = New Point(16, 352)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.TabIndex = 1
        Me.btnAdd.Text = "Add"
        '
        'btnSave
        '
        Me.btnSave.Location = New Point(352, 352)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.TabIndex = 3
        Me.btnSave.Text = "Save"
        '
        'btnClose
        '
        Me.btnClose.Location = New Point(440, 352)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.TabIndex = 4
        Me.btnClose.Text = "Close"
        '
        'cbxCompNo
        '
        Me.cbxCompNo.DataSource = Logic.ContractorNameList
        Me.cbxCompNo.DropDownStyle = ComboBoxStyle.DropDownList
        Me.cbxCompNo.Location = New Point(144, 48)
        Me.cbxCompNo.Name = "cbxCompNo"
        Me.cbxCompNo.Size = New Size(344, 21)
        Me.cbxCompNo.TabIndex = 5
        '
        'btnEdit
        '
        Me.btnEdit.Location = New Point(104, 352)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "Edit"
        '
        'frmContractor
        '
        Me.AutoScaleBaseSize = New Size(5, 13)
        Me.ClientSize = New Size(530, 392)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.cbxCompNo)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnAdd)
        Me.Controls.Add(Me.grpContractorDetails)
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "frmContractor"
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.Text = "Contractor Maintenance"
        Me.grpContractorDetails.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public Sub New(ByVal caller As Object)
        MyBase.New()
        InitializeComponent()

        CallingForm = caller
    End Sub

    ''' <summary>
    ''' Converts Enter key to Tab key when pressed.
    ''' </summary>
    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, ByVal keyData As Keys) As Boolean
        If keyData = Keys.Enter Then
            SendKeys.Send("{Tab}")
            Return True
        End If

        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function

    Private Sub frmContractor_Closing(ByVal sender As Object, ByVal e As ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If Not IsNothing(CallingForm) Then
            CallingForm.Show()
        End If

        CallingForm = Nothing
    End Sub

    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
        Close()
    End Sub

    Private Sub frmContractor_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        AddDataBindings()
        grpContractorDetails.Enabled = False
    End Sub

    ' clears text from all text fields
    Private Sub ClearTextFields()
        txtContractorNo.Clear()
        txtContractorName.Clear()
        txtVATNo.Clear()
        txtAddress1.Clear()
        txtAddress2.Clear()
        txtAddress3.Clear()
        txtAddress4.Clear()
        txtPostalCode.Clear()
        txtTelNo.Clear()
    End Sub

    ' binds database data to text fields
    Private Sub AddDataBindings()
        txtContractorNo.DataBindings.Add("Text", Logic, "ContractorNumber", False, DataSourceUpdateMode.OnPropertyChanged)
        txtContractorName.DataBindings.Add("Text", Logic, "ContractorName", False, DataSourceUpdateMode.OnPropertyChanged)
        txtVATNo.DataBindings.Add("Text", Logic, "VatNumber", False, DataSourceUpdateMode.OnPropertyChanged)
        txtAddress1.DataBindings.Add("Text", Logic, "AddressLine1", False, DataSourceUpdateMode.OnPropertyChanged)
        txtAddress2.DataBindings.Add("Text", Logic, "AddressLine2", False, DataSourceUpdateMode.OnPropertyChanged)
        txtAddress3.DataBindings.Add("Text", Logic, "AddressLine3", False, DataSourceUpdateMode.OnPropertyChanged)
        txtAddress4.DataBindings.Add("Text", Logic, "AddressLine4", False, DataSourceUpdateMode.OnPropertyChanged)
        txtPostalCode.DataBindings.Add("Text", Logic, "PostalCode", False, DataSourceUpdateMode.OnPropertyChanged)
        txtTelNo.DataBindings.Add("Text", Logic, "Telephone", False, DataSourceUpdateMode.OnPropertyChanged)
    End Sub

    Private Sub btnAdd_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnAdd.Click
        If FormState = FormStates.Empty Then
            cbxCompNo.SendToBack()
            cbxCompNo.Enabled = False

            grpContractorDetails.Enabled = True
            ClearTextFields()

            FormState = FormStates.Add
            txtContractorNo.Focus()
        End If
    End Sub

    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnSave.Click
        If FormState = FormStates.Add Then
            If txtContractorNo.Text = String.Empty Then
                MsgBox("A Contractor Number is required", MsgBoxStyle.Critical, "Error")
                txtContractorNo.Focus()
            Else
                Dim count As New Integer

                Logic.GetCount(count)

                If count > 0 Then
                    MsgBox("Contractor Number entered is already used", MsgBoxStyle.Critical, "Error")
                    txtContractorNo.Focus()
                Else
                    Logic.AddRowToTable()
                    MsgBox("Record was successfully saved", MsgBoxStyle.Information, "Information")

                    cbxCompNo.BringToFront()
                    cbxCompNo.Enabled = True

                    grpContractorDetails.Enabled = False
                    FormState = FormStates.Empty
                End If
            End If
        End If

        If FormState = FormStates.Edit Then
            Logic.SaveEditToTable()
            MsgBox("Record was successfully saved", MsgBoxStyle.Information, "Information")

            cbxCompNo.BringToFront()
            cbxCompNo.Enabled = True

            grpContractorDetails.Enabled = False
            txtContractorNo.Enabled = True

            FormState = FormStates.Empty
        End If

        Logic.InitializeProperties(0)
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnEdit.Click
        If FormState = FormStates.Empty Then
            cbxCompNo.SendToBack()
            cbxCompNo.Enabled = False

            grpContractorDetails.Enabled = True
            txtContractorNo.Enabled = False

            FormState = FormStates.Edit
            txtContractorName.Focus()
        End If
    End Sub

    Private Sub cbxCompNo_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles cbxCompNo.SelectedIndexChanged
        If FormState = FormStates.Empty Then
            Logic.InitializeProperties(cbxCompNo.SelectedIndex)
        End If
    End Sub
End Class
