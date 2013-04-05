Public Class frmInvoicing
    Inherits System.Windows.Forms.Form
    Private CallingForm As Object
    Dim dbconnection As OleDb.OleDbConnection

#Region " Windows Form Designer generated code "

    Public Sub New(ByVal caller As Object, ByRef dbc As OleDb.OleDbConnection)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        CallingForm = caller
        dbConnection = dbc
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
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btnCSInv As System.Windows.Forms.Button
    Friend WithEvents btnMInv As System.Windows.Forms.Button
    Friend WithEvents btnSInv As System.Windows.Forms.Button
    Friend WithEvents lblHeading As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents btnInvCancel As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btnExit = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.btnCSInv = New System.Windows.Forms.Button
        Me.btnMInv = New System.Windows.Forms.Button
        Me.btnSInv = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.lblHeading = New System.Windows.Forms.Label
        Me.btnInvCancel = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnExit
        '
        Me.btnExit.Location = New System.Drawing.Point(168, 416)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(168, 40)
        Me.btnExit.TabIndex = 10
        Me.btnExit.Text = "Back to Main Menu"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btnCSInv)
        Me.GroupBox1.Controls.Add(Me.btnMInv)
        Me.GroupBox1.Controls.Add(Me.btnSInv)
        Me.GroupBox1.Controls.Add(Me.Button1)
        Me.GroupBox1.Controls.Add(Me.btnInvCancel)
        Me.GroupBox1.Location = New System.Drawing.Point(16, 112)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(480, 224)
        Me.GroupBox1.TabIndex = 8
        Me.GroupBox1.TabStop = False
        '
        'btnCSInv
        '
        Me.btnCSInv.Location = New System.Drawing.Point(32, 32)
        Me.btnCSInv.Name = "btnCSInv"
        Me.btnCSInv.Size = New System.Drawing.Size(168, 40)
        Me.btnCSInv.TabIndex = 2
        Me.btnCSInv.Text = "Cutting Sheet Invoice"
        '
        'btnMInv
        '
        Me.btnMInv.Location = New System.Drawing.Point(32, 96)
        Me.btnMInv.Name = "btnMInv"
        Me.btnMInv.Size = New System.Drawing.Size(168, 40)
        Me.btnMInv.TabIndex = 3
        Me.btnMInv.Text = "Mesh Invoice"
        '
        'btnSInv
        '
        Me.btnSInv.Location = New System.Drawing.Point(280, 32)
        Me.btnSInv.Name = "btnSInv"
        Me.btnSInv.Size = New System.Drawing.Size(168, 40)
        Me.btnSInv.TabIndex = 3
        Me.btnSInv.Text = "Sundry Invoice"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(280, 96)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(168, 40)
        Me.Button1.TabIndex = 3
        Me.Button1.Text = "Escalation Invoice"
        '
        'lblHeading
        '
        Me.lblHeading.BackColor = System.Drawing.Color.White
        Me.lblHeading.Dock = System.Windows.Forms.DockStyle.Top
        Me.lblHeading.Font = New System.Drawing.Font("Times New Roman", 27.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHeading.ForeColor = System.Drawing.Color.RoyalBlue
        Me.lblHeading.Location = New System.Drawing.Point(0, 0)
        Me.lblHeading.Name = "lblHeading"
        Me.lblHeading.Size = New System.Drawing.Size(514, 80)
        Me.lblHeading.TabIndex = 11
        Me.lblHeading.Text = "Invoicing"
        Me.lblHeading.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnInvCancel
        '
        Me.btnInvCancel.Location = New System.Drawing.Point(280, 160)
        Me.btnInvCancel.Name = "btnInvCancel"
        Me.btnInvCancel.Size = New System.Drawing.Size(168, 40)
        Me.btnInvCancel.TabIndex = 3
        Me.btnInvCancel.Text = "Invoice Cancellation"
        '
        'frmInvoicing
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(514, 480)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.lblHeading)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "frmInvoicing"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Create an Invoice"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnCSInv_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCSInv.Click
        Dim Form As GenCutSheetInvoice = New GenCutSheetInvoice(Me, dbconnection)
        Me.Hide()
        Form.Show()
    End Sub



    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Close()
        CallingForm.show()
        CallingForm = Nothing
    End Sub

    Private Sub btnMInv_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMInv.Click
        Dim form As frmMeshInvoice = New frmMeshInvoice(Me, DBConnection)
        Me.Hide()
        form.Show()
    End Sub

    Private Sub btnSInv_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSInv.Click
        Dim Form As New frmSundryInv(Me, dbconnection)
        Me.Hide()
        Form.Show()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim eForm As frmEscalation = New frmEscalation(Me, dbconnection)
        Me.Hide()
        eForm.Show()
    End Sub

    Private Sub btnInvCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInvCancel.Click
        Dim cForm As InvoiceCancel = New InvoiceCancel(Me, dbconnection)
        Me.Hide()
        cForm.Show()

    End Sub
End Class
