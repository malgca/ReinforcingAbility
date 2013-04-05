Imports System.Data.Oledb
Public Class frmNotInvoiced
    Inherits System.Windows.Forms.Form
    Shared cnnReinforcing As New _
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
    Friend WithEvents ListView1 As System.Windows.Forms.ListView
    Friend WithEvents btnClose As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btnClose = New System.Windows.Forms.Button
        Me.ListView1 = New System.Windows.Forms.ListView
        Me.SuspendLayout()
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(152, 224)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(72, 32)
        Me.btnClose.TabIndex = 3
        Me.btnClose.Text = "Close"
        '
        'ListView1
        '
        Me.ListView1.ForeColor = System.Drawing.SystemColors.InactiveCaption
        Me.ListView1.GridLines = True
        Me.ListView1.Location = New System.Drawing.Point(24, 24)
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(296, 168)
        Me.ListView1.TabIndex = 5
        '
        'frmNotInvoiced
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(384, 266)
        Me.Controls.Add(Me.ListView1)
        Me.Controls.Add(Me.btnClose)
        Me.Name = "frmNotInvoiced"
        Me.Text = "Cutting Sheets Not Invoiced"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private CallingForm As Object

    Public Sub New(ByVal caller As Object)
        MyBase.New()
        InitializeComponent()

        CallingForm = caller
    End Sub

    Private Sub frmNotInvoiced_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        CuttingSheet.Initialize()
        ListView1.View = View.Details
        Dim Col As New ListView.ColumnHeaderCollection(ListView1)
        Col.Clear()
        Col.Add("Cut Sheet", 80, HorizontalAlignment.Right)
        Col.Add("Date", 80, HorizontalAlignment.Right)
        Col.Add("Invoice No", 80, HorizontalAlignment.Right)

        Dim dsCutSheet As New DataSet
        Dim cutDetails As ListViewItem
        Dim cutDate As Date
        Dim cutSheet, invoiceNo, cutJob As String
        Dim sql As String = "SELECT * FROM Cuttingsheet WHERE invoiceNo = 0"

        Try
            Dim adpCutSheet As New _
                       OleDbDataAdapter(sql, cnnReinforcing)
            adpCutSheet.Fill(dsCutSheet, "Cuttingsheet")
            If dsCutSheet.Tables("CuttingSheet").Rows.Count > 0 Then
                Dim dsRow As DataRow
                For Each dsRow In dsCutSheet.Tables("CuttingSheet").Rows
                    cutSheet = dsRow("CutSheetNo")
                    invoiceNo = dsRow("InvoiceNo")
                    'cutJob = dsRow("[Job No]")
                    cutDetails = ListView1.Items.Add(cutSheet)
                    cutDate = dsRow("cutDate")
                    cutDetails.SubItems.Add(cutDate)
                    cutDetails.SubItems.Add(invoiceNo)
                    cutDetails.SubItems.Add(cutJob)
                Next
            End If
            dsCutSheet = Nothing
        Catch ex As Exception
            MessageBox.Show(ex.ToString())
        End Try

    End Sub
    Private Sub frmAddClient_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If Not IsNothing(CallingForm) Then
            CallingForm.Show()
        End If

        CallingForm = Nothing
    End Sub
    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click

        CuttingSheet.Terminate()
        Close()

    End Sub

End Class
