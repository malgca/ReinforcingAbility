Imports DataTier
Imports System.Data
Imports System.Drawing
Imports System.ComponentModel

Public Class CutInvoice
    Implements INotifyPropertyChanged

    Private Property CutInvoiceData As CutInvoiceData

    Public Property InvoiceCount As Integer
    Public Property InvoiceNumber As Long

    Private Sub GenerateCuttingSheetInvoice(ByVal cuttingSheetNumber As String, ByVal deliveryNoteNumber As String, ByVal orderNo As String, ByVal invoiceHeader As String, ByVal design As String, ByVal invoiceDate As Date)
        Dim FullInvoiceSet As New DataSet

        InvoiceCount = GetInvoiceCount(cuttingSheetNumber, design, FullInvoiceSet)

        If FullInvoiceSet.Tables(0).Rows(0).Item("InvoiceNo").ToString = "0" Or FullInvoiceSet.Tables(0).Rows(0).Item("InvoiceNo").ToString = String.Empty Then
            'This IF STATEMENT checks if the cutting sheet has not been invoiced.

            Try
                invoiceNumber = Long.Parse(FullInvoiceSet.Tables(0).Rows(0).Item("LastInvNum").ToString()) + 1
            Catch ex As Exception
                Throw ex ' rethrow exception so it can be handeld in presentation tier
            End Try

            Dim CalcTotal As Integer = 0
            Dim CalcVat As Integer = 0
            Dim CalcNett As Integer = 0

            Dim JobNo As String = FullInvoiceSet.Tables(0).Rows(0).Item("Job.JobNo").ToString() ' used once but requires FullInvoiceSet
            Dim VAT As String = FullInvoiceSet.Tables(0).Rows(0).Item("VatPerc").ToString() ' used multiple times

            Dim command As New OleDb.OleDbCommand(SQL4NewInvoice, DbConnection)
            Try
                DbConnection.Open()
                command.ExecuteNonQuery()

                'Create Invoice Lines

                Dim lcv As Integer
                Dim CurTypeCode As String
                Dim NextTypeCode As String
                Dim TotalLengthForType As Integer = 0
                Dim TypeMass As Integer = 0
                Dim TotalMassForType As Integer = 0
                Dim TotalCostForType As Integer = 0
                Dim LineNumberCounter As Integer = 1
                Dim DESCRIPTION As String = "Description"

                For lcv = 0 To FullInvoiceSet.Tables(0).Rows.Count - 1

                    CurTypeCode = FullInvoiceSet.Tables(0).Rows(lcv).Item("CutItem.TypeCode").ToString()
                    'MessageBox.Show("Current Type = " + CurTypeCode)

                    If lcv + 1 <= FullInvoiceSet.Tables(0).Rows.Count - 1 Then
                        NextTypeCode = FullInvoiceSet.Tables(0).Rows(lcv + 1).Item("CutItem.TypeCode").ToString()
                    Else
                        NextTypeCode = "NO NEXT"
                    End If

                    TotalLengthForType += Double.Parse(FullInvoiceSet.Tables(0).Rows(lcv).Item("Qty").ToString) * Double.Parse(FullInvoiceSet.Tables(0).Rows(lcv).Item("Length").ToString)

                    'If NextTypeCode <> CurTypeCode Or lcv = DataSet.Tables(0).Rows.Count - 1 Then
                    If NextTypeCode <> CurTypeCode Or NextTypeCode = "NO NEXT" Then
                        TotalLengthForType /= 1000
                        TypeMass = Double.Parse(FullInvoiceSet.Tables(0).Rows(lcv).Item("Weight").ToString())
                        If FullInvoiceSet.Tables(0).Rows(lcv).Item("Tons Or Kilograms").ToString() = "T" Then
                            TotalMassForType = TotalLengthForType * TypeMass / 1000
                            TotalMassForType = Decimal.Round(TotalMassForType, 3)
                            TotalCostForType = TotalMassForType * FullInvoiceSet.Tables(0).Rows(lcv).Item("Rate")
                        Else ' KG
                            TotalMassForType = TotalLengthForType * TypeMass
                            TotalMassForType = Decimal.Round(TotalMassForType, 1)
                            TotalCostForType = Math.Round(TotalMassForType, 1) * FullInvoiceSet.Tables(0).Rows(lcv).Item("Rate")
                        End If

                        Dim SQL4NewInvoiceLine As String = "INSERT INTO InvoiceLine(InvNo,[Line#],TypeCode,Description,Qty,TonsorKg,CostPerUnit,Total) VALUES " & _
                        "(" & InvoiceNumber & _
                        "," & LineNumberCounter & _
                        ",'" & CurTypeCode & _
                        "','" & DESCRIPTION & _
                        "'," & TotalMassForType & _
                        ",'" & DataSet.Tables(0).Rows(lcv).Item("Tons Or Kilograms").ToString() & _
                        "'," & DataSet.Tables(0).Rows(lcv).Item("Rate").ToString() & _
                        "," & TotalCostForType & _
                        ")"

                        LineNumberCounter += 1
                        CalcTotal += TotalCostForType

                        TotalLengthForType = 0
                        TypeMass = 0
                        TotalMassForType = 0
                        TotalCostForType = 0

                        Dim InvLineCommand As New OleDb.OleDbCommand(SQL4NewInvoiceLine, DbConnection)
                        Try
                            InvLineCommand.ExecuteNonQuery()
                            'MessageBox.Show("Executed : " + SQL4NewInvoiceLine)
                        Catch mex As Exception
                            MessageBox.Show(mex.Message)
                        End Try


                    End If
                Next lcv

                CalcNett = CalcTotal * (1 + Single.Parse(VAT))
                CalcVat = CalcTotal * Single.Parse(VAT)

                Dim SQL4Totals As String = "UPDATE Invoice SET InvTotal = " & Decimal.Round(CalcTotal, 2) & ", InvVatAmt = " & Decimal.Round(CalcVat, 2) & " , InvNett = " & Decimal.Round(CalcNett, 2) & " WHERE InvoiceNo = " & invoiceNumber
                command = New OleDb.OleDbCommand(SQL4Totals, DbConnection)

                Try
                    command.ExecuteNonQuery()
                Catch ex As Exception
                    MessageBox.Show(ex.Message)

                End Try

                'Update Company's Last Invoice Number
                Dim SQL4CompanyUpdate As String = "UPDATE Company SET Company.LastInvNum = " + InvoiceNumber.ToString() + " WHERE Company.CompanyNo = '" + DataSet.Tables(0).Rows(0).Item("Company.CompanyNo").ToString() + "'"
                Dim UpdateCommand As New OleDb.OleDbCommand(SQL4CompanyUpdate, DbConnection)
                Try
                    UpdateCommand.ExecuteNonQuery()
                Catch MEEE As Exception
                    MessageBox.Show(MEEE.Message)
                End Try

                'Update Cutting Sheet Invoice Number
                Dim SQL4CuttingSheetUpdate As String = "UPDATE CuttingSheet SET CuttingSheet.InvoiceNo = " + InvoiceNumber.ToString() + " WHERE CuttingSheet.CutSheetNo = " + DataSet.Tables(0).Rows(0).Item("CuttingSheet.CutSheetNo").ToString()
                Dim UpdateInvCommand As New OleDb.OleDbCommand(SQL4CuttingSheetUpdate, DbConnection)
                Try
                    UpdateInvCommand.ExecuteNonQuery()

                Catch MEE As Exception
                    MessageBox.Show(MEE.Message)
                End Try

                'MessageBox.Show("Cutting Sheet Number: " + CuttingSheetNumber.ToString + " has been successfully invoiced - Invoice Number: " + InvoiceNumber.ToString())
                popCMB()
                lblNotify.Text = "Invoice Number: " + InvoiceNumber.ToString()
                btnPrint.Enabled = True
            Catch Myerror As Exception
                MessageBox.Show(Myerror.Message)
            Finally
                DbConnection.Close()
            End Try
        Else
            MessageBox.Show("Cutting sheet " + cuttingSheetNumber + " has already been invoiced. See invoice no. " + DataSet.Tables(0).Rows(0).Item("InvoiceNo").ToString)
        End If
    End Sub

    Private Function GetInvoiceCount(ByRef cuttingSheetNumber As String, ByRef design As String, ByRef FullInvoiceSet As DataSet)
        If cuttingSheetNumber = String.Empty Or cuttingSheetNumber = "Please Select..." Then
            cuttingSheetNumber = "-1"
        End If

        If design = String.Empty Then
            design = "0"
        End If

        CutInvoiceData.PopulateFullInvoice(Long.Parse(cuttingSheetNumber))

        CutInvoiceData.Adapter.Fill(FullInvoiceSet)

        Return FullInvoiceSet.Tables(0).Rows.Count
    End Function

    Private Sub popCMB()
        cmb_AllCutSheets.Items.Clear()
        Dim sql As String = "SELECT CutSheetNo FROM CuttingSheet WHERE InvoiceNo = 0 ORDER BY CutSheetNo ASC"
        Dim DataSet As New Data.DataSet
        Dim Adapter As New OleDb.OleDbDataAdapter(sql, DbConnection)
        Adapter.Fill(DataSet)
        Dim d As Integer
        For d = 0 To DataSet.Tables(0).Rows.Count - 1
            cmb_AllCutSheets.Items.Add(DataSet.Tables(0).Rows(d).Item("CutSheetNo").ToString())
        Next d
    End Sub

    Private Sub cmb_AllCutSheets_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles cmb_AllCutSheets.SelectedIndexChanged
        Dim sql As String = "SELECT OrderNo FROM (CuttingSheet INNER JOIN Job ON CuttingSheet.[Job No] = Job.JobNo) WHERE CuttingSheet.CutSheetNo = " & cmb_AllCutSheets.Text
        Dim ds As New Data.DataSet
        Dim Adapter As New OleDb.OleDbDataAdapter(sql, DbConnection)
        Adapter.Fill(ds)

        If ds.Tables(0).Rows.Count <> 0 Then
            txtOrderNum.Text = ds.Tables(0).Rows(0).Item("OrderNo").ToString()
        End If

        lblNotify.Text = ""
        btnPrint.Enabled = False

    End Sub

    Public Event PropertyChanged(sender As Object, e As PropertyChangedEventArgs) Implements INotifyPropertyChanged.PropertyChanged

    ''' <summary>
    ''' Notifies listener of a change in a property
    ''' </summary>
    Private Sub NotifyPropertyChanged(ByVal ParamArray Properties() As String)
        For Each Prop As String In Properties
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(Prop))
        Next
    End Sub

    ' Checks if a DB field is not equal to null
    Private Function IsNotNull(ByRef param As Object) As Boolean
        Return Not param.Equals(DBNull.Value)
    End Function
End Class
