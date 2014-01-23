Imports DataTier
Imports System.Data
Imports System.Drawing
Imports System.ComponentModel

Public Class CutInvoice
    Implements INotifyPropertyChanged

    Private Sub generateCuttingSheetInvoice(ByVal CuttingSheetNumber As String, ByVal DeliveryNoteNumber As String, ByVal OrderNo As String, ByVal InvoiceHeader As String, ByVal Design As String, ByVal InvoiceDate As Date)

        If CuttingSheetNumber = "" Or CuttingSheetNumber = "Please Select..." Then
            CuttingSheetNumber = "-1"
        End If
        If Design = "" Then
            Design = "0"
        End If

        Dim SQL As String = "SELECT * FROM ((((((CuttingSheet INNER JOIN Job ON [CuttingSheet].[Job No] = Job.JobNo ) INNER JOIN Company ON Job.CompanyNo = Company.CompanyNo) INNER JOIN SchedItem ON CuttingSheet.CutSheetNo = SchedItem.CutSheetNo) INNER JOIN CutItem ON SchedItem.CutSheetNo = CutItem.CutSheetNo AND SchedItem.ScheduleNo = CutItem.ScheduleNo) INNER JOIN ProductType ON CutItem.TypeCode = ProductType.TypeCode) INNER JOIN JobRate ON Job.JobNo = JobRate.JobNo AND JobRate.TypeCode = ProductType.TypeCode ) WHERE CuttingSheet.CutSheetNo = " + CuttingSheetNumber + " ORDER BY ProductType.TypeCode"
        Dim DataSet As New Data.DataSet
        Dim Adapter As New OleDb.OleDbDataAdapter(SQL, DbConnection)
        Adapter.Fill(DataSet)

        If DataSet.Tables(0).Rows.Count = 0 Then
            If CuttingSheetNumber = "-1" Then
                MessageBox.Show("Please enter a valid Cutting Sheet Number.", "Invalid Cutting Sheet Number.")
            Else
                MessageBox.Show("Cannot generate Invoice. Possible reason :" & Microsoft.VisualBasic.Chr(13) & "1. Rates are not specified for the job." & Microsoft.VisualBasic.Chr(13) & "2. The cutting sheet has no Schedules and/or Items.")
            End If

        Else
            If DataSet.Tables(0).Rows(0).Item("InvoiceNo").ToString = "0" Or DataSet.Tables(0).Rows(0).Item("InvoiceNo").ToString = "" Then
                'This IF STATEMENT checks if the cutting sheet has not been invoiced.

                Try
                    InvoiceNumber = Long.Parse(DataSet.Tables(0).Rows(0).Item("LastInvNum").ToString()) + 1
                Catch ex As Exception
                    MessageBox.Show("Gotcha!")
                End Try

                Dim InvoiceType As String = "Cutting Sheet"
                Dim Factor As Int16 = 1.5
                Dim EscMoAndDa As Date = New Date(1999, 12, 11)
                Dim Work As Int16 = 75
                Dim RefNo As String = "Ref"
                Dim JobNo As String = DataSet.Tables(0).Rows(0).Item("Job.JobNo").ToString()
                Dim VAT As String = DataSet.Tables(0).Rows(0).Item("VatPerc").ToString()
                Dim Active As String = "Yes"
                Dim Escalated As String = "No"
                Dim OnSummary As String = "Yes"
                Dim Comments As String = "Comments"


                Dim CalcTotal As Integer = 0
                Dim CalcVat As Integer = 0
                Dim CalcNett As Integer = 0

                Dim SQL4NewInvoice As String = "INSERT INTO Invoice(InvoiceNo,InvoiceType,InvDate,InvDeliveryNoteNo,InvFactor,Invmonthandyear,InvWork,InvOrdNum,InvRefNum,InvoiceHeading,InvTotal,InvVatAmt,InvDesign,InvNett,InvJobNo,InvActive,InvEscalated,InvOnSummary,InvComments) VALUES " & _
                "(  " & _
                    InvoiceNumber.ToString & _
                    ",'" & InvoiceType & _
                    "',#" & InvoiceDate.ToLongDateString & _
                    "#,'" & DeliveryNoteNumber & _
                    "'," & Factor & _
                    ",#" & EscMoAndDa & _
                    "#," & Work & _
                    ",'" & OrderNo & _
                    "','" & RefNo & _
                    "','" & InvoiceHeader & _
                    "'," & "-1" & _
                    "," & "-1" & _
                    "," & Design & _
                    "," & "-1" & _
                    ",'" & JobNo & _
                    "'," & Active & _
                    "," & Escalated & _
                    "," & OnSummary & _
                    ",'" & Comments & _
                    "')"


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

                    For lcv = 0 To DataSet.Tables(0).Rows.Count - 1

                        CurTypeCode = DataSet.Tables(0).Rows(lcv).Item("CutItem.TypeCode").ToString()
                        'MessageBox.Show("Current Type = " + CurTypeCode)

                        If lcv + 1 <= DataSet.Tables(0).Rows.Count - 1 Then
                            NextTypeCode = DataSet.Tables(0).Rows(lcv + 1).Item("CutItem.TypeCode").ToString()
                        Else
                            NextTypeCode = "NO NEXT"
                        End If

                        TotalLengthForType += Double.Parse(DataSet.Tables(0).Rows(lcv).Item("Qty").ToString) * Double.Parse(DataSet.Tables(0).Rows(lcv).Item("Length").ToString)

                        'If NextTypeCode <> CurTypeCode Or lcv = DataSet.Tables(0).Rows.Count - 1 Then
                        If NextTypeCode <> CurTypeCode Or NextTypeCode = "NO NEXT" Then
                            TotalLengthForType /= 1000
                            TypeMass = Double.Parse(DataSet.Tables(0).Rows(lcv).Item("Weight").ToString())
                            If DataSet.Tables(0).Rows(lcv).Item("Tons Or Kilograms").ToString() = "T" Then
                                TotalMassForType = TotalLengthForType * TypeMass / 1000
                                TotalMassForType = Decimal.Round(TotalMassForType, 3)
                                TotalCostForType = TotalMassForType * DataSet.Tables(0).Rows(lcv).Item("Rate")
                            Else ' KG
                                TotalMassForType = TotalLengthForType * TypeMass
                                TotalMassForType = Decimal.Round(TotalMassForType, 1)
                                TotalCostForType = Math.Round(TotalMassForType, 1) * DataSet.Tables(0).Rows(lcv).Item("Rate")
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

                    Dim SQL4Totals As String = "UPDATE Invoice SET InvTotal = " & Decimal.Round(CalcTotal, 2) & ", InvVatAmt = " & Decimal.Round(CalcVat, 2) & " , InvNett = " & Decimal.Round(CalcNett, 2) & " WHERE InvoiceNo = " & InvoiceNumber
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
                MessageBox.Show("Cutting sheet " + CuttingSheetNumber + " has already been invoiced. See invoice no. " + DataSet.Tables(0).Rows(0).Item("InvoiceNo").ToString)
            End If

        End If
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
