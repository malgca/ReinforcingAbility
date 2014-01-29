Imports System
Imports System.Globalization
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Common

''' <summary>
''' CutInvoice database operations
''' </summary>
Public Class CutInvoiceData
    Public Property Adapter As New OleDbDataAdapter

    Private Property FullInvoiceCommand As New OleDbCommand
    Private Property InsertInvoiceCommand As New OleDbCommand

    ''' <summary>
    ''' Retrieves information from a number of tables in order to populate a full invoice.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub PopulateFullInvoice(ByRef cuttingSheetNumber As Long)
        Me.Adapter.SelectCommand = Me.FullInvoiceCommand

        Me.FullInvoiceCommand.CommandText = "SELECT * FROM ((((((CuttingSheet " & _
            "INNER JOIN Job ON [CuttingSheet].[Job No] = Job.JobNo) " & _
            "INNER JOIN Company ON Job.CompanyNo = Company.CompanyNo) " & _
            "INNER JOIN SchedItem ON CuttingSheet.CutSheetNo = SchedItem.CutSheetNo) " & _
            "INNER JOIN CutItem ON SchedItem.CutSheetNo = CutItem.CutSheetNo AND SchedItem.ScheduleNo = CutItem.ScheduleNo) " & _
            "INNER JOIN ProductType ON CutItem.TypeCode = ProductType.TypeCode) " & _
            "INNER JOIN JobRate ON Job.JobNo = JobRate.JobNo AND JobRate.TypeCode = ProductType.TypeCode ) " & _
            "WHERE CuttingSheet.CutSheetNo = ? " & _
            "ORDER BY ProductType.TypeCode"

        Dim newParam As New OleDbParameter()
        newParam.SourceColumn = "CuttingSheet.CutSheetNo"
        newParam.OleDbType = OleDbType.Numeric
        newParam.Value = cuttingSheetNumber
        Me.FullInvoiceCommand.Parameters.Add(newParam)
    End Sub

    Public Sub InsertInvoice(ByRef JobNo As String, ByRef VAT As String, ByVal invoiceNumber As Long, ByVal invoiceDate As Date, ByRef design As String, ByRef deliveryNoteNumber As String, ByRef orderNo As String, ByRef invoiceHeader As String)
        Me.Adapter.InsertCommand = Me.InsertInvoiceCommand

        Me.InsertInvoiceCommand.CommandText = "INSERT INTO Invoice" & _
            "(InvoiceNo, InvoiceType, InvDate, InvDeliveryNoteNo, InvFactor, Invmonthandyear, InvWork, InvOrdNum, InvRefNum, InvoiceHeading, InvTotal, InvVatAmt, InvDesign, InvNett, InvJobNo, InvActive, InvEscalated, InvOnSummary, InvComments) " & _
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"

        Dim Active As String = "Yes"
        Dim Escalated As String = "No"
        Dim OnSummary As String = "Yes"
        Dim Comments As String = "Comments"

"','" & invoiceHeader & _
            "'," & "-1" & _
            "," & "-1" & _
            "," & design & _
            "," & "-1" & _
            ",'" & JobNo & _
            "'," & Active & _
            "," & Escalated & _
            "," & OnSummary & _
            ",'" & Comments & _
            "')"

        Dim newParam As New OleDbParameter()
        newParam.SourceColumn = "InvoiceNo"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 10
        newParam.Value = invoiceNumber.ToString()
        Me.InsertInvoiceCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "InvoiceType"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 10
        newParam.Value = "Cutting Sheet"
        Me.InsertInvoiceCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "InvDate"
        newParam.OleDbType = OleDbType.Date
        newParam.Value = "'#" & invoiceDate.ToLongDateString & "#'"
        Me.InsertInvoiceCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "InvDeliveryNoteNo"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 10
        newParam.Value = deliveryNoteNumber
        Me.InsertInvoiceCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "InvFactor"
        newParam.OleDbType = OleDbType.Numeric
        newParam.Value = 1.5
        Me.InsertInvoiceCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = " Invmonthandyear"
        newParam.OleDbType = OleDbType.Date
        newParam.Value = "#" & New Date(1999, 12, 11) & "#"
        Me.InsertInvoiceCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = " InvWork"
        newParam.OleDbType = OleDbType.Numeric
        newParam.Value = 75
        Me.InsertInvoiceCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "InvOrdNum"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 10
        newParam.Value = orderNo
        Me.InsertInvoiceCommand.Parameters.Add(newParam)

        newParam = New OleDbParameter()
        newParam.SourceColumn = "InvRefNum"
        newParam.OleDbType = OleDbType.VarWChar
        newParam.Size = 10
        newParam.Value = "Ref"
        Me.InsertInvoiceCommand.Parameters.Add(newParam)


    End Sub
End Class
