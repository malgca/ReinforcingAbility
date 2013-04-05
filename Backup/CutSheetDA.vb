' Chapter 13- Example6
' DA classes using dbms
' 
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections
Public Class CutSheetDA
    ' Declare a connection
    Shared cnnReinforcing As New _
        OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; " & _
        "Data Source=winsteelVers5.mdb")
    '-------------------------------------------------------------------
   
    '-------------------------------------------------------------------
    ' Declare a new ArrayList instance
    Shared cutSheet As New ArrayList
    ' Declare an instance of Customer
    Shared aCutSheet As CuttingSheet
    Shared cutSheetNo, invoiceNo As String
    Shared cutDate As Date

    ' Initialize Method
    Public Shared Sub Initialize()

        Try
            ' Try to open the connection
            cnnReinforcing.Open()
        Catch e As Exception
            Console.WriteLine(e.ToString)
        End Try
    End Sub

    ' Terminate Method
    Public Shared Sub Terminate()

        Try
            cnnReinforcing.Close()
            cnnReinforcing = Nothing
        Catch e As Exception
            Console.WriteLine(e.Message.ToString)
        End Try
    End Sub

    ' AddNew Method --Throws DuplicateException if exists
    ' GetAll Method
    Public Shared Function GetAll() As ArrayList
        Dim dsCutSheet As New DataSet
        Dim sqlQuery As String = "SELECT CuttingSheet, InvoiceNo, cutDate " & _
            "FROM CuttingSheet"

        Try
            Dim adpCutSheet As New _
                OleDbDataAdapter(sqlQuery, cnnReinforcing)
            adpCutSheet.Fill(dsCutSheet, "CuttingSheet")
            If dsCutSheet.Tables("CustTable").Rows.Count > 0 Then
                Dim dsRow As DataRow
                ' Clear the array list
                cutSheet.Clear()
                For Each dsRow In dsCutSheet.Tables("CuttingSheet").Rows
                    cutSheetNo = dsRow("CutSheetNo")
                    invoiceNo = dsRow("invoiceNo")
                    cutDate = dsRow("cutDate")
                    MessageBox.Show("cutsheetNo")
                Next
            Else
                ' No records in database
            End If
            dsCutSheet = Nothing
        Catch e As Exception
            Console.WriteLine(e.ToString)
        End Try
        Return cutSheet
    End Function

End Class
