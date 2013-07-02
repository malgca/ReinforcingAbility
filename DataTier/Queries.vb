Imports System.Text

Namespace DataTier
    Public Class Queries
        Private queryBuilder As New StringBuilder
        ''' <summary>
        ''' Basic SQL Select statement.
        ''' </summary>
        ''' <param name="queryCriteria">Parameter on which to query the database.</param>
        ''' <param name="table">Table to be queried</param>
        ''' <param name="order">Parameter on which to order the table.</param>
        ''' <remarks></remarks>
        Public Function SQLSelectQuery(ByRef queryCriteria, ByRef table, ByRef order) As String
            queryBuilder.Append("Select " & queryCriteria & " FROM " & table & " ORDER BY " & order)
            Return queryBuilder.ToString()
        End Function
    End Class
End Namespace
