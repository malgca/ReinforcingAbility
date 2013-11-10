Imports System
Imports System.Data
Imports System.Data.Common

Public Class DBOperations
    ' Private connection variable
    Private _connection As OleDbConnection

    Public Property Connection As OleDbConnection
        Get
            Return _connection
        End Get
        Private Set(value As OleDbConnection)
            If (Not (value.Equals(_connection))) Then
                _connection = value
            End If
        End Set
    End Property

    Private Sub New()
        'initiate connectionstring
        Me._connection = New OleDbConnection("Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Registry Path=;Jet OLEDB:Database Locking Mode=1;Data Source=""winsteelVers5.mdb"";Mode=Share Deny None;Jet OLEDB:Engine Type=5;Provider=""Microsoft.Jet.OLEDB.4.0"";Jet OLEDB:System database=;Jet OLEDB:SFP=False;persist security info=False;Extended Properties=;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Create System Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;User ID=Admin;Jet OLEDB:Global Bulk Transactions=1")

        ' Make sure only a single instance of this class may exist
    End Sub

    ''' <summary>
    ''' Get the running instance of the DBOperations class
    ''' </summary>
    ''' <returns>Singleton instance of DBOperations class</returns>
    Public Shared ReadOnly Property GetInstance As DBOperations
        Get
            Static Instance As DBOperations = New DBOperations
            Return Instance
        End Get
    End Property
End Class