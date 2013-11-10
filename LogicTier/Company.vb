Imports System
Imports System.Data
Imports DataTier
Imports System.Data.OleDb
Imports System.ComponentModel

''' <summary>
''' Company Form States
''' </summary>
Public Enum FormStates
    Empty = 0
    Add = 2
    Edit = 4
End Enum

''' <summary>
''' Handles logical operations related to the company
''' </summary>
Public Class Company
    Implements INotifyPropertyChanged

    Private _companyNumber,
        _regNumber,
        _vatNumber,
        _companyName,
        _addressLine1,
        _addressLine2,
        _addressLine3,
        _addressLine4,
        _telephone,
        _fax,
        _message,
        _email,
        _website As String

    Private _postalCode,
        _lastInvoiceNumber,
        _lastCuttingSheetNumber As Integer

    Private _unitOfMeasurement As MeasurementUnit

    Public Property CompanyNumber As String
        Get
            Return _companyNumber
        End Get
        Set(value As String)
            If _companyNumber <> value Then
                _companyNumber = value
                NotifyPropertyChanged("CompanyNumber")
            End If
        End Set
    End Property
    Public Property RegNumber As String
        Get
            Return _regNumber
        End Get
        Set(value As String)
            If (value <> _regNumber) Then
                _regNumber = value
                NotifyPropertyChanged("RegNumber")
            End If
        End Set
    End Property
    Public Property VatNumber As String
        Get
            Return _vatNumber
        End Get
        Set(value As String)
            If (value <> _vatNumber) Then
                _vatNumber = value
                NotifyPropertyChanged("VatNumber")
            End If
        End Set
    End Property
    Public Property CompanyName As String
        Get
            Return _companyName
        End Get
        Set(value As String)
            If (value <> _companyName) Then
                _companyName = value
                NotifyPropertyChanged("CompanyName")
            End If
        End Set
    End Property
    Public Property AddressLine1 As String
        Get
            Return _addressLine1
        End Get
        Set(value As String)
            If value <> _addressLine1 Then
                _addressLine1 = value
                NotifyPropertyChanged("AddressLine1")
            End If
        End Set
    End Property
    Public Property AddressLine2 As String
        Get
            Return _addressLine2
        End Get
        Set(value As String)
            If value <> _addressLine2 Then
                _addressLine2 = value
                NotifyPropertyChanged("AddressLine2")
            End If
        End Set
    End Property
    Public Property AddressLine3 As String
        Get
            Return _addressLine3
        End Get
        Set(value As String)
            If value <> _addressLine3 Then
                _addressLine3 = value
                NotifyPropertyChanged("AddressLine3")
            End If
        End Set
    End Property
    Public Property AddressLine4 As String
        Get
            Return _addressLine4
        End Get
        Set(value As String)
            If value <> _addressLine4 Then
                _addressLine4 = value
                NotifyPropertyChanged("AddressLine4")
            End If
        End Set
    End Property
    Public Property PostalCode As Integer
        Get
            Return _postalCode
        End Get
        Set(value As Integer)
            If value <> _postalCode Then
                _postalCode = value
                NotifyPropertyChanged("PostalCode")
            End If
        End Set
    End Property
    Public Property Telephone As String
        Get
            Return _telephone
        End Get
        Set(value As String)
            If value <> _telephone Then
                _telephone = value
                NotifyPropertyChanged("Telephone")
            End If
        End Set
    End Property
    Public Property Fax As String
        Get
            Return _fax
        End Get
        Set(value As String)
            If value <> _fax Then
                _fax = value
                NotifyPropertyChanged("Fax")
            End If
        End Set
    End Property
    Public Property Message As String
        Get
            Return _message
        End Get
        Set(value As String)
            If value <> _message Then
                _message = value
                NotifyPropertyChanged("Message")
            End If
        End Set
    End Property
    Public Property Email As String
        Get
            Return _email
        End Get
        Set(value As String)
            If value <> _email Then
                _email = value
                NotifyPropertyChanged("Email")
            End If
        End Set
    End Property
    Public Property Website As String
        Get
            Return _website
        End Get
        Set(value As String)
            If value <> _website Then
                _website = value
                NotifyPropertyChanged("Website")
            End If
        End Set
    End Property
    Public ReadOnly Property VAT As Double
        Get
            Return 0.14
        End Get
    End Property
    Public Property LastInvoiceNumber As Integer
        Get
            Return _lastInvoiceNumber
        End Get
        Set(value As Integer)
            If value <> _lastInvoiceNumber Then
                _lastInvoiceNumber = value
                NotifyPropertyChanged("LastInvoiceNumber")
            End If
        End Set
    End Property
    Public Property LastCuttingSheetNumber As Integer
        Get
            Return _lastCuttingSheetNumber
        End Get
        Set(value As Integer)
            If value <> _lastCuttingSheetNumber Then
                _lastCuttingSheetNumber = value
                NotifyPropertyChanged("LastCuttingSheetNumber")
            End If
        End Set
    End Property
    Public Property UnitOfMeasurement As MeasurementUnit
        Get
            Return _unitOfMeasurement
        End Get
        Set(value As MeasurementUnit)
            If value <> _unitOfMeasurement Then
                _unitOfMeasurement = value
                NotifyPropertyChanged("UnitOfMeasurement")
            End If
        End Set
    End Property

    Public Sub New()
        InitializeCompanyProperties()
    End Sub

    ' Initialize company parameters
    Private Sub InitializeCompanyProperties()
        Dim companyData As CompanyData
        companyData = New CompanyData()

        Dim companySet As New DataSet
        companyData.Adapter.Fill(companySet)

        Dim tableData = companySet.Tables.Item(0).Rows.Item(0)

        Console.WriteLine(companySet.Tables.Item(0).Constraints)
        Console.WriteLine(companySet.Tables.Item(0).PrimaryKey.Length)
        Console.WriteLine(companySet.Tables.Item(0).PrimaryKey.GetLength(0))
        Console.WriteLine(companySet.Tables.Item(0).Columns)
    End Sub

    Public Sub LinkAdapter(ByRef adapter As OleDbDataAdapter)
        Dim companyData As CompanyData
        companyData = New CompanyData()
        adapter = companyData.Adapter
    End Sub

    Public Event PropertyChanged(sender As Object, e As PropertyChangedEventArgs) Implements INotifyPropertyChanged.PropertyChanged

    ''' <summary>
    ''' Notifies listener of a change in a property
    ''' </summary>
    Public Sub NotifyPropertyChanged(ByVal ParamArray Properties() As String)
        For Each Prop As String In Properties
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(Prop))
        Next
    End Sub
End Class

