Imports DataTier
Imports System.Data
Imports System.ComponentModel

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

    Private _vat As Double

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
    Public Property VAT As Double
        Get
            Return _vat
        End Get
        Private Set(value As Double)
            If value <> _vat Then
                _vat = value
                NotifyPropertyChanged("VAT")
            End If
        End Set
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
    Public Property CompanyNameList As New List(Of String)

    Private Property CompanyData As CompanyData
    Private Property CompanySet As New DataSet

    Public Sub New()
        CompanyData = New CompanyData(CompanyNumber)

        InitializeProperties(0)
    End Sub

    ''' <summary>
    ''' Update company property parameters
    ''' </summary>
    Public Sub InitializeProperties(ByVal index As Integer)
        CompanyData.Adapter.Fill(CompanySet)

        Dim row = CompanySet.Tables.Item(0).Rows.Item(index)

        ' map properties to database fields
        If IsNotNull(row("CompanyNo")) Then
            CompanyNumber = row("CompanyNo")
        End If

        If IsNotNull(row("RegNo")) Then
            RegNumber = row("RegNo")
        End If

        If IsNotNull(row("VatNo")) Then
            VatNumber = row("VatNo")
        End If

        If IsNotNull(row("CompanyName")) Then
            CompanyName = row("CompanyName")
        End If

        If IsNotNull(row("Address")) Then
            AddressLine1 = row("Address")
        End If

        If IsNotNull(row("AddressLine2")) Then
            AddressLine2 = row("AddressLine2")
        End If

        If IsNotNull(row("AddressLine3")) Then
            AddressLine3 = row("AddressLine3")
        End If

        If IsNotNull(row("AddressLine4")) Then
            AddressLine4 = row("AddressLine4")
        End If

        If IsNotNull(row("PostalCode")) Then
            PostalCode = row("PostalCode")
        End If

        If IsNotNull(row("Telephone")) Then
            Telephone = row("Telephone")
        End If

        If IsNotNull(row("Fax")) Then
            Fax = row("Fax")
        End If

        If IsNotNull(row("Message")) Then
            Message = row("Message")
        End If

        'If IsNotNull(row("Email")) Then
        '    Email = row("Email")
        'End If

        'If IsNotNull(row("Website")) Then
        '    Website = row("Website")
        'End If

        If IsNotNull(row("VatPerc")) Then
            VAT = row("VatPerc")
        End If

        If IsNotNull(row("LastInvNum")) Then
            LastInvoiceNumber = row("LastInvNum")
        End If

        If IsNotNull(row("LastCutNum")) Then
            LastCuttingSheetNumber = row("LastCutNum")
        End If

        'If IsNotNull(row("UnitOfMeas")) Then
        '    UnitOfMeasurement = row("UnitOfMeas")
        'End If

        If IsNotNull(row("CompanyNo")) And IsNotNull(row("CompanyName")) Then
            For Each newRow As DataRow In CompanySet.Tables.Item(0).Rows
                CompanyNameList.Add(String.Format("[{0}] {1}", newRow("CompanyNo"), newRow("CompanyName")))
            Next
        End If
    End Sub

    ''' <summary>
    ''' Get the number of companies in the company table in the database
    ''' </summary>
    Public Sub GetCount(ByRef count As Integer)
        ' get the number of available companies in the database
        CompanyData.GetNumberOfCompanies(CompanyNumber, count)
    End Sub

    ''' <summary>
    ''' Adds a row to the company table
    ''' </summary>
    Public Sub AddRowToTable()
        ' update the company table with data currently in the company fields
        CompanyData.AddRow(CompanyNumber, CompanyName, RegNumber, VatNumber, AddressLine1, AddressLine2, AddressLine3, AddressLine4, PostalCode, Telephone, Email, Fax, Website, Message, VAT, LastInvoiceNumber, LastCuttingSheetNumber)
    End Sub

    ''' <summary>
    ''' Save an edit to the company table
    ''' </summary>
    Public Sub SaveEditToTable()
        ' save the editted row to the table
        CompanyData.SaveRowEdit(CompanyNumber, CompanyName, RegNumber, VatNumber, AddressLine1, AddressLine2, AddressLine3, AddressLine4, PostalCode, Telephone, Email, Fax, Website, Message, VAT, LastInvoiceNumber, LastCuttingSheetNumber)
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