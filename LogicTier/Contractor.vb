Imports DataTier
Imports System.Data
Imports System.ComponentModel

Public Class Contractor
    Implements INotifyPropertyChanged

    Private _contractorNumber,
        _contractorName,
        _addressLine1,
        _addressLine2,
        _addressLine3,
        _addressLine4,
        _telephone,
        _vatNumber,
        _regNumber As String

    Private _postalCode As Integer

    Private _isActive As Boolean

    Public Property ContractorNumber As String
        Get
            Return _contractorNumber
        End Get
        Set(value As String)
            If (value <> _contractorNumber) Then
                _contractorNumber = value
                NotifyPropertyChanged("ContractorNumber")
            End If
        End Set
    End Property
    Public Property ContractorName As String
        Get
            Return _contractorName
        End Get
        Set(value As String)
            If value <> _contractorName Then
                _contractorName = value
                NotifyPropertyChanged("ContractorName")
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
                _addressLine1 = value
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
    Public Property IsActive As Boolean
        Get
            Return _isActive
        End Get
        Set(value As Boolean)
            If value <> _isActive Then
                _isActive = value
                NotifyPropertyChanged("IsActive")
            End If
        End Set
    End Property
    Public Property VatNumber As String
        Get
            Return _vatNumber
        End Get
        Set(value As String)
            If value <> _vatNumber Then
                _vatNumber = value
                NotifyPropertyChanged("VatNumber")
            End If
        End Set
    End Property
    Public Property RegNumber As String
        Get
            Return _regNumber
        End Get
        Set(value As String)
            If value <> _regNumber Then
                _regNumber = value
                NotifyPropertyChanged("RegNumber")
            End If
        End Set
    End Property
    Public Property ContractorNameList As New List(Of String)

    Private Property ContractorData As ContractorData
    Private Property ContractorSet As DataSet

    Public Sub New()
        ContractorData = New ContractorData(ContractorNumber, ContractorName, AddressLine1, AddressLine2, AddressLine3, AddressLine4, PostalCode, Telephone, IsActive, VatNumber, RegNumber)

        InitializeContractorProperties(0)
    End Sub

    ''' <summary>
    ''' Update company property parameters
    ''' </summary>
    Public Sub InitializeContractorProperties(ByRef index As Integer)
        ContractorData.Adapter.Fill(ContractorSet)

        Dim row = ContractorSet.Tables.Item(0).Rows.Item(index)

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
            For Each newRow As DataRow In ContractorSet.Tables.Item(0).Rows
                CompanyNameList.Add(String.Format("[{0}] {1}", newRow("CompanyNo"), newRow("CompanyName")))
            Next
        End If
    End Sub

    ''' <summary>
    ''' Get the number of companies in the company table in the database
    ''' </summary>
    Public Sub GetCompanyCount(ByRef count As Integer)
        ' get the number of available companies in the database
        ContractorData.GetNumberOfContractors(ContractorNumber, count)
    End Sub

    ''' <summary>
    ''' Adds a row to the company table
    ''' </summary>
    Public Sub AddRowToCompanyTable()
        ' update the company table with data currently in the company fields
        ContractorData.AddContractorRow(ContractorNumber, ContractorName, AddressLine1, AddressLine2, AddressLine3, AddressLine4, PostalCode, Telephone, IsActive, VatNumber, RegNumber)
    End Sub

    ''' <summary>
    ''' Save an edit to the company table
    ''' </summary>
    Public Sub SaveEditToCompanyTable()
        ' save the editted row to the table
        ContractorData.SaveCompanyRowEdit(ContractorNumber, ContractorName, AddressLine1, AddressLine2, AddressLine3, AddressLine4, PostalCode, Telephone, IsActive, VatNumber, RegNumber)
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