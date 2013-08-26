

Imports System
Imports System.Data


Namespace LogicTier
    Public Class Company
        Public Property CompanyNo As String
        Public Property CompanyName As String
        Public Property Address As String
        Public Property AddressLine2 As String
        Public Property AddressLine3 As String
        Public Property AddressLine4 As String
        Public Property Email As String
        Public Property Fax As String
        Public Property LastCutNum As Integer
        Public Property LastInvNum As Integer
        Public Property Message As String
        Public Property PostalCode As String
        Public Property RegNo As String
        Public Property TelNo As String
        Public Property UnitofMeas As String
        Public Property VatNo As String
        Public Property VatPerc As Decimal
        Public Property Website As String
        Public Property NoAndName As String
        Dim dtInstance = DataTier.DataTier.DBOperations.GetInstance


        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="CompanyNo">Number assinged to this company</param>
        ''' <param name="CompanyName">Name of this company</param>
        ''' <param name="Address">Physical Address (Line 1) of company</param>
        ''' <param name="AddressLine2">Physical Address (Line 2) of company</param>
        ''' <param name="AddressLine3">Physical Address (Line 3) of company</param>
        ''' <param name="AddressLine4">Physical Address (Line 4) of company</param>
        ''' <param name="Email">Email Address of the company</param>
        ''' <param name="Fax">Fax Number of the Company -- Is this still relevant today?</param>
        ''' <param name="Message">Other uncalssafied information about the company</param>
        ''' <param name="PostalCode">Postal code for city/suburb of company</param>
        ''' <param name="LastCutNum">Number assigned to the last cutting sheet for this company</param>
        ''' <param name="LastInvNum">Number assinged to the last invoice for this company</param>
        ''' <param name="RegNo">Registration Number of this company</param>
        ''' <param name="TelNo">Telephone Number for this company</param>
        ''' <param name="UnitofMeas">The unit of measurement this company uses</param>
        ''' <param name="VatNo">Value-Added Tax Number assigend to this company by SARS</param>
        ''' <param name="VatPerc">Value-Added Tax Percentage used by this company</param>
        ''' <param name="Website">Website Address for this company</param>
        ''' <param name="NoAndName"></param>
        ''' <remarks></remarks>
        Public Sub NewCompany(ByRef CompanyNo As String, ByRef CompanyName As String, ByRef Address As String, ByRef AddressLine2 As String, ByRef AddressLine3 As String, ByRef AddressLine4 As String, ByRef Email As String, ByRef Fax As String, ByRef Message As String, ByRef PostalCode As String, ByRef LastCutNum As Integer, ByRef LastInvNum As Integer, ByRef RegNo As String, ByRef TelNo As String, ByRef UnitofMeas As String, ByRef VatNo As String, ByRef VatPerc As Decimal, ByRef Website As String, ByRef NoAndName As String)

            Me.CompanyNo = CompanyNo
            Me.CompanyName = CompanyName
            Me.Address = Address
            Me.AddressLine2 = AddressLine2
            Me.AddressLine3 = AddressLine3
            Me.AddressLine4 = AddressLine4
            Me.Email = Email
            Me.Fax = Fax
            Me.LastInvNum = LastInvNum
            Me.LastCutNum = LastCutNum
            Me.Message = Message
            Me.PostalCode = PostalCode

        End Sub

        Public Sub saveNewCompany()

        End Sub

        Public Sub updateCompany()

        End Sub

        Public Function loadCompanies()
            Dim dsCompany As DataSet = dtInstance.DBOperations.GetInstance.getCompanyDataSet()
            Return dsCompany
        End Function
    End Class
End Namespace

