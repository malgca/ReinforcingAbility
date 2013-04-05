' CuttingSheet Class with added methods for DA Class
' Chapter 13
Public Class CuttingSheet
    'attributes
    Private cutSheetNo As String
    Private invoiceNo As String
    Private cutDate As Date

    'get accessor methods 
    Public Function GetCutSheet() As String
        Return cutSheetNo
    End Function
    Public Function GetInvoiceNo() As String
        Return invoiceNo
    End Function
    Public Function GetDate() As String
        Return cutDate
    End Function

    'set accessor methods 
    Public Function TellAboutSelf() As String
        Dim info As String
        info = "C/Sheet = " & GetCutSheet() & _
               ", Inv = " & GetInvoiceNo() & _
               ", Date = " & GetDate()
        Return info
    End Function

    ' property named CuttingSheetName
   
    '------------------Begin Constructors---------------
    'default constructor
   
    'constructor (3 parameters)
    '------------Begin Data Access Shared Methods-----
    Public Shared Sub Initialize()
        CutSheetDA.Initialize()
    End Sub
    Public Shared Function GetAll() As ArrayList
        Return CutSheetDA.GetAll
    End Function
    Public Shared Sub Terminate()
        CutSheetDA.Terminate()
    End Sub
    '------------End Data Access Shared Methods---------
    '------------Begin Data Access Instance Methods-----

    '------------End Data Access Instance Methods-----
End Class







