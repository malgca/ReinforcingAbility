Public Class Job

    Private _jobRate As New Double

    Public Property JobNumber As String
    Public Property JobName As String
    Public Property Discount As Double
    Public Property AddedDiscount As Double
    Public ReadOnly Property CompanyNumber As Integer
        Get
            Return 1
        End Get
    End Property
    Public Property ContractorNumber As String
    Public Property OrderNumber As String
    Public Property MeasurementUnit As MeasurementUnit
    Public Property SteelTypes As List(Of BeamTypes)
    Public Property JobRate As Double
        Get
            Return _jobRate
        End Get
        Private Set(value As Double)

        End Set
    End Property
    ''' <summary>
    ''' Full Job constructor.
    ''' </summary>
    ''' <param name="JobNumber">Number code for this job.</param>
    ''' <param name="JobName">Name of a particular job.</param>
    ''' <param name="Discount">Discount given to Reinforcing Ability clients.</param>
    ''' <param name="AddedDiscount">Added Discount given to Reinforcing Ability clients.</param>
    ''' <param name="ContractorNumber">Number code assigned to client.</param>
    ''' <param name="OrderNumber">Number code assigned to Job orderr.</param>
    ''' <param name="MeasurementUnit">Unit of measurement to be used for a specific job.</param>
    Public Sub New(ByRef JobNumber As String, ByRef JobName As String, ByVal Discount As Double, ByVal AddedDiscount As Double, ByRef ContractorNumber As String, ByRef OrderNumber As String, ByVal MeasurementUnit As MeasurementUnit)
        Me.JobNumber = JobNumber
        Me.JobName = JobName
        Me.Discount = Discount
        Me.AddedDiscount = AddedDiscount
        Me.ContractorNumber = ContractorNumber
        Me.OrderNumber = OrderNumber
        Me.MeasurementUnit = MeasurementUnit
    End Sub
End Class