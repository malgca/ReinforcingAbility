Imports DataTier
Imports System.Data
Imports System.ComponentModel

Public Class BendingSchedule
    Implements INotifyPropertyChanged

    Public Property JobNumber As String
    Public Property JobNameList As New List(Of String)

    Private Property BendingScheduleData As BendingScheduleData
    Private Property BendingScheduleSet As New DataSet

    Public Sub New()
        BendingScheduleData = New BendingScheduleData(JobNumber)

        InitializeProperties(0)
    End Sub

    Public Sub InitializeProperties(ByVal index As Integer)

        PopulateJobList()
    End Sub

    Public Sub PopulateJobList()
        JobNameList.Clear()
        BendingScheduleData.PopulateJobsSet(BendingScheduleSet)

        For Each newRow As DataRow In BendingScheduleSet.Tables.Item(0).Rows
            If (IsNotNull(newRow("JobNo"))) Then
                JobNameList.Add(newRow("JobNo"))
            End If
        Next
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
