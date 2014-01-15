Imports System.Drawing

Public Class PageElement

    Public Class LineConstants
        Public Const Line As String = "#LINE__"
        Public Const DoubleLine As String = "#DOUBLELINE__"
        Public Const Space As String = "<SPACE>"
        Public Const HR As String = "<HR/>"
        Public Const HRBlack As String = "<HR/BLACK>"
        Public Const HRLight As String = "<HR/LIGHT>"
    End Class

    ' fonts used frequently in reinforcing
    Public Class PageFonts
        ''' <summary>
        ''' Arial 10, Normal
        ''' </summary>
        Public Shared Normal As New Font("Arial", 10)

        ''' <summary>
        ''' Arial 8, Italic
        ''' </summary>
        Public Shared SmallItalic As New Font("Arial", 8, FontStyle.Italic)
    End Class

    Public Property ElementText As String
    Public Property ElementFont As Font

    Public Property xPositionOne As Integer
    Public Property xPositionTwo As Integer

    Public Property IncludeEOL As Boolean

    Public Property ImageIndex As Integer
    Public Property ImageHeight As Integer

    Public Property YGap As Integer
    Public Property CenterElement As Boolean
    Public Property RightAlignElement As Boolean

    Sub New(ByRef elementText As String, ByRef elementFont As Font, ByVal xPosition As Integer, ByVal endOfLine As Boolean, ByVal center As Boolean, ByVal rightAlign As Boolean)
        Me.ElementText = elementText
        Me.ElementFont = elementFont
        Me.xPositionOne = xPosition
        Me.IncludeEOL = endOfLine
        Me.YGap = 0
        Me.CenterElement = center
        Me.RightAlignElement = rightAlign
    End Sub

    Sub New(ByRef elementText As String, ByRef elementFont As Font, ByVal xPosition As Integer, ByVal endOfLine As Boolean, ByVal center As Boolean)
        Me.ElementText = elementText
        Me.ElementFont = elementFont
        Me.xPositionOne = xPosition
        Me.IncludeEOL = endOfLine
        Me.YGap = 0
        Me.CenterElement = center
        Me.RightAlignElement = False
    End Sub

    Sub New(ByVal x1 As Integer, ByVal x2 As Integer, ByVal endOfLine As Boolean)
        Me.ElementText = LineConstants.Line
        Me.ElementFont = PageFonts.Normal
        Me.xPositionOne = x1
        Me.xPositionTwo = x2
        Me.IncludeEOL = endOfLine
        Me.YGap = 0
        Me.CenterElement = False
        Me.RightAlignElement = False
    End Sub

    Sub New(ByVal endOfLine As Boolean, ByVal x1 As Integer, ByVal x2 As Integer)
        Me.ElementText = LineConstants.DoubleLine
        Me.ElementFont = PageFonts.Normal
        Me.xPositionOne = x1
        Me.xPositionTwo = x2
        Me.IncludeEOL = endOfLine
        Me.YGap = 0
        Me.CenterElement = False
        Me.RightAlignElement = False
    End Sub

    Sub New(ByRef elementText As String, ByRef elementFont As Font, ByVal xPosition As Integer, ByVal center As Boolean)
        Me.ElementText = elementText
        Me.ElementFont = elementFont
        Me.xPositionOne = xPosition
        Me.IncludeEOL = False
        Me.YGap = 0
        Me.CenterElement = center
        Me.RightAlignElement = False
    End Sub

    Sub New(ByRef elementText As String, ByRef elementFont As Font, ByVal xPosition As Integer, ByVal yGap As Integer, ByVal center As Boolean)
        Me.ElementText = elementText
        Me.ElementFont = elementFont
        Me.xPositionOne = xPosition
        Me.IncludeEOL = True
        Me.YGap = yGap
        Me.CenterElement = center
        Me.RightAlignElement = False
    End Sub

    Sub New(ByRef elementText As String, ByVal imgIndex As Integer, ByVal imgHeight As Integer, ByVal xPosition As Integer, ByVal endOfLine As Boolean, ByVal center As Boolean)
        Me.ElementText = elementText
        Me.xPositionOne = xPosition
        Me.ImageHeight = imgHeight
        Me.ImageIndex = imgIndex
        Me.IncludeEOL = endOfLine
        Me.YGap = 0
        Me.CenterElement = center
        Me.RightAlignElement = False
    End Sub
End Class