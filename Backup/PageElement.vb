Public Class PageElement

    Public Text As String
    Public Font As Font
    Public x As Integer
    Public x2 As Integer
    Public includeEOL As Boolean
    Public Imageindex As Integer
    Public ImageHeight As Integer
    Public Ygap As Integer
    Public center As Boolean
    Public rAlign As Boolean


    Sub New(ByVal txt As String, ByVal Fnt As Font, ByVal xc As Integer, ByVal EOL As Boolean, ByVal cntr As Boolean, ByVal RightAlign As Boolean)
        Text = txt
        Font = Fnt
        x = xc
        includeEOL = EOL
        Me.Ygap = 0
        center = cntr
        rAlign = RightAlign
    End Sub
    Sub New(ByVal txt As String, ByVal Fnt As Font, ByVal xc As Integer, ByVal EOL As Boolean, ByVal cntr As Boolean)
        Text = txt
        Font = Fnt
        x = xc
        includeEOL = EOL
        Me.Ygap = 0
        center = cntr
        rAlign = False
    End Sub
    Sub New(ByVal x1 As Integer, ByVal x_2 As Integer, ByVal EOL As Boolean)
        Text = "#LINE__"
        Font = New Font("Arial", 10)
        x = x1
        x2 = x_2
        includeEOL = EOL
        Me.Ygap = 0
        center = False
        rAlign = False
    End Sub
    Sub New(ByVal EOL As Boolean, ByVal x1 As Integer, ByVal x_2 As Integer)
        Text = "#DOUBLELINE__"
        Font = New Font("Arial", 10)
        x = x1
        x2 = x_2
        includeEOL = EOL
        Me.Ygap = 0
        center = False
        rAlign = False
    End Sub
    Sub New(ByVal txt As String, ByVal Fnt As Font, ByVal xc As Integer, ByVal cntr As Boolean)
        Text = txt
        Font = Fnt
        x = xc
        includeEOL = False
        Me.Ygap = 0
        center = cntr
        rAlign = False
    End Sub
    Sub New(ByVal txt As String, ByVal Fnt As Font, ByVal xc As Integer, ByVal yGap As Integer, ByVal cntr As Boolean)
        Text = txt
        Font = Fnt
        x = xc
        includeEOL = True
        Me.Ygap = yGap
        center = cntr
        rAlign = False
    End Sub
    Sub New(ByVal txt As String, ByVal imgIndex As Integer, ByVal imgHeight As Integer, ByVal xc As Integer, ByVal EOL As Boolean, ByVal cntr As Boolean)
        Text = txt
        x = xc
        ImageHeight = imgHeight
        Imageindex = imgIndex
        includeEOL = EOL
        Me.Ygap = 0
        center = cntr
        rAlign = False
    End Sub
End Class
