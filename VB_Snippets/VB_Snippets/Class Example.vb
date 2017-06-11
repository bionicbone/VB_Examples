Public Class Car

    Public Property Make As String
    Public Property Model As String
    Public Property Year As Integer
    Public Property Colour As String

    'Private _BluebookValue As Decimal
    'Public Property BlueBookValue() As Decimal
    '    Get
    '        Return _BluebookValue
    '    End Get
    '    Set(ByVal value As Decimal)
    '        _BluebookValue = value
    '    End Set
    'End Property

    Public Function DetermineMarketValue() As Decimal

        'Return 100.0

        Dim carValue As Decimal
        If Me.Year > 2008 Then
            carValue = 10000
        Else
            carValue = 2000
        End If
        Return carValue
    End Function
End Class
