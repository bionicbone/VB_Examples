Public NotInheritable Class Car
    Inherits Vehicle


    'Private _BluebookValue As Decimal
    'Public Property BlueBookValue() As Decimal
    '    Get
    '        Return _BluebookValue
    '    End Get
    '    Set(ByVal value As Decimal)
    '        _BluebookValue = value
    '    End Set
    'End Property

    ' When you see the "New" sub in a class this is the "Consructor" sub.
    ' When even a new instance of the Class is created the "Constructor" or New Sub 
    ' will be executed, this allows some logic to be applied during the creation of the object. 
    ' NOTE: A "Constructor" - "New Sub" is not actually required.
    Public Sub New()
        ' NOTE: use "Me" keyword to make the code easy to read that we are talking
        '       about the the "Make" as part of this class.
        Me.Make = "Nissan"
    End Sub

    ' Overload the Constructor by creating two "New" subs with different parameters 
    Public Sub New(ByVal _make As String,
                   ByVal _model As String,
                   ByVal _year As Integer,
                   ByVal _colour As String)
        Me.Make = _make
        Me.Model = _model
        Me.Year = _year
        Me.Colour = _colour
    End Sub


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

    ' NOTE: Both track class and car class has a FormatMe function
    '       Therefore we need to use the Overrides statement to ensure
    '       the correct function is used depending on if we are passing 
    '       a truck or a car.
    Public Overrides Function FormatMe() As String
        Return String.Format("{0} - {1} - {2} - {3}",
                             Me.Make,
                             Me.Model,
                             Me.Year,
                             Me.Colour)
    End Function

End Class
