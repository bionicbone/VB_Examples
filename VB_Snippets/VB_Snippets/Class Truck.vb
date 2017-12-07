Public NotInheritable Class Truck
    Inherits Vehicle
    Public Property TowingCapacity As Integer

    ' NOTE: Both track class and car class has a FormatMe function
    '       Therefore we need to use the Overrides statement to ensure
    '       the correct function is used depending on if we are passing 
    '       a truck or a car.
    Public Overrides Function FormatMe() As String
        Return String.Format("{0} - {1} - {2}",
                             Me.Make,
                             Me.Model,
                             Me.TowingCapacity)
    End Function
End Class
