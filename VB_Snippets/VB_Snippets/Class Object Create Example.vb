Module ClassObjectCreateExample

    Sub CreateObjectFromClass()

        ' You can instantiate many objects from the same class, and each 
        ' can hold different values in each of its properties as long as
        ' the "new" command is used.

        ' NOTE: Both the following statments do the smae thing!
        Dim myNewCar1 As Car = New Car()
        Dim myNewCar2 As New Car()

        myNewCar1.Make = "Audi"
        myNewCar1.Model = "A4"
        myNewCar1.year = 2017
        myNewCar1.Colour = "Red"

        myNewCar2.Make = "Nissan"
        myNewCar2.Model = "Almera Tino"
        myNewCar2.Year = 2004
        myNewCar2.Colour = "Silver"

        Debug.Print("{0} - {1} - {2} - {3}",
                    myNewCar1.Make,
                    myNewCar1.Model,
                    myNewCar1.Year,
                    myNewCar1.Colour)
        Debug.Print("{0} - {1} - {2} - {3}",
                    myNewCar2.Make,
                    myNewCar2.Model,
                    myNewCar2.Year,
                    myNewCar2.Colour)


        ' using a Function in a Module, its a bit old fashioned now
        Debug.Print("Car 1 Value: {0:C}", determineMarketValue(myNewCar1))

        ' but normally the Fuction would be in the Class, and called like this.
        ' NOTE: the statement is formated like the inbuilt VB classes.
        '       i.e. console.writeline()
        Debug.Print("Car 1 Value: {0:C}", myNewCar1.DetermineMarketValue())


        ' and the same for Car2
        Debug.Print("Car 2 Value: {0:C}", determineMarketValue(myNewCar2))
        Debug.Print("Car 2 Value: {0:C}", myNewCar2.DetermineMarketValue())


        ' Speed test for each way of doing this
        ' CONCLUSION: no difference.
        Dim Result As Decimal
        Dim sWatch As New Stopwatch

        ' Module Function
        sWatch.Stop()
        sWatch.Reset()
        sWatch.Start()
        For i = 1 To 100000
            Result = determineMarketValue(myNewCar1)
        Next
        Debug.Print("Execution Time: {0}ms", sWatch.ElapsedMilliseconds)

        ' Module Function
        sWatch.Stop()
        sWatch.Reset()
        sWatch.Start()
        For i = 1 To 100000
            Result = myNewCar1.DetermineMarketValue()
        Next
        Debug.Print("Execution Time: {0}ms", sWatch.ElapsedMilliseconds)

    End Sub

    Function determineMarketValue(ByVal myCar As Car) As Decimal

        Dim carValue As Decimal
        If myCar.Year > 2008 Then
            carValue = 10000
        Else
            carValue = 2000
        End If
        Return carValue

    End Function

End Module
