Module ClassObjectCreateExample

    Sub CreateObjectFromClass()

        ' You can instantiate many objects from the same class, and each 
        ' can hold different values in each of its properties as long as
        ' the "new" command is used.

        ' NOTE: Both the following statments do the same thing!
        Dim myNewCar1 As Car = New Car()
        Dim myNewCar2 As New Car()
        ' This one uses the "overloaded constructor" (the 2nd Sub New in Car Class) to set the parameters in the one statement.
        Dim myNewCar3 As New Car("Ford", "Excape", 2005, "White")


        Debug.Print("The ""Constructor"" - ""Sub New"" in the Class set the make to: {0}", myNewCar1.Make)

        myNewCar1.Make = "Audi"
        myNewCar1.Model = "A4"
        myNewCar1.Year = 2017
        myNewCar1.Colour = "Red"

        myNewCar2.Make = "Nissan"
        myNewCar2.Model = "Almera Tino"
        myNewCar2.Year = 2004
        myNewCar2.Colour = "Silver"

        Debug.Print("These Parameters for the 1st Car are set in the Module: {0} - {1} - {2} - {3}",
                    myNewCar1.Make,
                    myNewCar1.Model,
                    myNewCar1.Year,
                    myNewCar1.Colour)
        Debug.Print("These Parameters for the 2nd Car are set in the Module: {0} - {1} - {2} - {3}",
                    myNewCar2.Make,
                    myNewCar2.Model,
                    myNewCar2.Year,
                    myNewCar2.Colour)


        ' using a Function in a Module, its a bit old fashioned now
        ' NOTE: :C} set the format "Currency"
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
        Debug.Print("Execution Time for Module: {0}ms", sWatch.ElapsedMilliseconds)

        ' Class Function
        sWatch.Stop()
        sWatch.Reset()
        sWatch.Start()
        For i = 1 To 100000
            Result = myNewCar1.DetermineMarketValue()
        Next
        Debug.Print("Execution Time for Class: {0}ms", sWatch.ElapsedMilliseconds)

        Dim myNewCar4 As New Car()
        myNewCar4.Make = "BMW"
        myNewCar4.Model = "745Li"
        myNewCar4.Year = 2005
        myNewCar4.Colour = "Black"
        ' printVehicleDetails will use the FormatMe function from the Car Class 
        ' because that is what we are passing.
        printVehicleDetails(myNewCar4)

        Dim myNewTruck1 As New Truck()
        With myNewTruck1
            .Make = "Ford"
            .Model = "F950"
            .Year = 2006
            .Colour = "Black"
            .TowingCapacity = 1200
        End With
        ' printVehicleDetails will use the FormatMe function from the Truck Class 
        ' because that is what we are passing.
        printVehicleDetails(myNewTruck1)
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

    Sub printVehicleDetails(ByVal _vehicle As Vehicle)
        ' This will use the FormatMe function from either Car or Truck
        ' depending on which Vehicle has been passed. 
        Debug.Print("Here are the details: {0}",
                    _vehicle.FormatMe())
    End Sub

End Module
