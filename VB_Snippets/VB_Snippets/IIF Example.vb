Module IIF_Example

    Sub IFF_Example()
        Dim uservalue As String = InputBox("Enter Value, 1 or anything else")
        Dim message As String = IIf(uservalue = "1", "car", "cat")
        Debug.WriteLine("you picked {1} so... you won a {0}!", message, uservalue)
    End Sub


End Module
