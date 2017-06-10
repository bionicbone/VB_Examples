Module Array_For_Each_Example
    Sub ForEach_Example()
        Dim Names() As String = {"Eddie", "Alex", "Michael", "David Lee"}
        For Each Name As String In Names
            Debug.WriteLine(Name)
        Next
    End Sub
End Module
