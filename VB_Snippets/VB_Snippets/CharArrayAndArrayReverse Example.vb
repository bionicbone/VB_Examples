Module CharArrayAndArrayReverse_Example

    Sub ReverseArray()
        Dim MyText As String = "You can get what you want out of life " &
            "if you help enough people get what they want."
        Dim CharArray() As Char = MyText.ToCharArray

        For Each MyChar As Char In CharArray
            Debug.Write(MyChar)
        Next

        Debug.Print("")

        Array.Reverse(CharArray)

        For Each MyChar As Char In CharArray
            Debug.Write(MyChar)
        Next
    End Sub

End Module

