Imports System.Text
Imports System.Diagnostics

Module String_Concatenation_Example

    Sub concatenationPerformance()
        Dim sWatch As Stopwatch = New Stopwatch()
        Dim myString As String = ""
        Dim myStringBuilder As StringBuilder = New StringBuilder

        sWatch.Reset()
        sWatch.Start()
        myString = ""
        For I = 1 To 50000
            myString = myString & CStr(I)
        Next
        sWatch.Stop()
        Debug.Print("50,000 x myString = myString & CStr(I) = {0}ms", sWatch.ElapsedMilliseconds.ToString)

        sWatch.Reset()
        sWatch.Start()
        myString = ""
        For I = 1 To 50000
            myString += CStr(I)
        Next
        sWatch.Stop()
        Debug.Print("50,000 x myString += CStr(I) = {0}ms", sWatch.ElapsedMilliseconds.ToString)

        sWatch.Reset()
        sWatch.Start()
        myString = ""
        For I = 1 To 50000
            myStringBuilder.Append(CStr(I))
        Next
        sWatch.Stop()
        Debug.Print("50,000 x mystringbuilder.Append(CStr(I)) = {0}ms", sWatch.ElapsedMilliseconds.ToString)
        myStringBuilder.Clear()
        myStringBuilder = Nothing
        myString = Nothing
    End Sub


End Module
