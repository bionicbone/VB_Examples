Module StopWatch_Example

    Sub myStopWatch()

        'Elapsed.TotalMilliseconds (double) returns the total number of whole and fractional milliseconds elapsed since inception
        'e.g. a stopwatch stopped at 1.23456 seconds would return 1234.56 in this property. See TimeSpan.TotalMilliseconds on MSDN

        'Elapsed.Milliseconds (int) returns the number of whole milliseconds in the current second
        'e.g. a stopwatch at 1.234 seconds would return 234 in this property. See TimeSpan.ElapsedMilliseconds

        'ElapsedTicks (long) returns the ticks since start of the stopwatch.
        'In the context of the original question, pertaining to the Stopwatch class, ElapsedTicks is the number of ticks elapsed. 
        'Ticks occur at the rate of Stopwatch.Frequency, so, to compute seconds elapsed, calculate: numSeconds = stopwatch.ElapsedTicks / Stopwatch.Frequency.
        'The old answer defined ticks as the number of 100 nanosecond periods, which is correct in the context of the DateTime class, 
        'but not correct in the context of the Stopwatch class. See Stopwatch.ElapsedTicks on MSDN.

        'ElapsedMilliseconds returns a rounded number to the nearest full millisecond, so this might lack precision Elapsed.TotalMilliseconds property can give.

        'Elapsed.TotalMilliseconds is a double that can return execution times to the partial millisecond while ElapsedMilliseconds is Int64. 
        'e.g. a stopwatch at 0.0007 milliseconds would return 0 , or 1234.56 milliseconds would return 1234 in this property. 
        'So for precision always use Elapsed.TotalMilliseconds.

        'See Stopwatch.ElapsedMilliseconds on MSDN for clarification.

        Dim Watch As Stopwatch = Stopwatch.StartNew()
        Dim Cycles As Integer = 100000

        For n = 1 To Cycles
        Next
        Watch.Stop()
        Debug.Print("For / Next = " & Cycles & ": " & Watch.Elapsed.TotalMilliseconds & " ms")

        Watch.Reset()
        Watch.Start()

        For n = 1 To Cycles
            Application.DoEvents()
        Next
        Watch.Stop()
        Debug.Print("For / Next = " & Cycles & " with DoEvents: " & Watch.Elapsed.TotalMilliseconds & " Total ms")

    End Sub

End Module
