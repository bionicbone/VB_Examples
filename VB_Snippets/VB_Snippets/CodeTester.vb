'Imports System.Diagnostics

Module CodeTester

    Sub TestThis()
        Dim sWatch As System.Diagnostics.Stopwatch
        Dim TestPlusPlus As Integer = 0
        sWatch = New System.Diagnostics.Stopwatch()

        sWatch.Start()

        For N = 1 To 10000


            'If N / 100 - Int(N / 100) = 0 Then
            '    Debug.Print(N)
            'End If


            If N Mod 100 = 0 Then
                Debug.Print(N)
                TestPlusPlus += N
            End If

        Next

        sWatch.Stop()

        Debug.Print("Total = {0}", TestPlusPlus)
        Debug.Print("Timer = {0}ms", sWatch.ElapsedMilliseconds.ToString)


    End Sub

End Module
