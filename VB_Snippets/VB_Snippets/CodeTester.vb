'Imports System.Diagnostics

Module CodeTester

    Sub TestThis()
        Dim sWatch As System.Diagnostics.Stopwatch

        sWatch = New System.Diagnostics.Stopwatch()

        sWatch.Start()

        For N = 1 To 10000


            'If N / 100 - Int(N / 100) = 0 Then
            '    Debug.Print(N)
            'End If


            If N Mod 100 = 0 Then
                Debug.Print(N)
            End If

        Next

        sWatch.Stop()

        Debug.Print(sWatch.Elapsed.ToString)


    End Sub

End Module
