Module CPU_Example

    Sub CPU()

        Dim CPU As New PerformanceCounter()
        Dim MaxCPU As Integer = 0
        Dim CurrentCPU As Double = 0
        With CPU
            .CategoryName = "Processor"
            .CounterName = "% Processor Time"
            .InstanceName = "_Total"
        End With

        Dim Cycles As Integer = 10000

        For n = 1 To Cycles
            CurrentCPU += CPU.NextValue
        Next
        Debug.Print("For / Next = " & Cycles & " Avg CPU: " & CurrentCPU / Cycles & "%")

        MaxCPU = 0

        For n = 1 To Cycles
            Application.DoEvents()
            CurrentCPU += CPU.NextValue
        Next
        Debug.Print("For / Next = " & Cycles & " with DoEvents " & " Avg CPU: " & CurrentCPU / Cycles & "%")

    End Sub


End Module
