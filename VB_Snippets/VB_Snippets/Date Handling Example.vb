Module Date_Handling_Example

    Sub Dates()

        Dim myDateValue As Date = Now

        Debug.Print(myDateValue.ToShortDateString())
        Debug.Print(myDateValue.ToShortTimeString())
        Debug.Print(myDateValue.ToLongDateString())
        Debug.Print(myDateValue.AddDays(3).ToShortDateString())
        Debug.Print(myDateValue.AddHours(3).ToShortTimeString())
        Debug.Print(myDateValue.AddDays(-3).ToShortDateString())
        Debug.Print(myDateValue.Month())
        Debug.Print(myDateValue.DayOfWeek())
        Debug.Print(Format(myDateValue, "dddd"))

        Dim myBirthDate1 As New Date
        myBirthDate1 = Date.Parse("21/8/1967")
        Debug.Print(myBirthDate1)

        ' Note american standard formating!!
        Dim myBirthDate2 = #8/21/1967#
        Debug.Print(myBirthDate2)

        'best way because I dont think tis takes into account any local settings on the PC
        Dim myBirthDate3 = New Date(1967, 8, 21)
        Debug.Print(myBirthDate3)

        Dim myAge As TimeSpan = Date.Now.Subtract(myBirthDate3)
        Debug.Print(myAge.TotalDays)

    End Sub

End Module
