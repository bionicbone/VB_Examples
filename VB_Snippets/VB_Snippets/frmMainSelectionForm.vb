Public Class frmMainSelectionForm

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        frmProgressBar.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call SQL_Read()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Call SQL_Write()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Call SQL_Update()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Call TestThis()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Call myStopWatch()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Call CPU()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If CheckDatabaseExists("LT5-PC\SQLEXPRESS", "HobbyKingDB") <> False Then
            Debug.Print("Database Exists")
        Else
            Debug.Print("Database Does Not Exist")
        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        If CheckTableExists("LT5-PC\SQLEXPRESS", "HobbyKingDB", "tblHKPrices") <> False Then
            Debug.Print("Table Exists")
        Else
            Debug.Print("Table Does Not Exist")
        End If
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Call SQL_Create_Table()
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Call SQL_Create_Database()
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click

    ' Return the Web Page HTML in wData
        Dim wData As String = WRequest("http://www.winzip.com", "GET", "")
        If Mid(wData, 1, 18) <> "An error occurred:" Then
            Debug.Print("Web Page Html:" & vbNewLine)
            Debug.Print(wData)
        Else
            Debug.Print(wData)
        End If
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        ' Use WRequest to bring back the Web Page as HTML, here we'll just prime the variable.
        Dim wData As String = ""
        wData = "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" />" & _
                "<!-- InstanceBeginEditable name=""doctitle"" -->" & _
                "<title>Zip Files, Unzip Files, Compress Files and Share Files with WinZip</title>" & _
                "<script type=""text/javascript"">" & _
                "// Global Variables" & _
                "var entryModal = false;" & _
                "var timerModal  = true;" & _
                "var currentLang = ""en"";" & _
                "var exeName = winzip19.exe ;" & _
                "var pubName = WinZip Computing ;" & _
                "var fileFrom = www.winzip.com ;" & _
                "var fileSize = 802 KB ;" & _
                "</script>"

        ' Get the exe name from the above HTML, using the code line below.
        ' Search for "// Global Variables"
        ' Then from that point search for "var exeName ="
        ' The first character we want will be 14 characters on from this point, because "var exeName =" is 13 charaters long.
        ' The Search for ";"
        ' The last character we want will be -1 characters before this point, because ";" is 1 character long.
        ' The result from the above HTML snippet would be "winzip19.exe"
        Dim rData As String = Search_Web_Data(wData, "// Global Variables", 0, "var exeName =", 14, ";", -1)
        If rData <> "N/A" Then
            Debug.Print(rData)
        Else
            Debug.Print("The data is not available...")
        End If
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        If WriteFile = False Then
            Debug.Print("File Failed to Write")
        End If
    End Sub


    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        Call overloadedMethodsExample()
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        Call concatenationPerformance()
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        Call Dates()
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        Call CreateObjectFromClass()
    End Sub
End Class