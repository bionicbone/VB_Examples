Public Class frmMainSelectionForm

    Private Sub btnProgressBar_Click(sender As Object, e As EventArgs) Handles btnProgressBar.Click
        frmProgressBar.Show()
    End Sub

    Private Sub btnSQLRead_Click(sender As Object, e As EventArgs) Handles btnSQLRead.Click
        Call SQL_Read()
    End Sub

    Private Sub btnSQLWrite_Click(sender As Object, e As EventArgs) Handles btnSQLWrite.Click
        Call SQL_Write()
    End Sub

    Private Sub btnSQLUpdate_Click(sender As Object, e As EventArgs) Handles btnSQLUpdate.Click
        Call SQL_Update()
    End Sub

    Private Sub btnTestThis_Click(sender As Object, e As EventArgs) Handles btnTestThis.Click
        Call TestThis()
    End Sub

    Private Sub btnStopWatch_Click(sender As Object, e As EventArgs) Handles btnStopWatch.Click
        Call myStopWatch()
    End Sub

    Private Sub btnCPU_Click(sender As Object, e As EventArgs) Handles btnCPU.Click
        Call CPU()
    End Sub

    Private Sub btnDataBaseExists_Click(sender As Object, e As EventArgs) Handles btnDataBaseExists.Click
        If CheckDatabaseExists("LT5-PC\SQLEXPRESS", "HobbyKingDB") <> False Then
            Debug.Print("Database Exists")
        Else
            Debug.Print("Database Does Not Exist")
        End If
    End Sub

    Private Sub btnTableExists_Click(sender As Object, e As EventArgs) Handles btnTableExists.Click
        If CheckTableExists("LT5-PC\SQLEXPRESS", "HobbyKingDB", "tblHKPrices") <> False Then
            Debug.Print("Table Exists")
        Else
            Debug.Print("Table Does Not Exist")
        End If
    End Sub

    Private Sub btnCreateTable_Click(sender As Object, e As EventArgs) Handles btnCreateTable.Click
        Call SQL_Create_Table()
    End Sub

    Private Sub btnCreateDataBase_Click(sender As Object, e As EventArgs) Handles btnCreateDataBase.Click
        Call SQL_Create_Database()
    End Sub

    Private Sub btnGetWebPage_Click(sender As Object, e As EventArgs) Handles btnGetWebPage.Click

        ' Return the Web Page HTML in wData
        Dim wData As String = WRequest("http://www.winzip.com", "GET", "")
        If Mid(wData, 1, 18) <> "An error occurred:" Then
            Debug.Print("Web Page Html:" & vbNewLine)
            Debug.Print(wData)
        Else
            Debug.Print(wData)
        End If
    End Sub

    Private Sub btnSearchWebData_Click(sender As Object, e As EventArgs) Handles btnSearchWebData.Click
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

    Private Sub btnWriteFile_Click(sender As Object, e As EventArgs) Handles btnWriteFile.Click
        If WriteFile() = False Then
            Debug.Print("File Failed to Write")
        End If
    End Sub


    Private Sub btnOverloadedMethods_Click(sender As Object, e As EventArgs) Handles btnOverloadedMethods.Click
        Call overloadedMethodsExample()
    End Sub

    Private Sub btnConcatenationPerf_Click(sender As Object, e As EventArgs) Handles btnConcatenationPerf.Click
        Call concatenationPerformance()
    End Sub

    Private Sub btnHandleDates_Click(sender As Object, e As EventArgs) Handles btnHandleDates.Click
        Call Dates()
    End Sub

    Private Sub btnClasses_Click(sender As Object, e As EventArgs) Handles btnClasses.Click
        Call CreateObjectFromClass()
    End Sub

    Private Sub frmMainSelectionForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnArrayReversal_Click(sender As Object, e As EventArgs) Handles btnArrayReversal.Click
        Call ReverseArray()
    End Sub


    Private Sub btnIIF_Example_Click(sender As Object, e As EventArgs) Handles btnIIF_Example.Click
        Call IFF_Example()
    End Sub

    Private Sub btnArrayForEach_Click(sender As Object, e As EventArgs) Handles btnArrayForEach.Click
        Call ForEach_Example()
    End Sub

    Private Sub btnReadFile_Click(sender As Object, e As EventArgs) Handles btnReadFile.Click
        Call ReadFile()
    End Sub
End Class