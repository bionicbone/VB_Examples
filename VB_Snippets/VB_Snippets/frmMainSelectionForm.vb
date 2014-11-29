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
End Class