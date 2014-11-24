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
End Class