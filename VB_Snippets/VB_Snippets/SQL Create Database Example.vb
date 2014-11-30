Imports System.Data.SqlClient

Module SQL_Create_Database_Example

    Sub SQL_Create_Database()

        Dim DatabaseName As String = "HobbyKingDB2"
        Dim ServerName As String = "(localdb)\v11.0"
        ' LT5-PC\SQLEXPRESS
        ' (localdb)\ProjectsV12
        '(localdb)\MSSQLLocalDB
        '(localdb)\v11.0

        Dim mySQLConnection As SqlConnection
        Dim mySqlCommand As SqlCommand

        Dim str As String = ""
        ' Ensure the Databases Exist, if not create them and the table structure
        If CheckDatabaseExists(ServerName, DatabaseName) = False Then
            ' ########  Create the Database  #########
            mySQLConnection = New SqlConnection("Server=" & ServerName & _
                                                ";database=master;Integrated Security=True;")
            str = "CREATE DATABASE " & DatabaseName
            mySQLConnection.Open()
            mySqlCommand = New SqlCommand(str, mySQLConnection)
            mySqlCommand.ExecuteNonQuery()
            mySQLConnection.Close()
        End If

    End Sub

End Module
