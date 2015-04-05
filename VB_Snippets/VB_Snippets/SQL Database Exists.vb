Imports System.Data.SqlClient

Module SQLDatabaseExists

    Public Function CheckDatabaseExists(ByVal server As String, _
                                       ByVal database As String) As Boolean
        ' ********************************************************************************************
        ' *** Check if an SQL database exists                                                      ***
        ' ***       Kevin Guest - Nov 2014                                                         ***
        ' ***    Windows XP, 7 , and 8.1 Compatible                                                ***
        ' ***                                                                                      ***
        ' *** To test this example:                                                                ***
        ' *** --- Call this function with either a valid or not valid SQL Database.                ***
        ' ***   --- For Example: CheckDatabaseExists("LT5-PC\SQLEXPRESS", "HobbyKingDB")           ***
        ' ***   --- The correct Server Name can be found by opening the SQL Server Object Explorer *** 
        ' ***   --- and finding the property called Server.                                        ***
        ' ***                                                                                      ***
        ' *** Result: <> False if Database Exists                                                  ***
        ' ********************************************************************************************

        Dim mySqlConnection As New SqlConnection
        Dim mysqlCommand As SqlCommand

        Dim connString As String = ("Data Source=" _
                    + (server + ";Initial Catalog=master;Integrated Security=True;"))

        Dim cmdText As String = _
           ("SELECT COUNT (*) FROM sys.sysdatabases WHERE NAME='" & database & "'")

        mySqlConnection.ConnectionString = connString
        mySqlConnection.Open()

        mysqlCommand = New SqlCommand(cmdText, mySqlConnection)
        Return CByte(mysqlCommand.ExecuteScalar())

        mySqlConnection.Close()
    End Function


End Module
