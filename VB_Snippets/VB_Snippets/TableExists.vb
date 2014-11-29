Imports System.Data.SqlClient

Module SQLTableExists

    Public Function CheckTableExists(ByVal server As String, _
                                   ByVal database As String, _
                                   ByVal table As String) As Boolean
        ' ********************************************************************************************
        ' *** Check if an SQL database table                                                       ***
        ' ***       Kevin Guest - Nov 2014                                                         ***
        ' ***    Windows XP, 7 , and 8.1 Compatible                                                ***
        ' ***                                                                                      ***
        ' *** To test this example:                                                                ***
        ' *** --- Call this function with either a valid or not valid SQL Database.                ***
        ' ***   --- Example: CheckTableExists("LT5-PC\SQLEXPRESS", "HobbyKingDB","tblHKPrices")    ***
        ' ***   --- The correct Server Name can be found by opening the SQL Server Object Explorer *** 
        ' ***   --- and finding the property called Server.                                        ***
        ' ***   --- NOTE: Always use CheckDatabaseExists example before checking the tables        *** 
        ' ***                                                                                      ***
        ' *** Result: <> False if Database Exists                                                  ***
        ' ********************************************************************************************

        Dim mySqlConnection As New SqlConnection

        Dim connString As String = ("Data Source=" _
                    & (server & ";Initial Catalog=" & database & ";Integrated Security=True;"))

        mySqlConnection.ConnectionString = connString
        mySqlConnection.Open()

        Dim restrictions(3) As String
        restrictions(2) = table
        Dim dbTbl As DataTable = mySqlConnection.GetSchema("Tables", restrictions)

        If dbTbl.Rows.Count = 0 Then
            'Table does not exist
            Return False
        Else
            'Table exists
            Return True
        End If

        dbTbl.Dispose()
        mySqlConnection.Close()
    End Function


End Module
