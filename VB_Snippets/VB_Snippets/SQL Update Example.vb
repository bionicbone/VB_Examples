Imports System.Data.SqlClient

Module SQL_Update_Example

    Dim xData As SqlDataReader
    Dim mySqlConnection As New SqlConnection
    Dim mysqlCommand As SqlCommand

    Public Sub SQL_Update()

        ' *************************************************************************************
        ' *** SQL Update Example                                                              ***
        ' ***       Kevin Guest - Nov 2014                                                  ***
        ' ***    Windows XP, 7 , and 8.1 Compatible                                         ***
        ' ***                                                                               ***
        ' *** To test this example:                                                         ***
        ' *** --- Create a Database called HobbyKingDB.                                     ***
        ' *** --- Within that BD there must be a Table called Table_1.                      ***
        ' *** --- Within that table create 3 fields called Field_0, Field_1 and Field_2.    ***
        ' ***                                                                               ***
        ' *************************************************************************************

        ' NOTE: The correct connection string can be found by opening the SQL Server Object Explorer
        ' from with in Visual Studio 2013.
        ' View --- SQL Server Object Explorer
        ' Select the Database you want to connect to and select properties.
        ' Find the property "Connection String" and use that here. 

        ' Connect to the DB.
        mySqlConnection.ConnectionString = _
            "Data Source=LT5-PC\SQLEXPRESS;Initial Catalog=HobbyKingDB;Integrated Security=True;Connect Timeout=15;Encrypt=False;TrustServerCertificate=False"
        mySqlConnection.Open()

        ' Set up some Variables to write.
        Dim Data_Name As String = "John"
        Dim Data_Age As String = "87"
        Dim Data_Month As String = "June"

        ' Create the UPDATE Statement.
        mysqlCommand = New SqlCommand(("UPDATE Table_1 " & _
                                       "SET Field_1 = '" & Data_Age + 1 & "', " & _
                                       "Field_2 = '" & Data_Month & "' " & _
                                       "WHERE (Field_0 = '" & Data_Name & "')"), mySqlConnection)
        mysqlCommand.ExecuteNonQuery()

        ' Alway Close the Connection after use or when terminating the program.
        mySqlConnection.Close()

    End Sub

End Module
