Imports System.Data.SqlClient

Module SQL_Read_Example

    Dim xData As SqlDataReader
    Dim mySqlConnection As New SqlConnection
    Dim mysqlCommand As SqlCommand

    Public Sub SQL_Read()

        ' *************************************************************************************
        ' *** SQL Read Example                                                              ***
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

        ' Create the SELECT Statement.
        mysqlCommand = New SqlCommand(("SELECT Table_1.* FROM Table_1 WHERE (Field_0 = 'Kit')"), mySqlConnection)
        xData = mysqlCommand.ExecuteReader()

        ' Ensure "some" data has been returned before attempting to read.
        If xData.HasRows = True Then

            ' Items are returned in the same order as the Table Fields, starting from 0.
            xData.Read()
            Debug.Print("Name: " & xData.Item(0))
            Debug.Print("Age: " & xData.Item(1))
            Debug.Print("Birth Month: " & xData.Item(2))
        End If

        ' Always Close the DataSet after use ready for the next DataSet to be loaded.
        xData.Close()

        ' Alway Close the Connection after use or when terminating the program.
        mySqlConnection.Close()

    End Sub

End Module
