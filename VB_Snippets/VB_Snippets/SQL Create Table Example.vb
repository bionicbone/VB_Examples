Imports System.Data.SqlClient

Module SQL_Create_Table_Example

    Public Sub SQL_Create_Table()

        Dim DatabaseName As String = "HobbyKingDB2"
        Dim ServerName As String = "(localdb)\v11.0"
        ' LT5-PC\SQLEXPRESS
        ' (localdb)\ProjectsV12
        '(localdb)\MSSQLLocalDB
        '(localdb)\v11.0

        Dim mySqlConnection2 As SqlConnection = New SqlConnection("Data Source=" & ServerName & _
                                                             ";Initial Catalog=" & DatabaseName & _
                                                             ";Integrated Security=True" & _
                                                             ";Connect Timeout=15" & _
                                                             ";Encrypt=False" & _
                                                             ";TrustServerCertificate=False")
        mySqlConnection2.Open()

        Dim str As String = ""
        ' Ensure the Tables Exist, the names are the same even in the SandBox database
        If CheckTableExists(ServerName, DatabaseName, "tblHKDownloadProcess") = False Then
            ' ########  Create This Table  #########
            str = "CREATE TABLE tblHKDownloadProcess (" & _
                                "ProcessID          NVARCHAR (50) NOT NULL PRIMARY KEY," & _
                                "Running            CHAR (3)      NOT NULL," & _
                                "LastRun            SMALLDATETIME NOT NULL," & _
                                "Start              INT           NOT NULL," & _
                                "Finish             INT           NOT NULL," & _
                                "TimeTaken          INT           NOT NULL," & _
                                "TimeTakenEstimated CHAR (3)      NOT NULL," & _
                                "ActiveSKUsInRange  INT           NULL," & _
                                "ScanningID         INT           NULL," & _
                                "WebErrors          INT           NULL," & _
                                "TotalItemsFound    INT           NULL," & _
                                "IntItemsFound      INT           NULL," & _
                                "UKItemsFound       INT           NULL," & _
                                "EUItemsFound       INT           NULL," & _
                                "LastHit            INT           NULL," & _
                                "FinishTime         INT           NULL," & _
                                "TotalLastRunTime   INT           NULL" & _
                                ")"

            Dim myCommand As SqlCommand = New SqlCommand(str, mySqlConnection2)
            myCommand.ExecuteNonQuery()
        End If
    End Sub
End Module
