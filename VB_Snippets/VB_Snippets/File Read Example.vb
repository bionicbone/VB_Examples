Imports System.IO

Module File_Read_Example

    Public Function ReadFile() As Boolean
        ' ########################################################
        ' ### Kevin Guest - 03 Jan 2015                        ###
        ' ### The function will read a text file               ###
        ' ### If it successful then the return value is true,  ###
        ' ### otherwise it will be false.                      ###
        ' ########################################################

        Dim Path As String = "C:\Temp\"
        Const TEXTFILE As String = "Test.txt"

        Try
            If File.Exists(Path & TEXTFILE) Then
                Debug.Print("Text file exists, reading file...")
                Dim objReader As New System.IO.StreamReader(Path & TEXTFILE)
                Dim Data As String = ""
                If objReader.Peek() <> -1 Then
                    Data = objReader.ReadLine()
                    Debug.Print(Data)
                End If
                objReader.Close()
                Debug.Print("Successful...")
            Else
                ' ### WRITE A STANDARD TEXT FILE HERE ###
                Return False
            End If
        Catch e As Exception
            Dim str As String = ""
            str = "An error occured while reading the Text file." & vbNewLine & e.Message
            MsgBox(str, vbOKOnly + vbExclamation, "Error")
            Return False
        End Try
        Return True
    End Function

End Module
