Imports System.IO

Module File_Write_Example
    Public Function WriteFile() As Boolean
        ' ################################################################
        ' ### Kevin Guest - 03 Jan 2015                                ###
        ' ### The function will write a text file.                     ###
        ' ### NOTE: If the file already exists it will be overwritten! ###
        ' ### If it successful then the return value is true,          ###
        ' ### otherwise it will be false.                              ###
        ' ################################################################

        Dim Path As String = "C:\Temp\"
        Const TEXTFILE As String = "Test.txt"

        Try
            Debug.Print("Write file " & Path & TEXTFILE & " ...")
            Dim objWriter As StreamWriter = File.CreateText(Path & TEXTFILE)
            Dim Data As String = "-- This text will be in the file " & Path & TEXTFILE & " --"
            objWriter.WriteLine(Data)
            objWriter.Flush()
            objWriter.Close()
            Debug.Print(Data)
            Debug.Print("Successful...")
        Catch e As Exception
            Dim str As String = ""
            str = "An error occured while creating the Text file." & vbNewLine & e.Message
            MsgBox(str, vbOKOnly + vbExclamation, "Error")
            Return False
        End Try
        Return True
    End Function
End Module
