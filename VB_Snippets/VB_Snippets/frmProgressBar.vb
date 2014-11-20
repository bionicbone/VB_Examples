Imports Microsoft.WindowsAPICodePack
Imports Microsoft.WindowsAPICodePack.Taskbar


Public Class frmProgressBar

    ' *************************************************************************************
    ' *** Progress Bar Solution for XP / Windows 7                                      ***
    ' ***       Kevin Guest - 18/01/2013                                                ***
    ' ***    Windows XP, 7 , and 8.1 Compatible                                         ***
    ' ***                                                                               ***
    ' *** Under Windows XP only the Form Progress Bar will be displayed.                ***
    ' *** Under Windows 7 / 8 the Task Bar Icon will also be updated.                   ***
    ' ***                                                                               ***
    ' *** To import this example into local code you need to install the                ***
    ' *** Windows API Code Pack into the Current VB Program.                            ***
    ' ***   --- Tools Menu - NuGet Package Manager - Manage NuGet Packages for Solution ***
    ' ***   --- Install the Windows 7 API Code Pack.                                    ***
    ' *************************************************************************************

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        'Declare the Variables
        Dim osVer As Version = Environment.OSVersion.Version                            'Detect Windows Environment.
        Dim frmMainFormHandle As Double = Me.Handle()                                   ' The Handle of the main form.

        For N = 1 To 100
            ProgressBar1.Value = N

            If osVer.Major >= 6 And osVer.Minor >= 1 Then                               'Only allowed in Windows 7 or above
                TaskbarManager.Instance.SetProgressValue(N, 100, frmMainFormHandle)
            End If

            Application.DoEvents()
            Threading.Thread.Sleep(100)
        Next N

        ' Reset the Task Bar otherwise it will stay green.
        If osVer.Major >= 6 And osVer.Minor >= 1 Then                               'Only allowed in Windows 7 or above
            TaskbarManager.Instance.SetProgressState(TaskbarProgressBarState.Normal, frmMainFormHandle)
        End If

    End Sub
End Class
