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
    ' ***                                                                               ***
    ' *** To test this example:                                                         ***
    ' *** --- Create a Form called frmMainFormHandle                                    ***
    ' *** --- Add a form Progress Bar                                                   ***
    ' *** --- Paste this code                                                           ***
    ' *************************************************************************************

    ' Create a public variable to hold the handle of the form.
    ' this is the form on the Task Bar that will show the progress.
    Public frmMainFormHandle As Double = 0


    Private Sub frmProgressBar_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' For Windows 7 or higher it is very important to catch the handle of the form
        ' so we can point to the correct instance on the TaskBar Progress Bar.
        frmMainFormHandle = Me.Handle()
    End Sub

    Private Sub frmProgressBar_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated

        ' This is not required, it's just part of the example to force the form to diplay in the Activated Event.
        Me.Show()

        ' Declare the Local Variables
        Dim osVer As Version = Environment.OSVersion.Version                    'Detect Windows Environment.

        For N = 1 To 100

            ' This would update a standard progress bar on the form.
            ProgressBar1.Value = N

            ' NOTE: NEVER use the -1 fix to make the progress bar keep up with code!
            '       The -1 fix is extremely resourse hungry and WILL slow down fast code by 500%-1000%

            If osVer.Major >= 6 And osVer.Minor >= 1 Then                       'Only allowed in Windows 7 or above
                ' This updates the Task Bar from icon to show the progress.
                TaskbarManager.Instance.SetProgressValue(N, 100, frmMainFormHandle)
            End If

            ' This is not required, it's just part of this example and to slow everything down.
            Me.Refresh()
            Threading.Thread.Sleep(200)
        Next N

        ' Standard Form Progress Bars are notorious for not keeping up with code.
        ' Try altering the above Sleep to just 20, note the Task Bar progress keeps up but the form one does not!
        ' These five lines just makes sure the user see's 100% in the form before leaving the code.
        ProgressBar1.Value = 100
        ProgressBar1.Value = 99
        ProgressBar1.Value = 100
        Me.Refresh()
        Threading.Thread.Sleep(1000)

        If osVer.Major >= 6 And osVer.Minor >= 1 Then                           'Only allowed in Windows 7 or above
            ' Reset the Task Bar otherwise it will stay 100% green!
            TaskbarManager.Instance.SetProgressState(TaskbarProgressBarState.Normal, frmMainFormHandle)

            ' NOTE:
            ' Also see TaskbarProgressBarState.Error, TaskbarProgressBarState.Paused, etc. for different colours / states.

        End If

        ' This is not required, it's just part of the example to force the form to close in the Activated Event.
        Me.Close()

    End Sub
End Class
