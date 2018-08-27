Imports System.ServiceProcess
Imports System.Windows.Forms

Partial Public Class WinOfflineUI

    Public Sub LaunchPadUI(ByVal UIOutputControl As Control, ByVal LaunchContext As String, ByVal AppToLaunch As String, Optional ByVal AppArguments As String = "")

        ' Local variables
        Dim LaunchControl As ServiceController      ' Service controller for the launch process.
        Dim ExecutionString As String               ' Command line to be executed externally to the application.
        Dim ArgumentString As String                ' Arguments passed on the command line for the external execution.
        Dim ProcessStartInfo As ProcessStartInfo    ' Process startup settings for configuring the bahviour of the process.
        Dim RunningProcess As Process               ' A process shell for executing the command line.

        ' Write debug
        Delegate_Sub_Append_Text(UIOutputControl, "Launch Context: " + LaunchContext)
        Delegate_Sub_Append_Text(UIOutputControl, "Application: " + AppToLaunch)
        Delegate_Sub_Append_Text(UIOutputControl, "Arguments: " + AppArguments)

        ' *****************************
        ' - Launch Service pre-check.
        ' *****************************

        ' Check if the launch service is already installed
        If WinOffline.Utility.ServiceExists(Globals.ProcessFriendlyName + " Launch Service") Then

            ' Acquire launch service
            LaunchControl = New ServiceController(Globals.ProcessFriendlyName + " Launch Service")

            ' Build execution string
            ExecutionString = "sc.exe"
            ArgumentString = "delete """ + Globals.ProcessFriendlyName + " Launch Service"""

            ' Write debug
            Delegate_Sub_Append_Text(UIOutputControl, "Detached process: " + ExecutionString + " " + ArgumentString)

            ' Create detached process
            ProcessStartInfo = New ProcessStartInfo(ExecutionString, ArgumentString)
            ProcessStartInfo.WorkingDirectory = Globals.WorkingDirectory
            ProcessStartInfo.UseShellExecute = False
            ProcessStartInfo.RedirectStandardOutput = True
            ProcessStartInfo.CreateNoWindow = True

            ' Start detached process
            RunningProcess = Process.Start(ProcessStartInfo)

            ' Wait for detached process to exit
            RunningProcess.WaitForExit()

            ' Write debug
            Delegate_Sub_Append_Text(UIOutputControl, "------------------------------------------------------------")
            Delegate_Sub_Append_Text(UIOutputControl, RunningProcess.StandardOutput.ReadToEnd)
            Delegate_Sub_Append_Text(UIOutputControl, "------------------------------------------------------------")

            ' Write debug
            Delegate_Sub_Append_Text(UIOutputControl, "Exit code: " + RunningProcess.ExitCode.ToString)

            ' Verify installation was successful
            If (RunningProcess.ExitCode <> 0) Then

                ' Close detached process
                RunningProcess.Close()

                ' Return
                Return

            Else

                ' Close detached process
                RunningProcess.Close()

            End If

        End If

        ' *****************************
        ' - Install the Launch Service.
        ' *****************************

        ' Build execution string
        ExecutionString = "sc.exe"
        ArgumentString = "create """ + Globals.ProcessFriendlyName + " Launch Service"" binPath= """ +
                Globals.WinOfflineTemp + "\LaunchService.exe"" " +
                "start= demand"

        ' Write debug
        Delegate_Sub_Append_Text(UIOutputControl, "Detached process: " + ExecutionString + " " + ArgumentString)

        ' Create detached process
        ProcessStartInfo = New ProcessStartInfo(ExecutionString, ArgumentString)
        ProcessStartInfo.WorkingDirectory = Globals.WorkingDirectory
        ProcessStartInfo.UseShellExecute = False
        ProcessStartInfo.RedirectStandardOutput = True
        ProcessStartInfo.CreateNoWindow = True

        ' Start detached process
        RunningProcess = Process.Start(ProcessStartInfo)

        ' Wait for detached process to exit
        RunningProcess.WaitForExit()

        ' Write debug
        Delegate_Sub_Append_Text(UIOutputControl, "------------------------------------------------------------")
        Delegate_Sub_Append_Text(UIOutputControl, RunningProcess.StandardOutput.ReadToEnd)
        Delegate_Sub_Append_Text(UIOutputControl, "------------------------------------------------------------")

        ' Write debug
        Delegate_Sub_Append_Text(UIOutputControl, "Exit code: " + RunningProcess.ExitCode.ToString)

        ' Verify installation was successful
        If (RunningProcess.ExitCode <> 0) Then

            ' Close detached process
            RunningProcess.Close()

            ' Return
            Return

        Else

            ' Close detached process
            RunningProcess.Close()

        End If

        ' *****************************
        ' - Start the Launch Service.
        ' *****************************

        ' Assign service controller
        LaunchControl = New ServiceController(Globals.ProcessFriendlyName + " Launch Service")

        ' Start the service
        LaunchControl.Start({LaunchContext, AppToLaunch, AppArguments})

        ' Wait for service to startup
        LaunchControl.WaitForStatus(ServiceControllerStatus.Stopped)

        ' *****************************
        ' - Delete the Launcher Service.
        ' *****************************

        ' Build execution string
        ExecutionString = "sc.exe"
        ArgumentString = "delete """ + Globals.ProcessFriendlyName + " Launch Service"""

        ' Write debug
        Delegate_Sub_Append_Text(UIOutputControl, "Detached process: " + ExecutionString + " " + ArgumentString)

        ' Create detached process
        ProcessStartInfo = New ProcessStartInfo(ExecutionString, ArgumentString)
        ProcessStartInfo.WorkingDirectory = Globals.WorkingDirectory
        ProcessStartInfo.UseShellExecute = False
        ProcessStartInfo.RedirectStandardOutput = True
        ProcessStartInfo.CreateNoWindow = True

        ' Start detached process
        RunningProcess = Process.Start(ProcessStartInfo)

        ' Wait for detached process to exit
        RunningProcess.WaitForExit()

        ' Write debug
        Delegate_Sub_Append_Text(UIOutputControl, "------------------------------------------------------------")
        Delegate_Sub_Append_Text(UIOutputControl, RunningProcess.StandardOutput.ReadToEnd)
        Delegate_Sub_Append_Text(UIOutputControl, "------------------------------------------------------------")

        ' Write debug
        Delegate_Sub_Append_Text(UIOutputControl, "Exit code: " + RunningProcess.ExitCode.ToString)

        ' Close detached process
        RunningProcess.Close()

        ' Return
        Return

    End Sub

End Class