'****************************** Class Header *******************************\
' Project Name: WinOffline
' Class Name:   WinOffline
' File Name:    LaunchPad.vb
' Author:       Brian Fontana
'***************************************************************************/

Imports System.ServiceProcess

Partial Public Class WinOffline

    Public Shared Function LaunchPad(ByVal CallStack As String,
                                     ByVal LaunchContext As String,
                                     ByVal AppToLaunch As String,
                                     Optional ByVal WorkingDirectory As String = "",
                                     Optional ByVal AppArguments As String = "") As Integer

        ' Local variables
        Dim LaunchControl As ServiceController      ' Service controller for the launch process.
        Dim ExecutionString As String               ' Command line to be executed externally to the application.
        Dim ArgumentString As String                ' Arguments passed on the command line for the external execution.
        Dim ProcessStartInfo As ProcessStartInfo    ' Process startup settings for configuring the bahviour of the process.
        Dim RunningProcess As Process               ' A process shell for executing the command line.

        ' Update call stack
        CallStack += "LaunchPad|"

        ' *****************************
        ' - Launch Service pre-check.
        ' *****************************

        ' Check if the launch service is already installed
        If Utility.ServiceExists(Globals.ProcessFriendlyName + " Launch Service") Then

            ' Acquire launch service
            LaunchControl = New ServiceController(Globals.ProcessFriendlyName + " Launch Service")

            ' Build execution string
            ExecutionString = "sc.exe"
            ArgumentString = "delete """ + Globals.ProcessFriendlyName + " Launch Service"""

            ' Write debug
            Logger.WriteDebug(CallStack, "Detached process: " + ExecutionString + " " + ArgumentString)

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
            Logger.WriteDebug("------------------------------------------------------------")
            Logger.WriteDebug(RunningProcess.StandardOutput.ReadToEnd)
            Logger.WriteDebug("------------------------------------------------------------")

            ' Write debug
            Logger.WriteDebug(CallStack, "Exit code: " + RunningProcess.ExitCode.ToString)

            ' Verify installation was successful
            If (RunningProcess.ExitCode <> 0) Then

                ' Close detached process
                RunningProcess.Close()

                ' Return
                Return 1

            Else

                ' Close detached process
                RunningProcess.Close()

            End If

        End If

        ' Verify executable availability
        If Not System.IO.File.Exists(Globals.WinOfflineTemp + "\LaunchService.exe") Then

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Launch service executable not found.")

            ' Return
            Return 2

        End If

        ' If unspecified, assign global working directory
        If WorkingDirectory Is Nothing OrElse WorkingDirectory.Equals("") Then WorkingDirectory = Globals.WorkingDirectory

        ' Write debug
        Logger.WriteDebug(CallStack, "Launch Context: " + LaunchContext)
        Logger.WriteDebug(CallStack, "Application: " + AppToLaunch)
        Logger.WriteDebug(CallStack, "Working Directory: " + WorkingDirectory)
        Logger.WriteDebug(CallStack, "Arguments: " + AppArguments)

        ' *****************************
        ' - Install the Launch Service.
        ' *****************************

        ' Build execution string
        ExecutionString = "sc.exe"
        ArgumentString = "create """ + Globals.ProcessFriendlyName + " Launch Service"" binPath= """ +
                Globals.WinOfflineTemp + "\LaunchService.exe"" " +
                "start= demand"

        ' Write debug
        Logger.WriteDebug(CallStack, "Detached process: " + ExecutionString + " " + ArgumentString)

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
        Logger.WriteDebug("------------------------------------------------------------")
        Logger.WriteDebug(RunningProcess.StandardOutput.ReadToEnd)
        Logger.WriteDebug("------------------------------------------------------------")

        ' Write debug
        Logger.WriteDebug(CallStack, "Exit code: " + RunningProcess.ExitCode.ToString)

        ' Verify installation was successful
        If (RunningProcess.ExitCode <> 0) Then

            ' Close detached process
            RunningProcess.Close()

            ' Return
            Return 3

        Else

            ' Close detached process
            RunningProcess.Close()

        End If

        ' *****************************
        ' - Start the Launch Service.
        ' *****************************

        ' Assign service controller
        LaunchControl = New ServiceController(Globals.ProcessFriendlyName + " Launch Service")

        ' Write debug
        Logger.WriteDebug(CallStack, "Launching..")

        ' Check if arguments were passed
        If AppArguments Is Nothing OrElse AppArguments.Equals("") Then

            ' Start the service (no args)
            LaunchControl.Start({LaunchContext, AppToLaunch, WorkingDirectory})

        Else

            ' Start the service (with args)
            LaunchControl.Start({LaunchContext, AppToLaunch, WorkingDirectory, AppArguments})

        End If

        ' Wait for service to startup
        LaunchControl.WaitForStatus(ServiceControllerStatus.Stopped)

        ' Write debug
        Logger.WriteDebug(CallStack, "Mission completed.")

        ' *****************************
        ' - Delete the Launcher Service.
        ' *****************************

        ' Build execution string
        ExecutionString = "sc.exe"
        ArgumentString = "delete """ + Globals.ProcessFriendlyName + " Launch Service"""

        ' Write debug
        Logger.WriteDebug(CallStack, "Detached process: " + ExecutionString + " " + ArgumentString)

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
        Logger.WriteDebug("------------------------------------------------------------")
        Logger.WriteDebug(RunningProcess.StandardOutput.ReadToEnd)
        Logger.WriteDebug("------------------------------------------------------------")

        ' Write debug
        Logger.WriteDebug(CallStack, "Exit code: " + RunningProcess.ExitCode.ToString)

        ' Close detached process
        RunningProcess.Close()

        ' Return
        Return 0

    End Function

End Class