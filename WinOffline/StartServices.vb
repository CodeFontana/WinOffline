Imports System.ServiceProcess

Partial Public Class WinOffline

    Public Shared Function StartServices(ByVal CallStack As String) As Integer

        Dim ExecutionString As String                           ' Command line to be executed externally to the application.
        Dim ArgumentString As String                            ' Arguments passed on the command line for the external execution.
        Dim RunningProcess As Process                           ' A process shell for executing the command line.
        Dim ConsoleOutput As String                             ' Live output from an external application execution.
        Dim StandardOutput As String = ""                       ' Output from an external application execution.
        Dim RemainingOutput As String                           ' Additional output flushed after a process has exited.
        Dim PMUXService As ServiceController                    ' Port multiplexer service controller.
        Dim ProcessStartInfo As ProcessStartInfo                ' Process startup settings for configuring the bahviour of the process.
        Dim CAMDirectory As String                              ' CAM service installation directory.
        Dim CAFExitCode As Integer = 0                          ' Exit code from caf service startup.
        Dim LoopCounter As Integer = 0                          ' Re-usable loop counter.
        Dim RunLevel As Integer = 0                             ' Tracks the health of the function and calls to external functions.

        CallStack += "StartServices|"
        Logger.SetCurrentTask("Starting services..")

        ' Start the PMUX service
        Try
            If Not Utility.ServiceExists("CA-SAM-Pmux") Then
                Logger.WriteDebug(CallStack, "Port multiplexer service: NOT INSTALLED")
                Exit Try
            End If
            If Not Utility.IsServiceEnabled("CA-SAM-Pmux") Then
                Logger.WriteDebug(CallStack, "Port multiplexer service: DISABLED")
                Exit Try
            End If
            PMUXService = New ServiceController("CA-SAM-Pmux")
            If PMUXService.Status = ServiceControllerStatus.Running Then
                Logger.WriteDebug(CallStack, "Port multiplexer service: ACTIVE.")
            Else
                PMUXService.Start()
                Logger.WriteDebug(CallStack, "Port multiplexer service: STARTED.")
            End If
        Catch ex As Exception
            Logger.WriteDebug(CallStack, "Error: Exception caught while starting port multiplexer service.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})
        End Try

        ' Start the CAM service
        Try
            If Not Utility.ServiceExists("CA-MessageQueuing") Then
                Logger.WriteDebug(CallStack, "CAM service: NOT INSTALLED")
                Exit Try
            End If
            If Utility.IsProcessRunning("cam") Then
                Logger.WriteDebug(CallStack, "CAM service: ACTIVE")
                Exit Try
            End If
            CAMDirectory = Environment.GetEnvironmentVariable("CAI_MSQ")
            ExecutionString = CAMDirectory + "\bin\cam.exe"
            ArgumentString = "change auto"
            Logger.WriteDebug(CallStack, "Detached process: " + ExecutionString + " " + ArgumentString)
            ProcessStartInfo = New ProcessStartInfo(ExecutionString, ArgumentString)
            ProcessStartInfo.WorkingDirectory = CAMDirectory + "\bin"
            ProcessStartInfo.UseShellExecute = False
            ProcessStartInfo.RedirectStandardOutput = True
            ProcessStartInfo.CreateNoWindow = True
            StandardOutput = ""
            RemainingOutput = ""
            Logger.WriteDebug("------------------------------------------------------------")
            RunningProcess = Process.Start(ProcessStartInfo)
            While RunningProcess.HasExited = False
                ConsoleOutput = RunningProcess.StandardOutput.ReadLine
                Logger.WriteDebug(ConsoleOutput)
                StandardOutput += ConsoleOutput + Environment.NewLine
                Threading.Thread.Sleep(Globals.THREAD_REST_INTERVAL)
            End While
            RunningProcess.WaitForExit()
            RemainingOutput = RunningProcess.StandardOutput.ReadToEnd.ToString
            StandardOutput += RemainingOutput
            Logger.WriteDebug(RemainingOutput)
            Logger.WriteDebug("------------------------------------------------------------")
            Logger.WriteDebug(CallStack, "Exit code: " + RunningProcess.ExitCode.ToString)
            RunningProcess.Close()
            Logger.WriteDebug(CallStack, "CAM service: ENABLED")
            ExecutionString = CAMDirectory + "\bin\cam.exe"
            ArgumentString = "start -c -l"
            Logger.WriteDebug(CallStack, "Detached process: " + ExecutionString + " " + ArgumentString)
            ProcessStartInfo = New ProcessStartInfo(ExecutionString, ArgumentString)
            ProcessStartInfo.WorkingDirectory = CAMDirectory + "\bin"
            ProcessStartInfo.UseShellExecute = False
            ProcessStartInfo.RedirectStandardOutput = True
            ProcessStartInfo.CreateNoWindow = True
            StandardOutput = ""
            RemainingOutput = ""
            Logger.WriteDebug("------------------------------------------------------------")
            RunningProcess = Process.Start(ProcessStartInfo)
            While RunningProcess.HasExited = False
                ConsoleOutput = RunningProcess.StandardOutput.ReadLine
                Logger.WriteDebug(ConsoleOutput)
                StandardOutput += ConsoleOutput + Environment.NewLine
                Threading.Thread.Sleep(Globals.THREAD_REST_INTERVAL)
            End While
            RunningProcess.WaitForExit()
            RemainingOutput = RunningProcess.StandardOutput.ReadToEnd.ToString
            StandardOutput += RemainingOutput
            Logger.WriteDebug(RemainingOutput)
            Logger.WriteDebug("------------------------------------------------------------")
            Logger.WriteDebug(CallStack, "Exit code: " + RunningProcess.ExitCode.ToString)
            RunningProcess.Close()
            Logger.WriteDebug(CallStack, "CAM service: STARTED")
        Catch ex As Exception
            Logger.WriteDebug(CallStack, "Error: Exception caught while starting CAM service.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})
        End Try

        ' Start the CAF service
        Try
            If Not Utility.ServiceExists("caf") Then
                Logger.WriteDebug(CallStack, "CAF service: NOT INSTALLED")
                Exit Try
            End If
            If Not Utility.IsServiceEnabled("caf") Then
                Logger.WriteDebug(CallStack, "CAF service: DISABLED")
                Exit Try
            End If
            If Utility.IsProcessRunning("caf") Then
                Logger.WriteDebug(CallStack, "CAF service: ACTIVE")
                Exit Try
            End If
            If Globals.SkipCAFStartUpSwitch Then
                Logger.WriteDebug(CallStack, "CAF service: STOPPED")
                Exit Try
            End If
            ExecutionString = Globals.DSMFolder + "bin\caf.exe"
            ArgumentString = "start"
            Logger.WriteDebug(CallStack, "Detached process: " + ExecutionString + " " + ArgumentString)
            ProcessStartInfo = New ProcessStartInfo(ExecutionString, ArgumentString)
            ProcessStartInfo.WorkingDirectory = Globals.DSMFolder + "bin"
            ProcessStartInfo.UseShellExecute = False
            ProcessStartInfo.RedirectStandardOutput = True
            ProcessStartInfo.CreateNoWindow = True
            StandardOutput = ""
            RemainingOutput = ""
            Logger.WriteDebug("------------------------------------------------------------")
            RunningProcess = Process.Start(ProcessStartInfo)
            While RunningProcess.HasExited = False
                ConsoleOutput = RunningProcess.StandardOutput.ReadLine
                Logger.WriteDebug(ConsoleOutput)
                StandardOutput += ConsoleOutput + Environment.NewLine
                Threading.Thread.Sleep(Globals.THREAD_REST_INTERVAL)
            End While
            RunningProcess.WaitForExit()
            RemainingOutput = RunningProcess.StandardOutput.ReadToEnd.ToString
            StandardOutput += RemainingOutput
            Logger.WriteDebug(RemainingOutput)
            Logger.WriteDebug("------------------------------------------------------------")
            CAFExitCode = RunningProcess.ExitCode
            Logger.WriteDebug(CallStack, "Exit code: " + RunningProcess.ExitCode.ToString)
            RunningProcess.Close()
            Logger.WriteDebug(CallStack, "CAF service: STARTED")
        Catch ex As Exception
            Logger.WriteDebug(CallStack, "Error: Exception caught while starting CAF service.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})
            Return 1

        End Try

        ' Start the health monitoring service
        Try
            If Globals.SkiphmAgent Then
                Logger.WriteDebug(CallStack, "Health monitoring service: BYPASS")
                Exit Try
            End If
            If Not Utility.ServiceExists("hmAgent") Then
                Logger.WriteDebug(CallStack, "Health monitoring service: NOT INSTALLED")
                Exit Try
            End If
            If Not Utility.IsServiceEnabled("hmAgent") Then
                Logger.WriteDebug(CallStack, "Health monitoring service: DISABLED")
                Exit Try
            End If
            If Utility.IsProcessRunning("hmAgent") Then
                Logger.WriteDebug(CallStack, "Health monitoring service: ACTIVE")
                Exit Try
            End If
            ExecutionString = Globals.DSMFolder + "bin\hmagent.exe"
            ArgumentString = "start"
            Logger.WriteDebug(CallStack, "Detached process: " + ExecutionString + " " + ArgumentString)
            ProcessStartInfo = New ProcessStartInfo(ExecutionString, ArgumentString)
            ProcessStartInfo.WorkingDirectory = Globals.DSMFolder + "bin"
            ProcessStartInfo.UseShellExecute = False
            ProcessStartInfo.RedirectStandardOutput = True
            ProcessStartInfo.CreateNoWindow = True
            StandardOutput = ""
            RemainingOutput = ""
            Logger.WriteDebug("------------------------------------------------------------")
            RunningProcess = Process.Start(ProcessStartInfo)
            While RunningProcess.HasExited = False
                ConsoleOutput = RunningProcess.StandardOutput.ReadLine
                Logger.WriteDebug(ConsoleOutput)
                StandardOutput += ConsoleOutput + Environment.NewLine
                Threading.Thread.Sleep(Globals.THREAD_REST_INTERVAL)
            End While
            RunningProcess.WaitForExit()
            RemainingOutput = RunningProcess.StandardOutput.ReadToEnd.ToString
            StandardOutput += RemainingOutput
            Logger.WriteDebug(RemainingOutput)
            Logger.WriteDebug("------------------------------------------------------------")
            CAFExitCode = RunningProcess.ExitCode
            Logger.WriteDebug(CallStack, "Exit code: " + RunningProcess.ExitCode.ToString)
            RunningProcess.Close()
            Logger.WriteDebug(CallStack, "Health monitoring service: STARTED")
        Catch ex As Exception
            Logger.WriteDebug(CallStack, "Error: Exception caught while starting health monitoring service.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})
        End Try

        ' Start the system tray process
        Try
            RunLevel = LaunchPad(CallStack, "all", Globals.DSMFolder + "bin\cfsystray.exe", Globals.DSMFolder + "bin")
            If RunLevel = 0 Then
                Logger.WriteDebug(CallStack, "System tray: STARTED")
            Else
                Logger.WriteDebug(CallStack, "System tray: FAILED TO START")
            End If
        Catch ex As Exception
            Logger.WriteDebug(CallStack, "Error: Exception caught while starting system tray process.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})
        End Try

        ' Check caf startup exit code
        If CAFExitCode <> 0 Then
            Logger.WriteDebug(CallStack, "CAF service: WARNING -- Non-zero startup code detected.")
        End If

        Return CAFExitCode

    End Function

    Public Shared Function StartCAFOnDemand(ByVal CallStack As String) As Integer

        Dim ExecutionString As String                           ' Command line to be executed externally to the application.
        Dim ArgumentString As String                            ' Arguments passed on the command line for the external execution.
        Dim RunningProcess As Process                           ' A process shell for executing the command line.
        Dim ConsoleOutput As String                             ' Live output from an external application execution.
        Dim StandardOutput As String = ""                       ' Output from an external application execution.
        Dim RemainingOutput As String                           ' Additional output flushed after a process has exited.
        Dim PMUXService As ServiceController                    ' Port multiplexer service controller.
        Dim ProcessStartInfo As ProcessStartInfo                ' Process startup settings for configuring the bahviour of the process.
        Dim CAMDirectory As String                              ' CAM service installation directory.
        Dim CAFExitCode As Integer = 0                          ' Exit code from caf service startup.

        CallStack += "StartCAF|"

        ' Start the PMUX service
        Try
            If Not Utility.ServiceExists("CA-SAM-Pmux") Then
                Logger.WriteDebug(CallStack, "Port multiplexer service: NOT INSTALLED")
                Exit Try
            End If
            If Not Utility.IsServiceEnabled("CA-SAM-Pmux") Then
                Logger.WriteDebug(CallStack, "Port multiplexer service: DISABLED")
                Exit Try
            End If
            PMUXService = New ServiceController("CA-SAM-Pmux")
            If PMUXService.Status = ServiceControllerStatus.Running Then
                Logger.WriteDebug(CallStack, "Port multiplexer service: ACTIVE.")
            Else
                PMUXService.Start()
                Logger.WriteDebug(CallStack, "Port multiplexer service: STARTED.")
            End If
        Catch ex As Exception
            Logger.WriteDebug(CallStack, "Error: Exception caught while starting port multiplexer service.")
            Logger.WriteDebug(ex.Message)
        End Try

        ' Start the CAM service
        Try
            If Not Utility.ServiceExists("CA-MessageQueuing") Then
                Logger.WriteDebug(CallStack, "CAM service: NOT INSTALLED")
                Exit Try
            End If
            If Utility.IsProcessRunning("cam") Then
                Logger.WriteDebug(CallStack, "CAM service: ACTIVE")
                Exit Try
            End If
            CAMDirectory = Environment.GetEnvironmentVariable("CAI_MSQ")
            ExecutionString = CAMDirectory + "\bin\cam.exe"
            ArgumentString = "change auto"
            Logger.WriteDebug(CallStack, "Detached process: " + ExecutionString + " " + ArgumentString)
            ProcessStartInfo = New ProcessStartInfo(ExecutionString, ArgumentString)
            ProcessStartInfo.WorkingDirectory = CAMDirectory + "\bin"
            ProcessStartInfo.UseShellExecute = False
            ProcessStartInfo.RedirectStandardOutput = True
            ProcessStartInfo.CreateNoWindow = True
            StandardOutput = ""
            RemainingOutput = ""
            Logger.WriteDebug("------------------------------------------------------------")
            RunningProcess = Process.Start(ProcessStartInfo)
            While RunningProcess.HasExited = False
                ConsoleOutput = RunningProcess.StandardOutput.ReadLine
                Logger.WriteDebug(ConsoleOutput)
                StandardOutput += ConsoleOutput + Environment.NewLine
                Threading.Thread.Sleep(Globals.THREAD_REST_INTERVAL)
            End While
            RunningProcess.WaitForExit()
            RemainingOutput = RunningProcess.StandardOutput.ReadToEnd.ToString
            StandardOutput += RemainingOutput
            Logger.WriteDebug(RemainingOutput)
            Logger.WriteDebug("------------------------------------------------------------")
            Logger.WriteDebug(CallStack, "Exit code: " + RunningProcess.ExitCode.ToString)
            RunningProcess.Close()
            Logger.WriteDebug(CallStack, "CAM service: ENABLED")
            ExecutionString = CAMDirectory + "\bin\cam.exe"
            ArgumentString = "start -c -l"
            Logger.WriteDebug(CallStack, "Detached process: " + ExecutionString + " " + ArgumentString)
            ProcessStartInfo = New ProcessStartInfo(ExecutionString, ArgumentString)
            ProcessStartInfo.WorkingDirectory = CAMDirectory + "\bin"
            ProcessStartInfo.UseShellExecute = False
            ProcessStartInfo.RedirectStandardOutput = True
            ProcessStartInfo.CreateNoWindow = True
            StandardOutput = ""
            RemainingOutput = ""
            Logger.WriteDebug("------------------------------------------------------------")
            RunningProcess = Process.Start(ProcessStartInfo)
            While RunningProcess.HasExited = False
                ConsoleOutput = RunningProcess.StandardOutput.ReadLine
                Logger.WriteDebug(ConsoleOutput)
                StandardOutput += ConsoleOutput + Environment.NewLine
                Threading.Thread.Sleep(Globals.THREAD_REST_INTERVAL)
            End While
            RunningProcess.WaitForExit()
            RemainingOutput = RunningProcess.StandardOutput.ReadToEnd.ToString
            StandardOutput += RemainingOutput
            Logger.WriteDebug(RemainingOutput)
            Logger.WriteDebug("------------------------------------------------------------")
            Logger.WriteDebug(CallStack, "Exit code: " + RunningProcess.ExitCode.ToString)
            RunningProcess.Close()
            Logger.WriteDebug(CallStack, "CAM service: STARTED")
        Catch ex As Exception
            Logger.WriteDebug(CallStack, "Error: Exception caught while starting CAM service.")
            Logger.WriteDebug(ex.Message)
        End Try

        ' Start the CAF service

        Try
            If Not Utility.ServiceExists("caf") Then
                Logger.WriteDebug(CallStack, "CAF service: NOT INSTALLED")
                Exit Try
            End If
            If Not Utility.IsServiceEnabled("caf") Then
                Logger.WriteDebug(CallStack, "CAF service: DISABLED")
                Exit Try
            End If
            If Utility.IsProcessRunning("caf") Then
                Logger.WriteDebug(CallStack, "CAF service: ACTIVE")
                Exit Try
            End If
            If Globals.SkipCAFStartUpSwitch Then
                Logger.WriteDebug(CallStack, "CAF service: STOPPED")
                Exit Try
            End If
            ExecutionString = Globals.DSMFolder + "bin\caf.exe"
            ArgumentString = "start"
            Logger.WriteDebug(CallStack, "Detached process: " + ExecutionString + " " + ArgumentString)
            ProcessStartInfo = New ProcessStartInfo(ExecutionString, ArgumentString)
            ProcessStartInfo.WorkingDirectory = Globals.DSMFolder + "bin"
            ProcessStartInfo.UseShellExecute = False
            ProcessStartInfo.RedirectStandardOutput = True
            ProcessStartInfo.CreateNoWindow = True
            StandardOutput = ""
            RemainingOutput = ""
            Logger.WriteDebug("------------------------------------------------------------")
            RunningProcess = Process.Start(ProcessStartInfo)
            While RunningProcess.HasExited = False
                ConsoleOutput = RunningProcess.StandardOutput.ReadLine
                Logger.WriteDebug(ConsoleOutput)
                StandardOutput += ConsoleOutput + Environment.NewLine
                Threading.Thread.Sleep(Globals.THREAD_REST_INTERVAL)
            End While
            RunningProcess.WaitForExit()
            RemainingOutput = RunningProcess.StandardOutput.ReadToEnd.ToString
            StandardOutput += RemainingOutput
            Logger.WriteDebug(RemainingOutput)
            Logger.WriteDebug("------------------------------------------------------------")
            CAFExitCode = RunningProcess.ExitCode
            Logger.WriteDebug(CallStack, "Exit code: " + RunningProcess.ExitCode.ToString)
            RunningProcess.Close()
            Logger.WriteDebug(CallStack, "CAF service: STARTED")
        Catch ex As Exception
            Logger.WriteDebug(CallStack, "Error: Exception caught while starting CAF service.")
            Logger.WriteDebug(ex.Message)
        End Try

        Return CAFExitCode

    End Function

End Class