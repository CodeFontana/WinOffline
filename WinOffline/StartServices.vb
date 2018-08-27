Imports System.ServiceProcess

Partial Public Class WinOffline

    Public Shared Function StartServices(ByVal CallStack As String) As Integer

        ' Local variables
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

        ' Update call stack
        CallStack += "StartServices|"

        ' Write debug
        Logger.SetCurrentTask("Starting services..")

        ' *****************************
        ' - Start the PMUX service.
        ' *****************************

        Try

            ' Check if PMUX service exists
            If Not Utility.ServiceExists("CA-SAM-Pmux") Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Port multiplexer service: NOT INSTALLED")

                ' Skip this step
                Exit Try

            End If

            ' Check if PMUX service is disabled
            If Not Utility.IsServiceEnabled("CA-SAM-Pmux") Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Port multiplexer service: DISABLED")

                ' Skip this step
                Exit Try

            End If

            ' Create the service controller object
            PMUXService = New ServiceController("CA-SAM-Pmux")

            ' Check if service is already running
            If PMUXService.Status = ServiceControllerStatus.Running Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Port multiplexer service: ACTIVE.")

            Else

                ' Start the service
                PMUXService.Start()

                ' Write debug
                Logger.WriteDebug(CallStack, "Port multiplexer service: STARTED.")

            End If

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Exception caught while starting port multiplexer service.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

            ' Create exception
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

            ' Don't return because CAF is stopped
            ' Return 1

        End Try

        ' *****************************
        ' - Start the CAM service.
        ' *****************************

        Try

            ' Check if CAM service exists
            If Not Utility.ServiceExists("CA-MessageQueuing") Then

                ' Write debug
                Logger.WriteDebug(CallStack, "CAM service: NOT INSTALLED")

                ' Skip this step
                Exit Try

            End If

            ' Check if CAM service is already running
            If Utility.IsProcessRunning("cam") Then

                ' Write debug
                Logger.WriteDebug(CallStack, "CAM service: ACTIVE")

                ' Skip this step
                Exit Try

            End If

            ' Build execution string
            CAMDirectory = Environment.GetEnvironmentVariable("CAI_MSQ")
            ExecutionString = CAMDirectory + "\bin\cam.exe"
            ArgumentString = "change auto"

            ' Write debug
            Logger.WriteDebug(CallStack, "Detached process: " + ExecutionString + " " + ArgumentString)

            ' Create detached process
            ProcessStartInfo = New ProcessStartInfo(ExecutionString, ArgumentString)
            ProcessStartInfo.WorkingDirectory = CAMDirectory + "\bin"
            ProcessStartInfo.UseShellExecute = False
            ProcessStartInfo.RedirectStandardOutput = True
            ProcessStartInfo.CreateNoWindow = True

            ' Reset standard output
            StandardOutput = ""
            RemainingOutput = ""

            ' Write debug
            Logger.WriteDebug("------------------------------------------------------------")

            ' Start detached process
            RunningProcess = Process.Start(ProcessStartInfo)

            ' Read live output
            While RunningProcess.HasExited = False

                ' Read output
                ConsoleOutput = RunningProcess.StandardOutput.ReadLine

                ' Write debug
                Logger.WriteDebug(ConsoleOutput)

                ' Update standard output
                StandardOutput += ConsoleOutput + Environment.NewLine

                ' Rest thread
                Threading.Thread.Sleep(Globals.THREAD_REST_INTERVAL)

            End While

            ' Wait for detached process to exit
            RunningProcess.WaitForExit()

            ' Append remaining standard output
            RemainingOutput = RunningProcess.StandardOutput.ReadToEnd.ToString
            StandardOutput += RemainingOutput

            ' Write debug
            Logger.WriteDebug(RemainingOutput)
            Logger.WriteDebug("------------------------------------------------------------")
            Logger.WriteDebug(CallStack, "Exit code: " + RunningProcess.ExitCode.ToString)

            ' Close detached process
            RunningProcess.Close()

            ' Write debug
            Logger.WriteDebug(CallStack, "CAM service: ENABLED")

            ' Build execution string
            ExecutionString = CAMDirectory + "\bin\cam.exe"
            ArgumentString = "start -c -l"

            ' Write debug
            Logger.WriteDebug(CallStack, "Detached process: " + ExecutionString + " " + ArgumentString)

            ' Create detached process
            ProcessStartInfo = New ProcessStartInfo(ExecutionString, ArgumentString)
            ProcessStartInfo.WorkingDirectory = CAMDirectory + "\bin"
            ProcessStartInfo.UseShellExecute = False
            ProcessStartInfo.RedirectStandardOutput = True
            ProcessStartInfo.CreateNoWindow = True

            ' Reset standard output
            StandardOutput = ""
            RemainingOutput = ""

            ' Write debug
            Logger.WriteDebug("------------------------------------------------------------")

            ' Start detached process
            RunningProcess = Process.Start(ProcessStartInfo)

            ' Read live output
            While RunningProcess.HasExited = False

                ' Read output
                ConsoleOutput = RunningProcess.StandardOutput.ReadLine

                ' Write debug
                Logger.WriteDebug(ConsoleOutput)

                ' Update standard output
                StandardOutput += ConsoleOutput + Environment.NewLine

                ' Rest thread
                Threading.Thread.Sleep(Globals.THREAD_REST_INTERVAL)

            End While

            ' Wait for detached process to exit
            RunningProcess.WaitForExit()

            ' Append remaining standard output
            RemainingOutput = RunningProcess.StandardOutput.ReadToEnd.ToString
            StandardOutput += RemainingOutput

            ' Write debug
            Logger.WriteDebug(RemainingOutput)
            Logger.WriteDebug("------------------------------------------------------------")
            Logger.WriteDebug(CallStack, "Exit code: " + RunningProcess.ExitCode.ToString)

            ' Close detached process
            RunningProcess.Close()

            ' Write debug
            Logger.WriteDebug(CallStack, "CAM service: STARTED")

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Exception caught while starting CAM service.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

            ' Create exception
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

            ' Don't return because CAF is stopped.
            ' Return 2

        End Try

        ' *****************************
        ' - Start the CAF service.
        ' *****************************

        Try

            ' Check if CAF service exists
            If Not Utility.ServiceExists("caf") Then

                ' Write debug
                Logger.WriteDebug(CallStack, "CAF service: NOT INSTALLED")

                ' Skip this step
                Exit Try

            End If

            ' Check if CAF service is disabled
            If Not Utility.IsServiceEnabled("caf") Then

                ' Write debug
                Logger.WriteDebug(CallStack, "CAF service: DISABLED")

                ' Skip this step
                Exit Try

            End If

            ' Check if CAF service is already running
            If Utility.IsProcessRunning("caf") Then

                ' Write debug
                Logger.WriteDebug(CallStack, "CAF service: ACTIVE")

                ' Skip this step
                Exit Try

            End If

            ' Check CAF startup switch
            If Globals.SkipCAFStartUpSwitch Then

                ' Write debug
                Logger.WriteDebug(CallStack, "CAF service: STOPPED")

                ' Skip this step
                Exit Try

            End If

            ' Build execution string
            ExecutionString = Globals.DSMFolder + "bin\caf.exe"
            ArgumentString = "start"

            ' Write debug
            Logger.WriteDebug(CallStack, "Detached process: " + ExecutionString + " " + ArgumentString)

            ' Create detached process
            ProcessStartInfo = New ProcessStartInfo(ExecutionString, ArgumentString)
            ProcessStartInfo.WorkingDirectory = Globals.DSMFolder + "bin"
            ProcessStartInfo.UseShellExecute = False
            ProcessStartInfo.RedirectStandardOutput = True
            ProcessStartInfo.CreateNoWindow = True

            ' Reset standard output
            StandardOutput = ""
            RemainingOutput = ""

            ' Write debug
            Logger.WriteDebug("------------------------------------------------------------")

            ' Start detached process
            RunningProcess = Process.Start(ProcessStartInfo)

            ' Read live output
            While RunningProcess.HasExited = False

                ' Read output
                ConsoleOutput = RunningProcess.StandardOutput.ReadLine

                ' Write debug
                Logger.WriteDebug(ConsoleOutput)

                ' Update standard output
                StandardOutput += ConsoleOutput + Environment.NewLine

                ' Rest thread
                Threading.Thread.Sleep(Globals.THREAD_REST_INTERVAL)

            End While

            ' Wait for detached process to exit
            RunningProcess.WaitForExit()

            ' Append remaining standard output
            RemainingOutput = RunningProcess.StandardOutput.ReadToEnd.ToString
            StandardOutput += RemainingOutput

            ' Write debug
            Logger.WriteDebug(RemainingOutput)
            Logger.WriteDebug("------------------------------------------------------------")

            ' Assign exit code
            CAFExitCode = RunningProcess.ExitCode

            ' Write debug
            Logger.WriteDebug(CallStack, "Exit code: " + RunningProcess.ExitCode.ToString)

            ' Close detached process
            RunningProcess.Close()

            ' Write debug
            Logger.WriteDebug(CallStack, "CAF service: STARTED")

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Exception caught while starting CAF service.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

            ' Create exception
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

            ' Return
            Return 5

        End Try

        ' *****************************
        ' - Start the health monitoring service.
        ' *****************************

        Try

            ' Check if hm service was bypassed
            If Globals.SkiphmAgent Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Health monitoring service: BYPASS")

                ' Skip this step
                Exit Try

            End If

            ' Check if hm service exists
            If Not Utility.ServiceExists("hmAgent") Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Health monitoring service: NOT INSTALLED")

                ' Skip this step
                Exit Try

            End If

            ' Check if hm service is disabled
            If Not Utility.IsServiceEnabled("hmAgent") Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Health monitoring service: DISABLED")

                ' Skip this step
                Exit Try

            End If

            ' Check if hm service is already started
            If Utility.IsProcessRunning("hmAgent") Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Health monitoring service: ACTIVE")

                ' Skip this step
                Exit Try

            End If

            ' Build execution string
            ExecutionString = Globals.DSMFolder + "bin\hmagent.exe"
            ArgumentString = "start"

            ' Write debug
            Logger.WriteDebug(CallStack, "Detached process: " + ExecutionString + " " + ArgumentString)

            ' Create detached process
            ProcessStartInfo = New ProcessStartInfo(ExecutionString, ArgumentString)
            ProcessStartInfo.WorkingDirectory = Globals.DSMFolder + "bin"
            ProcessStartInfo.UseShellExecute = False
            ProcessStartInfo.RedirectStandardOutput = True
            ProcessStartInfo.CreateNoWindow = True

            ' Reset standard output
            StandardOutput = ""
            RemainingOutput = ""

            ' Write debug
            Logger.WriteDebug("------------------------------------------------------------")

            ' Start detached process
            RunningProcess = Process.Start(ProcessStartInfo)

            ' Read live output
            While RunningProcess.HasExited = False

                ' Read output
                ConsoleOutput = RunningProcess.StandardOutput.ReadLine

                ' Write debug
                Logger.WriteDebug(ConsoleOutput)

                ' Update standard output
                StandardOutput += ConsoleOutput + Environment.NewLine

                ' Rest thread
                Threading.Thread.Sleep(Globals.THREAD_REST_INTERVAL)

            End While

            ' Wait for detached process to exit
            RunningProcess.WaitForExit()

            ' Append remaining standard output
            RemainingOutput = RunningProcess.StandardOutput.ReadToEnd.ToString
            StandardOutput += RemainingOutput

            ' Write debug
            Logger.WriteDebug(RemainingOutput)
            Logger.WriteDebug("------------------------------------------------------------")

            ' Assign exit code
            CAFExitCode = RunningProcess.ExitCode

            ' Write debug
            Logger.WriteDebug(CallStack, "Exit code: " + RunningProcess.ExitCode.ToString)

            ' Close detached process
            RunningProcess.Close()

            ' Write debug
            Logger.WriteDebug(CallStack, "Health monitoring service: STARTED")

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Exception caught while starting health monitoring service.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

            ' Create exception
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

            ' Don't return
            ' Return 8

        End Try

        ' *****************************
        ' - Start the system tray process.
        ' *****************************

        Try

            ' Call the launch pad
            RunLevel = LaunchPad(CallStack, "all", Globals.DSMFolder + "bin\cfsystray.exe", Globals.DSMFolder + "bin")

            ' Check runlevel
            If RunLevel = 0 Then

                ' Write debug
                Logger.WriteDebug(CallStack, "System tray: STARTED")

            Else

                ' Write debug
                Logger.WriteDebug(CallStack, "System tray: FAILED TO START")

            End If

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Exception caught while starting system tray process.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

            ' Create exception
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

            ' Don't return
            ' Return 10

        End Try

        ' Check caf startup exit code
        If CAFExitCode <> 0 Then

            ' Write debug
            Logger.WriteDebug(CallStack, "CAF service: WARNING -- Non-zero startup code detected.")

        End If

        ' Return
        Return CAFExitCode

    End Function

    Public Shared Function StartCAFOnDemand(ByVal CallStack As String) As Integer

        ' Local variables
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

        ' Update call stack
        CallStack += "StartCAF|"

        ' *****************************
        ' - Start the PMUX service.
        ' *****************************

        Try

            ' Check if PMUX service exists
            If Not Utility.ServiceExists("CA-SAM-Pmux") Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Port multiplexer service: NOT INSTALLED")

                ' Skip this step
                Exit Try

            End If

            ' Check if PMUX service is disabled
            If Not Utility.IsServiceEnabled("CA-SAM-Pmux") Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Port multiplexer service: DISABLED")

                ' Skip this step
                Exit Try

            End If

            ' Create the service controller object
            PMUXService = New ServiceController("CA-SAM-Pmux")

            ' Check if service is already running
            If PMUXService.Status = ServiceControllerStatus.Running Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Port multiplexer service: ACTIVE.")

            Else

                ' Start the service
                PMUXService.Start()

                ' Write debug
                Logger.WriteDebug(CallStack, "Port multiplexer service: STARTED.")

            End If

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Exception caught while starting port multiplexer service.")
            Logger.WriteDebug(ex.Message)

        End Try

        ' *****************************
        ' - Start the CAM service.
        ' *****************************

        Try

            ' Check if CAM service exists
            If Not Utility.ServiceExists("CA-MessageQueuing") Then

                ' Write debug
                Logger.WriteDebug(CallStack, "CAM service: NOT INSTALLED")

                ' Skip this step
                Exit Try

            End If

            ' Check if CAM service is already running
            If Utility.IsProcessRunning("cam") Then

                ' Write debug
                Logger.WriteDebug(CallStack, "CAM service: ACTIVE")

                ' Skip this step
                Exit Try

            End If

            ' Build execution string
            CAMDirectory = Environment.GetEnvironmentVariable("CAI_MSQ")
            ExecutionString = CAMDirectory + "\bin\cam.exe"
            ArgumentString = "change auto"

            ' Write debug
            Logger.WriteDebug(CallStack, "Detached process: " + ExecutionString + " " + ArgumentString)

            ' Create detached process
            ProcessStartInfo = New ProcessStartInfo(ExecutionString, ArgumentString)
            ProcessStartInfo.WorkingDirectory = CAMDirectory + "\bin"
            ProcessStartInfo.UseShellExecute = False
            ProcessStartInfo.RedirectStandardOutput = True
            ProcessStartInfo.CreateNoWindow = True

            ' Reset standard output
            StandardOutput = ""
            RemainingOutput = ""

            ' Write debug
            Logger.WriteDebug("------------------------------------------------------------")

            ' Start detached process
            RunningProcess = Process.Start(ProcessStartInfo)

            ' Read live output
            While RunningProcess.HasExited = False

                ' Read output
                ConsoleOutput = RunningProcess.StandardOutput.ReadLine

                ' Write debug
                Logger.WriteDebug(ConsoleOutput)

                ' Update standard output
                StandardOutput += ConsoleOutput + Environment.NewLine

                ' Rest thread
                Threading.Thread.Sleep(Globals.THREAD_REST_INTERVAL)

            End While

            ' Wait for detached process to exit
            RunningProcess.WaitForExit()

            ' Append remaining standard output
            RemainingOutput = RunningProcess.StandardOutput.ReadToEnd.ToString
            StandardOutput += RemainingOutput

            ' Write debug
            Logger.WriteDebug(RemainingOutput)
            Logger.WriteDebug("------------------------------------------------------------")
            Logger.WriteDebug(CallStack, "Exit code: " + RunningProcess.ExitCode.ToString)

            ' Close detached process
            RunningProcess.Close()

            ' Write debug
            Logger.WriteDebug(CallStack, "CAM service: ENABLED")

            ' Build execution string
            ExecutionString = CAMDirectory + "\bin\cam.exe"
            ArgumentString = "start -c -l"

            ' Write debug
            Logger.WriteDebug(CallStack, "Detached process: " + ExecutionString + " " + ArgumentString)

            ' Create detached process
            ProcessStartInfo = New ProcessStartInfo(ExecutionString, ArgumentString)
            ProcessStartInfo.WorkingDirectory = CAMDirectory + "\bin"
            ProcessStartInfo.UseShellExecute = False
            ProcessStartInfo.RedirectStandardOutput = True
            ProcessStartInfo.CreateNoWindow = True

            ' Reset standard output
            StandardOutput = ""
            RemainingOutput = ""

            ' Write debug
            Logger.WriteDebug("------------------------------------------------------------")

            ' Start detached process
            RunningProcess = Process.Start(ProcessStartInfo)

            ' Read live output
            While RunningProcess.HasExited = False

                ' Read output
                ConsoleOutput = RunningProcess.StandardOutput.ReadLine

                ' Write debug
                Logger.WriteDebug(ConsoleOutput)

                ' Update standard output
                StandardOutput += ConsoleOutput + Environment.NewLine

                ' Rest thread
                Threading.Thread.Sleep(Globals.THREAD_REST_INTERVAL)

            End While

            ' Wait for detached process to exit
            RunningProcess.WaitForExit()

            ' Append remaining standard output
            RemainingOutput = RunningProcess.StandardOutput.ReadToEnd.ToString
            StandardOutput += RemainingOutput

            ' Write debug
            Logger.WriteDebug(RemainingOutput)
            Logger.WriteDebug("------------------------------------------------------------")
            Logger.WriteDebug(CallStack, "Exit code: " + RunningProcess.ExitCode.ToString)

            ' Close detached process
            RunningProcess.Close()

            ' Write debug
            Logger.WriteDebug(CallStack, "CAM service: STARTED")

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Exception caught while starting CAM service.")
            Logger.WriteDebug(ex.Message)

        End Try

        ' *****************************
        ' - Start the CAF service.
        ' *****************************

        Try

            ' Check if CAF service exists
            If Not Utility.ServiceExists("caf") Then

                ' Write debug
                Logger.WriteDebug(CallStack, "CAF service: NOT INSTALLED")

                ' Skip this step
                Exit Try

            End If

            ' Check if CAF service is disabled
            If Not Utility.IsServiceEnabled("caf") Then

                ' Write debug
                Logger.WriteDebug(CallStack, "CAF service: DISABLED")

                ' Skip this step
                Exit Try

            End If

            ' Check if CAF service is already running
            If Utility.IsProcessRunning("caf") Then

                ' Write debug
                Logger.WriteDebug(CallStack, "CAF service: ACTIVE")

                ' Skip this step
                Exit Try

            End If

            ' Check CAF startup switch
            If Globals.SkipCAFStartUpSwitch Then

                ' Write debug
                Logger.WriteDebug(CallStack, "CAF service: STOPPED")

                ' Skip this step
                Exit Try

            End If

            ' Build execution string
            ExecutionString = Globals.DSMFolder + "bin\caf.exe"
            ArgumentString = "start"

            ' Write debug
            Logger.WriteDebug(CallStack, "Detached process: " + ExecutionString + " " + ArgumentString)

            ' Create detached process
            ProcessStartInfo = New ProcessStartInfo(ExecutionString, ArgumentString)
            ProcessStartInfo.WorkingDirectory = Globals.DSMFolder + "bin"
            ProcessStartInfo.UseShellExecute = False
            ProcessStartInfo.RedirectStandardOutput = True
            ProcessStartInfo.CreateNoWindow = True

            ' Reset standard output
            StandardOutput = ""
            RemainingOutput = ""

            ' Write debug
            Logger.WriteDebug("------------------------------------------------------------")

            ' Start detached process
            RunningProcess = Process.Start(ProcessStartInfo)

            ' Read live output
            While RunningProcess.HasExited = False

                ' Read output
                ConsoleOutput = RunningProcess.StandardOutput.ReadLine

                ' Write debug
                Logger.WriteDebug(ConsoleOutput)

                ' Update standard output
                StandardOutput += ConsoleOutput + Environment.NewLine

                ' Rest thread
                Threading.Thread.Sleep(Globals.THREAD_REST_INTERVAL)

            End While

            ' Wait for detached process to exit
            RunningProcess.WaitForExit()

            ' Append remaining standard output
            RemainingOutput = RunningProcess.StandardOutput.ReadToEnd.ToString
            StandardOutput += RemainingOutput

            ' Write debug
            Logger.WriteDebug(RemainingOutput)
            Logger.WriteDebug("------------------------------------------------------------")

            ' Assign exit code
            CAFExitCode = RunningProcess.ExitCode

            ' Write debug
            Logger.WriteDebug(CallStack, "Exit code: " + RunningProcess.ExitCode.ToString)

            ' Close detached process
            RunningProcess.Close()

            ' Write debug
            Logger.WriteDebug(CallStack, "CAF service: STARTED")

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Exception caught while starting CAF service.")
            Logger.WriteDebug(ex.Message)

        End Try

        ' Return
        Return CAFExitCode

    End Function

End Class