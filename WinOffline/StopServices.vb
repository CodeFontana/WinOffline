'****************************** Class Header *******************************\
' Project Name: WinOffline
' Class Name:   WinOffline
' File Name:    StopServices.vb
' Author:       Brian Fontana
'***************************************************************************/

Imports System.ServiceProcess
Imports System.Threading

Partial Public Class WinOffline

    Public Shared Function StopServices(ByVal CallStack As String) As Integer

        ' Local variables
        Dim ExecutionString As String               ' Command line to be executed externally to the application.
        Dim ArgumentString As String                ' Arguments passed on the command line for the external execution.
        Dim RunningProcess As Process               ' A process shell for executing the command line.
        Dim ConsoleOutput As String                 ' Live output from an external application execution.
        Dim StandardOutput As String                ' Output from an external application execution.
        Dim RemainingOutput As String               ' Additional output flushed after a process has exited.
        Dim PMLAService As ServiceController        ' Performance lite agent service controller.
        Dim ProcessStartInfo As ProcessStartInfo    ' Process startup settings for configuring the bahviour of the process.
        Dim CAMDirectory As String                  ' CAM service installation directory.
        Dim CAFExitCode As Integer                  ' CAF command exit code.
        Dim CAFStopHelperThread As Thread = Nothing ' Helper thread for expediting "caf stop" operation. 
        Dim LoopCounter As Integer = 0              ' Re-usable loop counter.
        Dim RunLevel As Integer = 0                 ' Tracks the health of the function and calls to external functions.

        ' Update call stack
        CallStack += "StopServices|"

        ' Write debug
        Logger.SetCurrentTask("Stopping services..")

        ' *****************************
        ' - Stop health monitoring agent.
        ' *****************************

        Try

            ' Check for running hmagent service
            If Utility.IsProcessRunning("hmagent") And Not Globals.SkiphmAgent Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Health monitoring service: ACTIVE")

                ' Build execution string
                ExecutionString = Globals.DSMFolder + "bin\hmagent.exe"
                ArgumentString = "stop"

                ' Write debug
                Logger.WriteDebug(CallStack, "Detached process: " + ExecutionString + " " + ArgumentString)

                ' Create detached process
                ProcessStartInfo = New ProcessStartInfo(ExecutionString, ArgumentString)
                ProcessStartInfo.WorkingDirectory = ExecutionString.Substring(0, ExecutionString.LastIndexOf("\"))
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
                    Thread.Sleep(Globals.THREAD_REST_INTERVAL)

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

                ' Delay execution
                Thread.Sleep(2000)

                ' Check for hmagent process
                If Utility.IsProcessRunning("hmagent") Then

                    ' Reset the loop counter
                    LoopCounter = 0

                    ' Wait for hmagent to stop
                    While Utility.IsProcessRunning("hmagent")

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Health monitoring service: STOPPING")

                        ' Delay execution
                        Thread.Sleep(5000)

                        ' Increment the loop counter
                        LoopCounter += 1

                        ' Check the loop counter
                        If LoopCounter >= 3 Then

                            ' Get process listing
                            Dim TrayProcess() As Process = System.Diagnostics.Process.GetProcessesByName("hmagent")

                            ' Process the list
                            For Each pProcess As Process In TrayProcess

                                ' Write debug
                                Logger.WriteDebug(CallStack, "Kill process: " + pProcess.ProcessName.ToString + "," + pProcess.Id.ToString)

                                ' Kill process
                                pProcess.Kill()

                            Next

                            ' Write debug
                            Logger.WriteDebug(CallStack, "Health monitoring service: TERMINATED")

                            ' Infinite loop safety -- Exit while condition
                            Exit While

                        End If

                    End While

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Health monitoring service: STOPPED")

                Else

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Health monitoring service: STOPPED")

                End If

            ElseIf Globals.SkiphmAgent Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Health monitoring service: BYPASS")

            Else

                ' Write debug
                Logger.WriteDebug(CallStack, "Health monitoring service: INACTIVE")

            End If

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Exception caught stopping the health monitoring service.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

            ' Create exception
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

            ' Return
            Return 1

        End Try

        ' *****************************
        ' - Stop alert collector service.
        ' *****************************

        Try

            ' Check for running alert collector service
            If Utility.IsProcessRunning("AlertCollector") And Not Globals.SkiphmAgent Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Alert collector service: ACTIVE")

                ' Build execution string
                ExecutionString = Globals.DSMFolder + "bin\AlertCollector.exe"
                ArgumentString = "stop"

                ' Write debug
                Logger.WriteDebug(CallStack, "Detached process: " + ExecutionString + " " + ArgumentString)

                ' Create detached process
                ProcessStartInfo = New ProcessStartInfo(ExecutionString, ArgumentString)
                ProcessStartInfo.WorkingDirectory = ExecutionString.Substring(0, ExecutionString.LastIndexOf("\"))
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
                    Thread.Sleep(Globals.THREAD_REST_INTERVAL)

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

                ' Delay execution
                Thread.Sleep(2000)

                ' Check for alertcollector process
                If Utility.IsProcessRunning("AlertCollector") Then

                    ' Reset the loop counter
                    LoopCounter = 0

                    ' Wait for alert collector to stop
                    While Utility.IsProcessRunning("AlertCollector")

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Alert collector service: STOPPING")

                        ' Delay execution
                        Thread.Sleep(5000)

                        ' Increment the loop counter
                        LoopCounter += 1

                        ' Check the loop counter
                        If LoopCounter >= 3 Then

                            ' Get process listing
                            Dim TrayProcess() As Process = System.Diagnostics.Process.GetProcessesByName("AlertCollector")

                            ' Process the list
                            For Each pProcess As Process In TrayProcess

                                ' Write debug
                                Logger.WriteDebug(CallStack, "Kill process: " + pProcess.ProcessName.ToString + "," + pProcess.Id.ToString)

                                ' Kill process
                                pProcess.Kill()

                            Next

                            ' Write debug
                            Logger.WriteDebug(CallStack, "Alert collector service: TERMINATED")

                            ' Infinite loop safety -- Exit while condition
                            Exit While

                        End If

                    End While

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Alert collector service: STOPPED")

                Else

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Alert collector service: STOPPED")

                End If

            ElseIf Globals.SkiphmAgent Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Alert collector service: BYPASS")

            Else

                ' Write debug
                Logger.WriteDebug(CallStack, "Alert collector service: INACTIVE")

            End If

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Exception caught stopping the alert collector service.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

            ' Create exception
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

            ' Return
            Return 2

        End Try

        ' *****************************
        ' - Terminate any active software delivery execution.
        ' *****************************

        Try

            ' Check for running software delivery executioner
            If Utility.IsProcessRunning("sd_jexec") Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Software delivery plugin: ACTIVE")

                ' Delay execution
                Thread.Sleep(5000)

                ' Check for additional SD execution processes
                If Utility.IsProcessRunning("sd_jexec") Then

                    ' Reset the loop counter
                    LoopCounter = 0

                    ' Wait for the SD execution to finish
                    While Utility.IsProcessRunning("sd_jexec")

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Software delivery plugin: WAITING")

                        ' Avoid deadlock for operations longer than 60 seconds
                        System.Windows.Forms.Application.DoEvents()

                        ' Delay execution
                        Thread.Sleep(5000)

                        ' Increment the loop counter
                        LoopCounter += 1

                        ' Check the loop counter
                        If LoopCounter >= 6 Then

                            ' Get process listing
                            Dim SDPluginProcesses() As Process = System.Diagnostics.Process.GetProcessesByName("sd_jexec")

                            ' Process the list
                            For Each pProcess As Process In SDPluginProcesses

                                ' Write debug
                                Logger.WriteDebug(CallStack, "Kill process: " + pProcess.ProcessName.ToString + "," + pProcess.Id.ToString)

                                ' Kill process
                                pProcess.Kill()

                            Next

                            ' Write debug
                            Logger.WriteDebug(CallStack, "Software delivery plugin: TERMINATED")

                            ' Infinite loop safety -- Exit while condition
                            Exit While

                        End If

                    End While

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Software delivery plugin: STOPPED")

                Else

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Software delivery plugin: STOPPED")

                End If

            Else

                ' Write debug
                Logger.WriteDebug(CallStack, "Software delivery plugin: INACTIVE")

            End If

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Exception caught terminating the software delivery plugin.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

            ' Create exception
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

            ' Return
            Return 3

        End Try

        ' *****************************
        ' - Stop DSM Explorer processes.
        ' *****************************

        Try

            ' Check for running explorer processes
            If Utility.IsProcessRunning("egc30n") Then

                ' Write debug
                Logger.WriteDebug(CallStack, "DSM Explorer: ACTIVE")

                ' Get process listing
                Dim ExplorerProcesses() As Process = System.Diagnostics.Process.GetProcessesByName("EGC30N")

                ' Process the list
                For Each pProcess As Process In ExplorerProcesses

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Terminate process: " + pProcess.ProcessName.ToString + "," + pProcess.Id.ToString)

                    ' Kill processes
                    pProcess.Kill()

                Next

                ' Write debug
                Logger.WriteDebug(CallStack, "DSM Explorer: TERMINATED")

            Else

                ' Write debug
                Logger.WriteDebug(CallStack, "DSM Explorer: INACTIVE")

            End If

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Exception caught terminating DSM Explorer processes.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

            ' Create exception
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

            ' Return
            Return 4

        End Try

        ' *****************************
        ' - Stop the dmprimer service.
        ' *****************************

        Try

            ' Check for running dmprimer service
            If Utility.IsProcessRunning("dm_primer") AndAlso Not Globals.SkipDMPrimer Then

                ' Write debug
                Logger.WriteDebug(CallStack, "DMPrimer service: ACTIVE")

                ' Set the service marker
                ' No need for a marker -- this service is only used by the
                ' deployment wizard. It will restart automatically if the 
                ' deployment wizard scans the agent.

                ' Verify primer path is available
                If Globals.DMPrimerFolder IsNot Nothing Then

                    ' Build execution string
                    ExecutionString = Globals.DMPrimerFolder + "dm_primer.exe"
                    ArgumentString = "stop"

                Else

                    ' Build execution string
                    ExecutionString = Utility.GetProcessFileName("dm_primer")
                    ArgumentString = "stop"

                End If

                ' Write debug
                Logger.WriteDebug(CallStack, "Detached process: " + ExecutionString + " " + ArgumentString)

                ' Create detached process
                ProcessStartInfo = New ProcessStartInfo(ExecutionString, ArgumentString)
                ProcessStartInfo.WorkingDirectory = ExecutionString.Substring(0, ExecutionString.LastIndexOf("\"))
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

                ' Delay execution
                Thread.Sleep(2000)

                ' Check for primer process
                If Utility.IsProcessRunning("dm_primer") Then

                    ' Reset the loop counter
                    LoopCounter = 0

                    ' Wait for dmprimer to stop
                    While Utility.IsProcessRunning("dm_primer")

                        ' Write debug
                        Logger.WriteDebug(CallStack, "DMPrimer service: STOPPING")

                        ' Delay execution
                        Thread.Sleep(5000)

                        ' Increment the loop counter
                        LoopCounter += 1

                        ' Check the loop counter
                        If LoopCounter >= 3 Then

                            ' Get process listing
                            Dim TrayProcess() As Process = System.Diagnostics.Process.GetProcessesByName("dm_primer")

                            ' Process the list
                            For Each pProcess As Process In TrayProcess

                                ' Write debug
                                Logger.WriteDebug(CallStack, "Kill process: " + pProcess.ProcessName.ToString + "," + pProcess.Id.ToString)

                                ' Kill process
                                pProcess.Kill()

                            Next

                            ' Write debug
                            Logger.WriteDebug(CallStack, "DMPrimer service: TERMINATED")

                            ' Infinite loop safety -- Exit while condition
                            Exit While

                        End If

                    End While

                    ' Write debug
                    Logger.WriteDebug(CallStack, "DMPrimer service: STOPPED")

                Else

                    ' Write debug
                    Logger.WriteDebug(CallStack, "DMPrimer service: STOPPED")

                End If

            ElseIf Globals.SkipDMPrimer Then

                ' Write debug
                Logger.WriteDebug(CallStack, "DMPrimer service: SKIPPED")

            Else

                ' Write debug
                Logger.WriteDebug(CallStack, "DMPrimer service: INACTIVE")

            End If

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Exception caught stopping the dmprimer service.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

            ' Create exception
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

            ' Return
            Return 5

        End Try

        ' *****************************
        ' - Stop the performance lite agent service.
        ' *****************************

        Try

            ' Check for PMLA process
            If Utility.IsProcessRunning("casplitegent") Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Performance lite agent service: ACTIVE")

                ' Create the service controller object
                PMLAService = New ServiceController("CASPLiteAgent")

                ' Stop the service
                PMLAService.Stop()

                ' Write debug
                Logger.WriteDebug(CallStack, "Performance lite agent service: STOPPING")

                ' Delay execution
                Thread.Sleep(2000)

                ' *****************************
                ' - Verify performance lite agent is down.
                ' *****************************

                ' Check for performance lite agent process
                If Utility.IsProcessRunning("casplitegent") Then

                    ' Reset the loop counter
                    LoopCounter = 0

                    ' Wait for PMLA to stop
                    While Utility.IsProcessRunning("casplitegent")

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Performance lite agent service: STOPPING")

                        ' Delay execution
                        Thread.Sleep(5000)

                        ' Increment the loop counter
                        LoopCounter += 1

                        ' Check the loop counter
                        If LoopCounter >= 3 Then

                            ' Get process listing
                            Dim PerfProcesses() As Process = System.Diagnostics.Process.GetProcessesByName("casplitegent")

                            ' Process the list
                            For Each pProcess As Process In PerfProcesses

                                ' Write debug
                                Logger.WriteDebug(CallStack, "Kill process: " + pProcess.ProcessName.ToString + "," + pProcess.Id.ToString)

                                ' Kill process
                                pProcess.Kill()

                            Next

                            ' Write debug
                            Logger.WriteDebug(CallStack, "Performance lite agent service: TERMINATED")

                            ' Infinite loop safety -- Exit while condition
                            Exit While

                        End If

                    End While

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Performance lite agent service: STOPPED")

                Else

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Performance lite agent service: STOPPED")

                End If

            Else

                ' Write debug
                Logger.WriteDebug(CallStack, "Performance lite agent service: INACTIVE")

            End If

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Exception caught while stopping the performance lite agent service.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

            ' Create exception
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

            ' Return
            Return 6

        End Try

        ' *****************************
        ' - Stop the system tray processes.
        ' *****************************

        Try

            ' Check for running tray processes
            If Utility.IsProcessRunning("cfsystray") Then

                ' Write debug
                Logger.WriteDebug(CallStack, "System tray: ACTIVE")

                ' Get process listing
                Dim TrayProcesses() As Process = Process.GetProcessesByName("cfsystray")

                ' Process the list
                For Each pProcess As Process In TrayProcesses

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Terminate process: " + pProcess.ProcessName.ToString + "," + pProcess.Id.ToString)

                    ' Kill process
                    pProcess.Kill()

                Next

                ' Write debug
                Logger.WriteDebug(CallStack, "System tray: TERMINATED")

            Else

                ' Write debug
                Logger.WriteDebug(CallStack, "System tray: INACTIVE")

            End If

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Exception caught stopping the system tray process.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

            ' Create exception
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

            ' Return
            Return 7

        End Try

        ' *****************************
        ' - Refresh the notification area.
        ' *****************************

        Try

            ' Check for user-based execution mode
            If Not Globals.RunningAsSystemIdentity Then

                ' Refresh notification tray area
                WindowsAPI.RefreshNotificationArea()

                ' Write debug
                Logger.WriteDebug(CallStack, "Notification area refreshed.")

            End If

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Warning: Exception caught refreshing notification tray.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

        End Try

        ' *****************************
        ' - Stop external CAF processes.
        ' *****************************

        ' Kill external utilities
        Utility.KillProcess("sd_jexec")
        Utility.KillProcess("dsmdiag")
        Utility.KillProcess("cmDirEngJob")
        Utility.KillProcess("ACServer")
        Utility.KillProcess("amSvpCmd")
        Utility.KillProcess("amswsigscan")
        Utility.KillProcess("amsoftscan")
        Utility.KillProcess("cfCafDialog")
        Utility.KillProcess("dmscript")
        Utility.KillProcess("ContentUtility")
        Utility.KillProcess("EGC30N")
        Utility.KillProcess("cfsystray")
        Utility.KillProcessByPath("javaw.exe", "\CIC\jre\")
        Utility.KillProcessByCommandLine("iexplore.exe", "Embedding")
        Utility.KillProcessByCommandLine("w3wp.exe", "DSM_WebService_HM")
        Utility.KillProcessByCommandLine("w3wp.exe", "DSM_WebService")
        Utility.KillProcessByCommandLine("w3wp.exe", "ITCM_Application_Pool")

        ' *****************************
        ' - Stop the CAF service.
        ' *****************************

        Try

            ' Check for CAF process
            If Utility.IsProcessRunning("caf") Then

                ' Write debug
                Logger.WriteDebug(CallStack, "CAF service: ACTIVE")

                ' Build execution string
                ExecutionString = Globals.DSMFolder + "bin\caf.exe"
                ArgumentString = "stop"

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

                    ' Analyze standard output
                    If ConsoleOutput IsNot Nothing AndAlso ConsoleOutput.ToLower.Contains(" 260 ") Then

                        ' Create helper thread for "caf stop" operation
                        CAFStopHelperThread = New Thread(Sub() CafStopWorker())

                        ' Start helper thread
                        CAFStopHelperThread.Start()

                    End If

                    ' Rest thread
                    Thread.Sleep(Globals.THREAD_REST_INTERVAL)

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

                ' Assign caf exit code
                CAFExitCode = RunningProcess.ExitCode

                ' Close detached process
                RunningProcess.Close()

                ' Delay execution
                Thread.Sleep(2000)

                ' Encapsulate non-essenial thread check
                Try

                    ' Check helper thread status
                    If CAFStopHelperThread IsNot Nothing AndAlso CAFStopHelperThread.IsAlive Then

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Waiting for helper thread to terminate.")

                        ' Join helper thread before continuing
                        CAFStopHelperThread.Join()

                    End If

                Catch ex As Exception

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Warning: Exception caught while joining with helper thread.")
                    Logger.WriteDebug(ex.Message)
                    Logger.WriteDebug(ex.StackTrace)

                End Try

                ' Check for non-zero exit code
                If CAFExitCode <> 0 Then

                    ' Check for access denied error
                    If StandardOutput.ToLower.Contains("access is denied") Then

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Administrator unable to stop CAF service.")

                        ' Call the launch pad
                        RunLevel = LaunchPad(CallStack, "System", Globals.DSMFolder + "bin\caf.exe", Globals.DSMFolder + "bin", "stop")

                        ' Reset the loop counter
                        LoopCounter = 0

                        ' Loop while CAF is running
                        While Utility.IsProcessRunning("caf.exe", "stop") Or Utility.IsProcessRunning("caf")

                            ' Check the loop counter
                            If LoopCounter = 0 Then

                                ' Write debug
                                Logger.WriteDebug(CallStack, "CAF service: STOPPING  [This may take up to 5 minutes!]")

                            Else

                                ' Write debug
                                Logger.WriteDebug(CallStack, "CAF service: STOPPING")

                            End If

                            ' Increment loop counter
                            LoopCounter += 1

                            ' Check the loop counter
                            If LoopCounter >= 10 Then

                                ' Write debug
                                Logger.WriteDebug(CallStack, "CAF service is not stopping gracefully.")
                                Logger.WriteDebug(CallStack, "Send kill request to the CAF service..")

                                ' Build execution string
                                ExecutionString = Globals.DSMFolder + "bin\caf.exe"
                                ArgumentString = "kill all"

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
                                Logger.WriteDebug(CallStack, "Exit code: " + RunningProcess.ExitCode.ToString)

                                ' Close detached process
                                RunningProcess.Close()

                                ' Exit loop after kill request
                                Exit While

                            End If

                            ' Rest the thread
                            Thread.Sleep(250)

                        End While

                        ' Check for CAF process
                        If Utility.IsProcessRunning("caf") Then

                            ' Write debug
                            Logger.WriteDebug(CallStack, "Error: Unable to terminate CAF service.")

                            ' Create exception
                            Manifest.UpdateManifest(CallStack,
                                                    Manifest.EXCEPTION_MANIFEST,
                                                    {"Error: Unable to terminate CAF service.",
                                                    "Reason: Access to the CAF service was denied."})

                            ' Return
                            Return 8

                        End If

                    Else

                        ' Write debug
                        Logger.WriteDebug(CallStack, "CAF service: WARNING [Terminated with non-zero exit code]")

                    End If

                Else

                    ' Verify graceful stop was successful
                    If Utility.IsProcessRunning("caf") Then

                        ' Write debug
                        Logger.WriteDebug(CallStack, "CAF service has not stopped gracefully.")
                        Logger.WriteDebug(CallStack, "Send first kill request to the CAF service..")

                        ' Build execution string
                        ExecutionString = Globals.DSMFolder + "bin\caf.exe"
                        ArgumentString = "kill all"

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
                        Logger.WriteDebug(CallStack, "Exit code: " + RunningProcess.ExitCode.ToString)

                        ' Close detached process
                        RunningProcess.Close()

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Send second kill request to the CAF service..")

                        ' Build execution string
                        ExecutionString = Globals.DSMFolder + "bin\caf.exe"
                        ArgumentString = "kill all"

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
                        Logger.WriteDebug(CallStack, "Exit code: " + RunningProcess.ExitCode.ToString)

                        ' Close detached process
                        RunningProcess.Close()

                        ' Delay execution
                        Thread.Sleep(2000)

                    End If

                    ' Check for CAF process
                    If Utility.IsProcessRunning("caf") Then

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Error: Unable to terminate CAF service.")

                        ' Create exception
                        Manifest.UpdateManifest(CallStack,
                                                Manifest.EXCEPTION_MANIFEST,
                                                {"Error: Unable to terminate CAF service.",
                                                "Reason: Access to the CAF service was denied."})

                        ' Return
                        Return 8

                    End If

                End If

                ' Write debug
                Logger.WriteDebug(CallStack, "CAF service: STOPPED")

            Else

                ' Write debug
                Logger.WriteDebug(CallStack, "CAF service: INACTIVE")

            End If

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Exception caught while stopping the CAF service.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

            ' Create exception
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

            ' Return
            Return 9

        End Try

        ' *****************************
        ' - Stop external CAF processes.
        ' *****************************

        ' Kill external utilities
        Utility.KillProcess("tngdta")

        ' *****************************
        ' - Stop the CAM service.
        ' *****************************

        Try

            ' Check for CAM process
            If Utility.IsProcessRunning("cam") AndAlso Not Globals.SkipCAM Then

                ' Write debug
                Logger.WriteDebug(CallStack, "CAM service: ACTIVE")

                ' *****************************
                ' - Disable the CAM service.
                ' *****************************

                ' Build execution string
                CAMDirectory = Environment.GetEnvironmentVariable("CAI_MSQ")
                ExecutionString = CAMDirectory + "\bin\cam.exe"
                ArgumentString = "change disabled"

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
                Logger.WriteDebug(CallStack, "CAM service has been disabled.")

                ' *****************************
                ' - Close the CAM service.
                ' *****************************

                ' Build execution string
                ExecutionString = CAMDirectory + "\bin\camclose.exe"
                ArgumentString = ""

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

                ' Delay execution
                Thread.Sleep(2000)

                ' Check for active CAM service
                If Utility.IsProcessRunning("cam") Then

                    ' Reset the loop counter
                    LoopCounter = 0

                    ' Wait for CAM to stop
                    While Utility.IsProcessRunning("cam")

                        ' Write debug
                        Logger.WriteDebug(CallStack, "CAM service: STOPPING")

                        ' Pause execution
                        Thread.Sleep(5000)

                        ' Increment the loop counter
                        LoopCounter += 1

                        ' Check the loop counter
                        If LoopCounter >= 5 Then

                            ' Get list of all tray processes
                            Dim TrayProcess() As Process = Process.GetProcessesByName("cam")

                            ' Process the list
                            For Each pProcess As Process In TrayProcess

                                ' Write debug
                                Logger.WriteDebug(CallStack, "Kill process: " + pProcess.ProcessName.ToString + "," + pProcess.Id.ToString)

                                ' Kill process
                                pProcess.Kill()

                            Next

                            ' Write debug
                            Logger.WriteDebug(CallStack, "CAM service: TERMINATED")

                            ' Infinite loop safety -- Exit while condition
                            Exit While

                        End If

                    End While

                    ' Write debug
                    Logger.WriteDebug(CallStack, "CAM service: STOPPED")

                Else

                    ' Write debug
                    Logger.WriteDebug(CallStack, "CAM service: STOPPED")

                End If

            ElseIf Globals.SkipCAM Then

                ' Write debug
                Logger.WriteDebug(CallStack, "CAM service: SKIPPED")

            Else

                ' Write debug
                Logger.WriteDebug(CallStack, "CAM service: INACTIVE")

            End If

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Exception caught while closing the CAM service.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

            ' Create exception
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

            ' Don't return because caf is stopped.
            ' Return 10

        End Try

        ' *****************************
        ' - Stop CSAM service.
        ' *****************************

        Try

            ' Check for CSAM process
            If Utility.IsProcessRunning("csampmux") AndAlso Not Globals.SkipCAM Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Port multiplexer service: ACTIVE")

                ' Build execution string
                ExecutionString = Globals.SSAFolder + "bin\csampmux.exe"
                ArgumentString = "stop"

                ' Write debug
                Logger.WriteDebug(CallStack, "Detached process: " + ExecutionString + " " + ArgumentString)

                ' Create detached process
                ProcessStartInfo = New ProcessStartInfo(ExecutionString, ArgumentString)
                ProcessStartInfo.WorkingDirectory = Globals.SSAFolder + "bin"
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
                    Thread.Sleep(Globals.THREAD_REST_INTERVAL)

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

                ' Delay execution
                Thread.Sleep(2000)

                ' Check for active CSAM service
                If Utility.IsProcessRunning("csampmux") Then

                    ' Reset the loop counter
                    LoopCounter = 0

                    ' Wait for CSAM to stop
                    While Utility.IsProcessRunning("csampmux")

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Port multiplexer service: STOPPING")

                        ' Delay execution
                        Thread.Sleep(5000)

                        ' Increment the loop counter
                        LoopCounter += 1

                        ' Check the loop counter
                        If LoopCounter >= 5 Then

                            ' Get process listing
                            Dim TrayProcess() As Process = Process.GetProcessesByName("csampmux")

                            ' Process the list
                            For Each pProcess As Process In TrayProcess

                                ' Write debug
                                Logger.WriteDebug(CallStack, "Kill process: " + pProcess.ProcessName.ToString + "," + pProcess.Id.ToString)

                                ' Kill process
                                pProcess.Kill()

                            Next

                            ' Write debug
                            Logger.WriteDebug(CallStack, "Port multiplexer service: TERMINATED")

                            ' Infinite loop safety -- Exit while condition
                            Exit While

                        End If

                    End While

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Port multiplexer service: STOPPED")

                Else

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Port multiplexer service: STOPPED")

                End If

            Else

                ' Write debug
                Logger.WriteDebug(CallStack, "Port multiplexer service: INACTIVE")

            End If

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Exception caught while stopping the port multiplexer service.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

            ' Create exception
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

            ' Don't return because caf is stopped.
            ' Return 11

        End Try

        ' Return
        Return RunLevel

    End Function

    Public Shared Sub CafStopWorker()

        ' Local variables
        Dim CafProcessList As New ArrayList
        Dim LoopCounter As Integer = 0
        Dim ChildProcess As ArrayList
        Dim WorkingMemoryBase As Double
        Dim WorkingMemoryCurrent As Double

        ' CafProcessList is an ArrayList with each element containing an Arraylist with the following contents:
        ' (0) = ProcessID, Ex: "1234"
        ' (1) = Name, Ex: "cmEngine.exe"
        ' (2) = ExecutablePath, Ex: "C:\Program Files (x86)\CA\DSM\Bin\cmEngine.exe"
        ' (3) = CommandLine, Ex: "C:\Program Files (x86)\CA\DSM\Bin\cmEngine.exe -name SystemEngine"
        ' (4) = WorkingSetSize, in bytes, Ex: 1234567

        ' Encapsulate thread operations
        Try

            ' Write debug
            Logger.WriteDebug("Helper thread: Expediting ""caf stop"" operation..")

            ' Loop while "caf stop" operation is running
            While Utility.IsProcessRunning("caf.exe", "stop")

                ' Obtain process arraylist (filtered for caf subsystem processes)
                CafProcessList = Utility.GetProcessChildren("caf.exe", "service", New ArrayList({"cfsmsmd.exe", "ccnfagent.exe", "encclient.exe"}))

                ' Check for an empty list
                If CafProcessList.Count = 0 Then

                    ' Write debug
                    Logger.WriteDebug("Helper thread: Finished.")

                    ' Return
                    Return

                End If

                ' Write debug
                Logger.WriteDebug("Helper thread: Monitoring (" + CafProcessList.Count.ToString + ") processes.")

                ' Reset loop counter
                LoopCounter = 0

                ' Rest for a short, but reasonable interval
                While LoopCounter < 50

                    ' Check for termination signal
                    If Not Utility.IsProcessRunning("caf.exe", "stop") Then

                        ' Write debug
                        Logger.WriteDebug("Helper thread: Finished.")

                        ' Return
                        Return

                    End If

                    ' Increment loop counter
                    LoopCounter += 1

                    ' Rest the thread
                    Thread.Sleep(Globals.THREAD_REST_INTERVAL)

                End While

                ' Iterate the process list
                For x As Integer = CafProcessList.Count - 1 To 0 Step -1

                    ' Obtain a child process arraylist
                    ChildProcess = CafProcessList(x)

                    ' Write debug
                    Logger.WriteDebug("Helper thread: Analyzing PID " + ChildProcess.Item(0).ToString + " -- " + ChildProcess.Item(1).ToString)

                    ' Check if the process is still running
                    If Not Utility.IsProcessRunning(Integer.Parse(ChildProcess.Item(0))) Then

                        ' Write debug
                        Logger.WriteDebug("Helper thread: Self-terminated PID " + ChildProcess.Item(0).ToString + " -- " + ChildProcess.Item(1).ToString)

                        ' Remove from the list
                        CafProcessList.RemoveAt(x)

                        ' Continue loop
                        Continue For

                    End If

                    ' Obtain base working set memory size
                    WorkingMemoryBase = Double.Parse(ChildProcess.Item(4))

                    ' Obtain current snapshot of working set memory size
                    WorkingMemoryCurrent = Utility.GetProcessWorkingSetMemorySize(ChildProcess.Item(0))

                    ' Check if working set memory size has increased or remained the same
                    If WorkingMemoryCurrent >= WorkingMemoryBase Then

                        ' Kill the process
                        Utility.KillProcess(Integer.Parse(ChildProcess.Item(0)))

                        ' Remove from the list
                        CafProcessList.RemoveAt(x)

                        ' Write debug
                        Logger.WriteDebug("Helper thread: Killed PID " + ChildProcess.Item(0).ToString + " -- " + ChildProcess.Item(1).ToString)

                        ' Continue loop
                        Continue For

                    End If

                    ' Write debug
                    Logger.WriteDebug("Helper thread: Memory decrease PID " + ChildProcess.Item(0).ToString + " -- " + ChildProcess.Item(1).ToString)

                Next

            End While

            ' Write debug
            Logger.WriteDebug("Helper thread: Finished.")

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug("Helper thread: Exception caught monitoring caf.exe child processes.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

        End Try

    End Sub

    Public Shared Function StopCAFOnDemand(ByVal CallStack As String) As Integer

        ' Local variables
        Dim ExecutionString As String               ' Command line to be executed externally to the application.
        Dim ArgumentString As String                ' Arguments passed on the command line for the external execution.
        Dim RunningProcess As Process               ' A process shell for executing the command line.
        Dim ConsoleOutput As String                 ' Live output from an external application execution.
        Dim StandardOutput As String                ' Output from an external application execution.
        Dim RemainingOutput As String               ' Additional output flushed after a process has exited.
        Dim ProcessStartInfo As ProcessStartInfo    ' Process startup settings for configuring the bahviour of the process.
        Dim CAFExitCode As Integer                  ' CAF command exit code.
        Dim CAFStopHelperThread As Thread = Nothing ' Helper thread for expediting "caf stop" operation. 
        Dim LoopCounter As Integer = 0              ' Re-usable loop counter.
        Dim RestCounter As Integer = 0              ' Used for resting the thread in short intervals.

        ' Update call stack
        CallStack += "StopCAF|"

        ' *****************************
        ' - Check CAF service status.
        ' *****************************

        ' Check if CAF service exists
        If Not Utility.ServiceExists("caf") Then

            ' Write debug
            Logger.WriteDebug(CallStack, "CAF service: NOT INSTALLED")

            ' Return
            Return 0

        End If

        ' Check if CAF service is disabled
        If Not Utility.IsServiceEnabled("caf") Then

            ' Write debug
            Logger.WriteDebug(CallStack, "CAF service: DISABLED")

            ' Return
            Return 0

        End If

        ' Check if CAF service is already stopped
        If Not Utility.IsProcessRunning("caf") Then

            ' Write debug
            Logger.WriteDebug(CallStack, "CAF service: ALREADY STOPPED")

            ' Return
            Return 0

        End If

        ' *****************************
        ' - Stop DSM Explorer processes.
        ' *****************************

        Try

            ' Check for running explorer processes
            If Utility.IsProcessRunning("egc30n") Then

                ' Write debug
                Logger.WriteDebug(CallStack, "DSM Explorer: ACTIVE")

                ' Get process listing
                Dim ExplorerProcesses() As Process = System.Diagnostics.Process.GetProcessesByName("EGC30N")

                ' Process the list
                For Each pProcess As Process In ExplorerProcesses

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Terminate process: " + pProcess.ProcessName.ToString + "," + pProcess.Id.ToString)

                    ' Kill processes
                    pProcess.Kill()

                Next

                ' Write debug
                Logger.WriteDebug(CallStack, "DSM Explorer: TERMINATED")

            Else

                ' Write debug
                Logger.WriteDebug(CallStack, "DSM Explorer: INACTIVE")

            End If

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Exception caught terminating DSM Explorer processes.")
            Logger.WriteDebug(ex.Message)

        End Try

        ' *****************************
        ' - Stop external CAF processes.
        ' *****************************

        ' Kill external caf processes
        Utility.KillProcess("sd_jexec")
        Utility.KillProcess("dsmdiag")
        Utility.KillProcess("cmDirEngJob")
        Utility.KillProcess("ACServer")
        Utility.KillProcess("amSvpCmd")
        Utility.KillProcess("amswsigscan")
        Utility.KillProcess("amsoftscan")
        Utility.KillProcess("cfCafDialog")
        Utility.KillProcess("dmscript")
        Utility.KillProcess("ContentUtility")
        Utility.KillProcess("EGC30N")
        Utility.KillProcessByPath("javaw.exe", "\CIC\jre\")
        Utility.KillProcessByCommandLine("iexplore.exe", "Embedding")
        Utility.KillProcessByCommandLine("w3wp.exe", "DSM_WebService_HM")
        Utility.KillProcessByCommandLine("w3wp.exe", "DSM_WebService")
        Utility.KillProcessByCommandLine("w3wp.exe", "ITCM_Application_Pool")

        ' *****************************
        ' - Stop the CAF service.
        ' *****************************

        Try

            ' Check for CAF process
            If Utility.IsProcessRunning("caf") Then

                ' Write debug
                Logger.WriteDebug(CallStack, "CAF service: ACTIVE")

                ' Build execution string
                ExecutionString = Globals.DSMFolder + "bin\caf.exe"
                ArgumentString = "stop"

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

                    ' Analyze standard output
                    If ConsoleOutput IsNot Nothing AndAlso ConsoleOutput.ToLower.Contains(" 260 ") Then

                        ' Create helper thread for "caf stop" operation
                        CAFStopHelperThread = New Thread(Sub() CafStopWorker())

                        ' Start helper thread
                        CAFStopHelperThread.Start()

                    End If

                    ' Rest thread
                    Thread.Sleep(Globals.THREAD_REST_INTERVAL)

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

                ' Assign caf exit code
                CAFExitCode = RunningProcess.ExitCode

                ' Close detached process
                RunningProcess.Close()

                ' Delay execution
                Thread.Sleep(2000)

                ' Encapsulate non-essenial thread check
                Try

                    ' Check helper thread status
                    If CAFStopHelperThread IsNot Nothing AndAlso CAFStopHelperThread.IsAlive Then

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Waiting for helper thread to terminate.")

                        ' Join helper thread before continuing
                        CAFStopHelperThread.Join()

                    End If

                Catch ex As Exception

                    Logger.WriteDebug(CallStack, "Warning: Exception caught while joining with helper thread.")
                    Logger.WriteDebug(ex.Message)
                    Logger.WriteDebug(ex.StackTrace)

                End Try

                ' Check for non-zero exit code
                If CAFExitCode <> 0 Then

                    ' Check for access denied error
                    If StandardOutput.ToLower.Contains("access is denied") Then

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Administrator unable to stop CAF service.")

                        ' Call the launch pad
                        LaunchPad(CallStack, "System", Globals.DSMFolder + "bin\caf.exe", Globals.DSMFolder + "bin\", "stop")

                        ' Reset the loop counter
                        LoopCounter = 0

                        ' Loop while CAF is running
                        While Utility.IsProcessRunning("caf.exe", "stop") Or Utility.IsProcessRunning("caf")

                            ' Check the loop counter
                            If LoopCounter = 0 Then

                                ' Write debug
                                Logger.WriteDebug(CallStack, "CAF service: STOPPING  [This may take up to 5 minutes!]")

                            Else

                                ' Write debug
                                Logger.WriteDebug(CallStack, "CAF service: STOPPING")

                            End If

                            ' Increment loop counter
                            LoopCounter += 1

                            ' Check the loop counter
                            If LoopCounter >= 10 Then

                                ' Write debug
                                Logger.WriteDebug(CallStack, "CAF service is not stopping gracefully.")
                                Logger.WriteDebug(CallStack, "Send kill request to the CAF service..")

                                ' Build execution string
                                ExecutionString = Globals.DSMFolder + "bin\caf.exe"
                                ArgumentString = "kill all"

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
                                Logger.WriteDebug(CallStack, "Exit code: " + RunningProcess.ExitCode.ToString)

                                ' Close detached process
                                RunningProcess.Close()

                                ' Exit loop after kil request
                                Exit While

                            End If

                            ' Reset rest counter
                            RestCounter = 0

                            ' Rest for a short, but reasonable interval
                            While RestCounter < 199

                                ' Increment rest counter
                                RestCounter += 1

                                ' Rest the thread
                                Thread.Sleep(Globals.THREAD_REST_INTERVAL)

                            End While

                        End While

                        ' Check for CAF process
                        If Utility.IsProcessRunning("caf") Then

                            ' Write debug
                            Logger.WriteDebug(CallStack, "Error: Unable to terminate CAF service.")

                            ' Return
                            Return 1

                        End If

                    Else

                        ' Write debug
                        Logger.WriteDebug(CallStack, "CAF service: WARNING [Terminated with non-zero exit code]")

                    End If

                Else

                    ' Verify graceful stop was successful
                    If Utility.IsProcessRunning("caf") Then

                        ' Write debug
                        Logger.WriteDebug(CallStack, "CAF service has not stopped gracefully.")
                        Logger.WriteDebug(CallStack, "Send first kill request to the CAF service..")

                        ' Build execution string
                        ExecutionString = Globals.DSMFolder + "bin\caf.exe"
                        ArgumentString = "kill all"

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
                        Logger.WriteDebug(CallStack, "Exit code: " + RunningProcess.ExitCode.ToString)

                        ' Close detached process
                        RunningProcess.Close()

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Send second kill request to the CAF service..")

                        ' Build execution string
                        ExecutionString = Globals.DSMFolder + "bin\caf.exe"
                        ArgumentString = "kill all"

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
                        Logger.WriteDebug(CallStack, "Exit code: " + RunningProcess.ExitCode.ToString)

                        ' Close detached process
                        RunningProcess.Close()

                    End If

                    ' Check for CAF process
                    If Utility.IsProcessRunning("caf") Then

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Error: Unable to terminate CAF service.")

                        ' Return
                        Return 2

                    End If

                End If

                ' Write debug
                Logger.WriteDebug(CallStack, "CAF service: STOPPED")

            Else

                ' Write debug
                Logger.WriteDebug(CallStack, "CAF service: INACTIVE")

            End If

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Exception caught while stopping the CAF service.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

            ' Return
            Return 3

        End Try

        ' *****************************
        ' - Stop external CAF processes.
        ' *****************************

        ' Kill external utilities
        Utility.KillProcess("tngdta")

        ' *****************************
        ' - Clear notification server cache data.
        ' *****************************

        ' Delete cache file
        Utility.DeleteFile(CallStack, Globals.DSMFolder + "appdata\cfnotsrvd.dat")

        ' Return
        Return 0

    End Function

End Class