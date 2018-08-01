'****************************** Class Header *******************************\
' Project Name: WinOffline
' Class Name:   WinOfflineUI
' File Name:    WinOfflineUI-CafStopStart.vb
' Author:       Brian Fontana
'***************************************************************************/

Imports System.Threading
Imports System.Windows.Forms

Partial Public Class WinOfflineUI

    Private Sub StartCAF(ByVal UIOutputControl As Control)

        ' Local variables
        Dim ExecutionString As String
        Dim ArgumentString As String
        Dim RunningProcess As Process
        Dim ConsoleOutput As String
        Dim StandardOutput As String = ""
        Dim RemainingOutput As String
        Dim ProcessStartInfo As ProcessStartInfo
        Dim LoopCounter As Integer = 0

        ' Encapsulate operation
        Try

            ' Check if CAF service exists
            If Not WinOffline.Utility.ServiceExists("caf") Then

                ' Write debug
                Delegate_Sub_Append_Text(UIOutputControl, "CAF service: NOT INSTALLED")

                ' Return
                Return

            End If

            ' Check if CAF service is disabled
            If Not WinOffline.Utility.IsServiceEnabled("caf") Then

                ' Write debug
                Delegate_Sub_Append_Text(UIOutputControl, "CAF service: DISABLED")

                ' Return
                Return

            End If

            ' Check if CAF service is already running
            If WinOffline.Utility.IsProcessRunning("caf") Then

                ' Write debug
                Delegate_Sub_Append_Text(UIOutputControl, "CAF service: ALREADY RUNNING")

                ' Return
                Return

            End If

            ' Build execution string
            ExecutionString = Globals.DSMFolder + "bin\caf.exe"
            ArgumentString = "start"

            ' Write debug
            Delegate_Sub_Append_Text(UIOutputControl, "Detached process: " + ExecutionString + " " + ArgumentString + Environment.NewLine)

            ' Create detached process
            ProcessStartInfo = New ProcessStartInfo(ExecutionString, ArgumentString)
            ProcessStartInfo.WorkingDirectory = Globals.DSMFolder + "bin"
            ProcessStartInfo.UseShellExecute = False
            ProcessStartInfo.RedirectStandardOutput = True
            ProcessStartInfo.CreateNoWindow = True

            ' Reset standard output
            StandardOutput = ""
            RemainingOutput = ""

            ' Start detached process
            RunningProcess = Process.Start(ProcessStartInfo)

            ' Read live output
            While RunningProcess.HasExited = False

                ' Read output
                ConsoleOutput = RunningProcess.StandardOutput.ReadLine

                ' Write debug
                Delegate_Sub_Append_Text(UIOutputControl, ConsoleOutput)

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
            Delegate_Sub_Append_Text(UIOutputControl, RemainingOutput + Environment.NewLine)
            Delegate_Sub_Append_Text(UIOutputControl, "Exit code: " + RunningProcess.ExitCode.ToString + Environment.NewLine)

            ' Close detached process
            RunningProcess.Close()

        Catch ex As Exception

            ' Write debug
            Delegate_Sub_Append_Text(UIOutputControl, "Error: Exception caught while starting CAF service." + Environment.NewLine)
            Delegate_Sub_Append_Text(UIOutputControl, ex.Message)
            Delegate_Sub_Append_Text(UIOutputControl, ex.StackTrace)

            ' Return
            Return

        End Try

        ' Write debug
        Delegate_Sub_Append_Text(UIOutputControl, "CAF service: STARTUP COMPLETE" + Environment.NewLine)

    End Sub

    Private Sub StopCAF(ByVal UIOutputControl As Control)

        ' Local variables
        Dim ExecutionString As String
        Dim ArgumentString As String
        Dim RunningProcess As Process
        Dim ConsoleOutput As String
        Dim StandardOutput As String
        Dim RemainingOutput As String
        Dim ProcessStartInfo As ProcessStartInfo
        Dim CAFExitCode As Integer
        Dim CAFStopHelperThread As Thread = Nothing
        Dim LoopCounter As Integer = 0
        Dim RestCounter As Integer = 0

        ' Encapsulate operation
        Try

            ' Check if CAF service exists
            If Not WinOffline.Utility.ServiceExists("caf") Then

                ' Write debug
                Delegate_Sub_Append_Text(UIOutputControl, "CAF service: NOT INSTALLED")

                ' Return
                Return

            End If

            ' Check if CAF service is disabled
            If Not WinOffline.Utility.IsServiceEnabled("caf") Then

                ' Write debug
                Delegate_Sub_Append_Text(UIOutputControl, "CAF service: DISABLED")

                ' Return
                Return

            End If

            ' Check if CAF service is already stopped
            If Not WinOffline.Utility.IsProcessRunning("caf") Then

                ' Write debug
                Delegate_Sub_Append_Text(UIOutputControl, "CAF service: ALREADY STOPPED")

                ' Return
                Return

            End If

            ' Kill external caf processes
            WinOffline.Utility.KillProcess("sd_jexec")
            WinOffline.Utility.KillProcess("dsmdiag")
            WinOffline.Utility.KillProcess("cmDirEngJob")
            WinOffline.Utility.KillProcess("ACServer")
            WinOffline.Utility.KillProcess("amSvpCmd")
            WinOffline.Utility.KillProcess("amswsigscan")
            WinOffline.Utility.KillProcess("amsoftscan")
            WinOffline.Utility.KillProcess("cfCafDialog")
            WinOffline.Utility.KillProcess("dmscript")
            WinOffline.Utility.KillProcess("ContentUtility")
            WinOffline.Utility.KillProcess("EGC30N")
            WinOffline.Utility.KillProcessByPath("javaw.exe", "\CIC\jre\")
            WinOffline.Utility.KillProcessByCommandLine("iexplore.exe", "Embedding")
            WinOffline.Utility.KillProcessByCommandLine("w3wp.exe", "DSM_WebService_HM")
            WinOffline.Utility.KillProcessByCommandLine("w3wp.exe", "DSM_WebService")
            WinOffline.Utility.KillProcessByCommandLine("w3wp.exe", "ITCM_Application_Pool")

            ' Check for CAF process
            If WinOffline.Utility.IsProcessRunning("caf") Then

                ' Build execution string
                ExecutionString = Globals.DSMFolder + "bin\caf.exe"
                ArgumentString = "stop"

                ' Write debug
                Delegate_Sub_Append_Text(UIOutputControl, "Detached process: " + ExecutionString + " " + ArgumentString + Environment.NewLine)

                ' Create detached process
                ProcessStartInfo = New ProcessStartInfo(ExecutionString, ArgumentString)
                ProcessStartInfo.WorkingDirectory = Globals.DSMFolder + "bin"
                ProcessStartInfo.UseShellExecute = False
                ProcessStartInfo.RedirectStandardOutput = True
                ProcessStartInfo.CreateNoWindow = True

                ' Reset standard output
                StandardOutput = ""
                RemainingOutput = ""

                ' Start detached process
                RunningProcess = Process.Start(ProcessStartInfo)

                ' Read live output
                While RunningProcess.HasExited = False

                    ' Read output
                    ConsoleOutput = RunningProcess.StandardOutput.ReadLine

                    ' Write debug
                    Delegate_Sub_Append_Text(UIOutputControl, ConsoleOutput)

                    ' Update standard output
                    StandardOutput += ConsoleOutput + Environment.NewLine

                    ' Analyze standard output
                    If ConsoleOutput IsNot Nothing AndAlso ConsoleOutput.ToLower.Contains(" 260 ") Then

                        ' Create helper thread for "caf stop" operation
                        CAFStopHelperThread = New Thread(Sub() CafStopWorker(UIOutputControl))

                        ' Start helper thread
                        CAFStopHelperThread.Start()

                    End If

                    ' Rest thread
                    Threading.Thread.Sleep(Globals.THREAD_REST_INTERVAL)

                End While

                ' Wait for detached process to exit
                RunningProcess.WaitForExit()

                ' Append remaining standard output
                RemainingOutput = RunningProcess.StandardOutput.ReadToEnd.ToString
                StandardOutput += RemainingOutput

                ' Write debug
                Delegate_Sub_Append_Text(UIOutputControl, RemainingOutput + Environment.NewLine)
                Delegate_Sub_Append_Text(UIOutputControl, "Exit code: " + RunningProcess.ExitCode.ToString + Environment.NewLine)

                ' Assign caf exit code
                CAFExitCode = RunningProcess.ExitCode

                ' Close detached process
                RunningProcess.Close()

                ' Encapsulate non-essenial thread check
                Try

                    ' Check helper thread status
                    If CAFStopHelperThread IsNot Nothing AndAlso CAFStopHelperThread.IsAlive Then

                        ' Write debug
                        Delegate_Sub_Append_Text(UIOutputControl, "Waiting for helper thread to terminate.")

                        ' Join helper thread before continuing
                        CAFStopHelperThread.Join()

                    End If

                Catch ex As Exception

                    ' Write debug
                    Delegate_Sub_Append_Text(UIOutputControl, "Warning: Exception caught while joining with helper thread.")
                    Delegate_Sub_Append_Text(UIOutputControl, ex.Message)
                    Delegate_Sub_Append_Text(UIOutputControl, ex.StackTrace)

                End Try

                ' Check for non-zero exit code
                If CAFExitCode <> 0 Then

                    ' Check for access denied error
                    If StandardOutput.ToLower.Contains("access is denied") Then

                        ' Write debug
                        Delegate_Sub_Append_Text(UIOutputControl, "Administrator unable to stop CAF service.")

                        ' Call the launch pad
                        LaunchPadUI(UIOutputControl, "System", Globals.DSMFolder + "bin\caf.exe", "stop")

                        ' Reset the loop counter
                        LoopCounter = 0

                        ' Loop while CAF is running
                        While WinOffline.Utility.IsProcessRunning("caf.exe", "stop") Or WinOffline.Utility.IsProcessRunning("caf")

                            ' Check loop counter
                            If LoopCounter = 0 Then

                                ' Write debug
                                Delegate_Sub_Append_Text(UIOutputControl, "CAF service: STOPPING  [This may take up to 5 minutes!]")

                            Else

                                ' Write debug
                                Delegate_Sub_Append_Text(UIOutputControl, "CAF service: STOPPING")

                            End If

                            ' Increment loop counter
                            LoopCounter += 1

                            ' Check the loop counter
                            If LoopCounter >= 10 Then

                                ' Write debug
                                Delegate_Sub_Append_Text(UIOutputControl, "CAF service is not stopping gracefully.")
                                Delegate_Sub_Append_Text(UIOutputControl, "Send first kill request to the CAF service..")

                                ' Build execution string
                                ExecutionString = Globals.DSMFolder + "bin\caf.exe"
                                ArgumentString = "kill all"

                                ' Write debug
                                Delegate_Sub_Append_Text(UIOutputControl, "Detached process: " + ExecutionString + " " + ArgumentString + Environment.NewLine)

                                ' Create detached process
                                ProcessStartInfo = New ProcessStartInfo(ExecutionString, ArgumentString)
                                ProcessStartInfo.WorkingDirectory = Globals.DSMFolder + "bin"
                                ProcessStartInfo.UseShellExecute = False
                                ProcessStartInfo.RedirectStandardOutput = True
                                ProcessStartInfo.CreateNoWindow = True

                                ' Reset standard output
                                StandardOutput = ""
                                RemainingOutput = ""

                                ' Start detached process
                                RunningProcess = Process.Start(ProcessStartInfo)

                                ' Read live output
                                While RunningProcess.HasExited = False

                                    ' Read output
                                    ConsoleOutput = RunningProcess.StandardOutput.ReadLine

                                    ' Write debug
                                    Delegate_Sub_Append_Text(UIOutputControl, ConsoleOutput)

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
                                Delegate_Sub_Append_Text(UIOutputControl, RemainingOutput + Environment.NewLine)
                                Delegate_Sub_Append_Text(UIOutputControl, "Exit code: " + RunningProcess.ExitCode.ToString + Environment.NewLine)

                                ' Close detached process
                                RunningProcess.Close()

                                ' Write debug
                                Delegate_Sub_Append_Text(UIOutputControl, "Send second kill request to the CAF service..")

                                ' Build execution string
                                ExecutionString = Globals.DSMFolder + "bin\caf.exe"
                                ArgumentString = "kill all"

                                ' Write debug
                                Delegate_Sub_Append_Text(UIOutputControl, "Detached process: " + ExecutionString + " " + ArgumentString + Environment.NewLine)

                                ' Create detached process
                                ProcessStartInfo = New ProcessStartInfo(ExecutionString, ArgumentString)
                                ProcessStartInfo.WorkingDirectory = Globals.DSMFolder + "bin"
                                ProcessStartInfo.UseShellExecute = False
                                ProcessStartInfo.RedirectStandardOutput = True
                                ProcessStartInfo.CreateNoWindow = True

                                ' Reset standard output
                                StandardOutput = ""
                                RemainingOutput = ""

                                ' Start detached process
                                RunningProcess = Process.Start(ProcessStartInfo)

                                ' Read live output
                                While RunningProcess.HasExited = False

                                    ' Read output
                                    ConsoleOutput = RunningProcess.StandardOutput.ReadLine

                                    ' Write debug
                                    Delegate_Sub_Append_Text(UIOutputControl, ConsoleOutput)

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
                                Delegate_Sub_Append_Text(UIOutputControl, RemainingOutput + Environment.NewLine)
                                Delegate_Sub_Append_Text(UIOutputControl, "Exit code: " + RunningProcess.ExitCode.ToString + Environment.NewLine)

                                ' Close detached process
                                RunningProcess.Close()

                                ' After kill request, exit the loop
                                Exit While

                            End If

                            ' Reset rest counter
                            RestCounter = 0

                            ' Rest for a short, but reasonable interval
                            While RestCounter < 199

                                ' Check for termination signal
                                If TerminateSignal Then Return

                                ' Increment rest counter
                                RestCounter += 1

                                ' Rest the thread
                                Thread.Sleep(Globals.THREAD_REST_INTERVAL)

                            End While

                        End While

                        ' Check for CAF process
                        If WinOffline.Utility.IsProcessRunning("caf") Then

                            ' Write debug
                            Delegate_Sub_Append_Text(UIOutputControl, "Error: Unable to terminate CAF service." + Environment.NewLine)

                            ' Return
                            Return

                        End If

                    Else

                        ' Write debug
                        Delegate_Sub_Append_Text(UIOutputControl, "CAF service: WARNING [Terminated with non-zero exit code]" + Environment.NewLine)

                    End If

                Else

                    ' Delay execution
                    Thread.Sleep(5000)

                    ' Verify graceful stop was successful
                    If WinOffline.Utility.IsProcessRunning("caf") Then

                        ' Write debug
                        Delegate_Sub_Append_Text(UIOutputControl, "CAF service has not stopped gracefully.")
                        Delegate_Sub_Append_Text(UIOutputControl, "Send kill request to the CAF service..")

                        ' Build execution string
                        ExecutionString = Globals.DSMFolder + "bin\caf.exe"
                        ArgumentString = "kill all"

                        ' Write debug
                        Delegate_Sub_Append_Text(UIOutputControl, "Detached process: " + ExecutionString + " " + ArgumentString + Environment.NewLine)

                        ' Create detached process
                        ProcessStartInfo = New ProcessStartInfo(ExecutionString, ArgumentString)
                        ProcessStartInfo.WorkingDirectory = Globals.DSMFolder + "bin"
                        ProcessStartInfo.UseShellExecute = False
                        ProcessStartInfo.RedirectStandardOutput = True
                        ProcessStartInfo.CreateNoWindow = True

                        ' Reset standard output
                        StandardOutput = ""
                        RemainingOutput = ""

                        ' Start detached process
                        RunningProcess = Process.Start(ProcessStartInfo)

                        ' Read live output
                        While RunningProcess.HasExited = False

                            ' Read output
                            ConsoleOutput = RunningProcess.StandardOutput.ReadLine

                            ' Write debug
                            Delegate_Sub_Append_Text(UIOutputControl, ConsoleOutput)

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
                        Delegate_Sub_Append_Text(UIOutputControl, RemainingOutput + Environment.NewLine)
                        Delegate_Sub_Append_Text(UIOutputControl, "Exit code: " + RunningProcess.ExitCode.ToString + Environment.NewLine)

                        ' Close detached process
                        RunningProcess.Close()

                    End If

                    ' Check for CAF process
                    If WinOffline.Utility.IsProcessRunning("caf") Then

                        ' Write debug
                        Delegate_Sub_Append_Text(UIOutputControl, "Error: Unable to terminate CAF service." + Environment.NewLine)

                        ' Return
                        Return

                    End If

                End If

            End If

        Catch ex As Exception

            ' Write debug
            Delegate_Sub_Append_Text(UIOutputControl, "Error: Exception caught while stopping the CAF service." + Environment.NewLine)
            Delegate_Sub_Append_Text(UIOutputControl, ex.Message)
            Delegate_Sub_Append_Text(UIOutputControl, ex.StackTrace)

            ' Return
            Return

        End Try

        ' *****************************
        ' - Stop external CAF processes.
        ' *****************************

        ' Kill external utilities
        WinOffline.Utility.KillProcess("tngdta")

        ' *****************************
        ' - Clear notification server cache data.
        ' *****************************

        Try

            ' Verify cache file exists
            If System.IO.File.Exists(Globals.DSMFolder + "appdata\cfnotsrvd.dat") Then

                ' Write debug
                Delegate_Sub_Append_Text(UIOutputControl, "Delete file: " + Globals.DSMFolder + "appdata\cfnotsrvd.dat" + Environment.NewLine)

                ' Delete cache file
                System.IO.File.Delete(Globals.DSMFolder + "appdata\cfnotsrvd.dat")

            End If

        Catch ex As Exception

            ' Write debug
            Delegate_Sub_Append_Text(UIOutputControl, "Exception caught cleaning up notification server data file.")
            Delegate_Sub_Append_Text(UIOutputControl, ex.Message)
            Delegate_Sub_Append_Text(UIOutputControl, ex.StackTrace)

        End Try

        ' Write debug
        Delegate_Sub_Append_Text(UIOutputControl, "CAF service: STOPPED" + Environment.NewLine)

    End Sub

    Private Sub CafStopWorker(ByVal UIOutputControl As Control)

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
            Delegate_Sub_Append_Text(UIOutputControl, "Helper thread: Expediting ""caf stop"" operation..")

            ' Loop while "caf stop" operation is running
            While WinOffline.Utility.IsProcessRunning("caf.exe", "stop")

                ' Obtain process arraylist (filtered for caf subsystem processes)
                CafProcessList = WinOffline.Utility.GetProcessChildren("caf.exe", "service", New ArrayList({"cfsmsmd.exe", "ccnfagent.exe", "encclient.exe"}))

                ' Check for an empty list
                If CafProcessList.Count = 0 Then

                    ' Write debug
                    Delegate_Sub_Append_Text(UIOutputControl, "Helper thread: Finished.")

                    ' Return
                    Return

                End If

                ' Write debug
                Delegate_Sub_Append_Text(UIOutputControl, "Helper thread: Monitoring (" + CafProcessList.Count.ToString + ") processes.")

                ' Reset loop counter
                LoopCounter = 0

                ' Rest for a short, but reasonable interval
                While LoopCounter < 50

                    ' Check for termination signal
                    If Not WinOffline.Utility.IsProcessRunning("caf.exe", "stop") Then

                        ' Write debug
                        Delegate_Sub_Append_Text(UIOutputControl, "Helper thread: Finished.")

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
                    Delegate_Sub_Append_Text(UIOutputControl, "Helper thread: Analyzing PID " + ChildProcess.Item(0).ToString + " -- " + ChildProcess.Item(1).ToString)

                    ' Check if the process is still running
                    If Not WinOffline.Utility.IsProcessRunning(Integer.Parse(ChildProcess.Item(0))) Then

                        ' Write debug
                        Delegate_Sub_Append_Text(UIOutputControl, "Helper thread: Self-terminated PID " + ChildProcess.Item(0).ToString + " -- " + ChildProcess.Item(1).ToString)

                        ' Remove from the list
                        CafProcessList.RemoveAt(x)

                        ' Continue loop
                        Continue For

                    End If

                    ' Obtain base working set memory size
                    WorkingMemoryBase = Double.Parse(ChildProcess.Item(4))

                    ' Obtain current snapshot of working set memory size
                    WorkingMemoryCurrent = WinOffline.Utility.GetProcessWorkingSetMemorySize(ChildProcess.Item(0))

                    ' Check if working set memory size has increased or remained the same
                    If WorkingMemoryCurrent >= WorkingMemoryBase Then

                        ' Kill the process
                        WinOffline.Utility.KillProcess(Integer.Parse(ChildProcess.Item(0)))

                        ' Remove from the list
                        CafProcessList.RemoveAt(x)

                        ' Write debug
                        Delegate_Sub_Append_Text(UIOutputControl, "Helper thread: Killed PID " + ChildProcess.Item(0).ToString + " -- " + ChildProcess.Item(1).ToString)

                        ' Continue loop
                        Continue For

                    End If

                    ' Write debug
                    Delegate_Sub_Append_Text(UIOutputControl, "Helper thread: Memory decreasing PID " + ChildProcess.Item(0).ToString + " -- " + ChildProcess.Item(1).ToString)

                Next

            End While

            ' Write debug
            Delegate_Sub_Append_Text(UIOutputControl, "Helper thread: Finished.")

        Catch ex As Exception

            ' Write debug
            Delegate_Sub_Append_Text(UIOutputControl, "Helper thread: Exception caught monitoring caf.exe child processes.")
            Delegate_Sub_Append_Text(UIOutputControl, ex.Message)
            Delegate_Sub_Append_Text(UIOutputControl, ex.StackTrace)

        End Try

    End Sub

End Class