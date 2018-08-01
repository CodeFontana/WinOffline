'****************************** Class Header *******************************\
' Project Name: WinOffline
' Class Name:   WinOffline/PipeClient
' File Name:    PipeClient.vb
' Author:       Brian Fontana
'***************************************************************************/

Imports System.IO
Imports System.Threading

Partial Public Class WinOffline

    Public Class PipeClient

        Private Shared InactivityCounter As Integer = 0

        Public Shared Sub InitPipeClient(ByVal CallStack As String)

            ' Update call stack
            CallStack += "PipeClient|"

            ' *****************************
            ' - Remove previous WinOffline.
            ' *****************************

            ' Check if WinOffline already running from temp folder
            If Not Utility.IsProcessRunningEx(Globals.WindowsTemp + "\" + Globals.ProcessShortName) Then

                ' Delete it
                Utility.DeleteFile(CallStack, Globals.WindowsTemp + "\" + Globals.ProcessShortName)

            End If

            ' *****************************
            ' - Deploy pipe client executable.
            ' *****************************

            ' Check if pipe client app is already deployed
            If Not System.IO.File.Exists(Globals.WindowsTemp + "\" + Globals.ProcessShortName) Then

                ' Deploy pipe client
                Try

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Copy File: " + Globals.ProcessName)
                    Logger.WriteDebug(CallStack, "To: " + Globals.WindowsTemp + "\" + Globals.ProcessShortName)

                    ' Copy to temp
                    System.IO.File.Copy(Globals.ProcessName, Globals.WindowsTemp + "\" + Globals.ProcessShortName, True)

                Catch ex As Exception

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Error: Exception caught copying " + Globals.ProcessShortName + "to temporary directory.")
                    Logger.WriteDebug(ex.Message)
                    Logger.WriteDebug(ex.StackTrace)

                    ' Create exception
                    Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

                    ' Return
                    Return

                End Try

            End If

            ' *****************************
            ' - Launch pipe client sessions.
            ' *****************************

            ' Check for active clients
            If Utility.IsProcessRunning(Globals.ProcessShortName, "-client") Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Active pipe client(s) detected.")

            Else

                ' Call the launch pad
                If LaunchPad(CallStack, "all", Globals.WindowsTemp + "\" + Globals.ProcessShortName, Globals.WindowsTemp + "\", "-client") = 0 Then

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Pipe client(s) deployed for all users.")

                Else

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Warning: Pipe client(s) deployment failed.")

                End If

            End If

        End Sub

        Public Shared Sub PipeClientApp()

            ' Local variables
            Dim ServerPipe As String
            Dim TerminateSignal As Boolean = False
            Dim StrLine As String = ""
            Dim LastMaxOffset As Long = 0
            Dim InactivityTimer As Timers.Timer
            Dim Autoclose As Boolean = False

            ' Create the progress gui
            Globals.ProgressGUI = New ProgressUI()
            Globals.ProgressGUI.Hide()

            ' Build server pipe string
            ServerPipe = Globals.WorkingDirectory + Globals.ProcessFriendlyName + ".pipe"

            ' Check if server pipe file exists
            If System.IO.File.Exists(ServerPipe) Then

                ' Open server pipe stream
                Globals.PipeClientStream = New System.IO.FileStream(ServerPipe, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                Globals.PipeClientReader = New System.IO.StreamReader(Globals.PipeClientStream)

                ' Set inactivity timer interval and enable it
                InactivityTimer = New System.Timers.Timer(30000)
                InactivityTimer.Enabled = True

                ' Link the elapsed event for the timer to our handler
                AddHandler InactivityTimer.Elapsed, AddressOf ElapsedPipeClient

                ' Loop indefinitely
                While Not TerminateSignal

                    ' Check the inactivity count
                    If InactivityCounter >= 5 Then

                        ' Close the GUI
                        Globals.ProgressGUI.Delegate_ProgressUI_Close()

                        ' Terminate
                        TerminateSignal = True

                    End If

                    ' Check if file size has changed
                    If Globals.PipeClientReader.BaseStream.Length = LastMaxOffset Then

                        ' Check if inactivity timer is already enabled
                        If InactivityTimer.Enabled = False Then

                            ' Start the inactivity timer
                            InactivityTimer.Enabled = True
                            InactivityTimer.Start()

                        End If

                        ' Allow progress gui to process its message queue
                        Globals.ProgressGUI.DoEvents()

                        ' Delay execution
                        Thread.Sleep(Globals.THREAD_REST_INTERVAL)

                        ' Restart loop
                        Continue While

                    Else

                        ' Stop the inactivity timer, reset loop counter
                        InactivityTimer.Stop()
                        InactivityTimer.Enabled = False
                        InactivityCounter = 0

                        ' Seek the previous maximum offset
                        Globals.PipeClientReader.BaseStream.Seek(LastMaxOffset, SeekOrigin.Begin)

                        ' Read messages until EOF
                        While Utility.InlineAssignHelper(StrLine, Globals.PipeClientReader.ReadLine()) IsNot Nothing AndAlso Not TerminateSignal

                            ' Dispatch by request type
                            If StrLine.StartsWith(PipeServer.SEND_SETTING) Then

                                ' Parse the settings
                                If StrLine.Contains("ShowGUISwitch:") Then

                                    ' Set global
                                    Globals.ShowGUISwitch = Boolean.Parse(StrLine.Substring(StrLine.LastIndexOf(":") + 1))

                                ElseIf StrLine.Contains("SDOnlineMode:") Then

                                    ' Set global
                                    Globals.SDBasedMode = Boolean.Parse(StrLine.Substring(StrLine.LastIndexOf(":") + 1))

                                ElseIf StrLine.Contains(Environment.UserName) Then

                                    ' Check for progress gui switch
                                    If StrLine.Contains("ShowGUISwitch:") Then

                                        ' Set global
                                        Globals.ShowGUISwitch = Boolean.Parse(StrLine.Substring(StrLine.LastIndexOf(":") + 1))

                                    End If

                                Else

                                    ' Unrecognized setting -- ignore

                                End If

                                ' Check conditions for launching progress gui
                                If Globals.ShowGUISwitch AndAlso Globals.ProgressUIThread Is Nothing Then

                                    ' Start progress gui
                                    Globals.ProgressGUI.SetFormTitle(Globals.ProcessFriendlyName + " -- " + Globals.AppVersion)
                                    Globals.ProgressUIThread = New System.Threading.Thread(AddressOf Globals.ProgressGUI.ShowDialog)
                                    Globals.ProgressUIThread.Start()

                                End If

                                ' Check for tray icon policy
                                If Globals.TrayIconVisible And Not Globals.ProgressGUI.TrayIcon.Visible Then

                                    ' Enable progress gui notification icon
                                    Globals.ProgressGUI.TrayIcon.Visible = True

                                End If

                            ElseIf StrLine.StartsWith(PipeServer.SEND_MESSAGE) Then

                                ' Forward filtered message to gui
                                Globals.ProgressGUI.Delegate_ProgressUI_WriteDebug(StrLine.Remove(0, PipeServer.SEND_MESSAGE.Length))

                            ElseIf StrLine.StartsWith(PipeServer.SEND_STATUS) Then

                                ' Forward status update to gui
                                Globals.ProgressGUI.Delegate_ProgressUI_UpdateCurrentTask(StrLine.Remove(0, PipeServer.SEND_STATUS.Length))

                            ElseIf StrLine.StartsWith(PipeServer.REFRESH_SYSTRAY) Then

                                ' Encapsulate refresh
                                Try

                                    ' Refresh notification tray area
                                    WindowsAPI.RefreshNotificationArea()

                                Catch ex As Exception

                                    ' Do nothing

                                End Try

                            ElseIf StrLine.StartsWith(PipeServer.CURTAINS_CALL) Then

                                ' Parse autoclose setting 
                                Autoclose = Boolean.Parse(StrLine.Substring(StrLine.LastIndexOf(":") + 1))

                                ' Autoclose the debug console
                                '   [TRUE = Autoclose based on Timer, FALSE = Hide immediately]
                                Globals.ProgressGUI.Delegate_ProgressUI_GoModeless(Autoclose)

                                ' Exit condition
                                TerminateSignal = True

                                ' Exit while
                                Exit While

                            Else

                                ' Unrecognized message -- ignore

                            End If

                        End While

                        ' Update max offset
                        LastMaxOffset = Globals.PipeClientReader.BaseStream.Position

                        ' Check for termination signal
                        If TerminateSignal Then Exit While

                        ' Allow progress gui to process its message queue
                        Globals.ProgressGUI.DoEvents()

                    End If

                End While

                ' Close client pipe stream
                Globals.PipeClientReader.Close()
                Globals.PipeClientStream.Close()

            Else

                ' Return -- Server pipe is required
                Throw New Exception("Server pipe is unavailable: " + ServerPipe)

            End If

        End Sub

        Private Shared Sub ElapsedPipeClient()

            ' Local variables
            Dim ExecutionString As String
            Dim ArgumentString As String
            Dim RunningProcess As Process
            Dim ProcessStartInfo As ProcessStartInfo

            ' Is anyone else out there?
            If Globals.SDBasedMode AndAlso
                    Not (Utility.IsProcessRunning(Globals.ProcessFriendlyName) OrElse
                    Utility.IsProcessRunning("sd_jexec") OrElse
                    Utility.IsProcessRunning("cfbasichwwnt.exe")) Then

                ' *****************************
                ' - Trigger an agent registration
                ' *****************************

                ' Check ITCM path for trailing '\'
                If Globals.DSMFolder.EndsWith("\") Then

                    ' Build execution string
                    ExecutionString = Globals.DSMFolder + "bin\caf.exe"
                    ArgumentString = "register"

                Else

                    ' Build execution string
                    ExecutionString = Globals.DSMFolder + "\bin\caf.exe"
                    ArgumentString = "register"

                End If

                ' Create detached process
                ProcessStartInfo = New ProcessStartInfo(ExecutionString, ArgumentString)
                ProcessStartInfo.WorkingDirectory = FileVector.GetFilePath(ExecutionString)
                ProcessStartInfo.UseShellExecute = False
                ProcessStartInfo.RedirectStandardOutput = False
                ProcessStartInfo.CreateNoWindow = True

                ' Start detached process
                RunningProcess = Process.Start(ProcessStartInfo)

                ' Wait for detached process to exit
                RunningProcess.WaitForExit()

                ' Close detached process
                RunningProcess.Close()

            End If

            ' Increment inactivity counter
            PipeClient.InactivityCounter += 1

        End Sub

    End Class

End Class