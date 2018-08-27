Partial Public Class WinOffline

    Public Shared Function StageI(ByVal CallStack As String) As Integer

        ' Local variables
        Dim StateStreamWriter As System.IO.StreamWriter         ' Output stream for writing the execution state marker.
        Dim ExecutionString As String                           ' Command line to be executed externally to the application.
        Dim ArgumentString As String                            ' Arguments passed on the command line for the external execution.
        Dim RunningProcess As Process                           ' A process shell for executing the command line.
        Dim SignalReRunProcessStartInfo As ProcessStartInfo     ' Process startup settings for configuring the bahviour of the process.
        Dim OfflineProcessStartInfo As ProcessStartInfo         ' Process startup settings for configuring the bahviour of the process.
        Dim RunLevel As Integer = 0                             ' Tracks the health of the function and calls to external functions.

        ' Update call stack
        CallStack += "StageI|"

        ' Write debug
        Logger.SetCurrentTask("StageI..")

        ' *****************************
        ' - Switch: Report agent patch history only.
        ' *****************************

        ' Check for SD mode and user get history switch
        If Globals.SDBasedMode AndAlso Globals.GetHistorySwitch Then

            ' Analyze patch history files
            RunLevel = ProcessHistoryFiles(CallStack)

            ' Check the run level
            If RunLevel <> 0 Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Error: Failed to process patch history files.")

                ' Return
                Return 1

            End If

            ' Set final stage indicator
            Globals.FinalStage = True

            ' Return
            Return 0

        End If

        ' Check patch removal switch
        If Not Globals.RemovePatchSwitch Then

            ' *****************************
            ' - Process patches.
            ' *****************************

            Try

                ' Call patch processor
                RunLevel = PatchProcessor(CallStack)

                ' Check the run level
                If RunLevel <> 0 Then

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Error: Patch processor failure.")

                    ' Return
                    Return 2

                End If

            Catch ex As Exception

                ' Write debug
                Logger.WriteDebug(CallStack, "Error: Exception caught during patch processing.")
                Logger.WriteDebug(ex.Message)
                Logger.WriteDebug(ex.StackTrace)

                ' Create exception
                Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

                ' Return
                Return 3

            End Try

        End If

        ' *****************************
        ' - Copy WinOffline to temp. [SD EXECUTION ONLY]
        ' *****************************

        ' SD execution only
        If Globals.SDBasedMode Then

            ' Check if WinOffline already running from temp folder
            If Not Utility.IsProcessRunningEx(Globals.WindowsTemp + "\" + Globals.ProcessShortName) Then

                ' Delete it
                Utility.DeleteFile(CallStack, Globals.WindowsTemp + "\" + Globals.ProcessShortName)

            End If

            ' Copy WinOffline to temp
            Try

                ' Check if WinOffline already exists
                If Not System.IO.File.Exists(Globals.WindowsTemp + "\" + Globals.ProcessShortName) Then

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Copy File: " + Globals.ProcessName)
                    Logger.WriteDebug(CallStack, "To: " + Globals.WindowsTemp + "\" + Globals.ProcessShortName)

                    ' Copy to temp
                    System.IO.File.Copy(Globals.ProcessName, Globals.WindowsTemp + "\" + Globals.ProcessShortName, True)

                End If

            Catch ex As Exception

                ' Write debug
                Logger.WriteDebug(CallStack, "Error: Exception caught copying " + Globals.ProcessFriendlyName + " to temporary directory.")
                Logger.WriteDebug(ex.Message)
                Logger.WriteDebug(ex.StackTrace)

                ' Create exception
                Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

                ' Return
                Return 4

            End Try

        End If

        ' *****************************
        ' - Signal software delivery re-run request. [SD EXECUTION ONLY]
        ' *****************************

        ' Software delivery execution only (also excludes uninstall/removal)
        If Globals.SDBasedMode Then

            ' Build stage I execution string
            ExecutionString = Globals.DSMFolder + "bin\" + "sd_acmd.exe"
            ArgumentString = "signal rerun"

            ' Write debug
            Logger.WriteDebug(CallStack, "Detached process: " + ExecutionString + " " + ArgumentString)

            ' Create detached process for sd agent rerun signal
            SignalReRunProcessStartInfo = New ProcessStartInfo(ExecutionString, ArgumentString)
            SignalReRunProcessStartInfo.WorkingDirectory = Globals.DSMFolder + "bin"
            SignalReRunProcessStartInfo.UseShellExecute = False
            SignalReRunProcessStartInfo.RedirectStandardOutput = True
            SignalReRunProcessStartInfo.CreateNoWindow = True

            ' Start detached process
            RunningProcess = Process.Start(SignalReRunProcessStartInfo)

            ' Wait for detached process to exit
            RunningProcess.WaitForExit()

            ' Write debug
            Logger.WriteDebug("------------------------------------------------------------")
            Logger.WriteDebug(RunningProcess.StandardOutput.ReadToEnd.ToString)
            Logger.WriteDebug("------------------------------------------------------------")

            ' Get exit code
            RunLevel = RunningProcess.ExitCode

            ' Write debug
            Logger.WriteDebug(CallStack, "Exit code: " + RunLevel.ToString)

            ' Close detached process
            RunningProcess.Close()

            ' Check runlevel
            If RunLevel <> 0 Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Error: Failed to signal rerun request.")

                ' Create exception
                Manifest.UpdateManifest(CallStack,
                                        Manifest.EXCEPTION_MANIFEST,
                                        {"Error: Failed to signal rerun request.",
                                        "Reason: Non-zero return code from sd_acmd."})

                ' Return
                Return 5

            End If

        End If

        ' *****************************
        ' - Write cache files.
        ' *****************************

        ' Write patch manifest cache file
        Manifest.WriteCache(CallStack, Manifest.PATCH_MANIFEST)

        ' Write removal manifest cache file
        Manifest.WriteCache(CallStack, Manifest.REMOVAL_MANIFEST)

        ' *****************************
        ' - Write execution state file. [SD EXECUTION ONLY]
        ' *****************************

        ' Check for software delivery based mode
        If Globals.SDBasedMode Then

            ' Write execution state file
            Try

                ' Write debug
                Logger.WriteDebug(CallStack, "Open file: " + Globals.WinOfflineTemp + "\" + Globals.ProcessFriendlyName + ".state")
                Logger.WriteDebug(CallStack, "Write: StageI completed")

                ' Write execution state marker to file
                StateStreamWriter = New System.IO.StreamWriter(Globals.WinOfflineTemp + "\" + Globals.ProcessFriendlyName + ".state", False)
                StateStreamWriter.WriteLine("StageI completed.")

                ' Write debug
                Logger.WriteDebug(CallStack, "Close file: " + Globals.WinOfflineTemp + "\" + Globals.ProcessFriendlyName + ".state")

                ' Close execution state marker stream
                StateStreamWriter.Close()

            Catch ex As Exception

                ' Write debug
                Logger.WriteDebug(CallStack, "Error: Exception caught writing state file.")
                Logger.WriteDebug(ex.Message)
                Logger.WriteDebug(ex.StackTrace)

                ' Create exception
                Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

                ' Return
                Return 6

            End Try

        End If

        ' *****************************
        ' - Create detached offline process. [SD EXECUTION ONLY]
        ' *****************************

        ' Software delivery execution only
        If Globals.SDBasedMode Then

            ' Build stageII execution string
            ExecutionString = Globals.WindowsTemp + "\" + Globals.ProcessShortName
            ArgumentString = ""

            ' Build stageII argument string
            For Each arg As String In Globals.CommandLineArgs

                ' Add the argument to the list
                ArgumentString += " " + arg

            Next

            ' Create process based on process identity
            If Globals.RunningAsSystemIdentity Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Detached process: " + ExecutionString + " " + ArgumentString)

                ' Create detached process For stage II execution
                OfflineProcessStartInfo = New ProcessStartInfo(ExecutionString)
                OfflineProcessStartInfo.WorkingDirectory = System.IO.Path.GetDirectoryName(ExecutionString)
                OfflineProcessStartInfo.Arguments = ArgumentString

                ' Start detached process
                Process.Start(OfflineProcessStartInfo)

            Else

                ' Call the launch pad
                RunLevel = LaunchPad(CallStack, "system", Globals.WindowsTemp + "\" + Globals.ProcessShortName, Globals.WindowsTemp + "\", ArgumentString)

                ' Check runlevel
                If RunLevel = 0 Then

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Detached process launched as system identity.")

                Else

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Warning: Failed to launch detached process.")

                End If

            End If

        End If

        ' Set flag
        Globals.StageICompleted = True

        ' Return
        Return RunLevel

    End Function

End Class