'****************************** Class Header *******************************\
' Project Name: WinOffline
' Class Name:   WinOffline
' File Name:    StageIII.vb
' Author:       Brian Fontana
'***************************************************************************/

Partial Public Class WinOffline

    Public Shared Function StageIII(ByVal CallStack As String) As Integer

        ' Local variables
        Dim ExecutionString As String                           ' Command line to be executed externally to the application.
        Dim ArgumentString As String                            ' Arguments passed on the command line for the external execution.
        Dim RunningProcess As Process                           ' A process shell for executing the command line.
        Dim SignalReRunProcessStartInfo As ProcessStartInfo     ' Process startup settings for configuring the bahviour of the process.
        Dim RunLevel As Integer = 0

        ' Update call stack
        CallStack += "StageIII|"

        ' Write debug
        Logger.SetCurrentTask("StageIII..")

        ' *****************************
        ' - Read exception cache.
        ' *****************************

        Try

            ' Restore from cache
            Manifest.ReadCache(CallStack, Manifest.EXCEPTION_MANIFEST)

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Exception caught reading cache file.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

            ' Create exception
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

            ' Increment RunLevel to indicate error
            RunLevel += 1

        End Try

        ' *****************************
        ' - Read history cache.
        ' *****************************

        Try

            ' Restore history cache
            Manifest.ReadCache(CallStack, Manifest.HISTORY_MANIFEST)

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Exception caught reading cache file.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

            ' Create exception
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

            ' Increment RunLevel to indicate error
            RunLevel += 1

        End Try

        ' *****************************
        ' - Read patch cache.
        ' *****************************

        Try

            ' Read patch cache
            Manifest.ReadCache(CallStack, Manifest.PATCH_MANIFEST)

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Exception caught reading cache file.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

            ' Create exception
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

            ' Increment RunLevel to indicate error
            RunLevel += 1

        End Try

        ' *****************************
        ' - Read removal cache.
        ' *****************************

        Try

            ' Read patch cache
            Manifest.ReadCache(CallStack, Manifest.REMOVAL_MANIFEST)

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Exception caught reading cache file.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

            ' Create exception
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

            ' Increment RunLevel to indicate error
            RunLevel += 1

        End Try

        ' *****************************
        ' - Switch: Perform software library cleanup.
        ' *****************************

        Try

            ' Check for library cleanup switch
            If Globals.CheckSDLibrarySwitch Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Switch: Analyze software library without making changes.")

                ' Analyze library
                LibraryManager.RepairLibrary(CallStack)

            ElseIf Globals.CleanupSDLibrarySwitch Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Switch: Perform software library cleanup.")

                ' Repair library
                LibraryManager.RepairLibrary(CallStack)

            End If

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Warning: Exception caught during software library cleanup.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

            ' Create exception
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

        End Try

        ' *****************************
        ' - Switch: Signal software delivery reboot request. [SD EXECUTION ONLY]
        ' *****************************

        ' Software delivery execution only
        If Globals.SDBasedMode And Globals.SDSignalRebootSwitch Then

            ' Build stage I execution string
            ExecutionString = Globals.DSMFolder + "bin\" + "sd_acmd.exe"
            ArgumentString = "signal reboot"

            ' Write debug
            Logger.WriteDebug(CallStack, "Detached process: " + ExecutionString + " " + ArgumentString)

            ' Create detached process for sd agent reboot signal
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
                Logger.WriteDebug(CallStack, "Error: Failed to signal reboot request.")

                ' Create exception
                Manifest.UpdateManifest(CallStack,
                                        Manifest.EXCEPTION_MANIFEST,
                                        {"Error: Failed to signal reboot request.",
                                        "Reason: Non-zero return code from sd_acmd."})

                ' Increment RunLevel to indicate error
                RunLevel += 1

            End If

        End If

        ' *****************************
        ' - Switch: Initiate a system reboot.
        ' *****************************

        ' Check for system reboot switch
        If Globals.RebootSystemSwitch Then

            ' Set reboot on termination flag
            Globals.RebootOnTermination = True

        End If

        ' *****************************
        ' - Check runlevel and exit.
        ' *****************************

        ' Check RunLevel
        If RunLevel <> 0 Then

            ' Write debug
            Logger.WriteDebug(CallStack, "StageIII execution completed with an error.")

        End If

        ' Set flag
        Globals.StageIIICompleted = True

        ' Return
        Return RunLevel

    End Function

End Class