Partial Public Class WinOffline

    Public Shared Function StageII(ByVal CallStack As String) As Integer

        ' Local variables
        Dim StateStreamWriter As System.IO.StreamWriter             ' Output stream for writing the execution state marker.
        Dim ExecutionString As String                               ' Command line to be executed externally to the application.
        Dim ArgumentString As String                                ' Arguments passed on the command line for the external execution.
        Dim RunningProcess As Process                               ' A process shell for executing the command line.
        Dim ProcessStartInfo As ProcessStartInfo                    ' Process startup settings for configuring the bahviour of the process.
        Dim PatchErrorReported As Boolean = False                   ' Flag indicating a patch error was reported.
        Dim HistoryFile As String                                   ' Name of patch history file.
        Dim BackupFileStamp As String                               ' Timestamp for saving a copy of the cam.cfg file
        Dim FileList As String()                                    ' Array for storing list of files.
        Dim strFile As String                                       ' Used for filename comparison.
        Dim hexString As String                                     ' Used for BHI temp file cleanup.
        Dim HostUUIDKey As Microsoft.Win32.RegistryKey = Nothing    ' Used for HostUUID regeneration.
        Dim RunLevel As Integer = 0                                 ' Tracks the health of the function and calls to external functions.

        ' Update call stack
        CallStack += "StageII|"

        ' Write debug
        Logger.SetCurrentTask("StageII..")

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

            ' Return
            Return 1

        End Try

        ' *****************************
        ' - Read patch cache.
        ' *****************************

        Try

            ' Read the patch manifest
            Manifest.ReadCache(CallStack, Manifest.PATCH_MANIFEST)

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Exception caught reading cache file.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

            ' Create exception
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

            ' Return
            Return 2

        End Try

        ' *****************************
        ' - Read removal cache.
        ' *****************************

        Try

            ' Read the removal manifest
            Manifest.ReadCache(CallStack, Manifest.REMOVAL_MANIFEST)

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Exception caught reading cache file.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

            ' Create exception
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

            ' Return
            Return 3

        End Try

        ' *****************************
        ' - Disable ENC client.
        ' *****************************

        ' Check for disable ENC switch
        If Globals.DisableENCSwitch AndAlso System.IO.File.Exists(Globals.DSMFolder + "bin\encutilcmd.exe") Then

            ' Write debug
            Logger.WriteDebug(CallStack, "Switch: Disable ENC client.")

            ' Build execution string
            ExecutionString = Globals.DSMFolder + "bin\encutilcmd.exe"
            ArgumentString = "client -state disabled"

            ' Write debug
            Logger.WriteDebug(CallStack, "Detached process: " + ExecutionString + " " + ArgumentString)

            ' Create detached process
            ProcessStartInfo = New ProcessStartInfo(ExecutionString, ArgumentString)
            ProcessStartInfo.WorkingDirectory = Globals.DSMFolder + "bin"
            ProcessStartInfo.UseShellExecute = False
            ProcessStartInfo.RedirectStandardOutput = True
            ProcessStartInfo.CreateNoWindow = True

            ' Start detached process
            RunningProcess = Process.Start(ProcessStartInfo)

            ' Wait for detached process to exit
            RunningProcess.WaitForExit()

            ' Write debug
            Logger.WriteDebug("------------------------------------------------------------")
            Logger.WriteDebug(RunningProcess.StandardOutput.ReadToEnd.ToString)
            Logger.WriteDebug("------------------------------------------------------------")

            ' Write debug
            Logger.WriteDebug(CallStack, "Exit code: " + RunningProcess.ExitCode.ToString)

            ' Close detached process
            RunningProcess.Close()

        ElseIf Not System.IO.File.Exists(Globals.DSMFolder + "bin\encutilcmd.exe") Then

            ' Write debug
            Logger.WriteDebug(CallStack, "Switch ignored: Disable ENC client.")
            Logger.WriteDebug(CallStack, "The encutilcmd.exe binary is not available.")

        End If

        ' *****************************
        ' - Stop client automation services.
        ' *****************************

        ' Check for simulation
        If Globals.SimulateCafStopSwitch Then

            ' Write debug
            Logger.WriteDebug(CallStack, "Switch: Simulate stopping services.")

        Else

            ' Stop services
            RunLevel = StopServices(CallStack)

        End If

        ' Write debug
        Logger.SetCurrentTask("StageII..")

        ' Check the run level
        If RunLevel <> 0 Then

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Client Automation service termination failure.")

            ' Attempt restart of all services
            RunLevel = StartServices(CallStack)

            ' Return -- Don't proceed any further. Report the problem stopping services.
            Return 4

        End If

        ' *****************************
        ' - Automatic: Clear notification server cache data.
        ' *****************************

        ' Delete notification cache file
        Utility.DeleteFile(CallStack, Globals.DSMFolder + "appdata\cfnotsrvd.dat")

        ' *****************************
        ' - Automatic: Cleanup basic agent temporary files.
        ' *****************************

        Try

            ' Write debug
            Logger.WriteDebug(CallStack, "Automatic: Cleanup basic agent temporary files.")

            ' Get the directory listing
            FileList = System.IO.Directory.GetFiles(Globals.WindowsTemp)

            ' Check for positive files
            If FileList.Length > 0 Then

                ' Loop the file list
                For n As Integer = 0 To FileList.Length - 1

                    ' Get a filename
                    strFile = FileList(n).ToString.ToLower
                    strFile = strFile.Substring(strFile.LastIndexOf("\") + 1)

                    ' Example of files:
                    ' bhiA.tmp
                    ' bhiA0.tmp
                    ' bhiA01.tmp
                    ' bhiA012.tmp

                    ' Check if filename starts with 'bhi' and ends with '.tmp'
                    If strFile.StartsWith("bhi") And
                        strFile.EndsWith(".tmp") Then

                        ' Parse the HEX substring
                        hexString = strFile.Substring(3, strFile.Length - 3)
                        hexString = hexString.Substring(0, hexString.IndexOf("."))

                        ' Check if filename only contains HEX
                        If Utility.IsHexString(hexString) Then

                            ' Attempt to delete each file
                            Try

                                ' Unset read-only parameter (in case it's set)
                                System.IO.File.SetAttributes(FileList(n), IO.FileAttributes.Normal)

                                ' Delete the file
                                Utility.DeleteFile(CallStack, FileList(n))

                            Catch ex2 As Exception

                                ' Write debug
                                Logger.WriteDebug(CallStack, "Warning: Exception caught deleting file: " + FileList(n).ToString)
                                Logger.WriteDebug(ex2.Message)
                                Logger.WriteDebug(ex2.StackTrace)

                            End Try

                        End If

                    End If

                Next

            End If

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Warning: Exception caught cleaning up temporary files.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

            ' Create exception
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

        End Try

        ' *****************************
        ' - Switch: Reset cftrace level.
        ' *****************************

        Try

            ' Check for cftrace reset switch
            If Globals.ResetCftraceSwitch Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Switch: Reset cftrace level.")

                ' Build execution string
                ExecutionString = Globals.DSMFolder + "bin\cftrace.exe"
                ArgumentString = "-c reset"

                ' Write debug
                Logger.WriteDebug(CallStack, "Detached process: " + ExecutionString + " " + ArgumentString)

                ' Create detached process
                ProcessStartInfo = New ProcessStartInfo(ExecutionString, ArgumentString)
                ProcessStartInfo.WorkingDirectory = Globals.DSMFolder + "bin"
                ProcessStartInfo.UseShellExecute = False
                ProcessStartInfo.RedirectStandardOutput = True
                ProcessStartInfo.CreateNoWindow = True

                ' Start detached process
                RunningProcess = Process.Start(ProcessStartInfo)

                ' Wait for detached process to exit
                RunningProcess.WaitForExit()

                ' Write debug
                Logger.WriteDebug("------------------------------------------------------------")
                Logger.WriteDebug(RunningProcess.StandardOutput.ReadToEnd.ToString)
                Logger.WriteDebug("------------------------------------------------------------")

                ' Write debug
                Logger.WriteDebug(CallStack, "Exit code: " + RunningProcess.ExitCode.ToString)

                ' Close detached process
                RunningProcess.Close()

            End If

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Warning: Exception caught resetting cftrace level.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

            ' Create exception
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

        End Try

        ' *****************************
        ' - Switch: Cleanup DSM logs folder.
        ' *****************************

        Try

            ' Check for DSM logs cleanup switch
            If Globals.CleanupLogsSwitch Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Switch: Cleanup various log folders.")

                ' Check if logs folder exists
                If System.IO.Directory.Exists(Globals.DSMFolder + "logs") Then

                    ' Delete folder contents
                    Utility.DeleteFolderContents(CallStack, Globals.DSMFolder + "logs", Nothing)

                End If

                ' Check if logs folder exists
                If System.IO.Directory.Exists(Globals.CAMFolder + "\logs") Then

                    ' Delete folder contents
                    Utility.DeleteFolderContents(CallStack, Globals.CAMFolder + "\logs", Nothing)

                End If

                ' Check if logs folder exists
                If System.IO.Directory.Exists(Globals.SSAFolder + "log") Then

                    ' Delete folder contents
                    Utility.DeleteFolderContents(CallStack, Globals.SSAFolder + "log", Nothing)

                End If

            End If

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Warning: Exception caught deleting logs folder.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

            ' Create exception
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

        End Try

        ' *****************************
        ' - Switch: Remove CAM configuration file.
        ' *****************************

        Try

            ' Check for remove CAM config switch
            If Globals.RemoveCAMConfigSwitch Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Switch: Remove CAM configuration file.")

                ' Check if cam.cfg exists
                If System.IO.File.Exists(Globals.CAMFolder + "\cam.cfg") Then

                    ' Create backup extension
                    BackupFileStamp = DateTime.Now.Year.ToString() + "." +
                            DateTime.Now.Month.ToString() + "." +
                            DateTime.Now.Day.ToString() + "--" +
                            DateTime.Now.TimeOfDay.ToString()

                    ' Prune insignificant time digits
                    BackupFileStamp = BackupFileStamp.Substring(0, BackupFileStamp.LastIndexOf("."))

                    ' Append file extenstion
                    BackupFileStamp = BackupFileStamp + ".bak"

                    ' Update time delimiters
                    BackupFileStamp = BackupFileStamp.Replace(":", ".")

                    ' Save a backup file
                    System.IO.File.Move(Globals.CAMFolder + "\cam.cfg", Globals.CAMFolder + "\cam--" + BackupFileStamp)

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Backup of CAM configuration file saved: " + Globals.CAMFolder + "\cam--" + BackupFileStamp)

                End If

                ' Check if cam.cfg exists
                If System.IO.File.Exists(Globals.CAMFolder + "\cam.cfg") Then

                    ' Unset read-only parameter (in case it's set)
                    System.IO.File.SetAttributes(Globals.CAMFolder + "\cam.cfg", IO.FileAttributes.Normal)

                    ' Delete the file
                    Utility.DeleteFile(CallStack, Globals.CAMFolder + "\cam.cfg")

                End If

                ' Check if cam.bak exists
                If System.IO.File.Exists(Globals.CAMFolder + "\cam.bak") Then

                    ' Unset read-only parameter (in case it's set)
                    System.IO.File.SetAttributes(Globals.CAMFolder + "\cam.bak", IO.FileAttributes.Normal)

                    ' Delete the file
                    Utility.DeleteFile(CallStack, Globals.CAMFolder + "\cam.bak")

                End If

            End If

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Warning: Exception caught deleting the CAM configuration file.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

            ' Create exception
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

        End Try

        ' *****************************
        ' - Switch: Perform agent cleanup.
        ' *****************************

        Try

            ' Check for agent cleanup switch
            If Globals.CleanupAgentSwitch Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Switch: Perform agent cleanup.")

                ' Cleanup AM/SD units folder
                If System.IO.Directory.Exists(Globals.DSMFolder + "Agent\units") Then

                    ' Delete unit account folders
                    Utility.DeleteFolderContents(CallStack, Globals.DSMFolder + "Agent\units", Nothing)

                End If

                ' Cleanup TMP activate folder
                If Globals.SDFolder IsNot Nothing AndAlso System.IO.Directory.Exists(Globals.SDFolder) Then

                    ' Check if the agent activate folder exists
                    If System.IO.Directory.Exists(Globals.SDFolder + "\TMP\activate") Then

                        ' Delete folder contents
                        Utility.DeleteFolderContents(CallStack, Globals.SDFolder + "\TMP\activate", Nothing)

                    End If

                End If

                ' Cleanup DTS agent folder
                If Globals.DTSFolder IsNot Nothing Then

                    ' Delete folder contents
                    Utility.DeleteFolderContents(CallStack, Globals.DSMFolder + "dts\dta\staging", Nothing)

                    ' Delete folder contents
                    Utility.DeleteFolderContents(CallStack, Globals.DSMFolder + "dts\dta\status", Nothing)

                    ' Recreate index file
                    System.IO.File.Create(Globals.DSMFolder + "dts\dta\status\index")

                    ' Write debug
                    Logger.WriteDebug(CallStack, "DTS index file recreated: " + Globals.DSMFolder + "dts\dta\status\index")

                End If

            End If

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Warning: Exception caught during agent cleanup.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

            ' Create exception
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

        End Try

        ' *****************************
        ' - Switch: Perform scalability server cleanup.
        ' *****************************

        Try

            ' Check for server cleanup switch
            If Globals.CleanupServerSwitch Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Switch: Perform scalability server cleanup.")

                ' Check if SDROOT exists
                If Globals.SDFolder IsNot Nothing AndAlso System.IO.Directory.Exists(Globals.SDFolder) Then

                    ' Check if file transfer journal folder exists
                    If System.IO.Directory.Exists(Globals.SDFolder + "\ASM\DATABASE\ftm_journals") Then

                        ' Create CLEANSTART folder
                        System.IO.Directory.CreateDirectory(Globals.SDFolder + "\ASM\DATABASE\ftm_journals\CLEANSTART")

                    End If

                    ' Check if deliveries folder exists
                    If System.IO.Directory.Exists(Globals.SDFolder + "\ASM\D") Then

                        ' Delete folder contents
                        Utility.DeleteFolderContents(CallStack, Globals.SDFolder + "\ASM\D", Nothing)

                    End If

                    ' Check if the server activate folder exists
                    If System.IO.Directory.Exists(Globals.SDLibraryFolder + "\activate") Then

                        ' Delete folder contents
                        Utility.DeleteFolderContents(CallStack, Globals.SDLibraryFolder + "\activate", {"lgzip", "scripts", "scripts.zip"})

                    End If

                    ' Check if DTS is installed (and the agent cleanup didn't already handle this)
                    If Globals.DTSFolder IsNot Nothing And Not Globals.CleanupAgentSwitch Then

                        ' Delete folder contents
                        Utility.DeleteFolderContents(CallStack, Globals.DSMFolder + "dts\dta\staging", Nothing)

                        ' Delete folder contents
                        Utility.DeleteFolderContents(CallStack, Globals.DSMFolder + "dts\dta\status", Nothing)

                        ' Recreate index file
                        System.IO.File.Create(Globals.DSMFolder + "dts\dta\status\index")

                        ' Write debug
                        Logger.WriteDebug(CallStack, "DTS index file recreated: " + Globals.DSMFolder + "dts\dta\status\index")

                    End If

                End If

            End If

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Warning: Exception caught during scalability server cleanup.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

            ' Create exception
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

        End Try

        ' *****************************
        ' - Switch: Certificate store cleanup.
        ' *****************************

        Try

            ' Check for cert store cleanup switch
            If Globals.CleanupCertStoreSwitch Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Switch: Certificate store cleanup.")

                ' Perform repair operation
                CertManager.RepairStore(CallStack)

            End If

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Warning: Exception caught repairing certificate store.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

            ' Create exception
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

        End Try

        ' *****************************
        ' - Switch: Remove patch history file, BEFORE patch operations.
        ' *****************************

        Try

            ' Check for history cleanup switch
            If Globals.RemoveHistoryBeforeSwitch Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Switch: Remove patch history file, BEFORE patch operations.")

                ' Calculate history file
                HistoryFile = Globals.DSMFolder + Globals.HostName + ".his"

                ' Delete history file
                Utility.DeleteFile(CallStack, HistoryFile)

                ' Check if REPLACED folder exists
                If System.IO.Directory.Exists(Globals.DSMFolder + "REPLACED") Then

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Delete folder: " + Globals.DSMFolder + "REPLACED")

                    ' Delete folder contents
                    Utility.DeleteFolderContents(CallStack, Globals.DSMFolder + "REPLACED", Nothing)

                    ' Delete the folder
                    System.IO.Directory.Delete(Globals.DSMFolder + "REPLACED", True)

                End If

            End If

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Warning: Exception caught deleting patch history.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

            ' Create exception
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

        End Try

        ' *****************************
        ' - Analyze patch history files.
        ' *****************************

        ' Analyze patch history files
        RunLevel = ProcessHistoryFiles(CallStack)

        ' Check the run level
        If RunLevel <> 0 Then

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Failed to process patch history files.")

            ' Attempt restart of all services
            RunLevel = StartServices(CallStack)

            ' Return -- Don't proceed any further. Report the problem parsing history files.
            Return 5

        End If

        ' *****************************
        ' - Perform patch operations.
        ' *****************************

        ' Perform patch operations
        RunLevel = PatchOperations(CallStack)

        ' Write debug
        Logger.SetCurrentTask("StageII..")

        ' Check the run level
        If RunLevel <> 0 Then

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Patch operation method reports a failure..")

            ' Set patch error flag -- Because we don't want to return without starting services. Just set the marker.
            PatchErrorReported = True

        End If

        ' *****************************
        ' - Switch: Remove patch history file, AFTER patch operations.
        ' *****************************

        Try

            ' Check for history cleanup switch
            If Globals.RemoveHistoryAfterSwitch Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Switch: Remove patch history file, AFTER patch operations.")

                ' Calculate history file
                HistoryFile = Globals.DSMFolder + Globals.HostName + ".his"

                ' Delete history file
                Utility.DeleteFile(CallStack, HistoryFile)

                ' Check if REPLACED folder exists
                If System.IO.Directory.Exists(Globals.DSMFolder + "REPLACED") Then

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Delete folder: " + Globals.DSMFolder + "REPLACED")

                    ' Delete folder contents
                    Utility.DeleteFolderContents(CallStack, Globals.DSMFolder + "REPLACED", Nothing)

                    ' Delete the folder
                    System.IO.Directory.Delete(Globals.DSMFolder + "REPLACED", True)

                End If

            End If

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Warning: Exception caught deleting patch history.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

            ' Create exception
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

        End Try

        ' *****************************
        ' - Switch: Regenerate agent HostUUID.
        ' *****************************

        Try

            ' Check for HostUUID regeneration switch
            If Globals.RegenHostUUIDSwitch Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Switch: Regenerate HostUUID.")

                ' Perform comstore operations
                Logger.WriteDebug(CallStack, "Delete comstore parameter: itrm/rc/host/managed/convertedhostuid")
                ComstoreAPI.DeleteParameter("itrm/rc/host/managed", "convertedhostuid")
                Logger.WriteDebug(CallStack, "Delete comstore parameter: itrm/rc/host/managed/hostuuidcopy")
                ComstoreAPI.DeleteParameter("itrm/rc/host/managed", "hostuuidcopy")
                Logger.WriteDebug(CallStack, "Delete comstore parameter: itrm/cfencrypt/LOCALID")
                ComstoreAPI.DeleteParameter("itrm/cfencrypt", "LOCALID")
                Logger.WriteDebug(CallStack, "Delete comstore section: itrm/rc/security/providers/winnt/users")
                ComstoreAPI.DeleteParameterSection("itrm/rc/security/providers/winnt/users")

                ' Perform registry operations
                Logger.WriteDebug(CallStack, "Delete registry key: " + "HKLM\SOFTWARE\Wow6432Node\ComputerAssociates\hostUUID")
                Microsoft.Win32.Registry.LocalMachine.DeleteSubKey("SOFTWARE\Wow6432Node\ComputerAssociates\hostUUID", False)
                Logger.WriteDebug(CallStack, "Delete registry key: " + "HKLM\SOFTWARE\ComputerAssociates\hostUUID")
                Microsoft.Win32.Registry.LocalMachine.DeleteSubKey("SOFTWARE\ComputerAssociates\hostUUID", False)

            End If

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Warning: Exception caught regenerating HostUUID.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

            ' Create exception
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

        End Try

        ' *****************************
        ' - Start client automation services.
        ' *****************************

        ' Check for simulation
        If Globals.SimulateCafStopSwitch Then

            ' Write debug
            Logger.WriteDebug(CallStack, "Switch: Simulate starting services.")

        Else

            ' Start services
            RunLevel = StartServices(CallStack)

        End If

        ' Check the run level
        If RunLevel <> 0 Then

            ' Write debug
            Logger.WriteDebug(CallStack, "Warning: Start services method reports a failure.")

        End If

        ' *****************************
        ' - Switch: Launch DSM Explorer
        ' *****************************

        ' User-based execution only
        If Globals.LaunchGuiSwitch And
            Not Globals.RunningAsSystemIdentity And
            System.IO.File.Exists(Globals.DSMFolder + "bin\dsmgui.exe") Then

            ' Write debug
            Logger.WriteDebug(CallStack, "Switch: Launch DSM Explorer.")

            ' Build execution string
            ExecutionString = Globals.DSMFolder + "bin\dsmgui.exe"
            ArgumentString = ""

            ' Write debug
            Logger.WriteDebug(CallStack, "Detached process: " + ExecutionString + " " + ArgumentString)

            ' Create detached process
            ProcessStartInfo = New ProcessStartInfo(ExecutionString, ArgumentString)
            ProcessStartInfo.WorkingDirectory = Globals.DSMFolder + "bin"
            ProcessStartInfo.UseShellExecute = False
            ProcessStartInfo.RedirectStandardOutput = True
            ProcessStartInfo.CreateNoWindow = True

            ' Start detached process
            RunningProcess = Process.Start(ProcessStartInfo)

            ' Wait for detached process to exit
            RunningProcess.WaitForExit()

            ' Close detached process
            RunningProcess.Close()

            ' Write debug
            Logger.WriteDebug(CallStack, "DSM Explorer launched.")

        End If

        ' *****************************
        ' - Write cache files
        ' *****************************

        ' Write patch manifest cache file
        Manifest.WriteCache(CallStack, Manifest.PATCH_MANIFEST)

        ' Write history manifest cache file
        Manifest.WriteCache(CallStack, Manifest.HISTORY_MANIFEST)

        ' Write exception manifest cache file
        Manifest.WriteCache(CallStack, Manifest.EXCEPTION_MANIFEST)

        ' *****************************
        ' - Write execution state file. [SD EXECUTION ONLY]
        ' *****************************

        ' Check execution level
        If Globals.SDBasedMode Then

            ' Write debug
            Logger.WriteDebug(CallStack, "Open file: " + Globals.WinOfflineTemp + "\" + Globals.ProcessFriendlyName + ".state")

            ' Try to write execution state file
            Try

                ' Write debug
                Logger.WriteDebug(CallStack, "Write: StageII completed")

                ' Write execution shate marker to file
                StateStreamWriter = New System.IO.StreamWriter(Globals.WinOfflineTemp + "\" + Globals.ProcessFriendlyName + ".state", False)
                StateStreamWriter.WriteLine("StageII completed.")

                ' Write debug
                Logger.WriteDebug(CallStack, "Close file: " + Globals.WinOfflineTemp + "\" + Globals.ProcessFriendlyName + ".state")

                ' Close execution state stream writer
                StateStreamWriter.Close()

            Catch ex As Exception

                ' Write debug
                Logger.WriteDebug(CallStack, "Error: Exception caught writing the state file.")
                Logger.WriteDebug(ex.Message)
                Logger.WriteDebug(ex.StackTrace)

                ' Create exception
                Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

                ' Return
                Return 6

            End Try

        End If

        ' Check if a patch error was reported
        If PatchErrorReported = True Then

            ' Write debug
            Logger.WriteDebug(CallStack, "StageII reports a patching error.")

            ' Return
            Return 7

        Else

            ' Set flag
            Globals.StageIICompleted = True

            ' Return
            Return 0

        End If

    End Function

End Class