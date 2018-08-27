Partial Public Class WinOffline

    Public Shared Function PatchOperations(ByVal CallStack As String) As Integer

        ' Local variables
        Dim Runlevel As Integer = 0

        ' Update call stack
        CallStack += "PatchOperations|"

        ' Write debug
        Logger.SetCurrentTask("Executing..")

        ' Check for removal switch
        If Globals.RemovePatchSwitch Then

            ' *****************************
            ' - Execute removal operations.
            ' *****************************

            ' Encapsulate removal operations
            Try

                ' Iterate the removal list
                For i As Integer = 0 To Manifest.RemovalManifestCount - 1

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Read removal manifest index: [" + i.ToString() + "]")
                    Logger.WriteDebug(CallStack, "Remove patch: " + Manifest.GetRemovalFromManifest(i).RemovalItem)

                    ' Check for simulation
                    If Globals.SimulatePatchSwitch Then

                        If Globals.SimulatePatchErrorSwitch Then

                            ' Write debug
                            Logger.WriteDebug(CallStack, "Switch: Simulate patching error.")

                            ' Update removal action
                            Manifest.GetRemovalFromManifest(i).RemovalAction = RemovalVector.REMOVAL_FAIL

                            ' Update reason comment
                            Manifest.GetRemovalFromManifest(i).CommentString = "Reason: This is a patch error simulation."

                        Else

                            ' Write debug
                            Logger.WriteDebug(CallStack, "Switch: Simulate removing patch.")

                            ' Update removal action
                            Manifest.GetRemovalFromManifest(i).RemovalAction = RemovalVector.SKIPPED

                            ' Update reason comment
                            Manifest.GetRemovalFromManifest(i).CommentString = "Reason: This is a simulation."

                        End If

                    Else

                        ' Remove patch
                        Runlevel = RemovePatch(CallStack, Manifest.GetRemovalFromManifest(i))

                        ' Check runlevel
                        If Runlevel = 1 Then

                            ' Write debug
                            Logger.WriteDebug(CallStack, "Result: SKIPPED.")

                            ' Update removal action
                            Manifest.GetRemovalFromManifest(i).RemovalAction = RemovalVector.SKIPPED

                        ElseIf Runlevel <> 0 Then

                            ' Write debug
                            Logger.WriteDebug(CallStack, "Result: FAILED TO REMOVE.")

                            ' Update removal action
                            Manifest.GetRemovalFromManifest(i).RemovalAction = RemovalVector.REMOVAL_FAIL

                        Else

                            ' Write debug
                            Logger.WriteDebug(CallStack, "Result: REMOVED.")

                            ' Update removal action
                            Manifest.GetRemovalFromManifest(i).RemovalAction = RemovalVector.REMOVAL_OK

                            ' Update patch history file
                            RemoveFromHistory(CallStack, Manifest.GetRemovalFromManifest(i))

                        End If

                    End If

                Next

            Catch ex As Exception

                ' Write debug
                Logger.WriteDebug(CallStack, "Error: Exception caught during removal operations.")
                Logger.WriteDebug(ex.Message)
                Logger.WriteDebug(ex.StackTrace)

                ' Create exception
                Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

                ' Return
                Return 1

            End Try

        Else

            ' *****************************
            ' - Execute apply operations.
            ' *****************************

            ' Encapsulate apply operations
            Try

                ' Iterate the patch manifest
                For i As Integer = 0 To Manifest.PatchManifestCount - 1

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Read patch manifest index: [" + i.ToString() + "]")
                    Logger.WriteDebug(CallStack, "Patch file: " + Manifest.GetPatchFromManifest(i).PatchFile.GetFileName)

                    ' Check for simulation
                    If Globals.SimulatePatchSwitch Then

                        ' Check for patch error simulation
                        If Globals.SimulatePatchErrorSwitch Then

                            ' Write debug
                            Logger.WriteDebug(CallStack, "Switch: Simulate patching error.")

                            ' Update patch action
                            Manifest.GetPatchFromManifest(i).PatchAction = PatchVector.APPLY_FAIL

                            ' Update reason comment
                            Manifest.GetPatchFromManifest(i).CommentString = "Reason: This is a patch error simulation."

                        Else

                            ' Write debug
                            Logger.WriteDebug(CallStack, "Switch: Simulate applying patch.")

                            ' Update patch action
                            Manifest.GetPatchFromManifest(i).PatchAction = PatchVector.SKIPPED

                            ' Update reason comment
                            Manifest.GetPatchFromManifest(i).CommentString = "Reason: This is a simulation."

                        End If

                    ElseIf Not Manifest.GetPatchFromManifest(i).GetInstruction("VERSIONCHECK").Equals("") AndAlso
                        Not Manifest.GetPatchFromManifest(i).GetInstruction("VERSIONCHECK").Equals(Globals.ITCMComstoreVersion) Then

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Result: SKIPPED.")

                        ' Update patch action
                        Manifest.GetPatchFromManifest(i).PatchAction = PatchVector.SKIPPED

                        ' Update reason comment
                        Manifest.GetPatchFromManifest(i).CommentString = "Reason: Agent version does not match JCL specification.NEWLINE"
                        Manifest.GetPatchFromManifest(i).CommentString += "- Agent version (registry): " + Globals.ITCMVersion + "NEWLINE"
                        Manifest.GetPatchFromManifest(i).CommentString += "- Agent version (comstore): " + Globals.ITCMComstoreVersion + "  **USED FOR VERSION CHECK**" + "NEWLINE"
                        Manifest.GetPatchFromManifest(i).CommentString += "- JCL ""VERSIONCHECK:"" specification: " +
                            Manifest.GetPatchFromManifest(i).GetInstruction("VERSIONCHECK").ToString

                    Else

                        ' Apply patch
                        Runlevel = ApplyPatch(CallStack, Manifest.GetPatchFromManifest(i))

                        ' Check runlevel
                        If Runlevel = 1 Then

                            ' Write debug
                            Logger.WriteDebug(CallStack, "Result: SKIPPED.")

                            ' Update patch action
                            Manifest.GetPatchFromManifest(i).PatchAction = PatchVector.SKIPPED

                        ElseIf Runlevel = 2 Then

                            ' Write debug
                            Logger.WriteDebug(CallStack, "Result: NOT APPLICABLE.")

                            ' Update patch action
                            Manifest.GetPatchFromManifest(i).PatchAction = PatchVector.NOT_APPLICABLE

                        ElseIf Runlevel = 3 Then

                            ' Write debug
                            Logger.WriteDebug(CallStack, "Result: ALREADY APPLIED.")

                            ' Update patch action
                            Manifest.GetPatchFromManifest(i).PatchAction = PatchVector.ALREADY_APPLIED

                        ElseIf Runlevel <> 0 Then

                            ' Check for file replacements
                            If Manifest.GetPatchFromManifest(i).SourceReplaceList.Count > 0 Then

                                ' Write debug
                                Logger.WriteDebug(CallStack, "Result: FAILED TO APPLY.")

                                ' Update patch action
                                Manifest.GetPatchFromManifest(i).PatchAction = PatchVector.APPLY_FAIL

                            Else

                                ' Write debug
                                Logger.WriteDebug(CallStack, "Result: FAILED TO EXECUTE.")

                                ' Update patch action
                                Manifest.GetPatchFromManifest(i).PatchAction = PatchVector.EXECUTE_FAIL

                            End If

                        Else

                            ' Check for file replacements
                            If Manifest.GetPatchFromManifest(i).SourceReplaceList.Count > 0 And
                                Manifest.GetPatchFromManifest(i).WereAllFileReplacementsSkipped Then

                                ' Write debug
                                Logger.WriteDebug(CallStack, "Result: SKIPPED.")

                                ' Update patch action
                                Manifest.GetPatchFromManifest(i).PatchAction = PatchVector.SKIPPED

                            ElseIf Manifest.GetPatchFromManifest(i).SourceReplaceList.Count > 0 Then

                                ' Write debug
                                Logger.WriteDebug(CallStack, "Result: APPLIED.")

                                ' Update patch action
                                Manifest.GetPatchFromManifest(i).PatchAction = PatchVector.APPLY_OK

                                ' Update patch history file
                                AddtoHistory(CallStack, Manifest.GetPatchFromManifest(i))

                            Else

                                ' Write debug
                                Logger.WriteDebug(CallStack, "Result: EXECUTED.")

                                ' Update patch action
                                Manifest.GetPatchFromManifest(i).PatchAction = PatchVector.EXECUTE_OK

                            End If

                        End If

                    End If

                Next

            Catch ex As Exception

                ' Write debug
                Logger.WriteDebug(CallStack, "Error: Exception caught during apply operations.")
                Logger.WriteDebug(ex.Message)
                Logger.WriteDebug(ex.StackTrace)

                ' Create exception
                Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

                ' Return
                Return 2

            End Try

        End If

        ' *****************************
        ' - Write cache files.
        ' *****************************

        Try

            ' Write patch manifest cache
            Manifest.WriteCache(CallStack, Manifest.PATCH_MANIFEST)

            ' Write removal manifest cache
            Manifest.WriteCache(CallStack, Manifest.REMOVAL_MANIFEST)

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Exception caught writing cache file.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

            ' Create exception
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

        End Try

        ' Check for recycle CAF service only mode
        If Manifest.PatchManifestCount = 0 And Manifest.RemovalManifestCount = 0 Then

            ' Write debug
            Logger.WriteDebug(CallStack, "No patch operations to perform.")

        End If

        ' Return
        Return 0

    End Function

    Public Shared Sub ExecutePreCmd(ByVal CallStack As String, ByRef pVector As PatchVector)

        ' Local variables
        Dim ExecutionString As String
        Dim ProcessStartInfo As ProcessStartInfo
        Dim RunningProcess As Process
        Dim ConsoleOutput As String
        Dim StandardOutput As String
        Dim RemainingOutput As String

        ' *****************************
        ' - Run PRESYSCMD sripts.
        ' *****************************

        ' Perform PRESYSCMD instructions
        For Each strLine As String In pVector.GetPreCommandList

            ' Build execution string
            ExecutionString = pVector.PatchFile.GetFilePath + "\" + strLine

            ' Write debug
            Logger.WriteDebug(CallStack, "Execute pre-script: " + ExecutionString)

            ' Create detached process
            ProcessStartInfo = New ProcessStartInfo(ExecutionString)
            ProcessStartInfo.WorkingDirectory = pVector.PatchFile.GetFilePath
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

            ' Store return code
            pVector.PreCmdReturnCodes.Add(RunningProcess.ExitCode.ToString)

            ' Close detached process
            RunningProcess.Close()

        Next

    End Sub

    Public Shared Sub ExecutePostCmd(ByVal CallStack As String, ByRef pVector As PatchVector)

        ' Local variables
        Dim ExecutionString As String
        Dim ProcessStartInfo As ProcessStartInfo
        Dim RunningProcess As Process
        Dim ConsoleOutput As String
        Dim StandardOutput As String
        Dim RemainingOutput As String

        ' *****************************
        ' - Run SYSCMD sripts.
        ' *****************************

        ' Perform SYSCMD instructions
        For Each strLine As String In pVector.GetSysCommandList

            ' Build execution string
            ExecutionString = pVector.PatchFile.GetFilePath + "\" + strLine

            ' Write debug
            Logger.WriteDebug(CallStack, "Execute script: " + ExecutionString)

            ' Create detached process
            ProcessStartInfo = New ProcessStartInfo(ExecutionString)
            ProcessStartInfo.WorkingDirectory = pVector.PatchFile.GetFilePath
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

            ' Store return code
            pVector.SysCmdReturnCodes.Add(RunningProcess.ExitCode.ToString)

            ' Close detached process
            RunningProcess.Close()

        Next

    End Sub

    Public Shared Function ApplyPatch(ByVal CallStack As String, ByRef pVector As PatchVector) As Integer

        ' Local variables
        Dim AlreadyAppliedList As New List(Of Boolean)
        Dim ReplacedIncrement As Integer = 0
        Dim DestinationFileName As String = ""
        Dim ReplacedFolder As String = ""
        Dim NewFileList As New ArrayList
        Dim Subfolder As String = ""
        Dim RebootFileName As String = ""
        Dim RebootIncrement As Integer = 0

        ' Update call stack
        CallStack += "ApplyPatch|"

        ' *****************************
        ' - Validate patch component is installed.
        ' *****************************

        ' Check the patch component
        If (pVector.IsClientAuto AndAlso Globals.DSMFolder Is Nothing) Or
            (pVector.IsSharedComponent AndAlso Globals.SharedCompFolder Is Nothing) Or
            (pVector.IsCAM AndAlso Globals.CAMFolder Is Nothing) Or
            (pVector.IsSSA AndAlso Globals.SSAFolder Is Nothing) Or
            (pVector.IsDataTransport AndAlso Globals.DTSFolder Is Nothing) Or
            (pVector.IsExplorerGUI AndAlso Globals.EGCFolder Is Nothing) Then

            ' Write debug
            Logger.WriteDebug(CallStack, "Patch skipped: " + pVector.PatchFile.GetFriendlyName)
            Logger.WriteDebug(CallStack, "Reason: Product component (" + pVector.GetInstruction("PRODUCT") + ") is not installed.")

            ' Update reason comment
            pVector.CommentString = "Reason: Product component (" + pVector.GetInstruction("PRODUCT") + ") is not installed."

            ' Return
            Return 1

        End If

        ' *****************************
        ' - Check if patch if already applied (file replacement check).
        ' *****************************

        ' Check for file replacement operations
        If pVector.SourceReplaceList.Count > 0 Then

            ' Write debug
            Logger.WriteDebug(CallStack, "Check if patch is already applied..")

            ' Iterate file replacements
            For x As Integer = 0 To pVector.SourceReplaceList.Count - 1

                ' Write debug
                Logger.WriteDebug(CallStack, "Compare: " + pVector.SourceReplaceList.Item(x).ToString)
                Logger.WriteDebug(CallStack, "Compare: " + pVector.DestReplaceList.Item(x).ToString)

                ' Check for unequal replacement file
                If Not Utility.IsFileEqual(pVector.SourceReplaceList.Item(x).ToString, pVector.DestReplaceList.Item(x).ToString) Then

                    ' Verify unequal file is not a new file vs skippable file
                    If Not System.IO.File.Exists(pVector.DestReplaceList.Item(x).ToString) AndAlso
                        pVector.SkipIfNotFoundList.Contains(FileVector.GetShortName(pVector.DestReplaceList.Item(x).ToString.ToLower)) Then

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Result: Mismatch [Skip file]")

                        ' Skipped file -- no applicability information

                    ElseIf Not System.IO.File.Exists(pVector.DestReplaceList.Item(x).ToString) AndAlso
                        Not pVector.SkipIfNotFoundList.Contains(FileVector.GetShortName(pVector.DestReplaceList.Item(x).ToString.ToLower)) Then

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Result: Mismatch [New file]")

                        ' New file -- not applied
                        AlreadyAppliedList.Add(False)

                    Else

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Result: Mismatch")

                        ' Complete mismatch -- not applied
                        AlreadyAppliedList.Add(False)

                    End If

                Else

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Result: Match")

                    ' File match -- already applied
                    AlreadyAppliedList.Add(True)

                End If

            Next

            ' Check already applied list
            If AlreadyAppliedList.Count = 0 Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Patch is not applicable.")

                ' Update reason comment
                pVector.CommentString = "Reason: All patch files will be skipped."

                ' Return
                Return 2

            ElseIf Not AlreadyAppliedList.Contains(False) Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Patch is already applied.")

                ' Update reason comment
                pVector.CommentString = "Reason: All patch files match with destination files."

                ' Return
                Return 3

            Else

                ' Write debug
                Logger.WriteDebug(CallStack, "Patch is not already applied.")

            End If

        Else

            ' Write debug
            Logger.WriteDebug(CallStack, "Patch is not already applied.")

        End If

        ' *****************************
        ' - Run PRESYSCMD sripts.
        ' *****************************

        ' Run the PRESYSCMDs
        ExecutePreCmd(CallStack, pVector)

        ' *****************************
        ' - Prepare REPLACED subfolder.
        ' *****************************

        ' Check for file replacement operations
        If pVector.SourceReplaceList.Count > 0 Then

            ' Check if REPLACED subfolder exists
            If System.IO.Directory.Exists(pVector.ReplaceFolder + ".OLD") Then

                ' Loop dangerously for first available subfolder
                While True

                    ' Check if incremented folder name exists
                    If System.IO.Directory.Exists(pVector.ReplaceFolder + "-" + ReplacedIncrement.ToString + ".OLD") Then

                        ' Increment subfolder
                        ReplacedIncrement += 1

                    Else

                        ' Create the subfolder
                        Try

                            ' Write debug
                            Logger.WriteDebug(CallStack, "Create REPLACED subfolder: " + pVector.ReplaceFolder + "-" + ReplacedIncrement.ToString + ".OLD")

                            ' Set replaced folder
                            ReplacedFolder = pVector.ReplaceFolder + "-" + ReplacedIncrement.ToString + ".OLD"

                            ' Create the REPLACED subfolder
                            System.IO.Directory.CreateDirectory(ReplacedFolder)

                            ' Break loop
                            Exit While

                        Catch ex As Exception

                            ' Write debug
                            Logger.WriteDebug(CallStack, "Error: Failed to create REPLACED subfolder.")
                            Logger.WriteDebug(ex.Message)
                            Logger.WriteDebug(ex.StackTrace)

                            ' Create exception
                            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

                            ' Write debug
                            Logger.WriteDebug(CallStack, "Patch skipped: " + pVector.PatchFile.GetFriendlyName)
                            Logger.WriteDebug(CallStack, "Reason: Failed to create REPLACED subfolder.")

                            ' Update reason comment
                            pVector.CommentString = "Reason: Failed to create REPLACED subfolder."

                            ' Return
                            Return 4

                        End Try

                    End If

                End While

            Else

                ' Create the subfolder
                Try

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Create REPLACED subfolder: " + pVector.ReplaceFolder + ".OLD")

                    ' Set replaced folder
                    ReplacedFolder = pVector.ReplaceFolder + ".OLD"

                    ' Create the REPLACED subfolder
                    System.IO.Directory.CreateDirectory(ReplacedFolder)

                Catch ex As Exception

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Error: Failed to create REPLACED subfolder.")
                    Logger.WriteDebug(ex.Message)
                    Logger.WriteDebug(ex.StackTrace)

                    ' Create exception
                    Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Patch skipped: " + pVector.PatchFile.GetFriendlyName)
                    Logger.WriteDebug(CallStack, "Reason: Failed to create REPLACED subfolder.")

                    ' Update reason comment
                    pVector.CommentString = "Reason: Failed to create REPLACED subfolder."

                    ' Return
                    Return 5

                End Try

            End If

        End If

        ' *****************************
        ' - Save original files to REPLACED folder.
        ' *****************************

        ' Iterate file replacements -- Save original files to REPLACED folder
        For x As Integer = 0 To pVector.DestReplaceList.Count - 1

            ' Check if replacement file exists, or if the file is new
            If System.IO.File.Exists(pVector.DestReplaceList.Item(x)) Then

                ' Calculate destination filename
                DestinationFileName = ReplacedFolder + "\" + pVector.ReplaceSubFolder.Item(x) + "\" +
                        FileVector.GetShortName(pVector.DestReplaceList.Item(x))

                ' Check for double backslash
                DestinationFileName = DestinationFileName.Replace("\\", "\")

                ' Ensure destination subfolder is available
                If Not System.IO.Directory.Exists(FileVector.GetFilePath(DestinationFileName)) Then

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Create subfolder: " + FileVector.GetFilePath(DestinationFileName))

                    ' Create the subfolder
                    System.IO.Directory.CreateDirectory(FileVector.GetFilePath(DestinationFileName))

                End If

                ' Attempt save operation
                Try

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Save original file: " + pVector.DestReplaceList.Item(x))
                    Logger.WriteDebug(CallStack, "Save to: " + DestinationFileName)

                    ' Copy file to appropriate REPLACED subfolder
                    System.IO.File.Copy(pVector.DestReplaceList.Item(x), DestinationFileName)

                Catch ex As Exception

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Error: Failed to save original file(s).")
                    Logger.WriteDebug(ex.Message)
                    Logger.WriteDebug(ex.StackTrace)

                    ' Create exception
                    Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Delete REPLACED folder: " + pVector.ReplaceFolder)

                    ' Delete REPLACED subfolder
                    System.IO.Directory.Delete(ReplacedFolder, True)

                    ' Update file replacement result
                    pVector.FileReplaceResult.Item(x) = PatchVector.FILE_FAILED

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Patch failed: " + pVector.PatchFile.GetFriendlyName)
                    Logger.WriteDebug(CallStack, "Reason: Failed to save original file(s).")

                    ' Update reason comment
                    pVector.CommentString = "Reason: Failed to save original file(s)."

                    ' Return
                    Return 6

                End Try

            Else

                ' Add to new file list (We may need to skip replacing these new files later)
                NewFileList.Add(FileVector.GetShortName(pVector.DestReplaceList.Item(x).ToString.ToLower))

            End If

        Next

        ' *****************************
        ' - Remove original files before replacement.
        ' *****************************

        ' Iterate file replacements -- Remove originals ahead of replacements
        For x As Integer = 0 To pVector.DestReplaceList.Count - 1

            ' Check if replacement file exists, or skip if the file is new
            If System.IO.File.Exists(pVector.DestReplaceList.Item(x)) Then

                ' Attempt to delete original file (or mark it for deletion on next reboot)
                Try

                    ' Check if the file is in-use
                    If Utility.IsFileOpen(pVector.DestReplaceList.Item(x)) Then

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Delete original file (on next reboot): " + pVector.DestReplaceList.Item(x))

                        ' Set reboot filename
                        RebootFileName = pVector.DestReplaceList.Item(x) + ".delete_on_reboot"

                        ' Reset counter
                        RebootIncrement = 0

                        ' Loop dangerously for available filename
                        While True

                            ' Check if filename exists
                            If System.IO.File.Exists(RebootFileName) Then

                                ' Set new file name
                                RebootFileName = pVector.DestReplaceList.Item(x) + ".delete_on_reboot" + RebootIncrement.ToString

                                ' Increment counter
                                RebootIncrement += 1

                            Else

                                ' Stop condition
                                Exit While

                            End If

                        End While

                        ' Rename original file
                        System.IO.File.Move(pVector.DestReplaceList.Item(x), RebootFileName)

                        ' Schedule the file for removal on reboot
                        WindowsAPI.MoveFileEx(RebootFileName, Nothing, WindowsAPI.MoveFileFlags.DelayUntilReboot)

                        ' Update file replacement result
                        pVector.FileReplaceResult.Item(x) = PatchVector.FILE_REBOOT_REQUIRED

                    Else

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Delete original file: " + pVector.DestReplaceList.Item(x))

                        ' Delete original file
                        System.IO.File.Delete(pVector.DestReplaceList.Item(x))

                    End If

                Catch ex As Exception

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Error: Failed to delete original file(s), or schedule their removal.")
                    Logger.WriteDebug(ex.Message)
                    Logger.WriteDebug(ex.StackTrace)

                    ' Create exception
                    Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

                    ' Undo changes
                    If x > 0 Then

                        ' Restore original files
                        For y As Integer = 0 To pVector.DestReplaceList.Count - 1

                            ' Calculate original filename for restoration
                            DestinationFileName = ReplacedFolder + "\" + pVector.ReplaceSubFolder.Item(y) + "\" +
                                FileVector.GetShortName(pVector.DestReplaceList.Item(y))

                            ' Check for double backslash
                            DestinationFileName = DestinationFileName.Replace("\\", "\")

                            ' Verify the file exists
                            ' Note: In the case where the patch file is a new file, then nothing was
                            '       backed up to the REPLACED folder, to restore to the original path.
                            If System.IO.File.Exists(DestinationFileName) Then

                                ' Write debug
                                Logger.WriteDebug(CallStack, "Restore original file: " + DestinationFileName)
                                Logger.WriteDebug(CallStack, "To: " + pVector.DestReplaceList.Item(y))

                                ' Restore original file from REPLACED folder
                                System.IO.File.Copy(DestinationFileName, pVector.DestReplaceList.Item(y), True)

                            End If

                        Next

                    End If

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Delete REPLACED folder: " + pVector.ReplaceFolder)

                    ' Delete REPLACED subfolder
                    System.IO.Directory.Delete(ReplacedFolder, True)

                    ' Update file replacement result
                    pVector.FileReplaceResult.Item(x) = PatchVector.FILE_FAILED

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Patch failed: " + pVector.PatchFile.GetFriendlyName)
                    Logger.WriteDebug(CallStack, "Reason: Failed to delete original file(s), or schedule their removal.")

                    ' Update reason comment
                    pVector.CommentString = "Reason: Failed to delete original file(s), or schedule their removal."

                    ' Return
                    Return 7

                End Try

            End If

        Next

        ' *****************************
        ' - Copy in replacement files.
        ' *****************************

        ' Iterate file replacements -- Copy in the replacement files
        For x As Integer = 0 To pVector.SourceReplaceList.Count - 1

            ' Check if file replacement should be skipped
            If pVector.SkipIfNotFoundList.Contains(FileVector.GetShortName(pVector.DestReplaceList.Item(x).ToString.ToLower)) And
                NewFileList.Contains(FileVector.GetShortName(pVector.DestReplaceList.Item(x).ToString.ToLower)) Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Skip replacement file: " + pVector.SourceReplaceList.Item(x))

                ' Mark skipped
                pVector.FileReplaceResult.Item(x) = PatchVector.SKIPPED

                ' Skip to end
                Continue For

            End If

            ' Write debug
            Logger.WriteDebug(CallStack, "Copy replacement file: " + pVector.SourceReplaceList.Item(x))
            Logger.WriteDebug(CallStack, "To: " + pVector.DestReplaceList.Item(x))

            Try

                ' Copy patch files to directory
                System.IO.File.Copy(pVector.SourceReplaceList.Item(x), pVector.DestReplaceList.Item(x), True)

                ' Update file replacement result (mark successful, only if orignal file was not in use)
                If Not pVector.FileReplaceResult.Item(x) = PatchVector.FILE_REBOOT_REQUIRED Then

                    ' Mark successful
                    pVector.FileReplaceResult.Item(x) = PatchVector.FILE_OK

                End If

            Catch ex As Exception

                ' Write debug
                Logger.WriteDebug(CallStack, "Error: Failed to copy replacement file(s) to destination.")
                Logger.WriteDebug(ex.Message)
                Logger.WriteDebug(ex.StackTrace)

                ' Create exception
                Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

                ' Undo changes
                If x > 0 Then

                    ' Restore original files
                    For y As Integer = 0 To pVector.DestReplaceList.Count - 1

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Delete replacement file: " + pVector.DestReplaceList.Item(y))

                        ' Ensure any replacement file is first removed
                        System.IO.File.Delete(pVector.DestReplaceList.Item(y))

                        ' Calculate original filename for restoration
                        DestinationFileName = ReplacedFolder + "\" + pVector.ReplaceSubFolder.Item(y) + "\" +
                            FileVector.GetShortName(pVector.DestReplaceList.Item(y))

                        ' Check for double backslash
                        DestinationFileName = DestinationFileName.Replace("\\", "\")

                        ' Verify the file exists
                        ' Note: In the case where the patch file is a new file, then nothing was
                        '       backed up to the REPLACED folder, to restore to the original path.
                        If System.IO.File.Exists(DestinationFileName) Then

                            ' Write debug
                            Logger.WriteDebug(CallStack, "Restore original file: " + DestinationFileName)
                            Logger.WriteDebug(CallStack, "To: " + pVector.DestReplaceList.Item(y))

                            ' Restore original file from REPLACED folder
                            System.IO.File.Copy(DestinationFileName, pVector.DestReplaceList.Item(y), True)

                        End If

                    Next

                End If

                ' Write debug
                Logger.WriteDebug(CallStack, "Delete REPLACED subfolder: " + pVector.ReplaceFolder)

                ' Delete the REPLACED subfolder
                System.IO.Directory.Delete(ReplacedFolder, True)

                ' Update file replacement result
                pVector.FileReplaceResult.Item(x) = PatchVector.FILE_FAILED

                ' Write debug
                Logger.WriteDebug(CallStack, "Patch failed: " + pVector.PatchFile.GetFriendlyName)
                Logger.WriteDebug(CallStack, "Reason: Failed to copy replacement file(s) to destination.")

                ' Update reason comment
                pVector.CommentString = "Reason: Failed to copy replacement file(s) to destination."

                ' Return
                Return 8

            End Try

        Next

        ' *****************************
        ' - Special case: All file replacements skipped.
        ' *****************************

        ' Check if all file replacements were skipped
        If pVector.SourceReplaceList.Count > 0 And pVector.WereAllFileReplacementsSkipped Then

            ' Update reason comment
            pVector.CommentString = "Reason: All patch file replacements skipped."

            ' Write debug
            Logger.WriteDebug(CallStack, "Delete REPLACED subfolder: " + pVector.ReplaceFolder)

            ' Delete the REPLACED subfolder
            System.IO.Directory.Delete(ReplacedFolder, True)

        End If

        ' *****************************
        ' - Run SYSCMD sripts.
        ' *****************************

        ' Run the SYSCMDs
        ExecutePostCmd(CallStack, pVector)

        ' Return
        Return 0

    End Function

    Public Shared Function RemovePatch(ByVal CallStack As String, ByRef rVector As RemovalVector) As Integer

        ' Local variables
        Dim hVector As HistoryVector = Nothing
        Dim HistoryVector As HistoryVector = Nothing
        Dim RemovalMatchFound As Boolean = False
        Dim OriginalFilesFound As Boolean = False
        Dim ReplacedFolder As String = ""
        Dim ReplacedBaseFolder As String = ""
        Dim LatestReplacedFolder As String = ""
        Dim ReplacedIncrement As Integer = 0
        Dim SubFolder As String = ""
        Dim SubFolderList As New ArrayList
        Dim NewFileList As New ArrayList
        Dim RestoreList As New ArrayList
        Dim DestinationFolder As String
        Dim DestinationIncrement As Integer = 0
        Dim DestinationFileName As String
        Dim RebootFileName As String = ""
        Dim RebootIncrement As Integer = 0

        ' Update call stack
        CallStack += "RemovePatch|"

        ' *****************************
        ' - Search patch history.
        ' *****************************

        ' Iterate each history vector
        For i As Integer = 0 To Manifest.HistoryManifestCount - 1

            ' Read a history vector
            HistoryVector = Manifest.GetHistoryFromManifest(i)

            ' Check for a match
            If HistoryVector.GetPatchName.ToLower.Equals(rVector.RemovalItem.ToLower) Then

                ' Set flag
                RemovalMatchFound = True

                ' Update latest match
                hVector = HistoryVector

                ' Don't exit the loop premature.
                ' Note: The idea is to remove the "latest" application of the patch. It's possible it
                '       could have been applied more than once, and it's also possible the list of
                '       installed files may not be the same (i.e. the details of the patch have changed).
                '       Since we want to remove history from the bottom, up, we should pop the latest
                '       hVector/history file record that matches the title of the patch.
                ' Exit For

            End If

        Next

        ' Ensure a match was found
        If RemovalMatchFound = False Then

            ' Write debug
            Logger.WriteDebug(CallStack, "Removal skipped: " + rVector.RemovalItem)
            Logger.WriteDebug(CallStack, "Reason: Patch is not found in history file(s).")

            ' Update reason comment
            rVector.CommentString = "Reason: Patch is not found in history file(s)."

            ' Return -- Matching history record not found
            Return 1

        End If

        ' *****************************
        ' - Find latest REPLACED subfolder.
        ' *****************************

        ' Build REPLACED folder by component
        If hVector.GetProductComponent.Equals(IT_CLIENT_MANAGER) Then
            ReplacedBaseFolder = Globals.DSMFolder + "REPLACED"
            ReplacedFolder = Globals.DSMFolder + "REPLACED\" + hVector.GetPatchName
            rVector.HistoryFile = Globals.DSMFolder + Globals.HostName + ".his"
        ElseIf hVector.GetProductComponent.Equals(SHARED_COMPONENTS) Then
            ReplacedBaseFolder = Globals.SharedCompFolder + "REPLACED"
            ReplacedFolder = Globals.SharedCompFolder + "REPLACED\" + hVector.GetPatchName
            rVector.HistoryFile = Globals.SharedCompFolder + Globals.HostName + ".his"
        ElseIf hVector.GetProductComponent.Equals(CA_MESSAGE_QUEUING) Then
            ReplacedBaseFolder = Globals.CAMFolder + "\REPLACED"
            ReplacedFolder = Globals.CAMFolder + "\REPLACED\" + hVector.GetPatchName
            rVector.HistoryFile = Globals.CAMFolder + "\" + Globals.HostName + ".his"
        ElseIf hVector.GetProductComponent.Equals(SECURE_SOCKET_ADAPATER) Then
            ReplacedBaseFolder = Globals.SSAFolder + "REPLACED"
            ReplacedFolder = Globals.SSAFolder + "REPLACED\" + hVector.GetPatchName
            rVector.HistoryFile = Globals.SSAFolder + Globals.HostName + ".his"
        ElseIf hVector.GetProductComponent.Equals(DATA_TRANSPORT) Then
            ReplacedBaseFolder = Globals.DTSFolder + "\REPLACED\"
            ReplacedFolder = Globals.DTSFolder + "\REPLACED\" + hVector.GetPatchName
            rVector.HistoryFile = Globals.DTSFolder + "\" + Globals.HostName + ".his"
        ElseIf hVector.GetProductComponent.Equals(EXPLORER_GUI) Then
            ReplacedBaseFolder = Globals.EGCFolder + "REPLACED"
            ReplacedFolder = Globals.EGCFolder + "REPLACED\" + hVector.GetPatchName
            rVector.HistoryFile = Globals.EGCFolder + Globals.HostName + ".his"
        End If

        ' Set base case for latest folder
        LatestReplacedFolder = ReplacedFolder + ".OLD"

        ' Verify base folder exists
        If System.IO.Directory.Exists(ReplacedFolder + ".OLD") Then

            ' Loop dangerously for latest increment REPLACED folder
            While True

                ' Check if incremental folder exists
                If System.IO.Directory.Exists(ReplacedFolder + "-" + ReplacedIncrement.ToString + ".OLD") Then

                    ' Assign latest folder
                    LatestReplacedFolder = ReplacedFolder + "-" + ReplacedIncrement.ToString + ".OLD"

                    ' Increase increment
                    ReplacedIncrement += 1

                Else

                    ' Stop condition
                    Exit While

                End If

            End While

            ' Assign latest REPLACED folder
            ReplacedFolder = LatestReplacedFolder

        Else

            ' Write debug
            Logger.WriteDebug(CallStack, "Removal failed: " + rVector.RemovalItem)
            Logger.WriteDebug(CallStack, "Reason: Original files unavailable for restoration.")

            ' Update reason comment
            rVector.CommentString = "Reason: Original files unavailable for restoration."

            ' Return -- Original files unavailable
            Return 2

        End If

        ' *****************************
        ' - Generate NEW FILE list for removal.
        ' *****************************

        ' Note: Per testing with ApplyPTF, any "INSTALLEDFILE=" specified from the history
        '       file, that does not have a matching original file within its REPLACED folder,
        '       is to be considered a NEW FILE introduced by the application of the patch.
        '       Therefore, on removal of said patch, NEW FILES will simply be deleted, but
        '       also saved in the BACKOUT subfolder.
        '
        '       Maintaining a NewFile list is not really necessary, as lower sections of
        '       code automatically save them to the BACKOUT subsolder and delete them
        '       where they stand. But I'm leaving this code here for now, in case this
        '       functionality is changed in the future.

        ' Iterate installed files
        For Each InstalledFile As String In hVector.GetInstalledFiles

            ' Verify original file, by component subfolder in REPLACED
            If hVector.GetProductComponent.Equals(IT_CLIENT_MANAGER) Then

                ' Build subfolder
                SubFolder = InstalledFile.ToLower.Replace(Globals.DSMFolder.ToLower, "")
                SubFolder = SubFolder.Replace("\\", "\")

            ElseIf hVector.GetProductComponent.Equals(SHARED_COMPONENTS) Then

                ' Build subfolder
                SubFolder = InstalledFile.ToLower.Replace(Globals.SharedCompFolder.ToLower, "")
                SubFolder = SubFolder.Replace("\\", "\")

            ElseIf hVector.GetProductComponent.Equals(CA_MESSAGE_QUEUING) Then

                ' Build subfolder
                SubFolder = InstalledFile.ToLower.Replace(Globals.CAMFolder.ToLower, "")
                SubFolder = SubFolder.Replace("\\", "\")

            ElseIf hVector.GetProductComponent.Equals(SECURE_SOCKET_ADAPATER) Then

                ' Build subfolder
                SubFolder = InstalledFile.ToLower.Replace(Globals.SSAFolder.ToLower, "")
                SubFolder = SubFolder.Replace("\\", "\")

            ElseIf hVector.GetProductComponent.Equals(DATA_TRANSPORT) Then

                ' Build subfolder
                SubFolder = InstalledFile.ToLower.Replace(Globals.DTSFolder.ToLower, "")
                SubFolder = SubFolder.Replace("\\", "\")

            ElseIf hVector.GetProductComponent.Equals(EXPLORER_GUI) Then

                ' Build subfolder
                SubFolder = InstalledFile.ToLower.Replace(Globals.EGCFolder.ToLower, "")
                SubFolder = SubFolder.Replace("\\", "\")

            End If

            ' Check for original file
            If System.IO.File.Exists(ReplacedFolder + "\" + SubFolder) Then

                ' Set flag
                OriginalFilesFound = True

                ' Add to restoration list
                RestoreList.Add(ReplacedFolder + "\" + SubFolder)

                ' Add subfolder to list (so we don't have to recalculate this later)
                SubFolderList.Add(SubFolder)

            Else

                ' Add to new file removal list
                NewFileList.Add(InstalledFile)

                ' Add subfolder to list (so we don't have to recalculate this later)
                SubFolderList.Add(SubFolder)

            End If

        Next

        ' *****************************
        ' - Prepare BACKOUT subfolder.
        ' *****************************

        ' Build destination folder
        DestinationFolder = ReplacedBaseFolder + "\BACKOUT.OLD\" + hVector.GetPatchName

        ' Verify if destination already folder exists
        If System.IO.Directory.Exists(DestinationFolder + ".OLD") Then

            ' Loop dangerously for latest increment destination folder
            While True

                ' Check if incremental folder exists
                If System.IO.Directory.Exists(DestinationFolder + "-" + DestinationIncrement.ToString + ".OLD") Then

                    ' Increment subfolder
                    DestinationIncrement += 1

                Else

                    ' Create the subfolder
                    Try

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Create BACKOUT subfolder: " + DestinationFolder + "-" + DestinationIncrement.ToString + ".OLD")

                        ' Set destination folder
                        DestinationFolder = DestinationFolder + "-" + DestinationIncrement.ToString + ".OLD"

                        ' Create the destination subfolder
                        System.IO.Directory.CreateDirectory(DestinationFolder)

                        ' Break loop
                        Exit While

                    Catch ex As Exception

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Removal failed: " + rVector.RemovalItem)
                        Logger.WriteDebug(CallStack, "Error: Failed to create BACKOUT subfolder.")
                        Logger.WriteDebug(ex.Message)
                        Logger.WriteDebug(ex.StackTrace)

                        ' Create exception
                        Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

                        ' Update reason comment
                        rVector.CommentString = "Reason: Failed to create BACKOUT subfolder."

                        ' Return
                        Return 3

                    End Try

                    ' Stop condition
                    Exit While

                End If

            End While

        Else

            ' Create the subfolder
            Try

                ' Write debug
                Logger.WriteDebug(CallStack, "Create BACKOUT subfolder: " + DestinationFolder + ".OLD")

                ' Set destination folder
                DestinationFolder = DestinationFolder + ".OLD"

                ' Create the destination subfolder
                System.IO.Directory.CreateDirectory(DestinationFolder)

            Catch ex As Exception

                ' Write debug
                Logger.WriteDebug(CallStack, "Removal failed: " + rVector.RemovalItem)
                Logger.WriteDebug(CallStack, "Error: Failed to create BACKOUT subfolder.")
                Logger.WriteDebug(ex.Message)
                Logger.WriteDebug(ex.StackTrace)

                ' Create exception
                Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

                ' Update reason comment
                rVector.CommentString = "Reason: Failed to create BACKOUT subfolder."

                ' Return
                Return 4

            End Try

        End If

        ' *****************************
        ' - Save current files to BACKOUT folder.
        ' *****************************

        ' Iterate current files -- Save current files to BACKOUT folder
        For x As Integer = 0 To hVector.GetInstalledFiles.Count - 1

            ' Verify current file exists
            If System.IO.File.Exists(hVector.GetInstalledFiles.Item(x)) Then

                ' Calculate destination filename
                DestinationFileName = DestinationFolder + "\" + SubFolderList.Item(x)

                ' Check for double backslash
                DestinationFileName = DestinationFileName.Replace("\\", "\")

                ' Ensure destination subfolder is available
                If Not System.IO.Directory.Exists(FileVector.GetFilePath(DestinationFileName)) Then

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Create subfolder: " + FileVector.GetFilePath(DestinationFileName))

                    ' Create the subfolder
                    System.IO.Directory.CreateDirectory(FileVector.GetFilePath(DestinationFileName))

                End If

                ' Attempt save operation
                Try

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Save current file: " + hVector.GetInstalledFiles.Item(x))
                    Logger.WriteDebug(CallStack, "Save to: " + DestinationFileName)

                    ' Copy file to appropriate REPLACED subfolder
                    System.IO.File.Copy(hVector.GetInstalledFiles.Item(x), DestinationFileName, True)

                Catch ex As Exception

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Error: Failed to save current file(s).")
                    Logger.WriteDebug(ex.Message)
                    Logger.WriteDebug(ex.StackTrace)

                    ' Create exception
                    Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Delete BACKOUT folder: " + DestinationFolder)

                    ' Create the REPLACED subfolder
                    System.IO.Directory.Delete(DestinationFolder, True)

                    ' Update file replacement result
                    rVector.RemovalFileName.Add(hVector.GetInstalledFiles.Item(x))
                    rVector.FileRemovalResult.Add(RemovalVector.FILE_FAILED)

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Removal failed: " + rVector.RemovalItem)

                    ' Update reason comment
                    rVector.CommentString = "Reason: Failed to save current file(s)."

                    ' Return
                    Return 5

                End Try

            End If

        Next

        ' *****************************
        ' - Remove current files before replacement.
        ' *****************************

        ' Iterate current files -- Remove current files ahead of replacements
        For x As Integer = 0 To hVector.GetInstalledFiles.Count - 1

            ' Verify current file exists
            If System.IO.File.Exists(hVector.GetInstalledFiles.Item(x)) Then

                ' Attempt to delete original file (or mark it for deletion on next reboot)
                Try

                    ' Check if the file is in-use
                    If Utility.IsFileOpen(hVector.GetInstalledFiles.Item(x)) Then

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Delete original file (on next reboot): " + hVector.GetInstalledFiles.Item(x))

                        ' Set reboot filename
                        RebootFileName = hVector.GetInstalledFiles.Item(x) + ".delete_on_reboot"

                        ' Reset counter
                        RebootIncrement = 0

                        ' Loop dangerously for available filename
                        While True

                            ' Check if filename exists
                            If System.IO.File.Exists(RebootFileName) Then

                                ' Set new file name
                                RebootFileName = hVector.GetInstalledFiles.Item(x) + ".delete_on_reboot" + RebootIncrement.ToString

                                ' Increment counter
                                RebootIncrement += 1

                            Else

                                ' Stop condition
                                Exit While

                            End If

                        End While

                        ' Rename original file
                        System.IO.File.Move(hVector.GetInstalledFiles.Item(x), RebootFileName)

                        ' Schedule the file for removal on reboot
                        WindowsAPI.MoveFileEx(RebootFileName, Nothing, WindowsAPI.MoveFileFlags.DelayUntilReboot)

                        ' Update file replacement result
                        rVector.RemovalFileName.Add(hVector.GetInstalledFiles.Item(x))
                        rVector.FileRemovalResult.Add(RemovalVector.FILE_REBOOT_REQUIRED)

                    Else

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Delete current file: " + hVector.GetInstalledFiles.Item(x))

                        ' Delete current file
                        System.IO.File.Delete(hVector.GetInstalledFiles.Item(x))

                        ' Update file replacement result
                        rVector.RemovalFileName.Add(hVector.GetInstalledFiles.Item(x))
                        rVector.FileRemovalResult.Add(RemovalVector.FILE_OK)

                    End If

                Catch ex As Exception

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Error: Failed to delete current file(s), or schedule their removal.")
                    Logger.WriteDebug(ex.Message)
                    Logger.WriteDebug(ex.StackTrace)

                    ' Create exception
                    Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

                    ' Undo changes
                    If x > 0 Then

                        ' Restore original files
                        For y As Integer = 0 To hVector.GetInstalledFiles.Count - 1

                            ' Calculate current filename for restoration
                            DestinationFileName = DestinationFolder + "\" + SubFolderList.Item(y)

                            ' Check for double backslash
                            DestinationFileName = DestinationFileName.Replace("\\", "\")

                            ' Verify the file exists
                            If System.IO.File.Exists(DestinationFileName) Then

                                ' Write debug
                                Logger.WriteDebug(CallStack, "Restore current file: " + DestinationFileName)
                                Logger.WriteDebug(CallStack, "To: " + hVector.GetInstalledFiles.Item(y))

                                ' Restore original file from REPLACED folder
                                System.IO.File.Copy(DestinationFileName, hVector.GetInstalledFiles.Item(y), True)

                            End If

                        Next

                    End If

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Delete BACKOUT folder: " + DestinationFolder)

                    ' Delete the BACKOUT subfolder
                    System.IO.Directory.Delete(DestinationFolder, True)

                    ' Update file replacement result
                    rVector.RemovalFileName.Add(hVector.GetInstalledFiles.Item(x))
                    rVector.FileRemovalResult.Add(RemovalVector.FILE_FAILED)

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Removal failed: " + rVector.RemovalItem)

                    ' Update reason comment
                    rVector.CommentString = "Reason: Failed to delete current file(s), or schedule their removal."

                    ' Return
                    Return 6

                End Try

            End If

        Next

        ' *****************************
        ' - Restore original files.
        ' *****************************

        ' Iterate restoration list -- Restore original files
        For x As Integer = 0 To RestoreList.Count - 1

            ' Write debug
            Logger.WriteDebug(CallStack, "Restore file: " + RestoreList.Item(x))
            Logger.WriteDebug(CallStack, "To: " + hVector.GetInstalledFiles.Item(x))

            Try

                ' Restore file
                System.IO.File.Copy(RestoreList.Item(x), hVector.GetInstalledFiles.Item(x), True)

                ' Update file replacement result (mark successful, only if orignal file was not in use)
                If Not rVector.FileRemovalResult.Item(x) = RemovalVector.FILE_REBOOT_REQUIRED Then

                    ' Mark successful
                    rVector.FileRemovalResult.Item(x) = RemovalVector.FILE_OK

                End If

            Catch ex As Exception

                ' Write debug
                Logger.WriteDebug(CallStack, "Error: Failed to restore original file(s).")
                Logger.WriteDebug(ex.Message)
                Logger.WriteDebug(ex.StackTrace)

                ' Create exception
                Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

                ' Undo changes
                If x > 0 Then

                    ' Restore original files
                    For y As Integer = 0 To hVector.GetInstalledFiles.Count - 1

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Delete restored file: " + hVector.GetInstalledFiles.Item(y))

                        ' Ensure any restored file is first removed
                        System.IO.File.Delete(hVector.GetInstalledFiles.Item(y))

                        ' Calculate current filename for restoration
                        DestinationFileName = DestinationFolder + "\" + SubFolderList.Item(y)

                        ' Check for double backslash
                        DestinationFileName = DestinationFileName.Replace("\\", "\")

                        ' Verify the file exists
                        If System.IO.File.Exists(DestinationFileName) Then

                            ' Write debug
                            Logger.WriteDebug(CallStack, "Restore current file: " + DestinationFileName)
                            Logger.WriteDebug(CallStack, "To: " + hVector.GetInstalledFiles.Item(y))

                            ' Restore original file from REPLACED folder
                            System.IO.File.Copy(DestinationFileName, hVector.GetInstalledFiles.Item(y), True)

                        End If

                    Next

                End If

                ' Write debug
                Logger.WriteDebug(CallStack, "Delete BACKOUT folder: " + DestinationFolder)

                ' Delete the BACKOUT subfolder
                System.IO.Directory.Delete(DestinationFolder, True)

                ' Update file replacement result
                rVector.RemovalFileName.Add(hVector.GetInstalledFiles.Item(x))
                rVector.FileRemovalResult.Add(RemovalVector.FILE_FAILED)

                ' Write debug
                Logger.WriteDebug(CallStack, "Removal failed: " + rVector.RemovalItem)

                ' Update reason comment
                rVector.CommentString = "Reason: Failed to restore original file(s)."

                ' Return
                Return 7

            End Try

        Next

        ' *****************************
        ' - Remove REPLACED folder.
        ' *****************************

        ' Write debug
        Logger.WriteDebug(CallStack, "Delete REPLACED subfolder: " + ReplacedFolder)

        ' Delete REPLACED subfolder
        System.IO.Directory.Delete(ReplacedFolder, True)

        ' Return
        Return 0

    End Function

End Class