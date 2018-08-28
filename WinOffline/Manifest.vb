Partial Public Class WinOffline

    Public Class Manifest

        ' Class variables
        Private Shared PatchManifest As New ArrayList
        Private Shared RemovalManifest As New ArrayList
        Private Shared HistoryManifest As New ArrayList
        Private Shared ExceptionManifest As New ArrayList

        ' Class contstants
        Public Const PATCH_MANIFEST As Integer = 0
        Public Const REMOVAL_MANIFEST As Integer = 1
        Public Const HISTORY_MANIFEST As Integer = 2
        Public Const EXCEPTION_MANIFEST As Integer = 3

        Public Shared Sub UpdateManifest(ByVal CallStack As String,
                                         ByVal ManifestType As Integer,
                                         ByVal NewItem As Object)

            ' Update call stack
            CallStack += "UpdateManifest|"

            ' If object is a patch, verify it's unique with current manifest
            If ManifestType = PATCH_MANIFEST Then

                ' Cast the object
                NewItem = DirectCast(NewItem, PatchVector)

                ' Iterate patch manifest
                For Each pVector As PatchVector In PatchManifest

                    ' Check for a match
                    If NewItem.PatchFile.GetFriendlyName.Equals(pVector.PatchFile.GetFriendlyName) Then

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Warning: Manifest item is not unique-- skipping.")

                        ' Match found
                        Return

                    End If

                Next

            End If

            ' Add to appropriate manifest
            If ManifestType = PATCH_MANIFEST Then

                ' Cast the object
                NewItem = DirectCast(NewItem, PatchVector)

                ' Add to patch manifest
                PatchManifest.Add(NewItem)

            ElseIf ManifestType = REMOVAL_MANIFEST Then

                ' Cast the object
                NewItem = DirectCast(NewItem, RemovalVector)

                ' Add to removal manifest
                RemovalManifest.Add(NewItem)

            ElseIf ManifestType = HISTORY_MANIFEST Then

                ' Cast the object
                NewItem = DirectCast(NewItem, HistoryVector)

                ' Add to history manifest
                HistoryManifest.Add(NewItem)

            ElseIf ManifestType = EXCEPTION_MANIFEST Then

                ' Cast the object
                NewItem = DirectCast(NewItem, Array)

                ' Iterate the list
                For Each strLine As String In NewItem

                    ' Add to the manifest
                    ExceptionManifest.Add(strLine)

                Next

            End If

            ' Return success
            Return

        End Sub

        Public Shared Sub ReadCache(ByVal CallStack As String,
                                    ByVal ManifestType As Integer)

            ' Local variables
            Dim CacheFile As String = ""
            Dim CacheReader As System.IO.StreamReader
            Dim strLine As String = ""
            Dim RestoredManifest As New ArrayList

            ' Update call stack
            CallStack += "ReadCache|"

            ' Select cache file
            If ManifestType = PATCH_MANIFEST Then

                ' Read patch manifest cache
                PatchManifest = PatchVector.ReadCache(CallStack)

            ElseIf ManifestType = REMOVAL_MANIFEST Then

                ' Read removal manifest cache
                RemovalManifest = RemovalVector.ReadCache(CallStack)

            ElseIf ManifestType = HISTORY_MANIFEST Then

                ' Read history manifest cache
                HistoryManifest = HistoryVector.ReadCache(CallStack)

            ElseIf ManifestType = EXCEPTION_MANIFEST Then

                ' Exception cache file
                CacheFile = Globals.WinOfflineTemp + "\" + Globals.ProcessFriendlyName + "-exception.cache"

                ' Write debug
                Logger.WriteDebug(CallStack, "Open file: " + CacheFile)

                ' Check if cache file exists
                If Not System.IO.File.Exists(CacheFile) Then

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Cache file does not exist.")

                    ' Return
                    Return

                End If

                ' Open the cache file
                CacheReader = New System.IO.StreamReader(CacheFile)

                ' Iterate cache file contents
                Do While CacheReader.Peek() >= 0

                    ' Read a line
                    strLine = CacheReader.ReadLine()

                    ' Add to restored manifest
                    RestoredManifest.Add(strLine)

                Loop

                ' Write debug
                Logger.WriteDebug(CallStack, "Close file: " + CacheFile)

                ' Close the cache file
                CacheReader.Close()

                ' Replace exception manifest
                ExceptionManifest = RestoredManifest

            Else

                ' Unrecognized cache file
                Return

            End If

        End Sub

        Public Shared Sub WriteCache(ByVal CallStack As String,
                                     ByVal ManifestType As Integer)

            ' Local variables
            Dim CacheFile As String = ""
            Dim CacheWriter As System.IO.StreamWriter
            Dim strLine As String

            ' Update call stack
            CallStack += "WriteCache|"

            ' Select cache file
            If ManifestType = PATCH_MANIFEST Then

                ' Write patch manifest cache
                PatchVector.WriteCache(CallStack, PatchManifest)

            ElseIf ManifestType = REMOVAL_MANIFEST Then

                ' Write removal manifest cache
                RemovalVector.WriteCache(CallStack, RemovalManifest)

            ElseIf ManifestType = HISTORY_MANIFEST Then

                ' Write history manifest cache
                HistoryVector.WriteCache(CallStack, HistoryManifest)

            ElseIf ManifestType = EXCEPTION_MANIFEST Then

                ' *****************************
                ' - Flush prior cache file.
                ' *****************************

                ' Exception cache file
                CacheFile = Globals.WinOfflineTemp + "\" + Globals.ProcessFriendlyName + "-exception.cache"

                ' Verify any prior cache file is deleted before creating a new file
                If System.IO.File.Exists(CacheFile) Then

                    ' Delete prior cache file
                    Utility.DeleteFile(CallStack, CacheFile)

                ElseIf Not System.IO.Directory.Exists(Globals.WinOfflineTemp) Then

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Warning: Temporary folder already cleaned up!")
                    Logger.WriteDebug(CallStack, "Warning: Cache file write aborted!")

                    ' Return
                    Return

                End If

                ' *****************************
                ' - Create a new cache file.
                ' *****************************

                ' Write debug
                Logger.WriteDebug(CallStack, "Write cache file: " + CacheFile)

                ' Open the cache file for write
                CacheWriter = New System.IO.StreamWriter(CacheFile, False)

                ' Iterate the exception manifest
                For n As Integer = 0 To ExceptionManifest.Count() - 1

                    ' Get a line
                    strLine = ExceptionManifest.Item(n)

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Write: " + strLine)

                    ' Write exception to cache
                    CacheWriter.WriteLine(strLine)

                Next

                ' Write debug
                Logger.WriteDebug(CallStack, "Close file: " + CacheFile)

                ' Close cache file
                CacheWriter.Close()


            Else

                ' Unrecognized cache file
                Return

            End If

        End Sub

        Public Shared Function PatchManifestCount() As Integer

            ' Return the count
            Return PatchManifest.Count

        End Function

        Public Shared Function GetPatchFromManifest(index As Integer) As PatchVector

            ' Return requested PatchVector
            Return PatchManifest.Item(index)

        End Function

        Public Shared Function PatchManifestContains(ByVal PatchFriendlyName As String) As Boolean

            ' Iterate patch manifest
            For Each pVector As PatchVector In PatchManifest

                ' Check for a match
                If pVector.PatchFile.GetFriendlyName.ToLower.Equals(PatchFriendlyName.ToLower) Then

                    ' Return
                    Return True

                End If

            Next

            ' Return
            Return False

        End Function

        Public Shared Function RemovalManifestCount() As Integer

            ' Return the count
            Return RemovalManifest.Count

        End Function

        Public Shared Function GetRemovalFromManifest(index As Integer) As RemovalVector

            ' Return requested RemovalVector
            Return RemovalManifest.Item(index)

        End Function

        Public Shared Sub ResetRemovalManifest()

            ' Clear the removal manifest
            RemovalManifest.Clear()

        End Sub

        Public Shared Function HistoryManifestCount() As Integer

            ' Return the count
            Return HistoryManifest.Count

        End Function

        Public Shared Function GetHistoryFromManifest(index As Integer) As HistoryVector

            ' Return requested HisotryVector
            Return HistoryManifest.Item(index)

        End Function

        Public Shared Function IsAlreadyApplied(ByVal pVector As PatchVector) As Boolean

            ' Local variables
            Dim IsApplied As Boolean = False

            ' Iterate the history manifest
            For Each hVector As HistoryVector In HistoryManifest

                ' Match the patch, by name
                If hVector.GetPatchName.ToLower.Equals(pVector.PatchFile.GetFriendlyName.ToLower) Then

                    ' Check the history action
                    If hVector.IsApplyAction Then

                        ' Update status
                        IsApplied = True

                    Else

                        ' Update status
                        IsApplied = False

                    End If

                End If

            Next

            ' Return
            Return IsApplied

        End Function

        Public Shared Function DebugLogSummary() As String

            ' Local variables
            Dim PatchSummary As String = ""
            Dim CurrentComponent As String = ""
            Dim CommandFile As String = ""
            Dim ReturnCode As String = ""
            Dim ReplaceFile As String = ""
            Dim ExceptionCounter As Integer = 1

            ' Calculate overline
            For i As Integer = 0 To Globals.ProcessFriendlyName.Length + 14

                ' Add underline
                PatchSummary += "-"

            Next

            ' Update summary -- Patch manifest
            PatchSummary += Environment.NewLine + Globals.ProcessFriendlyName + " Event Summary:" + Environment.NewLine

            ' Calculate underline
            For i As Integer = 0 To Globals.ProcessFriendlyName.Length + 14

                ' Add underline
                PatchSummary += "-"

            Next

            ' Add a new line
            PatchSummary += Environment.NewLine

            ' Append patching summary
            PatchSummary += GeneratePatchingSummary()

            ' Append library cleanup summary
            If Globals.CheckSDLibrarySwitch OrElse Globals.CleanupSDLibrarySwitch Then

                ' Update summary
                PatchSummary += "Software library cleanup summary:" + Environment.NewLine

                ' Iterate library cleanup events
                For Each strLine As String In LibraryManager.LibraryEvents

                    ' Append event to summary
                    PatchSummary += "-" + strLine + Environment.NewLine

                Next

                ' Add additional line
                PatchSummary += Environment.NewLine

            End If

            ' Update summary -- History manifest
            For Each hVector As HistoryVector In HistoryManifest

                ' Check for product component change
                If Not hVector.GetProductComponent.Equals(CurrentComponent) Then

                    ' Set the new component
                    CurrentComponent = hVector.GetProductComponent

                    ' Calculate overline
                    For i As Integer = 0 To CurrentComponent.Length + 8

                        ' Add underline
                        PatchSummary += "-"

                    Next

                    ' Update summary
                    PatchSummary += Environment.NewLine + CurrentComponent + " History:" + Environment.NewLine

                    ' Calculate underline
                    For i As Integer = 0 To CurrentComponent.Length + 8

                        ' Add underline
                        PatchSummary += "-"

                    Next

                    ' Add a new line
                    PatchSummary += Environment.NewLine

                End If

                ' Check for an empty vector
                If hVector.IsEmpty Then

                    ' Update summary
                    PatchSummary += "No history." + Environment.NewLine + Environment.NewLine

                Else

                    ' Add summary header
                    PatchSummary += hVector.GetHeader + Environment.NewLine

                    ' Add summary installed files
                    PatchSummary += hVector.GetInstallFilesRaw + Environment.NewLine

                End If

            Next

            ' Calculate overline
            For i As Integer = 0 To Globals.ProcessFriendlyName.Length + 18

                ' Add underline
                PatchSummary += "-"

            Next

            ' Update summary -- Exception manifest
            PatchSummary += Environment.NewLine + Globals.ProcessFriendlyName + " Exception Summary:" + Environment.NewLine

            ' Calculate underline
            For i As Integer = 0 To Globals.ProcessFriendlyName.Length + 18

                ' Add underline
                PatchSummary += "-"

            Next

            ' Add a new line
            PatchSummary += Environment.NewLine

            ' Check for Exceptions
            If ExceptionManifest.Count = 0 Then

                ' Update summary
                PatchSummary += "No exceptions thrown." + Environment.NewLine

            Else

                ' Iterate the Exception Stack
                For i As Integer = 0 To ExceptionManifest.Count() - 2 Step 2

                    ' Update summary
                    PatchSummary += Environment.NewLine + "Exception #" + ExceptionCounter.ToString + ":" + Environment.NewLine
                    PatchSummary += "----------------" + Environment.NewLine

                    ' Update summary
                    PatchSummary += ExceptionManifest.Item(i).ToString + Environment.NewLine
                    PatchSummary += ExceptionManifest.Item(i + 1).ToString + Environment.NewLine

                    ' Update summary
                    PatchSummary += "----------------" + Environment.NewLine

                    ' Increment counter
                    ExceptionCounter += 1

                Next

            End If

            ' Add new lines
            PatchSummary += Environment.NewLine

            ' Calculate overline
            For i As Integer = 0 To Globals.ProcessFriendlyName.Length + 12

                ' Add underline
                PatchSummary += "-"

            Next

            ' Stub debug log header
            PatchSummary += Environment.NewLine + Globals.ProcessFriendlyName + " Debug trace:" + Environment.NewLine

            ' Calculate underline
            For i As Integer = 0 To Globals.ProcessFriendlyName.Length + 12

                ' Add underline
                PatchSummary += "-"

            Next

            ' Return
            Return PatchSummary

        End Function

        Public Shared Function DebugConsoleSummary() As String

            ' Local variables
            Dim PatchSummary As String = Environment.NewLine
            Dim CurrentComponent As String = ""
            Dim ExceptionCounter As Integer = 1

            ' Update summary -- History manifest
            For Each hVector As HistoryVector In HistoryManifest

                ' Check for product component change
                If Not hVector.GetProductComponent.Equals(CurrentComponent) Then

                    ' Set the new component
                    CurrentComponent = hVector.GetProductComponent

                    ' Calculate overline
                    For i As Integer = 0 To CurrentComponent.Length + 8

                        ' Add underline
                        PatchSummary += "-"

                    Next

                    ' Update summary
                    PatchSummary += Environment.NewLine + CurrentComponent + " History:" + Environment.NewLine

                    ' Calculate underline
                    For i As Integer = 0 To CurrentComponent.Length + 8

                        ' Add underline
                        PatchSummary += "-"

                    Next

                    ' Add new line
                    PatchSummary += Environment.NewLine

                End If

                ' Check for an empty vector
                If hVector.IsEmpty Then

                    ' Update summary
                    PatchSummary += "No history." + Environment.NewLine + Environment.NewLine

                Else

                    ' Add summary header
                    PatchSummary += hVector.GetHeader + Environment.NewLine

                    ' Add summary installed files
                    PatchSummary += hVector.GetInstallFilesRaw + Environment.NewLine

                End If

            Next

            ' Calculate overline
            For i As Integer = 0 To Globals.ProcessFriendlyName.Length + 18

                ' Add underline
                PatchSummary += "-"

            Next

            ' Update summary -- Exception manifest
            PatchSummary += Environment.NewLine + Globals.ProcessFriendlyName + " Exception Summary:" + Environment.NewLine

            ' Calculate overline
            For i As Integer = 0 To Globals.ProcessFriendlyName.Length + 18

                ' Add underline
                PatchSummary += "-"

            Next

            ' Add new line
            PatchSummary += Environment.NewLine

            ' Check for exceptions
            If ExceptionManifest.Count = 0 Then

                ' Update summary
                PatchSummary += "No exceptions thrown." + Environment.NewLine

            Else

                ' Iterate the exception Stack
                For i As Integer = 0 To ExceptionManifest.Count() - 2 Step 2

                    ' Update summary
                    PatchSummary += Environment.NewLine + "Exception #" + ExceptionCounter.ToString + ":" + Environment.NewLine
                    PatchSummary += "----------------" + Environment.NewLine

                    ' Update summary
                    PatchSummary += ExceptionManifest.Item(i).ToString + Environment.NewLine
                    PatchSummary += ExceptionManifest.Item(i + 1).ToString + Environment.NewLine

                    ' Update summary
                    PatchSummary += "----------------" + Environment.NewLine

                    ' Increment counter
                    ExceptionCounter += 1

                Next

            End If

            ' Add new line
            PatchSummary += Environment.NewLine

            ' Calculate overline
            For i As Integer = 0 To Globals.ProcessFriendlyName.Length + 14

                ' Add underline
                PatchSummary += "-"

            Next

            ' Update summary -- Patch manifest
            PatchSummary += Environment.NewLine + Globals.ProcessFriendlyName + " Event Summary:" + Environment.NewLine

            ' Calculate underline
            For i As Integer = 0 To Globals.ProcessFriendlyName.Length + 14

                ' Add underline
                PatchSummary += "-"

            Next

            ' Add new line
            PatchSummary += Environment.NewLine

            ' Append patching summary
            PatchSummary += GeneratePatchingSummary()

            ' Append library cleanup summary
            If Globals.CheckSDLibrarySwitch OrElse Globals.CleanupSDLibrarySwitch Then

                ' Update summary
                PatchSummary += "Software library cleanup summary:" + Environment.NewLine

                ' Iterate library cleanup events
                For Each strLine As String In LibraryManager.LibraryEvents

                    ' Append event to summary
                    PatchSummary += "-" + strLine + Environment.NewLine

                Next

                ' Add additional line
                PatchSummary += Environment.NewLine

            End If

            ' Return
            Return PatchSummary

        End Function

        Public Shared Function GeneratePatchingSummary() As String

            ' Local variables
            Dim PatchSummary As String = ""
            Dim CommandFile As String = ""
            Dim ReturnCode As String = ""
            Dim ReplaceFile As String = ""

            ' Check for patches
            If Globals.GetHistorySwitch Then

                ' Update summary
                PatchSummary += "Report agent patch history." + Environment.NewLine + Environment.NewLine

            ElseIf PatchManifest.Count = 0 AndAlso RemovalManifest.Count = 0 Then

                ' Update summary
                PatchSummary += "No patch operations performed." + Environment.NewLine + Environment.NewLine

            Else

                ' Check patch removal switch
                If Globals.RemovePatchSwitch Then

                    ' Iterate the removal manifest
                    For i As Integer = 0 To RemovalManifestCount() - 1

                        ' Add the removal name
                        PatchSummary += GetRemovalFromManifest(i).RemovalItem + ": "

                        ' Check the action
                        If GetRemovalFromManifest(i).RemovalAction = RemovalVector.SKIPPED Then

                            ' Update summary
                            PatchSummary += "SKIPPED" + Environment.NewLine
                            PatchSummary += GetRemovalFromManifest(i).CommentString + Environment.NewLine

                        ElseIf GetRemovalFromManifest(i).RemovalAction = RemovalVector.REMOVAL_OK Then

                            ' Update summary
                            PatchSummary += "REMOVED" + Environment.NewLine
                            PatchSummary += "History File: " + GetRemovalFromManifest(i).HistoryFile + Environment.NewLine

                        ElseIf GetRemovalFromManifest(i).RemovalAction = RemovalVector.REMOVAL_FAIL Then

                            ' Update summary
                            PatchSummary += "FAILED TO REMOVE" + Environment.NewLine
                            PatchSummary += GetRemovalFromManifest(i).CommentString + Environment.NewLine

                            ' Set patch error flag
                            Globals.PatchErrorDetected = True

                        Else

                            ' Update summary
                            PatchSummary += "UNKNOWN" + Environment.NewLine

                        End If

                        ' Iterate file removal results
                        For x As Integer = 0 To GetRemovalFromManifest(i).FileRemovalResult.Count - 1

                            ' Read file removal
                            ReplaceFile = GetRemovalFromManifest(i).RemovalFileName.Item(x)

                            ' Update summary
                            PatchSummary += "- " + ReplaceFile + " ["

                            ' Check file replacement results
                            If GetRemovalFromManifest(i).FileRemovalResult.Item(x) = RemovalVector.FILE_SKIPPED Then

                                ' Update summary
                                PatchSummary += "SKIPPED]" + Environment.NewLine

                            ElseIf GetRemovalFromManifest(i).FileRemovalResult.Item(x) = RemovalVector.FILE_OK Then

                                ' Update summary
                                PatchSummary += "OK]" + Environment.NewLine

                            ElseIf GetRemovalFromManifest(i).FileRemovalResult.Item(x) = RemovalVector.FILE_REBOOT_REQUIRED Then

                                ' Update summary
                                PatchSummary += "REBOOT REQUIRED]" + Environment.NewLine

                            ElseIf GetRemovalFromManifest(i).FileRemovalResult.Item(x) = RemovalVector.FILE_FAILED Then

                                ' Update summary
                                PatchSummary += "FAILED]" + Environment.NewLine

                            End If

                        Next

                        ' Update summary
                        PatchSummary += Environment.NewLine

                    Next

                Else

                    ' Iterate the patch manifest
                    For i As Integer = 0 To PatchManifest.Count - 1

                        ' Add the patch name
                        PatchSummary += GetPatchFromManifest(i).PatchFile.GetFriendlyName + ": "

                        ' Check the action
                        If GetPatchFromManifest(i).PatchAction = PatchVector.NOT_APPLICABLE Then

                            ' Update summary
                            PatchSummary += "NOT APPLICABLE" + Environment.NewLine
                            PatchSummary += GetPatchFromManifest(i).CommentString.Replace("NEWLINE", Environment.NewLine) + Environment.NewLine

                        ElseIf GetPatchFromManifest(i).PatchAction = PatchVector.ALREADY_APPLIED Then

                            ' Update summary
                            PatchSummary += "ALREADY APPLIED" + Environment.NewLine
                            PatchSummary += GetPatchFromManifest(i).CommentString.Replace("NEWLINE", Environment.NewLine) + Environment.NewLine

                        ElseIf GetPatchFromManifest(i).PatchAction = PatchVector.SKIPPED Then

                            ' Update summary
                            PatchSummary += "SKIPPED" + Environment.NewLine
                            PatchSummary += GetPatchFromManifest(i).CommentString.Replace("NEWLINE", Environment.NewLine) + Environment.NewLine

                        ElseIf GetPatchFromManifest(i).PatchAction = PatchVector.APPLY_OK Then

                            ' Update summary
                            PatchSummary += "APPLIED" + Environment.NewLine

                        ElseIf GetPatchFromManifest(i).PatchAction = PatchVector.APPLY_FAIL Then

                            ' Update summary
                            PatchSummary += "FAILED TO APPLY" + Environment.NewLine
                            PatchSummary += GetPatchFromManifest(i).CommentString.Replace("NEWLINE", Environment.NewLine) + Environment.NewLine

                            ' Set patch error flag
                            Globals.PatchErrorDetected = True

                        ElseIf GetPatchFromManifest(i).PatchAction = PatchVector.EXECUTE_OK Then

                            ' Update summary
                            PatchSummary += "EXECUTED" + Environment.NewLine

                        ElseIf GetPatchFromManifest(i).PatchAction = PatchVector.EXECUTE_FAIL Then

                            ' Update summary
                            PatchSummary += "FAILED TO EXECUTE" + Environment.NewLine
                            PatchSummary += GetPatchFromManifest(i).CommentString.Replace("NEWLINE", Environment.NewLine) + Environment.NewLine

                            ' Set patch error flag
                            Globals.PatchErrorDetected = True

                        Else

                            ' Update summary
                            PatchSummary += "UNKNOWN" + Environment.NewLine

                        End If

                        ' Append further details, based on the patch action
                        If GetPatchFromManifest(i).PatchAction = PatchVector.APPLY_OK Or
                            GetPatchFromManifest(i).PatchAction = PatchVector.APPLY_FAIL Or
                            GetPatchFromManifest(i).PatchAction = PatchVector.EXECUTE_OK Or
                            GetPatchFromManifest(i).PatchAction = PatchVector.EXECUTE_FAIL Then

                            ' Iterate PRESYSCMDs and return codes
                            For x As Integer = 0 To GetPatchFromManifest(i).GetPreCommandList.Count - 1

                                ' Read the command file
                                CommandFile = GetPatchFromManifest(i).GetPreCommandList.Item(x)

                                ' Ensure we have a return code
                                If GetPatchFromManifest(i).PreCmdReturnCodes.Count > x Then

                                    ' Read the return code
                                    ReturnCode = GetPatchFromManifest(i).PreCmdReturnCodes.Item(x)

                                Else

                                    ' Return code unavailable
                                    ReturnCode = "SKIPPED"

                                End If

                                ' Update summary
                                PatchSummary += "- " + CommandFile + " [Return Code: " + ReturnCode + "]" + Environment.NewLine

                            Next

                            ' Iterate file replacements
                            For x As Integer = 0 To GetPatchFromManifest(i).DestReplaceList.Count - 1

                                ' Read file replacement
                                ReplaceFile = GetPatchFromManifest(i).DestReplaceList.Item(x)

                                ' Update summary
                                PatchSummary += "- " + ReplaceFile + " ["

                                ' Check file replacement results
                                If GetPatchFromManifest(i).FileReplaceResult.Item(x) = PatchVector.FILE_SKIPPED Then

                                    ' Update summary
                                    PatchSummary += "SKIPPED]" + Environment.NewLine

                                ElseIf GetPatchFromManifest(i).FileReplaceResult.Item(x) = PatchVector.FILE_OK Then

                                    ' Update summary
                                    PatchSummary += "OK]" + Environment.NewLine

                                ElseIf GetPatchFromManifest(i).FileReplaceResult.Item(x) = PatchVector.FILE_REBOOT_REQUIRED Then

                                    ' Update summary
                                    PatchSummary += "REBOOT REQUIRED]" + Environment.NewLine

                                ElseIf GetPatchFromManifest(i).FileReplaceResult.Item(x) = PatchVector.FILE_FAILED Then

                                    ' Update summary
                                    PatchSummary += "FAILED]" + Environment.NewLine

                                End If

                            Next

                            ' Iterate SYSCMDs and return codes
                            For x As Integer = 0 To GetPatchFromManifest(i).GetSysCommandList.Count - 1

                                ' Read the command file
                                CommandFile = GetPatchFromManifest(i).GetSysCommandList.Item(x)

                                ' Ensure we have a return code
                                If GetPatchFromManifest(i).SysCmdReturnCodes.Count > x Then

                                    ' Read the return code
                                    ReturnCode = GetPatchFromManifest(i).SysCmdReturnCodes.Item(x)

                                Else

                                    ' Return code unavailable
                                    ReturnCode = "SKIPPED"

                                End If

                                ' Update summary
                                PatchSummary += "- " + CommandFile + " [Return Code: " + ReturnCode + "]" + Environment.NewLine

                            Next

                        End If

                        ' Update summary
                        PatchSummary += Environment.NewLine

                    Next

                End If

            End If

            ' Return
            Return PatchSummary

        End Function

        Public Shared Function FlushManifestData() As Boolean

            ' Local variables
            Dim DataCleared As Boolean = False

            ' Check if any manifest data will be cleared (simply to reduce a deinit message if datacleared is false)
            If PatchManifest.Count > 0 OrElse RemovalManifest.Count > 0 OrElse HistoryManifest.Count > 0 Then DataCleared = True

            ' Clear patch manifest
            PatchManifest = New ArrayList

            ' Clear removal manifest
            RemovalManifest = New ArrayList

            ' Clear the history manifest
            HistoryManifest = New ArrayList

            ' Don't clear the exception manifest, this should persist
            ' ExceptionStack = New ArrayList

            ' Return
            Return DataCleared

        End Function

    End Class

End Class