'****************************** Class Header *******************************\
' Project Name: WinOffline
' Class Name:   WinOffline
' File Name:    HistoryFileAPI.vb
' Author:       Brian Fontana
'***************************************************************************/

Partial Public Class WinOffline

    Public Const IT_CLIENT_MANAGER As String = "IT Client Manager -- ITCM"
    Public Const SHARED_COMPONENTS As String = "Shared Components -- SC"
    Public Const CA_MESSAGE_QUEUING As String = "CA Message Queuing Service -- CAM"
    Public Const SECURE_SOCKET_ADAPATER As String = "Secure Socket Adapter -- SSA"
    Public Const DATA_TRANSPORT As String = "Data Transport Service -- DTS"
    Public Const EXPLORER_GUI As String = "DSM Explorer GUI -- EGC"

    Public Shared Function ProcessHistoryFiles(ByVal CallStack As String)

        ' Local variables
        Dim Runlevel As Integer = 0

        ' Update call stack
        CallStack += "HistoryFile|"

        ' Read client auto history
        Runlevel = ReadHistoryFile(CallStack, IT_CLIENT_MANAGER, Globals.DSMFolder + Globals.HostName + ".his")

        ' Check the run level
        If Runlevel <> 0 Then

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Failed to process ITCM history file.")

            ' Return
            Return 1

        End If

        ' Read client auto shared component history
        Runlevel = ReadHistoryFile(CallStack, SHARED_COMPONENTS, Globals.SharedCompFolder + Globals.HostName + ".his")

        ' Check the run level
        If Runlevel <> 0 Then

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Failed to process shared components history file.")

            ' Return
            Return 2

        End If

        ' Read CAM history
        Runlevel = ReadHistoryFile(CallStack, CA_MESSAGE_QUEUING, Globals.CAMFolder + "\" + Globals.HostName + ".his")

        ' Check the run level
        If Runlevel <> 0 Then

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Failed to process CAM history file.")

            ' Return
            Return 3

        End If

        ' Check optional component -- SSA
        If Globals.SSAFolder IsNot Nothing Then

            ' Read SSA history
            Runlevel = ReadHistoryFile(CallStack, SECURE_SOCKET_ADAPATER, Globals.SSAFolder + Globals.HostName + ".his")

            ' Check the run level
            If Runlevel <> 0 Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Error: Failed to process SSA history file.")

                ' Return
                Return 4

            End If

        End If

        ' Check optional component -- DTS
        If Globals.DTSFolder IsNot Nothing Then

            ' Read DTS history
            Runlevel = ReadHistoryFile(CallStack, DATA_TRANSPORT, Globals.DTSFolder + "\" + Globals.HostName + ".his")

            ' Check the run level
            If Runlevel <> 0 Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Error: Failed to process DTS history file.")

                ' Return
                Return 5

            End If

        End If

        ' Check optional component -- EGC
        If Globals.EGCFolder IsNot Nothing Then

            ' Read EGC history
            Runlevel = ReadHistoryFile(CallStack, EXPLORER_GUI, Globals.EGCFolder + Globals.HostName + ".his")

            ' Check the run level
            If Runlevel <> 0 Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Error: Failed to process EGC history file.")

                ' Return
                Return 6

            End If

        End If

        ' Return
        Return 0

    End Function

    Public Shared Function ReadHistoryFile(ByVal CallStack As String, ByVal Component As String, ByVal HistoryFile As String)

        ' Local variables
        Dim HistoryFileFound As Boolean = False             ' Flag indicating if history file is available.
        Dim HistoryFileReader As System.IO.StreamReader     ' Input stream for reading history file.
        Dim HistoryArray As New ArrayList                   ' Array for storing raw history from file.
        Dim strLine As String                               ' Temp string for parsing the raw history.
        Dim HistoryEvents As New ArrayList                  ' Array for grouping related history file entries.
        Dim hVector As HistoryVector                        ' History vector object.
        Dim PatchCount As Integer = 0                       ' Patch counter (for summary purposes)

        ' Write debug
        Logger.WriteDebug(CallStack, "Component: " + Component)

        ' Check if patch history exists
        If System.IO.File.Exists(HistoryFile) Then

            ' Try processing patch history
            Try

                ' Set history file marker
                HistoryFileFound = True

                ' Write debug
                Logger.WriteDebug(CallStack, "Open file: " + HistoryFile)
                Logger.WriteDebug("############################################################")

                ' Open the patch history file
                HistoryFileReader = New System.IO.StreamReader(HistoryFile)

                ' Loop history file contents
                Do While HistoryFileReader.Peek() >= 0

                    ' Read a line
                    strLine = HistoryFileReader.ReadLine()

                    ' Write debug -- History file contents only
                    Logger.WriteDebug(strLine)

                    ' Update the array
                    HistoryArray.Add(strLine)

                Loop

                ' Write debug
                Logger.WriteDebug("############################################################")
                Logger.WriteDebug(CallStack, "Close file: " + HistoryFile)

                ' Close history file
                HistoryFileReader.Close()

                ' *****************************
                ' - Process the history records.
                ' *****************************

                ' Iterate history array
                For i As Integer = 0 To HistoryArray.Count - 1

                    ' Read an entry
                    strLine = HistoryArray.Item(i)

                    ' Check for major history entry (installed or backed out)
                    If strLine.StartsWith("[") And strLine.ToLower.Contains("installed") Then

                        ' Clear array for grouping related entries
                        HistoryEvents = New ArrayList

                        ' Add first entry to the array
                        HistoryEvents.Add(strLine)

                        ' Perform inner loop to add further related history entries
                        For j As Integer = i + 1 To HistoryArray.Count - 1

                            ' Read next line
                            strLine = HistoryArray.Item(j)

                            ' Check if next patch was reached
                            If strLine.StartsWith("[") Then

                                ' Break inner loop
                                Exit For

                            Else

                                ' Add related entry
                                HistoryEvents.Add(strLine)

                            End If

                            ' Skip ahead i, so we don't re-read the same patch
                            i = j

                        Next

                        ' Create history vector object
                        hVector = New HistoryVector(Component, HistoryEvents)

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Installed patch: " + hVector.GetPatchName)

                        ' Add history vector to manifest
                        Manifest.UpdateManifest(CallStack, Manifest.HISTORY_MANIFEST, hVector)

                        ' Tick the counter
                        PatchCount += 1

                    ElseIf strLine.StartsWith("[") And strLine.ToLower.Contains("backed out") Then

                        ' Clear array for grouping related entries
                        HistoryEvents = New ArrayList

                        ' Add first entry to the array
                        HistoryEvents.Add(strLine)

                        ' Create new history vector
                        hVector = New HistoryVector(Component, HistoryEvents)

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Removed patch: " + hVector.GetPatchName)

                        ' Add history vector to manifest
                        Manifest.UpdateManifest(CallStack, Manifest.HISTORY_MANIFEST, hVector)

                        ' Tick the counter
                        PatchCount += 1

                    End If

                Next

                ' Check for empty history file
                If PatchCount = 0 Then

                    ' Create an empty hisotry vector
                    hVector = New HistoryVector(Component, New ArrayList({"NoHistory"}))

                    ' Write debug
                    Logger.WriteDebug(CallStack, "History file is empty.")

                    ' Update the history table
                    Manifest.UpdateManifest(CallStack, Manifest.HISTORY_MANIFEST, hVector)

                End If

            Catch ex As Exception

                ' Write debug
                Logger.WriteDebug(CallStack, "Error: Exception caught during history file processing.")
                Logger.WriteDebug(ex.Message)
                Logger.WriteDebug(ex.StackTrace)

                ' Create exception
                Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

                ' Return
                Return 1

            End Try

        Else

            ' Create an empty hisotry vector
            hVector = New HistoryVector(Component, New ArrayList({"NoHistory"}))

            ' Write debug
            Logger.WriteDebug(CallStack, "History file: Unavailable.")

            ' Update the history table
            Manifest.UpdateManifest(CallStack, Manifest.HISTORY_MANIFEST, hVector)

        End If

        ' Return
        Return 0

    End Function

    Public Shared Sub AddtoHistory(ByVal CallStack As String, ByVal pVector As PatchVector)

        ' Local variables
        Dim HistoryFile As String = ""
        Dim HistoryFileInfo As System.IO.FileInfo
        Dim HistoryFileWriter As System.IO.StreamWriter
        Dim NewFile As Boolean = False

        ' Update call stack
        CallStack += "HistoryFile|"

        ' Build path to history file
        If pVector.IsClientAuto Then HistoryFile = Globals.DSMFolder + Globals.HostName + ".his"
        If pVector.IsSharedComponent Then HistoryFile = Globals.SharedCompFolder + Globals.HostName + ".his"
        If pVector.IsCAM Then HistoryFile = Globals.CAMFolder + "\" + Globals.HostName + ".his"
        If pVector.IsSSA Then HistoryFile = Globals.SSAFolder + Globals.HostName + ".his"
        If pVector.IsDataTransport Then HistoryFile = Globals.DTSFolder + "\" + Globals.HostName + ".his"
        If pVector.IsExplorerGUI Then HistoryFile = Globals.EGCFolder + Globals.HostName + ".his"

        ' Encapsulte history file update
        Try

            ' Write debug
            Logger.WriteDebug(CallStack, "Append history file: " + HistoryFile)

            ' Verify history file exists
            If System.IO.File.Exists(HistoryFile) Then

                ' Get history file info
                HistoryFileInfo = New System.IO.FileInfo(HistoryFile)

                ' Check for read-only attribute
                If HistoryFileInfo.Attributes And IO.FileAttributes.ReadOnly Then

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Unset read-only file attribute: " + HistoryFile)

                    ' Unset read-only
                    HistoryFileInfo.Attributes = HistoryFileInfo.Attributes Xor FileAttribute.ReadOnly

                End If

                ' Normalize the history file
                NormalizeHistoryFile(CallStack, HistoryFile)

            Else

                ' Set new file flag
                NewFile = True

            End If

            ' Open history file for append
            HistoryFileWriter = New IO.StreamWriter(HistoryFile, True)

            ' Add new file header
            If NewFile Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Write: " + "[" + pVector.GetApplyPTFDateFormat + "] - PTF Wizard created this history file.")

                ' Write header
                HistoryFileWriter.WriteLine("[" + pVector.GetApplyPTFDateFormat + "] - PTF Wizard created this history file." + Environment.NewLine)

            End If

            ' Write debug
            Logger.WriteDebug(CallStack, "Write patch summary: ")
            Logger.WriteDebug("############################################################")
            Logger.WriteDebug(pVector.GetApplyPTFSummary)
            Logger.WriteDebug("############################################################")

            ' Add the ApplyPTF formatted patch summary
            HistoryFileWriter.WriteLine(pVector.GetApplyPTFSummary)

            ' Write debug
            Logger.WriteDebug(CallStack, "Close file: " + HistoryFile)

            ' Close cache file
            HistoryFileWriter.Close()

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Exception caught appending history file.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

            ' Create exception
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

            ' Return
            Return

        End Try

        ' Return
        Return

    End Sub

    Public Shared Sub RemoveFromHistory(ByVal CallStack As String, ByVal rVector As RemovalVector)

        ' Local variables
        Dim HistoryFileReader As System.IO.StreamReader
        Dim HistoryFileWriter As System.IO.StreamWriter
        Dim strLine As String
        Dim HistoryFileArray As New ArrayList
        Dim BeginIndex As Integer = -1
        Dim EndIndex As Integer = -1
        Dim DelimCount As Integer = 0
        Dim HistoryFileInfo As System.IO.FileInfo

        ' Update call stack
        CallStack += "HistoryFile|"

        ' *****************************
        ' - Read history file.
        ' *****************************

        ' Check if patch history exists
        If System.IO.File.Exists(rVector.HistoryFile) Then

            ' Encapsulate history file read
            Try

                ' Write debug
                Logger.WriteDebug(CallStack, "Read history file: " + rVector.HistoryFile)

                ' Open the patch history file
                HistoryFileReader = New System.IO.StreamReader(rVector.HistoryFile)

                ' Loop history file contents
                Do While HistoryFileReader.Peek() >= 0

                    ' Read a line
                    strLine = HistoryFileReader.ReadLine()

                    ' Update the array
                    HistoryFileArray.Add(strLine)

                Loop

                ' Write debug
                Logger.WriteDebug(CallStack, "Close file: " + rVector.HistoryFile)

                ' Close history file
                HistoryFileReader.Close()

            Catch ex As Exception

                ' Write debug
                Logger.WriteDebug(CallStack, "Error: Exception caught reading history file.")
                Logger.WriteDebug(ex.Message)
                Logger.WriteDebug(ex.StackTrace)

                ' Create exception
                Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

                ' Return
                Return

            End Try

        End If

        ' *****************************
        ' - Remove last patch occurrence.
        ' *****************************

        ' Encapsulate array operations
        Try

            ' Iterate the HistoryFileArray
            For i As Integer = 0 To HistoryFileArray.Count - 1

                ' Check for begin index
                If HistoryFileArray.Item(i).ToString.ToLower.Contains("ptf wizard installed " + rVector.RemovalItem.ToLower) Then

                    ' Set begin index
                    BeginIndex = i

                    ' Reset end index
                    EndIndex = 0

                    ' Scan ahead for end index
                    For j As Integer = i + 1 To HistoryFileArray.Count - 1

                        ' Check for end index
                        If HistoryFileArray.Item(j).ToString.StartsWith("[") Or j = HistoryFileArray.Count - 1 Then

                            ' Set end index
                            EndIndex = j

                            ' Exit loop
                            Exit For

                        End If

                    Next

                End If

            Next

            ' Check index
            If BeginIndex <> -1 And EndIndex <> -1 And BeginIndex < EndIndex Then

                ' Remove the last occurrence of patch history
                HistoryFileArray.RemoveRange(BeginIndex, EndIndex - BeginIndex)

                ' Write debug
                Logger.WriteDebug(CallStack, "Last occurrence removed: " + rVector.RemovalItem)

            End If

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Exception caught removing last patch occurrence.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

            ' Create exception
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

            ' Return
            Return

        End Try

        ' *****************************
        ' - Write history file.
        ' *****************************

        ' Encapsulate history file rewrite
        Try

            ' Get history file info
            HistoryFileInfo = New System.IO.FileInfo(rVector.HistoryFile)

            ' Check for read-only attribute
            If HistoryFileInfo.Attributes And IO.FileAttributes.ReadOnly Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Unset read-only file attribute: " + rVector.HistoryFile)

                ' Unset read-only
                HistoryFileInfo.Attributes = HistoryFileInfo.Attributes Xor FileAttribute.ReadOnly

            End If

            ' Write debug
            Logger.WriteDebug(CallStack, "Write history file: " + rVector.HistoryFile)

            ' Open the patch history file
            HistoryFileWriter = New System.IO.StreamWriter(rVector.HistoryFile, False)

            ' Loop HistoryFileArray contents
            For Each HistoryLine In HistoryFileArray

                ' Write a line
                HistoryFileWriter.WriteLine(HistoryLine)

            Next

            ' Write debug
            Logger.WriteDebug(CallStack, "Close file: " + rVector.HistoryFile)

            ' Close history file
            HistoryFileWriter.Close()

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Exception caught rewriting history file.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

            ' Create exception
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

            ' Return
            Return

        End Try

        ' *****************************
        ' - Normalize history file.
        ' *****************************

        ' Normalize the history file
        NormalizeHistoryFile(CallStack, rVector.HistoryFile)

        ' Return
        Return

    End Sub

    Public Shared Sub NormalizeHistoryFile(ByVal CallStack As String, ByVal HistoryFile As String)

        ' Local variables
        Dim HistoryFileReader As System.IO.StreamReader
        Dim HistoryFileWriter As System.IO.StreamWriter
        Dim HistoryFileCreateDate As Date
        Dim CurrentLine As String
        Dim PreviousLine As String
        Dim ForwardLine As String
        Dim strLine As String
        Dim HistoryFileArray As New ArrayList
        Dim BeginIndex As Integer = 0
        Dim EndIndex As Integer = 0
        Dim DelimCount As Integer = 0

        ' *****************************
        ' - Read history file.
        ' *****************************

        ' Check if patch history exists
        If System.IO.File.Exists(HistoryFile) Then

            ' Read in history file
            Try

                ' Read history file creation date
                HistoryFileCreateDate = System.IO.File.GetCreationTime(HistoryFile)

                ' Write debug
                Logger.WriteDebug(CallStack, "History file creation date: " + HistoryFileCreateDate.ToString("F"))
                Logger.WriteDebug(CallStack, "Read history file contents: " + HistoryFile)

                ' Open the patch history file
                HistoryFileReader = New System.IO.StreamReader(HistoryFile)

                ' Loop history file contents
                Do While HistoryFileReader.Peek() >= 0

                    ' Read a line
                    strLine = HistoryFileReader.ReadLine()

                    ' Update the array
                    HistoryFileArray.Add(strLine)

                Loop

                ' Write debug
                Logger.WriteDebug(CallStack, "Close file: " + HistoryFile)

                ' Close history file
                HistoryFileReader.Close()

            Catch ex As Exception

                ' Write debug
                Logger.WriteDebug(CallStack, "Error: Exception caught reading history file.")
                Logger.WriteDebug(ex.Message)
                Logger.WriteDebug(ex.StackTrace)

                ' Create exception
                Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

                ' Return
                Return

            End Try

        End If

        ' *****************************
        ' - Normalize history file.
        ' *****************************

        ' Encapsulate normalization
        Try

            ' Iterate the HistoryFileArray
            For i As Integer = HistoryFileArray.Count - 1 To 0 Step -1

                ' Read current line
                CurrentLine = HistoryFileArray.Item(i)

                ' Read previous line
                If i < HistoryFileArray.Count - 1 Then

                    ' Read previous line
                    PreviousLine = HistoryFileArray.Item(i + 1)

                Else

                    ' Stub previous line
                    PreviousLine = "-1"

                End If

                ' Check for empty line
                If CurrentLine.Equals("") AndAlso PreviousLine.Equals("") Then

                    ' Remove extra unnecessary space
                    HistoryFileArray.RemoveAt(i)

                    ' Continue loop
                    Continue For

                ElseIf CurrentLine.ToLower.StartsWith("installedfile=") Then

                    ' Check for ceiling
                    If i - 8 < 0 Then

                        ' Ceiling reached -- remove stray entry
                        HistoryFileArray.RemoveAt(i)

                        ' Continue loop -- ceiling limit for install record
                        Continue For

                    End If

                    ' Iterate forward
                    For j As Integer = i - 1 To 0 Step -1

                        ' Read forward
                        ForwardLine = HistoryFileArray.Item(j)

                        ' Check for empty line
                        If ForwardLine.Equals("") Then

                            ' Remove extra unnecessary space
                            HistoryFileArray.RemoveAt(j)

                            ' Decrement i
                            i -= 1

                            ' Continue loop
                            Continue For

                        ElseIf ForwardLine.ToLower.StartsWith("installedfile=") Then

                            ' Entry valid -- continue loop
                            Continue For

                        ElseIf ForwardLine.ToLower.StartsWith("supersede=") Then

                            ' Check for minimum of (8) required lines preceding this record
                            If j - 8 >= 0 AndAlso
                                DirectCast(HistoryFileArray.Item(j - 1), String).ToLower.StartsWith("mcoreqs=") AndAlso
                                DirectCast(HistoryFileArray.Item(j - 2), String).ToLower.StartsWith("coreqs=") AndAlso
                                DirectCast(HistoryFileArray.Item(j - 3), String).ToLower.StartsWith("mprereqs=") AndAlso
                                DirectCast(HistoryFileArray.Item(j - 4), String).ToLower.StartsWith("prereqs=") AndAlso
                                DirectCast(HistoryFileArray.Item(j - 5), String).ToLower.StartsWith("component=") AndAlso
                                DirectCast(HistoryFileArray.Item(j - 6), String).ToLower.StartsWith("genlevel=") AndAlso
                                DirectCast(HistoryFileArray.Item(j - 7), String).ToLower.StartsWith("release=") AndAlso
                                DirectCast(HistoryFileArray.Item(j - 8), String).StartsWith("[") AndAlso
                                DirectCast(HistoryFileArray.Item(j - 8), String).Contains("]") AndAlso
                                DirectCast(HistoryFileArray.Item(j - 8), String).ToLower.Contains("installed") Then

                                ' Check forward for empty line
                                If j - 9 >= 0 AndAlso Not DirectCast(HistoryFileArray.Item(j - 9), String).Equals("") Then

                                    ' Insert empty line
                                    HistoryFileArray.Insert(j - 9, "")

                                End If

                                ' Decrement i to (j-9)
                                i = j - 9

                                ' Exit condition -- install record is valid
                                Exit For

                            Else

                                ' Invalid entry -- remove i thru j
                                For k As Integer = i To j Step -1

                                    ' Remove invalid entry
                                    HistoryFileArray.RemoveAt(k)

                                Next

                                ' Decrement i to j's position
                                i = j

                                ' Exit condition -- invalid set of records removed
                                Exit For

                            End If

                        Else

                            ' Invalid entry -- remove i thru j
                            For k As Integer = i To j Step -1

                                ' Remove invalid entry
                                HistoryFileArray.RemoveAt(k)

                            Next

                            ' Decrement i to j's position
                            i = j

                            ' Exit condition -- invalid set of records removed
                            Exit For

                        End If

                    Next

                    ' Continue loop
                    Continue For

                ElseIf CurrentLine.ToLower.StartsWith("#") OrElse
                    CurrentLine.Equals("") OrElse
                    CurrentLine.ToLower.Contains("ptf wizard created this history file") Then

                    ' Check if create event is only line in the file
                    If CurrentLine.ToLower.Contains("ptf wizard created this history file") AndAlso PreviousLine.Equals("-1") Then

                        ' Insert a new line
                        HistoryFileArray.Add(Environment.NewLine)

                    ElseIf CurrentLine.ToLower.Contains("ptf wizard created this history file") AndAlso i <> 0 Then

                        ' Remove misplaced line
                        HistoryFileArray.RemoveAt(i)

                        ' Continue loop
                        Continue For

                    End If

                    ' Valid entry -- Continue loop
                    Continue For

                Else

                    ' Remove unrecognized line
                    HistoryFileArray.RemoveAt(i)

                    ' Continue loop
                    Continue For

                End If

            Next

            ' Check first line for creation header
            If HistoryFileArray.Count = 0 OrElse Not DirectCast(HistoryFileArray.Item(0), String).ToLower.Contains("ptf wizard created this history file") Then

                ' Insert a new first line based on file creation date
                HistoryFileArray.Insert(0, HistoryFileCreateDate.ToString("[ddd MMM d HH:mm:ss yyyy]") + " - PTF Wizard created this history file.")

            End If

            ' Write debug
            Logger.WriteDebug(CallStack, "History file normalized.")

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Exception caught normalizing history file.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

            ' Create exception
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

            ' Return
            Return

        End Try

        ' *****************************
        ' - Write history file.
        ' *****************************

        ' Encapsulate history file rewrite
        Try

            ' Write debug
            Logger.WriteDebug(CallStack, "Rewrite history file: " + HistoryFile)

            ' Open the patch history file
            HistoryFileWriter = New System.IO.StreamWriter(HistoryFile, False)

            ' Loop HistoryFileArray contents
            For Each HistoryLine In HistoryFileArray

                ' Write a line
                HistoryFileWriter.WriteLine(HistoryLine)

            Next

            ' Write debug
            Logger.WriteDebug(CallStack, "Close file: " + HistoryFile)

            ' Close history file
            HistoryFileWriter.Close()

        Catch ex As Exception

            ' Write debug
            Logger.WriteDebug(CallStack, "Error: Exception caught rewriting history file.")
            Logger.WriteDebug(ex.Message)
            Logger.WriteDebug(ex.StackTrace)

            ' Create exception
            Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

            ' Return
            Return

        End Try

    End Sub

End Class