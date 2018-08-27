Imports System.Windows.Forms

Partial Public Class WinOfflineUI

    Private Sub InitStartPanel()

        ' Replace WinOffline referecnes with process name
        grpWinOfflineWelcome.Text = grpWinOfflineWelcome.Text.Replace("WinOffline", Globals.ProcessFriendlyName)
        grpHistory.Text = "Patch History for " + Globals.HostName

        ' Build WinOffline patch history tree
        BuildPatchHistoryTree()

    End Sub

    Private Sub BuildPatchHistoryTree()

        ' Check if client automation is installed
        If Globals.ITCMInstalled = False Then

            ' Local variables
            Dim ComponentRootNode As New TreeNode("Client Automation")

            ' Update UI
            treHistory.Nodes.Add(ComponentRootNode)
            ComponentRootNode.Nodes.Add(New TreeNode("Not Installed."))
            ComponentRootNode.Expand()

            ' Disable actions
            rbnApply.Enabled = False
            rbnBackout.Enabled = False
            rbnCLIHelp.Select()

            ' Return
            Exit Sub

        End If

        ' Read main component history files
        ReadHistoryFile("Client Automation", Globals.DSMFolder + Globals.HostName + ".his")
        ReadHistoryFile("Shared Components", Globals.SharedCompFolder + Globals.HostName + ".his")
        ReadHistoryFile("CA Message Queuing Service", Globals.CAMFolder + "\" + Globals.HostName + ".his")

        ' Read optional component history files
        If Globals.SSAFolder IsNot Nothing Then
            ReadHistoryFile("Secure Socket Adapter", Globals.SSAFolder + Globals.HostName + ".his")
        End If
        If Globals.DTSFolder IsNot Nothing Then
            ReadHistoryFile("Data Transport Service", Globals.DTSFolder + "\" + Globals.HostName + ".his")
        End If
        If Globals.EGCFolder IsNot Nothing Then
            ReadHistoryFile("DSM Explorer GUI", Globals.EGCFolder + "\" + Globals.HostName + ".his")
        End If

        ' Set scroll position
        treHistory.Nodes(0).EnsureVisible()

    End Sub

    Private Sub ReadHistoryFile(ByVal Component As String, ByVal HistoryFile As String)

        ' Local variables
        Dim ComponentRootNode As New TreeNode(Component)
        Dim HistoryFileReader As System.IO.StreamReader
        Dim HistoryArray As New ArrayList
        Dim strLine As String
        Dim PatchName As String
        Dim TempFileList As New ArrayList
        Dim PatchArray As New ArrayList
        Dim InstalledFileArray As New ArrayList
        Dim PatchFileName As String
        Dim DuplicateIndex As Integer = 0
        Dim DuplicateArray As New ArrayList
        Dim DuplicateFileArray As New ArrayList
        Dim DuplicateRemoved As Boolean = False

        ' Create the root node
        treHistory.Nodes.Add(ComponentRootNode)

        ' Check if the history file exists
        If System.IO.File.Exists(HistoryFile) Then

            ' Open the history file
            Try

                ' Open the history file
                HistoryFileReader = New System.IO.StreamReader(HistoryFile)

                ' Read contents of the history file
                Do While HistoryFileReader.Peek() >= 0

                    ' Add each line into the array
                    HistoryArray.Add(HistoryFileReader.ReadLine())

                Loop

                ' Close the history file
                HistoryFileReader.Close()

            Catch ex As Exception

                ' Update UI for failure
                Dim NewNode As New TreeNode("Failed to read history file.")
                ComponentRootNode.Nodes.Add(NewNode)
                ComponentRootNode.Expand()

            End Try

            ' Check for available history records
            If HistoryArray.Count > 0 Then

                ' Iterate the history array
                For i As Integer = 0 To HistoryArray.Count - 1

                    ' Read an outer history entry
                    strLine = HistoryArray.Item(i).ToString.ToLower

                    ' Check for installed patch header
                    If strLine.ToLower.Contains("ptf wizard installed") Then

                        Try

                            ' Parse patch name
                            PatchName = strLine.ToUpper.Substring(strLine.IndexOf("installed") + 10)

                            ' Trim product code
                            PatchName = PatchName.Substring(0, PatchName.IndexOf("(") - 1)

                            ' Check for a duplicate
                            If PatchArray.Contains(PatchName) Then

                                ' Obtain the index of the first duplicate
                                DuplicateIndex = PatchArray.IndexOf(PatchName)

                                ' Add to duplicate array
                                ' Note: We need to track duplicates in the history file. In the
                                '       case where a patch is applied twice, but removed only
                                '       once, means it should still be a candidate for removal,
                                '       as long as the original files are available in REPLACED.
                                DuplicateArray.Add(PatchArray.Item(DuplicateIndex))
                                DuplicateFileArray.Add(DirectCast(InstalledFileArray.Item(DuplicateIndex), ArrayList).Clone)

                                ' Remove the duplicate patch
                                PatchArray.RemoveAt(DuplicateIndex)
                                InstalledFileArray.RemoveAt(DuplicateIndex)

                            End If

                            ' Add patch name to array
                            PatchArray.Add(PatchName)

                        Catch ex As Exception

                            ' Unexpected format, skip the line
                            Continue For

                        End Try

                        ' Clear the temporary file list
                        TempFileList = New ArrayList

                        ' Perform inner loop for installed files
                        For j As Integer = i + 1 To HistoryArray.Count - 1

                            ' Read an inner history entry
                            strLine = HistoryArray.Item(j).ToString

                            ' Check for INSTALLEDFILE= property
                            If strLine.ToLower.Contains("installedfile=") Then

                                Try

                                    ' Trim the property tag
                                    PatchFileName = strLine.Substring(strLine.IndexOf("=") + 1)

                                    ' Trim the timestamp
                                    PatchFileName = PatchFileName.Substring(0, PatchFileName.IndexOf(","))

                                    ' Add patch filename to list
                                    TempFileList.Add(PatchFileName)

                                Catch ex As Exception

                                    ' Unexpected format, skip it
                                    Continue For

                                End Try

                            End If

                            ' Stop processing inner loop after reaching next patch
                            If strLine.ToLower.Contains("ptf wizard") Then

                                ' We've gone into the next patch, break inner loop
                                Exit For

                            End If

                            ' Increment i (to keep pace with j)
                            i = j

                        Next

                        ' Add file list to patch file array
                        InstalledFileArray.Add(TempFileList)

                    ElseIf strLine.ToLower.Contains("backed out") Then

                        Try

                            ' Parse patch name
                            PatchName = strLine.ToLower.Substring(strLine.IndexOf("backed out") + 11)

                            ' Trim product code
                            PatchName = PatchName.Substring(0, PatchName.IndexOf("(") - 1)

                            ' Reset duplicate flag
                            DuplicateRemoved = False

                            ' Iterate the duplicate array
                            ' Note: Check the duplicate array first. Example: If a patch was
                            '       applied twice, but removed once, the resulting history
                            '       should show one positive apply of the patch, which may
                            '       qualify for removal if the original files are avalable in
                            '       in REPLACED.
                            For k As Integer = 0 To DuplicateArray.Count - 1

                                ' Check for a match
                                If DuplicateArray.Item(k).ToString.ToLower.Equals(PatchName) Then

                                    ' Remove the patch from history
                                    DuplicateArray.RemoveAt(k)
                                    DuplicateFileArray.RemoveAt(k)

                                    ' Set flag
                                    DuplicateRemoved = True

                                    ' Exit loop -- Only one removal
                                    Exit For

                                End If

                            Next

                            ' Check if a match was found from the duplicate pile
                            If Not DuplicateRemoved Then

                                ' Iterate the patch array and remove it
                                For k As Integer = 0 To PatchArray.Count - 1

                                    ' Check for a match
                                    If PatchArray.Item(k).ToString.ToLower.Equals(PatchName) Then

                                        ' Remove the patch from history
                                        PatchArray.RemoveAt(k)
                                        InstalledFileArray.RemoveAt(k)

                                        ' Since removing elements, decrement k
                                        k -= 1

                                        ' Check k for exit condition
                                        If k >= PatchArray.Count - 1 Then Exit For

                                    End If

                                Next

                            End If

                        Catch ex As Exception

                            ' Do nothing
                            Continue For

                        End Try

                    End If

                Next

                ' Iterate the resulting patch array
                For i As Integer = 0 To PatchArray.Count - 1

                    ' Create a node for the patch
                    Dim PatchNode As New TreeNode(PatchArray.Item(i))

                    ' Populate array of Installed Files
                    TempFileList = InstalledFileArray.Item(i)

                    ' Iterate installed files
                    For Each InstalledFile As String In TempFileList

                        ' Create a node for the installed file
                        Dim FileNode As New TreeNode(InstalledFile)

                        ' Add to patch node
                        PatchNode.Nodes.Add(FileNode)

                    Next

                    ' Add patch list, with installed files, to component root node
                    ComponentRootNode.Nodes.Add(PatchNode)

                Next

                ' Expand root node
                ComponentRootNode.Expand()

            Else

                ' Update UI for 'No history'
                Dim NewNode As New TreeNode("No history available.")
                ComponentRootNode.Nodes.Add(NewNode)
                ComponentRootNode.Expand()

            End If

        Else

            ' Update UI for 'No history'
            Dim NewNode As New TreeNode("No history available.")
            ComponentRootNode.Nodes.Add(NewNode)
            ComponentRootNode.Expand()

        End If

        ' Check scenario where history file exists, but is empty
        If ComponentRootNode.Nodes.Count = 0 Then

            ' Update UI for 'No history'
            Dim NewNode As New TreeNode("No history available.")
            ComponentRootNode.Nodes.Add(NewNode)
            ComponentRootNode.Expand()

        End If

    End Sub

End Class