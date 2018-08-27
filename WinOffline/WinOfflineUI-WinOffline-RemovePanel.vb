Imports System.IO
Imports System.Windows.Forms
Imports System.Drawing
Imports WinOffline.WinOffline

Partial Public Class WinOfflineUI

    Private Sub InitRemovePanel()

        ' Populate options list
        CopyListView(lvwApplyOptions, lvwRemoveOptions)

        ' Set initial switch settings
        If Globals.CleanupLogsSwitch Then GetListViewItemByName("CleanupDSMLogsItem", lvwApplyOptions).Checked = True
        If Globals.ResetCftraceSwitch Then GetListViewItemByName("ResetCftraceItem", lvwApplyOptions).Checked = True
        If Globals.RemoveCAMConfigSwitch Then GetListViewItemByName("RemoveCAMConfigItem", lvwApplyOptions).Checked = True
        If Globals.CleanupAgentSwitch Then GetListViewItemByName("CleanupAgentItem", lvwApplyOptions).Checked = True
        If Globals.CleanupServerSwitch Then GetListViewItemByName("CleanupServerItem", lvwApplyOptions).Checked = True
        If Globals.CheckSDLibrarySwitch Then GetListViewItemByName("CheckLibraryItem", lvwApplyOptions).Checked = True
        If Globals.CleanupSDLibrarySwitch Then GetListViewItemByName("CleanLibraryItem", lvwApplyOptions).Checked = True
        If Globals.CleanupCertStoreSwitch Then GetListViewItemByName("CleanupCertItem", lvwApplyOptions).Checked = True
        If Globals.SkipCAFStartUpSwitch Then GetListViewItemByName("SkipCAFStartUpItem", lvwApplyOptions).Checked = True
        If Globals.SkipCAM Then GetListViewItemByName("SkipCAMItem", lvwApplyOptions).Checked = True
        If Globals.SkipDMPrimer Then GetListViewItemByName("SkipDMPrimerItem", lvwApplyOptions).Checked = True
        If Globals.SkiphmAgent Then GetListViewItemByName("SkiphmAgentItem", lvwApplyOptions).Checked = True
        If Globals.LaunchGuiSwitch Then GetListViewItemByName("LaunchGuiItem", lvwApplyOptions).Checked = True
        If Globals.SimulateCafStopSwitch Then GetListViewItemByName("SimulateCafStopItem", lvwApplyOptions).Checked = True
        If Globals.SimulatePatchSwitch Then GetListViewItemByName("SimulatePatchItem", lvwApplyOptions).Checked = True
        If Globals.SimulatePatchErrorSwitch Then GetListViewItemByName("SimulatePatchErrorItem", lvwApplyOptions).Checked = True
        If Globals.RemoveHistoryBeforeSwitch Then GetListViewItemByName("RemoveHistoryBeforeItem", lvwApplyOptions).Checked = True
        If Globals.RemoveHistoryAfterSwitch Then GetListViewItemByName("RemoveHistoryAfterItem", lvwApplyOptions).Checked = True
        If Globals.DumpCazipxpSwitch Then GetListViewItemByName("DumpCazipxpItem", lvwApplyOptions).Checked = True
        If Globals.RegenHostUUIDSwitch Then GetListViewItemByName("RegenHostUUIDItem", lvwApplyOptions).Checked = True
        If Globals.SDSignalRebootSwitch Then GetListViewItemByName("SDSignalRebootItem", lvwApplyOptions).Checked = True
        If Globals.RebootSystemSwitch Then GetListViewItemByName("RebootSystemItem", lvwApplyOptions).Checked = True

        ' Set OptionsListView column resize
        ColumnOptions.AutoResize(ColumnHeaderAutoResizeStyle.ColumnContent)

        ' Recurse history tree
        RecurseHistoryTree()


    End Sub

    Private Sub CopyListView(ByVal ListView1 As ListView, ByVal ListView2 As ListView)

        ' Iterate ListView 1 items
        For Each ListItem As ListViewItem In ListView1.Items

            ' Clone each item
            Dim CloneItem As ListViewItem = ListItem.Clone()

            ' Transcribe the item's name
            CloneItem.Name = ListItem.Name

            ' Iterate the sub-items
            For Each SubItem As ListViewItem.ListViewSubItem In ListItem.SubItems

                ' Create matching sub items
                CloneItem.SubItems.Add(New ListViewItem.ListViewSubItem(CloneItem, SubItem.Text, SubItem.ForeColor, SubItem.BackColor, SubItem.Font))

            Next

            ' Add clone to ListView2
            ListView2.Items.Add(CloneItem)

        Next

    End Sub

    Private Sub RecurseHistoryTree()

        ' Local variables
        Dim OriginalFileFound As Boolean = False
        Dim SupplementalPath As String = ""
        Dim LatestFolder As String = ""
        Dim FolderIncrement As Integer = 0
        Dim RemovalCount As Integer = 0

        ' Iterate each component of the HistoryTree
        For Each Component As TreeNode In treHistory.Nodes

            ' Iterate each patch within each component
            For Each Patch As TreeNode In Component.Nodes

                ' Check for base case -- "No history available"
                If Patch.Text.ToLower.Equals("no history available.") Then

                    ' Break loop, process next component
                    Exit For

                End If

                ' Reset flag
                OriginalFileFound = False

                ' Iterate the installed files within each patch
                ' Note: While the majority of this routine verifies all replaced files reported in the
                '       history file are available in each respective REPLACED folder, this is really
                '       not necessary. It's possible to have a patch that applies new files, in which
                '       case, no original file is saved in the REPLACED folder. Some testing with
                '       ApplyPTF shows having the REPLACED subfolder is sufficient. If some original
                '       files are missing, it doesn't complain, and just deletes the replacement files
                '       outlined in the history file. I'm leaving this routine in WinOffline, as this
                '       might change in the future, but for now, I will blindly delete replacement files
                '       that are suspected to be new.
                For Each InstalledFile As TreeNode In Patch.Nodes

                    ' Reset folder increment
                    FolderIncrement = 0

                    ' Verify a matching original file, to each installed file, by component REPLACED folder
                    If Component.Text.ToLower.Equals("client automation") Then

                        ' Read supplemental path
                        SupplementalPath = InstalledFile.Text.Replace(Globals.DSMFolder, "")

                        ' Build base path to REPLACED folder
                        LatestFolder = Globals.DSMFolder + "REPLACED\" + Patch.Text + ".OLD"

                        ' Verify if base folder exists
                        If System.IO.Directory.Exists(LatestFolder) Then

                            ' Loop dangerously for latest REPLACED folder
                            While True

                                ' Check if incremental folder exists
                                If System.IO.Directory.Exists(Globals.DSMFolder + "REPLACED\" + Patch.Text + "-" + FolderIncrement.ToString + ".OLD") Then

                                    ' Assign latest folder
                                    LatestFolder = Globals.DSMFolder + "REPLACED\" + Patch.Text + "-" + FolderIncrement.ToString + ".OLD"

                                    ' Increase increment
                                    FolderIncrement += 1

                                Else

                                    ' Stop condition
                                    Exit While

                                End If

                            End While

                            ' Check for original file
                            If File.Exists(LatestFolder + "\" + SupplementalPath) Then

                                ' Set flag
                                OriginalFileFound = True

                            Else

                                ' Set flag
                                OriginalFileFound = False

                            End If

                        End If

                    ElseIf Component.Text.ToLower.Equals("shared components") Then

                        ' Read supplemental path
                        SupplementalPath = InstalledFile.Text.Replace(Globals.SharedCompFolder, "")

                        ' Build base path to REPLACED folder
                        LatestFolder = Globals.SharedCompFolder + "REPLACED\" + Patch.Text + ".OLD"

                        ' Verify if base folder exists
                        If System.IO.Directory.Exists(LatestFolder) Then

                            ' Loop dangerously for latest REPLACED folder
                            While True

                                ' Check if incremental folder exists
                                If System.IO.Directory.Exists(Globals.SharedCompFolder + "REPLACED\" + Patch.Text + "-" + FolderIncrement.ToString + ".OLD") Then

                                    ' Assign latest folder
                                    LatestFolder = Globals.SharedCompFolder + "REPLACED\" + Patch.Text + "-" + FolderIncrement.ToString + ".OLD"

                                    ' Increase increment
                                    FolderIncrement += 1

                                Else

                                    ' Stop condition
                                    Exit While

                                End If

                            End While

                            ' Check for original file
                            If File.Exists(LatestFolder + "\" + SupplementalPath) Then

                                ' Set flag
                                OriginalFileFound = True

                            End If

                        End If

                    ElseIf Component.Text.ToLower.Equals("ca message queuing service") Then

                        ' Read supplemental path
                        SupplementalPath = InstalledFile.Text.Replace(Globals.CAMFolder + "\", "")

                        ' Build base path to REPLACED folder
                        LatestFolder = Globals.CAMFolder + "\REPLACED\" + Patch.Text + ".OLD"

                        ' Verify if base folder exists
                        If System.IO.Directory.Exists(LatestFolder) Then

                            ' Loop dangerously for latest REPLACED folder
                            While True

                                ' Check if incremental folder exists
                                If System.IO.Directory.Exists(Globals.CAMFolder + "\REPLACED\" + Patch.Text + "-" + FolderIncrement.ToString + ".OLD") Then

                                    ' Assign latest folder
                                    LatestFolder = Globals.CAMFolder + "\REPLACED\" + Patch.Text + "-" + FolderIncrement.ToString + ".OLD"

                                    ' Increase increment
                                    FolderIncrement += 1

                                Else

                                    ' Stop condition
                                    Exit While

                                End If

                            End While

                            ' Check for original file
                            If File.Exists(LatestFolder + "\" + SupplementalPath) Then

                                ' Set flag
                                OriginalFileFound = True

                            End If

                        End If

                    ElseIf Component.Text.ToLower.Equals("secure socket adapter") Then

                        ' Read supplemental path
                        SupplementalPath = InstalledFile.Text.Replace(Globals.SSAFolder, "")

                        ' Build base path to REPLACED folder
                        LatestFolder = Globals.SSAFolder + "REPLACED\" + Patch.Text + ".OLD"

                        ' Verify if base folder exists
                        If System.IO.Directory.Exists(LatestFolder) Then

                            ' Loop dangerously for latest REPLACED folder
                            While True

                                ' Check if incremental folder exists
                                If System.IO.Directory.Exists(Globals.SSAFolder + "REPLACED\" + Patch.Text + "-" + FolderIncrement.ToString + ".OLD") Then

                                    ' Assign latest folder
                                    LatestFolder = Globals.SSAFolder + "REPLACED\" + Patch.Text + "-" + FolderIncrement.ToString + ".OLD"

                                    ' Increase increment
                                    FolderIncrement += 1

                                Else

                                    ' Stop condition
                                    Exit While

                                End If

                            End While

                            ' Check for original file
                            If File.Exists(LatestFolder + "\" + SupplementalPath) Then

                                ' Set flag
                                OriginalFileFound = True

                            End If

                        End If

                    ElseIf Component.Text.ToLower.Equals("data transport service") Then

                        ' Read supplemental path
                        SupplementalPath = InstalledFile.Text.Replace(Globals.DTSFolder + "\", "")

                        ' Build base path to REPLACED folder
                        LatestFolder = Globals.DTSFolder + "\REPLACED\" + Patch.Text + ".OLD"

                        ' Verify if base folder exists
                        If System.IO.Directory.Exists(LatestFolder) Then

                            ' Loop dangerously for latest REPLACED folder
                            While True

                                ' Check if incremental folder exists
                                If System.IO.Directory.Exists(Globals.DTSFolder + "\REPLACED\" + Patch.Text + "-" + FolderIncrement.ToString + ".OLD") Then

                                    ' Assign latest folder
                                    LatestFolder = Globals.DTSFolder + "\REPLACED\" + Patch.Text + "-" + FolderIncrement.ToString + ".OLD"

                                    ' Increase increment
                                    FolderIncrement += 1

                                Else

                                    ' Stop condition
                                    Exit While

                                End If

                            End While

                            ' Check for original file
                            If File.Exists(LatestFolder + "\" + SupplementalPath) Then

                                ' Set flag
                                OriginalFileFound = True

                            End If

                        End If

                    ElseIf Component.Text.ToLower.Equals("dsm explorer gui") Then

                        ' Read supplemental path
                        SupplementalPath = InstalledFile.Text.Replace(Globals.EGCFolder, "")

                        ' Build base path to REPLACED folder
                        LatestFolder = Globals.EGCFolder + "REPLACED\" + Patch.Text + ".OLD"

                        ' Verify if base folder exists
                        If System.IO.Directory.Exists(LatestFolder) Then

                            ' Loop dangerously for latest REPLACED folder
                            While True

                                ' Check if incremental folder exists
                                If System.IO.Directory.Exists(Globals.EGCFolder + "REPLACED\" + Patch.Text + "-" + FolderIncrement.ToString + ".OLD") Then

                                    ' Assign latest folder
                                    LatestFolder = Globals.EGCFolder + "REPLACED\" + Patch.Text + "-" + FolderIncrement.ToString + ".OLD"

                                    ' Increase increment
                                    FolderIncrement += 1

                                Else

                                    ' Stop condition
                                    Exit While

                                End If

                            End While

                            ' Check for original file
                            If File.Exists(LatestFolder + "\" + SupplementalPath) Then

                                ' Set flag
                                OriginalFileFound = True

                            End If

                        End If

                    End If

                    ' If any original file is missing, break.
                    If OriginalFileFound = False Then Exit For

                Next

                ' All originals are availalbe -- This is good.
                If OriginalFileFound Then

                    ' Create a list view item
                    Dim RemovalListItem As ListViewItem = lvwHistory.Items.Add(Patch.Text)

                    ' Add sub-items
                    RemovalListItem.SubItems.Add(Component.Text)
                    RemovalListItem.SubItems.Add("OK for removal [" + LatestFolder.Substring(LatestFolder.LastIndexOf("\") + 1) + "]")

                    ' Iterate removal manifest
                    For i As Integer = 0 To Manifest.RemovalManifestCount - 1

                        ' Get a removal vector
                        Dim rVector As RemovalVector = Manifest.GetRemovalFromManifest(i)

                        ' Check for a match
                        If rVector.RemovalItem.ToLower.Equals(Patch.Text.ToLower) Then

                            ' Tick the removal box
                            RemovalListItem.Checked = True

                        End If

                    Next

                ElseIf OriginalFileFound = False And System.IO.Directory.Exists(LatestFolder) Then

                    ' Note: Refer to the note at the top of this routine. Although not all files
                    '       identified by the history file match the contents saved in the REPLACED
                    '       subfolder, we will allow removal of the patch, based purely on the
                    '       existence of the matching REPLACED subfolder. Any file not backed up
                    '       in REPLACED will assume to have been a new file, introduced when the
                    '       patch was applied, and therefore will be discarded.
                    '
                    '       Though I'm not throwing away the code that checks for all this ;)
                    '       Perhaps it will have some future use to separate these two scenarios.

                    ' Create a list view item
                    Dim RemovalListItem As ListViewItem = lvwHistory.Items.Add(Patch.Text)

                    ' Add sub-items
                    RemovalListItem.SubItems.Add(Component.Text)
                    RemovalListItem.SubItems.Add("OK for removal [" + LatestFolder.Substring(LatestFolder.LastIndexOf("\") + 1) + "]")

                    ' Iterate removal manifest
                    For i As Integer = 0 To Manifest.RemovalManifestCount - 1

                        ' Get a removal vector
                        Dim rVector As RemovalVector = Manifest.GetRemovalFromManifest(i)

                        ' Check for a match
                        If rVector.RemovalItem.ToLower.Equals(Patch.Text.ToLower) Then

                            ' Tick the removal box
                            RemovalListItem.Checked = True

                        End If

                    Next

                Else

                    ' Create a list view item
                    Dim RemovalListItem As ListViewItem = lvwHistory.Items.Add(Patch.Text)

                    ' Add sub-items
                    RemovalListItem.SubItems.Add(Component.Text)
                    RemovalListItem.SubItems.Add("Missing REPLACED subfolder for removal.")

                    ' Disable selection
                    RemovalListItem.ForeColor = SystemColors.GrayText

                End If

                ' Imcrement patch count
                RemovalCount += 1

            Next

        Next

        ' Resize the columns, if any removal history is available
        If RemovalCount > 0 Then

            ' Set PatchListView column resize
            ColumnHistoryPatch.AutoResize(ColumnHeaderAutoResizeStyle.ColumnContent)
            ColumnHistoryComp.AutoResize(ColumnHeaderAutoResizeStyle.ColumnContent)
            ColumnHistoryStatus.AutoResize(ColumnHeaderAutoResizeStyle.ColumnContent)

        Else

            ' Disable the removal start button
            Delegate_Sub_Enable_Remove_Start_Button(False)

        End If

    End Sub

    Private Sub HistoryListView_ItemCheck(sender As Object, e As ItemCheckEventArgs) Handles lvwHistory.ItemCheck

        ' For any other status than 'Ok to remove.', restrict the checkbox.
        If Not lvwHistory.Items(e.Index).SubItems(2).Text.ToLower.Contains("ok for removal") Then

            ' Keep current checkbox value (which is unchecked)
            e.NewValue = e.CurrentValue

        End If

    End Sub

    Private Sub HistoryListView_ItemChecked(sender As Object, e As ItemCheckedEventArgs) Handles lvwHistory.ItemChecked

        ' Disable the removal start button
        Delegate_Sub_Enable_Remove_Start_Button(False)

        ' Iterate HistoryListView items
        For Each ListItem As ListViewItem In lvwHistory.Items

            ' If any item is checked, enable the start button
            If ListItem.Checked Then Delegate_Sub_Enable_Remove_Start_Button(True)

        Next

    End Sub

End Class