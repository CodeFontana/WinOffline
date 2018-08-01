'****************************** Class Header *******************************\
' Project Name: WinOffline
' Class Name:   WinOfflineUI
' File Name:    WinOfflineUI-WinOffline-MainPanel.vb
' Author:       Brian Fontana
'***************************************************************************/

Imports System.Threading
Imports System.Windows.Forms
Imports WinOffline.WinOffline

Partial Public Class WinOfflineUI

    Private Sub InitMainPanel()

        ' Local variables
        Dim Runlevel As Integer = 0

        ' Create advanced option list view items
        Dim CleanupDSMLogsItem As ListViewItem = lvwApplyOptions.Items.Add("-cleanlogs: Cleanup DSM logs folder.")
        Dim ResetCftraceItem As ListViewItem = lvwApplyOptions.Items.Add("-resetcftrace: Reset cftrace level.")
        Dim RemoveCAMConfigItem As ListViewItem = lvwApplyOptions.Items.Add("-rmcamcfg: Remove CAM configuration file.")
        Dim CleanupAgentItem As ListViewItem = lvwApplyOptions.Items.Add("-cleanagent: Perform agent cleanup.")
        Dim CleanupServerItem As ListViewItem = lvwApplyOptions.Items.Add("-cleanserver: Perform scalability server cleanup.")
        Dim CheckLibraryItem As ListViewItem = lvwApplyOptions.Items.Add("-checklibrary: Analyze software library without making changes.")
        Dim CleanLibraryItem As ListViewItem = lvwApplyOptions.Items.Add("-cleanlibrary: Perform software library cleanup.")
        Dim CleanupCertStoreItem As ListViewItem = lvwApplyOptions.Items.Add("-cleancerts: Certificate store cleanup.")
        Dim SkipCAFStartUpItem As ListViewItem = lvwApplyOptions.Items.Add("-skipcafstartup: Skip CAF startup.")
        Dim SkipCAMItem As ListViewItem = lvwApplyOptions.Items.Add("-skipcam: Don't stop CAM service.")
        Dim SkipDMPrimerItem As ListViewItem = lvwApplyOptions.Items.Add("-skipdmprimer: Don't stop DMPrimer service.")
        Dim SkiphmAgentItem As ListViewItem = lvwApplyOptions.Items.Add("-skiphm: Don't stop hmAgent.")
        Dim LaunchGuiItem As ListViewItem = lvwApplyOptions.Items.Add("-loadgui: Launch DSM Explorer.")
        Dim SimulateCafStopItem As ListViewItem = lvwApplyOptions.Items.Add("-simulatestop: Simulate recycling CAF.")
        Dim SimulatePatchItem As ListViewItem = lvwApplyOptions.Items.Add("-simulatepatch: Simulate patch operations.")
        Dim SimulatePatchErrorItem As ListViewItem = lvwApplyOptions.Items.Add("-simulatepatcherror: Simulate a patching error.")
        Dim RemoveHistoryBeforeItem As ListViewItem = lvwApplyOptions.Items.Add("-rmhistorybefore: Remove patch history file, BEFORE patch operations.")
        Dim RemoveHistoryAfterItem As ListViewItem = lvwApplyOptions.Items.Add("-rmhistoryafter: Remove patch history file, AFTER patch operations.")
        Dim DumpCazipxpItem As ListViewItem = lvwApplyOptions.Items.Add("-dumpcazipxp: Extract cazipxp.exe utility and exit without changes.")
        Dim RegenHostUUIDItem As ListViewItem = lvwApplyOptions.Items.Add("-regenuuid: Regenerate HostUUID.")
        Dim SDSignalRebootItem As ListViewItem = lvwApplyOptions.Items.Add("-signalreboot: Signal software delivery reboot request.")
        Dim RebootSystemItem As ListViewItem = lvwApplyOptions.Items.Add("-reboot: Reboot system after completion.")

        ' Set list view item names
        CleanupDSMLogsItem.Name = "CleanupDSMLogsItem"
        ResetCftraceItem.Name = "ResetCftraceItem"
        RemoveCAMConfigItem.Name = "RemoveCAMConfigItem"
        CleanupAgentItem.Name = "CleanupAgentItem"
        CleanupServerItem.Name = "CleanupServerItem"
        CheckLibraryItem.Name = "CheckLibraryItem"
        CleanLibraryItem.Name = "CleanLibraryItem"
        CleanupCertStoreItem.Name = "CleanupCertItem"
        SkipCAFStartUpItem.Name = "SkipCAFStartUpItem"
        SkipCAMItem.Name = "SkipCAMItem"
        SkipDMPrimerItem.Name = "SkipDMPrimerItem"
        SkiphmAgentItem.Name = "SkiphmAgentItem"
        LaunchGuiItem.Name = "LaunchGuiItem"
        SimulateCafStopItem.Name = "SimulateCafStopItem"
        SimulatePatchItem.Name = "SimulatePatchItem"
        SimulatePatchErrorItem.Name = "SimulatePatchErrorItem"
        RemoveHistoryBeforeItem.Name = "RemoveHistoryBeforeItem"
        RemoveHistoryAfterItem.Name = "RemoveHistoryAfterItem"
        DumpCazipxpItem.Name = "DumpCazipxpItem"
        RegenHostUUIDItem.Name = "RegenHostUUIDItem"
        SDSignalRebootItem.Name = "SDSignalRebootItem"
        RebootSystemItem.Name = "RebootSystemItem"

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

        ' Update group label
        grpPatchView.Text = "Working Directory: " + Globals.ProcessFilePath

        ' Spin off thread for updating the PatchListView
        PatchListViewThread = New Thread(AddressOf PatchScanThread)
        PatchListViewThread.Start()

    End Sub

    Private Sub PatchScanThread()

        ' Local variables
        Dim FileList As String()
        Dim PreviousFileList As String()
        Dim PatchFileName As String         ' Name of patch file (i.e. C:\SomeDirectory\SomePatch.caz).
        Dim PatchPath As String             ' Path to patch file (i.e. C:\SomeDirectory).
        Dim PatchShortName As String        ' Name of patch file (i.e. SomePatch.caz).
        Dim PatchFriendlyName As String     ' Name of patch file (i.e. SomePatch).
        Dim PatchManifest As New ArrayList
        Dim LoopCounter As Integer = 0
        Dim PatchErrorDetected As Boolean
        Dim Runlevel As Integer = 0

        ' Set initial value
        PreviousFileList = {}

        Try

            ' Patch list view will refresh for duration of WinOfflineUI lifecycle
            While (Not TerminateSignal)

                ' Reset patch error flag
                PatchErrorDetected = False

                ' Get a directory listing
                FileList = System.IO.Directory.GetFiles(Globals.WorkingDirectory)

                ' Check if listing has changed
                If Not WinOffline.Utility.IsArrayEqual(FileList, PreviousFileList) Then

                    ' Clear PatchListView
                    Delegate_Sub_Clear_ListView(lvwPatchList)

                    ' Clear the patch manifest
                    PatchManifest.Clear()

                    ' Iterate the file list
                    For n As Integer = 0 To FileList.Length - 1

                        ' Check for a patch file
                        If FileList(n).ToLower.EndsWith(".caz") OrElse FileList(n).ToLower.EndsWith(".jcl") Then

                            ' Get all versions of the filename
                            PatchFileName = FileList(n)
                            PatchPath = FileVector.GetFilePath(PatchFileName)
                            PatchShortName = FileVector.GetShortName(PatchFileName)
                            PatchFriendlyName = FileVector.GetFriendlyName(PatchFileName)

                            ' Check for duplicate patch
                            If PatchManifest.Contains(PatchFriendlyName.ToLower) Then

                                ' Update result
                                Delegate_Sub_Set_Patch_ListView(lvwPatchList, New ArrayList({PatchShortName,
                                                                                   "N/A",
                                                                                   "Duplicate"}))

                                ' Continue loop
                                Continue For

                            End If

                            ' Check if patch is compressed
                            If PatchShortName.ToLower.EndsWith(".caz") Then

                                ' Decompress the CAZ file
                                Runlevel = DecompressPatch(PatchFileName)

                                ' Check result
                                If Runlevel = 1 Then

                                    ' Update result
                                    Delegate_Sub_Set_Patch_ListView(lvwPatchList, New ArrayList({PatchShortName,
                                                                                   "N/A",
                                                                                   "Failed to create " + Globals.WinOfflineTemp + "\UITemp."}))

                                    ' Set patch error flag
                                    PatchErrorDetected = True

                                ElseIf Runlevel = 2 Then

                                    ' Update result
                                    Delegate_Sub_Set_Patch_ListView(lvwPatchList, New ArrayList({PatchShortName,
                                                                                   "N/A",
                                                                                   "Decompression failed"}))

                                    ' Set patch error flag
                                    PatchErrorDetected = True

                                ElseIf Runlevel = 3 Then

                                    ' Update result
                                    Delegate_Sub_Set_Patch_ListView(lvwPatchList, New ArrayList({PatchShortName,
                                                                                   "N/A",
                                                                                   "Missing JCL file"}))

                                    ' Set patch error flag
                                    PatchErrorDetected = True

                                ElseIf Runlevel = 4 Then

                                    ' Update result
                                    Delegate_Sub_Set_Patch_ListView(lvwPatchList, New ArrayList({PatchShortName,
                                                                                   "N/A",
                                                                                   "Multiple JCL files"}))

                                    ' Set patch error flag
                                    PatchErrorDetected = True

                                ElseIf Runlevel <> 0 Then

                                    ' Update result
                                    Delegate_Sub_Set_Patch_ListView(lvwPatchList, New ArrayList({PatchShortName,
                                                                                   "N/A",
                                                                                   "Unknown error"}))

                                Else

                                    ' Adjust file information -- use decompressed patch
                                    PatchFileName = Globals.WinOfflineTemp + "\UITemp\" + PatchShortName
                                    PatchPath = FileVector.GetFilePath(PatchFileName)
                                    PatchShortName = FileVector.GetShortName(PatchFileName)
                                    PatchFriendlyName = FileVector.GetFriendlyName(PatchFileName)

                                End If

                                ' Check for any decompression failure
                                If Runlevel <> 0 Then

                                    ' Cleanup temporary space
                                    System.IO.Directory.Delete(Globals.WinOfflineTemp + "\UITemp", True)

                                    ' Continue loop
                                    Continue For

                                End If

                            End If

                            ' Evaluate patch
                            Runlevel = EvaluatePatch(PatchFileName)

                            ' Check result
                            If Runlevel = 0 Then

                                ' Check for versioncheck tag
                                If Not GetPatchVector(PatchFileName).GetInstruction("VERSIONCHECK").Equals("") Then

                                    ' Update result -- display versioncheck confirmation
                                    Delegate_Sub_Set_Patch_ListView(lvwPatchList, New ArrayList({PatchShortName,
                                                                                        GetPatchVector(PatchFileName).GetInstruction("PRODUCT"),
                                                                                        "READY [VERSIONCHECK=" + GetPatchVector(PatchFileName).GetInstruction("VERSIONCHECK") + "]"}))

                                Else

                                    ' Update result
                                    Delegate_Sub_Set_Patch_ListView(lvwPatchList, New ArrayList({PatchShortName,
                                                                                        GetPatchVector(PatchFileName).GetInstruction("PRODUCT"),
                                                                                        "READY"}))

                                End If

                                ' Add to manifest
                                PatchManifest.Add(PatchFriendlyName.ToLower)

                            ElseIf Runlevel = 1 Then

                                ' Update result
                                Delegate_Sub_Set_Patch_ListView(lvwPatchList, New ArrayList({PatchShortName,
                                                                                "N/A",
                                                                                "Failed to open JCL file"}))

                                ' Set patch error flag
                                PatchErrorDetected = True

                            ElseIf Runlevel = 2 Then

                                ' Update result
                                Delegate_Sub_Set_Patch_ListView(lvwPatchList, New ArrayList({PatchShortName,
                                                                                "N/A",
                                                                                "JCL missing PRODUCT tag"}))

                                ' Set patch error flag
                                PatchErrorDetected = True

                            ElseIf Runlevel = 3 Then

                                ' Update result
                                Delegate_Sub_Set_Patch_ListView(lvwPatchList, New ArrayList({PatchShortName,
                                                                                GetPatchVector(PatchFileName).GetInstruction("PRODUCT"),
                                                                                "Unsupported product code"}))

                                ' Set patch error flag
                                PatchErrorDetected = True

                            ElseIf Runlevel = 4 Then

                                ' Update result
                                Delegate_Sub_Set_Patch_ListView(lvwPatchList, New ArrayList({PatchShortName,
                                                                                "N/A",
                                                                                "JCL missing instructions"}))

                                ' Set patch error flag
                                PatchErrorDetected = True

                            ElseIf Runlevel = 5 Then

                                ' Update result
                                Delegate_Sub_Set_Patch_ListView(lvwPatchList, New ArrayList({PatchShortName,
                                                                                GetPatchVector(PatchFileName).GetInstruction("PRODUCT"),
                                                                                "JCL missing dependencies"}))

                                ' Set patch error flag
                                PatchErrorDetected = True

                            ElseIf Runlevel = 6 Then

                                ' Update result
                                Delegate_Sub_Set_Patch_ListView(lvwPatchList, New ArrayList({PatchShortName,
                                                                                GetPatchVector(PatchFileName).GetInstruction("PRODUCT"),
                                                                                "Product component (" + GetPatchVector(PatchFileName).GetInstruction("PRODUCT") + ") is not installed."}))

                                ' Set patch error flag
                                PatchErrorDetected = True

                            ElseIf Runlevel = 7 Then

                                ' Update result
                                Delegate_Sub_Set_Patch_ListView(lvwPatchList, New ArrayList({PatchShortName,
                                                                                GetPatchVector(PatchFileName).GetInstruction("PRODUCT"),
                                                                                "Not applicable [VERSIONCHECK=" + GetPatchVector(PatchFileName).GetInstruction("VERSIONCHECK") + "]"}))

                            ElseIf Runlevel = 8 Then

                                ' Update result
                                Delegate_Sub_Set_Patch_ListView(lvwPatchList, New ArrayList({PatchShortName,
                                                                                GetPatchVector(PatchFileName).GetInstruction("PRODUCT"),
                                                                                "Not applicable."}))

                            ElseIf Runlevel = 9 Then

                                ' Update result
                                Delegate_Sub_Set_Patch_ListView(lvwPatchList, New ArrayList({PatchShortName,
                                                                                GetPatchVector(PatchFileName).GetInstruction("PRODUCT"),
                                                                                "Already applied."}))

                            End If

                            ' Check for UI temp cleanup
                            If FileList(n).ToLower.EndsWith(".caz") Then

                                ' Cleanup temporary space
                                System.IO.Directory.Delete(Globals.WinOfflineTemp + "\UITemp", True)

                            End If

                        End If

                    Next

                    ' Save file listing
                    PreviousFileList = FileList.Clone

                    ' Enable or disable start button
                    Delegate_Sub_Enable_Start_Button(Not PatchErrorDetected)

                End If

                ' Reset LoopCounter
                LoopCounter = 0

                ' Check for termination signal every half second.
                ' But only run an update on the working directory every 5 seconds.
                While LoopCounter < 99

                    ' Check for termination signal
                    If TerminateSignal Then Return

                    ' Increment loop counter
                    LoopCounter += 1

                    ' Rest the thread
                    Thread.Sleep(Globals.THREAD_REST_INTERVAL)

                End While

            End While

        Catch ex As Exception

            ' Write debug
            Delegate_Sub_Append_Text(rtbDebug, "PatchListViewThread --> " + ex.Message)
            Delegate_Sub_Append_Text(rtbDebug, ex.StackTrace)

        End Try

    End Sub

    Private Shared Function DecompressPatch(ByVal CAZFileName As String) As Integer

        ' Local variables
        Dim DestShortName As String                         ' Shortname for destination filename (i.e. SomePatch.caz).
        Dim CAZDestPath As String                           ' Destination path after decompression (i.e. C:\Windows\Temp\WinOffline\SomePatch).
        Dim DestFileName As String                          ' Destination filename after decompression (i.e. C:\Windows\Temp\WinOffline\SomePatch\SomePatch.caz).
        Dim ExecutionString As String                       ' Command line to be executed externally to the application.
        Dim ArgumentString As String                        ' Arguments passed on the command line for the external execution.
        Dim CAZipXPProcessStartInfo As ProcessStartInfo     ' Process startup settings for configuring the bahviour of the process.
        Dim RunningProcess As Process                       ' A process shell for executing the command line.
        Dim ExitCode As Integer                             ' External process return code.
        Dim FileList As String()                            ' Directory listing of files.
        Dim JCLShortName As String                          ' Name of jcl file found in caz (i.e. SomePatch.jcl)
        Dim JCLCounter As Integer = 0                       ' Counter of JCL files.

        ' Calculate names
        DestShortName = FileVector.GetShortName(CAZFileName)
        CAZDestPath = Globals.WinOfflineTemp + "\UITemp"
        DestFileName = CAZDestPath + "\" + DestShortName

        ' Prepare decompression destination
        Try

            ' Create patch directory
            System.IO.Directory.CreateDirectory(CAZDestPath)

            ' Copy patch files to directory
            System.IO.File.Copy(CAZFileName, DestFileName, True)

        Catch ex As Exception

            ' Return -- Unable to create UI temporary directory
            Return 1

        End Try

        ' Build cazipxp execution string
        ExecutionString = Globals.WinOfflineTemp + "\cazipxp.exe"
        ArgumentString = "-u " + """" + DestFileName + """"

        ' Create detached process for cazipxp
        CAZipXPProcessStartInfo = New ProcessStartInfo(ExecutionString, ArgumentString)
        CAZipXPProcessStartInfo.WorkingDirectory = CAZDestPath
        CAZipXPProcessStartInfo.UseShellExecute = False
        CAZipXPProcessStartInfo.RedirectStandardOutput = True
        CAZipXPProcessStartInfo.CreateNoWindow = True

        ' Start detached process
        RunningProcess = Process.Start(CAZipXPProcessStartInfo)

        ' Wait for detached process to exit
        RunningProcess.WaitForExit()

        ' Get the exitcode
        ExitCode = RunningProcess.ExitCode

        ' Close detached process
        RunningProcess.Close()

        ' Check cazipxp exit code
        If ExitCode <> 0 Then

            ' Return -- Failed to decompress .caz file
            Return 2

        End If

        ' Get the patch directory file listing
        FileList = System.IO.Directory.GetFiles(CAZDestPath)

        ' Loop the file list -- Check and decompress any CAZ files
        For n As Integer = 0 To FileList.Length - 1

            ' Check for a JCL file
            If FileList(n).ToLower.Contains(".jcl") Then

                ' Read the JCL name
                JCLShortName = FileVector.GetShortName(FileList(n))

                ' Increment the counter
                JCLCounter = JCLCounter + 1

            End If

        Next

        ' Check the JCL counter
        If JCLCounter = 0 Then

            ' Return -- No JCL files found.
            Return 3

        ElseIf JCLCounter > 1 Then

            ' Return -- Too many JCL files found.
            Return 4

        End If

        ' Return
        Return 0

    End Function

    Public Shared Function EvaluatePatch(ByVal PatchFileName As String) As Integer

        ' Local variables
        Dim FileList As String()
        Dim PatchStreamReader As System.IO.StreamReader
        Dim InstructionList As New ArrayList
        Dim NewPatch As PatchVector
        Dim strLine As String
        Dim DependentFlag As Boolean = False
        Dim AlreadyAppliedList As New List(Of Boolean)

        ' Check for .caz file (retrieve the .jcl file for evaluation)
        If PatchFileName.ToLower.EndsWith(".caz") Then

            ' Get the patch directory file listing
            FileList = System.IO.Directory.GetFiles(FileVector.GetFilePath(PatchFileName))

            ' Loop the file list -- find the .jcl file
            For n As Integer = 0 To FileList.Length - 1

                ' Match the JCL file
                If FileList(n).ToLower.Contains(".jcl") Then

                    ' Assign .jcl file
                    PatchFileName = FileList(n)

                End If

            Next

        End If

        ' Open the JCL file
        Try

            ' Open the .jcl reader stream
            PatchStreamReader = New System.IO.StreamReader(PatchFileName)

        Catch ex As Exception

            ' Return -- Unable open JCL file
            Return 1

        End Try

        ' Loop patch file contents
        Do While PatchStreamReader.Peek() >= 0

            ' Read a line
            strLine = PatchStreamReader.ReadLine()

            ' Update the array
            InstructionList.Add(strLine)

        Loop

        ' Close the patch stream reader
        PatchStreamReader.Close()

        ' Create a new patch object
        NewPatch = New PatchVector(PatchFileName, InstructionList)

        ' Verify a product code is specified
        If NewPatch.GetInstruction("product").Equals("") Then

            ' Return -- JCL missing PRODUCT tag
            Return 2

        End If

        ' Verify patch contains a supported product code
        If Not (NewPatch.IsClientAuto Or
                NewPatch.IsSharedComponent Or
                NewPatch.IsCAM Or
                NewPatch.IsSSA Or
                NewPatch.IsDataTransport Or
                NewPatch.IsExplorerGUI) Then

            ' Return -- Unsupported product code
            Return 3

        End If

        ' Verify the patch contains at least one valid instruction
        If NewPatch.GetShortNameReplaceList.Count +
            NewPatch.GetPreCommandList.Count +
            NewPatch.GetSysCommandList.Count = 0 Then

            ' Return -- Patch contains NO instructions 
            Return 4

        End If

        ' Iterate the file replacement list
        For Each Dependent As String In NewPatch.GetAllDependentFiles

            ' Reset dependency flag
            DependentFlag = False

            ' List files from the patch folder
            FileList = System.IO.Directory.GetFiles(NewPatch.PatchFile.GetFilePath)

            ' Loop the file list -- Resolve file dependency
            For n As Integer = 0 To FileList.Length - 1

                ' Check for a match
                If FileList(n).ToLower.Contains(Dependent.ToLower) Then

                    ' Set flag
                    DependentFlag = True

                    ' Break loop
                    Exit For

                End If

            Next

            ' Dependent not found
            If DependentFlag = False Then

                ' Return -- Missing dependent files
                Return 5

            End If

        Next

        ' Verify the patch component is installed
        If (NewPatch.IsClientAuto AndAlso Globals.DSMFolder Is Nothing) Or
            (NewPatch.IsSharedComponent AndAlso Globals.SharedCompFolder Is Nothing) Or
            (NewPatch.IsCAM AndAlso Globals.CAMFolder Is Nothing) Or
            (NewPatch.IsSSA AndAlso Globals.SSAFolder Is Nothing) Or
            (NewPatch.IsDataTransport AndAlso Globals.DTSFolder Is Nothing) Or
            (NewPatch.IsExplorerGUI AndAlso Globals.EGCFolder Is Nothing) Then

            ' Return -- patch component is not installed
            Return 6

        End If

        ' Check for VERSIONCHECK tag
        If Not NewPatch.GetInstruction("VERSIONCHECK").Equals("") AndAlso Not NewPatch.GetInstruction("VERSIONCHECK").Equals(Globals.ITCMComstoreVersion) Then

            ' Return -- VERSIONCHECK mismatch
            Return 7

        End If

        ' Check for file replacement operations
        If NewPatch.SourceReplaceList.Count > 0 Then

            ' Iterate file replacements
            For x As Integer = 0 To NewPatch.SourceReplaceList.Count - 1

                ' Check for unequal replacement file
                If Not Utility.IsFileEqual(NewPatch.SourceReplaceList.Item(x).ToString, NewPatch.DestReplaceList.Item(x).ToString) Then

                    ' Verify unequal file is not a new file vs skippable file
                    If Not System.IO.File.Exists(NewPatch.DestReplaceList.Item(x).ToString) AndAlso
                        NewPatch.SkipIfNotFoundList.Contains(FileVector.GetShortName(NewPatch.DestReplaceList.Item(x).ToString.ToLower)) Then

                        ' Skipped file -- no applicability information

                    ElseIf Not System.IO.File.Exists(NewPatch.DestReplaceList.Item(x).ToString) AndAlso
                        Not NewPatch.SkipIfNotFoundList.Contains(FileVector.GetShortName(NewPatch.DestReplaceList.Item(x).ToString.ToLower)) Then

                        ' New file -- not applied
                        AlreadyAppliedList.Add(False)

                    Else

                        ' Complete mismatch -- not applied
                        AlreadyAppliedList.Add(False)

                    End If

                Else

                    ' File match -- already applied
                    AlreadyAppliedList.Add(True)

                End If

            Next

            ' Check already applied list
            If AlreadyAppliedList.Count = 0 Then

                ' Return -- not applicable
                Return 8

            ElseIf Not AlreadyAppliedList.Contains(False) Then

                ' Return -- already applied
                Return 9

            End If

        End If

        ' Return
        Return 0

    End Function

    Public Shared Function GetPatchVector(ByVal PatchFileName As String) As PatchVector

        ' Local variables
        Dim FileList As String()
        Dim PatchStreamReader As System.IO.StreamReader
        Dim InstructionList As New ArrayList
        Dim strLine As String

        ' Check for .caz file (retrieve the .jcl file for evaluation)
        If PatchFileName.ToLower.EndsWith(".caz") Then

            ' Get the patch directory file listing
            FileList = System.IO.Directory.GetFiles(FileVector.GetFilePath(PatchFileName))

            ' Loop the file list -- find the .jcl file
            For n As Integer = 0 To FileList.Length - 1

                ' Match the JCL file
                If FileList(n).ToLower.Contains(".jcl") Then

                    ' Assign .jcl file
                    PatchFileName = FileList(n)

                End If

            Next

        End If

        ' Open the JCL file
        Try

            ' Open the .jcl reader stream
            PatchStreamReader = New System.IO.StreamReader(PatchFileName)

        Catch ex As Exception

            ' Return -- Unable open JCL file
            Return Nothing

        End Try

        ' Loop patch file contents
        Do While PatchStreamReader.Peek() >= 0

            ' Read a line
            strLine = PatchStreamReader.ReadLine()

            ' Update the array
            InstructionList.Add(strLine)

        Loop

        ' Close the patch stream reader
        PatchStreamReader.Close()

        ' Return
        Return New PatchVector(PatchFileName, InstructionList)

    End Function

    Private Function GetListViewItemByName(ByVal ItemName As String, ByVal oListView As ListView) As ListViewItem
        For Each item As ListViewItem In oListView.Items
            If item.Name.Equals(ItemName) Then Return item
        Next
        Return Nothing
    End Function

    Private Sub xListView_ItemChecked(sender As Object, e As ItemCheckedEventArgs) Handles lvwApplyOptions.ItemChecked, lvwRemoveOptions.ItemChecked

        ' Remove the check handlers during mirror operation
        RemoveHandler lvwApplyOptions.ItemChecked, AddressOf xListView_ItemChecked
        RemoveHandler lvwRemoveOptions.ItemChecked, AddressOf xListView_ItemChecked

        ' Check the sender, and mirror the check configuration to the other ListView
        If sender.Equals(lvwApplyOptions) Then
            For Each ListItem As ListViewItem In lvwRemoveOptions.Items
                If ListItem.Name.Equals(e.Item.Name) Then
                    ListItem.Checked = e.Item.Checked
                End If
            Next
        ElseIf sender.Equals(lvwRemoveOptions) Then
            For Each ListItem As ListViewItem In lvwApplyOptions.Items
                If ListItem.Name.Equals(e.Item.Name) Then
                    ListItem.Checked = e.Item.Checked
                End If
            Next
        End If

        ' Resume check handlers
        AddHandler lvwApplyOptions.ItemChecked, AddressOf xListView_ItemChecked
        AddHandler lvwRemoveOptions.ItemChecked, AddressOf xListView_ItemChecked

    End Sub

End Class