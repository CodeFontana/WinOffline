'****************************** Class Header *******************************\
' Project Name: WinOffline
' Class Name:   WinOffline
' File Name:    WinOffline.vb
' Author:       Brian Fontana
'***************************************************************************/

Public Class WinOffline

    Public Shared Function Main(ByVal CommandLineArguments As String()) As Integer

        ' Attach debugger
        If Utility.StringArrayContains(CommandLineArguments, "attachdebug") Then MsgBox("Attach debugger.", MsgBoxStyle.Information, "Attach debugger")

        ' Local variables
        Dim CallStack As String = "|Main|"
        Dim RunLevel As Integer = 0

        ' Write debug
        Logger.SetCurrentTask("Starting up..")

        ' Obtain process globals
        Globals.ProcessName = Process.GetCurrentProcess.MainModule.FileName
        Globals.ProcessFilePath = FileVector.GetFilePath(Globals.ProcessName)
        Globals.ProcessShortName = FileVector.GetShortName(Globals.ProcessName)
        Globals.ProcessFriendlyName = FileVector.GetFriendlyName(Globals.ProcessName)
        Globals.DotNetVersion = Environment.Version.ToString
        Globals.CommandLineArgs = CommandLineArguments

        ' Attach to console (if available), and set flag
        If WindowsAPI.AttachConsole(-1) Then Globals.AttachedtoConsole = True Else Globals.AttachedtoConsole = False

        ' *****************************
        ' - Main process initialization.
        ' *****************************

        ' Initialize
        RunLevel = Init.Init(CallStack)

        ' Check the run level
        If RunLevel <> 0 Then

            ' Deinitialize
            Init.DeInit(CallStack, True, False, True, False)

            ' Return [1-10]
            Return RunLevel

        End If

        ' *****************************
        ' - Check for removal tool execution.
        ' *****************************

        ' Check removal tool switches
        If Globals.RemoveITCM OrElse Globals.UninstallITCM Then

            ' Remove ITCM
            RunLevel = RemoveITCM(CallStack)

            ' Deinitialize
            Init.DeInit(CallStack, False, False, True, False)

            ' Return
            Return RunLevel

        End If

        ' *****************************
        ' - Check for caf on-demand switches.
        ' *****************************

        ' Check for caf on-demand switches
        If Globals.StopCAFSwitch Then

            ' Remove ITCM
            RunLevel = StopCAFOnDemand(CallStack)

            ' Deinitialize
            Init.DeInit(CallStack, False, False, True, False)

            ' Return
            Return RunLevel

        ElseIf Globals.StartCAFSwitch Then

            ' Remove ITCM
            RunLevel = StartCAFOnDemand(CallStack)

            ' Deinitialize
            Init.DeInit(CallStack, False, False, True, False)

            ' Return
            Return RunLevel

        End If

        ' *****************************
        ' - Check for SQL execution switches.
        ' *****************************

        ' Check for SQL execution switches
        If Globals.AttachedtoConsole AndAlso (Globals.DbTestConnectionSwitch OrElse Globals.MdbOverviewSwitch OrElse Globals.MdbCleanAppsSwitch) Then

            ' Call SQL function dispatcher
            RunLevel = DatabaseAPI.SQLFunctionDispatch(CallStack)

            ' Deinitialize (keep debug log)
            Init.DeInit(CallStack, True, False, True, True)

            ' Return
            Return RunLevel

        End If

        ' *****************************
        ' - Check for launch app switch.
        ' *****************************

        ' Check for launch app switch
        If Globals.LaunchAppSwitch Then

            ' Laumch app
            RunLevel = LaunchPad(CallStack, Globals.LaunchAppContext, Globals.LaunchAppFileName, FileVector.GetFilePath(Globals.LaunchAppFileName), Globals.LaunchAppArguments)

            ' Deinitialize
            Init.DeInit(CallStack, False, False, True, False)

            ' Return
            Return RunLevel

        End If

        ' *****************************
        ' - Check for software library cleanup execution.
        ' *****************************

        ' Check for libray analysis or cleanup switches
        If Globals.AttachedtoConsole AndAlso (Globals.CheckSDLibrarySwitch OrElse Globals.CleanupSDLibrarySwitch) Then

            ' Cleanup library
            LibraryManager.RepairLibrary(CallStack)

            ' Deinitialize
            Init.DeInit(CallStack, False, False, True, True)

            ' Return
            Return RunLevel

        End If

        ' *****************************
        ' - Determine entry point. (Using very expanded logic)
        ' *****************************

        ' Entry point check
        If Globals.ParentProcessName.ToLower.Equals("sd_jexec") Then

            ' Write debug
            Logger.WriteDebug(CallStack, "Entry point: Software Delivery")
            Logger.WriteDebug(CallStack, "User identity: " + Globals.ProcessIdentity.Name)

            ' Set execution mode flag
            Globals.SDBasedMode = True

            ' Identity check
            If Globals.RunningAsSystemIdentity Then

                ' Silent switch check
                If Globals.HideGUISwitch Then

                    ' Write debug
                    Logger.WriteDebug(CallStack, "UI interaction: Silent")

                Else

                    ' ShowGUI switch check
                    If Globals.ShowGUISwitch Then

                        ' Write debug
                        Logger.WriteDebug(CallStack, "UI interaction: Progress GUI")

                        ' Deploy pipe clients
                        PipeClient.InitPipeClient(CallStack)

                    Else

                        ' Write debug
                        Logger.WriteDebug(CallStack, "UI interaction: Tray icon")

                        ' Deploy pipe clients
                        PipeClient.InitPipeClient(CallStack)

                    End If

                End If

            Else

                ' Silent switch check
                If Globals.HideGUISwitch Then

                    ' Write debug
                    Logger.WriteDebug(CallStack, "UI interaction: Silent")

                Else

                    ' ShowGUI switch check
                    If Globals.ShowGUISwitch Then

                        ' Write debug
                        Logger.WriteDebug(CallStack, "UI interaction: Progress GUI")

                        ' Deploy pipe clients
                        PipeClient.InitPipeClient(CallStack)

                    Else

                        ' Write debug
                        Logger.WriteDebug(CallStack, "UI interaction: Tray icon")

                        ' Deploy pipe clients
                        PipeClient.InitPipeClient(CallStack)

                    End If

                End If

            End If

            ' Update pipe clients for SD online mode
            PipeServer.SendMessage(PipeServer.SEND_SETTING, "SDOnlineMode:" + Globals.SDBasedMode.ToString)
            Logger.WriteDebug(CallStack, "Write setting:  " + "SDOnlineMode:" + Globals.SDBasedMode.ToString)

            ' *****************************
            ' - Establish software delivery output ID and location.
            ' *****************************

            ' Process the active container file (.cwf)
            RunLevel = JobContainer(CallStack)

            ' Verify SD execution is synchronized--
            '   *****
            '   "FalseStart" SD executions occur when the Administrator terminates the
            '   WinOffline job container, before all stages have run thru completion,
            '   and later, a subsequent WinOffline execution is delivered again.
            '   *****
            '   This scenario implies the C:\Windows\Temp\WinOffline\* temporary folder,
            '   was never cleaned up, as execution never ran thru completion.
            '   *****
            '   The Dispatcher() flags a "dirty" execution when StageII is requested in
            '   Software Delivery ONLINE mode, which is not per design. The Dispatcher
            '   will call Init.SDStageIReInit() to attempt correction, and carry out the
            '   orders of the current job, after cleaning up the old one.
            '   *****
            '   Since StageII is an OFFLINE mode, this means the job output ID was read
            '   directly from cache. Setting the "dirty" flag informs WinOffline main,
            '   after the Dispatcher() has returned, to dump the current cached output 
            '   ID and process the job container file to get the current/correct job 
            '   output ID.
            '   *****
            '   Similarly, during StageI and StageIII executions, which are ONLINE
            '   execution modes, during the processing of the job container file from
            '   software delivery, if a cached output ID is present, but mismatches
            '   with the ID parsed from the current container, the execution is
            '   considered "dirty". Once again, Init.SDStageIReInit() is called for
            '   remediation, but if remediation fails for some reason, "FalseStart"
            '   is set, and WinOffline shuts down gracefully, reporting the problem.
            '   *****
            If Globals.FalseStart Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Error: Exception caught processing software delivery container file.")

                ' Perform cleanup
                Init.DeInit(CallStack, True, False, True, True)

                ' Return
                Return 20

            ElseIf RunLevel <> 0 Or Globals.JobOutputFile Is Nothing Or Globals.JobOutputFile.Equals("") Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Warning: Software delivery job output will be unavailable.")

                ' Set job output flag
                Globals.WriteSDJobOutput = False

                ' Create exception
                Manifest.UpdateManifest(CallStack,
                                        Manifest.EXCEPTION_MANIFEST,
                                        {"Error: Exception caught processing software delivery container file.",
                                        "Reason: Please analyze the debug log for more information."})

            Else

                ' Write debug
                Logger.WriteDebug(CallStack, "Software delivery job output file: " + Globals.JobOutputFile)

                ' Set job output flag
                Globals.WriteSDJobOutput = True

            End If

            ' Call the execution state dispatcher
            Globals.DispatcherReturnCode = Dispatcher(CallStack)

        Else

            ' Write debug
            Logger.WriteDebug(CallStack, "Entry point: Non-Software Delivery")
            Logger.WriteDebug(CallStack, "User identity: " + Globals.ProcessIdentity.Name)

            ' Set execution mode flag
            Globals.SDBasedMode = False

            ' Identity check
            If Globals.RunningAsSystemIdentity Then

                ' Silent switch check
                If Globals.HideGUISwitch Then

                    ' Write debug
                    Logger.WriteDebug(CallStack, "UI interaction: Silent")

                Else

                    ' ShowGUI switch check
                    If Globals.ShowGUISwitch Then

                        ' Write debug
                        Logger.WriteDebug(CallStack, "UI interaction: Progress GUI")

                        ' Deploy pipe clients
                        PipeClient.InitPipeClient(CallStack)

                    Else

                        ' Write debug
                        Logger.WriteDebug(CallStack, "UI interaction: Tray icon")

                        ' Deploy pipe clients
                        PipeClient.InitPipeClient(CallStack)

                    End If

                End If

            Else

                ' Silent switch check
                If Globals.HideGUISwitch Then

                    ' Write debug
                    Logger.WriteDebug(CallStack, "UI interaction: Silent")

                Else

                    ' Check console attachment
                    If Globals.AttachedtoConsole Then

                        ' Check if ITCM is not installed
                        If Globals.ITCMInstalled = False Then

                            ' Write debug
                            Logger.WriteDebug(CallStack, "ITCM is not installed.")

                            ' Deinitialize
                            Init.DeInit(CallStack, True, False, True, False)

                            ' Return
                            Return 0

                        End If

                        ' Write debug
                        Logger.WriteDebug(CallStack, "UI interaction: Tray icon")

                        ' Deploy pipe clients
                        PipeClient.InitPipeClient(CallStack)

                    Else

                        ' Write debug
                        Logger.WriteDebug(CallStack, "UI interaction: WinOffline Explorer")

                        ' Enable log cleanup switch by default
                        Globals.CleanupLogsSwitch = True

                        ' Initialize WinOfflineUI
                        Globals.WinOfflineExplorer = New WinOfflineUI

                        ' Wait for result..
                        If Globals.WinOfflineExplorer.ShowDialog() = Windows.Forms.DialogResult.Abort Or Globals.DumpCazipxpSwitch Then

                            ' Check if user selected resource dump execution
                            If Globals.DumpCazipxpSwitch Then

                                ' Copy Cazipxp to working directory
                                System.IO.File.Copy(Globals.WinOfflineTemp + "\Cazipxp.exe", Globals.WorkingDirectory + "Cazipxp.exe", True)

                            End If

                            ' Perform cleanup (don't generate a patching summary)
                            Init.DeInit(CallStack, True, False, True, False)

                            ' Return
                            Return 0

                        End If

                        ' Create the progress gui
                        Globals.ProgressGUI = New ProgressUI()
                        Globals.ProgressGUI.Hide()
                        Globals.ProgressGUI.SetFormTitle(Globals.ProcessFriendlyName + " -- " + Globals.AppVersion)

                        ' Start the progress gui thread
                        Globals.ProgressUIThread = New System.Threading.Thread(AddressOf Globals.ProgressGUI.ShowDialog)
                        Globals.ProgressUIThread.Start()

                        ' Check tray icon policy
                        If Globals.TrayIconVisible Then

                            ' Enable debug gui notification icon
                            Globals.ProgressGUI.TrayIcon.Visible = True

                        End If

                    End If

                End If

            End If

            ' Call the execution state dispatcher
            Globals.DispatcherReturnCode = Dispatcher(CallStack)

        End If

        ' *****************************
        ' - Check for DIRTY execution flag.
        ' *****************************

        ' Check the dirty flag
        If Globals.DirtyFlag Then

            ' Write debug
            Logger.WriteDebug(CallStack, "Dirty execution flag was set.")

            ' *****************************
            ' - Reset the job output ID.
            ' *****************************

            ' Process the active container work file (.cwf)
            RunLevel = JobContainer(CallStack)

            ' Check the run level
            If RunLevel <> 0 Or Globals.JobOutputFile Is Nothing Or Globals.JobOutputFile.Equals("") Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Error: Exception caught processing software delivery container file.")
                Logger.WriteDebug(CallStack, "Warning: Software delivery job output will be unavailable.")

                ' Set job output flag
                Globals.WriteSDJobOutput = False

            Else

                ' Write debug
                Logger.WriteDebug(CallStack, "Software delivery container file processed.")
                Logger.WriteDebug(CallStack, "Software delivery job output file: " + Globals.JobOutputFile)

                ' Set job output flag
                Globals.WriteSDJobOutput = True

            End If

            ' *****************************
            ' - Reset global flags.
            ' *****************************

            ' Unset DIRTY flag
            Globals.DirtyFlag = False

            ' Reset patch error status
            Globals.PatchErrorDetected = False

        End If

        ' *****************************
        ' - Update logger task based on execution conditions.
        ' *****************************

        ' Set task based on execution type
        If Globals.SDBasedMode AndAlso Globals.StageIICompleted Then

            ' Set current task
            Logger.SetCurrentTask("Waiting for software delivery..")

        ElseIf Globals.SDBasedMode AndAlso Globals.StageICompleted Then

            ' Set current task
            Logger.SetCurrentTask("Waiting for next stage..")

        Else

            ' Set current task
            Logger.SetCurrentTask("Wrapping up..")

        End If

        ' *****************************
        ' - Process dispatcher return code and deinitialize.
        ' *****************************

        ' Check the dispatcher return
        If RunLevel <> 0 Then

            ' Write debug
            Logger.WriteDebug(CallStack, Globals.ProcessFriendlyName + " completed with an error.")

        Else

            ' Write debug
            Logger.WriteDebug(CallStack, Globals.ProcessFriendlyName + " completed.")

        End If

        ' Check final stage
        If Globals.FinalStage Then

            ' Deinitialize
            Init.DeInit(CallStack, True, False, False, True)

        End If

        ' Dispatch debug gui [User execution only]
        If Not Globals.RunningAsSystemIdentity AndAlso Globals.ProgressGUI IsNot Nothing Then

            ' Check execution type
            If Globals.FinalStage Then

                ' Autoclose the debug console [Autoclose = True]
                Globals.ProgressGUI.Delegate_ProgressUI_GoModeless(True)

            Else

                ' Hide the debug console [Autoclose = False]
                Globals.ProgressGUI.Delegate_ProgressUI_GoModeless(False)

            End If

        End If

        ' *****************************
        ' - Switch: Initiate a system reboot.
        ' *****************************

        ' Check for system reboot switch
        If Globals.RebootOnTermination Then

            ' Initiate a reboot before termination
            Utility.SystemReboot(30)

        End If

        ' *****************************
        ' - Terminate program.
        ' *****************************

        ' Check the dispatcher return code
        If Globals.DispatcherReturnCode <> 0 Then

            ' Return
            Return 21

        ElseIf Globals.PatchErrorDetected Then

            ' Return
            Return 100

        Else

            ' Return
            Return 0

        End If

    End Function

End Class