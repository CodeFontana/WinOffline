'****************************** Class Header *******************************\
' Project Name: WinOffline
' Class Name:   WinOffline/Logger
' File Name:    Logger.vb
' Author:       Brian Fontana
'***************************************************************************/

Partial Public Class WinOffline

    Public Class Logger

        ' Class variables
        Private Shared DebugLogBuffer As New ArrayList
        Private Shared DebugGUIBuffer As New ArrayList
        Private Shared DebugPipeBuffer As New ArrayList
        Private Shared ActiveLogBuffer As Boolean = False
        Private Shared ActiveGUIBuffer As Boolean = False
        Private Shared ActivePipeBuffer As Boolean = False
        Private Shared DebugGUITask As String = ""
        Private Shared PreviousTask As String = ""
        Private Shared ReInitDebugHistory As New ArrayList
        Public Shared LastCallStack As String = ""

        ' GetNowTime function
        Private Shared Function GetNowTime() As String

            ' Build custom timestamp
            Dim RightNow As String = DateTime.Now.Year.ToString() + "." +
                                     DateTime.Now.Month.ToString() + "." +
                                     DateTime.Now.Day.ToString() + "--" +
                                     DateTime.Now.TimeOfDay.ToString

            ' Return formatted string
            Return RightNow.Replace(":", ".")

        End Function

        ' Update current task
        Public Shared Sub SetCurrentTask(ByVal CurrentTask As String)

            ' Set the current task
            DebugGUITask = CurrentTask

        End Sub

        ' Write debug function (Without CallStack)
        Public Shared Sub WriteDebug(ByVal Message As String)

            ' Validate message
            If Message Is Nothing Then Return

            ' *****************************
            ' - Write debug to console.
            ' *****************************

            ' Write debug to output console
            System.Console.WriteLine(Message)

            ' *****************************
            ' - Write debug to progress ui.
            ' *****************************

            ' Write debug to debug console
            If Globals.ProgressGUI IsNot Nothing AndAlso
                Globals.ProgressUIThread IsNot Nothing Then

                ' Check if current task has changed
                If Not DebugGUITask.Equals(PreviousTask) Then

                    ' Update the GUI
                    Globals.ProgressGUI.Delegate_ProgressUI_UpdateCurrentTask(DebugGUITask)

                    ' Update the previous task
                    PreviousTask = DebugGUITask

                End If

                ' Check for an active buffer
                If ActiveGUIBuffer Then

                    ' Purge contents of buffer
                    For Each LineItem As String In DebugGUIBuffer

                        ' Write buffer
                        Globals.ProgressGUI.Delegate_ProgressUI_WriteDebug(LineItem)

                    Next

                    ' Clear the buffer
                    DebugGUIBuffer.Clear()

                    ' Deactivate buffer
                    ActiveGUIBuffer = False

                End If

                ' Write debug
                Globals.ProgressGUI.Delegate_ProgressUI_WriteDebug(Message)

            Else

                ' Activate buffer
                ActiveGUIBuffer = True

                ' Add to buffer
                DebugGUIBuffer.Add(Message)

            End If

            ' *****************************
            ' - Write debug to debug log.
            ' *****************************

            ' Write debug to log file
            If Globals.DebugStreamWriter IsNot Nothing AndAlso
                Globals.DebugStreamWriter.BaseStream IsNot Nothing Then

                Try

                    ' Check for an active buffer
                    If ActiveLogBuffer Then

                        ' Purge contents of buffer
                        For Each LineItem As String In DebugLogBuffer

                            ' Write buffer
                            Globals.DebugStreamWriter.WriteLine(LineItem)

                        Next

                        ' Clear the buffer
                        DebugLogBuffer.Clear()

                        ' Deactivate buffer
                        ActiveLogBuffer = False

                    End If

                    ' Write debug
                    Globals.DebugStreamWriter.WriteLine(Message)

                Catch ex As Exception

                    ' Activate buffer
                    ActiveLogBuffer = True

                    ' Add to buffer
                    DebugLogBuffer.Add(Message)

                End Try

            Else

                ' Activate buffer
                ActiveLogBuffer = True

                ' Add to buffer
                DebugLogBuffer.Add(Message)

            End If

            ' *****************************
            ' - Write debug to pipe server.
            ' *****************************

            ' Write debug to pipe server
            If Globals.PipeServerWriter IsNot Nothing AndAlso
                Globals.PipeServerWriter.BaseStream IsNot Nothing Then

                Try

                    ' Check for a GUI task update
                    If Not DebugGUITask.Equals(PreviousTask) Then

                        ' Send a message
                        PipeServer.SendMessage(PipeServer.SEND_STATUS, DebugGUITask)

                        ' Update the previous task
                        PreviousTask = DebugGUITask

                    End If

                    ' Check for an active buffer
                    If ActivePipeBuffer Then

                        ' Purge contents of buffer
                        For Each LineItem As String In DebugPipeBuffer

                            ' Write buffer
                            PipeServer.SendMessage(PipeServer.SEND_MESSAGE, LineItem)

                        Next

                        ' Clear the buffer
                        DebugPipeBuffer.Clear()

                        ' Deactivate buffer
                        ActivePipeBuffer = False

                    End If

                    ' Write debug
                    PipeServer.SendMessage(PipeServer.SEND_MESSAGE, Message)

                Catch ex As Exception

                    ' Activate buffer
                    ActivePipeBuffer = True

                    ' Add to buffer
                    DebugPipeBuffer.Add(Message)

                End Try

            Else

                ' Activate buffer
                ActivePipeBuffer = True

                ' Add to buffer
                DebugPipeBuffer.Add(Message)

            End If

            ' *****************************
            ' - Write debug to reinit history.
            ' *****************************

            ' Add to reinit buffer
            ReInitDebugHistory.Add(Message)

            ' *****************************
            ' - Special case: RemoveITCM UI.
            ' *****************************

            ' Check for removal UI interaction
            If LastCallStack.Contains("RemoveITCM|") AndAlso Globals.WinOfflineExplorer IsNot Nothing Then

                ' Write output to GUI control
                Globals.WinOfflineExplorer.Delegate_Sub_Append_Text(Globals.WinOfflineExplorer.txtRemovalTool, Message)

            End If

        End Sub

        ' Write debug function (With CallStack)
        Public Shared Sub WriteDebug(ByVal CallStack As String, ByVal Message As String)

            ' Local variables
            Dim FilterMessage As String = ""
            Dim FinalMessage As String = ""
            Dim CallStackChanged As Boolean = False

            ' Validate message
            If Message Is Nothing Then Return

            ' Check if callstack has changed
            If Not CallStack.ToLower.Equals(LastCallStack.ToLower) Then

                ' Set marker
                CallStackChanged = True

            End If

            ' Set last callstack
            LastCallStack = CallStack

            ' Get current timestamp
            Dim CurrentTime = GetNowTime()

            ' Setup the filter message
            FilterMessage = CallStack + Message

            ' Filter the callstack to create the final message
            While FilterMessage.Contains("|")

                ' Update final message
                FinalMessage = FilterMessage

                ' Keep filtering the message
                FilterMessage = FilterMessage.Substring(FilterMessage.IndexOf("|") + 1)

            End While

            ' *****************************
            ' - Write debug to console.
            ' *****************************

            ' Check if callstack changed
            If CallStackChanged Then

                ' Output blank line
                System.Console.WriteLine()

            End If

            ' Write debug to output console
            System.Console.WriteLine(FinalMessage)

            ' *****************************
            ' - Write debug to progress ui.
            ' *****************************

            ' Write debug to debug console
            If Globals.ProgressGUI IsNot Nothing AndAlso
                Globals.ProgressUIThread IsNot Nothing Then

                ' Check if current task has changed
                If Not DebugGUITask.Equals(PreviousTask) Then

                    ' Update the GUI
                    Globals.ProgressGUI.Delegate_ProgressUI_UpdateCurrentTask(DebugGUITask)

                    ' Update the previous task
                    PreviousTask = DebugGUITask

                End If

                ' Check for an active buffer
                If ActiveGUIBuffer Then

                    ' Purge contents of buffer
                    For Each LineItem As String In DebugGUIBuffer

                        ' Write buffer
                        Globals.ProgressGUI.Delegate_ProgressUI_WriteDebug(LineItem)

                    Next

                    ' Clear the buffer
                    DebugGUIBuffer.Clear()

                    ' Deactivate buffer
                    ActiveGUIBuffer = False

                End If

                ' Check if callstack changed
                If CallStackChanged Then

                    ' Output blank line
                    Globals.ProgressGUI.Delegate_ProgressUI_WriteDebug("")

                End If

                ' Write debug
                Globals.ProgressGUI.Delegate_ProgressUI_WriteDebug(CurrentTime + CallStack + Message)

            Else

                ' Activate buffer
                ActiveGUIBuffer = True

                ' Check if callstack changed
                If CallStackChanged Then

                    ' Append blank line to buffer
                    DebugGUIBuffer.Add("")

                End If

                ' Add to buffer
                DebugGUIBuffer.Add(CurrentTime + CallStack + Message)

            End If

            ' *****************************
            ' - Write debug to debug log.
            ' *****************************

            ' Write debug to log file
            If Globals.DebugStreamWriter IsNot Nothing AndAlso
                Globals.DebugStreamWriter.BaseStream IsNot Nothing Then

                Try

                    ' Check for an active buffer
                    If ActiveLogBuffer Then

                        ' Purge contents of buffer
                        For Each LineItem As String In DebugLogBuffer

                            ' Write buffer
                            Globals.DebugStreamWriter.WriteLine(LineItem)

                        Next

                        ' Clear the buffer
                        DebugLogBuffer.Clear()

                        ' Deactivate buffer
                        ActiveLogBuffer = False

                    End If

                    ' Check if callstack changed
                    If CallStackChanged Then

                        ' Output blank line
                        Globals.DebugStreamWriter.WriteLine("")

                    End If

                    ' Write debug
                    Globals.DebugStreamWriter.WriteLine(CurrentTime + CallStack + Message)

                Catch ex As Exception

                    ' Activate buffer
                    ActiveLogBuffer = True

                    ' Check if callstack changed
                    If CallStackChanged Then

                        ' Append blank line to buffer
                        DebugLogBuffer.Add("")

                    End If

                    ' Add to buffer
                    DebugLogBuffer.Add(CurrentTime + CallStack + Message)

                End Try

            Else

                ' Activate buffer
                ActiveLogBuffer = True

                ' Check if callstack changed
                If CallStackChanged Then

                    ' Append blank line to buffer
                    DebugLogBuffer.Add("")

                End If

                ' Add to buffer
                DebugLogBuffer.Add(CurrentTime + CallStack + Message)

            End If

            ' *****************************
            ' - Write debug to pipe server.
            ' *****************************

            ' Write debug to pipe server
            If Globals.PipeServerWriter IsNot Nothing AndAlso
                Globals.PipeServerWriter.BaseStream IsNot Nothing Then

                Try

                    ' Check for a GUI task update
                    If Not DebugGUITask.Equals(PreviousTask) Then

                        ' Send a message
                        PipeServer.SendMessage(PipeServer.SEND_STATUS, DebugGUITask)

                        ' Update the previous task
                        PreviousTask = DebugGUITask

                    End If

                    ' Check for an active buffer
                    If ActivePipeBuffer Then

                        ' Purge contents of buffer
                        For Each LineItem As String In DebugPipeBuffer

                            ' Write buffer
                            PipeServer.SendMessage(PipeServer.SEND_MESSAGE, LineItem)

                        Next

                        ' Clear the buffer
                        DebugPipeBuffer.Clear()

                        ' Deactivate buffer
                        ActivePipeBuffer = False

                    End If

                    ' Check if callstack changed
                    If CallStackChanged Then

                        ' Output blank line
                        PipeServer.SendMessage(PipeServer.SEND_MESSAGE, "")

                    End If

                    ' Write debug
                    PipeServer.SendMessage(PipeServer.SEND_MESSAGE, CurrentTime + CallStack + Message)

                Catch ex As Exception

                    ' Activate buffer
                    ActivePipeBuffer = True

                    ' Check if callstack changed
                    If CallStackChanged Then

                        ' Append blank line to buffer
                        DebugPipeBuffer.Add("")

                    End If

                    ' Add to buffer
                    DebugPipeBuffer.Add(CurrentTime + CallStack + Message)

                End Try

            Else

                ' Activate buffer
                ActivePipeBuffer = True

                ' Check if callstack changed
                If CallStackChanged Then

                    ' Append blank line to buffer
                    DebugPipeBuffer.Add("")

                End If

                ' Add to buffer
                DebugPipeBuffer.Add(CurrentTime + CallStack + Message)

            End If

            ' *****************************
            ' - Write debug to reinit history.
            ' *****************************

            ' Check if callstack changed
            If CallStackChanged Then

                ' Append blank line to reinit buffer
                ReInitDebugHistory.Add("")

            End If

            ' Add to reinit buffer
            ReInitDebugHistory.Add(CurrentTime + CallStack + Message)

            ' *****************************
            ' - Special case: RemoveITCM UI.
            ' *****************************

            ' Check for removal UI interaction
            If CallStack.Contains("RemoveITCM|") AndAlso Globals.WinOfflineExplorer IsNot Nothing Then

                ' Check if callstack changed
                If CallStackChanged Then

                    ' Output blank line
                    Globals.WinOfflineExplorer.Delegate_Sub_Append_Text(Globals.WinOfflineExplorer.txtRemovalTool, "")

                End If

                ' Write output to GUI control
                Globals.WinOfflineExplorer.Delegate_Sub_Append_Text(Globals.WinOfflineExplorer.txtRemovalTool, FinalMessage)

            End If

        End Sub

        ' Insert Summary
        Public Shared Sub InsertSummary(ByVal CallStack As String)

            ' Local variables
            Dim DebugStreamReader As System.IO.StreamReader = Nothing
            Dim CloneStreamWriter As System.IO.StreamWriter = Nothing
            Dim strLine As String

            ' Update call stack
            CallStack += "Logger|"

            ' Mess with debug stream
            Try

                ' Close the debug writer stream
                Globals.DebugStreamWriter.Close()

                ' Write debug
                Logger.WriteDebug(CallStack, "Debug stream closed.")

                ' Open the debug reader stream
                DebugStreamReader = New System.IO.StreamReader(Globals.WindowsTemp + "\" + Globals.ProcessFriendlyName + ".txt")

                ' Open the clone writer stream
                CloneStreamWriter = New System.IO.StreamWriter(Globals.WindowsTemp + "\" + Globals.ProcessFriendlyName + ".clone", False)
                CloneStreamWriter.AutoFlush = True

                ' Write debug
                Logger.WriteDebug(CallStack, "Clone stream created.")

                ' Read the log file summary
                Globals.DebugLogSummary = Manifest.DebugLogSummary

                ' Write debug
                Logger.WriteDebug(CallStack, "Debug summary created.")

                ' Completely unrelated, but for consistency, get the console at the same time
                Globals.ProgressGUISummary = Manifest.DebugConsoleSummary

                ' Insert the summary
                CloneStreamWriter.Write(Globals.DebugLogSummary)

                ' Write debug
                Logger.WriteDebug(CallStack, "Debug summary written to clone stream.")

                ' Copy the debug log
                Do While DebugStreamReader.Peek() >= 0

                    ' Read a line from the debug log
                    strLine = DebugStreamReader.ReadLine

                    ' Write the line to the clone log
                    CloneStreamWriter.WriteLine(strLine)

                Loop

                ' Write debug
                Logger.WriteDebug(CallStack, "Debug log appended to clone stream.")

                ' Close the debug reader stream
                DebugStreamReader.Close()

                ' Close the clone writer stream
                CloneStreamWriter.Close()

                ' Delete the existing debug log
                Utility.DeleteFile(CallStack, Globals.WindowsTemp + "\" + Globals.ProcessFriendlyName + ".txt")

                ' Write debug
                Logger.WriteDebug(CallStack, "Move file: " + Globals.WindowsTemp + "\" + Globals.ProcessFriendlyName + ".clone")
                Logger.WriteDebug(CallStack, "To: " + Globals.WindowsTemp + "\" + Globals.ProcessFriendlyName + ".txt")

                ' Set the clone as the new debug log
                System.IO.File.Move(Globals.WindowsTemp + "\" + Globals.ProcessFriendlyName + ".clone",
                                    Globals.WindowsTemp + "\" + Globals.ProcessFriendlyName + ".txt")

                ' Write debug
                Logger.WriteDebug(CallStack, "Debug log successfully recreated with summary.")

            Catch ex As Exception

                ' Write debug
                Logger.WriteDebug(ex.Message)
                Logger.WriteDebug(ex.StackTrace)
                Logger.WriteDebug(CallStack, "Exception caught creating debug log summary.")

                ' Close debug stream reader
                If DebugStreamReader IsNot Nothing Then

                    ' Close debug stream reader
                    DebugStreamReader.Close()

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Debug stream reader closed.")

                End If

                ' Close clone stream writer
                If CloneStreamWriter IsNot Nothing Then

                    ' Close clone stream writer
                    CloneStreamWriter.Close()

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Clone stream writer closed.")

                End If

                ' Check if clone file exists
                If System.IO.File.Exists(Globals.WindowsTemp + "\" + Globals.ProcessFriendlyName + ".clone") Then

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Delete file: " + Globals.WindowsTemp + "\" + Globals.ProcessFriendlyName + ".clone")

                    ' Delete the existing debug log
                    Utility.DeleteFile(CallStack, Globals.WindowsTemp + "\" + Globals.ProcessFriendlyName + ".clone")

                End If

            Finally

                ' Reopen the debug writer stream
                Globals.DebugStreamWriter = New System.IO.StreamWriter(Globals.WindowsTemp + "\" + Globals.ProcessFriendlyName + ".txt", True)
                Globals.DebugStreamWriter.AutoFlush = True

                ' Write debug
                Logger.WriteDebug(CallStack, "Debug stream has been restored.")

            End Try

        End Sub

        ' Initialize debug log
        Public Shared Sub InitDebugLog(ByVal CallStack As String)

            ' Local variables
            Dim DebugLog As String = Globals.WindowsTemp + "\" + Globals.ProcessFriendlyName + ".txt"

            ' Update call stack
            CallStack += "InitDebug|"

            ' Try establishing the debug stream
            Try

                ' Open the debug stream writer
                Globals.DebugStreamWriter = New System.IO.StreamWriter(DebugLog, True)

                ' Set autoflush on the debug
                Globals.DebugStreamWriter.AutoFlush = True

                ' Write debug
                Logger.WriteDebug(CallStack, "Debug stream: " + DebugLog)

            Catch ex As Exception

                ' Write debug
                Logger.WriteDebug(ex.Message)
                Logger.WriteDebug(ex.StackTrace)

                ' Write debug
                Logger.WriteDebug(CallStack, "Exception caught opening debug stream.")

                ' Throw exception
                Throw New Exception("Exception caught creating debug log: " + DebugLog +
                                    Environment.NewLine + Environment.NewLine + ex.Message)

            End Try

        End Sub

        ' Initialize alternate debug log
        Public Shared Function InitAlternateDebugLog(ByVal CallStack As String) As Integer

            ' Local variables
            Dim DebugIncrement As Integer = 0
            Dim DebugLog As String = Globals.WindowsTemp + "\" + Globals.ProcessFriendlyName + DebugIncrement.ToString + ".txt"

            ' Update call stack
            CallStack += "InitAltDebug|"

            ' Check if primary debug log already exists
            While System.IO.File.Exists(DebugLog)

                ' Write debug
                Logger.WriteDebug(CallStack, "Alternate debug stream already exists: " + DebugLog)

                ' Update increment
                DebugIncrement += 1

                ' Build alternative
                DebugLog = Globals.WindowsTemp + "\" + Globals.ProcessFriendlyName + DebugIncrement.ToString + ".txt"

            End While

            ' Try establishing the debug stream
            Try

                ' Open the debug stream writer
                Globals.DebugStreamWriter = New System.IO.StreamWriter(DebugLog, True)

                ' Set autoflush on the debug
                Globals.DebugStreamWriter.AutoFlush = True

                ' Write debug
                Logger.WriteDebug(CallStack, "Alternate debug stream: " + DebugLog)

            Catch ex As Exception

                ' Write debug
                Logger.WriteDebug(ex.Message)
                Logger.WriteDebug(ex.StackTrace)

                ' Write debug
                Logger.WriteDebug(CallStack, "Exception caught opening alternate debug stream.")

                ' Return
                Return 1

            End Try

            ' Return
            Return 0

        End Function

        ' Merge alternate debug logs
        Public Shared Function ReadAdditionalDebug(ByVal CallStack As String) As Boolean

            ' Local variables
            Dim AdditionalFound As Boolean = False
            Dim DebugIncrement As Integer = 0
            Dim DebugLog As String = Globals.WindowsTemp + "\" + Globals.ProcessFriendlyName + DebugIncrement.ToString + ".txt"
            Dim AlternateReader As System.IO.StreamReader
            Dim strLine As String

            ' Update call stack
            CallStack += "Logger|"

            ' Iterate additional debug logs
            While System.IO.File.Exists(DebugLog)

                ' Set flag
                AdditionalFound = True

                ' Write header
                Globals.AdditionalDebug.Add("Parallel Process Debug (" + Globals.ProcessFriendlyName + DebugIncrement.ToString + ".txt):")

                ' Read in alternate debug log
                Try

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Open file: " + DebugLog)

                    ' Open the debug log
                    AlternateReader = New System.IO.StreamReader(DebugLog)

                    ' Loop debug log contents
                    Do While AlternateReader.Peek() >= 0

                        ' Read a line
                        strLine = AlternateReader.ReadLine()

                        ' Add to array
                        Globals.AdditionalDebug.Add(strLine)

                    Loop

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Finished reading in contents of log.")
                    Logger.WriteDebug(CallStack, "Close file: " + DebugLog)

                    ' Close the debug log
                    AlternateReader.Close()

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Delete file: " + DebugLog)

                    ' Delete the alternate debug log
                    Utility.DeleteFile(CallStack, DebugLog)

                Catch ex As Exception

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Error: Exception caught reading additional debug log.")
                    Logger.WriteDebug(ex.Message)
                    Logger.WriteDebug(ex.StackTrace)

                    ' Create exception
                    Manifest.UpdateManifest(CallStack, Manifest.EXCEPTION_MANIFEST, {ex.Message, ex.StackTrace})

                End Try

                ' Write footer
                Globals.AdditionalDebug.Add("")

                ' Update increment
                DebugIncrement += 1

                ' Build next log
                DebugLog = Globals.WindowsTemp + "\" + Globals.ProcessFriendlyName + DebugIncrement.ToString + ".txt"

            End While

            ' Check if any alternates were found
            If AdditionalFound = False Then

                ' Write debug
                Logger.WriteDebug(CallStack, "No additional debug logs found.")

                ' Return
                Return False

            Else

                ' Write debug
                Logger.WriteDebug(CallStack, "Additional debug logs have been processed.")

                ' Return
                Return True

            End If

        End Function

        ' Terminate debug log
        Public Shared Sub TermDebugLog(ByVal CallStack As String, ByVal KeepDebugLog As Boolean)

            ' Local variables
            Dim DebugLog As String

            ' Skip if debug log does not exist
            If Globals.DebugStreamWriter Is Nothing OrElse Not System.IO.File.Exists(Globals.WindowsTemp + "\" + Globals.ProcessFriendlyName + ".txt") Then

                ' Return
                Return

            End If

            ' Read in alternate debug logs
            If ReadAdditionalDebug(CallStack) Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Contents of additional debug logs:" + Environment.NewLine)
                Logger.WriteDebug("############################################################")

                ' Iterate array contents
                For Each strLine As String In Globals.AdditionalDebug

                    ' Write debug
                    Logger.WriteDebug(strLine)

                Next

                ' Write debug
                Logger.WriteDebug("############################################################" + Environment.NewLine)

            End If

            ' Check for job output marker
            If Globals.WriteSDJobOutput Then

                ' Write debug
                Logger.WriteDebug(CallStack, "Copy File: " + Globals.WindowsTemp + "\" + Globals.ProcessFriendlyName + ".txt")
                Logger.WriteDebug(CallStack, "To: " + Globals.JobOutputFile)

                ' Copy contents of debug file to job output file
                System.IO.File.Copy(Globals.WindowsTemp + "\" + Globals.ProcessFriendlyName + ".txt", Globals.JobOutputFile, True)

            End If

            ' Check if debug log gets archived
            If KeepDebugLog Then

                ' Assign debug log completion timestamp
                DebugLog = Globals.ProcessFriendlyName + "-" +
                    DateTime.Now.Year.ToString() + "." +
                    DateTime.Now.Month.ToString() + "." +
                    DateTime.Now.Day.ToString() + "--" +
                    DateTime.Now.TimeOfDay.ToString()

                ' Prune insignificant time digits
                DebugLog = DebugLog.Substring(0, DebugLog.LastIndexOf("."))

                ' Append file extenstion
                DebugLog = DebugLog + ".txt"

                ' Filter timestamp
                DebugLog = DebugLog.Replace(":", ".")

                ' Write debug
                Logger.WriteDebug(CallStack, "Archive Log: " + Globals.DSMFolder + DebugLog)

                ' Close debug stream
                Globals.DebugStreamWriter.Close()

                ' Move the debug log
                Try

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Move File: " + Globals.WindowsTemp + "\" + Globals.ProcessFriendlyName + ".txt")
                    Logger.WriteDebug(CallStack, "To: " + Globals.DSMFolder + DebugLog)

                    ' Move debug log to dsm directory
                    System.IO.File.Move(Globals.WindowsTemp + "\" + Globals.ProcessFriendlyName + ".txt", Globals.DSMFolder + DebugLog)

                Catch ex As Exception

                    ' Write debug - suppress
                    'Logger.WriteDebug(ex.Message)
                    'Logger.WriteDebug(ex.StackTrace)
                    'Logger.WriteDebug(CallStack, "Error: Exception caught moving debug log to archive.")

                    ' Re-open the debug stream
                    Globals.DebugStreamWriter = New System.IO.StreamWriter(Globals.WindowsTemp + "\" + Globals.ProcessFriendlyName + ".txt", True)

                    ' Set autoflush on the debug
                    Globals.DebugStreamWriter.AutoFlush = True

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Debug log abandoned: " + Globals.WindowsTemp + "\" + Globals.ProcessFriendlyName + ".txt")

                    ' Close debug stream
                    Globals.DebugStreamWriter.Close()

                End Try

            Else

                ' Close debug stream
                Globals.DebugStreamWriter.Close()

                ' Delete the debug log
                Try

                    ' Rename debug log in temp directory
                    Utility.DeleteFile(CallStack, Globals.WindowsTemp + "\" + Globals.ProcessFriendlyName + ".txt", True)

                Catch ex As Exception

                    ' Write debug
                    Logger.WriteDebug(ex.Message)
                    Logger.WriteDebug(ex.StackTrace)
                    Logger.WriteDebug(CallStack, "Error: Exception caught deleting debug log.")

                    ' Re-open the debug stream
                    Globals.DebugStreamWriter = New System.IO.StreamWriter(Globals.WindowsTemp + "\" + Globals.ProcessFriendlyName + ".txt", True)

                    ' Set autoflush on the debug
                    Globals.DebugStreamWriter.AutoFlush = True

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Debug log abandoned.")

                    ' Close debug stream
                    Globals.DebugStreamWriter.Close()

                End Try

            End If

        End Sub

        ' Reinit debug log
        '   Referenced by Init.SDStageIReInit().
        '   Allows the prior debugging from a failed job to be purged.
        Public Shared Sub SDStageIReInitDebugLog()

            ' Close the debug writer stream
            Globals.DebugStreamWriter.Close()

            ' Reopen the debug writer stream
            Globals.DebugStreamWriter = New System.IO.StreamWriter(Globals.WindowsTemp + "\" + Globals.ProcessFriendlyName + ".txt", False)
            Globals.DebugStreamWriter.AutoFlush = True

            ' Rewrite debug log using buffer
            For Each LineItem As String In Logger.ReInitDebugHistory

                ' Write buffer
                Globals.DebugStreamWriter.WriteLine(LineItem)

            Next

        End Sub

    End Class

End Class