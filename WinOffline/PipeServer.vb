'****************************** Class Header *******************************\
' Project Name: WinOffline
' Class Name:   WinOffline/PipeServer
' File Name:    PipeServer.vb
' Author:       Brian Fontana
'***************************************************************************/

Imports System.Threading

Partial Public Class WinOffline

    Public Class PipeServer

        Public Const SEND_SETTING As String = "SEND_SETTING:"
        Public Const SEND_MESSAGE As String = "SEND_MESSAGE:"
        Public Const SEND_STATUS As String = "SEND_STATUS:"
        Public Const CURTAINS_CALL As String = "CURTAINS_CALL:"
        Public Const REFRESH_SYSTRAY As String = "REFRESH_SYSTRAY"
        Private Shared ReInitPipeHistory As New ArrayList

        Public Shared Sub InitPipeServer(ByVal CallStack As String)

            ' *****************************
            ' - Create pipe server.
            ' *****************************

            ' Update call stack
            CallStack += "InitPipe|"

            ' Set pipe server filename
            Globals.PipeServerFileName = Globals.WindowsTemp + "\" + Globals.ProcessFriendlyName + ".pipe"

            ' Check if pipe file already exsits
            If System.IO.File.Exists(Globals.PipeServerFileName) Then

                ' Set flag
                Globals.PipeServerAlreadyExists = True

                ' Write debug
                Logger.WriteDebug(CallStack, "Pipe file already exists.")
                Logger.WriteDebug(CallStack, "Restore pipe: " + Globals.PipeServerFileName)

                ' Establish pipe server
                Try

                    ' Open the stream writer
                    Globals.PipeServerWriter = New System.IO.StreamWriter(Globals.PipeServerFileName, True)

                    ' Set autoflush on the writer
                    Globals.PipeServerWriter.AutoFlush = True

                Catch ex As Exception

                    ' Write debug
                    Logger.WriteDebug(ex.Message)
                    Logger.WriteDebug(ex.StackTrace)
                    Logger.WriteDebug(CallStack, "Exception caught restoring pipe server.")

                    ' Throw exception
                    Throw New Exception("Exception caught restoring existing pipe server: " + Globals.PipeServerFileName +
                                        Environment.NewLine + Environment.NewLine + ex.Message)

                End Try

            Else

                ' Write debug
                Logger.WriteDebug(CallStack, "Establish pipe: " + Globals.PipeServerFileName)

                ' Establish pipe server
                Try

                    ' Open the stream writer
                    Globals.PipeServerWriter = New System.IO.StreamWriter(Globals.PipeServerFileName, True)

                    ' Set autoflush on the writer
                    Globals.PipeServerWriter.AutoFlush = True

                    ' Write settings (and debug)
                    SendMessage(SEND_SETTING, "ShowGUISwitch:" + Globals.ShowGUISwitch.ToString)
                    Logger.WriteDebug(CallStack, "Write setting:  " + "ShowGUISwitch:" + Globals.ShowGUISwitch.ToString)

                Catch ex As Exception

                    ' Write debug
                    Logger.WriteDebug(ex.Message)
                    Logger.WriteDebug(ex.StackTrace)
                    Logger.WriteDebug(CallStack, "Exception caught creating pipe server.")

                    ' Throw exception
                    Throw New Exception("Exception caught establishing pipe server: " + Globals.PipeServerFileName +
                                        Environment.NewLine + Environment.NewLine + ex.Message)

                End Try

            End If

        End Sub

        Public Shared Function SendMessage(ByVal RequestType As String, ByVal Message As String) As Integer

            ' Try sending a message
            Try

                ' Ensure a valid message
                If Message IsNot Nothing AndAlso Globals.PipeServerWriter IsNot Nothing Then

                    ' Write message
                    Globals.PipeServerWriter.WriteLine(RequestType + Message.Replace(Environment.NewLine, Environment.NewLine + RequestType))

                    ' Save history
                    ReInitPipeHistory.Add(RequestType + Message)

                End If

            Catch ex As Exception

                ' Write debug
                Logger.WriteDebug(ex.Message)
                Logger.WriteDebug(ex.StackTrace)
                Logger.WriteDebug("Exception caught sending pipe request.")

                ' Return
                Return 1

            End Try

            ' Return
            Return 0

        End Function

        Public Shared Function TermPipeServer(ByVal CallStack As String, ByVal Autoclose As Boolean) As Integer

            ' Local variables
            Dim LoopCounter As Integer = 0

            ' Skip if pipe server does not exist
            If Globals.PipeServerFileName Is Nothing OrElse Not System.IO.File.Exists(Globals.PipeServerFileName) Then

                ' Return
                Return 0

            End If

            ' Write debug
            Logger.WriteDebug(CallStack, "Broadcast term signal to pipe clients..")

            ' Broadcast termination message to pipe clients
            '   [TRUE = Autoclose based on Timer, FALSE = Hide immediately]
            SendMessage(CURTAINS_CALL, Autoclose.ToString)

            ' Check autoclose setting
            If Autoclose Then

                ' Loop for graceful pipe client termination
                While Utility.IsProcessRunning(Globals.ProcessShortName, "-client")

                    ' Delay execution
                    Thread.Sleep(Globals.THREAD_REST_INTERVAL)

                End While

            End If

            ' Check for active pipe client processes
            If Utility.IsProcessRunning(Globals.ProcessShortName, "-client") Then

                ' Loop for graceful pipe client termination
                While Utility.IsProcessRunning(Globals.ProcessShortName, "-client")

                    ' Write debug
                    Logger.WriteDebug(CallStack, "Pipe client(s): ACTIVE")

                    ' Delay execution
                    Thread.Sleep(5000)

                    ' Increment the loop counter
                    LoopCounter += 1

                    ' Check the loop counter
                    If LoopCounter >= 5 Then

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Terminate pipe client processes..")

                        ' Kill pipe client processes
                        Utility.KillProcessByCommandLine(Globals.ProcessShortName, "-client")

                        ' Write debug
                        Logger.WriteDebug(CallStack, "Pipe client(s): TERMINATED")

                        ' Infinite loop safety -- Exit while condition
                        Exit While

                    End If

                End While

                ' Write debug
                Logger.WriteDebug(CallStack, "Pipe client(s): CLOSED")

            Else

                ' Write debug
                Logger.WriteDebug(CallStack, "Pipe client(s): CLOSED")

            End If

            ' Write debug
            Logger.WriteDebug(CallStack, "Close pipe server..")

            ' Verify stream is open
            If Globals.PipeServerWriter IsNot Nothing Then

                ' Close the pipe server stream
                Globals.PipeServerWriter.Close()

            End If

            ' Attempt server side cleanup
            Utility.DeleteFile(CallStack, Globals.PipeServerFileName)

            ' Write debug
            Logger.WriteDebug(CallStack, "Pipe server closed.")

            ' Return
            Return 0

        End Function

    End Class

End Class