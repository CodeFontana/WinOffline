'****************************** Class Header *******************************\
' Project Name: WinOffline
' Class Name:   WinOfflineUI
' File Name:    WinOfflineUI-Sql-CleanApps.vb
' Author:       Brian Fontana
'***************************************************************************/

Imports System.Data.SqlClient
Imports System.Text.RegularExpressions
Imports System.Threading

Partial Public Class WinOfflineUI

    Private Shared SqlCleanAppsThread As Thread
    Private Shared CafControlThread As Thread
    Private Shared ButtonMonitorThread As Thread

    Private Sub InitSqlCleanApps()

        ' Disable buttons
        Delegate_Sub_Enable_Red_Button(btnSqlDisconnectCleanApps, False)
        Delegate_Sub_Enable_Yellow_Button(btnSqlCafStopCleanApps, False)
        Delegate_Sub_Enable_Tan_Button(btnSqlRunCleanApps, False)

        ' Check caf status
        If WinOffline.Utility.IsProcessRunning("caf.exe", "service") Then

            ' Update button text
            Delegate_Sub_Set_Text(btnSqlCafStopCleanApps, "&Stop CAF")

        Else

            ' Update button text
            Delegate_Sub_Set_Text(btnSqlCafStopCleanApps, "&Start CAF")

        End If

    End Sub

    Private Sub btnSqlConnectCleanApps_Click(sender As Object, e As EventArgs) Handles btnSqlConnectCleanApps.Click

        ' Perform SQL connection
        SqlConnect()

    End Sub

    Private Sub btnSqlDisconnectCleanApps_Click(sender As Object, e As EventArgs) Handles btnSqlDisconnectCleanApps.Click

        ' Perform disconnect method
        SqlDisconnect()

    End Sub

    Private Sub btnSqlCafStopCleanApps_Click(sender As Object, e As EventArgs) Handles btnSqlCafStopCleanApps.Click

        ' Check if manager name matches the current hostname
        If Not ManagerName.ToLower.Equals(Globals.HostName.ToLower) Then

            ' Push user alert
            AlertBox.CreateUserAlert("Functionality unavailable. You are running " +
                                     Globals.ProcessFriendlyName + " remotely from " + ManagerName + ". " +
                                     "It is recommended to stop CAF on " + ManagerName + " before running the cleanup.")

            ' Disable button
            Delegate_Sub_Enable_Yellow_Button(btnSqlCafStopCleanApps, False)

            ' Return
            Return

        End If

        ' Check button status -- are we stop or start
        If btnSqlCafStopCleanApps.Text.ToLower.Equals("&stop caf") Then

            ' Disable buttons
            Delegate_Sub_Enable_Yellow_Button(btnSqlCafStopCleanApps, False)
            Delegate_Sub_Enable_Tan_Button(btnSqlRunCleanApps, False)

            ' Create and initialize a new thread to stop caf
            CafControlThread = New Thread(Sub() StopCAF(txtSqlCleanApps))
            CafControlThread.Start()

            ' Create and initialize a new thread for updating button status
            ButtonMonitorThread = New Thread(Sub() ButtonMonitorWorker())
            ButtonMonitorThread.Start()

        ElseIf btnSqlCafStopCleanApps.Text.ToLower.Equals("&start caf") Then

            ' Disable button
            Delegate_Sub_Enable_Yellow_Button(btnSqlCafStopCleanApps, False)
            Delegate_Sub_Enable_Tan_Button(btnSqlRunCleanApps, False)

            ' Create and initialize a new thread to stop caf
            CafControlThread = New Thread(Sub() StartCAF(txtSqlCleanApps))
            CafControlThread.Start()

            ' Create and initialize a new thread for updating button status
            ButtonMonitorThread = New Thread(Sub() ButtonMonitorWorker())
            ButtonMonitorThread.Start()

        End If

    End Sub

    Private Sub ButtonMonitorWorker()

        ' Wait for CAFControlThread to complete
        CafControlThread.Join()

        ' Check caf status
        If WinOffline.Utility.IsProcessRunning("caf.exe", "service") Then

            ' Update button text
            Delegate_Sub_Set_Text(btnSqlCafStopCleanApps, "&Stop CAF")

        Else

            ' Update button text
            Delegate_Sub_Set_Text(btnSqlCafStopCleanApps, "&Start CAF")

        End If

        ' Enable buttons
        Delegate_Sub_Enable_Yellow_Button(btnSqlCafStopCleanApps, True)
        Delegate_Sub_Enable_Tan_Button(btnSqlRunCleanApps, True)

    End Sub

    Private Sub btnSqlRunCleanApps_Click(sender As Object, e As EventArgs) Handles btnSqlRunCleanApps.Click

        ' Local variables
        Dim SqlScript As String = My.Resources.SQL_CleanApps_Unified
        Dim ConnectionString As String

        ' Check button status -- are we run or cancel
        If btnSqlRunCleanApps.Text.ToLower.Equals("&run cleanapps") Then

            ' Build connection string
            ConnectionString = BuildSqlConnectionString()

            ' Check connection string
            If ConnectionString Is Nothing Then

                ' Return
                Return

            End If

            ' Update button
            Delegate_Sub_Set_Text(btnSqlRunCleanApps, "C&ancel CleanApps")

            ' Disable button
            Delegate_Sub_Enable_Yellow_Button(btnSqlCafStopCleanApps, False)

            ' Clear output
            Delegate_Sub_Set_Text(txtSqlCleanApps, "")

            ' Initialize and start a submit thread
            SqlCleanAppsThread = New Thread(Sub() SqlCleanAppsWorker(ConnectionString, SqlScript))
            SqlCleanAppsThread.Start()

        ElseIf btnSqlRunCleanApps.Text.ToLower.Equals("c&ancel cleanapps") Then

            ' Set cancel flag
            CancelSignal = True

            ' Update results to reflect cancellation
            Delegate_Sub_Append_Text(txtSqlCleanApps, "Cleanup cancelled by user.")

            ' Update button
            Delegate_Sub_Set_Text(btnSqlRunCleanApps, "&Run CleanApps")

            ' Check if manager name matches the current hostname
            If Not ManagerName.ToLower.Equals(Globals.HostName.ToLower) Then

                ' Disable button
                Delegate_Sub_Enable_Yellow_Button(btnSqlCafStopCleanApps, False)

            End If

        End If

    End Sub

    Private Sub btnSqlExitCleanApps_Click(sender As Object, e As EventArgs) Handles btnSqlExitCleanApps.Click

        ' Close the dialog
        Close()

    End Sub

    Private Sub SqlCleanAppsWorker(ByVal ConnectionString As String, ByVal SqlScript As String)

        ' Local variables
        Dim DbConnection As SqlConnection
        Dim SqlBatches As String()
        Dim SqlCmd As SqlCommand
        Dim SqlData As SqlDataReader
        Dim CallStack As String = "SqlCleanAppsWorker --> "

        ' Encapsulate sql operation
        Try

            ' Create sql connection
            DbConnection = New SqlConnection(ConnectionString)

            ' Attach informational message handler to connection
            AddHandler DbConnection.InfoMessage, New SqlInfoMessageEventHandler(AddressOf OnSqlCleanAppsInfoMessage)

            ' Open sql connection
            DbConnection.Open()

            ' Write debug
            Delegate_Sub_Append_Text(rtbDebug, CallStack + "Connected successful: " + SqlServer)

            ' Enumerate sql batches based on "GO"
            SqlBatches = Regex.Split(SqlScript, "^\s*GO\s*\d*\s*($|\-\-.*$)", RegexOptions.Multiline Or RegexOptions.IgnorePatternWhitespace Or RegexOptions.IgnoreCase)

            ' Iterate each batch of queries
            For Each QueryLine As String In SqlBatches

                ' Check for empty line
                If QueryLine.Equals("") Then Continue For

                ' Build sql command from user query
                SqlCmd = New SqlCommand(QueryLine, DbConnection)

                ' Write debug
                Delegate_Sub_Append_Text(rtbDebug, CallStack + "Execute statement(s):" + Environment.NewLine + QueryLine)

                ' Execute query -- ExecuteReader method
                SqlData = SqlCmd.ExecuteReader()

                ' Check for termination or cancellation signal
                If TerminateSignal Or CancelSignal Then

                    ' Close sql data reader
                    SqlData.Close()

                    ' Close the database connection
                    DbConnection.Close()

                    ' Write debug
                    Delegate_Sub_Append_Text(rtbDebug, CallStack + "Database connection closed.")

                    ' Return
                    Return

                End If

                ' Iterate each record
                While SqlData.Read

                    ' Check for termination or cancellation signal
                    If TerminateSignal Or CancelSignal Then

                        ' Close sql data reader
                        SqlData.Close()

                        ' Close the database connection
                        DbConnection.Close()

                        ' Write debug
                        Delegate_Sub_Append_Text(rtbDebug, CallStack + "Database connection closed.")

                        ' Return
                        Return

                    End If

                End While

                ' Close sql data reader
                SqlData.Close()

            Next

            ' Close the database connection
            DbConnection.Close()

            ' Write debug
            Delegate_Sub_Append_Text(rtbDebug, CallStack + "Database connection closed.")

        Catch ex As Exception

            ' Write debug
            Delegate_Sub_Append_Text(rtbDebug, CallStack + "Exception:" + Environment.NewLine + ex.Message)
            Delegate_Sub_Append_Text(rtbDebug, CallStack + "Stack trace: " + Environment.NewLine + ex.StackTrace)

        Finally

            ' Check for termination signal
            If Not TerminateSignal Then

                ' Update button
                Delegate_Sub_Set_Text(btnSqlRunCleanApps, "&Run CleanApps")

                ' Enable button
                Delegate_Sub_Enable_Yellow_Button(btnSqlCafStopCleanApps, True)

            End If

        End Try

    End Sub

    Private Sub OnSqlCleanAppsInfoMessage(ByVal sender As Object, ByVal e As SqlInfoMessageEventArgs)

        ' Write debug
        Delegate_Sub_Append_Text(txtSqlCleanApps, e.Message)

    End Sub

End Class