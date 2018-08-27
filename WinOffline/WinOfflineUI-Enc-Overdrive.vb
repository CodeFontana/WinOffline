Imports System.Windows.Forms
Imports System.Threading

Partial Public Class WinOfflineUI

    Private Shared InternalTermSignal As Boolean = False
    Private Shared EncClientMonitorThread As Thread
    Private Shared EncOverdriveThread As Thread
    Private Shared dgvEncStressTable_mutex As Mutex

    Private Sub InitEncOverdrive()

        ' Local variables
        Dim DataGridViewCellStyle1 As DataGridViewCellStyle = New DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As DataGridViewCellStyle = New DataGridViewCellStyle()

        ' Replace WinOffline references with process name
        lblEncStressIntro.Text = WinOffline.Utility.ReplaceWord(lblEncStressIntro.Text, "WinOffline", Globals.ProcessFriendlyName)

        ' Set initial button availability
        Delegate_Sub_Enable_Blue_Button(btnEncStressStart, True)
        Delegate_Sub_Enable_Blue_Button(btnEncStressStop, False)

        ' Add columns to data grid
        Delegate_Sub_Add_DataGridView_Column(dgvEncStressTable, "col_encClient", "Client")
        Delegate_Sub_Add_DataGridView_Column(dgvEncStressTable, "col_active_threads", "Active Thread(s)")
        Delegate_Sub_Add_DataGridView_Column(dgvEncStressTable, "col_completed_threads", "Completed Thread(s)")
        Delegate_Sub_Add_DataGridView_Column(dgvEncStressTable, "col_successful_pings", "OK Pings")
        Delegate_Sub_Add_DataGridView_Column(dgvEncStressTable, "col_failed_pings", "Failed Pings")
        Delegate_Sub_Add_DataGridView_Column(dgvEncStressTable, "col_total_bytes", "Total KB")

        ' Set data grid column header style (since visual studio tends to reset this)
        DataGridViewCellStyle1.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New Drawing.Font("Calibri", 9.0!, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = DataGridViewTriState.[False]
        dgvEncStressTable.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1

        ' Set data grid default cell style (since visual studio tends to reset these too)
        DataGridViewCellStyle2.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = Drawing.Color.Beige
        DataGridViewCellStyle2.Font = New Drawing.Font("Calibri", 9.0!, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = DataGridViewTriState.[False]
        dgvEncStressTable.DefaultCellStyle = DataGridViewCellStyle2

        ' Set data grid column sizing modes
        dgvEncStressTable.Columns(0).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        dgvEncStressTable.Columns(1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        dgvEncStressTable.Columns(2).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        dgvEncStressTable.Columns(3).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        dgvEncStressTable.Columns(4).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        dgvEncStressTable.Columns(5).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells

        ' Set data grid column sorting modes
        dgvEncStressTable.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
        dgvEncStressTable.Columns(1).SortMode = DataGridViewColumnSortMode.NotSortable
        dgvEncStressTable.Columns(2).SortMode = DataGridViewColumnSortMode.NotSortable
        dgvEncStressTable.Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable
        dgvEncStressTable.Columns(4).SortMode = DataGridViewColumnSortMode.NotSortable
        dgvEncStressTable.Columns(5).SortMode = DataGridViewColumnSortMode.NotSortable

        ' Override default and align the first column (enc client) to the left
        dgvEncStressTable.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        ' Initialize the threads and supporting mutex
        EncClientMonitorThread = New Thread(AddressOf EncClientMonitorWorker)
        EncOverdriveThread = New Thread(AddressOf EncOverdriveWorker)
        dgvEncStressTable_mutex = New Mutex()

    End Sub

    Private Sub PingWorker(ByVal Victim As String)

        ' Local variables
        Dim ClientIndex As Integer
        Dim NumPings As Integer
        Dim MsgSize As Integer
        Dim ReplyTimeout As Integer
        Dim WaitTime As Integer
        Dim ActiveThreadCount As Integer
        Dim CompletedThreadCount As Integer
        Dim CompletedPingCount As Integer
        Dim FailedPingCount As Integer
        Dim CurrentBytes As Double
        Dim Output As String = ""
        Dim TokenizedOutput As String()
        Dim NumPingsOK As Integer = 0
        Dim NumPingsFailed As Integer = 0

        ' Wait for the mutex -- data grid access
        dgvEncStressTable_mutex.WaitOne()

        ' Get the client index in the data grid
        ClientIndex = GetEncClientList.IndexOf(Victim)

        ' Parse ping settings
        NumPings = Integer.Parse(txtEncStressNumPings.Text)
        MsgSize = Integer.Parse(txtEncStressMsgSize.Text)
        ReplyTimeout = Integer.Parse(txtEncStressReplyTimeout.Text)
        WaitTime = Integer.Parse(txtEncStressDelayTime.Text)

        ' Parse active thread count
        ActiveThreadCount = dgvEncStressTable.Rows(ClientIndex).Cells("col_active_threads").Value

        ' Increment the active thread count
        Delegate_Sub_Update_DataGridView_Value(dgvEncStressTable, ClientIndex, "col_active_threads", ActiveThreadCount + 1)

        ' Release the mutex -- data grid access
        dgvEncStressTable_mutex.ReleaseMutex()

        ' Check ping type settings
        If rbnEncStressCamPing.Checked Then

            ' Run camping and read output
            WinOffline.Utility.RunCommand(Globals.CAMFolder + "\bin\camping.exe", Victim +
                                          " -r " + NumPings.ToString +
                                          " -s " + MsgSize.ToString +
                                          " -w " + WaitTime.ToString +
                                          " -t " + ReplyTimeout.ToString,
                                          ,
                                          Output)

            ' Tokenize the output
            TokenizedOutput = Output.Split(New String() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)

            ' Iterate the output
            For Each strLine In TokenizedOutput

                ' Count each successful or failed ping
                If strLine.Contains("rtt") Then NumPingsOK += 1 Else NumPingsFailed += 1

            Next

        Else

            ' Run caf ping and read output
            WinOffline.Utility.RunCommand(Globals.DSMFolder + "bin\caf.exe", "ping " + Victim +
                                          " timeout " + ReplyTimeout.ToString +
                                          " repeat " + NumPings.ToString,
                                          Output)

            ' Tokenize the output
            TokenizedOutput = Output.Split(New String() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)

            ' Iterate the output
            For Each strLine In TokenizedOutput

                ' Inspect each line of output
                If strLine.Contains("ms:") And strLine.Contains("Failed to send a message to caf") Then

                    ' Increment failed ping count
                    NumPingsFailed += 1

                ElseIf strLine.Contains("ms:") Then

                    ' Increment successful ping count
                    NumPingsOK += 1

                End If

            Next

        End If

        ' Wait for the mutex -- data grid access
        dgvEncStressTable_mutex.WaitOne()

        ' Verify the client is still in the data grid
        If GetEncClientList.Contains(Victim) Then

            ' Get the client index in the data grid
            ClientIndex = GetEncClientList.IndexOf(Victim)

            ' Parse current statistics
            ActiveThreadCount = dgvEncStressTable.Rows(ClientIndex).Cells("col_active_threads").Value
            CompletedThreadCount = dgvEncStressTable.Rows(ClientIndex).Cells("col_completed_threads").Value
            CompletedPingCount = dgvEncStressTable.Rows(ClientIndex).Cells("col_successful_pings").Value
            FailedPingCount = dgvEncStressTable.Rows(ClientIndex).Cells("col_failed_pings").Value
            CurrentBytes = dgvEncStressTable.Rows(ClientIndex).Cells("col_total_bytes").Value

            ' Write updated statistics
            Delegate_Sub_Update_DataGridView_Value(dgvEncStressTable, ClientIndex, "col_active_threads", ActiveThreadCount - 1)
            Delegate_Sub_Update_DataGridView_Value(dgvEncStressTable, ClientIndex, "col_completed_threads", CompletedThreadCount + 1)
            Delegate_Sub_Update_DataGridView_Value(dgvEncStressTable, ClientIndex, "col_successful_pings", CompletedPingCount + NumPingsOK)
            Delegate_Sub_Update_DataGridView_Value(dgvEncStressTable, ClientIndex, "col_failed_pings", FailedPingCount + NumPingsFailed)
            Delegate_Sub_Update_DataGridView_Value(dgvEncStressTable, ClientIndex, "col_total_bytes", CurrentBytes + NumPingsOK * MsgSize / 1000)

        End If

        ' Release the mutex -- data grid access
        dgvEncStressTable_mutex.ReleaseMutex()

    End Sub

    Private Sub TargetWorker(ByVal Victim As String)

        ' Local variables
        Dim LocalThreadList As New ArrayList
        Dim PingThread As Thread

        ' Iterate based on user setting for number of threads to create
        For i = 0 To Integer.Parse(txtEncStressNumPingThreads.Text) - 1

            ' Create a new ping thread
            PingThread = New Thread(Sub() PingWorker(Victim))

            ' Add it to the local list
            LocalThreadList.Add(PingThread)

            ' Start the thread
            PingThread.Start()

            ' Rest this thread
            Thread.Sleep(Globals.THREAD_REST_INTERVAL)

        Next

        ' Loop until all ping threads have completed
        While LocalThreadList.Count > 0

            ' Iterate the local list
            For x As Integer = LocalThreadList.Count - 1 To 0 Step -1

                ' Check if the thread is alive or dead
                If Not DirectCast(LocalThreadList.Item(x), Thread).IsAlive Then

                    ' Remove the thread from the local list
                    LocalThreadList.RemoveAt(x)

                End If

            Next

            ' Rest this thread
            Thread.Sleep(Globals.THREAD_REST_INTERVAL)

        End While

    End Sub

    Private Sub EncOverdriveWorker()

        ' Local variables
        Dim LocalThreadList As New ArrayList
        Dim RandomGenerator As New Random
        Dim Victim As String
        Dim TargetThread As Thread

        ' Loop dangerously
        While True

            ' Check for conditions to create new target threads
            If LocalThreadList.Count / 2 < Integer.Parse(txtEncStressNumTargetThreads.Text) AndAlso
                LocalThreadList.Count / 2 < GetEncClientList.Count AndAlso
                GetEncClientList.Count > 0 AndAlso
                Not (TerminateSignal Or InternalTermSignal) Then

                ' Loop dangerously for a new victim (enc client)
                Do
                    ' Randomly choose a victim
                    Victim = GetEncClientList.Item(RandomGenerator.Next(0, GetEncClientList.Count))

                    ' Verify the victim is not already a victim of another thread
                    If Not LocalThreadList.Contains(Victim) Or (TerminateSignal Or InternalTermSignal) Then Exit Do

                Loop

                ' Check for termination signal (could have occurred while looping for a new victim)
                If Not (TerminateSignal Or InternalTermSignal) Then

                    ' Create a new target thread
                    TargetThread = New Thread(Sub() TargetWorker(Victim))

                    ' Add the thread and victim to the local list
                    LocalThreadList.Add(TargetThread)
                    LocalThreadList.Add(Victim)

                    ' Start the thread
                    TargetThread.Start()
                End If

            ElseIf TerminateSignal Or InternalTermSignal Then

                ' Loop until all target threads have completed
                While LocalThreadList.Count > 0

                    ' Iterate backwards through the local list (checking threads and skipping victims)
                    For x As Integer = LocalThreadList.Count - 2 To 0 Step -2

                        ' Check if the thread is alive or dead
                        If Not DirectCast(LocalThreadList.Item(x), Thread).IsAlive Then

                            ' Remove the thread and victim from the list
                            LocalThreadList.RemoveAt(x)
                            LocalThreadList.RemoveAt(x)

                        End If

                    Next

                    ' Rest the thread
                    Thread.Sleep(Globals.THREAD_REST_INTERVAL)

                End While

                ' Stop condition -- all target threads have finished
                Exit While

            Else

                ' Iterate backwards through the local list (checking threads and skipping victims)
                For x As Integer = LocalThreadList.Count - 2 To 0 Step -2

                    ' Check if the thread is alive or dead
                    If Not DirectCast(LocalThreadList.Item(x), Thread).IsAlive Then

                        ' Remove the thread and victim from the list
                        LocalThreadList.RemoveAt(x)
                        LocalThreadList.RemoveAt(x)

                    End If

                Next

            End If

            ' Rest the thread
            Thread.Sleep(Globals.THREAD_REST_INTERVAL)

        End While

        ' Enable the start button
        Delegate_Sub_Enable_Blue_Button(btnEncStressStart, True)

        ' Re-enable ping type setting
        Delegate_Sub_Enable_Control(grpEncStressPingType, True)

    End Sub

    Private Sub EncClientMonitorWorker()

        ' Local variables
        Dim Output As String = ""
        Dim TokenizedOutput As String()
        Dim LocalClientList As New ArrayList
        Dim DataGridClientList As New ArrayList
        Dim CheckClient As String

        ' Loop until termination signal
        While Not TerminateSignal And Not InternalTermSignal

            ' Clear the local client list
            LocalClientList.Clear()

            ' Run encUtilCmd to retrieve the client list and capture the output
            WinOffline.Utility.RunCommand(Globals.DSMFolder + "bin\encUtilCmd.exe", "server -client -port " + Globals.ENCServerTCPPort,, Output)

            ' Tokenize the output
            TokenizedOutput = Output.Split(New String() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)

            ' Iterate the output
            For Each strLine As String In TokenizedOutput

                ' Check for a client listing
                If strLine.ToLower.Contains("enc gateway client :") Then

                    ' Add the client to the local client list
                    LocalClientList.Add(strLine.Substring(strLine.IndexOf(":") + 2))

                End If

            Next

            ' Check for an empty client list
            If LocalClientList.Count = 0 Then

                ' Signal internal termination
                InternalTermSignal = True

                ' Push a user alert message
                AlertBox.CreateUserAlert("The ENC gateway server is configured, but the client list is empty.")

                ' Change button availability
                Delegate_Sub_Enable_Blue_Button(btnEncStressStop, False)
                Delegate_Sub_Enable_Control(grpEncStressPingType, True)

            End If

            ' Iterate the client list
            For Each client In LocalClientList

                ' Check if the client is new to the data grid
                If Not GetEncClientList().Contains(client) Then

                    ' Grab the mutex for accessing the data grid
                    dgvEncStressTable_mutex.WaitOne()

                    ' Add the new client
                    Delegate_Sub_Add_DataGridView_Row(dgvEncStressTable, {client, 0, 0, 0, 0})

                    ' Release the data grid mutex
                    dgvEncStressTable_mutex.ReleaseMutex()

                End If

            Next

            ' Grab the current client list from the data grid
            DataGridClientList = GetEncClientList.Clone

            ' Iterate the data grid client list
            For x = DataGridClientList.Count - 1 To 0 Step -1

                ' Pop a client
                CheckClient = DataGridClientList.Item(x)

                ' Verify the client is also on the local client list
                If Not LocalClientList.Contains(CheckClient) Then

                    ' Grab the data grid mutex
                    dgvEncStressTable_mutex.WaitOne()

                    ' Remove the dead client
                    Delegate_Sub_Remove_DataGridView_Row(dgvEncStressTable, x)

                    ' Release the data grid mutex
                    dgvEncStressTable_mutex.ReleaseMutex()

                End If

            Next

            ' Update the status table text
            Delegate_Sub_Set_Text(grpEncStressStatus, "Status Table (" + LocalClientList.Count.ToString + " clients)")

            ' Rest with frequent checks for the terminate signal
            For i = 0 To 599

                ' Rest the thread
                Thread.Sleep(Globals.THREAD_REST_INTERVAL)

                ' Check for termination signal
                If TerminateSignal Then Exit While

            Next

        End While

        ' Reset status table text
        Delegate_Sub_Set_Text(grpEncStressStatus, "Status Table")

    End Sub

    Private Function GetEncClientList() As ArrayList

        ' Local variables
        Dim ClientList As New ArrayList

        ' Iterate each row of the data grid
        For Each row As DataGridViewRow In dgvEncStressTable.Rows

            ' Add to the local list
            ClientList.Add(row.Cells(0).Value)

        Next

        ' Return the list
        Return ClientList

    End Function

    Private Sub btnEncStressStart_Click(sender As Object, e As EventArgs) Handles btnEncStressStart.Click

        ' Verify the encClient process is running
        If Not WinOffline.Utility.IsProcessRunning("encClient") Then

            ' Push user alert
            AlertBox.CreateUserAlert("The ENC gateway client plugin (encClient.exe) is not running.")

            ' Return
            Return

        End If

        ' Verify an enc gateway server is configured
        If Not Globals.ENCGatewayServer.Length > 0 Then

            ' Push user alert
            AlertBox.CreateUserAlert("The ENC gateway server address is not configured.")

            ' Return
            Return

        End If

        ' Disable the ping type setting
        grpEncStressPingType.Enabled = False

        ' Check ping type user setting
        If rbnEncStressCafPing.Checked Then

            ' For caf ping, disable the total bytes column, as caf pings vary in size
            dgvEncStressTable.Columns(5).Visible = False

            ' Set autosize for the enc client to fill the extra space
            dgvEncStressTable.Columns(0).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            dgvEncStressTable.Columns(1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            dgvEncStressTable.Columns(2).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            dgvEncStressTable.Columns(3).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            dgvEncStressTable.Columns(4).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells

        Else

            ' For camping, enable the total bytes column, as cam pings are specific sizes
            dgvEncStressTable.Columns(5).Visible = True

            ' Set autosize for the number of bytes to fill the extra space
            dgvEncStressTable.Columns(0).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            dgvEncStressTable.Columns(1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            dgvEncStressTable.Columns(2).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            dgvEncStressTable.Columns(3).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            dgvEncStressTable.Columns(4).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            dgvEncStressTable.Columns(5).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells

        End If

        ' Reset internal terminate signal
        InternalTermSignal = False

        ' Clear the current data grid
        Delegate_Sub_Clear_DataGridView_Rows(dgvEncStressTable)

        ' Set button availability
        Delegate_Sub_Enable_Blue_Button(btnEncStressStart, False)
        Delegate_Sub_Enable_Blue_Button(btnEncStressStop, True)

        ' Initialize and start the threads
        EncClientMonitorThread = New Thread(AddressOf EncClientMonitorWorker)
        EncOverdriveThread = New Thread(AddressOf EncOverdriveWorker)
        EncClientMonitorThread.Start()
        EncOverdriveThread.Start()

    End Sub

    Private Sub btnEncStressStop_Click(sender As Object, e As EventArgs) Handles btnEncStressStop.Click

        ' Disable the stop button
        Delegate_Sub_Enable_Blue_Button(btnEncStressStop, False)

        ' Set internal termination signal
        InternalTermSignal = True

    End Sub

    Private Sub btnEncStressTargetThreadMinus_Click(sender As Object, e As EventArgs) Handles btnEncStressTargetThreadMinus.Click
        Dim numTargets As Integer = Integer.Parse(txtEncStressNumTargetThreads.Text)
        If numTargets > 1 Then numTargets -= 1
        txtEncStressNumTargetThreads.Text = numTargets.ToString
    End Sub

    Private Sub btnEncStressTargetThreadPlus_Click(sender As Object, e As EventArgs) Handles btnEncStressTargetThreadPlus.Click
        Dim numTargets As Integer = Integer.Parse(txtEncStressNumTargetThreads.Text)
        If numTargets < 100 Then numTargets += 1
        txtEncStressNumTargetThreads.Text = numTargets.ToString
    End Sub

    Private Sub txtEncStressNumTargets_KeyDown(sender As Object, e As KeyEventArgs) Handles txtEncStressNumTargetThreads.KeyDown
        If e.KeyCode = Keys.Add Then
            btnEncStressTargetThreadPlus_Click(sender, e)
        ElseIf e.KeyCode = Keys.Subtract Then
            btnEncStressTargetThreadMinus_Click(sender, e)
        End If
    End Sub

    Private Sub btnEncStressPingThreadMinus_Click(sender As Object, e As EventArgs) Handles btnEncStressPingThreadMinus.Click
        Dim numThreads As Integer = Integer.Parse(txtEncStressNumPingThreads.Text)
        If numThreads > 1 Then numThreads -= 1
        txtEncStressNumPingThreads.Text = numThreads.ToString
    End Sub

    Private Sub btnEncStressPingThreadPlus_Click(sender As Object, e As EventArgs) Handles btnEncStressPingThreadPlus.Click
        Dim numThreads As Integer = Integer.Parse(txtEncStressNumPingThreads.Text)
        If numThreads < 50 Then numThreads += 1
        txtEncStressNumPingThreads.Text = numThreads.ToString
    End Sub

    Private Sub txtNumThreads_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txtEncStressNumPingThreads.KeyDown
        If e.KeyCode = Keys.Add Then
            btnEncStressPingThreadPlus_Click(sender, e)
        ElseIf e.KeyCode = Keys.Subtract Then
            btnEncStressPingThreadMinus_Click(sender, e)
        End If
    End Sub

    Private Sub btnEncStressNumPingMinus_Click(sender As Object, e As EventArgs) Handles btnEncStressNumPingMinus.Click
        Dim numPings As Integer = Integer.Parse(txtEncStressNumPings.Text)
        If numPings <= 10 And numPings > 1 Then
            numPings -= 1
        ElseIf numPings <= 100 And numPings > 1 Then
            numPings -= 10
        ElseIf numPings > 100 And numPings > 1 Then
            numPings -= 100
        End If
        txtEncStressNumPings.Text = numPings.ToString
    End Sub

    Private Sub btnEncStressNumPingPlus_Click(sender As Object, e As EventArgs) Handles btnEncStressNumPingPlus.Click
        Dim numPings As Integer = Integer.Parse(txtEncStressNumPings.Text)
        If numPings < 10 Then
            numPings += 1
        ElseIf numPings < 100 Then
            numPings += 10
        ElseIf numPings < 10000 Then
            numPings += 100
        End If
        txtEncStressNumPings.Text = numPings.ToString
    End Sub

    Private Sub txtEncStressNumPings_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txtEncStressNumPings.KeyDown
        If e.KeyCode = Keys.Add Then
            btnEncStressNumPingPlus_Click(sender, e)
        ElseIf e.KeyCode = Keys.Subtract Then
            btnEncStressNumPingMinus_Click(sender, e)
        End If
    End Sub

    Private Sub btnEncStressMsgSizeMinus_Click(sender As Object, e As EventArgs) Handles btnEncStressMsgSizeMinus.Click
        Dim numBytes As Integer = Integer.Parse(txtEncStressMsgSize.Text)
        If numBytes > 2 Then numBytes /= 2
        txtEncStressMsgSize.Text = numBytes.ToString
    End Sub

    Private Sub btnEncStressMsgSizePlus_Click(sender As Object, e As EventArgs) Handles btnEncStressMsgSizePlus.Click
        Dim numBytes As Integer = Integer.Parse(txtEncStressMsgSize.Text)
        If numBytes < 16777216 Then numBytes *= 2
        txtEncStressMsgSize.Text = numBytes.ToString
    End Sub

    Private Sub txtMsgSize_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txtEncStressMsgSize.KeyDown
        If e.KeyCode = Keys.Add Then
            btnEncStressMsgSizePlus_Click(sender, e)
        ElseIf e.KeyCode = Keys.Subtract Then
            btnEncStressMsgSizeMinus_Click(sender, e)
        End If
    End Sub

    Private Sub btnEncStressReplyTimeoutMinus_Click(sender As Object, e As EventArgs) Handles btnEncStressReplyTimeoutMinus.Click
        Dim timeout As Integer = Integer.Parse(txtEncStressReplyTimeout.Text)
        If timeout > 1 Then timeout -= 1
        txtEncStressReplyTimeout.Text = timeout.ToString
    End Sub

    Private Sub btnEncStressReplyTimeoutPlus_Click(sender As Object, e As EventArgs) Handles btnEncStressReplyTimeoutPlus.Click
        Dim timeout As Integer = Integer.Parse(txtEncStressReplyTimeout.Text)
        If timeout < 60 Then timeout += 1
        txtEncStressReplyTimeout.Text = timeout.ToString
    End Sub

    Private Sub btnEncStressReplyTimeoutPlus_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txtEncStressReplyTimeout.KeyDown
        If e.KeyCode = Keys.Add Then
            btnEncStressReplyTimeoutPlus_Click(sender, e)
        ElseIf e.KeyCode = Keys.Subtract Then
            btnEncStressReplyTimeoutMinus_Click(sender, e)
        End If
    End Sub

    Private Sub btnEncStressDelayTimeMinus_Click(sender As Object, e As EventArgs) Handles btnEncStressDelayTimeMinus.Click
        Dim waittime As Integer = Integer.Parse(txtEncStressDelayTime.Text)
        If waittime > 0 Then waittime -= 1
        txtEncStressDelayTime.Text = waittime.ToString
    End Sub

    Private Sub btnEncStressDelayTimePlus_Click(sender As Object, e As EventArgs) Handles btnEncStressDelayTimePlus.Click
        Dim waittime As Integer = Integer.Parse(txtEncStressDelayTime.Text)
        If waittime < 60 Then waittime += 1
        txtEncStressDelayTime.Text = waittime.ToString
    End Sub

    Private Sub btnEncStressWaitTimePlus_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txtEncStressDelayTime.KeyDown
        If e.KeyCode = Keys.Add Then
            btnEncStressDelayTimePlus_Click(sender, e)
        ElseIf e.KeyCode = Keys.Subtract Then
            btnEncStressDelayTimeMinus_Click(sender, e)
        End If
    End Sub

    Private Sub rbnEncStressCamPing_CheckedChanged(sender As Object, e As EventArgs) Handles rbnEncStressCamPing.CheckedChanged
        If rbnEncStressCamPing.Checked Then
            rbnEncStressCafPing.Checked = False
            txtEncStressMsgSize.Enabled = True
            btnEncStressMsgSizeMinus.Enabled = True
            btnEncStressMsgSizePlus.Enabled = True
            txtEncStressDelayTime.Enabled = True
            btnEncStressDelayTimeMinus.Enabled = True
            btnEncStressDelayTimePlus.Enabled = True
        End If
    End Sub

    Private Sub rbnEncStressCafPing_CheckedChanged(sender As Object, e As EventArgs) Handles rbnEncStressCafPing.CheckedChanged
        If rbnEncStressCafPing.Checked Then
            txtEncStressMsgSize.Enabled = False
            btnEncStressMsgSizeMinus.Enabled = False
            btnEncStressMsgSizePlus.Enabled = False
            txtEncStressDelayTime.Enabled = False
            btnEncStressDelayTimeMinus.Enabled = False
            btnEncStressDelayTimePlus.Enabled = False
        End If
    End Sub

End Class