'****************************** Class Header *******************************\
' Project Name: WinOffline
' Class Name:   WinOfflineUI
' File Name:    WinOfflineUI-Sql-GroupEvalGrid.vb
' Author:       Brian Fontana
'***************************************************************************/

Imports System.Data
Imports System.Data.SqlClient
Imports System.Threading
Imports System.Windows.Forms

Partial Public Class WinOfflineUI

    Private Shared GroupEvalGridThread As Thread
    Private Shared GroupEvalUpdateThread As Thread
    Private EngineGrid As New List(Of List(Of String))
    Private OriginalValues As New List(Of String)
    Private ChangeTracker As New List(Of String)
    Private SkipSelectionChange As Boolean = False

    Private Sub InitSqlGroupEvalGrid()

        ' Disable buttons
        Delegate_Sub_Enable_Red_Button(btnSqlDisconnectGroupEvalGrid, False)
        Delegate_Sub_Enable_Blue_Button(btnSqlRefreshGroupEvalGrid, False)
        Delegate_Sub_Enable_Blue_Button(btnSqlExportGroupEvalGrid, False)
        Delegate_Sub_Enable_CadetBlue_Button(btnGroupEvalGridCommit, False)
        Delegate_Sub_Enable_Peru_Button(btnGroupEvalGridPreview, False)
        Delegate_Sub_Enable_LightCoral_Button(btnGroupEvalGridDiscard, False)

        ' Set grid properties
        dgvGroupEvalGrid.AllowUserToAddRows = False
        dgvGroupEvalGrid.AllowUserToDeleteRows = False
        dgvGroupEvalGrid.AllowUserToResizeRows = False
        dgvGroupEvalGrid.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
        dgvGroupEvalGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgvGroupEvalGrid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
        dgvGroupEvalGrid.BorderStyle = BorderStyle.Fixed3D
        dgvGroupEvalGrid.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText
        dgvGroupEvalGrid.ColumnHeadersDefaultCellStyle.Font = New Drawing.Font("Calibri", 9)
        dgvGroupEvalGrid.ColumnHeadersDefaultCellStyle.WrapMode = False
        dgvGroupEvalGrid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        dgvGroupEvalGrid.DefaultCellStyle.BackColor = Drawing.Color.Beige
        dgvGroupEvalGrid.DefaultCellStyle.Font = New Drawing.Font("Calibri", 9)
        dgvGroupEvalGrid.DefaultCellStyle.WrapMode = False
        dgvGroupEvalGrid.EnableHeadersVisualStyles = False
        dgvGroupEvalGrid.MultiSelect = False
        dgvGroupEvalGrid.ReadOnly = False
        dgvGroupEvalGrid.RowHeadersVisible = False
        dgvGroupEvalGrid.ScrollBars = ScrollBars.Both
        dgvGroupEvalGrid.ShowCellErrors = False
        dgvGroupEvalGrid.ShowCellToolTips = False
        dgvGroupEvalGrid.ShowEditingIcon = False
        dgvGroupEvalGrid.ShowRowErrors = False

    End Sub

    Private Sub GroupEvalGridWorker(ByVal ConnectionString As String)

        ' Local variables
        Dim DbConnection As SqlConnection = New SqlConnection(ConnectionString)
        Dim CallStack As String = "GroupEvalGridWorker --> "
        Dim dgvCmbCell As DataGridViewComboBoxCell

        ' Encapsulate grid worker
        Try

            ' Open sql connection
            DbConnection.Open()

            ' Write debug
            Delegate_Sub_Append_Text(rtbDebug, CallStack + "Connected successful: " + SqlServer)

            ' Reveal progress bar
            grpSqlGroupEvalGrid.Invoke(Sub() grpSqlGroupEvalGrid.Height = grpSqlGroupEvalGrid.Height - prgGroupEvalGrid.Height - 4)

            ' Query group evaluation chart
            SqlPopulateGridWorker(CallStack,
                                  DbConnection,
                                  "select eng.engine_uuid as 'Engine UUID', eng.label as 'Engine Name', gd.eval_freq as 'Eval Interval (s)', gd.group_uuid as 'Group UUID', gd.label as 'Group Name', dateadd(s, datediff(s, getutcdate(), getdate()), dateadd(s, gd.last_eval_date_time, '19700101')) as 'Last Evaluated', qd.label as 'Query Name', dateadd(s, datediff(s, getutcdate(), getdate()), dateadd(s, qd.creation_date, '19700101')) as 'Query Creation Date' from ca_group_def gd with (nolock) inner join ca_query_def qd with (nolock) on gd.query_uuid=qd.query_uuid left join ca_engine eng with (nolock) on gd.evaluation_uuid=eng.engine_uuid order by gd.label",
                                  dgvGroupEvalGrid,
                                  "Group Name",,,
                                  prgGroupEvalGrid,
                                  grpSqlGroupEvalGrid,
                                  New List(Of String) From {"Group UUID", "Engine UUID"})

            ' Query available engines
            EngineGrid = SqlSelectGrid(CallStack, DbConnection, "select engine_uuid as 'Engine UUID', label as 'Engine Name' from ca_engine where domain_uuid in (select set_val_uuid from ca_settings where set_id=1)")

            ' Insert new data grid column
            Delegate_Sub_Insert_DataGridView_ComboBoxColumn(dgvGroupEvalGrid, dgvGroupEvalGrid.Columns("Engine Name").Index, "EvalEngine", "Evaluation Engine")

            ' Hide column
            Delegate_Sub_Hide_DataGridView_Column(dgvGroupEvalGrid, "Engine Name")

            ' Verify engine grid is available (terminate or cancel signal was given)
            If EngineGrid IsNot Nothing Then

                ' Iterate data grid rows
                For Each dgvRow As DataGridViewRow In dgvGroupEvalGrid.Rows

                    ' Create a new data grid combo box
                    dgvCmbCell = New DataGridViewComboBoxCell

                    ' Add "All Engines" option
                    dgvCmbCell.Items.Add("All Engines")

                    ' Iterate engines
                    For Each EngineList As List(Of String) In EngineGrid

                        ' Add engines to combo box (0=Engine UUID, 1=Engine Name)
                        dgvCmbCell.Items.Add(EngineList(1).ToString)

                        ' Check if engine matches row specification
                        If EngineList(0).ToString.Equals(dgvRow.Cells("Engine UUID").Value.ToString) Then

                            ' Set default value
                            dgvCmbCell.Value = EngineList(1).ToString

                        End If

                    Next

                    ' Check for "All Engines" setting (EngineUUID=NULL, EngineName=NULL, EvalInterval=0)
                    If dgvRow.Cells("Engine UUID").Value.ToString.ToLower.Equals("null") Then

                        ' Set "all engines"
                        dgvCmbCell.Value = "All Engines"

                    End If

                    ' Change eval engine cell to the new combo box
                    dgvRow.Cells("EvalEngine") = dgvCmbCell

                    ' Store original values for later comparison
                    OriginalValues.Add(dgvRow.Cells("Group UUID").Value.ToString + "," +
                                       dgvRow.Cells("EvalEngine").Value.ToString + "," +
                                       dgvRow.Cells("Eval Interval (s)").Value.ToString)

                Next

            End If

            ' Re-init progress bar
            prgGroupEvalGrid.Invoke(Sub() prgGroupEvalGrid.Value = 0)
            prgGroupEvalGrid.Invoke(Sub() prgGroupEvalGrid.Maximum = 10)

            ' Bold editable columns
            dgvGroupEvalGrid.Columns("EvalEngine").DefaultCellStyle.Font = New System.Drawing.Font(dgvGroupEvalGrid.DefaultCellStyle.Font, Drawing.FontStyle.Bold)
            dgvGroupEvalGrid.Columns("Eval Interval (s)").DefaultCellStyle.Font = New System.Drawing.Font(dgvGroupEvalGrid.DefaultCellStyle.Font, Drawing.FontStyle.Bold)
            prgGroupEvalGrid.Invoke(Sub() prgGroupEvalGrid.Increment(1))

            ' Color editable columns
            dgvGroupEvalGrid.Columns("EvalEngine").DefaultCellStyle.BackColor = Drawing.Color.Yellow
            dgvGroupEvalGrid.Columns("Eval Interval (s)").DefaultCellStyle.BackColor = Drawing.Color.Yellow
            prgGroupEvalGrid.Invoke(Sub() prgGroupEvalGrid.Increment(1))

            ' Resize rows and columns after changes (and increment progress bar)
            Delegate_Sub_UnSet_DataGrid_Fill_Column(dgvGroupEvalGrid, "Group Name")
            prgGroupEvalGrid.Invoke(Sub() prgGroupEvalGrid.Increment(1))
            Delegate_Sub_Resize_DataGrid_Rows(dgvGroupEvalGrid, DataGridViewAutoSizeRowsMode.AllCells)
            prgGroupEvalGrid.Invoke(Sub() prgGroupEvalGrid.Increment(1))
            Delegate_Sub_Resize_DataGrid_Column(dgvGroupEvalGrid, "EvalEngine", DataGridViewAutoSizeColumnsMode.AllCells)
            prgGroupEvalGrid.Invoke(Sub() prgGroupEvalGrid.Increment(1))
            Delegate_Sub_Resize_DataGrid_Column(dgvGroupEvalGrid, "Eval Interval (s)", DataGridViewAutoSizeColumnsMode.AllCells)
            prgGroupEvalGrid.Invoke(Sub() prgGroupEvalGrid.Increment(1))
            Delegate_Sub_Resize_DataGrid_Column(dgvGroupEvalGrid, "Group Name", 250)
            prgGroupEvalGrid.Invoke(Sub() prgGroupEvalGrid.Increment(1))
            Delegate_Sub_Resize_DataGrid_Column(dgvGroupEvalGrid, "Last Evaluated", DataGridViewAutoSizeColumnsMode.AllCells)
            prgGroupEvalGrid.Invoke(Sub() prgGroupEvalGrid.Increment(1))
            Delegate_Sub_Resize_DataGrid_Column(dgvGroupEvalGrid, "Query Name", 300)
            prgGroupEvalGrid.Invoke(Sub() prgGroupEvalGrid.Increment(1))
            Delegate_Sub_Resize_DataGrid_Column(dgvGroupEvalGrid, "Query Creation Date", DataGridViewAutoSizeColumnsMode.AllCells)
            prgGroupEvalGrid.Invoke(Sub() prgGroupEvalGrid.Increment(1))

            ' Hide progress bar
            grpSqlGroupEvalGrid.Invoke(Sub() grpSqlGroupEvalGrid.Height = grpSqlGroupEvalGrid.Height + prgGroupEvalGrid.Height + 4)

        Catch ex As Exception

            ' Write debug
            Delegate_Sub_Append_Text(rtbDebug, CallStack + "Exception:" + Environment.NewLine + ex.Message)
            Delegate_Sub_Append_Text(rtbDebug, CallStack + "Stack trace: " + Environment.NewLine + ex.StackTrace)

        Finally

            ' Check if database connection is open
            If Not DbConnection.State = ConnectionState.Closed Then

                ' Close the database connection
                DbConnection.Close()

                ' Write debug
                Delegate_Sub_Append_Text(rtbDebug, CallStack + "Database connection closed.")

            End If

            ' Enable buttons
            Delegate_Sub_Enable_Yellow_Button(btnSqlRefreshGroupEvalGrid, True)
            Delegate_Sub_Enable_Tan_Button(btnSqlExportGroupEvalGrid, True)

        End Try

    End Sub

    Private Sub dgvGroupEvalGrid_CellEnter(ByVal sender As System.Object, ByVal e As DataGridViewCellEventArgs) Handles dgvGroupEvalGrid.CellEnter

        ' Send edit key (avoid need for two clicks)
        SendKeys.Send("{F4}")

    End Sub

    Private Sub dgvGroupEvalGrid_EditingControlShowing(ByVal sender As System.Object, ByVal e As DataGridViewEditingControlShowingEventArgs) Handles dgvGroupEvalGrid.EditingControlShowing

        ' Local variables
        Dim dgvCmbCell As ComboBox

        ' Check for evaluation engine column
        If dgvGroupEvalGrid.CurrentCell.ColumnIndex = dgvGroupEvalGrid.Columns("EvalEngine").Index Then

            ' Convert control to combo box
            dgvCmbCell = CType(e.Control, ComboBox)

            ' Check combo box
            If dgvCmbCell IsNot Nothing Then

                ' Remove handler, if existing, to prevent duplication
                RemoveHandler dgvCmbCell.SelectionChangeCommitted, New EventHandler(AddressOf EvalEngine_SelectionChangeCommitted)

                ' Add selection changed handler
                AddHandler dgvCmbCell.SelectionChangeCommitted, New EventHandler(AddressOf EvalEngine_SelectionChangeCommitted)

            End If

        End If

    End Sub

    Private Sub EvalEngine_SelectionChangeCommitted(ByVal sender As System.Object, ByVal e As System.EventArgs)

        ' Local variables
        Dim dgvCmbCell As ComboBox = CType(sender, ComboBox)

        ' Commit cell value change (which will fire cell value changed event)
        dgvGroupEvalGrid.Rows(dgvGroupEvalGrid.CurrentCell.RowIndex).Cells("EvalEngine").Value = dgvCmbCell.SelectedItem.ToString

    End Sub

    Private Sub dgvGroupEvalGrid_CellValueChanged(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles dgvGroupEvalGrid.CellValueChanged

        ' Local variables
        Dim RowIndex As Integer = dgvGroupEvalGrid.CurrentCell.RowIndex
        Dim ColIndex As Integer = dgvGroupEvalGrid.CurrentCell.ColumnIndex
        Dim GroupUUID As String = dgvGroupEvalGrid.Rows(RowIndex).Cells("Group UUID").Value.ToString
        Dim EngineUUID As String = dgvGroupEvalGrid.Rows(RowIndex).Cells("Engine UUID").Value.ToString
        Dim EvalEngine As String = dgvGroupEvalGrid.Rows(RowIndex).Cells("EvalEngine").Value.ToString
        Dim EvalInterval As String = dgvGroupEvalGrid.Rows(RowIndex).Cells("Eval Interval (s)").Value.ToString
        Dim NumberCheck As Integer = -1

        ' Skip cascading calls
        If SkipSelectionChange Then Return

        ' Lock cascading calls
        SkipSelectionChange = True

        Try

            ' Match column having value change
            If ColIndex = dgvGroupEvalGrid.Columns("EvalEngine").Index Then

                ' Check if "All Engines" was selected
                If EvalEngine.Equals("All Engines") Then

                    ' Update engine uuid column
                    dgvGroupEvalGrid.Rows(RowIndex).Cells("Engine UUID").Value = "Null"

                    ' Check if evaluation interval is 0
                    If Not dgvGroupEvalGrid.Rows(RowIndex).Cells("Eval Interval (s)").Value.ToString.Equals("0") Then

                        ' Set eval interval to 0
                        dgvGroupEvalGrid.Rows(RowIndex).Cells("Eval Interval (s)").Value = "0"

                    End If

                Else

                    ' Update engine uuid column
                    dgvGroupEvalGrid.Rows(RowIndex).Cells("Engine UUID").Value = GetEngineUUID(EvalEngine)

                    ' Check if eval interval is 0 (since this is non-"All Engines" case and valid values are 60-604800)
                    If dgvGroupEvalGrid.Rows(RowIndex).Cells("Eval Interval (s)").Value.ToString.Equals("0") Then

                        ' Check if original value is 0
                        If GetOrigEvalInterval(GroupUUID) = 0 Then

                            ' Set default value (1 day = 86400 seconds)
                            dgvGroupEvalGrid.Rows(RowIndex).Cells("Eval Interval (s)").Value = 86400

                        Else

                            ' Restore engine eval interval
                            dgvGroupEvalGrid.Rows(RowIndex).Cells("Eval Interval (s)").Value = GetOrigEvalInterval(GroupUUID)

                        End If

                    End If

                End If

            ElseIf ColIndex = dgvGroupEvalGrid.Columns("Eval Interval (s)").Index Then

                ' Check value is a number
                If Integer.TryParse(EvalInterval, NumberCheck) Then

                    ' Check number is within valid range
                    If NumberCheck = 0 AndAlso Not dgvGroupEvalGrid.Rows(RowIndex).Cells("EvalEngine").Value.Equals("All Engines") Then

                        ' Alert user
                        MsgBox("An evaluation interval of '0' is only valid when the evaluation engine is set to 'All Engines'.", MsgBoxStyle.Exclamation, "Nope")

                        ' Restore engine eval interval
                        dgvGroupEvalGrid.Rows(RowIndex).Cells("Eval Interval (s)").Value = GetOrigEvalInterval(GroupUUID)

                    ElseIf NumberCheck <> 0 AndAlso dgvGroupEvalGrid.Rows(RowIndex).Cells("EvalEngine").Value.Equals("All Engines") Then

                        ' Alert user
                        MsgBox("Evaluation interval must be '0' when the evaluation engine is set to 'All Engines'.", MsgBoxStyle.Exclamation, "Nope")

                        ' Restore engine eval interval
                        dgvGroupEvalGrid.Rows(RowIndex).Cells("Eval Interval (s)").Value = 0

                    ElseIf Not (NumberCheck >= 60 AndAlso NumberCheck <= 604800) Then

                        ' Alert user
                        MsgBox("Evaluation interval must be an integer between 60 seconds and 604800 seconds.", MsgBoxStyle.Exclamation, "Nope")

                        ' Restore engine eval interval
                        dgvGroupEvalGrid.Rows(RowIndex).Cells("Eval Interval (s)").Value = GetOrigEvalInterval(GroupUUID)

                    End If

                Else

                    ' Alert user
                    MsgBox("Evaluation interval must be an integer between 60 seconds and 604800 seconds.", MsgBoxStyle.Exclamation, "Nope")

                    ' Restore engine eval interval
                    dgvGroupEvalGrid.Rows(RowIndex).Cells(ColIndex).Value = GetOrigEvalInterval(GroupUUID)

                End If

            End If

            ' Check if evaluation engine has changed
            If dgvGroupEvalGrid.Rows(RowIndex).Cells("EvalEngine").Value.Equals(GetOrigEvalEngine(GroupUUID)) Then

                ' Back to orignal value -- restore font color
                dgvGroupEvalGrid.Rows(RowIndex).Cells("EvalEngine").Style.Font = New System.Drawing.Font(dgvGroupEvalGrid.DefaultCellStyle.Font, Drawing.FontStyle.Bold)
                dgvGroupEvalGrid.Rows(RowIndex).Cells("EvalEngine").Style.ForeColor = System.Drawing.Color.Black

            Else

                ' Change font color
                dgvGroupEvalGrid.Rows(RowIndex).Cells("EvalEngine").Style.Font = New System.Drawing.Font(dgvGroupEvalGrid.DefaultCellStyle.Font, Drawing.FontStyle.Bold Or Drawing.FontStyle.Italic)
                dgvGroupEvalGrid.Rows(RowIndex).Cells("EvalEngine").Style.ForeColor = System.Drawing.Color.Red

            End If

            ' Check if evaluation interval has changed
            If Integer.Parse(dgvGroupEvalGrid.Rows(RowIndex).Cells("Eval Interval (s)").Value) = GetOrigEvalInterval(GroupUUID) Then

                ' Back to orignal value -- restore font color
                dgvGroupEvalGrid.Rows(RowIndex).Cells("Eval Interval (s)").Style.Font = New System.Drawing.Font(dgvGroupEvalGrid.DefaultCellStyle.Font, Drawing.FontStyle.Bold)
                dgvGroupEvalGrid.Rows(RowIndex).Cells("Eval Interval (s)").Style.ForeColor = System.Drawing.Color.Black

            Else

                ' Change font
                dgvGroupEvalGrid.Rows(RowIndex).Cells("Eval Interval (s)").Style.Font = New System.Drawing.Font(dgvGroupEvalGrid.DefaultCellStyle.Font, Drawing.FontStyle.Bold Or Drawing.FontStyle.Italic)
                dgvGroupEvalGrid.Rows(RowIndex).Cells("Eval Interval (s)").Style.ForeColor = System.Drawing.Color.Red

            End If

            ' Update the tracker only if there are changes from original values
            If Not dgvGroupEvalGrid.Rows(RowIndex).Cells("EvalEngine").Value.Equals(GetOrigEvalEngine(GroupUUID)) OrElse
                Not Integer.Parse(dgvGroupEvalGrid.Rows(RowIndex).Cells("Eval Interval (s)").Value) = GetOrigEvalInterval(GroupUUID) Then

                ' Refresh values before updating tracker
                EngineUUID = dgvGroupEvalGrid.Rows(RowIndex).Cells("Engine UUID").Value.ToString
                EvalInterval = dgvGroupEvalGrid.Rows(RowIndex).Cells("Eval Interval (s)").Value.ToString

                ' Update change tracker
                RemoveFromChangeTracker(GroupUUID)
                AddtoChangeTracker(GroupUUID, EngineUUID, EvalInterval)

            Else

                ' Remove from change tracker
                RemoveFromChangeTracker(GroupUUID)

            End If

            ' Check if there are changes
            If ChangeTracker.Count > 0 Then

                ' Enable buttons
                Delegate_Sub_Enable_CadetBlue_Button(btnGroupEvalGridCommit, True)
                Delegate_Sub_Enable_Peru_Button(btnGroupEvalGridPreview, True)
                Delegate_Sub_Enable_LightCoral_Button(btnGroupEvalGridDiscard, True)

            Else

                ' Disable buttons
                Delegate_Sub_Enable_CadetBlue_Button(btnGroupEvalGridCommit, False)
                Delegate_Sub_Enable_Peru_Button(btnGroupEvalGridPreview, False)
                Delegate_Sub_Enable_LightCoral_Button(btnGroupEvalGridDiscard, False)

            End If

        Finally

            ' Unlock cascading calls
            SkipSelectionChange = False

        End Try

    End Sub

    Private Sub AddtoChangeTracker(ByVal GroupUUID As String,
                                   ByVal EngineUUID As String,
                                   ByVal EvalInterval As Integer)

        ' Iterate change tracker
        For i As Integer = 0 To ChangeTracker.Count - 1

            ' Check if passed group uuid is already sepcified
            If ChangeTracker.Item(i).ToString.Equals(GroupUUID) Then

                ' Remove the group uuid, engine uuid and eval interval
                ChangeTracker.RemoveAt(i)
                ChangeTracker.RemoveAt(i)
                ChangeTracker.RemoveAt(i)

            End If

        Next

        ' Add values to change tracker
        ChangeTracker.Add(GroupUUID)
        ChangeTracker.Add(EngineUUID)
        ChangeTracker.Add(EvalInterval.ToString)

    End Sub

    Private Sub RemoveFromChangeTracker(ByVal GroupUUID As String)

        ' Iterate change tracker
        For i As Integer = 0 To ChangeTracker.Count - 1

            ' Check if passed group uuid is already sepcified
            If ChangeTracker.Item(i).ToString.Equals(GroupUUID) Then

                ' Remove the group uuid, engine uuid and eval interval
                ChangeTracker.RemoveAt(i)
                ChangeTracker.RemoveAt(i)
                ChangeTracker.RemoveAt(i)

                ' Exit loop
                Exit For

            End If

        Next

    End Sub

    Private Function GetOrigEvalEngine(ByVal GroupUUID As String) As String

        ' Iterate original values
        For Each value As String In OriginalValues

            ' Check for group uuid match
            If GroupUUID.Equals(value.Substring(0, value.IndexOf(","))) Then

                ' Return
                Return value.Substring(value.IndexOf(",") + 1, value.LastIndexOf(",") - value.Substring(0, value.IndexOf(",")).Length - 1)

            End If

        Next

        ' Not found
        Return Nothing

    End Function

    Private Function GetOrigEvalInterval(ByVal GroupUUID As String) As Integer

        ' Iterate original values
        For Each value As String In OriginalValues

            ' Match by group uuid
            If GroupUUID.Equals(value.Substring(0, value.IndexOf(","))) Then

                ' Return eval interval
                Return Integer.Parse(value.Substring(value.LastIndexOf(",") + 1))

            End If

        Next

        ' Not found
        Return -1

    End Function

    Private Function GetEngineUUID(ByVal EngineName As String) As String

        ' Iterate engine grid (0=Engine UUID, 1=Engine Name)
        For Each EngineList As List(Of String) In EngineGrid

            ' Check for name match
            If EngineList(1).ToString.ToLower.Equals(EngineName.ToLower) Then

                ' Return
                Return EngineList(0)

            End If

        Next

        ' Return
        Return Nothing

    End Function

    Private Sub btnGroupEvalGridCommit_Click(sender As Object, e As EventArgs) Handles btnGroupEvalGridCommit.Click

        ' Disable buttons
        Delegate_Sub_Enable_CadetBlue_Button(btnGroupEvalGridCommit, False)
        Delegate_Sub_Enable_Peru_Button(btnGroupEvalGridPreview, False)
        Delegate_Sub_Enable_LightCoral_Button(btnGroupEvalGridDiscard, False)

        ' New thread to update SQL
        GroupEvalUpdateThread = New Thread(Sub() GroupEvalUpdateWorker(ConnectionString))
        GroupEvalUpdateThread.Start()

    End Sub

    Private Sub GroupEvalUpdateWorker(ByVal ConnectionString As String)

        ' Local variables
        Dim CallStack As String = "GroupEvalUpdateWorker --> "
        Dim QueryText As String
        Dim DatabaseConnection As SqlConnection = New SqlConnection(ConnectionString)
        Dim SqlCmd As SqlCommand
        Dim GroupUUID As String
        Dim EngineUUID As String
        Dim EvalInterval As String
        Dim GroupsAffected As Integer = 0
        Dim RecordsAffected As Integer = 0

        ' Encapsulate connection operation
        Try

            ' Open sql connection
            DatabaseConnection.Open()

            ' Write debug
            Delegate_Sub_Append_Text(rtbDebug, CallStack + "Connected successful: " + SqlServer)

            ' Iterate change tracker
            For i As Integer = 0 To ChangeTracker.Count - 1 Step 3

                ' Fetch values
                GroupUUID = ChangeTracker(i)
                EngineUUID = ChangeTracker(i + 1)
                EvalInterval = ChangeTracker(i + 2)

                ' Build query text
                QueryText = "update ca_group_def set eval_freq=" + EvalInterval + ", evaluation_uuid=" + EngineUUID +
                    " where group_uuid=" + GroupUUID

                ' Build sql command from user query
                SqlCmd = New SqlCommand(QueryText, DatabaseConnection)

                ' Write debug
                Delegate_Sub_Append_Text(rtbDebug, CallStack + "Execute statement:" + Environment.NewLine + QueryText)

                ' Execute query
                RecordsAffected += SqlCmd.ExecuteNonQuery

                ' Increment counter
                GroupsAffected += 1

            Next

            ' Alert user
            MsgBox(RecordsAffected.ToString + " record(s) affected, " + GroupsAffected.ToString + " group(s) affected.", MsgBoxStyle.Information, "Changes committed.")

        Catch ex As Exception

            ' Write debug
            Delegate_Sub_Append_Text(rtbDebug, CallStack + "Exception:" + Environment.NewLine + ex.Message)
            Delegate_Sub_Append_Text(rtbDebug, CallStack + "Stack trace: " + Environment.NewLine + ex.StackTrace)

        Finally

            ' Check if database connection is open
            If Not DatabaseConnection.State = ConnectionState.Closed Then

                ' Close the database connection
                DatabaseConnection.Close()

                ' Write debug
                Delegate_Sub_Append_Text(rtbDebug, CallStack + "Database connection closed.")

            End If

        End Try

        ' Disable buttons
        Delegate_Sub_Enable_Yellow_Button(btnSqlRefreshGroupEvalGrid, False)
        Delegate_Sub_Enable_Tan_Button(btnSqlExportGroupEvalGrid, False)
        Delegate_Sub_Enable_CadetBlue_Button(btnGroupEvalGridCommit, False)
        Delegate_Sub_Enable_Peru_Button(btnGroupEvalGridPreview, False)
        Delegate_Sub_Enable_LightCoral_Button(btnGroupEvalGridDiscard, False)

        ' Clear lists
        OriginalValues = New List(Of String)
        ChangeTracker = New List(Of String)

        ' Restart thread
        GroupEvalGridThread = New Thread(Sub() GroupEvalGridWorker(ConnectionString))
        GroupEvalGridThread.Start()

    End Sub

    Private Sub btnGroupEvalGridPreview_Click(sender As Object, e As EventArgs) Handles btnGroupEvalGridPreview.Click

        ' Local variables
        Dim GroupUUID As String
        Dim EngineUUID As String
        Dim EvalInterval As String
        Dim PreviewString As String = ""

        ' Iterate change tracker
        For i As Integer = 0 To ChangeTracker.Count - 1 Step 3

            ' Fetch values
            GroupUUID = ChangeTracker(i)
            EngineUUID = ChangeTracker(i + 1)
            EvalInterval = ChangeTracker(i + 2)

            ' Build query text
            PreviewString += "update ca_group_def set eval_freq=" + EvalInterval + ", evaluation_uuid=" + EngineUUID +
                " where group_uuid=" + GroupUUID + Environment.NewLine

        Next

        ' Show alert box
        AlertBox.CreateUserAlert(PreviewString, 0, HorizontalAlignment.Left, False, 9, True)

    End Sub

    Private Sub btnGroupEvalGridDiscard_Click(sender As Object, e As EventArgs) Handles btnGroupEvalGridDiscard.Click

        ' Clear change tracker
        ChangeTracker = New List(Of String)

        ' Perform refresh
        btnSqlRefreshGroupEvalGrid.PerformClick()

        ' Disable buttons
        Delegate_Sub_Enable_CadetBlue_Button(btnGroupEvalGridCommit, False)
        Delegate_Sub_Enable_Peru_Button(btnGroupEvalGridPreview, False)
        Delegate_Sub_Enable_LightCoral_Button(btnGroupEvalGridDiscard, False)

    End Sub

    Private Sub btnSqlConnectGroupEvalGrid_Click(sender As Object, e As EventArgs) Handles btnSqlConnectGroupEvalGrid.Click

        ' Perform SQL connection
        SqlConnect()

    End Sub

    Private Sub btnSqlDisconnectGroupEvalGrid_Click(sender As Object, e As EventArgs) Handles btnSqlDisconnectGroupEvalGrid.Click

        ' Perform disconnect method
        SqlDisconnect()

    End Sub

    Private Sub btnSqlRefreshGroupEvalGrid_Click(sender As Object, e As EventArgs) Handles btnSqlRefreshGroupEvalGrid.Click

        ' Disable buttons
        Delegate_Sub_Enable_Yellow_Button(btnSqlRefreshGroupEvalGrid, False)
        Delegate_Sub_Enable_Tan_Button(btnSqlExportGroupEvalGrid, False)
        Delegate_Sub_Enable_CadetBlue_Button(btnGroupEvalGridCommit, False)
        Delegate_Sub_Enable_LightCoral_Button(btnGroupEvalGridDiscard, False)

        ' Clear original values list
        OriginalValues = New List(Of String)

        ' Restart thread
        GroupEvalGridThread = New Thread(Sub() GroupEvalGridWorker(ConnectionString))
        GroupEvalGridThread.Start()

    End Sub

    Private Sub btnSqlExportGroupEvalGrid_Click(sender As Object, e As EventArgs) Handles btnSqlExportGroupEvalGrid.Click

        ' Local variables
        Dim saveFileDialog1 As New SaveFileDialog()
        Dim StateStreamWriter As System.IO.StreamWriter

        ' Set dialog properties
        saveFileDialog1.Filter = "CSV (Comma delimited)|*.csv"
        saveFileDialog1.Title = "Save a CSV File"

        ' Launch dialog and check result
        If saveFileDialog1.ShowDialog() = DialogResult.Cancel Then Return

        ' Encapsulate file operation
        Try

            ' Open output stream
            StateStreamWriter = New System.IO.StreamWriter(saveFileDialog1.FileName, False)

            ' Iterate datagrid column headers
            For Each dgvColumn As DataGridViewColumn In dgvGroupEvalGrid.Columns

                ' Write values
                StateStreamWriter.Write(dgvColumn.HeaderText + ",")

            Next

            ' Write newline
            StateStreamWriter.Write(Environment.NewLine)

            ' Iterate datagrid rows
            For Each dgvRecord As DataGridViewRow In dgvGroupEvalGrid.Rows

                ' Iterate cells
                For Each CellItem As DataGridViewCell In dgvRecord.Cells

                    ' Write values
                    StateStreamWriter.Write(CellItem.Value.ToString + ",")

                Next

                ' Write newline
                StateStreamWriter.Write(Environment.NewLine)

            Next

            ' Close output stream
            StateStreamWriter.Close()

        Catch ex As Exception

            ' Push user alert
            AlertBox.CreateUserAlert("Export failed." + Environment.NewLine + Environment.NewLine + "Exception: " + ex.Message)

        End Try

    End Sub

    Private Sub btnSqlExitGroupEvalGrid_Click(sender As Object, e As EventArgs) Handles btnSqlExitGroupEvalGrid.Click

        ' Close the dialog
        Close()

    End Sub

End Class