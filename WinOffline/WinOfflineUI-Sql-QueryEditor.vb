Imports System.Threading
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Text.RegularExpressions

Partial Public Class WinOfflineUI

    Private Shared SqlQueryEditorThread As Thread
    Private Shared CancelQuery As Boolean = False

    Private Sub InitSqlQueryEditor()

        ' Disable buttons
        Delegate_Sub_Enable_Red_Button(btnSqlDisconnectQueryEditor, False)
        Delegate_Sub_Enable_Yellow_Button(btnSqlSubmitQueryEditor, False)
        Delegate_Sub_Enable_Tan_Button(btnSqlCancelQueryEditor, False)

    End Sub

    Private Sub btnSqlConnectQueryEditor_Click(sender As Object, e As EventArgs) Handles btnSqlConnectQueryEditor.Click

        ' Perform SQL connection
        SqlConnect()

    End Sub

    Private Sub btnSqlDisconnectQueryEditor_Click(sender As Object, e As EventArgs) Handles btnSqlDisconnectQueryEditor.Click

        ' Perform disconnect method
        SqlDisconnect()

    End Sub

    Private Sub btnSqlSubmitQueryEditor_Click(sender As Object, e As EventArgs) Handles btnSqlSubmitQueryEditor.Click

        ' Local variables
        Dim ConnectionString As String

        ' Build connection string
        ConnectionString = BuildSqlConnectionString()

        ' Check connection string
        If ConnectionString Is Nothing Then

            ' Return
            Return

        End If

        ' Set sql query editor function availability
        Delegate_Sub_Enable_Control(rtbSqlQuery, False)
        Delegate_Sub_Enable_Yellow_Button(btnSqlSubmitQueryEditor, False)
        Delegate_Sub_Enable_Tan_Button(btnSqlCancelQueryEditor, True)
        Delegate_Select_Button(btnSqlCancelQueryEditor)

        ' Initialize and start a submit thread
        SqlQueryEditorThread = New Thread(Sub() SqlQueryEditorWorker(ConnectionString, Callback_Function_Read_Selected_Text(rtbSqlQuery)))
        SqlQueryEditorThread.Start()

    End Sub

    Private Sub btnSqlCancelQueryEditor_Click(sender As Object, e As EventArgs) Handles btnSqlCancelQueryEditor.Click

        ' Set cancel flag
        CancelQuery = True

        ' Update results to reflect cancellation
        Delegate_Sub_Set_Text(txtSqlMessage, "Query cancelled by user.")

        ' Change the active tab to the messages tab
        Delegate_Sub_Set_Active_Tab(TabControlSql, TabPageMessages)

        ' Disable button
        Delegate_Sub_Enable_Tan_Button(btnSqlCancelQueryEditor, False)

    End Sub

    Private Sub btnSqlExitQueryEditor_Click(sender As Object, e As EventArgs) Handles btnSqlExitQueryEditor.Click

        ' Close the dialog
        Close()

    End Sub

    Private Sub SqlQueryEditorWorker(ByVal ConnectionString As String, ByVal QueryText As String)

        ' Local variables
        Dim DbConnection As SqlConnection
        Dim SqlBatches As String()
        Dim SqlCmd As SqlCommand
        Dim SqlData As SqlDataReader
        Dim SqlResultRow As New ArrayList
        Dim CellCounter As Integer = 0
        Dim SqlResultGrid As DataGridView
        Dim TotalRecordReturnCount As Integer = 0
        Dim BatchRecordReturnCount As Integer = 0
        Dim pnlTableLayout As New TableLayoutPanel
        Dim CallStack As String = "SqlQueryEditorWorker --> "

        ' Initialize table layout panel
        pnlTableLayout.Dock = DockStyle.Fill
        pnlTableLayout.AutoScroll = True
        pnlTableLayout.AutoSize = True
        pnlTableLayout.AutoSizeMode = AutoSizeMode.GrowAndShrink

        ' Results tab -- Reset tab text
        Delegate_Sub_Set_Text(TabPageGrid, "Results")

        ' Results tab -- Clear all child controls
        Delegate_Sub_Remove_Child_Controls(TabPageGrid)

        ' Results tab -- Add new table layout panel
        Delegate_Sub_Add_Control(TabPageGrid, pnlTableLayout)

        ' Messages tab -- Clear all messages
        Delegate_Sub_Set_Text(txtSqlMessage, "")

        ' Switch to messages tab
        Delegate_Sub_Set_Active_Tab(TabControlSql, TabPageMessages)

        ' Encapsulate sql operation
        Try

            ' Create sql connection
            DbConnection = New SqlConnection(ConnectionString)

            ' Attach informational message handler to connection
            AddHandler DbConnection.InfoMessage, New SqlInfoMessageEventHandler(AddressOf OnSqlQueryEditorInfoMessage)

            ' Open sql connection
            DbConnection.Open()

            ' Write debug
            Delegate_Sub_Append_Text(rtbDebug, CallStack + "Connected successful: " + SqlServer)

            ' Enumerate sql batches based on "GO"
            SqlBatches = Regex.Split(QueryText, "^\s*GO\s*\d*\s*($|\-\-.*$)", RegexOptions.Multiline Or RegexOptions.IgnorePatternWhitespace Or RegexOptions.IgnoreCase)

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
                If TerminateSignal Or CancelSignal Or CancelQuery Then

                    ' Close sql data reader
                    SqlData.Close()

                    ' Close the database connection
                    DbConnection.Close()

                    ' Write debug
                    Delegate_Sub_Append_Text(rtbDebug, CallStack + "Database connection closed.")

                    ' Return
                    Return

                End If

                ' Loop through sql data results
                Do

                    ' Check if data returns any rows
                    If SqlData.HasRows Then

                        ' Initialize a new data grid
                        SqlResultGrid = New DataGridView

                        ' Set data grid properties
                        SqlResultGrid.AllowUserToAddRows = False
                        SqlResultGrid.AllowUserToDeleteRows = False
                        SqlResultGrid.AllowUserToResizeRows = False
                        SqlResultGrid.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
                        SqlResultGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                        SqlResultGrid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                        SqlResultGrid.BorderStyle = BorderStyle.Fixed3D
                        SqlResultGrid.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText
                        SqlResultGrid.ColumnHeadersDefaultCellStyle.Font = New Drawing.Font("Calibri", 9)
                        SqlResultGrid.ColumnHeadersDefaultCellStyle.WrapMode = False
                        SqlResultGrid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
                        SqlResultGrid.DefaultCellStyle.BackColor = Drawing.Color.Beige
                        SqlResultGrid.DefaultCellStyle.Font = New Drawing.Font("Calibri", 8)
                        SqlResultGrid.DefaultCellStyle.WrapMode = False
                        SqlResultGrid.EnableHeadersVisualStyles = False
                        SqlResultGrid.ReadOnly = True
                        SqlResultGrid.RowHeadersVisible = False
                        SqlResultGrid.ScrollBars = ScrollBars.Both
                        SqlResultGrid.ShowCellErrors = False
                        SqlResultGrid.ShowCellToolTips = False
                        SqlResultGrid.ShowEditingIcon = False
                        SqlResultGrid.ShowRowErrors = False

                        ' Set data grid position in the table layout
                        pnlTableLayout.SetCellPosition(SqlResultGrid, New TableLayoutPanelCellPosition(0, CellCounter))

                        ' Add the data grid to the table layout
                        Delegate_Sub_Add_Control(pnlTableLayout, SqlResultGrid)

                        ' Switch active tab to results tab
                        Delegate_Sub_Set_Active_Tab(TabControlSql, TabPageGrid)

                        ' Increment cell counter
                        CellCounter += 1

                        ' Iterate through table columns
                        For i As Integer = 0 To SqlData.FieldCount - 1

                            ' Check for a column name
                            If SqlData.GetName(i).Equals("") Then

                                ' Stub an empty column name
                                Delegate_Sub_Add_DataGridView_Column(SqlResultGrid, "col_" + i.ToString, "(No column name)")

                            Else

                                ' Set column name
                                Delegate_Sub_Add_DataGridView_Column(SqlResultGrid, "col_" + i.ToString, SqlData.GetName(i))

                            End If

                        Next

                        ' Iterate each record
                        While SqlData.Read

                            ' Iterate each column of the record
                            For i As Integer = 0 To SqlData.FieldCount - 1

                                ' Check column position for null
                                If SqlData.IsDBNull(i) Then

                                    ' Add string null instead of actual null
                                    SqlResultRow.Add("Null")

                                ElseIf SqlData.GetFieldType(i) Is GetType(Byte()) Then

                                    ' Add string representation of binary data
                                    SqlResultRow.Add("0x" + BitConverter.ToString(SqlData.GetSqlBinary(i).Value).Replace("-", ""))

                                Else

                                    ' Add raw data
                                    SqlResultRow.Add(SqlData(i).ToString)

                                End If

                            Next

                            ' Add the row to the data grid
                            Delegate_Sub_Add_DataGridView_Row(SqlResultGrid, SqlResultRow.ToArray)

                            ' Clear the result list
                            SqlResultRow.Clear()

                            ' Increment the record counter
                            BatchRecordReturnCount += 1

                            ' Check for termination or cancellation signal
                            If TerminateSignal Or CancelSignal Or CancelQuery Then

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

                    End If

                    ' Check for positive number of affected records
                    If SqlData.RecordsAffected <> -1 Then

                        ' Messages -- inform user of affected record count
                        Delegate_Sub_Append_Text(txtSqlMessage, Environment.NewLine +
                                                         String.Format("({0} record(s) affected)" +
                                                         Environment.NewLine, SqlData.RecordsAffected.ToString))

                        ' Write debug
                        Delegate_Sub_Append_Text(rtbDebug, CallStack + String.Format("({0} record(s) affected)", SqlData.RecordsAffected.ToString))

                    Else

                        ' Messages -- inform user no records were affected
                        Delegate_Sub_Append_Text(txtSqlMessage, Environment.NewLine +
                                                         String.Format("(0 record(s) affected, {0} row(s) returned)", BatchRecordReturnCount.ToString) +
                                                         Environment.NewLine)

                        ' Write debug
                        Delegate_Sub_Append_Text(rtbDebug, CallStack + String.Format("(0 record(s) affected, {0} row(s) returned)", BatchRecordReturnCount.ToString))

                    End If

                    ' Increment total record return count
                    TotalRecordReturnCount += BatchRecordReturnCount

                    ' Reset batch record counter
                    BatchRecordReturnCount = 0

                    ' Check for termination or cancellation signal
                    If TerminateSignal Or CancelSignal Or CancelQuery Then

                        ' Close sql data reader
                        SqlData.Close()

                        ' Close the database connection
                        DbConnection.Close()

                        ' Write debug
                        Delegate_Sub_Append_Text(rtbDebug, CallStack + "Database connection closed.")

                        ' Return
                        Return

                    End If

                Loop While SqlData.NextResult

                ' Close sql data reader
                SqlData.Close()

            Next

            ' Check total record returned count
            If TotalRecordReturnCount > 0 Then

                ' Results tab -- Update tab text
                Delegate_Sub_Set_Text(TabPageGrid, "Results (" + TotalRecordReturnCount.ToString + " records)")

            Else

                ' Results tab -- Reset tab text
                Delegate_Sub_Set_Text(TabPageGrid, "Results")

            End If

        Catch ex As Exception

            ' Write debug
            Delegate_Sub_Set_Text(txtSqlMessage, ex.Message)
            Delegate_Sub_Append_Text(rtbDebug, CallStack + "Exception:" + Environment.NewLine + ex.Message)
            Delegate_Sub_Append_Text(rtbDebug, CallStack + "Stack trace: " + Environment.NewLine + ex.StackTrace)

        Finally

            ' Check for termination signal
            If Not TerminateSignal Then

                ' Set sql query editor function availability
                Delegate_Sub_Enable_Control(rtbSqlQuery, True)
                Delegate_Sub_Enable_Blue_Button(btnSqlSubmitQueryEditor, True)
                Delegate_Sub_Enable_Blue_Button(btnSqlCancelQueryEditor, False)
                Delegate_Select_Button(btnSqlSubmitQueryEditor)

                ' Reset cancellation flag
                CancelQuery = False

            End If

        End Try

    End Sub

    Private Sub txtSqlQuery_TextChanged(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles rtbSqlQuery.KeyDown, btnSqlSubmitQueryEditor.KeyDown

        ' Check for F5 key
        If e.KeyCode = Keys.F5 Then

            ' Press submit button
            btnSqlSubmitQueryEditor.PerformClick()

            ' Cancel the key
            e.SuppressKeyPress = True

        End If

    End Sub

    Private Sub OnSqlQueryEditorInfoMessage(ByVal sender As Object, ByVal e As SqlInfoMessageEventArgs)

        ' Write debug
        Delegate_Sub_Append_Text(txtSqlMessage, e.Message)

    End Sub

End Class