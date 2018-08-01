﻿'****************************** Class Header *******************************\
' Project Name: WinOffline
' Class Name:   WinOfflineUI
' File Name:    WinOfflineUI-Sql-UnUsedSoftGrid.vb
' Author:       Brian Fontana
'***************************************************************************/

Imports System.Data
Imports System.Data.SqlClient
Imports System.Threading
Imports System.Windows.Forms

Partial Public Class WinOfflineUI

    Private Shared UnUsedSoftGridThread As Thread

    Private Sub InitSqlUnUsedSoftGrid()

        ' Disable buttons
        Delegate_Sub_Enable_Red_Button(btnSqlDisconnectUnUsedSoftGrid, False)
        Delegate_Sub_Enable_Blue_Button(btnSqlRefreshUnUsedSoftGrid, False)
        Delegate_Sub_Enable_Blue_Button(btnSqlExportUnUsedSoftGrid, False)

        ' Set grid properties
        dgvSwNotUsed.AllowUserToAddRows = False
        dgvSwNotUsed.AllowUserToDeleteRows = False
        dgvSwNotUsed.AllowUserToResizeRows = False
        dgvSwNotUsed.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgvSwNotUsed.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
        dgvSwNotUsed.BorderStyle = BorderStyle.Fixed3D
        dgvSwNotUsed.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText
        dgvSwNotUsed.ColumnHeadersDefaultCellStyle.Font = New Drawing.Font("Calibri", 9)
        dgvSwNotUsed.ColumnHeadersDefaultCellStyle.WrapMode = False
        dgvSwNotUsed.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        dgvSwNotUsed.DefaultCellStyle.BackColor = Drawing.Color.Beige
        dgvSwNotUsed.DefaultCellStyle.Font = New Drawing.Font("Calibri", 9)
        dgvSwNotUsed.DefaultCellStyle.WrapMode = False
        dgvSwNotUsed.EnableHeadersVisualStyles = False
        dgvSwNotUsed.ReadOnly = True
        dgvSwNotUsed.RowHeadersVisible = False
        dgvSwNotUsed.ScrollBars = ScrollBars.Both
        dgvSwNotUsed.ShowCellErrors = False
        dgvSwNotUsed.ShowCellToolTips = False
        dgvSwNotUsed.ShowEditingIcon = False
        dgvSwNotUsed.ShowRowErrors = False

        ' Set grid properties
        dgvSwNotInst.AllowUserToAddRows = False
        dgvSwNotInst.AllowUserToDeleteRows = False
        dgvSwNotInst.AllowUserToResizeRows = False
        dgvSwNotInst.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgvSwNotInst.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
        dgvSwNotInst.BorderStyle = BorderStyle.Fixed3D
        dgvSwNotInst.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText
        dgvSwNotInst.ColumnHeadersDefaultCellStyle.Font = New Drawing.Font("Calibri", 9)
        dgvSwNotInst.ColumnHeadersDefaultCellStyle.WrapMode = False
        dgvSwNotInst.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        dgvSwNotInst.DefaultCellStyle.BackColor = Drawing.Color.Beige
        dgvSwNotInst.DefaultCellStyle.Font = New Drawing.Font("Calibri", 9)
        dgvSwNotInst.DefaultCellStyle.WrapMode = False
        dgvSwNotInst.EnableHeadersVisualStyles = False
        dgvSwNotInst.ReadOnly = True
        dgvSwNotInst.RowHeadersVisible = False
        dgvSwNotInst.ScrollBars = ScrollBars.Both
        dgvSwNotInst.ShowCellErrors = False
        dgvSwNotInst.ShowCellToolTips = False
        dgvSwNotInst.ShowEditingIcon = False
        dgvSwNotInst.ShowRowErrors = False

        ' Set grid properties
        dgvSwNotStaged.AllowUserToAddRows = False
        dgvSwNotStaged.AllowUserToDeleteRows = False
        dgvSwNotStaged.AllowUserToResizeRows = False
        dgvSwNotStaged.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgvSwNotStaged.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
        dgvSwNotStaged.BorderStyle = BorderStyle.Fixed3D
        dgvSwNotStaged.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText
        dgvSwNotStaged.ColumnHeadersDefaultCellStyle.Font = New Drawing.Font("Calibri", 9)
        dgvSwNotStaged.ColumnHeadersDefaultCellStyle.WrapMode = False
        dgvSwNotStaged.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        dgvSwNotStaged.DefaultCellStyle.BackColor = Drawing.Color.Beige
        dgvSwNotStaged.DefaultCellStyle.Font = New Drawing.Font("Calibri", 9)
        dgvSwNotStaged.DefaultCellStyle.WrapMode = False
        dgvSwNotStaged.EnableHeadersVisualStyles = False
        dgvSwNotStaged.ReadOnly = True
        dgvSwNotStaged.RowHeadersVisible = False
        dgvSwNotStaged.ScrollBars = ScrollBars.Both
        dgvSwNotStaged.ShowCellErrors = False
        dgvSwNotStaged.ShowCellToolTips = False
        dgvSwNotStaged.ShowEditingIcon = False
        dgvSwNotStaged.ShowRowErrors = False

        ' Force all tabs to draw (otherwise data grid views won't format during population)
        Delegate_Sub_Set_Active_Tab(tabCtrlSwNotUsed, tabSwNotStaged)
        Delegate_Sub_Set_Active_Tab(tabCtrlSwNotUsed, tabSwNotInst)
        Delegate_Sub_Set_Active_Tab(tabCtrlSwNotUsed, tabSwNotUsed)

    End Sub

    Private Sub UnUsedSoftGridWorker(ByVal ConnectionString As String)

        ' Local variables
        Dim DbConnection As SqlConnection = New SqlConnection(ConnectionString)
        Dim RecordCount As Integer = 0
        Dim CallStack As String = "UnUsedSoftGridWorker --> "

        ' Encapsulate grid worker
        Try

            ' Open sql connection
            DbConnection.Open()

            ' Write debug
            Delegate_Sub_Append_Text(rtbDebug, CallStack + "Connected successful: " + SqlServer)

            ' Reveal progress bar
            tabCtrlSwNotUsed.Invoke(Sub() tabCtrlSwNotUsed.Height = tabCtrlSwNotUsed.Height - prgUnUsedSoftGrid.Height - 4)

            ' Query software not used
            SqlPopulateGridWorker(CallStack,
                                          DbConnection,
                                          "select itemname as 'Software Title', itemversion as 'Software Version', DATEADD(s, DATEDIFF(s, getutcdate(), getdate()), DATEADD(s, creationdate, '19700101')) as 'Registration Date' from usd_rsw with (nolock) where objectid not in (select rsw from usd_actproc with (nolock) where objectid in (select actproc from usd_applic with (nolock))) and itemtype<>5 order by itemname, itemversion",
                                          dgvSwNotUsed,
                                          "Software Title",
                                          DataGridViewAutoSizeColumnsMode.AllCells,
                                          DataGridViewAutoSizeRowsMode.AllCells,
                                          prgUnUsedSoftGrid,
                                          tabSwNotUsed)

            ' Query software not installed
            SqlPopulateGridWorker(CallStack,
                                          DbConnection,
                                          "select itemname as 'Software Title', itemversion as 'Software Version', DATEADD(s, DATEDIFF(s, getutcdate(), getdate()), DATEADD(s, creationdate, '19700101')) as 'Registration Date' from usd_rsw with (nolock) where objectid not in (select rsw from usd_actproc with (nolock) where objectid in (select actproc from usd_applic with (nolock) where status in (1,7,8,9,27))) and itemtype<>5 order by itemname, itemversion",
                                          dgvSwNotInst,
                                          "Software Title",
                                          DataGridViewAutoSizeColumnsMode.AllCells,
                                          DataGridViewAutoSizeRowsMode.AllCells,
                                          prgUnUsedSoftGrid,
                                          tabSwNotInst)

            ' Query software not staged
            SqlPopulateGridWorker(CallStack,
                                          DbConnection,
                                          "select itemname as 'Software Title', itemversion as 'Software Version', DATEADD(s, DATEDIFF(s, getutcdate(), getdate()), DATEADD(s, creationdate, '19700101')) as 'Registration Date' from usd_rsw with (nolock) where objectid not in (select rsw from usd_actproc with (nolock) where objectid in (select actproc from usd_applic with (nolock) where status in (1,2,3,4,17,18,19,20,21,27))) and itemtype<>5 order by itemname, itemversion",
                                          dgvSwNotStaged,
                                          "Software Title",
                                          DataGridViewAutoSizeColumnsMode.AllCells,
                                          DataGridViewAutoSizeRowsMode.AllCells,
                                          prgUnUsedSoftGrid,
                                          tabSwNotStaged)

            ' Hide progress bar
            tabCtrlSwNotUsed.Invoke(Sub() tabCtrlSwNotUsed.Height = pnlSqlUnUsedSoftGrid.Height - pnlSqlUnUsedSoftGridButtons.Height - 3)

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
            Delegate_Sub_Enable_Yellow_Button(btnSqlRefreshUnUsedSoftGrid, True)
            Delegate_Sub_Enable_Tan_Button(btnSqlExportUnUsedSoftGrid, True)

        End Try

    End Sub

    Private Sub btnSqlConnectUnUsedSoftGrid_Click(sender As Object, e As EventArgs) Handles btnSqlConnectUnUsedSoftGrid.Click

        ' Perform SQL connection
        SqlConnect()

    End Sub

    Private Sub btnSqlDisconnectUnUsedSoftGrid_Click(sender As Object, e As EventArgs) Handles btnSqlDisconnectUnUsedSoftGrid.Click

        ' Perform disconnect method
        SqlDisconnect()

    End Sub

    Private Sub btnSqlRefreshUnUsedSoftGrid_Click(sender As Object, e As EventArgs) Handles btnSqlRefreshUnUsedSoftGrid.Click

        ' Disable buttons
        Delegate_Sub_Enable_Yellow_Button(btnSqlRefreshUnUsedSoftGrid, False)
        Delegate_Sub_Enable_Tan_Button(btnSqlExportUnUsedSoftGrid, False)

        ' Reset tab text
        If tabSwNotUsed.Text.Contains("(") Then Delegate_Sub_Set_Text(tabSwNotUsed, tabSwNotUsed.Text.Substring(0, tabSwNotUsed.Text.IndexOf("(") - 1))
        If tabSwNotInst.Text.Contains("(") Then Delegate_Sub_Set_Text(tabSwNotInst, tabSwNotInst.Text.Substring(0, tabSwNotInst.Text.IndexOf("(") - 1))
        If tabSwNotStaged.Text.Contains("(") Then Delegate_Sub_Set_Text(tabSwNotStaged, tabSwNotStaged.Text.Substring(0, tabSwNotStaged.Text.IndexOf("(") - 1))

        ' Restart thread
        UnUsedSoftGridThread = New Thread(Sub() UnUsedSoftGridWorker(ConnectionString))
        UnUsedSoftGridThread.Start()

    End Sub

    Private Sub btnSqlExportUnUsedSoftGrid_Click(sender As Object, e As EventArgs) Handles btnSqlExportUnUsedSoftGrid.Click

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

            ' Check selected tab
            If tabCtrlSwNotUsed.SelectedTab.Equals(tabSwNotUsed) Then

                ' Iterate datagrid column headers
                For Each dgvColumn As DataGridViewColumn In dgvSwNotUsed.Columns

                    ' Write values
                    StateStreamWriter.Write(dgvColumn.HeaderText + ",")

                Next

                ' Write newline
                StateStreamWriter.Write(Environment.NewLine)

                ' Iterate datagrid rows
                For Each dgvRecord As DataGridViewRow In dgvSwNotUsed.Rows

                    ' Iterate cells
                    For Each CellItem As DataGridViewCell In dgvRecord.Cells

                        ' Write values
                        StateStreamWriter.Write(CellItem.Value.ToString.Replace(",", "+") + ",")

                    Next

                    ' Write newline
                    StateStreamWriter.Write(Environment.NewLine)

                Next

            ElseIf tabCtrlSwNotUsed.SelectedTab.Equals(tabSwNotInst) Then

                ' Iterate datagrid column headers
                For Each dgvColumn As DataGridViewColumn In dgvSwNotInst.Columns

                    ' Write values
                    StateStreamWriter.Write(dgvColumn.HeaderText + ",")

                Next

                ' Write newline
                StateStreamWriter.Write(Environment.NewLine)

                ' Iterate datagrid rows
                For Each dgvRecord As DataGridViewRow In dgvSwNotInst.Rows

                    ' Iterate cells
                    For Each CellItem As DataGridViewCell In dgvRecord.Cells

                        ' Write values
                        StateStreamWriter.Write(CellItem.Value.ToString.Replace(",", "+") + ",")

                    Next

                    ' Write newline
                    StateStreamWriter.Write(Environment.NewLine)

                Next

            ElseIf tabCtrlSwNotUsed.SelectedTab.Equals(tabSwNotStaged) Then

                ' Iterate datagrid column headers
                For Each dgvColumn As DataGridViewColumn In dgvSwNotStaged.Columns

                    ' Write values
                    StateStreamWriter.Write(dgvColumn.HeaderText + ",")

                Next

                ' Write newline
                StateStreamWriter.Write(Environment.NewLine)

                ' Iterate datagrid rows
                For Each dgvRecord As DataGridViewRow In dgvSwNotStaged.Rows

                    ' Iterate cells
                    For Each CellItem As DataGridViewCell In dgvRecord.Cells

                        ' Write values
                        StateStreamWriter.Write(CellItem.Value.ToString.Replace(",", "+") + ",")

                    Next

                    ' Write newline
                    StateStreamWriter.Write(Environment.NewLine)

                Next

            End If

            ' Close output stream
            StateStreamWriter.Close()

        Catch ex As Exception

            ' Push user alert
            AlertBox.CreateUserAlert("Export failed." + Environment.NewLine + Environment.NewLine + "Exception: " + ex.Message)

        End Try

    End Sub

    Private Sub btnSqlExitUnUsedSoftGrid_Click(sender As Object, e As EventArgs) Handles btnSqlExitUnUsedSoftGrid.Click

        ' Close the dialog
        Close()

    End Sub

End Class