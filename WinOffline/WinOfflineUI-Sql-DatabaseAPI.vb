'****************************** Class Header *******************************\
' Project Name: WinOffline
' Class Name:   WinOfflineUI
' File Name:    WinOfflineUI-Sql-DatabaseAPI.vb
' Author:       Brian Fontana
'***************************************************************************/

Imports System.Data
Imports System.Data.SqlClient
Imports System.Threading
Imports System.Windows.Forms

Partial Public Class WinOfflineUI

    Public Const WINDOWS_AUTHENTICATION As String = "Windows Authentication"
    Public Const SQL_AUTHENTICATION As String = "SQL Authentication"
    Public Shared SqlServer As String = ""
    Public Shared InstanceName As String = ""
    Public Shared PortNumber As String = "1433"
    Public Shared DatabaseName As String = "mdb"
    Public Shared AuthType As String = WINDOWS_AUTHENTICATION
    Public Shared SqlUser As String = "ca_itrm"
    Public Shared SqlPassword As String = ""
    Public Shared ConnectionString As String = Nothing
    Public Shared DatabaseThread As Thread = Nothing
    Public Shared CancelSignal As Boolean = False
    Public Shared DomainUuid As String = Nothing
    Public Shared DomainId As String = Nothing
    Public Shared DomainType As String = Nothing
    Public Shared ManagerName As String = Nothing

    Public Sub SqlConnect()

        ' Local variables
        Dim SqlConnectForm As SqlConnectUI = New SqlConnectUI

        ' Show connection form
        If SqlConnectForm.ShowDialog() = DialogResult.Abort Then

            ' Return
            Return

        End If

        ' Build connection string
        ConnectionString = BuildSqlConnectionString()

        ' Check connection string
        If ConnectionString Is Nothing Then

            ' Return
            Return

        End If

        ' Disable buttons
        Delegate_Sub_Enable_Green_Button(btnSqlConnectMdbOverview, False)
        Delegate_Sub_Enable_Green_Button(btnSqlConnectTableSpace, False)
        Delegate_Sub_Enable_Green_Button(btnSqlConnectAgentGrid, False)
        Delegate_Sub_Enable_Green_Button(btnSqlConnectUserGrid, False)
        Delegate_Sub_Enable_Green_Button(btnSqlConnectServerGrid, False)
        Delegate_Sub_Enable_Green_Button(btnSqlConnectEngineGrid, False)
        Delegate_Sub_Enable_Green_Button(btnSqlConnectGroupEvalGrid, False)
        Delegate_Sub_Enable_Green_Button(btnSqlConnectInstSoftGrid, False)
        Delegate_Sub_Enable_Green_Button(btnSqlConnectDiscSoftGrid, False)
        Delegate_Sub_Enable_Green_Button(btnSqlConnectDuplCompGrid, False)
        Delegate_Sub_Enable_Green_Button(btnSqlConnectUnUsedSoftGrid, False)
        Delegate_Sub_Enable_Green_Button(btnSqlConnectCleanApps, False)
        Delegate_Sub_Enable_Green_Button(btnSqlConnectQueryEditor, False)

        ' Enable buttons
        Delegate_Sub_Enable_Red_Button(btnSqlDisconnectMdbOverview, True)
        Delegate_Sub_Enable_Red_Button(btnSqlDisconnectTableSpace, True)
        Delegate_Sub_Enable_Red_Button(btnSqlDisconnectAgentGrid, True)
        Delegate_Sub_Enable_Red_Button(btnSqlDisconnectUserGrid, True)
        Delegate_Sub_Enable_Red_Button(btnSqlDisconnectServerGrid, True)
        Delegate_Sub_Enable_Red_Button(btnSqlDisconnectEngineGrid, True)
        Delegate_Sub_Enable_Red_Button(btnSqlDisconnectGroupEvalGrid, True)
        Delegate_Sub_Enable_Red_Button(btnSqlDisconnectInstSoftGrid, True)
        Delegate_Sub_Enable_Red_Button(btnSqlDisconnectDiscSoftGrid, True)
        Delegate_Sub_Enable_Red_Button(btnSqlDisconnectDuplCompGrid, True)
        Delegate_Sub_Enable_Red_Button(btnSqlDisconnectUnUsedSoftGrid, True)
        Delegate_Sub_Enable_Red_Button(btnSqlDisconnectCleanApps, True)
        Delegate_Sub_Enable_Tan_Button(btnSqlRunCleanApps, True)
        Delegate_Sub_Enable_Red_Button(btnSqlDisconnectQueryEditor, True)
        Delegate_Sub_Enable_Yellow_Button(btnSqlSubmitQueryEditor, True)

        ' Clear cancellation flag
        CancelSignal = False

        ' Start database connection thread
        DatabaseThread = New Thread(Sub() SqlDatabaseWorker())
        DatabaseThread.Start()

    End Sub

    Public Sub SqlDisconnect()

        ' Raise cancellation signals
        CancelSignal = True

        ' Loop while external threads are still running
        While True

            ' Check if external threads are still alive
            If (MdbOverviewThread IsNot Nothing AndAlso MdbOverviewThread.IsAlive) OrElse
               (TableSpaceGridThread IsNot Nothing AndAlso TableSpaceGridThread.IsAlive) OrElse
               (AgentGridThread IsNot Nothing AndAlso AgentGridThread.IsAlive) OrElse
               (UserGridThread IsNot Nothing AndAlso UserGridThread.IsAlive) OrElse
               (ServerGridThread IsNot Nothing AndAlso ServerGridThread.IsAlive) OrElse
               (EngineGridThread IsNot Nothing AndAlso EngineGridThread.IsAlive) OrElse
               (GroupEvalGridThread IsNot Nothing AndAlso GroupEvalGridThread.IsAlive) OrElse
               (InstSoftGridThread IsNot Nothing AndAlso InstSoftGridThread.IsAlive) OrElse
               (DiscSoftGridThread IsNot Nothing AndAlso DiscSoftGridThread.IsAlive) OrElse
               (UnUsedSoftGridThread IsNot Nothing AndAlso UnUsedSoftGridThread.IsAlive) OrElse
               (DuplCompGridThread IsNot Nothing AndAlso DuplCompGridThread.IsAlive) OrElse
               (SqlCleanAppsThread IsNot Nothing AndAlso SqlCleanAppsThread.IsAlive) OrElse
               (SqlQueryEditorThread IsNot Nothing AndAlso SqlQueryEditorThread.IsAlive) Then

                ' Process message queue
                Application.DoEvents()

            Else

                ' All threads finished -- exit loop
                Exit While

            End If

        End While

        ' Reset threads
        MdbOverviewThread = Nothing
        TableSpaceGridThread = Nothing
        AgentGridThread = Nothing
        UserGridThread = Nothing
        ServerGridThread = Nothing
        EngineGridThread = Nothing
        GroupEvalGridThread = Nothing
        InstSoftGridThread = Nothing
        DiscSoftGridThread = Nothing
        UnUsedSoftGridThread = Nothing
        DuplCompGridThread = Nothing
        SqlCleanAppsThread = Nothing
        SqlQueryEditorThread = Nothing

        ' Clear views
        Delegate_Sub_Set_Text(txtMdbVersion, "")
        Delegate_Sub_Set_Text(txtITCMVersion, "")
        Delegate_Sub_Set_Text(txtMdbInstallDate, "")
        Delegate_Sub_Set_Text(txtMdbType, "")
        Delegate_Sub_Clear_ListView(lvwITCMSummary)
        Delegate_Sub_Clear_DataGridView(dgvAgentVersion)
        Delegate_Sub_Clear_DataGridView(dgvContentSummary)
        Delegate_Sub_Clear_DataGridView(dgvTableSpaceGrid)
        Delegate_Sub_Clear_DataGridView(dgvAgentGrid)
        Delegate_Sub_Clear_DataGridView(dgvAgentObsolete90)
        Delegate_Sub_Clear_DataGridView(dgvAgentObsolete365)
        Delegate_Sub_Clear_DataGridView(dgvUserGrid)
        Delegate_Sub_Clear_DataGridView(dgvUserObsolete90)
        Delegate_Sub_Clear_DataGridView(dgvUserObsolete365)
        Delegate_Sub_Clear_DataGridView(dgvServerGrid)
        Delegate_Sub_Clear_DataGridView(dgvServerLastCollected24)
        Delegate_Sub_Clear_DataGridView(dgvServerSignature30)
        Delegate_Sub_Clear_DataGridView(dgvEngineGrid)
        Delegate_Sub_Clear_DataGridView(dgvGroupEvalGrid)
        Delegate_Sub_Clear_DataGridView(dgvSoftInst)
        Delegate_Sub_Clear_DataGridView(dgvDiscSignature)
        Delegate_Sub_Clear_DataGridView(dgvDiscCustom)
        Delegate_Sub_Clear_DataGridView(dgvDiscHeuristic)
        Delegate_Sub_Clear_DataGridView(dgvDiscIntellisig)
        Delegate_Sub_Clear_DataGridView(dgvDiscEverything)
        Delegate_Sub_Clear_DataGridView(dgvDuplHostname)
        Delegate_Sub_Clear_DataGridView(dgvDuplSerialNum)
        Delegate_Sub_Clear_DataGridView(dgvDuplBoth)
        Delegate_Sub_Clear_DataGridView(dgvDuplBlank)
        Delegate_Sub_Clear_DataGridView(dgvSwNotUsed)
        Delegate_Sub_Clear_DataGridView(dgvSwNotInst)
        Delegate_Sub_Clear_DataGridView(dgvSwNotStaged)

        ' Reset text
        If tabDiscSignature.Text.Contains("(") Then Delegate_Sub_Set_Text(tabDiscSignature, tabDiscSignature.Text.Substring(0, tabDiscSignature.Text.IndexOf("(") - 1))
        If tabDiscCustom.Text.Contains("(") Then Delegate_Sub_Set_Text(tabDiscCustom, tabDiscCustom.Text.Substring(0, tabDiscCustom.Text.IndexOf("(") - 1))
        If tabDiscHeuristic.Text.Contains("(") Then Delegate_Sub_Set_Text(tabDiscHeuristic, tabDiscHeuristic.Text.Substring(0, tabDiscHeuristic.Text.IndexOf("(") - 1))
        If tabDiscIntellisig.Text.Contains("(") Then Delegate_Sub_Set_Text(tabDiscIntellisig, tabDiscIntellisig.Text.Substring(0, tabDiscIntellisig.Text.IndexOf("(") - 1))
        If tabDiscEverything.Text.Contains("(") Then Delegate_Sub_Set_Text(tabDiscEverything, tabDiscEverything.Text.Substring(0, tabDiscEverything.Text.IndexOf("(") - 1))
        If tabDuplHostname.Text.Contains("(") Then Delegate_Sub_Set_Text(tabDuplHostname, tabDuplHostname.Text.Substring(0, tabDuplHostname.Text.IndexOf("(") - 1))
        If tabDuplSerialNum.Text.Contains("(") Then Delegate_Sub_Set_Text(tabDuplSerialNum, tabDuplSerialNum.Text.Substring(0, tabDuplSerialNum.Text.IndexOf("(") - 1))
        If tabDuplBoth.Text.Contains("(") Then Delegate_Sub_Set_Text(tabDuplBoth, tabDuplBoth.Text.Substring(0, tabDuplBoth.Text.IndexOf("(") - 1))
        If tabDuplBlank.Text.Contains("(") Then Delegate_Sub_Set_Text(tabDuplBlank, tabDuplBlank.Text.Substring(0, tabDuplBlank.Text.IndexOf("(") - 1))
        If grpSqlInstSoftGrid.Text.Contains("(") Then Delegate_Sub_Set_Text(grpSqlInstSoftGrid, grpSqlInstSoftGrid.Text.Substring(0, grpSqlInstSoftGrid.Text.IndexOf("(") - 1))
        If tabSwNotUsed.Text.Contains("(") Then Delegate_Sub_Set_Text(tabSwNotUsed, tabSwNotUsed.Text.Substring(0, tabSwNotUsed.Text.IndexOf("(") - 1))
        If tabSwNotInst.Text.Contains("(") Then Delegate_Sub_Set_Text(tabSwNotInst, tabSwNotInst.Text.Substring(0, tabSwNotInst.Text.IndexOf("(") - 1))
        If tabSwNotStaged.Text.Contains("(") Then Delegate_Sub_Set_Text(tabSwNotStaged, tabSwNotStaged.Text.Substring(0, tabSwNotStaged.Text.IndexOf("(") - 1))
        If tabAgentSummary.Text.Contains("(") Then Delegate_Sub_Set_Text(tabAgentSummary, tabAgentSummary.Text.Substring(0, tabAgentSummary.Text.IndexOf("(") - 1))
        If tabAgentObsolete90.Text.Contains("(") Then Delegate_Sub_Set_Text(tabAgentObsolete90, tabAgentObsolete90.Text.Substring(0, tabAgentObsolete90.Text.IndexOf("(") - 1))
        If tabAgentObsolete365.Text.Contains("(") Then Delegate_Sub_Set_Text(tabAgentObsolete365, tabAgentObsolete365.Text.Substring(0, tabAgentObsolete365.Text.IndexOf("(") - 1))
        If tabUserSummary.Text.Contains("(") Then Delegate_Sub_Set_Text(tabUserSummary, tabUserSummary.Text.Substring(0, tabUserSummary.Text.IndexOf("(") - 1))
        If tabUserObsolete90.Text.Contains("(") Then Delegate_Sub_Set_Text(tabUserObsolete90, tabUserObsolete90.Text.Substring(0, tabUserObsolete90.Text.IndexOf("(") - 1))
        If tabUserObsolete365.Text.Contains("(") Then Delegate_Sub_Set_Text(tabUserObsolete365, tabUserObsolete365.Text.Substring(0, tabUserObsolete365.Text.IndexOf("(") - 1))
        If tabServerSummary.Text.Contains("(") Then Delegate_Sub_Set_Text(tabServerSummary, tabServerSummary.Text.Substring(0, tabServerSummary.Text.IndexOf("(") - 1))
        If tabServerLastCollected24.Text.Contains("(") Then Delegate_Sub_Set_Text(tabServerLastCollected24, tabServerLastCollected24.Text.Substring(0, tabServerLastCollected24.Text.IndexOf("(") - 1))
        If tabServerSignature30.Text.Contains("(") Then Delegate_Sub_Set_Text(tabServerSignature30, tabServerSignature30.Text.Substring(0, tabServerSignature30.Text.IndexOf("(") - 1))
        If grpSqlTableSpace.Text.Contains("(") Then Delegate_Sub_Set_Text(grpSqlTableSpace, grpSqlTableSpace.Text.Substring(0, grpSqlTableSpace.Text.IndexOf("(") - 1))
        If grpSqlGroupEvalGrid.Text.Contains("(") Then Delegate_Sub_Set_Text(grpSqlGroupEvalGrid, grpSqlGroupEvalGrid.Text.Substring(0, grpSqlGroupEvalGrid.Text.IndexOf("(") - 1))

        ' Enable buttons
        Delegate_Sub_Enable_Green_Button(btnSqlConnectMdbOverview, True)
        Delegate_Sub_Enable_Green_Button(btnSqlConnectTableSpace, True)
        Delegate_Sub_Enable_Green_Button(btnSqlConnectAgentGrid, True)
        Delegate_Sub_Enable_Green_Button(btnSqlConnectUserGrid, True)
        Delegate_Sub_Enable_Green_Button(btnSqlConnectServerGrid, True)
        Delegate_Sub_Enable_Green_Button(btnSqlConnectEngineGrid, True)
        Delegate_Sub_Enable_Green_Button(btnSqlConnectGroupEvalGrid, True)
        Delegate_Sub_Enable_Green_Button(btnSqlConnectInstSoftGrid, True)
        Delegate_Sub_Enable_Green_Button(btnSqlConnectDiscSoftGrid, True)
        Delegate_Sub_Enable_Green_Button(btnSqlConnectDuplCompGrid, True)
        Delegate_Sub_Enable_Green_Button(btnSqlConnectUnUsedSoftGrid, True)
        Delegate_Sub_Enable_Green_Button(btnSqlConnectCleanApps, True)
        Delegate_Sub_Enable_Green_Button(btnSqlConnectQueryEditor, True)

        ' Disable buttons
        Delegate_Sub_Enable_Red_Button(btnSqlDisconnectMdbOverview, False)
        Delegate_Sub_Enable_Yellow_Button(btnSqlRefreshMdbOverview, False)
        Delegate_Sub_Enable_Tan_Button(btnSqlExportMdbOverview, False)
        Delegate_Sub_Enable_Red_Button(btnSqlDisconnectTableSpace, False)
        Delegate_Sub_Enable_Yellow_Button(btnSqlRefreshTableSpace, False)
        Delegate_Sub_Enable_Tan_Button(btnSqlExportTableSpace, False)
        Delegate_Sub_Enable_Red_Button(btnSqlDisconnectAgentGrid, False)
        Delegate_Sub_Enable_Yellow_Button(btnSqlRefreshAgentGrid, False)
        Delegate_Sub_Enable_Tan_Button(btnSqlExportAgentGrid, False)
        Delegate_Sub_Enable_Red_Button(btnSqlDisconnectUserGrid, False)
        Delegate_Sub_Enable_Yellow_Button(btnSqlRefreshUserGrid, False)
        Delegate_Sub_Enable_Tan_Button(btnSqlExportUserGrid, False)
        Delegate_Sub_Enable_Red_Button(btnSqlDisconnectServerGrid, False)
        Delegate_Sub_Enable_Yellow_Button(btnSqlRefreshServerGrid, False)
        Delegate_Sub_Enable_Tan_Button(btnSqlExportServerGrid, False)
        Delegate_Sub_Enable_Red_Button(btnSqlDisconnectEngineGrid, False)
        Delegate_Sub_Enable_Yellow_Button(btnSqlRefreshEngineGrid, False)
        Delegate_Sub_Enable_Tan_Button(btnSqlExportEngineGrid, False)
        Delegate_Sub_Enable_Red_Button(btnSqlDisconnectGroupEvalGrid, False)
        Delegate_Sub_Enable_Yellow_Button(btnSqlRefreshGroupEvalGrid, False)
        Delegate_Sub_Enable_Tan_Button(btnSqlExportGroupEvalGrid, False)
        Delegate_Sub_Enable_CadetBlue_Button(btnGroupEvalGridCommit, False)
        Delegate_Sub_Enable_LightCoral_Button(btnGroupEvalGridDiscard, False)
        Delegate_Sub_Enable_Red_Button(btnSqlDisconnectInstSoftGrid, False)
        Delegate_Sub_Enable_Yellow_Button(btnSqlRefreshInstSoftGrid, False)
        Delegate_Sub_Enable_Tan_Button(btnSqlExportInstSoftGrid, False)
        Delegate_Sub_Enable_Red_Button(btnSqlDisconnectDiscSoftGrid, False)
        Delegate_Sub_Enable_Yellow_Button(btnSqlRefreshDiscSoftGrid, False)
        Delegate_Sub_Enable_Tan_Button(btnSqlExportDiscSoftGrid, False)
        Delegate_Sub_Enable_Red_Button(btnSqlDisconnectDuplCompGrid, False)
        Delegate_Sub_Enable_Yellow_Button(btnSqlRefreshDuplCompGrid, False)
        Delegate_Sub_Enable_Tan_Button(btnSqlExportDuplCompGrid, False)
        Delegate_Sub_Enable_Red_Button(btnSqlDisconnectUnUsedSoftGrid, False)
        Delegate_Sub_Enable_Yellow_Button(btnSqlRefreshUnUsedSoftGrid, False)
        Delegate_Sub_Enable_Tan_Button(btnSqlExportUnUsedSoftGrid, False)
        Delegate_Sub_Enable_Red_Button(btnSqlDisconnectCleanApps, False)
        Delegate_Sub_Enable_Tan_Button(btnSqlRunCleanApps, False)
        Delegate_Sub_Enable_Red_Button(btnSqlDisconnectQueryEditor, False)
        Delegate_Sub_Enable_Yellow_Button(btnSqlSubmitQueryEditor, False)
        Delegate_Sub_Enable_Tan_Button(btnSqlCancelQueryEditor, False)

    End Sub

    Public Function BuildSqlConnectionString() As String

        ' Local variables
        Dim ConnectionString As String = Nothing

        ' Check for empty username or password
        If AuthType.Equals(SQL_AUTHENTICATION) AndAlso (SqlUser.Equals("") Or SqlPassword.Equals("")) Then

            ' Push user alert
            AlertBox.CreateUserAlert("Invalid sql username or password.")

            ' Return
            Return Nothing

        End If

        ' Check for instance or port
        If InstanceName.Equals("") And PortNumber.Equals("") And
                Not (SqlServer.Contains("\") Or SqlServer.Contains(",")) Then

            ' Push user alert
            AlertBox.CreateUserAlert("An instance name or port number is required.")

            ' Return
            Return Nothing

        End If

        ' Check for instance or port
        If DatabaseName.Equals("") Then

            ' Push user alert
            AlertBox.CreateUserAlert("A database name is required.")

            ' Return
            Return Nothing

        End If

        ' Check authentication type
        If AuthType.Equals(SQL_AUTHENTICATION) Then

            ' Check sql server user input
            If SqlServer.Contains("\") Or SqlServer.Contains(",") Then

                ' Build connection string (sql authentication)
                ConnectionString = "Server=" + SqlServer
                ConnectionString += ";Database=" + DatabaseName
                ConnectionString += ";User Id=" + SqlUser
                ConnectionString += ";Password=" + SqlPassword
                ConnectionString += ";Application Name=" + Globals.ProcessFriendlyName + " " + Globals.AppVersion

            Else

                ' Build connection string (sql authentication)
                ConnectionString = "Server=" + SqlServer
                ConnectionString += "\" + InstanceName
                ConnectionString += "," + PortNumber
                ConnectionString += ";Database=" + DatabaseName
                ConnectionString += ";User Id=" + SqlUser
                ConnectionString += ";Password=" + SqlPassword
                ConnectionString += ";Application Name=" + Globals.ProcessFriendlyName + " " + Globals.AppVersion

            End If

        Else

            'Check Sql server user input
            If SqlServer.Contains("\") Or SqlServer.Contains(",") Then

                ' Build connection string (windows authentication)
                ConnectionString = "Server=" + SqlServer
                ConnectionString += ";Database=" + DatabaseName
                ConnectionString += ";Trusted_Connection=True"
                ConnectionString += ";Application Name=" + Globals.ProcessFriendlyName + " " + Globals.AppVersion

            Else

                ' Build connection string (windows authentication)
                ConnectionString = "Server=" + SqlServer
                ConnectionString += "\" + InstanceName
                ConnectionString += "," + PortNumber
                ConnectionString += ";Database=" + DatabaseName
                ConnectionString += ";Trusted_Connection=True"
                ConnectionString += ";Application Name=" + Globals.ProcessFriendlyName + " " + Globals.AppVersion

            End If

        End If

        ' Return
        Return ConnectionString

    End Function

    Public Sub SqlDatabaseWorker()

        ' Local variables
        Dim DbConnection As SqlConnection = New SqlConnection(ConnectionString)
        Dim CallStack As String = "DatabaseWorker --> "

        ' Encapsulate connection operation
        Try

            ' Open sql connection
            DbConnection.Open()

            ' Write debug
            Delegate_Sub_Append_Text(rtbDebug, CallStack + "Connected successful: " + SqlServer)

            ' Query domain_uuid
            DomainUuid = SqlSelectScalar(CallStack, DbConnection, "select set_val_uuid from ca_settings with (nolock) where set_id=1")
            Delegate_Sub_Append_Text(rtbDebug, "DatabaseWorker --> Domain UUID: " + DomainUuid)

            ' Query domain_id
            DomainId = SqlSelectScalar(CallStack, DbConnection, "select domain_id from ca_n_tier with (nolock) where domain_uuid=" + DomainUuid)
            Delegate_Sub_Append_Text(rtbDebug, "DatabaseWorker --> Domain ID: " + DomainId)

            ' Query domain_type
            DomainType = SqlSelectScalar(CallStack, DbConnection, "select domain_type from ca_n_tier with (nolock) where domain_uuid=" + DomainUuid)
            Delegate_Sub_Append_Text(rtbDebug, "DatabaseWorker --> Domain Type: " + DomainType)

            ' Query manager name (label)
            ManagerName = SqlSelectScalar(CallStack, DbConnection, "select label from ca_manager with (nolock) where domain_uuid=" + DomainUuid)
            Delegate_Sub_Append_Text(rtbDebug, "DatabaseWorker --> Manager Name: " + ManagerName)

            ' Check selected node name
            If Callback_Function_Read_Selected_TreeNode(ExplorerTree).Equals("SqlMdbOverviewNode") Then

                ' Start thread
                MdbOverviewThread = New Thread(Sub() MdbOverviewWorker(ConnectionString))
                MdbOverviewThread.Start()

            ElseIf Callback_Function_Read_Selected_TreeNode(ExplorerTree).Equals("SqlTableSpaceGridNode") Then

                ' Start thread
                TableSpaceGridThread = New Thread(Sub() TableSpaceGridWorker(ConnectionString))
                TableSpaceGridThread.Start()

            ElseIf Callback_Function_Read_Selected_TreeNode(ExplorerTree).Equals("AgentGridNode") Then

                ' Start thread
                AgentGridThread = New Thread(Sub() AgentGridWorker(ConnectionString))
                AgentGridThread.Start()

            ElseIf Callback_Function_Read_Selected_TreeNode(ExplorerTree).Equals("UserGridNode") Then

                ' Start thread
                UserGridThread = New Thread(Sub() UserGridWorker(ConnectionString))
                UserGridThread.Start()

            ElseIf Callback_Function_Read_Selected_TreeNode(ExplorerTree).Equals("ServerGridNode") Then

                ' Start thread
                ServerGridThread = New Thread(Sub() ServerGridWorker(ConnectionString))
                ServerGridThread.Start()

            ElseIf Callback_Function_Read_Selected_TreeNode(ExplorerTree).Equals("EngineGridNode") Then

                ' Start thread
                EngineGridThread = New Thread(Sub() EngineGridWorker(ConnectionString))
                EngineGridThread.Start()

            ElseIf Callback_Function_Read_Selected_TreeNode(ExplorerTree).Equals("GroupEvalGridNode") Then

                ' Start thread
                GroupEvalGridThread = New Thread(Sub() GroupEvalGridWorker(ConnectionString))
                GroupEvalGridThread.Start()

            ElseIf Callback_Function_Read_Selected_TreeNode(ExplorerTree).Equals("InstSoftGridNode") Then

                ' Start thread
                InstSoftGridThread = New Thread(Sub() InstSoftGridWorker(ConnectionString))
                InstSoftGridThread.Start()

            ElseIf Callback_Function_Read_Selected_TreeNode(ExplorerTree).Equals("DiscSoftGridNode") Then

                ' Start thread
                DiscSoftGridThread = New Thread(Sub() DiscSoftGridWorker(ConnectionString))
                DiscSoftGridThread.Start()

            ElseIf Callback_Function_Read_Selected_TreeNode(ExplorerTree).Equals("DuplCompGridNode") Then

                ' Start thread
                DuplCompGridThread = New Thread(Sub() DuplCompGridWorker(ConnectionString))
                DuplCompGridThread.Start()

            ElseIf Callback_Function_Read_Selected_TreeNode(ExplorerTree).Equals("UnUsedSoftGridNode") Then

                ' Start thread
                UnUsedSoftGridThread = New Thread(Sub() UnUsedSoftGridWorker(ConnectionString))
                UnUsedSoftGridThread.Start()

            End If

            ' Sql cleanapps functionality
            Delegate_Sub_Enable_Yellow_Button(btnSqlCafStopCleanApps, True)
            Delegate_Sub_Enable_Tan_Button(btnSqlRunCleanApps, True)

        Catch ex As Exception

            ' Push user alert
            AlertBox.CreateUserAlert(ex.Message)

            ' Write debug
            Delegate_Sub_Append_Text(rtbDebug, CallStack + "Connected failed: " + ex.Message)

            ' Perform disconnect method
            SqlDisconnect()

            ' Return
            Return

        End Try

        ' Encapsulate monitoring loop
        While Not TerminateSignal Or CancelSignal

            ' Check database connection
            If CancelSignal OrElse DbConnection.State = ConnectionState.Broken OrElse DbConnection.State = ConnectionState.Closed Then

                ' Throw cancellation signal (if it wasn't thrown)
                CancelSignal = True

                ' Check if database connection is open
                If Not DbConnection.State = ConnectionState.Closed Then

                    ' Close the database connection
                    DbConnection.Close()

                    ' Write debug
                    Delegate_Sub_Append_Text(rtbDebug, CallStack + "Database connection closed.")

                End If

                ' Perform disconnect method
                SqlDisconnect()

                ' Exit while
                Exit While

            End If

            ' Rest thread
            Thread.Sleep(Globals.THREAD_REST_INTERVAL)

        End While

    End Sub

    Public Function SqlSelectScalar(ByVal CallStack As String,
                                    ByVal DatabaseConnection As SqlConnection,
                                    ByVal QueryText As String) As String

        ' Local variables
        Dim SqlCmd As SqlCommand
        Dim SqlData As SqlDataReader
        Dim ScalarValue As String = ""

        ' Update call stack
        CallStack += "SqlSelectScalar --> "

        ' Encapsulate sql operation
        Try

            ' Check connection state
            If DatabaseConnection.State = ConnectionState.Closed Or DatabaseConnection.State = ConnectionState.Broken Then

                ' Return
                Return Nothing

            End If

            ' Build sql command from user query
            SqlCmd = New SqlCommand(QueryText, DatabaseConnection)

            ' Write debug
            Delegate_Sub_Append_Text(rtbDebug, CallStack + "Execute statement:" + Environment.NewLine + QueryText)

            ' Execute query -- ExecuteReader method
            SqlData = SqlCmd.ExecuteReader()

            ' Check for termination or cancellation signal
            If TerminateSignal Or CancelSignal Then

                ' Close sql data reader
                SqlData.Close()

                ' Return
                Return Nothing

            End If

            ' Check if data returns any rows
            If SqlData.Read() AndAlso SqlData.HasRows Then

                ' Check column position for null
                If SqlData.IsDBNull(0) Then

                    ' Add string null instead of actual null
                    ScalarValue = "Null"

                ElseIf SqlData.GetFieldType(0) Is GetType(Byte()) Then

                    ' Add string representation of binary data
                    ScalarValue = "0x" + BitConverter.ToString(SqlData.GetSqlBinary(0).Value).Replace("-", "")

                Else

                    ' Add raw data
                    ScalarValue = SqlData(0).ToString

                End If

            End If

            ' Close sql data reader
            SqlData.Close()

        Catch ex As Exception

            ' Write debug
            Delegate_Sub_Append_Text(rtbDebug, CallStack + "Exception:" + Environment.NewLine + ex.Message)

            ' Return
            Return ex.Message

        End Try

        ' Return
        Return ScalarValue

    End Function

    Public Function SqlSelectScalarList(ByVal CallStack As String,
                                        ByVal DatabaseConnection As SqlConnection,
                                        ByVal QueryText As String) As List(Of String)

        ' Local variables
        Dim SqlCmd As SqlCommand
        Dim SqlData As SqlDataReader
        Dim ResultArray As New List(Of String)

        ' Update call stack
        CallStack += "SqlSelectScalarList --> "

        ' Encapsulate sql operation
        Try

            ' Check connection state
            If DatabaseConnection.State = ConnectionState.Closed Or DatabaseConnection.State = ConnectionState.Broken Then

                ' Return
                Return Nothing

            End If

            ' Build sql command from user query
            SqlCmd = New SqlCommand(QueryText, DatabaseConnection)

            ' Write debug
            Delegate_Sub_Append_Text(rtbDebug, CallStack + "Execute statement:" + Environment.NewLine + QueryText)

            ' Execute query -- ExecuteReader method
            SqlData = SqlCmd.ExecuteReader()

            ' Check for termination or cancellation signal
            If TerminateSignal Or CancelSignal Then

                ' Close sql data reader
                SqlData.Close()

                ' Return
                Return Nothing

            End If

            ' Loop through sql data results
            Do

                ' Check if data returns any rows
                If SqlData.HasRows Then

                    ' Iterate each record
                    While SqlData.Read

                        ' Check column position for null
                        If SqlData.IsDBNull(0) Then

                            ' Add string null instead of actual null
                            ResultArray.Add("Null")

                        ElseIf SqlData.GetFieldType(0) Is GetType(Byte()) Then

                            ' Add string representation of binary data
                            ResultArray.Add("0x" + BitConverter.ToString(SqlData.GetSqlBinary(0).Value).Replace("-", ""))

                        Else

                            ' Add raw data
                            ResultArray.Add(SqlData(0).ToString)

                        End If

                    End While

                End If

            Loop While SqlData.NextResult

            ' Close sql data reader
            SqlData.Close()

        Catch ex As Exception

            ' Write debug
            Delegate_Sub_Append_Text(rtbDebug, CallStack + "Exception:" + Environment.NewLine + ex.Message)

        End Try

        ' Return
        Return ResultArray

    End Function

    Public Function SqlSelectGrid(ByVal CallStack As String,
                                  ByVal DatabaseConnection As SqlConnection,
                                  ByVal QueryText As String,
                                  Optional ByVal IncludeColumnHeaders As Boolean = False) As List(Of List(Of String))

        ' Local variables
        Dim SqlCmd As SqlCommand
        Dim SqlData As SqlDataReader
        Dim RowArray As New List(Of String)
        Dim ResultArray As New List(Of List(Of String))

        ' Update call stack
        CallStack += "SqlSelectGrid --> "

        ' Encapsulate sql operation
        Try

            ' Check connection state
            If DatabaseConnection.State = ConnectionState.Closed Or DatabaseConnection.State = ConnectionState.Broken Then

                ' Return
                Return Nothing

            End If

            ' Build sql command from user query
            SqlCmd = New SqlCommand(QueryText, DatabaseConnection)

            ' Write debug
            Delegate_Sub_Append_Text(rtbDebug, CallStack + "Execute statement:" + Environment.NewLine + QueryText)

            ' Execute query -- ExecuteReader method
            SqlData = SqlCmd.ExecuteReader()

            ' Check for termination or cancellation signal
            If TerminateSignal Or CancelSignal Then

                ' Close sql data reader
                SqlData.Close()

                ' Return
                Return Nothing

            End If

            ' Loop through sql data results
            Do

                ' Check if data returns any rows
                If SqlData.HasRows Then

                    ' Check if column headers should be included in resultset
                    If IncludeColumnHeaders Then

                        ' Iterate through table columns names
                        For i As Integer = 0 To SqlData.FieldCount - 1

                            ' Check for a column name
                            If SqlData.GetName(i).Equals("") Then

                                ' Stub an empty column name
                                RowArray.Add("(No column name)")

                            Else

                                ' Set column name
                                RowArray.Add(SqlData.GetName(i))

                            End If

                            ' Process message queue
                            Application.DoEvents()

                        Next

                        ' Add columns names as first row of the result
                        ResultArray.Add(RowArray)

                    End If

                    ' Iterate each record
                    While SqlData.Read

                        ' Initialize array for row data
                        RowArray = New List(Of String)

                        ' Iterate each column of the record
                        For i As Integer = 0 To SqlData.FieldCount - 1

                            ' Check column position for null
                            If SqlData.IsDBNull(i) Then

                                ' Add string null instead of actual null
                                RowArray.Add("Null")

                            ElseIf SqlData.GetFieldType(i) Is GetType(Byte()) Then

                                ' Add string representation of binary data
                                RowArray.Add("0x" + BitConverter.ToString(SqlData.GetSqlBinary(i).Value).Replace("-", ""))

                            Else

                                ' Add raw data
                                RowArray.Add(SqlData(i))

                            End If

                            ' Process message queue
                            Application.DoEvents()

                        Next

                        ' Add row (arraylist) to result (arraylist)
                        ResultArray.Add(RowArray)

                        ' Check for termination or cancellation signal before processing next record
                        If TerminateSignal Or CancelSignal Then

                            ' Close sql data reader
                            SqlData.Close()

                            ' Return
                            Return Nothing

                        End If

                    End While

                End If

                ' Check for termination or cancellation signal between processing data results
                If TerminateSignal Or CancelSignal Then

                    ' Close sql data reader
                    SqlData.Close()

                    ' Return
                    Return Nothing

                End If

            Loop While SqlData.NextResult

            ' Close sql data reader
            SqlData.Close()

        Catch ex As Exception

            ' Write debug
            Delegate_Sub_Append_Text(rtbDebug, CallStack + "Exception:" + Environment.NewLine + ex.Message)

        End Try

        ' Return
        Return ResultArray

    End Function

    Public Sub SqlPopulateGridWorker(ByVal CallStack As String,
                                     ByVal DatabaseConnection As SqlConnection,
                                     ByVal QueryText As String,
                                     ByRef dgvResultGrid As DataGridView,
                                     Optional ByVal FillColumn As String = "",
                                     Optional ByVal AutoResizeColumns As DataGridViewAutoSizeColumnsMode = DataGridViewAutoSizeColumnMode.NotSet,
                                     Optional ByVal AutoResizeRows As DataGridViewAutoSizeRowsMode = DataGridViewAutoSizeColumnMode.NotSet,
                                     Optional ByVal LoadingBar As ProgressBar = Nothing,
                                     Optional ByVal LoadingControl As Control = Nothing,
                                     Optional ByVal HideColumns As List(Of String) = Nothing)

        ' Local variables
        Dim SqlCmd As SqlCommand
        Dim SqlReader As SqlDataReader
        Dim SqlResultTable As New DataTable
        Dim SqlDataRow As DataRow
        Dim SqlResultRow As New ArrayList
        Dim PercentLoaded As Integer = 0
        Dim PreviousPercentage As Integer = 0

        ' Update call stack
        CallStack += "SqlPopulateGridWorker --> "

        ' Clear data grid
        Delegate_Sub_Clear_DataGridView(dgvResultGrid)

        ' Encapsulate sql operation
        Try

            ' Check connection state
            If DatabaseConnection.State = ConnectionState.Closed Or DatabaseConnection.State = ConnectionState.Broken Then Return

            ' Set loading message on control's text field (if available)
            If LoadingControl IsNot Nothing Then

                ' Set initial percentage
                Delegate_Sub_Set_Text(LoadingControl, LoadingControl.Text + " (Loading)")

            End If

            ' Build sql command from provided query
            SqlCmd = New SqlCommand(QueryText, DatabaseConnection)

            ' Write debug
            Delegate_Sub_Append_Text(rtbDebug, CallStack + "Execute statement:" + Environment.NewLine + QueryText)

            ' Execute query -- ExecuteReader method
            SqlReader = SqlCmd.ExecuteReader()

            ' Populate result table
            SqlResultTable.Load(SqlReader)

            ' Close sql data reader
            SqlReader.Close()

            ' Check for termination or cancellation signal
            If TerminateSignal Or CancelSignal Then Return

            ' Set progress bar scale (if available)
            If LoadingBar IsNot Nothing AndAlso SqlResultTable.Rows.Count > 0 Then

                ' Update progress bar values
                LoadingBar.Invoke(Sub() LoadingBar.Value = 0)
                LoadingBar.Invoke(Sub() LoadingBar.Maximum = SqlResultTable.Rows.Count)

            End If

            ' Set initial percentage on control's text field (if available)
            If LoadingControl IsNot Nothing Then

                ' Set initial percentage
                Delegate_Sub_Set_Text(LoadingControl, LoadingControl.Text + " (" + PercentLoaded.ToString + "%)")

            End If

            ' Iterate result table columns
            For i As Integer = 0 To SqlResultTable.Columns.Count - 1

                ' Check for a column name
                If SqlResultTable.Columns(i).ColumnName.Equals("") Then

                    ' Stub an empty column name
                    Delegate_Sub_Add_DataGridView_Column(dgvResultGrid, "col_" + i.ToString, "(No column name)")

                Else

                    ' Provide actual column name
                    Delegate_Sub_Add_DataGridView_Column(dgvResultGrid, SqlResultTable.Columns(i).ColumnName, SqlResultTable.Columns(i).ColumnName)

                End If

                ' Process message queue
                Application.DoEvents()

            Next

            ' Check for hidden columns
            If HideColumns IsNot Nothing Then

                ' Iterate list of columns to be hidden
                For Each HiddenColumn As String In HideColumns

                    ' Hide column
                    Delegate_Sub_Hide_DataGridView_Column(dgvResultGrid, HiddenColumn)

                Next

            End If

            ' Iterate result table rows
            For i As Integer = 0 To SqlResultTable.Rows.Count - 1

                ' Fetch data row
                SqlDataRow = SqlResultTable.Rows(i)

                ' Iterate each column position of the row
                For j As Integer = 0 To SqlDataRow.ItemArray.Length - 1

                    ' Check column position for null
                    If SqlDataRow.IsNull(j) Then

                        ' Add string null instead of actual null
                        SqlResultRow.Add("Null")

                    ElseIf SqlDataRow.Item(j).GetType Is GetType(Byte()) Then

                        ' Add string representation of binary data
                        SqlResultRow.Add("0x" + BitConverter.ToString(SqlDataRow.Item(j)).Replace("-", ""))

                    Else

                        ' Add raw data
                        SqlResultRow.Add(SqlDataRow.Item(j))

                    End If

                    ' Check for termination or cancellation signal
                    If TerminateSignal Or CancelSignal Then Return

                    ' Process message queue
                    Application.DoEvents()

                Next

                ' Add the row to the data grid
                Delegate_Sub_Add_DataGridView_Row(dgvResultGrid, SqlResultRow.ToArray)

                ' Auto resize on first ten records and the last record
                If i < 10 OrElse i = SqlResultTable.Rows.Count - 1 Then

                    ' Check if column resize parameter was specified
                    If AutoResizeColumns <> 0 Then

                        ' Perform one-time auto resize of columns
                        Delegate_Sub_Resize_DataGrid_Columns(dgvResultGrid, AutoResizeColumns)

                    End If

                    ' Check if a row resize parameter was specified
                    If AutoResizeRows <> 0 Then

                        ' Perform one-time auto resize of rows
                        Delegate_Sub_Resize_DataGrid_Rows(dgvResultGrid, AutoResizeRows)

                        ' Adjust row template for future rows
                        If i = 0 Then Delegate_Sub_Resize_DataGrid_Row_Template_Height(dgvResultGrid, dgvResultGrid.Rows(0).Height)

                    End If

                    ' Check if fill column was specified
                    If Not FillColumn.Equals("") Then

                        ' Set fill column
                        Delegate_Sub_Set_DataGrid_Fill_Column(dgvResultGrid, FillColumn)

                    End If

                End If

                ' Clear the result list
                SqlResultRow.Clear()

                ' Calculate percent loaded
                PercentLoaded = (SqlResultTable.Rows.IndexOf(SqlDataRow) / SqlResultTable.Rows.Count) * 100

                ' Increment progress bar (if available)
                If LoadingBar IsNot Nothing Then

                    ' Increment progress bar
                    LoadingBar.Invoke(Sub() LoadingBar.Increment(1))

                End If

                ' Increment control text (if available) only if there's a change in percentage-- this is to reduce flickering of the control
                If LoadingControl IsNot Nothing AndAlso PreviousPercentage <> PercentLoaded Then

                    ' Update percentage
                    Delegate_Sub_Set_Text(LoadingControl, LoadingControl.Text.Substring(0, LoadingControl.Text.IndexOf("(") - 1) + " (" + PercentLoaded.ToString + "%)")

                End If

                ' Update previous percentage only if there's a full integer change
                If PreviousPercentage <> PercentLoaded Then

                    ' Move up previous percentage
                    PreviousPercentage = PercentLoaded

                End If

                ' Check for termination or cancellation signal
                If TerminateSignal Or CancelSignal Then Return

                ' Process message queue
                Application.DoEvents()

            Next

        Catch ex As Exception

            ' Write debug
            Delegate_Sub_Append_Text(rtbDebug, CallStack + "Exception:" + Environment.NewLine + ex.Message)

            ' Add an exception column
            Delegate_Sub_Add_DataGridView_Column(dgvResultGrid, "Exception", "Exception")

            ' Add error message as row
            Delegate_Sub_Add_DataGridView_Row(dgvResultGrid, {ex.Message})

            ' Perform one-time auto resize of columns
            Delegate_Sub_Resize_DataGrid_Columns(dgvResultGrid, AutoResizeColumns)

        Finally

            ' Reset progress bar (if available)
            If LoadingBar IsNot Nothing Then

                ' Reset progress bar
                LoadingBar.Invoke(Sub() LoadingBar.Value = 0)

            End If

            ' Reset control text (if available)
            If LoadingControl IsNot Nothing Then

                ' Reset control text
                Delegate_Sub_Set_Text(LoadingControl, LoadingControl.Text.Substring(0, LoadingControl.Text.IndexOf("(") - 1) + " (" + SqlResultTable.Rows.Count.ToString + ")")

            End If

        End Try

    End Sub

    Public Function SqlSimplePopulateGridWorker(ByVal CallStack As String,
                                                ByVal DatabaseConnection As SqlConnection,
                                                ByVal QueryText As String,
                                                ByRef dgvResultGrid As DataGridView,
                                                Optional ByVal FillColumn As String = "") As Integer

        ' Local variables
        Dim SqlCmd As SqlCommand
        Dim SqlReader As SqlDataReader
        Dim SqlResultRow As New ArrayList
        Dim TotalRecordReturnCount As Integer = 0
        Dim BatchRecordReturnCount As Integer = 0

        ' Update call stack
        CallStack += "SqlSimplePopulateGridWorker --> "

        ' Clear data grid
        Delegate_Sub_Clear_DataGridView(dgvResultGrid)

        ' Encapsulate sql operation
        Try

            ' Check connection state
            If DatabaseConnection.State = ConnectionState.Closed Or DatabaseConnection.State = ConnectionState.Broken Then

                ' Return
                Return -1

            End If

            ' Build sql command from provided query
            SqlCmd = New SqlCommand(QueryText, DatabaseConnection)

            ' Write debug
            Delegate_Sub_Append_Text(rtbDebug, CallStack + "Execute statement:" + Environment.NewLine + QueryText)

            ' Execute query -- ExecuteReader method
            SqlReader = SqlCmd.ExecuteReader()

            ' Check for termination or cancellation signal
            If TerminateSignal Or CancelSignal Then

                ' Close sql data reader
                SqlReader.Close()

                ' Return
                Return -1

            End If

            ' Iterate sql reader results
            Do

                ' Check if reader has any rows
                If SqlReader.HasRows Then

                    ' Iterate through table columns
                    For i As Integer = 0 To SqlReader.FieldCount - 1

                        ' Check for a column name
                        If SqlReader.GetName(i).Equals("") Then

                            ' Stub an empty column name
                            Delegate_Sub_Add_DataGridView_Column(dgvResultGrid, "col_" + i.ToString, "(No column name)")

                        Else

                            ' Provide actual column name
                            Delegate_Sub_Add_DataGridView_Column(dgvResultGrid, SqlReader.GetName(i), SqlReader.GetName(i))

                        End If

                        ' Process message queue
                        Application.DoEvents()

                    Next

                    ' Check if fill column was passed
                    If Not FillColumn.Equals("") Then

                        ' Set fill column
                        Delegate_Sub_Set_DataGrid_Fill_Column(dgvResultGrid, FillColumn)

                    End If

                    ' Iterate each row
                    While SqlReader.Read

                        ' Iterate each column position of the row
                        For i As Integer = 0 To SqlReader.FieldCount - 1

                            ' Check column position for null
                            If SqlReader.IsDBNull(i) Then

                                ' Add string null instead of actual null
                                SqlResultRow.Add("Null")

                            ElseIf SqlReader.GetFieldType(i) Is GetType(Byte()) Then

                                ' Add string representation of binary data
                                SqlResultRow.Add("0x" + BitConverter.ToString(SqlReader.GetSqlBinary(i).Value).Replace("-", ""))

                            Else

                                ' Add raw data
                                SqlResultRow.Add(SqlReader(i))

                            End If

                            ' Check for termination or cancellation signal before processing next record
                            If TerminateSignal Or CancelSignal Then

                                ' Close sql data reader
                                SqlReader.Close()

                                ' Return
                                Return -1

                            End If

                            ' Process message queue
                            Application.DoEvents()

                        Next

                        ' Add the row to the data grid
                        Delegate_Sub_Add_DataGridView_Row(dgvResultGrid, SqlResultRow.ToArray)

                        ' Clear the result list
                        SqlResultRow.Clear()

                        ' Increment the record counter
                        BatchRecordReturnCount += 1

                        ' Check for termination or cancellation signal before processing next record
                        If TerminateSignal Or CancelSignal Then

                            ' Close sql data reader
                            SqlReader.Close()

                            ' Return
                            Return -1

                        End If

                        ' Process message queue
                        Application.DoEvents()

                    End While

                End If

                ' Increment total record return count
                TotalRecordReturnCount += BatchRecordReturnCount

                ' Reset batch record counter
                BatchRecordReturnCount = 0

                ' Check for termination or cancellation signal between processing data results
                If TerminateSignal Or CancelSignal Then

                    ' Close sql data reader
                    SqlReader.Close()

                    ' Return
                    Return -1

                End If

            Loop While SqlReader.NextResult

            ' Close sql data reader
            SqlReader.Close()

        Catch ex As Exception

            ' Write debug
            Delegate_Sub_Append_Text(rtbDebug, CallStack + "Exception:" + Environment.NewLine + ex.Message)

        End Try

        ' Return
        Return TotalRecordReturnCount

    End Function

End Class