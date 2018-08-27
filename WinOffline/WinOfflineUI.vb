Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Windows.Forms

Partial Public Class WinOfflineUI

    Private Const MF_STRING As Integer = &H0
    Private Const MF_SEPARATOR As Integer = &H800
    Private Const WM_SYSCOMMAND As Int32 = &H112
    Private SystemInfoUpdateThread As Thread
    Private HiddenNodes As New ArrayList
    Private TerminateSignal As Boolean = False

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function GetSystemMenu(
        <[In]()> ByVal hWnd As IntPtr,
        <[In]()> ByVal bRevert As Boolean) _
        As IntPtr
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function AppendMenu(
        <[In]()> ByVal hMenu As IntPtr,
        <[In]()> ByVal uFlags As Int32,
        <[In]()> ByVal uIDNewItem As IntPtr,
        <[In](), [Optional]> ByVal lpNewItem As String) _
        As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    Private Sub WinOfflineUI_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' Add about menu option to system menu
        AddSysMenuItems()

        ' Set titlebar text
        Text = Globals.ProcessFriendlyName + " -- " + Globals.AppVersion

        ' Set autoscroll minimum size
        pnlWinOfflineStart.AutoScrollMinSize = pnlWinOfflineStart.Size
        pnlWinOfflineMain.AutoScrollMinSize = pnlWinOfflineMain.Size
        pnlWinOfflineRemove.AutoScrollMinSize = pnlWinOfflineRemove.Size
        pnlWinOfflineSDHelp.AutoScrollMinSize = pnlWinOfflineSDHelp.Size
        pnlWinOfflineCLIHelp.AutoScrollMinSize = pnlWinOfflineCLIHelp.Size
        pnlENCOverdrive.AutoScrollMinSize = pnlENCOverdrive.Size
        pnlSqlMdbOverview.AutoScrollMinSize = pnlSqlMdbOverview.Size
        pnlSqlTableSpaceGrid.AutoScrollMinSize = pnlSqlTableSpaceGrid.Size
        pnlSqlAgentGrid.AutoScrollMinSize = pnlSqlAgentGrid.Size
        pnlSqlUserGrid.AutoScrollMinSize = pnlSqlUserGrid.Size
        pnlSqlServerGrid.AutoScrollMinSize = pnlSqlServerGrid.Size
        pnlSqlEngineGrid.AutoScrollMinSize = pnlSqlEngineGrid.Size
        pnlSqlGroupEvalGrid.AutoScrollMinSize = pnlSqlGroupEvalGrid.Size
        pnlSqlInstSoftGrid.AutoScrollMinSize = pnlSqlInstSoftGrid.Size
        pnlSqlDiscSoftGrid.AutoScrollMinSize = pnlSqlDiscSoftGrid.Size
        pnlSqlDuplCompGrid.AutoScrollMinSize = pnlSqlDuplCompGrid.Size
        pnlSqlUnUsedSoftGrid.AutoScrollMinSize = pnlSqlUnUsedSoftGrid.Size
        pnlSqlCleanApps.AutoScrollMinSize = pnlSqlCleanApps.Size
        pnlSqlQueryEditor.AutoScrollMinSize = pnlSqlQueryEditor.Size
        pnlRemovalTool.AutoScrollMinSize = pnlRemovalTool.Size
        pnlSystemInfo.AutoScrollMinSize = pnlSystemInfo.Size
        pnlDebug.AutoScrollMinSize = pnlDebug.Size

        ' Perform panel initializations
        InitWinOffline()
        InitSystemInfo()
        InitEncOverdrive()
        InitSqlMdbOverview()
        InitSqlTableSpaceGrid()
        InitSqlAgentGrid()
        InitSqlUserGrid()
        InitSqlServerGrid()
        InitSqlEngineGrid()
        InitSqlGroupEvalGrid()
        InitSqlInstSoftGrid()
        InitSqlDiscSoftGrid()
        InitSqlDuplCompGrid()
        InitSqlUnUsedSoftGrid()
        InitSqlCleanApps()
        InitSqlQueryEditor()

        ' Set WinOffline as the default node
        ExplorerTree.SelectedNode = GetTreeNodebyName("WinOfflineNode")

        ' Expand all tree nodes automatically
        ExplorerTree.ExpandAll()

        ' Recurse tree node names for reflexsive "WinOffline" references
        RecurseTreeName(ExplorerTree)

        ' Hide the debug node
        HiddenNodes.Add(GetTreeNodebyName("DebugNode"))
        ExplorerTree.Nodes(0).Nodes.Remove(GetTreeNodebyName("DebugNode"))

        ' Resize form based on user DPI setting (100% = 96 DPI)
        Using myGraphics As Drawing.Graphics = CreateGraphics()
            Width = myGraphics.DpiX / 96 * Size.Width
            Height = myGraphics.DpiY / 96 * Size.Height
        End Using

    End Sub

    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)

        ' Feed message to base class
        MyBase.WndProc(m)

        ' Check for system command message (i.e. system/control menu)
        If (m.Msg = WM_SYSCOMMAND) Then

            ' Check WParam field of the message
            Select Case m.WParam.ToInt32

                ' String case
                Case MF_STRING

                    ' Create and show the about box
                    Dim myAboutBox As New AboutBox
                    myAboutBox.ShowDialog()

            End Select

        End If

    End Sub

    Private Sub AddSysMenuItems()

        ' Get system menu handle
        Dim hSysMenu As IntPtr = GetSystemMenu(Me.Handle, False)

        ' Apend a menu separator and "About" box link
        AppendMenu(hSysMenu, MF_SEPARATOR, IntPtr.Zero, Nothing)
        AppendMenu(hSysMenu, MF_STRING, IntPtr.Zero, "About")

    End Sub

    Private Sub WinOfflineUI_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        ' Hide the form
        Visible = False

        ' Set termination signals
        CancelSignal = True
        TerminateSignal = True

        ' Loop while external threads are still running
        While True

            ' Check if external threads are still alive
            If (SystemInfoUpdateThread IsNot Nothing AndAlso SystemInfoUpdateThread.IsAlive) OrElse
                (PatchListViewThread IsNot Nothing AndAlso PatchListViewThread.IsAlive) OrElse
                (EncClientMonitorThread IsNot Nothing AndAlso EncClientMonitorThread.IsAlive) OrElse
                (EncOverdriveThread IsNot Nothing AndAlso EncOverdriveThread.IsAlive) OrElse
                (SqlCleanAppsThread IsNot Nothing AndAlso SqlCleanAppsThread.IsAlive) OrElse
                (MdbOverviewThread IsNot Nothing AndAlso MdbOverviewThread.IsAlive) OrElse
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
                (DatabaseThread IsNot Nothing AndAlso DatabaseThread.IsAlive) OrElse
                (RemovalToolThread IsNot Nothing AndAlso RemovalToolThread.IsAlive) Then

                ' Process message queue
                Application.DoEvents()

            Else

                ' All threads finished -- exit loop
                Exit While

            End If

        End While

        ' Check dialog result
        If DialogResult <> DialogResult.OK Then

            ' Pass underlying WinOffline the about signal
            DialogResult = DialogResult.Abort

        End If

    End Sub

    Private Sub WinOfflineUI_KeyDown(sender As Object, ByVal e As KeyEventArgs) Handles MyBase.KeyDown

        ' Check for Ctrl+Shift+D signal
        If e.Control AndAlso e.Shift AndAlso e.KeyCode = Keys.D Then

            ' Check if debug panel is visible
            If pnlDebug.Visible Then

                ' Hide the debug panel and node
                HiddenNodes.Add(GetTreeNodebyName("DebugNode"))
                ExplorerTree.Nodes(0).Nodes.Remove(GetTreeNodebyName("DebugNode"))
                pnlDebug.Visible = False

            Else

                ' Show the debug panel and node
                ExplorerTree.Nodes(0).Nodes.Add(GetHiddenNodebyName("DebugNode"))
                HiddenNodes.Remove(GetHiddenNodebyName("DebugNode"))
                ExplorerTree.SelectedNode = GetTreeNodebyName("DebugNode")

            End If

        End If

        ' Check for enter key
        If e.KeyCode = Keys.Enter Then

            ' Check which panel is visible
            If pnlSqlMdbOverview.Visible Then

                ' Press the connect button
                btnSqlConnectMdbOverview.PerformClick()

                ' Cancel the key
                e.SuppressKeyPress = True

            ElseIf pnlSqlTableSpaceGrid.Visible Then

                ' Press the connect button
                btnSqlConnectTableSpace.PerformClick()

                ' Cancel the key
                e.SuppressKeyPress = True

            ElseIf pnlSqlAgentGrid.Visible Then

                ' Press the connect button
                btnSqlConnectAgentGrid.PerformClick()

                ' Cancel the key
                e.SuppressKeyPress = True

            ElseIf pnlSqlUserGrid.Visible Then

                ' Press the connect button
                btnSqlConnectUserGrid.PerformClick()

                ' Cancel the key
                e.SuppressKeyPress = True

            ElseIf pnlSqlServerGrid.Visible Then

                ' Press the connect button
                btnSqlConnectServerGrid.PerformClick()

                ' Cancel the key
                e.SuppressKeyPress = True

            ElseIf pnlSqlEngineGrid.Visible Then

                ' Press the connect button
                btnSqlConnectEngineGrid.PerformClick()

                ' Cancel the key
                e.SuppressKeyPress = True

            ElseIf pnlSqlGroupEvalGrid.Visible Then

                ' Press the connect button
                btnSqlConnectGroupEvalGrid.PerformClick()

                ' Cancel the key
                e.SuppressKeyPress = True

            ElseIf pnlSqlInstSoftGrid.Visible Then

                ' Press the connect button
                btnSqlConnectInstSoftGrid.PerformClick()

                ' Cancel the key
                e.SuppressKeyPress = True

            ElseIf pnlSqlDiscSoftGrid.Visible Then

                ' Press the connect button
                btnSqlConnectDiscSoftGrid.PerformClick()

                ' Cancel the key
                e.SuppressKeyPress = True

            ElseIf pnlSqlDuplCompGrid.Visible Then

                ' Press the connect button
                btnSqlConnectDuplCompGrid.PerformClick()

                ' Cancel the key
                e.SuppressKeyPress = True

            ElseIf pnlSqlUnUsedSoftGrid.Visible Then

                ' Press the connect button
                btnSqlConnectUnUsedSoftGrid.PerformClick()

                ' Cancel the key
                e.SuppressKeyPress = True

            ElseIf pnlSqlCleanApps.Visible Then

                ' Press the connect button
                btnSqlConnectCleanApps.PerformClick()

                ' Cancel the key
                e.SuppressKeyPress = True

            ElseIf pnlSqlQueryEditor.Visible Then

                ' Press the connect button
                btnSqlConnectQueryEditor.PerformClick()

                ' Cancel the key
                e.SuppressKeyPress = True

            End If

        End If

    End Sub

    Private Sub RecurseTreeName(ByRef t As TreeView)

        ' Iterate each tree node
        For Each n As TreeNode In t.Nodes

            ' Process tree node
            RecurseNodeNames(n)

        Next

    End Sub

    Private Sub RecurseNodeNames(ByRef n As TreeNode)

        ' Replace "WinOffline" references with process name
        If n.Text.Contains("WinOffline") Then n.Text = n.Text.Replace("WinOffline", Globals.ProcessFriendlyName)

        ' Iterate child tree nodes
        For Each t As TreeNode In n.Nodes

            ' Recurse child tree nodes
            If t.GetNodeCount(True) > 0 Then RecurseNodeNames(t)

            ' Replace "WinOffline" references with process name
            If t.Text.Contains("WinOffline") Then t.Text = t.Text.Replace("WinOffline", Globals.ProcessFriendlyName)

        Next

    End Sub

    Private Function GetTreeNodebyName(ByVal Name As String) As TreeNode

        ' Iterate each tree node
        For Each Node As TreeNode In ExplorerTree.Nodes(0).Nodes

            ' Return matching node
            If Node.Name.ToLower.Equals(Name.ToLower) Then Return Node

        Next

        ' Return
        Return Nothing

    End Function

    Private Function GetHiddenNodebyName(ByVal Name As String) As TreeNode

        ' Iterate each hidden tree node
        For Each Node As TreeNode In HiddenNodes

            ' Return matching node
            If Node.Name.ToLower.Equals(Name.ToLower) Then Return Node

        Next

        ' Return
        Return Nothing

    End Function

    Private Sub HideAllPanels()

        ' Hide all panels
        pnlWinOfflineStart.Visible = False
        pnlWinOfflineMain.Visible = False
        pnlWinOfflineRemove.Visible = False
        pnlWinOfflineCLIHelp.Visible = False
        pnlWinOfflineSDHelp.Visible = False
        pnlENCOverdrive.Visible = False
        pnlSqlMdbOverview.Visible = False
        pnlSqlTableSpaceGrid.Visible = False
        pnlSqlAgentGrid.Visible = False
        pnlSqlUserGrid.Visible = False
        pnlSqlServerGrid.Visible = False
        pnlSqlEngineGrid.Visible = False
        pnlSqlGroupEvalGrid.Visible = False
        pnlSqlInstSoftGrid.Visible = False
        pnlSqlDiscSoftGrid.Visible = False
        pnlSqlDuplCompGrid.Visible = False
        pnlSqlUnUsedSoftGrid.Visible = False
        pnlSqlCleanApps.Visible = False
        pnlSqlQueryEditor.Visible = False
        pnlRemovalTool.Visible = False
        pnlSystemInfo.Visible = False
        pnlDebug.Visible = False

    End Sub

    Private Sub ExplorerTree_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles ExplorerTree.AfterSelect

        ' Hide all current panels
        HideAllPanels()

        ' Check selected node name
        If e.Node.Name.Equals("RootNode") Or e.Node.Name.Equals("WinOfflineNode") Then

            ' WinOffline start panel visible
            pnlWinOfflineStart.Visible = True

        ElseIf e.Node.Name.Equals("SqlMdbOverviewNode") Then

            ' Sql mdb overview panel visible
            pnlSqlMdbOverview.Visible = True

            ' Check if database thread is running and panel thread is unintialized
            If DatabaseThread IsNot Nothing AndAlso DatabaseThread.IsAlive AndAlso MdbOverviewThread Is Nothing Then

                ' Start panel thread
                MdbOverviewThread = New Thread(Sub() MdbOverviewWorker(ConnectionString))
                MdbOverviewThread.Start()

            End If

        ElseIf e.Node.Name.Equals("SqlTableSpaceGridNode") Then

            ' Sql mdb tablespace panel visible
            pnlSqlTableSpaceGrid.Visible = True

            ' Check if database thread is running and panel thread is unintialized
            If DatabaseThread IsNot Nothing AndAlso DatabaseThread.IsAlive AndAlso TableSpaceGridThread Is Nothing Then

                ' Start panel thread
                TableSpaceGridThread = New Thread(Sub() TableSpaceGridWorker(ConnectionString))
                TableSpaceGridThread.Start()

            End If

        ElseIf e.Node.Name.Equals("AgentGridNode") Then

            ' Sql agent grid panel visible
            pnlSqlAgentGrid.Visible = True

            ' Check if database thread is running and panel thread is unintialized
            If DatabaseThread IsNot Nothing AndAlso DatabaseThread.IsAlive AndAlso AgentGridThread Is Nothing Then

                ' Start panel thread
                AgentGridThread = New Thread(Sub() AgentGridWorker(ConnectionString))
                AgentGridThread.Start()

            End If

        ElseIf e.Node.Name.Equals("UserGridNode") Then

            ' Sql user grid panel visible
            pnlSqlUserGrid.Visible = True

            ' Check if database thread is running and panel thread is unintialized
            If DatabaseThread IsNot Nothing AndAlso DatabaseThread.IsAlive AndAlso UserGridThread Is Nothing Then

                ' Start panel thread
                UserGridThread = New Thread(Sub() UserGridWorker(ConnectionString))
                UserGridThread.Start()

            End If

        ElseIf e.Node.Name.Equals("ServerGridNode") Then

            ' Sql server grid panel visible
            pnlSqlServerGrid.Visible = True

            ' Check if database thread is running and panel thread is unintialized
            If DatabaseThread IsNot Nothing AndAlso DatabaseThread.IsAlive AndAlso ServerGridThread Is Nothing Then

                ' Start panel thread
                ServerGridThread = New Thread(Sub() ServerGridWorker(ConnectionString))
                ServerGridThread.Start()

            End If

        ElseIf e.Node.Name.Equals("EngineGridNode") Then

            ' Sql engine grid panel visible
            pnlSqlEngineGrid.Visible = True

            ' Check if database thread is running and panel thread is unintialized
            If DatabaseThread IsNot Nothing AndAlso DatabaseThread.IsAlive AndAlso EngineGridThread Is Nothing Then

                ' Start panel thread
                EngineGridThread = New Thread(Sub() EngineGridWorker(ConnectionString))
                EngineGridThread.Start()

            End If

        ElseIf e.Node.Name.Equals("GroupEvalGridNode") Then

            ' Sql group eval grid panel visible
            pnlSqlGroupEvalGrid.Visible = True

            ' Check if database thread is running and panel thread is unintialized
            If DatabaseThread IsNot Nothing AndAlso DatabaseThread.IsAlive AndAlso GroupEvalGridThread Is Nothing Then

                ' Start panel thread
                GroupEvalGridThread = New Thread(Sub() GroupEvalGridWorker(ConnectionString))
                GroupEvalGridThread.Start()

            End If

        ElseIf e.Node.Name.Equals("InstSoftGridNode") Then

            ' Sql installed software grid panel visible
            pnlSqlInstSoftGrid.Visible = True

            ' Check if database thread is running and panel thread is unintialized
            If DatabaseThread IsNot Nothing AndAlso DatabaseThread.IsAlive AndAlso InstSoftGridThread Is Nothing Then

                ' Start panel thread
                InstSoftGridThread = New Thread(Sub() InstSoftGridWorker(ConnectionString))
                InstSoftGridThread.Start()

            End If

        ElseIf e.Node.Name.Equals("DiscSoftGridNode") Then

            ' Sql discovered software grid panel visible
            pnlSqlDiscSoftGrid.Visible = True

            ' Check if database thread is running and panel thread is unintialized
            If DatabaseThread IsNot Nothing AndAlso DatabaseThread.IsAlive AndAlso DiscSoftGridThread Is Nothing Then

                ' Start panel thread
                DiscSoftGridThread = New Thread(Sub() DiscSoftGridWorker(ConnectionString))
                DiscSoftGridThread.Start()

            End If

        ElseIf e.Node.Name.Equals("DuplCompGridNode") Then

            ' Sql duplicate computer grid panel visible
            pnlSqlDuplCompGrid.Visible = True

            ' Check if database thread is running and panel thread is unintialized
            If DatabaseThread IsNot Nothing AndAlso DatabaseThread.IsAlive AndAlso DuplCompGridThread Is Nothing Then

                ' Start panel thread
                DuplCompGridThread = New Thread(Sub() DuplCompGridWorker(ConnectionString))
                DuplCompGridThread.Start()

            End If

        ElseIf e.Node.Name.Equals("UnUsedSoftGridNode") Then

            ' Sql unused software grid panel visible
            pnlSqlUnUsedSoftGrid.Visible = True

            ' Check if database thread is running and panel thread is unintialized
            If DatabaseThread IsNot Nothing AndAlso DatabaseThread.IsAlive AndAlso UnUsedSoftGridThread Is Nothing Then

                ' Start panel thread
                UnUsedSoftGridThread = New Thread(Sub() UnUsedSoftGridWorker(ConnectionString))
                UnUsedSoftGridThread.Start()

            End If

        ElseIf e.Node.Name.Equals("SqlCleanAppsNode") Then

            ' Sql cleanApps panel visible
            pnlSqlCleanApps.Visible = True

        ElseIf e.Node.Name.Equals("SqlQueryNode") Then

            ' Sql query panel visible
            pnlSqlQueryEditor.Visible = True

        ElseIf e.Node.Name.Equals("EncToolsNode") Or e.Node.Name.Equals("EncStressNode") Then

            ' Enc overdrive panel visible
            pnlENCOverdrive.Visible = True

        ElseIf e.Node.Name.Equals("RemovalToolNode") Then

            ' Removal tool panel visible
            pnlRemovalTool.Visible = True

        ElseIf e.Node.Name.Equals("SysInfoNode") Then

            ' System info panel visible
            pnlSystemInfo.Visible = True

        ElseIf e.Node.Name.Equals("DebugNode") Then

            ' Debug panel visible
            pnlDebug.Visible = True

            ' Autoscroll
            rtbDebug.ScrollToCaret()

        End If

    End Sub

End Class