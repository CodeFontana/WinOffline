Imports System.Threading

Partial Public Class WinOfflineUI

    Private Shared RemovalToolThread As Thread

    Private Sub rbnRemoveITCM_CheckedChanged(sender As Object, e As EventArgs) Handles rbnRemoveITCM.CheckedChanged

        ' Check status
        If rbnRemoveITCM.Checked Then

            ' Update globals
            Globals.RemoveITCM = True
            Globals.UninstallITCM = False

            ' Enable hostuuid checkbox
            Delegate_Sub_Enable_Control(chkRetainHostUUID, True)

        End If

    End Sub

    Private Sub rbnUninstallITCM_CheckedChanged(sender As Object, e As EventArgs) Handles rbnUninstallITCM.CheckedChanged

        ' Check status
        If rbnUninstallITCM.Checked Then

            ' Update globals
            Globals.RemoveITCM = False
            Globals.UninstallITCM = True

            ' Disable hostuuid checkbox
            chkRetainHostUUID.Checked = False
            Delegate_Sub_Enable_Control(chkRetainHostUUID, False)

        End If

    End Sub

    Private Sub chkRetainHostUUID_CheckedChanged(sender As Object, e As EventArgs) Handles chkRetainHostUUID.CheckedChanged

        ' Check status
        If chkRetainHostUUID.Checked Then

            ' Update globals
            Globals.KeepHostUUIDSwitch = True

        Else

            ' Update globals
            Globals.KeepHostUUIDSwitch = False

        End If

    End Sub

    Private Sub btnRemoveITCM_Click(sender As Object, e As EventArgs) Handles btnRemoveITCM.Click

        ' Disable controls
        Delegate_Sub_Enable_Control(btnRemoveITCM, False)
        Delegate_Sub_Enable_Control(rbnRemoveITCM, False)
        Delegate_Sub_Enable_Control(chkRetainHostUUID, False)
        Delegate_Sub_Enable_Control(rbnUninstallITCM, False)
        Delegate_Sub_Enable_Control(lblRemoveITCMCaution, False)

        ' Clear output window contents
        txtRemovalTool.Text = ""

        ' Create external thread
        RemovalToolThread = New Thread(Sub() WinOffline.RemoveITCM(WinOffline.Logger.LastCallStack))
        RemovalToolThread.Start()

    End Sub

    Private Sub btnExitRemoveITCM_Click(sender As Object, e As EventArgs) Handles btnExitRemoveITCM.Click

        ' Close the dialog
        Close()

    End Sub

End Class