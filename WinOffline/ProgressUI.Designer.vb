Imports System.Windows.Forms

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ProgressUI
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ProgressUI))
        Me.grpProgressUI = New System.Windows.Forms.GroupBox()
        Me.rtbProgress = New System.Windows.Forms.RichTextBox()
        Me.AutoCloseTimer = New System.Windows.Forms.Timer(Me.components)
        Me.lblCurTaskPre = New System.Windows.Forms.Label()
        Me.lblCurTask = New System.Windows.Forms.Label()
        Me.TrayIcon = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.chkTailProgress = New System.Windows.Forms.CheckBox()
        Me.grpProgressUI.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpProgressUI
        '
        Me.grpProgressUI.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpProgressUI.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.grpProgressUI.Controls.Add(Me.rtbProgress)
        Me.grpProgressUI.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpProgressUI.Location = New System.Drawing.Point(5, 2)
        Me.grpProgressUI.Name = "grpProgressUI"
        Me.grpProgressUI.Size = New System.Drawing.Size(505, 506)
        Me.grpProgressUI.TabIndex = 4
        Me.grpProgressUI.TabStop = False
        Me.grpProgressUI.Text = "Status Console"
        '
        'rtbProgress
        '
        Me.rtbProgress.BackColor = System.Drawing.Color.Beige
        Me.rtbProgress.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.rtbProgress.DetectUrls = False
        Me.rtbProgress.Dock = System.Windows.Forms.DockStyle.Fill
        Me.rtbProgress.Font = New System.Drawing.Font("Consolas", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rtbProgress.Location = New System.Drawing.Point(3, 19)
        Me.rtbProgress.Name = "rtbProgress"
        Me.rtbProgress.ReadOnly = True
        Me.rtbProgress.Size = New System.Drawing.Size(499, 484)
        Me.rtbProgress.TabIndex = 0
        Me.rtbProgress.Text = ""
        Me.rtbProgress.WordWrap = False
        '
        'AutoCloseTimer
        '
        Me.AutoCloseTimer.Interval = 1000
        '
        'lblCurTaskPre
        '
        Me.lblCurTaskPre.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblCurTaskPre.AutoSize = True
        Me.lblCurTaskPre.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCurTaskPre.Location = New System.Drawing.Point(2, 515)
        Me.lblCurTaskPre.MaximumSize = New System.Drawing.Size(90, 18)
        Me.lblCurTaskPre.MinimumSize = New System.Drawing.Size(90, 18)
        Me.lblCurTaskPre.Name = "lblCurTaskPre"
        Me.lblCurTaskPre.Size = New System.Drawing.Size(90, 18)
        Me.lblCurTaskPre.TabIndex = 5
        Me.lblCurTaskPre.Text = "Current Task:"
        '
        'lblCurTask
        '
        Me.lblCurTask.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblCurTask.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCurTask.Location = New System.Drawing.Point(98, 515)
        Me.lblCurTask.Name = "lblCurTask"
        Me.lblCurTask.Size = New System.Drawing.Size(294, 18)
        Me.lblCurTask.TabIndex = 6
        '
        'TrayIcon
        '
        Me.TrayIcon.Icon = CType(resources.GetObject("TrayIcon.Icon"), System.Drawing.Icon)
        '
        'chkTailProgress
        '
        Me.chkTailProgress.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkTailProgress.AutoSize = True
        Me.chkTailProgress.Checked = True
        Me.chkTailProgress.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkTailProgress.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkTailProgress.Location = New System.Drawing.Point(398, 515)
        Me.chkTailProgress.MaximumSize = New System.Drawing.Size(104, 20)
        Me.chkTailProgress.MinimumSize = New System.Drawing.Size(104, 20)
        Me.chkTailProgress.Name = "chkTailProgress"
        Me.chkTailProgress.Size = New System.Drawing.Size(104, 20)
        Me.chkTailProgress.TabIndex = 7
        Me.chkTailProgress.Text = "Follow Output"
        Me.chkTailProgress.UseVisualStyleBackColor = True
        '
        'ProgressUI
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(514, 542)
        Me.Controls.Add(Me.grpProgressUI)
        Me.Controls.Add(Me.chkTailProgress)
        Me.Controls.Add(Me.lblCurTask)
        Me.Controls.Add(Me.lblCurTaskPre)
        Me.Font = New System.Drawing.Font("Calibri", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimumSize = New System.Drawing.Size(530, 250)
        Me.Name = "ProgressUI"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "WinOffline"
        Me.grpProgressUI.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
    Friend WithEvents grpProgressUI As System.Windows.Forms.GroupBox
    Friend WithEvents AutoCloseTimer As System.Windows.Forms.Timer
    Friend WithEvents rtbProgress As System.Windows.Forms.RichTextBox
    Friend WithEvents lblCurTaskPre As System.Windows.Forms.Label
    Friend WithEvents lblCurTask As System.Windows.Forms.Label
    Friend WithEvents TrayIcon As System.Windows.Forms.NotifyIcon
    Friend WithEvents chkTailProgress As System.Windows.Forms.CheckBox
End Class
