<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SqlConnectUI
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SqlConnectUI))
        Me.grpSqlLogin = New System.Windows.Forms.GroupBox()
        Me.lblDatabaseName = New System.Windows.Forms.Label()
        Me.txtDatabaseName = New System.Windows.Forms.TextBox()
        Me.txtPortNumber = New System.Windows.Forms.TextBox()
        Me.txtInstanceName = New System.Windows.Forms.TextBox()
        Me.lblPortNumber = New System.Windows.Forms.Label()
        Me.lblInstanceName = New System.Windows.Forms.Label()
        Me.lblBlackLine = New System.Windows.Forms.Label()
        Me.btnSqlConnect = New System.Windows.Forms.Button()
        Me.txtSqlPassword = New System.Windows.Forms.TextBox()
        Me.lblSqlPassword = New System.Windows.Forms.Label()
        Me.txtSqlUser = New System.Windows.Forms.TextBox()
        Me.lblSqlUser = New System.Windows.Forms.Label()
        Me.cmbAuthType = New System.Windows.Forms.ComboBox()
        Me.lblAuthType = New System.Windows.Forms.Label()
        Me.txtSqlServer = New System.Windows.Forms.TextBox()
        Me.lblSqlServer = New System.Windows.Forms.Label()
        Me.btnSqlCancel = New System.Windows.Forms.Button()
        Me.grpSqlLogin.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpSqlLogin
        '
        Me.grpSqlLogin.Controls.Add(Me.btnSqlCancel)
        Me.grpSqlLogin.Controls.Add(Me.lblDatabaseName)
        Me.grpSqlLogin.Controls.Add(Me.txtDatabaseName)
        Me.grpSqlLogin.Controls.Add(Me.txtPortNumber)
        Me.grpSqlLogin.Controls.Add(Me.txtInstanceName)
        Me.grpSqlLogin.Controls.Add(Me.lblPortNumber)
        Me.grpSqlLogin.Controls.Add(Me.lblInstanceName)
        Me.grpSqlLogin.Controls.Add(Me.lblBlackLine)
        Me.grpSqlLogin.Controls.Add(Me.btnSqlConnect)
        Me.grpSqlLogin.Controls.Add(Me.txtSqlPassword)
        Me.grpSqlLogin.Controls.Add(Me.lblSqlPassword)
        Me.grpSqlLogin.Controls.Add(Me.txtSqlUser)
        Me.grpSqlLogin.Controls.Add(Me.lblSqlUser)
        Me.grpSqlLogin.Controls.Add(Me.cmbAuthType)
        Me.grpSqlLogin.Controls.Add(Me.lblAuthType)
        Me.grpSqlLogin.Controls.Add(Me.txtSqlServer)
        Me.grpSqlLogin.Controls.Add(Me.lblSqlServer)
        Me.grpSqlLogin.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grpSqlLogin.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpSqlLogin.Location = New System.Drawing.Point(3, 3)
        Me.grpSqlLogin.Name = "grpSqlLogin"
        Me.grpSqlLogin.Size = New System.Drawing.Size(559, 300)
        Me.grpSqlLogin.TabIndex = 29
        Me.grpSqlLogin.TabStop = False
        Me.grpSqlLogin.Text = "SQL Login"
        '
        'lblDatabaseName
        '
        Me.lblDatabaseName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblDatabaseName.AutoSize = True
        Me.lblDatabaseName.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDatabaseName.Location = New System.Drawing.Point(24, 121)
        Me.lblDatabaseName.Name = "lblDatabaseName"
        Me.lblDatabaseName.Size = New System.Drawing.Size(62, 15)
        Me.lblDatabaseName.TabIndex = 48
        Me.lblDatabaseName.Text = "Database:"
        '
        'txtDatabaseName
        '
        Me.txtDatabaseName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDatabaseName.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDatabaseName.Location = New System.Drawing.Point(128, 118)
        Me.txtDatabaseName.Name = "txtDatabaseName"
        Me.txtDatabaseName.Size = New System.Drawing.Size(416, 23)
        Me.txtDatabaseName.TabIndex = 34
        '
        'txtPortNumber
        '
        Me.txtPortNumber.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPortNumber.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPortNumber.Location = New System.Drawing.Point(128, 89)
        Me.txtPortNumber.Name = "txtPortNumber"
        Me.txtPortNumber.Size = New System.Drawing.Size(416, 23)
        Me.txtPortNumber.TabIndex = 33
        '
        'txtInstanceName
        '
        Me.txtInstanceName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtInstanceName.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtInstanceName.Location = New System.Drawing.Point(128, 60)
        Me.txtInstanceName.Name = "txtInstanceName"
        Me.txtInstanceName.Size = New System.Drawing.Size(416, 23)
        Me.txtInstanceName.TabIndex = 32
        '
        'lblPortNumber
        '
        Me.lblPortNumber.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblPortNumber.AutoSize = True
        Me.lblPortNumber.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPortNumber.Location = New System.Drawing.Point(24, 92)
        Me.lblPortNumber.Name = "lblPortNumber"
        Me.lblPortNumber.Size = New System.Drawing.Size(33, 15)
        Me.lblPortNumber.TabIndex = 44
        Me.lblPortNumber.Text = "Port:"
        '
        'lblInstanceName
        '
        Me.lblInstanceName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblInstanceName.AutoSize = True
        Me.lblInstanceName.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInstanceName.Location = New System.Drawing.Point(24, 62)
        Me.lblInstanceName.Name = "lblInstanceName"
        Me.lblInstanceName.Size = New System.Drawing.Size(91, 15)
        Me.lblInstanceName.TabIndex = 43
        Me.lblInstanceName.Text = "Instance Name:"
        '
        'lblBlackLine
        '
        Me.lblBlackLine.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblBlackLine.BackColor = System.Drawing.Color.Black
        Me.lblBlackLine.Location = New System.Drawing.Point(14, 242)
        Me.lblBlackLine.Name = "lblBlackLine"
        Me.lblBlackLine.Size = New System.Drawing.Size(532, 1)
        Me.lblBlackLine.TabIndex = 40
        '
        'btnSqlConnect
        '
        Me.btnSqlConnect.BackColor = System.Drawing.Color.DarkSeaGreen
        Me.btnSqlConnect.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSqlConnect.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSqlConnect.ForeColor = System.Drawing.Color.White
        Me.btnSqlConnect.Location = New System.Drawing.Point(125, 259)
        Me.btnSqlConnect.Name = "btnSqlConnect"
        Me.btnSqlConnect.Size = New System.Drawing.Size(126, 30)
        Me.btnSqlConnect.TabIndex = 38
        Me.btnSqlConnect.Text = "&Connect Sql"
        Me.btnSqlConnect.UseVisualStyleBackColor = False
        '
        'txtSqlPassword
        '
        Me.txtSqlPassword.AcceptsReturn = True
        Me.txtSqlPassword.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSqlPassword.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSqlPassword.Location = New System.Drawing.Point(181, 205)
        Me.txtSqlPassword.Name = "txtSqlPassword"
        Me.txtSqlPassword.Size = New System.Drawing.Size(363, 23)
        Me.txtSqlPassword.TabIndex = 37
        Me.txtSqlPassword.UseSystemPasswordChar = True
        '
        'lblSqlPassword
        '
        Me.lblSqlPassword.AutoSize = True
        Me.lblSqlPassword.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSqlPassword.Location = New System.Drawing.Point(24, 207)
        Me.lblSqlPassword.Name = "lblSqlPassword"
        Me.lblSqlPassword.Size = New System.Drawing.Size(64, 15)
        Me.lblSqlPassword.TabIndex = 36
        Me.lblSqlPassword.Text = "Password:"
        '
        'txtSqlUser
        '
        Me.txtSqlUser.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSqlUser.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSqlUser.Location = New System.Drawing.Point(181, 176)
        Me.txtSqlUser.Name = "txtSqlUser"
        Me.txtSqlUser.Size = New System.Drawing.Size(363, 23)
        Me.txtSqlUser.TabIndex = 36
        '
        'lblSqlUser
        '
        Me.lblSqlUser.AutoSize = True
        Me.lblSqlUser.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSqlUser.Location = New System.Drawing.Point(24, 179)
        Me.lblSqlUser.Name = "lblSqlUser"
        Me.lblSqlUser.Size = New System.Drawing.Size(35, 15)
        Me.lblSqlUser.TabIndex = 34
        Me.lblSqlUser.Text = "User:"
        '
        'cmbAuthType
        '
        Me.cmbAuthType.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbAuthType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbAuthType.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbAuthType.FormattingEnabled = True
        Me.cmbAuthType.Location = New System.Drawing.Point(128, 147)
        Me.cmbAuthType.Name = "cmbAuthType"
        Me.cmbAuthType.Size = New System.Drawing.Size(416, 23)
        Me.cmbAuthType.TabIndex = 35
        '
        'lblAuthType
        '
        Me.lblAuthType.AutoSize = True
        Me.lblAuthType.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAuthType.Location = New System.Drawing.Point(24, 152)
        Me.lblAuthType.Name = "lblAuthType"
        Me.lblAuthType.Size = New System.Drawing.Size(89, 15)
        Me.lblAuthType.TabIndex = 32
        Me.lblAuthType.Text = "Authentication:"
        '
        'txtSqlServer
        '
        Me.txtSqlServer.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSqlServer.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSqlServer.Location = New System.Drawing.Point(128, 32)
        Me.txtSqlServer.Name = "txtSqlServer"
        Me.txtSqlServer.Size = New System.Drawing.Size(416, 23)
        Me.txtSqlServer.TabIndex = 31
        '
        'lblSqlServer
        '
        Me.lblSqlServer.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblSqlServer.AutoSize = True
        Me.lblSqlServer.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSqlServer.Location = New System.Drawing.Point(24, 34)
        Me.lblSqlServer.Name = "lblSqlServer"
        Me.lblSqlServer.Size = New System.Drawing.Size(67, 15)
        Me.lblSqlServer.TabIndex = 30
        Me.lblSqlServer.Text = "SQL Server:"
        '
        'btnSqlCancel
        '
        Me.btnSqlCancel.BackColor = System.Drawing.Color.IndianRed
        Me.btnSqlCancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSqlCancel.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSqlCancel.ForeColor = System.Drawing.Color.White
        Me.btnSqlCancel.Location = New System.Drawing.Point(307, 259)
        Me.btnSqlCancel.Name = "btnSqlCancel"
        Me.btnSqlCancel.Size = New System.Drawing.Size(126, 30)
        Me.btnSqlCancel.TabIndex = 49
        Me.btnSqlCancel.Text = "&Cancel"
        Me.btnSqlCancel.UseVisualStyleBackColor = False
        '
        'SqlConnectUI
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(565, 306)
        Me.Controls.Add(Me.grpSqlLogin)
        Me.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "SqlConnectUI"
        Me.Padding = New System.Windows.Forms.Padding(3)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "WinOffline"
        Me.grpSqlLogin.ResumeLayout(False)
        Me.grpSqlLogin.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents grpSqlLogin As Windows.Forms.GroupBox
    Friend WithEvents lblDatabaseName As Windows.Forms.Label
    Friend WithEvents txtDatabaseName As Windows.Forms.TextBox
    Friend WithEvents txtPortNumber As Windows.Forms.TextBox
    Friend WithEvents txtInstanceName As Windows.Forms.TextBox
    Friend WithEvents lblPortNumber As Windows.Forms.Label
    Friend WithEvents lblInstanceName As Windows.Forms.Label
    Friend WithEvents lblBlackLine As Windows.Forms.Label
    Friend WithEvents btnSqlConnect As Windows.Forms.Button
    Friend WithEvents txtSqlPassword As Windows.Forms.TextBox
    Friend WithEvents lblSqlPassword As Windows.Forms.Label
    Friend WithEvents txtSqlUser As Windows.Forms.TextBox
    Friend WithEvents lblSqlUser As Windows.Forms.Label
    Friend WithEvents cmbAuthType As Windows.Forms.ComboBox
    Friend WithEvents lblAuthType As Windows.Forms.Label
    Friend WithEvents txtSqlServer As Windows.Forms.TextBox
    Friend WithEvents lblSqlServer As Windows.Forms.Label
    Friend WithEvents btnSqlCancel As Windows.Forms.Button
End Class
