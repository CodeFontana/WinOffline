<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class HelpUI
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(HelpUI))
        Me.rtbHelp = New System.Windows.Forms.RichTextBox()
        Me.btnReturn = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'rtbHelp
        '
        Me.rtbHelp.BackColor = System.Drawing.Color.Beige
        Me.rtbHelp.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rtbHelp.Location = New System.Drawing.Point(12, 12)
        Me.rtbHelp.Name = "rtbHelp"
        Me.rtbHelp.ReadOnly = True
        Me.rtbHelp.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical
        Me.rtbHelp.Size = New System.Drawing.Size(410, 502)
        Me.rtbHelp.TabIndex = 0
        Me.rtbHelp.TabStop = False
        Me.rtbHelp.Text = resources.GetString("rtbHelp.Text")
        '
        'btnReturn
        '
        Me.btnReturn.BackColor = System.Drawing.Color.SteelBlue
        Me.btnReturn.FlatAppearance.BorderSize = 0
        Me.btnReturn.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnReturn.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnReturn.ForeColor = System.Drawing.Color.White
        Me.btnReturn.Location = New System.Drawing.Point(156, 520)
        Me.btnReturn.Name = "btnReturn"
        Me.btnReturn.Size = New System.Drawing.Size(123, 30)
        Me.btnReturn.TabIndex = 1
        Me.btnReturn.Text = "&Return.."
        Me.btnReturn.UseVisualStyleBackColor = False
        '
        'HelpUI
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.ClientSize = New System.Drawing.Size(434, 562)
        Me.Controls.Add(Me.btnReturn)
        Me.Controls.Add(Me.rtbHelp)
        Me.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "HelpUI"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Command Line Options"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents rtbHelp As System.Windows.Forms.RichTextBox
    Friend WithEvents btnReturn As System.Windows.Forms.Button
End Class
