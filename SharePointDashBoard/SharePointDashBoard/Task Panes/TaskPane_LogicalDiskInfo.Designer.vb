<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TaskPane_LogicalDiskInfo
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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
        Me.components = New System.ComponentModel.Container
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.btnGetAllInfo = New System.Windows.Forms.Button
        Me.clbHosts = New System.Windows.Forms.CheckedListBox
        Me.btnGetInfo = New System.Windows.Forms.Button
        Me.cbFloppyDisk = New System.Windows.Forms.CheckBox
        Me.cbCDROM = New System.Windows.Forms.CheckBox
        Me.cbHardDisk = New System.Windows.Forms.CheckBox
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.cbHardDisk)
        Me.GroupBox2.Controls.Add(Me.cbCDROM)
        Me.GroupBox2.Controls.Add(Me.cbFloppyDisk)
        Me.GroupBox2.Controls.Add(Me.btnGetAllInfo)
        Me.GroupBox2.Controls.Add(Me.clbHosts)
        Me.GroupBox2.Controls.Add(Me.btnGetInfo)
        Me.GroupBox2.Location = New System.Drawing.Point(3, 11)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(337, 403)
        Me.GroupBox2.TabIndex = 10
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "LogicalDisk Info"
        '
        'btnGetAllInfo
        '
        Me.btnGetAllInfo.Location = New System.Drawing.Point(6, 190)
        Me.btnGetAllInfo.Name = "btnGetAllInfo"
        Me.btnGetAllInfo.Size = New System.Drawing.Size(325, 23)
        Me.btnGetAllInfo.TabIndex = 4
        Me.btnGetAllInfo.Text = "Get All Info"
        Me.btnGetAllInfo.UseVisualStyleBackColor = True
        '
        'clbHosts
        '
        Me.clbHosts.FormattingEnabled = True
        Me.clbHosts.Location = New System.Drawing.Point(6, 19)
        Me.clbHosts.Name = "clbHosts"
        Me.clbHosts.Size = New System.Drawing.Size(325, 94)
        Me.clbHosts.TabIndex = 3
        '
        'btnGetInfo
        '
        Me.btnGetInfo.Location = New System.Drawing.Point(6, 161)
        Me.btnGetInfo.Name = "btnGetInfo"
        Me.btnGetInfo.Size = New System.Drawing.Size(325, 23)
        Me.btnGetInfo.TabIndex = 2
        Me.btnGetInfo.Text = "Get Info"
        Me.btnGetInfo.UseVisualStyleBackColor = True
        '
        'cbFloppyDisk
        '
        Me.cbFloppyDisk.AutoSize = True
        Me.cbFloppyDisk.Location = New System.Drawing.Point(6, 119)
        Me.cbFloppyDisk.Name = "cbFloppyDisk"
        Me.cbFloppyDisk.Size = New System.Drawing.Size(86, 17)
        Me.cbFloppyDisk.TabIndex = 5
        Me.cbFloppyDisk.Text = "Floppy Disks"
        Me.cbFloppyDisk.UseVisualStyleBackColor = True
        '
        'cbCDROM
        '
        Me.cbCDROM.AutoSize = True
        Me.cbCDROM.Location = New System.Drawing.Point(126, 119)
        Me.cbCDROM.Name = "cbCDROM"
        Me.cbCDROM.Size = New System.Drawing.Size(76, 17)
        Me.cbCDROM.TabIndex = 6
        Me.cbCDROM.Text = "CD ROMS"
        Me.cbCDROM.UseVisualStyleBackColor = True
        '
        'cbHardDisk
        '
        Me.cbHardDisk.AutoSize = True
        Me.cbHardDisk.Checked = True
        Me.cbHardDisk.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbHardDisk.Location = New System.Drawing.Point(245, 119)
        Me.cbHardDisk.Name = "cbHardDisk"
        Me.cbHardDisk.Size = New System.Drawing.Size(78, 17)
        Me.cbHardDisk.TabIndex = 7
        Me.cbHardDisk.Text = "Hard Disks"
        Me.cbHardDisk.UseVisualStyleBackColor = True
        '
        'TaskPane_LogicalDiskInfo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.GroupBox2)
        Me.Name = "TaskPane_LogicalDiskInfo"
        Me.Size = New System.Drawing.Size(343, 500)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents btnGetInfo As System.Windows.Forms.Button
    Friend WithEvents clbHosts As System.Windows.Forms.CheckedListBox
    Friend WithEvents btnGetAllInfo As System.Windows.Forms.Button
    Friend WithEvents cbHardDisk As System.Windows.Forms.CheckBox
    Friend WithEvents cbCDROM As System.Windows.Forms.CheckBox
    Friend WithEvents cbFloppyDisk As System.Windows.Forms.CheckBox

End Class
