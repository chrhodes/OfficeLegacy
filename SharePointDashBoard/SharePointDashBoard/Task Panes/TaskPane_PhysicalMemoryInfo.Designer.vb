<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TaskPane_PhysicalMemoryInfo
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
        Me.clbHosts = New System.Windows.Forms.CheckedListBox
        Me.btnGetInfo = New System.Windows.Forms.Button
        Me.btnGetAllInfo = New System.Windows.Forms.Button
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.btnGetAllInfo)
        Me.GroupBox2.Controls.Add(Me.clbHosts)
        Me.GroupBox2.Controls.Add(Me.btnGetInfo)
        Me.GroupBox2.Location = New System.Drawing.Point(3, 11)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(337, 403)
        Me.GroupBox2.TabIndex = 10
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Group Sheet"
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
        Me.btnGetInfo.Location = New System.Drawing.Point(6, 131)
        Me.btnGetInfo.Name = "btnGetInfo"
        Me.btnGetInfo.Size = New System.Drawing.Size(325, 23)
        Me.btnGetInfo.TabIndex = 2
        Me.btnGetInfo.Text = "Get Info"
        Me.btnGetInfo.UseVisualStyleBackColor = True
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
        'TaskPane_PhysicalMemoryInfo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.GroupBox2)
        Me.Name = "TaskPane_PhysicalMemoryInfo"
        Me.Size = New System.Drawing.Size(343, 500)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents btnGetInfo As System.Windows.Forms.Button
    Friend WithEvents clbHosts As System.Windows.Forms.CheckedListBox
    Friend WithEvents btnGetAllInfo As System.Windows.Forms.Button

End Class
