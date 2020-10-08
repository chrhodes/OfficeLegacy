<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TaskPane_CreateSheets
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
        Me.btnCreateWorksheet_X = New System.Windows.Forms.Button
        Me.btnCreateWorksheet_Y = New System.Windows.Forms.Button
        Me.btnCreateWorksheet_Z = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnCreateWorksheet_X
        '
        Me.btnCreateWorksheet_X.Location = New System.Drawing.Point(12, 21)
        Me.btnCreateWorksheet_X.Name = "btnCreateWorksheet_X"
        Me.btnCreateWorksheet_X.Size = New System.Drawing.Size(155, 25)
        Me.btnCreateWorksheet_X.TabIndex = 12
        Me.btnCreateWorksheet_X.Text = "X Worksheet"
        Me.btnCreateWorksheet_X.UseVisualStyleBackColor = True
        '
        'btnCreateWorksheet_Y
        '
        Me.btnCreateWorksheet_Y.Location = New System.Drawing.Point(12, 54)
        Me.btnCreateWorksheet_Y.Name = "btnCreateWorksheet_Y"
        Me.btnCreateWorksheet_Y.Size = New System.Drawing.Size(155, 25)
        Me.btnCreateWorksheet_Y.TabIndex = 13
        Me.btnCreateWorksheet_Y.Text = "Y Worksheet"
        Me.btnCreateWorksheet_Y.UseVisualStyleBackColor = True
        '
        'btnCreateWorksheet_Z
        '
        Me.btnCreateWorksheet_Z.Location = New System.Drawing.Point(12, 87)
        Me.btnCreateWorksheet_Z.Name = "btnCreateWorksheet_Z"
        Me.btnCreateWorksheet_Z.Size = New System.Drawing.Size(155, 25)
        Me.btnCreateWorksheet_Z.TabIndex = 14
        Me.btnCreateWorksheet_Z.Text = "Z Worksheet"
        Me.btnCreateWorksheet_Z.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btnCreateWorksheet_Z)
        Me.GroupBox1.Controls.Add(Me.btnCreateWorksheet_X)
        Me.GroupBox1.Controls.Add(Me.btnCreateWorksheet_Y)
        Me.GroupBox1.Location = New System.Drawing.Point(10, 14)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(179, 260)
        Me.GroupBox1.TabIndex = 25
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Create Worksheets"
        '
        'TaskPane_CreateSheets
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "TaskPane_CreateSheets"
        Me.Size = New System.Drawing.Size(200, 290)
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnCreateWorksheet_X As System.Windows.Forms.Button
    Friend WithEvents btnCreateWorksheet_Y As System.Windows.Forms.Button
    Friend WithEvents btnCreateWorksheet_Z As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox

End Class
