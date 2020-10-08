<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TaskPane_HRB
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
        Me.btnAddOfficeInfo = New System.Windows.Forms.Button
        Me.btnLoadLookupFile = New System.Windows.Forms.Button
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.btnRemoveRows = New System.Windows.Forms.Button
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnAddOfficeInfo
        '
        Me.btnAddOfficeInfo.Location = New System.Drawing.Point(8, 48)
        Me.btnAddOfficeInfo.Name = "btnAddOfficeInfo"
        Me.btnAddOfficeInfo.Size = New System.Drawing.Size(162, 23)
        Me.btnAddOfficeInfo.TabIndex = 3
        Me.btnAddOfficeInfo.Text = "Add Office Info"
        Me.ToolTip1.SetToolTip(Me.btnAddOfficeInfo, "Add Office, District, Region, Division Information")
        Me.btnAddOfficeInfo.UseVisualStyleBackColor = True
        '
        'btnLoadLookupFile
        '
        Me.btnLoadLookupFile.Location = New System.Drawing.Point(8, 19)
        Me.btnLoadLookupFile.Name = "btnLoadLookupFile"
        Me.btnLoadLookupFile.Size = New System.Drawing.Size(162, 23)
        Me.btnLoadLookupFile.TabIndex = 6
        Me.btnLoadLookupFile.Text = "Load Lookup File"
        Me.ToolTip1.SetToolTip(Me.btnLoadLookupFile, "Select Office Lookup file")
        Me.btnLoadLookupFile.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.btnRemoveRows)
        Me.GroupBox2.Controls.Add(Me.GroupBox3)
        Me.GroupBox2.Controls.Add(Me.btnLoadLookupFile)
        Me.GroupBox2.Controls.Add(Me.btnAddOfficeInfo)
        Me.GroupBox2.Location = New System.Drawing.Point(8, 11)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(182, 403)
        Me.GroupBox2.TabIndex = 10
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Group Sheet"
        '
        'btnRemoveRows
        '
        Me.btnRemoveRows.Location = New System.Drawing.Point(8, 77)
        Me.btnRemoveRows.Name = "btnRemoveRows"
        Me.btnRemoveRows.Size = New System.Drawing.Size(162, 23)
        Me.btnRemoveRows.TabIndex = 16
        Me.btnRemoveRows.Text = "Remove Non-Response Rows"
        Me.btnRemoveRows.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Location = New System.Drawing.Point(6, 243)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(167, 154)
        Me.GroupBox3.TabIndex = 15
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Debug"
        '
        'TaskPane_HRB
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.GroupBox2)
        Me.Name = "TaskPane_HRB"
        Me.Size = New System.Drawing.Size(200, 428)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnAddOfficeInfo As System.Windows.Forms.Button
    Friend WithEvents btnLoadLookupFile As System.Windows.Forms.Button
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents btnRemoveRows As System.Windows.Forms.Button

End Class
