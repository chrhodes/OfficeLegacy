<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TaskPane_Config
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
        Me.gbDebug = New System.Windows.Forms.GroupBox
        Me.chkDisplayDebugMessages = New System.Windows.Forms.CheckBox
        Me.chkScreenUpdatesOff = New System.Windows.Forms.CheckBox
        Me.btnFindLast = New System.Windows.Forms.Button
        Me.btnReLoadConfigData = New System.Windows.Forms.Button
        Me.btnLoadLookups = New System.Windows.Forms.Button
        Me.btnCreateDefinedNames = New System.Windows.Forms.Button
        Me.gbTeamInfo = New System.Windows.Forms.GroupBox
        Me.cbTeamName = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.gbDebug.SuspendLayout()
        Me.gbTeamInfo.SuspendLayout()
        Me.SuspendLayout()
        '
        'gbDebug
        '
        Me.gbDebug.Controls.Add(Me.chkDisplayDebugMessages)
        Me.gbDebug.Controls.Add(Me.chkScreenUpdatesOff)
        Me.gbDebug.Location = New System.Drawing.Point(9, 178)
        Me.gbDebug.Name = "gbDebug"
        Me.gbDebug.Size = New System.Drawing.Size(178, 69)
        Me.gbDebug.TabIndex = 5
        Me.gbDebug.TabStop = False
        Me.gbDebug.Text = "Debug Settings"
        '
        'chkDisplayDebugMessages
        '
        Me.chkDisplayDebugMessages.AutoSize = True
        Me.chkDisplayDebugMessages.Enabled = False
        Me.chkDisplayDebugMessages.Location = New System.Drawing.Point(6, 42)
        Me.chkDisplayDebugMessages.Name = "chkDisplayDebugMessages"
        Me.chkDisplayDebugMessages.Size = New System.Drawing.Size(146, 17)
        Me.chkDisplayDebugMessages.TabIndex = 1
        Me.chkDisplayDebugMessages.Text = "Display Debug Messages"
        Me.chkDisplayDebugMessages.UseVisualStyleBackColor = True
        '
        'chkScreenUpdatesOff
        '
        Me.chkScreenUpdatesOff.AutoSize = True
        Me.chkScreenUpdatesOff.Checked = True
        Me.chkScreenUpdatesOff.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkScreenUpdatesOff.Location = New System.Drawing.Point(6, 19)
        Me.chkScreenUpdatesOff.Name = "chkScreenUpdatesOff"
        Me.chkScreenUpdatesOff.Size = New System.Drawing.Size(126, 17)
        Me.chkScreenUpdatesOff.TabIndex = 0
        Me.chkScreenUpdatesOff.Text = "Screeen Updates Off"
        Me.chkScreenUpdatesOff.UseVisualStyleBackColor = True
        '
        'btnFindLast
        '
        Me.btnFindLast.Location = New System.Drawing.Point(9, 253)
        Me.btnFindLast.Name = "btnFindLast"
        Me.btnFindLast.Size = New System.Drawing.Size(76, 20)
        Me.btnFindLast.TabIndex = 7
        Me.btnFindLast.Text = "Find Last"
        Me.btnFindLast.UseVisualStyleBackColor = True
        '
        'btnReLoadConfigData
        '
        Me.btnReLoadConfigData.Location = New System.Drawing.Point(9, 72)
        Me.btnReLoadConfigData.Name = "btnReLoadConfigData"
        Me.btnReLoadConfigData.Size = New System.Drawing.Size(178, 25)
        Me.btnReLoadConfigData.TabIndex = 8
        Me.btnReLoadConfigData.Text = "Reload Config Data"
        Me.btnReLoadConfigData.UseVisualStyleBackColor = True
        '
        'btnLoadLookups
        '
        Me.btnLoadLookups.Location = New System.Drawing.Point(9, 103)
        Me.btnLoadLookups.Name = "btnLoadLookups"
        Me.btnLoadLookups.Size = New System.Drawing.Size(178, 25)
        Me.btnLoadLookups.TabIndex = 11
        Me.btnLoadLookups.Text = "Load Lookups"
        Me.btnLoadLookups.UseVisualStyleBackColor = True
        '
        'btnCreateDefinedNames
        '
        Me.btnCreateDefinedNames.Location = New System.Drawing.Point(9, 134)
        Me.btnCreateDefinedNames.Name = "btnCreateDefinedNames"
        Me.btnCreateDefinedNames.Size = New System.Drawing.Size(178, 25)
        Me.btnCreateDefinedNames.TabIndex = 12
        Me.btnCreateDefinedNames.Text = "Create Defined Names"
        Me.btnCreateDefinedNames.UseVisualStyleBackColor = True
        '
        'gbTeamInfo
        '
        Me.gbTeamInfo.Controls.Add(Me.Label1)
        Me.gbTeamInfo.Controls.Add(Me.cbTeamName)
        Me.gbTeamInfo.Location = New System.Drawing.Point(9, 0)
        Me.gbTeamInfo.Name = "gbTeamInfo"
        Me.gbTeamInfo.Size = New System.Drawing.Size(178, 97)
        Me.gbTeamInfo.TabIndex = 13
        Me.gbTeamInfo.TabStop = False
        Me.gbTeamInfo.Text = "Team Information"
        '
        'cbTeamName
        '
        Me.cbTeamName.FormattingEnabled = True
        Me.cbTeamName.Location = New System.Drawing.Point(6, 45)
        Me.cbTeamName.Name = "cbTeamName"
        Me.cbTeamName.Size = New System.Drawing.Size(166, 21)
        Me.cbTeamName.TabIndex = 14
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(65, 13)
        Me.Label1.TabIndex = 15
        Me.Label1.Text = "Team Name"
        '
        'TaskPane_Config
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.gbTeamInfo)
        Me.Controls.Add(Me.btnCreateDefinedNames)
        Me.Controls.Add(Me.btnLoadLookups)
        Me.Controls.Add(Me.btnReLoadConfigData)
        Me.Controls.Add(Me.btnFindLast)
        Me.Controls.Add(Me.gbDebug)
        Me.Name = "TaskPane_Config"
        Me.Size = New System.Drawing.Size(196, 294)
        Me.gbDebug.ResumeLayout(False)
        Me.gbDebug.PerformLayout()
        Me.gbTeamInfo.ResumeLayout(False)
        Me.gbTeamInfo.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents gbDebug As System.Windows.Forms.GroupBox
    Friend WithEvents chkDisplayDebugMessages As System.Windows.Forms.CheckBox
    Friend WithEvents chkScreenUpdatesOff As System.Windows.Forms.CheckBox
    Friend WithEvents btnFindLast As System.Windows.Forms.Button
    Friend WithEvents btnReLoadConfigData As System.Windows.Forms.Button
    Friend WithEvents btnLoadLookups As System.Windows.Forms.Button
    Friend WithEvents btnCreateDefinedNames As System.Windows.Forms.Button
    Friend WithEvents gbTeamInfo As System.Windows.Forms.GroupBox
    Friend WithEvents cbTeamName As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label

End Class
