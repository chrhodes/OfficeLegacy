<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TaskPane_ExcelUtil
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
        Me.components = New System.ComponentModel.Container()
        Me.gbLastRowColumn = New System.Windows.Forms.GroupBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtLastColSearch = New System.Windows.Forms.TextBox()
        Me.txtLastRowSearch = New System.Windows.Forms.TextBox()
        Me.txtLastColSpecial = New System.Windows.Forms.TextBox()
        Me.txtLastRowSpecial = New System.Windows.Forms.TextBox()
        Me.btnGetLastRowColInfo = New System.Windows.Forms.Button()
        Me.gbDeleteDuplicateRows = New System.Windows.Forms.GroupBox()
        Me.btnDeleteDuplicateRows = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.gbFolderMap = New System.Windows.Forms.GroupBox()
        Me.btnGroupDown = New System.Windows.Forms.Button()
        Me.btnSearchDown = New System.Windows.Forms.Button()
        Me.btnSearchUp = New System.Windows.Forms.Button()
        Me.btnUnGroupSelection = New System.Windows.Forms.Button()
        Me.btnCreateFolderMap = New System.Windows.Forms.Button()
        Me.gbLastRowColumn.SuspendLayout()
        Me.gbDeleteDuplicateRows.SuspendLayout()
        Me.gbFolderMap.SuspendLayout()
        Me.SuspendLayout()
        '
        'gbLastRowColumn
        '
        Me.gbLastRowColumn.Controls.Add(Me.Label4)
        Me.gbLastRowColumn.Controls.Add(Me.Label3)
        Me.gbLastRowColumn.Controls.Add(Me.Label2)
        Me.gbLastRowColumn.Controls.Add(Me.Label1)
        Me.gbLastRowColumn.Controls.Add(Me.txtLastColSearch)
        Me.gbLastRowColumn.Controls.Add(Me.txtLastRowSearch)
        Me.gbLastRowColumn.Controls.Add(Me.txtLastColSpecial)
        Me.gbLastRowColumn.Controls.Add(Me.txtLastRowSpecial)
        Me.gbLastRowColumn.Controls.Add(Me.btnGetLastRowColInfo)
        Me.gbLastRowColumn.Location = New System.Drawing.Point(17, 299)
        Me.gbLastRowColumn.Name = "gbLastRowColumn"
        Me.gbLastRowColumn.Size = New System.Drawing.Size(167, 154)
        Me.gbLastRowColumn.TabIndex = 16
        Me.gbLastRowColumn.TabStop = False
        Me.gbLastRowColumn.Text = "Last Row / Column"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 129)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(82, 13)
        Me.Label4.TabIndex = 27
        Me.Label4.Text = "Last Col Search"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(6, 106)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(89, 13)
        Me.Label3.TabIndex = 26
        Me.Label3.Text = "Last Row Search"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 83)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(83, 13)
        Me.Label2.TabIndex = 25
        Me.Label2.Text = "Last Col Special"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 60)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(90, 13)
        Me.Label1.TabIndex = 24
        Me.Label1.Text = "Last Row Special"
        '
        'txtLastColSearch
        '
        Me.txtLastColSearch.Location = New System.Drawing.Point(105, 126)
        Me.txtLastColSearch.Name = "txtLastColSearch"
        Me.txtLastColSearch.Size = New System.Drawing.Size(56, 20)
        Me.txtLastColSearch.TabIndex = 23
        '
        'txtLastRowSearch
        '
        Me.txtLastRowSearch.Location = New System.Drawing.Point(105, 103)
        Me.txtLastRowSearch.Name = "txtLastRowSearch"
        Me.txtLastRowSearch.Size = New System.Drawing.Size(56, 20)
        Me.txtLastRowSearch.TabIndex = 22
        '
        'txtLastColSpecial
        '
        Me.txtLastColSpecial.Location = New System.Drawing.Point(105, 80)
        Me.txtLastColSpecial.Name = "txtLastColSpecial"
        Me.txtLastColSpecial.Size = New System.Drawing.Size(56, 20)
        Me.txtLastColSpecial.TabIndex = 21
        '
        'txtLastRowSpecial
        '
        Me.txtLastRowSpecial.Location = New System.Drawing.Point(105, 57)
        Me.txtLastRowSpecial.Name = "txtLastRowSpecial"
        Me.txtLastRowSpecial.Size = New System.Drawing.Size(56, 20)
        Me.txtLastRowSpecial.TabIndex = 20
        '
        'btnGetLastRowColInfo
        '
        Me.btnGetLastRowColInfo.Location = New System.Drawing.Point(6, 19)
        Me.btnGetLastRowColInfo.Name = "btnGetLastRowColInfo"
        Me.btnGetLastRowColInfo.Size = New System.Drawing.Size(155, 23)
        Me.btnGetLastRowColInfo.TabIndex = 19
        Me.btnGetLastRowColInfo.Text = "Get Last Row / Column Info"
        Me.btnGetLastRowColInfo.UseVisualStyleBackColor = True
        '
        'gbDeleteDuplicateRows
        '
        Me.gbDeleteDuplicateRows.Controls.Add(Me.btnDeleteDuplicateRows)
        Me.gbDeleteDuplicateRows.Location = New System.Drawing.Point(17, 169)
        Me.gbDeleteDuplicateRows.Name = "gbDeleteDuplicateRows"
        Me.gbDeleteDuplicateRows.Size = New System.Drawing.Size(167, 100)
        Me.gbDeleteDuplicateRows.TabIndex = 28
        Me.gbDeleteDuplicateRows.TabStop = False
        Me.gbDeleteDuplicateRows.Text = "Delete Duplicate Rows"
        '
        'btnDeleteDuplicateRows
        '
        Me.btnDeleteDuplicateRows.Location = New System.Drawing.Point(6, 19)
        Me.btnDeleteDuplicateRows.Name = "btnDeleteDuplicateRows"
        Me.btnDeleteDuplicateRows.Size = New System.Drawing.Size(155, 23)
        Me.btnDeleteDuplicateRows.TabIndex = 29
        Me.btnDeleteDuplicateRows.Text = "Delete Duplicate Rows"
        Me.ToolTip1.SetToolTip(Me.btnDeleteDuplicateRows, "Delete Duplicate Rows")
        Me.btnDeleteDuplicateRows.UseVisualStyleBackColor = True
        '
        'gbFolderMap
        '
        Me.gbFolderMap.Controls.Add(Me.btnGroupDown)
        Me.gbFolderMap.Controls.Add(Me.btnSearchDown)
        Me.gbFolderMap.Controls.Add(Me.btnSearchUp)
        Me.gbFolderMap.Controls.Add(Me.btnUnGroupSelection)
        Me.gbFolderMap.Controls.Add(Me.btnCreateFolderMap)
        Me.gbFolderMap.Location = New System.Drawing.Point(17, 48)
        Me.gbFolderMap.Name = "gbFolderMap"
        Me.gbFolderMap.Size = New System.Drawing.Size(167, 115)
        Me.gbFolderMap.TabIndex = 30
        Me.gbFolderMap.TabStop = False
        Me.gbFolderMap.Text = "Folder Map"
        '
        'btnGroupDown
        '
        Me.btnGroupDown.Image = Global.SupportToolsExcel.My.Resources.Resources.group_down
        Me.btnGroupDown.Location = New System.Drawing.Point(6, 66)
        Me.btnGroupDown.Name = "btnGroupDown"
        Me.btnGroupDown.Size = New System.Drawing.Size(35, 43)
        Me.btnGroupDown.TabIndex = 7
        Me.btnGroupDown.UseVisualStyleBackColor = True
        '
        'btnSearchDown
        '
        Me.btnSearchDown.Image = Global.SupportToolsExcel.My.Resources.Resources.search_down
        Me.btnSearchDown.Location = New System.Drawing.Point(46, 66)
        Me.btnSearchDown.Name = "btnSearchDown"
        Me.btnSearchDown.Size = New System.Drawing.Size(35, 43)
        Me.btnSearchDown.TabIndex = 6
        Me.btnSearchDown.UseVisualStyleBackColor = True
        '
        'btnSearchUp
        '
        Me.btnSearchUp.Image = Global.SupportToolsExcel.My.Resources.Resources.search_up
        Me.btnSearchUp.Location = New System.Drawing.Point(86, 66)
        Me.btnSearchUp.Name = "btnSearchUp"
        Me.btnSearchUp.Size = New System.Drawing.Size(35, 43)
        Me.btnSearchUp.TabIndex = 5
        Me.btnSearchUp.UseVisualStyleBackColor = True
        '
        'btnUnGroupSelection
        '
        Me.btnUnGroupSelection.Image = Global.SupportToolsExcel.My.Resources.Resources.ungroup_selection
        Me.btnUnGroupSelection.Location = New System.Drawing.Point(126, 66)
        Me.btnUnGroupSelection.Name = "btnUnGroupSelection"
        Me.btnUnGroupSelection.Size = New System.Drawing.Size(35, 43)
        Me.btnUnGroupSelection.TabIndex = 4
        Me.btnUnGroupSelection.UseVisualStyleBackColor = True
        '
        'btnCreateFolderMap
        '
        Me.btnCreateFolderMap.Image = Global.SupportToolsExcel.My.Resources.Resources.folder_map
        Me.btnCreateFolderMap.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnCreateFolderMap.Location = New System.Drawing.Point(6, 19)
        Me.btnCreateFolderMap.Name = "btnCreateFolderMap"
        Me.btnCreateFolderMap.Size = New System.Drawing.Size(155, 41)
        Me.btnCreateFolderMap.TabIndex = 0
        Me.btnCreateFolderMap.Text = "Create Folder Map"
        Me.btnCreateFolderMap.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnCreateFolderMap.UseVisualStyleBackColor = True
        '
        'TaskPane_ExcelUtil
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.gbFolderMap)
        Me.Controls.Add(Me.gbDeleteDuplicateRows)
        Me.Controls.Add(Me.gbLastRowColumn)
        Me.Name = "TaskPane_ExcelUtil"
        Me.Size = New System.Drawing.Size(200, 468)
        Me.gbLastRowColumn.ResumeLayout(False)
        Me.gbLastRowColumn.PerformLayout()
        Me.gbDeleteDuplicateRows.ResumeLayout(False)
        Me.gbFolderMap.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents gbLastRowColumn As System.Windows.Forms.GroupBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtLastColSearch As System.Windows.Forms.TextBox
    Friend WithEvents txtLastRowSearch As System.Windows.Forms.TextBox
    Friend WithEvents txtLastColSpecial As System.Windows.Forms.TextBox
    Friend WithEvents txtLastRowSpecial As System.Windows.Forms.TextBox
    Friend WithEvents btnGetLastRowColInfo As System.Windows.Forms.Button
    Friend WithEvents gbDeleteDuplicateRows As System.Windows.Forms.GroupBox
    Friend WithEvents btnDeleteDuplicateRows As System.Windows.Forms.Button
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents gbFolderMap As System.Windows.Forms.GroupBox
    Friend WithEvents btnUnGroupSelection As System.Windows.Forms.Button
    Friend WithEvents btnCreateFolderMap As System.Windows.Forms.Button
    Friend WithEvents btnGroupDown As System.Windows.Forms.Button
    Friend WithEvents btnSearchDown As System.Windows.Forms.Button
    Friend WithEvents btnSearchUp As System.Windows.Forms.Button

End Class
