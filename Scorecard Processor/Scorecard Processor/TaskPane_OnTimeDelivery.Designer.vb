<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TaskPane_OnTimeDelivery
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
        Me.btnAddOnTimeDataSheets = New System.Windows.Forms.Button
        Me.btnAddOnTimeDataToPowerPoint = New System.Windows.Forms.Button
        Me.btnBrowseForOnTimeDataFile = New System.Windows.Forms.Button
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.btnValidateOnTimeDataFiles = New System.Windows.Forms.Button
        Me.btnOpenOnTimeDataFile = New System.Windows.Forms.Button
        Me.btnAddOnTimeChart = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.cbOnTimeTeams = New System.Windows.Forms.ComboBox
        Me.clbOnTimeTeams = New System.Windows.Forms.CheckedListBox
        Me.btnAddOnTimeCharts = New System.Windows.Forms.Button
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.btnDeleteOnTimeDataSheets = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnAddOnTimeDataSheets
        '
        Me.btnAddOnTimeDataSheets.Location = New System.Drawing.Point(8, 48)
        Me.btnAddOnTimeDataSheets.Name = "btnAddOnTimeDataSheets"
        Me.btnAddOnTimeDataSheets.Size = New System.Drawing.Size(162, 23)
        Me.btnAddOnTimeDataSheets.TabIndex = 3
        Me.btnAddOnTimeDataSheets.Text = "Add On-Time Data Sheets"
        Me.ToolTip1.SetToolTip(Me.btnAddOnTimeDataSheets, "Add On-Time Data Worksheet for selected file(s)")
        Me.btnAddOnTimeDataSheets.UseVisualStyleBackColor = True
        '
        'btnAddOnTimeDataToPowerPoint
        '
        Me.btnAddOnTimeDataToPowerPoint.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAddOnTimeDataToPowerPoint.Location = New System.Drawing.Point(8, 106)
        Me.btnAddOnTimeDataToPowerPoint.Name = "btnAddOnTimeDataToPowerPoint"
        Me.btnAddOnTimeDataToPowerPoint.Size = New System.Drawing.Size(162, 38)
        Me.btnAddOnTimeDataToPowerPoint.TabIndex = 5
        Me.btnAddOnTimeDataToPowerPoint.Text = "Add On-Time Data to PowerPoint"
        Me.btnAddOnTimeDataToPowerPoint.UseVisualStyleBackColor = True
        '
        'btnBrowseForOnTimeDataFile
        '
        Me.btnBrowseForOnTimeDataFile.Location = New System.Drawing.Point(8, 19)
        Me.btnBrowseForOnTimeDataFile.Name = "btnBrowseForOnTimeDataFile"
        Me.btnBrowseForOnTimeDataFile.Size = New System.Drawing.Size(162, 23)
        Me.btnBrowseForOnTimeDataFile.TabIndex = 6
        Me.btnBrowseForOnTimeDataFile.Text = "Browse for On-Time Data File"
        Me.ToolTip1.SetToolTip(Me.btnBrowseForOnTimeDataFile, "Select On-Time data file and place in current cell")
        Me.btnBrowseForOnTimeDataFile.UseVisualStyleBackColor = True
        '
        'btnValidateOnTimeDataFiles
        '
        Me.btnValidateOnTimeDataFiles.Location = New System.Drawing.Point(7, 64)
        Me.btnValidateOnTimeDataFiles.Name = "btnValidateOnTimeDataFiles"
        Me.btnValidateOnTimeDataFiles.Size = New System.Drawing.Size(153, 23)
        Me.btnValidateOnTimeDataFiles.TabIndex = 12
        Me.btnValidateOnTimeDataFiles.Text = "Validate On-Time Data Files"
        Me.ToolTip1.SetToolTip(Me.btnValidateOnTimeDataFiles, "Validate selected file(s) contain valid On-Time data")
        Me.btnValidateOnTimeDataFiles.UseVisualStyleBackColor = True
        '
        'btnOpenOnTimeDataFile
        '
        Me.btnOpenOnTimeDataFile.Location = New System.Drawing.Point(7, 93)
        Me.btnOpenOnTimeDataFile.Name = "btnOpenOnTimeDataFile"
        Me.btnOpenOnTimeDataFile.Size = New System.Drawing.Size(153, 23)
        Me.btnOpenOnTimeDataFile.TabIndex = 13
        Me.btnOpenOnTimeDataFile.Text = "Open On-Time Data File"
        Me.ToolTip1.SetToolTip(Me.btnOpenOnTimeDataFile, "Open selected file")
        Me.btnOpenOnTimeDataFile.UseVisualStyleBackColor = True
        '
        'btnAddOnTimeChart
        '
        Me.btnAddOnTimeChart.Location = New System.Drawing.Point(8, 19)
        Me.btnAddOnTimeChart.Name = "btnAddOnTimeChart"
        Me.btnAddOnTimeChart.Size = New System.Drawing.Size(162, 23)
        Me.btnAddOnTimeChart.TabIndex = 7
        Me.btnAddOnTimeChart.Text = "Add On-Time Chart"
        Me.btnAddOnTimeChart.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btnAddOnTimeChart)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 420)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(182, 58)
        Me.GroupBox1.TabIndex = 9
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Team Sheet"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.cbOnTimeTeams)
        Me.GroupBox2.Controls.Add(Me.clbOnTimeTeams)
        Me.GroupBox2.Controls.Add(Me.btnAddOnTimeCharts)
        Me.GroupBox2.Controls.Add(Me.GroupBox3)
        Me.GroupBox2.Controls.Add(Me.btnBrowseForOnTimeDataFile)
        Me.GroupBox2.Controls.Add(Me.btnAddOnTimeDataToPowerPoint)
        Me.GroupBox2.Controls.Add(Me.btnAddOnTimeDataSheets)
        Me.GroupBox2.Location = New System.Drawing.Point(8, 11)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(182, 403)
        Me.GroupBox2.TabIndex = 10
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Group Sheet"
        '
        'cbOnTimeTeams
        '
        Me.cbOnTimeTeams.FormattingEnabled = True
        Me.cbOnTimeTeams.Location = New System.Drawing.Point(6, 216)
        Me.cbOnTimeTeams.Name = "cbOnTimeTeams"
        Me.cbOnTimeTeams.Size = New System.Drawing.Size(164, 21)
        Me.cbOnTimeTeams.TabIndex = 18
        '
        'clbOnTimeTeams
        '
        Me.clbOnTimeTeams.FormattingEnabled = True
        Me.clbOnTimeTeams.Location = New System.Drawing.Point(6, 161)
        Me.clbOnTimeTeams.Name = "clbOnTimeTeams"
        Me.clbOnTimeTeams.ScrollAlwaysVisible = True
        Me.clbOnTimeTeams.Size = New System.Drawing.Size(163, 49)
        Me.clbOnTimeTeams.TabIndex = 17
        '
        'btnAddOnTimeCharts
        '
        Me.btnAddOnTimeCharts.Location = New System.Drawing.Point(8, 77)
        Me.btnAddOnTimeCharts.Name = "btnAddOnTimeCharts"
        Me.btnAddOnTimeCharts.Size = New System.Drawing.Size(162, 23)
        Me.btnAddOnTimeCharts.TabIndex = 16
        Me.btnAddOnTimeCharts.Text = "Add On-Time Charts"
        Me.btnAddOnTimeCharts.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.btnValidateOnTimeDataFiles)
        Me.GroupBox3.Controls.Add(Me.btnOpenOnTimeDataFile)
        Me.GroupBox3.Controls.Add(Me.btnDeleteOnTimeDataSheets)
        Me.GroupBox3.Location = New System.Drawing.Point(6, 243)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(167, 154)
        Me.GroupBox3.TabIndex = 15
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Debug"
        '
        'btnDeleteOnTimeDataSheets
        '
        Me.btnDeleteOnTimeDataSheets.Location = New System.Drawing.Point(7, 122)
        Me.btnDeleteOnTimeDataSheets.Name = "btnDeleteOnTimeDataSheets"
        Me.btnDeleteOnTimeDataSheets.Size = New System.Drawing.Size(153, 23)
        Me.btnDeleteOnTimeDataSheets.TabIndex = 14
        Me.btnDeleteOnTimeDataSheets.Text = "Delete On-Time Data Sheets"
        Me.btnDeleteOnTimeDataSheets.UseVisualStyleBackColor = True
        '
        'TaskPane_OnTimeDelivery
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.GroupBox2)
        Me.Name = "TaskPane_OnTimeDelivery"
        Me.Size = New System.Drawing.Size(200, 500)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnAddOnTimeDataSheets As System.Windows.Forms.Button
    Friend WithEvents btnAddOnTimeDataToPowerPoint As System.Windows.Forms.Button
    Friend WithEvents btnBrowseForOnTimeDataFile As System.Windows.Forms.Button
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents btnAddOnTimeChart As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents btnValidateOnTimeDataFiles As System.Windows.Forms.Button
    Friend WithEvents btnOpenOnTimeDataFile As System.Windows.Forms.Button
    Friend WithEvents btnDeleteOnTimeDataSheets As System.Windows.Forms.Button
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents btnAddOnTimeCharts As System.Windows.Forms.Button
    Friend WithEvents clbOnTimeTeams As System.Windows.Forms.CheckedListBox
    Friend WithEvents cbOnTimeTeams As System.Windows.Forms.ComboBox

End Class
