<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TaskPane_ITRs
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
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.btnProcessDARTOutput = New System.Windows.Forms.Button()
        Me.btnDuplicateInputSheet = New System.Windows.Forms.Button()
        Me.btnProcessDuplicateRows = New System.Windows.Forms.Button()
        Me.btnProcessDynamicOutput = New System.Windows.Forms.Button()
        Me.btnMergeDuplicateRows = New System.Windows.Forms.Button()
        Me.btnGetITRInformation = New System.Windows.Forms.Button()
        Me.btnDisplayITRDetail = New System.Windows.Forms.Button()
        Me.gbDARTReport = New System.Windows.Forms.GroupBox()
        Me.gbDebug = New System.Windows.Forms.GroupBox()
        Me.btnAddPageBreaks = New System.Windows.Forms.Button()
        Me.btnFormatSourceITRs = New System.Windows.Forms.Button()
        Me.btnAddPivotTables = New System.Windows.Forms.Button()
        Me.btnAddListObjects = New System.Windows.Forms.Button()
        Me.cmbTeamName = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.gbITRWork = New System.Windows.Forms.GroupBox()
        Me.gbDARTReport.SuspendLayout()
        Me.gbDebug.SuspendLayout()
        Me.gbITRWork.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnProcessDARTOutput
        '
        Me.btnProcessDARTOutput.Location = New System.Drawing.Point(8, 60)
        Me.btnProcessDARTOutput.Name = "btnProcessDARTOutput"
        Me.btnProcessDARTOutput.Size = New System.Drawing.Size(162, 23)
        Me.btnProcessDARTOutput.TabIndex = 21
        Me.btnProcessDARTOutput.Text = "Process DART Output"
        Me.ToolTip1.SetToolTip(Me.btnProcessDARTOutput, "Step One")
        Me.btnProcessDARTOutput.UseVisualStyleBackColor = True
        '
        'btnDuplicateInputSheet
        '
        Me.btnDuplicateInputSheet.Location = New System.Drawing.Point(10, 42)
        Me.btnDuplicateInputSheet.Name = "btnDuplicateInputSheet"
        Me.btnDuplicateInputSheet.Size = New System.Drawing.Size(142, 23)
        Me.btnDuplicateInputSheet.TabIndex = 18
        Me.btnDuplicateInputSheet.Text = "Duplicate Input Sheet"
        Me.ToolTip1.SetToolTip(Me.btnDuplicateInputSheet, "Step One")
        Me.btnDuplicateInputSheet.UseVisualStyleBackColor = True
        '
        'btnProcessDuplicateRows
        '
        Me.btnProcessDuplicateRows.Location = New System.Drawing.Point(10, 65)
        Me.btnProcessDuplicateRows.Name = "btnProcessDuplicateRows"
        Me.btnProcessDuplicateRows.Size = New System.Drawing.Size(142, 23)
        Me.btnProcessDuplicateRows.TabIndex = 17
        Me.btnProcessDuplicateRows.Text = "Process Duplicate Rows"
        Me.ToolTip1.SetToolTip(Me.btnProcessDuplicateRows, "Step One")
        Me.btnProcessDuplicateRows.UseVisualStyleBackColor = True
        '
        'btnProcessDynamicOutput
        '
        Me.btnProcessDynamicOutput.Location = New System.Drawing.Point(10, 19)
        Me.btnProcessDynamicOutput.Name = "btnProcessDynamicOutput"
        Me.btnProcessDynamicOutput.Size = New System.Drawing.Size(142, 23)
        Me.btnProcessDynamicOutput.TabIndex = 3
        Me.btnProcessDynamicOutput.Text = "Process Dynamic Output"
        Me.ToolTip1.SetToolTip(Me.btnProcessDynamicOutput, "Step One")
        Me.btnProcessDynamicOutput.UseVisualStyleBackColor = True
        '
        'btnMergeDuplicateRows
        '
        Me.btnMergeDuplicateRows.Location = New System.Drawing.Point(15, 19)
        Me.btnMergeDuplicateRows.Name = "btnMergeDuplicateRows"
        Me.btnMergeDuplicateRows.Size = New System.Drawing.Size(155, 23)
        Me.btnMergeDuplicateRows.TabIndex = 30
        Me.btnMergeDuplicateRows.Text = "Merge Duplicate Rows"
        Me.btnMergeDuplicateRows.UseVisualStyleBackColor = True
        '
        'btnGetITRInformation
        '
        Me.btnGetITRInformation.Location = New System.Drawing.Point(15, 48)
        Me.btnGetITRInformation.Name = "btnGetITRInformation"
        Me.btnGetITRInformation.Size = New System.Drawing.Size(155, 23)
        Me.btnGetITRInformation.TabIndex = 31
        Me.btnGetITRInformation.Text = "Get ITR Information"
        Me.btnGetITRInformation.UseVisualStyleBackColor = True
        '
        'btnDisplayITRDetail
        '
        Me.btnDisplayITRDetail.Location = New System.Drawing.Point(15, 77)
        Me.btnDisplayITRDetail.Name = "btnDisplayITRDetail"
        Me.btnDisplayITRDetail.Size = New System.Drawing.Size(155, 23)
        Me.btnDisplayITRDetail.TabIndex = 32
        Me.btnDisplayITRDetail.Text = "Display ITR Detail"
        Me.btnDisplayITRDetail.UseVisualStyleBackColor = True
        '
        'gbDARTReport
        '
        Me.gbDARTReport.Controls.Add(Me.gbDebug)
        Me.gbDARTReport.Controls.Add(Me.cmbTeamName)
        Me.gbDARTReport.Controls.Add(Me.Label1)
        Me.gbDARTReport.Controls.Add(Me.btnProcessDARTOutput)
        Me.gbDARTReport.Location = New System.Drawing.Point(15, 13)
        Me.gbDARTReport.Name = "gbDARTReport"
        Me.gbDARTReport.Size = New System.Drawing.Size(182, 325)
        Me.gbDARTReport.TabIndex = 21
        Me.gbDARTReport.TabStop = False
        Me.gbDARTReport.Text = "DART Report"
        '
        'gbDebug
        '
        Me.gbDebug.Controls.Add(Me.btnAddPageBreaks)
        Me.gbDebug.Controls.Add(Me.btnFormatSourceITRs)
        Me.gbDebug.Controls.Add(Me.btnDuplicateInputSheet)
        Me.gbDebug.Controls.Add(Me.btnProcessDuplicateRows)
        Me.gbDebug.Controls.Add(Me.btnAddPivotTables)
        Me.gbDebug.Controls.Add(Me.btnAddListObjects)
        Me.gbDebug.Controls.Add(Me.btnProcessDynamicOutput)
        Me.gbDebug.Location = New System.Drawing.Point(9, 106)
        Me.gbDebug.Name = "gbDebug"
        Me.gbDebug.Size = New System.Drawing.Size(162, 207)
        Me.gbDebug.TabIndex = 24
        Me.gbDebug.TabStop = False
        Me.gbDebug.Text = "Debug"
        '
        'btnAddPageBreaks
        '
        Me.btnAddPageBreaks.Location = New System.Drawing.Point(10, 173)
        Me.btnAddPageBreaks.Name = "btnAddPageBreaks"
        Me.btnAddPageBreaks.Size = New System.Drawing.Size(142, 23)
        Me.btnAddPageBreaks.TabIndex = 20
        Me.btnAddPageBreaks.Text = "Add Page Breaks"
        Me.btnAddPageBreaks.UseVisualStyleBackColor = True
        '
        'btnFormatSourceITRs
        '
        Me.btnFormatSourceITRs.Location = New System.Drawing.Point(10, 150)
        Me.btnFormatSourceITRs.Name = "btnFormatSourceITRs"
        Me.btnFormatSourceITRs.Size = New System.Drawing.Size(142, 23)
        Me.btnFormatSourceITRs.TabIndex = 19
        Me.btnFormatSourceITRs.Text = "Format SourceITRs"
        Me.btnFormatSourceITRs.UseVisualStyleBackColor = True
        '
        'btnAddPivotTables
        '
        Me.btnAddPivotTables.Location = New System.Drawing.Point(10, 111)
        Me.btnAddPivotTables.Name = "btnAddPivotTables"
        Me.btnAddPivotTables.Size = New System.Drawing.Size(142, 23)
        Me.btnAddPivotTables.TabIndex = 16
        Me.btnAddPivotTables.Text = "Add Pivot Tables"
        Me.btnAddPivotTables.UseVisualStyleBackColor = True
        '
        'btnAddListObjects
        '
        Me.btnAddListObjects.Location = New System.Drawing.Point(10, 88)
        Me.btnAddListObjects.Name = "btnAddListObjects"
        Me.btnAddListObjects.Size = New System.Drawing.Size(142, 23)
        Me.btnAddListObjects.TabIndex = 6
        Me.btnAddListObjects.Text = "Add ListObjects"
        Me.btnAddListObjects.UseVisualStyleBackColor = True
        '
        'cmbTeamName
        '
        Me.cmbTeamName.FormattingEnabled = True
        Me.cmbTeamName.Items.AddRange(New Object() {"Data Services", "Integration Services", "Reporting Services"})
        Me.cmbTeamName.Location = New System.Drawing.Point(9, 32)
        Me.cmbTeamName.Name = "cmbTeamName"
        Me.cmbTeamName.Size = New System.Drawing.Size(161, 21)
        Me.cmbTeamName.TabIndex = 23
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(65, 13)
        Me.Label1.TabIndex = 22
        Me.Label1.Text = "Team Name"
        '
        'gbITRWork
        '
        Me.gbITRWork.Controls.Add(Me.btnDisplayITRDetail)
        Me.gbITRWork.Controls.Add(Me.btnGetITRInformation)
        Me.gbITRWork.Controls.Add(Me.btnMergeDuplicateRows)
        Me.gbITRWork.Location = New System.Drawing.Point(15, 344)
        Me.gbITRWork.Name = "gbITRWork"
        Me.gbITRWork.Size = New System.Drawing.Size(182, 288)
        Me.gbITRWork.TabIndex = 22
        Me.gbITRWork.TabStop = False
        Me.gbITRWork.Text = "ITR Work"
        '
        'TaskPane_ITRs
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.gbITRWork)
        Me.Controls.Add(Me.gbDARTReport)
        Me.Name = "TaskPane_ITRs"
        Me.Size = New System.Drawing.Size(211, 646)
        Me.gbDARTReport.ResumeLayout(False)
        Me.gbDARTReport.PerformLayout()
        Me.gbDebug.ResumeLayout(False)
        Me.gbITRWork.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents gbDARTReport As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnProcessDARTOutput As System.Windows.Forms.Button
    Friend WithEvents cmbTeamName As System.Windows.Forms.ComboBox
    Friend WithEvents gbDebug As System.Windows.Forms.GroupBox
    Friend WithEvents btnAddPageBreaks As System.Windows.Forms.Button
    Friend WithEvents btnFormatSourceITRs As System.Windows.Forms.Button
    Friend WithEvents btnDuplicateInputSheet As System.Windows.Forms.Button
    Friend WithEvents btnProcessDuplicateRows As System.Windows.Forms.Button
    Friend WithEvents btnAddPivotTables As System.Windows.Forms.Button
    Friend WithEvents btnAddListObjects As System.Windows.Forms.Button
    Friend WithEvents btnProcessDynamicOutput As System.Windows.Forms.Button
    Friend WithEvents gbITRWork As System.Windows.Forms.GroupBox
    Friend WithEvents btnMergeDuplicateRows As System.Windows.Forms.Button
    Friend WithEvents btnGetITRInformation As System.Windows.Forms.Button
    Friend WithEvents btnDisplayITRDetail As System.Windows.Forms.Button

End Class
