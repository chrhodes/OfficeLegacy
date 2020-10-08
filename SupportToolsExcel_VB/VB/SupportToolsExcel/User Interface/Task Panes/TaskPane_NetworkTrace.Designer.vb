<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TaskPane_NetworkTrace
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
        Me.btnFormatColumns = New System.Windows.Forms.Button()
        Me.gbDebug = New System.Windows.Forms.GroupBox()
        Me.btnHilightLostFrames = New System.Windows.Forms.Button()
        Me.btnHilightErrorSheet = New System.Windows.Forms.Button()
        Me.btnRemoveColumns = New System.Windows.Forms.Button()
        Me.btnFormatSheet = New System.Windows.Forms.Button()
        Me.btnDuplicateColumns = New System.Windows.Forms.Button()
        Me.txtHostCount = New System.Windows.Forms.TextBox()
        Me.btnDetectHosts = New System.Windows.Forms.Button()
        Me.btnHilightTime = New System.Windows.Forms.Button()
        Me.btnRemoveHex = New System.Windows.Forms.Button()
        Me.btnHilightTraceSheet = New System.Windows.Forms.Button()
        Me.btnClearData = New System.Windows.Forms.Button()
        Me.gbActions = New System.Windows.Forms.GroupBox()
        Me.btnCreateAnalysisSheet = New System.Windows.Forms.Button()
        Me.txtSheetName = New System.Windows.Forms.TextBox()
        Me.btnFormatTrace = New System.Windows.Forms.Button()
        Me.gbDebug.SuspendLayout()
        Me.gbActions.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnFormatColumns
        '
        Me.btnFormatColumns.Location = New System.Drawing.Point(20, 356)
        Me.btnFormatColumns.Name = "btnFormatColumns"
        Me.btnFormatColumns.Size = New System.Drawing.Size(153, 23)
        Me.btnFormatColumns.TabIndex = 0
        Me.btnFormatColumns.Text = "Format Columns"
        Me.btnFormatColumns.UseVisualStyleBackColor = True
        '
        'gbDebug
        '
        Me.gbDebug.Controls.Add(Me.btnHilightLostFrames)
        Me.gbDebug.Controls.Add(Me.btnHilightErrorSheet)
        Me.gbDebug.Controls.Add(Me.btnRemoveColumns)
        Me.gbDebug.Controls.Add(Me.btnFormatSheet)
        Me.gbDebug.Controls.Add(Me.btnDuplicateColumns)
        Me.gbDebug.Controls.Add(Me.txtHostCount)
        Me.gbDebug.Controls.Add(Me.btnDetectHosts)
        Me.gbDebug.Controls.Add(Me.btnHilightTime)
        Me.gbDebug.Controls.Add(Me.btnRemoveHex)
        Me.gbDebug.Controls.Add(Me.btnHilightTraceSheet)
        Me.gbDebug.Controls.Add(Me.btnClearData)
        Me.gbDebug.Controls.Add(Me.btnFormatColumns)
        Me.gbDebug.Location = New System.Drawing.Point(3, 207)
        Me.gbDebug.Name = "gbDebug"
        Me.gbDebug.Size = New System.Drawing.Size(194, 414)
        Me.gbDebug.TabIndex = 1
        Me.gbDebug.TabStop = False
        Me.gbDebug.Text = "Debug"
        '
        'btnHilightLostFrames
        '
        Me.btnHilightLostFrames.Location = New System.Drawing.Point(20, 48)
        Me.btnHilightLostFrames.Name = "btnHilightLostFrames"
        Me.btnHilightLostFrames.Size = New System.Drawing.Size(153, 23)
        Me.btnHilightLostFrames.TabIndex = 13
        Me.btnHilightLostFrames.Text = "Hilight Lost Frames"
        Me.btnHilightLostFrames.UseVisualStyleBackColor = True
        '
        'btnHilightErrorSheet
        '
        Me.btnHilightErrorSheet.Location = New System.Drawing.Point(20, 77)
        Me.btnHilightErrorSheet.Name = "btnHilightErrorSheet"
        Me.btnHilightErrorSheet.Size = New System.Drawing.Size(153, 23)
        Me.btnHilightErrorSheet.TabIndex = 12
        Me.btnHilightErrorSheet.Text = "Hilight Error Sheet"
        Me.btnHilightErrorSheet.UseVisualStyleBackColor = True
        '
        'btnRemoveColumns
        '
        Me.btnRemoveColumns.Location = New System.Drawing.Point(20, 217)
        Me.btnRemoveColumns.Name = "btnRemoveColumns"
        Me.btnRemoveColumns.Size = New System.Drawing.Size(153, 23)
        Me.btnRemoveColumns.TabIndex = 10
        Me.btnRemoveColumns.Text = "Remove Columns"
        Me.btnRemoveColumns.UseVisualStyleBackColor = True
        '
        'btnFormatSheet
        '
        Me.btnFormatSheet.Location = New System.Drawing.Point(20, 385)
        Me.btnFormatSheet.Name = "btnFormatSheet"
        Me.btnFormatSheet.Size = New System.Drawing.Size(153, 23)
        Me.btnFormatSheet.TabIndex = 9
        Me.btnFormatSheet.Text = "Format Sheet"
        Me.btnFormatSheet.UseVisualStyleBackColor = True
        '
        'btnDuplicateColumns
        '
        Me.btnDuplicateColumns.Location = New System.Drawing.Point(20, 188)
        Me.btnDuplicateColumns.Name = "btnDuplicateColumns"
        Me.btnDuplicateColumns.Size = New System.Drawing.Size(153, 23)
        Me.btnDuplicateColumns.TabIndex = 8
        Me.btnDuplicateColumns.Text = "Duplicate Columns"
        Me.btnDuplicateColumns.UseVisualStyleBackColor = True
        '
        'txtHostCount
        '
        Me.txtHostCount.Location = New System.Drawing.Point(144, 161)
        Me.txtHostCount.Name = "txtHostCount"
        Me.txtHostCount.Size = New System.Drawing.Size(29, 20)
        Me.txtHostCount.TabIndex = 7
        '
        'btnDetectHosts
        '
        Me.btnDetectHosts.Location = New System.Drawing.Point(20, 159)
        Me.btnDetectHosts.Name = "btnDetectHosts"
        Me.btnDetectHosts.Size = New System.Drawing.Size(118, 23)
        Me.btnDetectHosts.TabIndex = 6
        Me.btnDetectHosts.Text = "Detect Hosts"
        Me.btnDetectHosts.UseVisualStyleBackColor = True
        '
        'btnHilightTime
        '
        Me.btnHilightTime.Location = New System.Drawing.Point(20, 106)
        Me.btnHilightTime.Name = "btnHilightTime"
        Me.btnHilightTime.Size = New System.Drawing.Size(153, 23)
        Me.btnHilightTime.TabIndex = 5
        Me.btnHilightTime.Text = "Hilight Time"
        Me.btnHilightTime.UseVisualStyleBackColor = True
        '
        'btnRemoveHex
        '
        Me.btnRemoveHex.Location = New System.Drawing.Point(20, 19)
        Me.btnRemoveHex.Name = "btnRemoveHex"
        Me.btnRemoveHex.Size = New System.Drawing.Size(153, 23)
        Me.btnRemoveHex.TabIndex = 4
        Me.btnRemoveHex.Text = "Remove Hex"
        Me.btnRemoveHex.UseVisualStyleBackColor = True
        '
        'btnHilightTraceSheet
        '
        Me.btnHilightTraceSheet.Location = New System.Drawing.Point(20, 275)
        Me.btnHilightTraceSheet.Name = "btnHilightTraceSheet"
        Me.btnHilightTraceSheet.Size = New System.Drawing.Size(153, 23)
        Me.btnHilightTraceSheet.TabIndex = 3
        Me.btnHilightTraceSheet.Text = "Hilight Trace Sheet"
        Me.btnHilightTraceSheet.UseVisualStyleBackColor = True
        '
        'btnClearData
        '
        Me.btnClearData.Location = New System.Drawing.Point(20, 246)
        Me.btnClearData.Name = "btnClearData"
        Me.btnClearData.Size = New System.Drawing.Size(153, 23)
        Me.btnClearData.TabIndex = 2
        Me.btnClearData.Text = "Clear Data"
        Me.btnClearData.UseVisualStyleBackColor = True
        '
        'gbActions
        '
        Me.gbActions.Controls.Add(Me.btnCreateAnalysisSheet)
        Me.gbActions.Controls.Add(Me.txtSheetName)
        Me.gbActions.Controls.Add(Me.btnFormatTrace)
        Me.gbActions.Location = New System.Drawing.Point(3, 23)
        Me.gbActions.Name = "gbActions"
        Me.gbActions.Size = New System.Drawing.Size(194, 148)
        Me.gbActions.TabIndex = 6
        Me.gbActions.TabStop = False
        Me.gbActions.Text = "Actions"
        '
        'btnCreateAnalysisSheet
        '
        Me.btnCreateAnalysisSheet.Location = New System.Drawing.Point(20, 42)
        Me.btnCreateAnalysisSheet.Name = "btnCreateAnalysisSheet"
        Me.btnCreateAnalysisSheet.Size = New System.Drawing.Size(153, 23)
        Me.btnCreateAnalysisSheet.TabIndex = 3
        Me.btnCreateAnalysisSheet.Text = "Create Analysis Sheet"
        Me.btnCreateAnalysisSheet.UseVisualStyleBackColor = True
        '
        'txtSheetName
        '
        Me.txtSheetName.Location = New System.Drawing.Point(20, 19)
        Me.txtSheetName.Name = "txtSheetName"
        Me.txtSheetName.Size = New System.Drawing.Size(153, 20)
        Me.txtSheetName.TabIndex = 2
        '
        'btnFormatTrace
        '
        Me.btnFormatTrace.Location = New System.Drawing.Point(20, 119)
        Me.btnFormatTrace.Name = "btnFormatTrace"
        Me.btnFormatTrace.Size = New System.Drawing.Size(153, 23)
        Me.btnFormatTrace.TabIndex = 1
        Me.btnFormatTrace.Text = "Format Trace"
        Me.btnFormatTrace.UseVisualStyleBackColor = True
        '
        'TaskPane_NetworkTrace
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.gbActions)
        Me.Controls.Add(Me.gbDebug)
        Me.Name = "TaskPane_NetworkTrace"
        Me.Size = New System.Drawing.Size(200, 691)
        Me.gbDebug.ResumeLayout(False)
        Me.gbDebug.PerformLayout()
        Me.gbActions.ResumeLayout(False)
        Me.gbActions.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnFormatColumns As System.Windows.Forms.Button
    Friend WithEvents gbDebug As System.Windows.Forms.GroupBox
    Friend WithEvents btnHilightTraceSheet As System.Windows.Forms.Button
    Friend WithEvents btnClearData As System.Windows.Forms.Button
    Friend WithEvents btnRemoveHex As System.Windows.Forms.Button
    Friend WithEvents btnHilightTime As System.Windows.Forms.Button
    Friend WithEvents gbActions As System.Windows.Forms.GroupBox
    Friend WithEvents btnFormatTrace As System.Windows.Forms.Button
    Friend WithEvents btnCreateAnalysisSheet As System.Windows.Forms.Button
    Friend WithEvents txtSheetName As System.Windows.Forms.TextBox
    Friend WithEvents txtHostCount As System.Windows.Forms.TextBox
    Friend WithEvents btnDetectHosts As System.Windows.Forms.Button
    Friend WithEvents btnDuplicateColumns As System.Windows.Forms.Button
    Friend WithEvents btnFormatSheet As System.Windows.Forms.Button
    Friend WithEvents btnRemoveColumns As System.Windows.Forms.Button
    Friend WithEvents btnHilightErrorSheet As System.Windows.Forms.Button
    Friend WithEvents btnHilightLostFrames As System.Windows.Forms.Button

End Class
