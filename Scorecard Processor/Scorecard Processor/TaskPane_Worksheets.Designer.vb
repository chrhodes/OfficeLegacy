<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TaskPane_Worksheets
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
        Me.btnCreateOnTimeDeliveryWorksheet = New System.Windows.Forms.Button
        Me.btnCreateBudgetVarianceWorksheet = New System.Windows.Forms.Button
        Me.btnCreateITRProcessingWorksheet = New System.Windows.Forms.Button
        Me.btnCreateScoreCardWorksheets = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.btnCreateSurveyWorksheets = New System.Windows.Forms.Button
        Me.btnCreateSurveyMappingWorksheet = New System.Windows.Forms.Button
        Me.btnCreateRollupWorksheets = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnCreateOnTimeDeliveryWorksheet
        '
        Me.btnCreateOnTimeDeliveryWorksheet.Location = New System.Drawing.Point(12, 159)
        Me.btnCreateOnTimeDeliveryWorksheet.Name = "btnCreateOnTimeDeliveryWorksheet"
        Me.btnCreateOnTimeDeliveryWorksheet.Size = New System.Drawing.Size(155, 25)
        Me.btnCreateOnTimeDeliveryWorksheet.TabIndex = 12
        Me.btnCreateOnTimeDeliveryWorksheet.Text = "On Time Delivery Worksheet"
        Me.btnCreateOnTimeDeliveryWorksheet.UseVisualStyleBackColor = True
        '
        'btnCreateBudgetVarianceWorksheet
        '
        Me.btnCreateBudgetVarianceWorksheet.Location = New System.Drawing.Point(12, 192)
        Me.btnCreateBudgetVarianceWorksheet.Name = "btnCreateBudgetVarianceWorksheet"
        Me.btnCreateBudgetVarianceWorksheet.Size = New System.Drawing.Size(155, 25)
        Me.btnCreateBudgetVarianceWorksheet.TabIndex = 13
        Me.btnCreateBudgetVarianceWorksheet.Text = "Budget Variance Worksheet"
        Me.btnCreateBudgetVarianceWorksheet.UseVisualStyleBackColor = True
        '
        'btnCreateITRProcessingWorksheet
        '
        Me.btnCreateITRProcessingWorksheet.Location = New System.Drawing.Point(12, 225)
        Me.btnCreateITRProcessingWorksheet.Name = "btnCreateITRProcessingWorksheet"
        Me.btnCreateITRProcessingWorksheet.Size = New System.Drawing.Size(155, 25)
        Me.btnCreateITRProcessingWorksheet.TabIndex = 14
        Me.btnCreateITRProcessingWorksheet.Text = "ITR Processing Worksheet"
        Me.btnCreateITRProcessingWorksheet.UseVisualStyleBackColor = True
        '
        'btnCreateScoreCardWorksheets
        '
        Me.btnCreateScoreCardWorksheets.Location = New System.Drawing.Point(12, 28)
        Me.btnCreateScoreCardWorksheets.Name = "btnCreateScoreCardWorksheets"
        Me.btnCreateScoreCardWorksheets.Size = New System.Drawing.Size(155, 25)
        Me.btnCreateScoreCardWorksheets.TabIndex = 24
        Me.btnCreateScoreCardWorksheets.Text = "ScoreCard Worksheets"
        Me.btnCreateScoreCardWorksheets.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btnCreateSurveyWorksheets)
        Me.GroupBox1.Controls.Add(Me.btnCreateSurveyMappingWorksheet)
        Me.GroupBox1.Controls.Add(Me.btnCreateRollupWorksheets)
        Me.GroupBox1.Controls.Add(Me.btnCreateScoreCardWorksheets)
        Me.GroupBox1.Controls.Add(Me.btnCreateITRProcessingWorksheet)
        Me.GroupBox1.Controls.Add(Me.btnCreateOnTimeDeliveryWorksheet)
        Me.GroupBox1.Controls.Add(Me.btnCreateBudgetVarianceWorksheet)
        Me.GroupBox1.Location = New System.Drawing.Point(10, 14)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(179, 261)
        Me.GroupBox1.TabIndex = 25
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Create Worksheets"
        '
        'btnCreateSurveyWorksheets
        '
        Me.btnCreateSurveyWorksheets.Location = New System.Drawing.Point(12, 61)
        Me.btnCreateSurveyWorksheets.Name = "btnCreateSurveyWorksheets"
        Me.btnCreateSurveyWorksheets.Size = New System.Drawing.Size(155, 25)
        Me.btnCreateSurveyWorksheets.TabIndex = 27
        Me.btnCreateSurveyWorksheets.Text = "Survey Worksheets (Teams)"
        Me.btnCreateSurveyWorksheets.UseVisualStyleBackColor = True
        '
        'btnCreateSurveyMappingWorksheet
        '
        Me.btnCreateSurveyMappingWorksheet.Location = New System.Drawing.Point(12, 126)
        Me.btnCreateSurveyMappingWorksheet.Name = "btnCreateSurveyMappingWorksheet"
        Me.btnCreateSurveyMappingWorksheet.Size = New System.Drawing.Size(154, 25)
        Me.btnCreateSurveyMappingWorksheet.TabIndex = 26
        Me.btnCreateSurveyMappingWorksheet.Text = "Survey Mapping Worksheet"
        Me.btnCreateSurveyMappingWorksheet.UseVisualStyleBackColor = True
        '
        'btnCreateRollupWorksheets
        '
        Me.btnCreateRollupWorksheets.Location = New System.Drawing.Point(12, 94)
        Me.btnCreateRollupWorksheets.Name = "btnCreateRollupWorksheets"
        Me.btnCreateRollupWorksheets.Size = New System.Drawing.Size(154, 24)
        Me.btnCreateRollupWorksheets.TabIndex = 25
        Me.btnCreateRollupWorksheets.Text = "Rollup Worksheets"
        Me.btnCreateRollupWorksheets.UseVisualStyleBackColor = True
        '
        'TaskPane_Worksheets
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "TaskPane_Worksheets"
        Me.Size = New System.Drawing.Size(200, 400)
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnCreateOnTimeDeliveryWorksheet As System.Windows.Forms.Button
    Friend WithEvents btnCreateBudgetVarianceWorksheet As System.Windows.Forms.Button
    Friend WithEvents btnCreateITRProcessingWorksheet As System.Windows.Forms.Button
    Friend WithEvents btnCreateScoreCardWorksheets As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btnCreateRollupWorksheets As System.Windows.Forms.Button
    Friend WithEvents btnCreateSurveyMappingWorksheet As System.Windows.Forms.Button
    Friend WithEvents btnCreateSurveyWorksheets As System.Windows.Forms.Button

End Class
