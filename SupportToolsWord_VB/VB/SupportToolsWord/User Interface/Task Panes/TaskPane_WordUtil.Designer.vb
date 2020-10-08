<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TaskPane_WordUtil
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.btnProduceIndividualTeamScorecard = New System.Windows.Forms.Button
        Me.cbTeams = New System.Windows.Forms.ComboBox
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.chkIncludeFeedback = New System.Windows.Forms.CheckBox
        Me.chkIncludeCharts = New System.Windows.Forms.CheckBox
        Me.btnAddSurveyResultsToWordIndividualTeam = New System.Windows.Forms.Button
        Me.btnAddSurveyResultsToPowerPointIndividualTeam = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.GroupBox6 = New System.Windows.Forms.GroupBox
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.CheckBox2 = New System.Windows.Forms.CheckBox
        Me.btnAllSurveyResultsToWordAllTeams = New System.Windows.Forms.Button
        Me.btnAddSurveyResultsToPowerPointAllTeams = New System.Windows.Forms.Button
        Me.GroupBox5 = New System.Windows.Forms.GroupBox
        Me.btnProduceAllTeamsScorecards = New System.Windows.Forms.Button
        Me.btnProduceAllTeamsSurveys = New System.Windows.Forms.Button
        Me.btnProduceAllTeamsScorecard = New System.Windows.Forms.Button
        Me.btnClearDestinationSheets = New System.Windows.Forms.Button
        Me.btnCopyValues = New System.Windows.Forms.Button
        Me.btnFormatAllSurveySheets = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.GroupBox4)
        Me.GroupBox1.Controls.Add(Me.cbTeams)
        Me.GroupBox1.Controls.Add(Me.GroupBox3)
        Me.GroupBox1.Location = New System.Drawing.Point(3, 100)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(190, 264)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Individual Team"
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.btnProduceIndividualTeamScorecard)
        Me.GroupBox4.Location = New System.Drawing.Point(13, 46)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(163, 55)
        Me.GroupBox4.TabIndex = 29
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Scorecard Results"
        '
        'btnProduceIndividualTeamScorecard
        '
        Me.btnProduceIndividualTeamScorecard.Location = New System.Drawing.Point(6, 18)
        Me.btnProduceIndividualTeamScorecard.Name = "btnProduceIndividualTeamScorecard"
        Me.btnProduceIndividualTeamScorecard.Size = New System.Drawing.Size(150, 30)
        Me.btnProduceIndividualTeamScorecard.TabIndex = 0
        Me.btnProduceIndividualTeamScorecard.Text = "Produce Scorecard"
        Me.btnProduceIndividualTeamScorecard.UseVisualStyleBackColor = True
        '
        'cbTeams
        '
        Me.cbTeams.FormattingEnabled = True
        Me.cbTeams.Location = New System.Drawing.Point(13, 19)
        Me.cbTeams.Name = "cbTeams"
        Me.cbTeams.Size = New System.Drawing.Size(163, 21)
        Me.cbTeams.TabIndex = 20
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.chkIncludeFeedback)
        Me.GroupBox3.Controls.Add(Me.chkIncludeCharts)
        Me.GroupBox3.Controls.Add(Me.btnAddSurveyResultsToWordIndividualTeam)
        Me.GroupBox3.Controls.Add(Me.btnAddSurveyResultsToPowerPointIndividualTeam)
        Me.GroupBox3.Location = New System.Drawing.Point(13, 107)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(164, 150)
        Me.GroupBox3.TabIndex = 28
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Survey Results"
        '
        'chkIncludeFeedback
        '
        Me.chkIncludeFeedback.AutoSize = True
        Me.chkIncludeFeedback.Checked = True
        Me.chkIncludeFeedback.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkIncludeFeedback.Location = New System.Drawing.Point(6, 43)
        Me.chkIncludeFeedback.Name = "chkIncludeFeedback"
        Me.chkIncludeFeedback.Size = New System.Drawing.Size(112, 17)
        Me.chkIncludeFeedback.TabIndex = 30
        Me.chkIncludeFeedback.Text = "Include Feedback"
        Me.chkIncludeFeedback.UseVisualStyleBackColor = True
        '
        'chkIncludeCharts
        '
        Me.chkIncludeCharts.AutoSize = True
        Me.chkIncludeCharts.Checked = True
        Me.chkIncludeCharts.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkIncludeCharts.Location = New System.Drawing.Point(6, 19)
        Me.chkIncludeCharts.Name = "chkIncludeCharts"
        Me.chkIncludeCharts.Size = New System.Drawing.Size(94, 17)
        Me.chkIncludeCharts.TabIndex = 29
        Me.chkIncludeCharts.Text = "Include Charts"
        Me.chkIncludeCharts.UseVisualStyleBackColor = True
        '
        'btnAddSurveyResultsToWordIndividualTeam
        '
        Me.btnAddSurveyResultsToWordIndividualTeam.Location = New System.Drawing.Point(6, 108)
        Me.btnAddSurveyResultsToWordIndividualTeam.Name = "btnAddSurveyResultsToWordIndividualTeam"
        Me.btnAddSurveyResultsToWordIndividualTeam.Size = New System.Drawing.Size(150, 30)
        Me.btnAddSurveyResultsToWordIndividualTeam.TabIndex = 22
        Me.btnAddSurveyResultsToWordIndividualTeam.Text = "Produce Word Output"
        Me.btnAddSurveyResultsToWordIndividualTeam.UseVisualStyleBackColor = True
        '
        'btnAddSurveyResultsToPowerPointIndividualTeam
        '
        Me.btnAddSurveyResultsToPowerPointIndividualTeam.Location = New System.Drawing.Point(6, 66)
        Me.btnAddSurveyResultsToPowerPointIndividualTeam.Name = "btnAddSurveyResultsToPowerPointIndividualTeam"
        Me.btnAddSurveyResultsToPowerPointIndividualTeam.Size = New System.Drawing.Size(150, 30)
        Me.btnAddSurveyResultsToPowerPointIndividualTeam.TabIndex = 8
        Me.btnAddSurveyResultsToPowerPointIndividualTeam.Text = "Produce PowerPoint Output"
        Me.btnAddSurveyResultsToPowerPointIndividualTeam.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.GroupBox6)
        Me.GroupBox2.Controls.Add(Me.GroupBox5)
        Me.GroupBox2.Controls.Add(Me.btnProduceAllTeamsSurveys)
        Me.GroupBox2.Controls.Add(Me.btnProduceAllTeamsScorecard)
        Me.GroupBox2.Location = New System.Drawing.Point(3, 370)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(190, 309)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "All Teams"
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.CheckBox1)
        Me.GroupBox6.Controls.Add(Me.CheckBox2)
        Me.GroupBox6.Controls.Add(Me.btnAllSurveyResultsToWordAllTeams)
        Me.GroupBox6.Controls.Add(Me.btnAddSurveyResultsToPowerPointAllTeams)
        Me.GroupBox6.Location = New System.Drawing.Point(12, 152)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(164, 150)
        Me.GroupBox6.TabIndex = 31
        Me.GroupBox6.TabStop = False
        Me.GroupBox6.Text = "Survey Results"
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Checked = True
        Me.CheckBox1.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox1.Location = New System.Drawing.Point(6, 43)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(112, 17)
        Me.CheckBox1.TabIndex = 30
        Me.CheckBox1.Text = "Include Feedback"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Checked = True
        Me.CheckBox2.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox2.Location = New System.Drawing.Point(6, 19)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(94, 17)
        Me.CheckBox2.TabIndex = 29
        Me.CheckBox2.Text = "Include Charts"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'btnAllSurveyResultsToWordAllTeams
        '
        Me.btnAllSurveyResultsToWordAllTeams.Location = New System.Drawing.Point(6, 108)
        Me.btnAllSurveyResultsToWordAllTeams.Name = "btnAllSurveyResultsToWordAllTeams"
        Me.btnAllSurveyResultsToWordAllTeams.Size = New System.Drawing.Size(150, 30)
        Me.btnAllSurveyResultsToWordAllTeams.TabIndex = 22
        Me.btnAllSurveyResultsToWordAllTeams.Text = "Produce Word Output"
        Me.btnAllSurveyResultsToWordAllTeams.UseVisualStyleBackColor = True
        '
        'btnAddSurveyResultsToPowerPointAllTeams
        '
        Me.btnAddSurveyResultsToPowerPointAllTeams.Location = New System.Drawing.Point(6, 66)
        Me.btnAddSurveyResultsToPowerPointAllTeams.Name = "btnAddSurveyResultsToPowerPointAllTeams"
        Me.btnAddSurveyResultsToPowerPointAllTeams.Size = New System.Drawing.Size(150, 30)
        Me.btnAddSurveyResultsToPowerPointAllTeams.TabIndex = 8
        Me.btnAddSurveyResultsToPowerPointAllTeams.Text = "Produce PowerPoint Output"
        Me.btnAddSurveyResultsToPowerPointAllTeams.UseVisualStyleBackColor = True
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.btnProduceAllTeamsScorecards)
        Me.GroupBox5.Location = New System.Drawing.Point(11, 55)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(163, 55)
        Me.GroupBox5.TabIndex = 30
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Scorecard Results"
        '
        'btnProduceAllTeamsScorecards
        '
        Me.btnProduceAllTeamsScorecards.Location = New System.Drawing.Point(6, 18)
        Me.btnProduceAllTeamsScorecards.Name = "btnProduceAllTeamsScorecards"
        Me.btnProduceAllTeamsScorecards.Size = New System.Drawing.Size(150, 30)
        Me.btnProduceAllTeamsScorecards.TabIndex = 0
        Me.btnProduceAllTeamsScorecards.Text = "Produce Scorecards"
        Me.btnProduceAllTeamsScorecards.UseVisualStyleBackColor = True
        '
        'btnProduceAllTeamsSurveys
        '
        Me.btnProduceAllTeamsSurveys.Location = New System.Drawing.Point(11, 116)
        Me.btnProduceAllTeamsSurveys.Name = "btnProduceAllTeamsSurveys"
        Me.btnProduceAllTeamsSurveys.Size = New System.Drawing.Size(166, 30)
        Me.btnProduceAllTeamsSurveys.TabIndex = 1
        Me.btnProduceAllTeamsSurveys.Text = "Produce All Teams Surveys"
        Me.btnProduceAllTeamsSurveys.UseVisualStyleBackColor = True
        '
        'btnProduceAllTeamsScorecard
        '
        Me.btnProduceAllTeamsScorecard.Location = New System.Drawing.Point(11, 19)
        Me.btnProduceAllTeamsScorecard.Name = "btnProduceAllTeamsScorecard"
        Me.btnProduceAllTeamsScorecard.Size = New System.Drawing.Size(165, 30)
        Me.btnProduceAllTeamsScorecard.TabIndex = 0
        Me.btnProduceAllTeamsScorecard.Text = "Produce All Teams Scorecard"
        Me.btnProduceAllTeamsScorecard.UseVisualStyleBackColor = True
        '
        'btnClearDestinationSheets
        '
        Me.btnClearDestinationSheets.Location = New System.Drawing.Point(3, 14)
        Me.btnClearDestinationSheets.Name = "btnClearDestinationSheets"
        Me.btnClearDestinationSheets.Size = New System.Drawing.Size(190, 20)
        Me.btnClearDestinationSheets.TabIndex = 33
        Me.btnClearDestinationSheets.Text = "Clear Destination Sheets"
        Me.btnClearDestinationSheets.UseVisualStyleBackColor = True
        '
        'btnCopyValues
        '
        Me.btnCopyValues.Location = New System.Drawing.Point(3, 37)
        Me.btnCopyValues.Name = "btnCopyValues"
        Me.btnCopyValues.Size = New System.Drawing.Size(190, 20)
        Me.btnCopyValues.TabIndex = 32
        Me.btnCopyValues.Text = "Copy Values"
        Me.btnCopyValues.UseVisualStyleBackColor = True
        '
        'btnFormatAllSurveySheets
        '
        Me.btnFormatAllSurveySheets.Location = New System.Drawing.Point(3, 61)
        Me.btnFormatAllSurveySheets.Name = "btnFormatAllSurveySheets"
        Me.btnFormatAllSurveySheets.Size = New System.Drawing.Size(190, 20)
        Me.btnFormatAllSurveySheets.TabIndex = 34
        Me.btnFormatAllSurveySheets.Text = "Format All Survey Sheets"
        Me.btnFormatAllSurveySheets.UseVisualStyleBackColor = True
        '
        'TaskPane_Results
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.btnFormatAllSurveySheets)
        Me.Controls.Add(Me.btnCopyValues)
        Me.Controls.Add(Me.btnClearDestinationSheets)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "TaskPane_Results"
        Me.Size = New System.Drawing.Size(200, 691)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox6.ResumeLayout(False)
        Me.GroupBox6.PerformLayout()
        Me.GroupBox5.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents cbTeams As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents chkIncludeFeedback As System.Windows.Forms.CheckBox
    Friend WithEvents chkIncludeCharts As System.Windows.Forms.CheckBox
    Friend WithEvents btnAddSurveyResultsToWordIndividualTeam As System.Windows.Forms.Button
    Friend WithEvents btnAddSurveyResultsToPowerPointIndividualTeam As System.Windows.Forms.Button
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents btnProduceIndividualTeamScorecard As System.Windows.Forms.Button
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents btnProduceAllTeamsScorecards As System.Windows.Forms.Button
    Friend WithEvents btnProduceAllTeamsSurveys As System.Windows.Forms.Button
    Friend WithEvents btnProduceAllTeamsScorecard As System.Windows.Forms.Button
    Friend WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents btnAllSurveyResultsToWordAllTeams As System.Windows.Forms.Button
    Friend WithEvents btnAddSurveyResultsToPowerPointAllTeams As System.Windows.Forms.Button
    Friend WithEvents btnClearDestinationSheets As System.Windows.Forms.Button
    Friend WithEvents btnCopyValues As System.Windows.Forms.Button
    Friend WithEvents btnFormatAllSurveySheets As System.Windows.Forms.Button

End Class
