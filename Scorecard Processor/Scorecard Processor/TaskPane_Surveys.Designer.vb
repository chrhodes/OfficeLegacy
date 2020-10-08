<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TaskPane_Surveys
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
        Me.btnLoadSurveyData = New System.Windows.Forms.Button
        Me.btnCreateDataWorksheet = New System.Windows.Forms.Button
        Me.btnAddFormulas = New System.Windows.Forms.Button
        Me.btnFormatWorksheet = New System.Windows.Forms.Button
        Me.cmbSurveyName = New System.Windows.Forms.ComboBox
        Me.btnAddCharts = New System.Windows.Forms.Button
        Me.btnLoadSurveyQuestions = New System.Windows.Forms.Button
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnLoadSurveyData
        '
        Me.btnLoadSurveyData.Location = New System.Drawing.Point(5, 77)
        Me.btnLoadSurveyData.Name = "btnLoadSurveyData"
        Me.btnLoadSurveyData.Size = New System.Drawing.Size(164, 23)
        Me.btnLoadSurveyData.TabIndex = 0
        Me.btnLoadSurveyData.Text = "Load Survey Data"
        Me.btnLoadSurveyData.UseVisualStyleBackColor = True
        '
        'btnCreateDataWorksheet
        '
        Me.btnCreateDataWorksheet.Location = New System.Drawing.Point(5, 106)
        Me.btnCreateDataWorksheet.Name = "btnCreateDataWorksheet"
        Me.btnCreateDataWorksheet.Size = New System.Drawing.Size(164, 23)
        Me.btnCreateDataWorksheet.TabIndex = 1
        Me.btnCreateDataWorksheet.Text = "CreateDataWorksheet"
        Me.btnCreateDataWorksheet.UseVisualStyleBackColor = True
        '
        'btnAddFormulas
        '
        Me.btnAddFormulas.Location = New System.Drawing.Point(5, 135)
        Me.btnAddFormulas.Name = "btnAddFormulas"
        Me.btnAddFormulas.Size = New System.Drawing.Size(164, 23)
        Me.btnAddFormulas.TabIndex = 2
        Me.btnAddFormulas.Text = "Add Formulas"
        Me.btnAddFormulas.UseVisualStyleBackColor = True
        '
        'btnFormatWorksheet
        '
        Me.btnFormatWorksheet.Location = New System.Drawing.Point(5, 164)
        Me.btnFormatWorksheet.Name = "btnFormatWorksheet"
        Me.btnFormatWorksheet.Size = New System.Drawing.Size(164, 23)
        Me.btnFormatWorksheet.TabIndex = 3
        Me.btnFormatWorksheet.Text = "Format Worksheet"
        Me.btnFormatWorksheet.UseVisualStyleBackColor = True
        '
        'cmbSurveyName
        '
        Me.cmbSurveyName.FormattingEnabled = True
        Me.cmbSurveyName.Items.AddRange(New Object() {"Partner Survey", "Business Survey", "IT Survey", "Help Desk Survey"})
        Me.cmbSurveyName.Location = New System.Drawing.Point(5, 50)
        Me.cmbSurveyName.Name = "cmbSurveyName"
        Me.cmbSurveyName.Size = New System.Drawing.Size(164, 21)
        Me.cmbSurveyName.TabIndex = 4
        '
        'btnAddCharts
        '
        Me.btnAddCharts.Location = New System.Drawing.Point(5, 193)
        Me.btnAddCharts.Name = "btnAddCharts"
        Me.btnAddCharts.Size = New System.Drawing.Size(164, 24)
        Me.btnAddCharts.TabIndex = 5
        Me.btnAddCharts.Text = "Add Charts"
        Me.btnAddCharts.UseVisualStyleBackColor = True
        '
        'btnLoadSurveyQuestions
        '
        Me.btnLoadSurveyQuestions.Location = New System.Drawing.Point(5, 21)
        Me.btnLoadSurveyQuestions.Name = "btnLoadSurveyQuestions"
        Me.btnLoadSurveyQuestions.Size = New System.Drawing.Size(164, 23)
        Me.btnLoadSurveyQuestions.TabIndex = 21
        Me.btnLoadSurveyQuestions.Text = "Load Survey Questions"
        Me.btnLoadSurveyQuestions.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.btnLoadSurveyQuestions)
        Me.GroupBox3.Controls.Add(Me.btnAddCharts)
        Me.GroupBox3.Controls.Add(Me.cmbSurveyName)
        Me.GroupBox3.Controls.Add(Me.btnFormatWorksheet)
        Me.GroupBox3.Controls.Add(Me.btnAddFormulas)
        Me.GroupBox3.Controls.Add(Me.btnCreateDataWorksheet)
        Me.GroupBox3.Controls.Add(Me.btnLoadSurveyData)
        Me.GroupBox3.Location = New System.Drawing.Point(10, 11)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(179, 223)
        Me.GroupBox3.TabIndex = 26
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Process Survey Data"
        '
        'TaskPane_Surveys
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.GroupBox3)
        Me.Name = "TaskPane_Surveys"
        Me.Size = New System.Drawing.Size(200, 263)
        Me.GroupBox3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnLoadSurveyData As System.Windows.Forms.Button
    Friend WithEvents btnCreateDataWorksheet As System.Windows.Forms.Button
    Friend WithEvents btnAddFormulas As System.Windows.Forms.Button
    Friend WithEvents btnFormatWorksheet As System.Windows.Forms.Button
    Friend WithEvents cmbSurveyName As System.Windows.Forms.ComboBox
    Friend WithEvents btnAddCharts As System.Windows.Forms.Button
    Friend WithEvents btnLoadSurveyQuestions As System.Windows.Forms.Button
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox

End Class
