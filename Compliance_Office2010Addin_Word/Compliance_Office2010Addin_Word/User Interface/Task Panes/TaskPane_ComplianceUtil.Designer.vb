<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TaskPane_ComplianceUtil
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
        Me.gbTagIndexWords = New System.Windows.Forms.GroupBox()
        Me.btnCreateIndexStyles = New System.Windows.Forms.Button()
        Me.btnMarkIndexWords = New System.Windows.Forms.Button()
        Me.btnCreateIndex = New System.Windows.Forms.Button()
        Me.btnTagIndexWords = New System.Windows.Forms.Button()
        Me.btnFindIndexWords = New System.Windows.Forms.Button()
        Me.gbImproveReadability = New System.Windows.Forms.GroupBox()
        Me.btnLoadReplacementWords = New System.Windows.Forms.Button()
        Me.ckIndexWordsOnly = New System.Windows.Forms.CheckBox()
        Me.txtReplacementWord = New System.Windows.Forms.TextBox()
        Me.btnSaveReplacementWords = New System.Windows.Forms.Button()
        Me.lblReplacementWord = New System.Windows.Forms.Label()
        Me.btnZapReplacementWords = New System.Windows.Forms.Button()
        Me.saveFileDialog = New System.Windows.Forms.SaveFileDialog()
        Me.openFileDialog = New System.Windows.Forms.OpenFileDialog()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.gbTagIndexWords.SuspendLayout
        Me.gbImproveReadability.SuspendLayout
        Me.SuspendLayout
        '
        'gbTagIndexWords
        '
        Me.gbTagIndexWords.Controls.Add(Me.btnCreateIndexStyles)
        Me.gbTagIndexWords.Controls.Add(Me.btnMarkIndexWords)
        Me.gbTagIndexWords.Controls.Add(Me.btnCreateIndex)
        Me.gbTagIndexWords.Controls.Add(Me.btnTagIndexWords)
        Me.gbTagIndexWords.Controls.Add(Me.btnFindIndexWords)
        Me.gbTagIndexWords.Location = New System.Drawing.Point(14, 16)
        Me.gbTagIndexWords.Name = "gbTagIndexWords"
        Me.gbTagIndexWords.Size = New System.Drawing.Size(200, 243)
        Me.gbTagIndexWords.TabIndex = 0
        Me.gbTagIndexWords.TabStop = false
        Me.gbTagIndexWords.Text = "Tag Index Words"
        '
        'btnCreateIndexStyles
        '
        Me.btnCreateIndexStyles.Location = New System.Drawing.Point(15, 58)
        Me.btnCreateIndexStyles.Name = "btnCreateIndexStyles"
        Me.btnCreateIndexStyles.Size = New System.Drawing.Size(171, 23)
        Me.btnCreateIndexStyles.TabIndex = 9
        Me.btnCreateIndexStyles.Text = "Create Index Styles"
        Me.ToolTip1.SetToolTip(Me.btnCreateIndexStyles, "Creates IndexStyle formatting style in current document")
        Me.btnCreateIndexStyles.UseVisualStyleBackColor = true
        '
        'btnMarkIndexWords
        '
        Me.btnMarkIndexWords.Location = New System.Drawing.Point(15, 95)
        Me.btnMarkIndexWords.Name = "btnMarkIndexWords"
        Me.btnMarkIndexWords.Size = New System.Drawing.Size(171, 23)
        Me.btnMarkIndexWords.TabIndex = 8
        Me.btnMarkIndexWords.Text = "Mark Index Words"
        Me.ToolTip1.SetToolTip(Me.btnMarkIndexWords, "Marks words contained in XML file with IndexStyle")
        Me.btnMarkIndexWords.UseVisualStyleBackColor = true
        '
        'btnCreateIndex
        '
        Me.btnCreateIndex.Location = New System.Drawing.Point(15, 203)
        Me.btnCreateIndex.Name = "btnCreateIndex"
        Me.btnCreateIndex.Size = New System.Drawing.Size(171, 23)
        Me.btnCreateIndex.TabIndex = 2
        Me.btnCreateIndex.Text = "Create Index"
        Me.btnCreateIndex.UseVisualStyleBackColor = true
        '
        'btnTagIndexWords
        '
        Me.btnTagIndexWords.Location = New System.Drawing.Point(15, 161)
        Me.btnTagIndexWords.Name = "btnTagIndexWords"
        Me.btnTagIndexWords.Size = New System.Drawing.Size(171, 23)
        Me.btnTagIndexWords.TabIndex = 1
        Me.btnTagIndexWords.Text = "Tag Index Words"
        Me.btnTagIndexWords.UseVisualStyleBackColor = true
        '
        'btnFindIndexWords
        '
        Me.btnFindIndexWords.Location = New System.Drawing.Point(15, 29)
        Me.btnFindIndexWords.Name = "btnFindIndexWords"
        Me.btnFindIndexWords.Size = New System.Drawing.Size(171, 23)
        Me.btnFindIndexWords.TabIndex = 0
        Me.btnFindIndexWords.Text = "Find Index Words"
        Me.btnFindIndexWords.UseVisualStyleBackColor = true
        '
        'gbImproveReadability
        '
        Me.gbImproveReadability.Controls.Add(Me.btnLoadReplacementWords)
        Me.gbImproveReadability.Controls.Add(Me.ckIndexWordsOnly)
        Me.gbImproveReadability.Controls.Add(Me.txtReplacementWord)
        Me.gbImproveReadability.Controls.Add(Me.btnSaveReplacementWords)
        Me.gbImproveReadability.Controls.Add(Me.lblReplacementWord)
        Me.gbImproveReadability.Controls.Add(Me.btnZapReplacementWords)
        Me.gbImproveReadability.Location = New System.Drawing.Point(14, 287)
        Me.gbImproveReadability.Name = "gbImproveReadability"
        Me.gbImproveReadability.Size = New System.Drawing.Size(200, 243)
        Me.gbImproveReadability.TabIndex = 1
        Me.gbImproveReadability.TabStop = false
        Me.gbImproveReadability.Text = "Improve Readability"
        '
        'btnLoadReplacementWords
        '
        Me.btnLoadReplacementWords.Location = New System.Drawing.Point(15, 83)
        Me.btnLoadReplacementWords.Name = "btnLoadReplacementWords"
        Me.btnLoadReplacementWords.Size = New System.Drawing.Size(171, 23)
        Me.btnLoadReplacementWords.TabIndex = 7
        Me.btnLoadReplacementWords.Text = "Load Replacement Words"
        Me.btnLoadReplacementWords.UseVisualStyleBackColor = true
        '
        'ckIndexWordsOnly
        '
        Me.ckIndexWordsOnly.AutoSize = true
        Me.ckIndexWordsOnly.Location = New System.Drawing.Point(15, 201)
        Me.ckIndexWordsOnly.Name = "ckIndexWordsOnly"
        Me.ckIndexWordsOnly.Size = New System.Drawing.Size(110, 17)
        Me.ckIndexWordsOnly.TabIndex = 6
        Me.ckIndexWordsOnly.Text = "Index Words Only"
        Me.ckIndexWordsOnly.UseVisualStyleBackColor = true
        '
        'txtReplacementWord
        '
        Me.txtReplacementWord.Location = New System.Drawing.Point(117, 144)
        Me.txtReplacementWord.Name = "txtReplacementWord"
        Me.txtReplacementWord.Size = New System.Drawing.Size(69, 20)
        Me.txtReplacementWord.TabIndex = 5
        Me.txtReplacementWord.Text = "Simple"
        '
        'btnSaveReplacementWords
        '
        Me.btnSaveReplacementWords.Location = New System.Drawing.Point(15, 19)
        Me.btnSaveReplacementWords.Name = "btnSaveReplacementWords"
        Me.btnSaveReplacementWords.Size = New System.Drawing.Size(171, 23)
        Me.btnSaveReplacementWords.TabIndex = 2
        Me.btnSaveReplacementWords.Text = "Save Replacement Words"
        Me.btnSaveReplacementWords.UseVisualStyleBackColor = true
        '
        'lblReplacementWord
        '
        Me.lblReplacementWord.AutoSize = true
        Me.lblReplacementWord.Location = New System.Drawing.Point(12, 147)
        Me.lblReplacementWord.Name = "lblReplacementWord"
        Me.lblReplacementWord.Size = New System.Drawing.Size(99, 13)
        Me.lblReplacementWord.TabIndex = 4
        Me.lblReplacementWord.Text = "Replacement Word"
        '
        'btnZapReplacementWords
        '
        Me.btnZapReplacementWords.Location = New System.Drawing.Point(15, 172)
        Me.btnZapReplacementWords.Name = "btnZapReplacementWords"
        Me.btnZapReplacementWords.Size = New System.Drawing.Size(171, 23)
        Me.btnZapReplacementWords.TabIndex = 3
        Me.btnZapReplacementWords.Text = "Zap Replacement Words"
        Me.btnZapReplacementWords.UseVisualStyleBackColor = true
        '
        'TaskPane_ComplianceUtil
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6!, 13!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.gbImproveReadability)
        Me.Controls.Add(Me.gbTagIndexWords)
        Me.Name = "TaskPane_ComplianceUtil"
        Me.Size = New System.Drawing.Size(227, 691)
        Me.gbTagIndexWords.ResumeLayout(false)
        Me.gbImproveReadability.ResumeLayout(false)
        Me.gbImproveReadability.PerformLayout
        Me.ResumeLayout(false)

End Sub
    Friend WithEvents gbTagIndexWords As System.Windows.Forms.GroupBox
    Friend WithEvents btnTagIndexWords As System.Windows.Forms.Button
    Friend WithEvents btnFindIndexWords As System.Windows.Forms.Button
    Friend WithEvents gbImproveReadability As System.Windows.Forms.GroupBox
    Friend WithEvents btnZapReplacementWords As System.Windows.Forms.Button
    Friend WithEvents btnSaveReplacementWords As System.Windows.Forms.Button
    Friend WithEvents btnCreateIndex As System.Windows.Forms.Button
    Friend WithEvents txtReplacementWord As System.Windows.Forms.TextBox
    Friend WithEvents lblReplacementWord As System.Windows.Forms.Label
    Friend WithEvents ckIndexWordsOnly As System.Windows.Forms.CheckBox
    Friend WithEvents saveFileDialog As System.Windows.Forms.SaveFileDialog
    Friend WithEvents btnLoadReplacementWords As System.Windows.Forms.Button
    Friend WithEvents openFileDialog As System.Windows.Forms.OpenFileDialog
    Friend WithEvents btnMarkIndexWords As System.Windows.Forms.Button
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents btnCreateIndexStyles As System.Windows.Forms.Button

End Class
