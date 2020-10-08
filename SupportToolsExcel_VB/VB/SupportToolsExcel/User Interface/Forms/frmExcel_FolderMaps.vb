Option Strict Off

Imports System.Windows.Forms

Public Class frmExcel_FolderMaps
    Inherits System.Windows.Forms.Form

    Public Class Regex
        ' SharePoint Folder/File/Document Libraries may not contain any of the following characters
        '   / \ : * ? " < > | <TAB> { } % ~ &
        ' nor may they end in periods or contain embedded double periods.
        ' The following regular expressions capture these rules.
        Public Const cIllegalFileCharacters As String = "[/\\:\*\?""<>\|#\{}%~&]|\.\."  ' SharePoint disallowed
        Public Const cIllegalFolderCharacters As String = "[:\*\?""<>\|#\{}%~&]"   ' SharePoint disallowed
    End Class

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents gbContents As System.Windows.Forms.GroupBox
    Friend WithEvents chkLimitLevels As System.Windows.Forms.CheckBox
    Friend WithEvents chkGroupResults As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowFiles As System.Windows.Forms.CheckBox
    Friend WithEvents txtGroupLevel As System.Windows.Forms.TextBox
    Friend WithEvents txtLimitLevel As System.Windows.Forms.TextBox
    Friend WithEvents spnGroupLevel As System.Windows.Forms.VScrollBar
    Friend WithEvents spnLimitLevel As System.Windows.Forms.VScrollBar
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdCreateFolderMap As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtMonthsSinceAccessed As System.Windows.Forms.TextBox
    Friend WithEvents txtMonthsSinceWritten As System.Windows.Forms.TextBox
    Friend WithEvents spnMonthsSinceAccessed As System.Windows.Forms.VScrollBar
    Friend WithEvents spnMonthsSinceWritten As System.Windows.Forms.VScrollBar
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents chkColorCodeDates As System.Windows.Forms.CheckBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtMonthsSinceCreated As System.Windows.Forms.TextBox
    Friend WithEvents spnMonthsSinceCreated As System.Windows.Forms.VScrollBar
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtIllegalFileCharacters As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents chkIllegalCharacters As System.Windows.Forms.CheckBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtIllegalFolderCharacters As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents chkFileNameLength As System.Windows.Forms.CheckBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents txtMaxFileNameLength As System.Windows.Forms.TextBox
    Friend WithEvents ColorDialog1 As System.Windows.Forms.ColorDialog
    Friend WithEvents pnlDefaultColor As System.Windows.Forms.Panel
    Friend WithEvents pnlPathTooLongColor As System.Windows.Forms.Panel
    Friend WithEvents pnlMonthAccessedColor As System.Windows.Forms.Panel
    Friend WithEvents pnlMonthWrittenColor As System.Windows.Forms.Panel
    Friend WithEvents pnlMonthCreatedColor As System.Windows.Forms.Panel
    Friend WithEvents pnlIllegalFileNameLengthColor As System.Windows.Forms.Panel
    Friend WithEvents pnlIllegalCharactersColor As System.Windows.Forms.Panel
    Friend WithEvents chkShowFolders As System.Windows.Forms.CheckBox
    Friend WithEvents pnlNoAccessColor As System.Windows.Forms.Panel
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents gbStartingFolder As System.Windows.Forms.GroupBox
    Friend WithEvents btnSelectFolder As System.Windows.Forms.Button
    Friend WithEvents txtStartingFolder As System.Windows.Forms.TextBox
    Friend WithEvents chkPatternMatchFileOutput As System.Windows.Forms.CheckBox
    Friend WithEvents txtFileMatchPattern As System.Windows.Forms.TextBox
    Friend WithEvents ToolTips As System.Windows.Forms.ToolTip
    Friend WithEvents chkSkipFoldersWithNoFiles As System.Windows.Forms.CheckBox
    Friend WithEvents pnlFolderHighlightColor As System.Windows.Forms.Panel
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents txtFolderMatchPattern As System.Windows.Forms.TextBox
    Friend WithEvents chkPatternMatchFolderHighlight As System.Windows.Forms.CheckBox
    Friend WithEvents pnlPatternMatchFileColor As System.Windows.Forms.Panel
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.gbContents = New System.Windows.Forms.GroupBox
        Me.pnlPatternMatchFileColor = New System.Windows.Forms.Panel
        Me.Label16 = New System.Windows.Forms.Label
        Me.pnlFolderHighlightColor = New System.Windows.Forms.Panel
        Me.Label15 = New System.Windows.Forms.Label
        Me.txtFolderMatchPattern = New System.Windows.Forms.TextBox
        Me.chkPatternMatchFolderHighlight = New System.Windows.Forms.CheckBox
        Me.chkSkipFoldersWithNoFiles = New System.Windows.Forms.CheckBox
        Me.chkPatternMatchFileOutput = New System.Windows.Forms.CheckBox
        Me.txtFileMatchPattern = New System.Windows.Forms.TextBox
        Me.pnlNoAccessColor = New System.Windows.Forms.Panel
        Me.Label14 = New System.Windows.Forms.Label
        Me.chkShowFolders = New System.Windows.Forms.CheckBox
        Me.pnlPathTooLongColor = New System.Windows.Forms.Panel
        Me.spnLimitLevel = New System.Windows.Forms.VScrollBar
        Me.spnGroupLevel = New System.Windows.Forms.VScrollBar
        Me.chkShowFiles = New System.Windows.Forms.CheckBox
        Me.chkGroupResults = New System.Windows.Forms.CheckBox
        Me.chkLimitLevels = New System.Windows.Forms.CheckBox
        Me.txtLimitLevel = New System.Windows.Forms.TextBox
        Me.txtGroupLevel = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.cmdCreateFolderMap = New System.Windows.Forms.Button
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.pnlMonthAccessedColor = New System.Windows.Forms.Panel
        Me.pnlMonthWrittenColor = New System.Windows.Forms.Panel
        Me.pnlMonthCreatedColor = New System.Windows.Forms.Panel
        Me.pnlDefaultColor = New System.Windows.Forms.Panel
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtMonthsSinceCreated = New System.Windows.Forms.TextBox
        Me.spnMonthsSinceCreated = New System.Windows.Forms.VScrollBar
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtMonthsSinceAccessed = New System.Windows.Forms.TextBox
        Me.txtMonthsSinceWritten = New System.Windows.Forms.TextBox
        Me.spnMonthsSinceAccessed = New System.Windows.Forms.VScrollBar
        Me.spnMonthsSinceWritten = New System.Windows.Forms.VScrollBar
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.chkColorCodeDates = New System.Windows.Forms.CheckBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.pnlIllegalFileNameLengthColor = New System.Windows.Forms.Panel
        Me.pnlIllegalCharactersColor = New System.Windows.Forms.Panel
        Me.Label13 = New System.Windows.Forms.Label
        Me.chkFileNameLength = New System.Windows.Forms.CheckBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.txtMaxFileNameLength = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.txtIllegalFolderCharacters = New System.Windows.Forms.TextBox
        Me.chkIllegalCharacters = New System.Windows.Forms.CheckBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.txtIllegalFileCharacters = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.ColorDialog1 = New System.Windows.Forms.ColorDialog
        Me.gbStartingFolder = New System.Windows.Forms.GroupBox
        Me.btnSelectFolder = New System.Windows.Forms.Button
        Me.txtStartingFolder = New System.Windows.Forms.TextBox
        Me.ToolTips = New System.Windows.Forms.ToolTip(Me.components)
        Me.gbContents.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.gbStartingFolder.SuspendLayout()
        Me.SuspendLayout()
        '
        'gbContents
        '
        Me.gbContents.Controls.Add(Me.pnlPatternMatchFileColor)
        Me.gbContents.Controls.Add(Me.Label16)
        Me.gbContents.Controls.Add(Me.pnlFolderHighlightColor)
        Me.gbContents.Controls.Add(Me.Label15)
        Me.gbContents.Controls.Add(Me.txtFolderMatchPattern)
        Me.gbContents.Controls.Add(Me.chkPatternMatchFolderHighlight)
        Me.gbContents.Controls.Add(Me.chkSkipFoldersWithNoFiles)
        Me.gbContents.Controls.Add(Me.chkPatternMatchFileOutput)
        Me.gbContents.Controls.Add(Me.txtFileMatchPattern)
        Me.gbContents.Controls.Add(Me.pnlNoAccessColor)
        Me.gbContents.Controls.Add(Me.Label14)
        Me.gbContents.Controls.Add(Me.chkShowFolders)
        Me.gbContents.Controls.Add(Me.pnlPathTooLongColor)
        Me.gbContents.Controls.Add(Me.spnLimitLevel)
        Me.gbContents.Controls.Add(Me.spnGroupLevel)
        Me.gbContents.Controls.Add(Me.chkShowFiles)
        Me.gbContents.Controls.Add(Me.chkGroupResults)
        Me.gbContents.Controls.Add(Me.chkLimitLevels)
        Me.gbContents.Controls.Add(Me.txtLimitLevel)
        Me.gbContents.Controls.Add(Me.txtGroupLevel)
        Me.gbContents.Controls.Add(Me.Label8)
        Me.gbContents.Location = New System.Drawing.Point(8, 59)
        Me.gbContents.Name = "gbContents"
        Me.gbContents.Size = New System.Drawing.Size(362, 217)
        Me.gbContents.TabIndex = 0
        Me.gbContents.TabStop = False
        Me.gbContents.Text = "Contents"
        '
        'pnlPatternMatchFileColor
        '
        Me.pnlPatternMatchFileColor.BackColor = System.Drawing.Color.Lime
        Me.pnlPatternMatchFileColor.Location = New System.Drawing.Point(326, 164)
        Me.pnlPatternMatchFileColor.Name = "pnlPatternMatchFileColor"
        Me.pnlPatternMatchFileColor.Size = New System.Drawing.Size(25, 17)
        Me.pnlPatternMatchFileColor.TabIndex = 29
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(210, 167)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(107, 13)
        Me.Label16.TabIndex = 30
        Me.Label16.Text = "Folder Highlight Color"
        Me.Label16.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlFolderHighlightColor
        '
        Me.pnlFolderHighlightColor.BackColor = System.Drawing.Color.Lime
        Me.pnlFolderHighlightColor.Location = New System.Drawing.Point(326, 104)
        Me.pnlFolderHighlightColor.Name = "pnlFolderHighlightColor"
        Me.pnlFolderHighlightColor.Size = New System.Drawing.Size(25, 17)
        Me.pnlFolderHighlightColor.TabIndex = 27
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(210, 107)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(107, 13)
        Me.Label15.TabIndex = 28
        Me.Label15.Text = "Folder Highlight Color"
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtFolderMatchPattern
        '
        Me.txtFolderMatchPattern.Location = New System.Drawing.Point(6, 129)
        Me.txtFolderMatchPattern.Name = "txtFolderMatchPattern"
        Me.txtFolderMatchPattern.Size = New System.Drawing.Size(345, 20)
        Me.txtFolderMatchPattern.TabIndex = 26
        Me.ToolTips.SetToolTip(Me.txtFolderMatchPattern, "Regular Expression to match folders")
        '
        'chkPatternMatchFolderHighlight
        '
        Me.chkPatternMatchFolderHighlight.Location = New System.Drawing.Point(8, 102)
        Me.chkPatternMatchFolderHighlight.Name = "chkPatternMatchFolderHighlight"
        Me.chkPatternMatchFolderHighlight.Size = New System.Drawing.Size(210, 24)
        Me.chkPatternMatchFolderHighlight.TabIndex = 25
        Me.chkPatternMatchFolderHighlight.Text = "RegEx Pattern Match Folder Highlight"
        '
        'chkSkipFoldersWithNoFiles
        '
        Me.chkSkipFoldersWithNoFiles.Location = New System.Drawing.Point(191, 72)
        Me.chkSkipFoldersWithNoFiles.Name = "chkSkipFoldersWithNoFiles"
        Me.chkSkipFoldersWithNoFiles.Size = New System.Drawing.Size(162, 24)
        Me.chkSkipFoldersWithNoFiles.TabIndex = 24
        Me.chkSkipFoldersWithNoFiles.Text = "Skip Folders with No Files"
        '
        'chkPatternMatchFileOutput
        '
        Me.chkPatternMatchFileOutput.Enabled = False
        Me.chkPatternMatchFileOutput.Location = New System.Drawing.Point(8, 162)
        Me.chkPatternMatchFileOutput.Name = "chkPatternMatchFileOutput"
        Me.chkPatternMatchFileOutput.Size = New System.Drawing.Size(195, 24)
        Me.chkPatternMatchFileOutput.TabIndex = 23
        Me.chkPatternMatchFileOutput.Text = "RegEx Pattern Match File Output"
        '
        'txtFileMatchPattern
        '
        Me.txtFileMatchPattern.Enabled = False
        Me.txtFileMatchPattern.Location = New System.Drawing.Point(8, 189)
        Me.txtFileMatchPattern.Name = "txtFileMatchPattern"
        Me.txtFileMatchPattern.Size = New System.Drawing.Size(343, 20)
        Me.txtFileMatchPattern.TabIndex = 2
        Me.ToolTips.SetToolTip(Me.txtFileMatchPattern, "Regular Expression to match files")
        '
        'pnlNoAccessColor
        '
        Me.pnlNoAccessColor.BackColor = System.Drawing.Color.Violet
        Me.pnlNoAccessColor.Location = New System.Drawing.Point(101, 14)
        Me.pnlNoAccessColor.Name = "pnlNoAccessColor"
        Me.pnlNoAccessColor.Size = New System.Drawing.Size(25, 17)
        Me.pnlNoAccessColor.TabIndex = 20
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(9, 16)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(83, 13)
        Me.Label14.TabIndex = 21
        Me.Label14.Text = "NoAccess Color"
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'chkShowFolders
        '
        Me.chkShowFolders.Checked = True
        Me.chkShowFolders.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowFolders.Location = New System.Drawing.Point(8, 72)
        Me.chkShowFolders.Name = "chkShowFolders"
        Me.chkShowFolders.Size = New System.Drawing.Size(93, 24)
        Me.chkShowFolders.TabIndex = 20
        Me.chkShowFolders.Text = "Show Folders"
        '
        'pnlPathTooLongColor
        '
        Me.pnlPathTooLongColor.BackColor = System.Drawing.Color.Cyan
        Me.pnlPathTooLongColor.Location = New System.Drawing.Point(251, 12)
        Me.pnlPathTooLongColor.Name = "pnlPathTooLongColor"
        Me.pnlPathTooLongColor.Size = New System.Drawing.Size(25, 17)
        Me.pnlPathTooLongColor.TabIndex = 19
        '
        'spnLimitLevel
        '
        Me.spnLimitLevel.Location = New System.Drawing.Point(139, 43)
        Me.spnLimitLevel.Name = "spnLimitLevel"
        Me.spnLimitLevel.Size = New System.Drawing.Size(16, 20)
        Me.spnLimitLevel.TabIndex = 5
        '
        'spnGroupLevel
        '
        Me.spnGroupLevel.Location = New System.Drawing.Point(330, 45)
        Me.spnGroupLevel.Name = "spnGroupLevel"
        Me.spnGroupLevel.Size = New System.Drawing.Size(16, 20)
        Me.spnGroupLevel.TabIndex = 4
        '
        'chkShowFiles
        '
        Me.chkShowFiles.Location = New System.Drawing.Point(107, 72)
        Me.chkShowFiles.Name = "chkShowFiles"
        Me.chkShowFiles.Size = New System.Drawing.Size(104, 24)
        Me.chkShowFiles.TabIndex = 3
        Me.chkShowFiles.Text = "Show Files"
        Me.ToolTips.SetToolTip(Me.chkShowFiles, "Output contains files")
        '
        'chkGroupResults
        '
        Me.chkGroupResults.Location = New System.Drawing.Point(191, 41)
        Me.chkGroupResults.Name = "chkGroupResults"
        Me.chkGroupResults.Size = New System.Drawing.Size(104, 24)
        Me.chkGroupResults.TabIndex = 2
        Me.chkGroupResults.Text = "Group Results"
        '
        'chkLimitLevels
        '
        Me.chkLimitLevels.Location = New System.Drawing.Point(8, 41)
        Me.chkLimitLevels.Name = "chkLimitLevels"
        Me.chkLimitLevels.Size = New System.Drawing.Size(89, 24)
        Me.chkLimitLevels.TabIndex = 1
        Me.chkLimitLevels.Text = "Limit Levels"
        '
        'txtLimitLevel
        '
        Me.txtLimitLevel.Location = New System.Drawing.Point(107, 43)
        Me.txtLimitLevel.Name = "txtLimitLevel"
        Me.txtLimitLevel.Size = New System.Drawing.Size(24, 20)
        Me.txtLimitLevel.TabIndex = 2
        Me.txtLimitLevel.Text = "1"
        Me.txtLimitLevel.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtGroupLevel
        '
        Me.txtGroupLevel.Location = New System.Drawing.Point(298, 45)
        Me.txtGroupLevel.Name = "txtGroupLevel"
        Me.txtGroupLevel.Size = New System.Drawing.Size(24, 20)
        Me.txtGroupLevel.TabIndex = 1
        Me.txtGroupLevel.Text = "1"
        Me.txtGroupLevel.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(143, 16)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(99, 13)
        Me.Label8.TabIndex = 8
        Me.Label8.Text = "PathTooLong Color"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdCancel
        '
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Location = New System.Drawing.Point(186, 590)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(64, 23)
        Me.cmdCancel.TabIndex = 6
        Me.cmdCancel.Text = "Cancel"
        '
        'cmdCreateFolderMap
        '
        Me.cmdCreateFolderMap.Location = New System.Drawing.Point(256, 590)
        Me.cmdCreateFolderMap.Name = "cmdCreateFolderMap"
        Me.cmdCreateFolderMap.Size = New System.Drawing.Size(112, 23)
        Me.cmdCreateFolderMap.TabIndex = 7
        Me.cmdCreateFolderMap.Text = "Create Folder Map"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.pnlMonthAccessedColor)
        Me.GroupBox1.Controls.Add(Me.pnlMonthWrittenColor)
        Me.GroupBox1.Controls.Add(Me.pnlMonthCreatedColor)
        Me.GroupBox1.Controls.Add(Me.pnlDefaultColor)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.txtMonthsSinceCreated)
        Me.GroupBox1.Controls.Add(Me.spnMonthsSinceCreated)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.txtMonthsSinceAccessed)
        Me.GroupBox1.Controls.Add(Me.txtMonthsSinceWritten)
        Me.GroupBox1.Controls.Add(Me.spnMonthsSinceAccessed)
        Me.GroupBox1.Controls.Add(Me.spnMonthsSinceWritten)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.chkColorCodeDates)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 282)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(362, 145)
        Me.GroupBox1.TabIndex = 8
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Date Information"
        '
        'pnlMonthAccessedColor
        '
        Me.pnlMonthAccessedColor.BackColor = System.Drawing.Color.Blue
        Me.pnlMonthAccessedColor.Location = New System.Drawing.Point(326, 114)
        Me.pnlMonthAccessedColor.Name = "pnlMonthAccessedColor"
        Me.pnlMonthAccessedColor.Size = New System.Drawing.Size(25, 17)
        Me.pnlMonthAccessedColor.TabIndex = 19
        '
        'pnlMonthWrittenColor
        '
        Me.pnlMonthWrittenColor.BackColor = System.Drawing.Color.Green
        Me.pnlMonthWrittenColor.Location = New System.Drawing.Point(326, 78)
        Me.pnlMonthWrittenColor.Name = "pnlMonthWrittenColor"
        Me.pnlMonthWrittenColor.Size = New System.Drawing.Size(25, 17)
        Me.pnlMonthWrittenColor.TabIndex = 19
        '
        'pnlMonthCreatedColor
        '
        Me.pnlMonthCreatedColor.BackColor = System.Drawing.Color.Red
        Me.pnlMonthCreatedColor.Location = New System.Drawing.Point(326, 48)
        Me.pnlMonthCreatedColor.Name = "pnlMonthCreatedColor"
        Me.pnlMonthCreatedColor.Size = New System.Drawing.Size(25, 17)
        Me.pnlMonthCreatedColor.TabIndex = 19
        '
        'pnlDefaultColor
        '
        Me.pnlDefaultColor.BackColor = System.Drawing.Color.Black
        Me.pnlDefaultColor.Location = New System.Drawing.Point(326, 15)
        Me.pnlDefaultColor.Name = "pnlDefaultColor"
        Me.pnlDefaultColor.Size = New System.Drawing.Size(25, 17)
        Me.pnlDefaultColor.TabIndex = 18
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(15, 52)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(112, 13)
        Me.Label6.TabIndex = 17
        Me.Label6.Text = "Months Since Created"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtMonthsSinceCreated
        '
        Me.txtMonthsSinceCreated.Location = New System.Drawing.Point(137, 49)
        Me.txtMonthsSinceCreated.Name = "txtMonthsSinceCreated"
        Me.txtMonthsSinceCreated.Size = New System.Drawing.Size(24, 20)
        Me.txtMonthsSinceCreated.TabIndex = 16
        Me.txtMonthsSinceCreated.Text = "1"
        Me.txtMonthsSinceCreated.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'spnMonthsSinceCreated
        '
        Me.spnMonthsSinceCreated.Location = New System.Drawing.Point(169, 48)
        Me.spnMonthsSinceCreated.Name = "spnMonthsSinceCreated"
        Me.spnMonthsSinceCreated.Size = New System.Drawing.Size(16, 21)
        Me.spnMonthsSinceCreated.TabIndex = 15
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(246, 52)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(71, 13)
        Me.Label7.TabIndex = 14
        Me.Label7.Text = "Created Color"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(15, 85)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(109, 13)
        Me.Label5.TabIndex = 12
        Me.Label5.Text = "Months Since Written"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(9, 118)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(122, 13)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "Months Since Accessed"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtMonthsSinceAccessed
        '
        Me.txtMonthsSinceAccessed.Location = New System.Drawing.Point(137, 115)
        Me.txtMonthsSinceAccessed.Name = "txtMonthsSinceAccessed"
        Me.txtMonthsSinceAccessed.Size = New System.Drawing.Size(24, 20)
        Me.txtMonthsSinceAccessed.TabIndex = 10
        Me.txtMonthsSinceAccessed.Text = "1"
        Me.txtMonthsSinceAccessed.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtMonthsSinceWritten
        '
        Me.txtMonthsSinceWritten.Location = New System.Drawing.Point(137, 82)
        Me.txtMonthsSinceWritten.Name = "txtMonthsSinceWritten"
        Me.txtMonthsSinceWritten.Size = New System.Drawing.Size(24, 20)
        Me.txtMonthsSinceWritten.TabIndex = 9
        Me.txtMonthsSinceWritten.Text = "1"
        Me.txtMonthsSinceWritten.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'spnMonthsSinceAccessed
        '
        Me.spnMonthsSinceAccessed.Location = New System.Drawing.Point(169, 115)
        Me.spnMonthsSinceAccessed.Name = "spnMonthsSinceAccessed"
        Me.spnMonthsSinceAccessed.Size = New System.Drawing.Size(16, 20)
        Me.spnMonthsSinceAccessed.TabIndex = 8
        '
        'spnMonthsSinceWritten
        '
        Me.spnMonthsSinceWritten.Location = New System.Drawing.Point(169, 82)
        Me.spnMonthsSinceWritten.Name = "spnMonthsSinceWritten"
        Me.spnMonthsSinceWritten.Size = New System.Drawing.Size(16, 20)
        Me.spnMonthsSinceWritten.TabIndex = 7
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(236, 118)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(81, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Accessed Color"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(249, 82)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(68, 13)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Written Color"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(255, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(62, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Defalt Color"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'chkColorCodeDates
        '
        Me.chkColorCodeDates.AutoSize = True
        Me.chkColorCodeDates.Location = New System.Drawing.Point(8, 19)
        Me.chkColorCodeDates.Name = "chkColorCodeDates"
        Me.chkColorCodeDates.Size = New System.Drawing.Size(109, 17)
        Me.chkColorCodeDates.TabIndex = 0
        Me.chkColorCodeDates.Text = "Color Code Dates"
        Me.chkColorCodeDates.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.pnlIllegalFileNameLengthColor)
        Me.GroupBox2.Controls.Add(Me.pnlIllegalCharactersColor)
        Me.GroupBox2.Controls.Add(Me.Label13)
        Me.GroupBox2.Controls.Add(Me.chkFileNameLength)
        Me.GroupBox2.Controls.Add(Me.Label12)
        Me.GroupBox2.Controls.Add(Me.txtMaxFileNameLength)
        Me.GroupBox2.Controls.Add(Me.Label11)
        Me.GroupBox2.Controls.Add(Me.txtIllegalFolderCharacters)
        Me.GroupBox2.Controls.Add(Me.chkIllegalCharacters)
        Me.GroupBox2.Controls.Add(Me.Label10)
        Me.GroupBox2.Controls.Add(Me.txtIllegalFileCharacters)
        Me.GroupBox2.Controls.Add(Me.Label9)
        Me.GroupBox2.Location = New System.Drawing.Point(8, 432)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(362, 146)
        Me.GroupBox2.TabIndex = 9
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "SharePoint Information"
        '
        'pnlIllegalFileNameLengthColor
        '
        Me.pnlIllegalFileNameLengthColor.BackColor = System.Drawing.Color.Cyan
        Me.pnlIllegalFileNameLengthColor.Location = New System.Drawing.Point(326, 94)
        Me.pnlIllegalFileNameLengthColor.Name = "pnlIllegalFileNameLengthColor"
        Me.pnlIllegalFileNameLengthColor.Size = New System.Drawing.Size(25, 17)
        Me.pnlIllegalFileNameLengthColor.TabIndex = 19
        '
        'pnlIllegalCharactersColor
        '
        Me.pnlIllegalCharactersColor.BackColor = System.Drawing.Color.Orange
        Me.pnlIllegalCharactersColor.Location = New System.Drawing.Point(326, 14)
        Me.pnlIllegalCharactersColor.Name = "pnlIllegalCharactersColor"
        Me.pnlIllegalCharactersColor.Size = New System.Drawing.Size(25, 17)
        Me.pnlIllegalCharactersColor.TabIndex = 19
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(173, 98)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(144, 13)
        Me.Label13.TabIndex = 20
        Me.Label13.Text = "Illegal FileName Length Color"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'chkFileNameLength
        '
        Me.chkFileNameLength.AutoSize = True
        Me.chkFileNameLength.Location = New System.Drawing.Point(8, 97)
        Me.chkFileNameLength.Name = "chkFileNameLength"
        Me.chkFileNameLength.Size = New System.Drawing.Size(140, 17)
        Me.chkFileNameLength.TabIndex = 18
        Me.chkFileNameLength.Text = "Check FileName Length"
        Me.chkFileNameLength.UseVisualStyleBackColor = True
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(7, 123)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(137, 13)
        Me.Label12.TabIndex = 17
        Me.Label12.Text = "Maximum File Name Length"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtMaxFileNameLength
        '
        Me.txtMaxFileNameLength.Location = New System.Drawing.Point(155, 120)
        Me.txtMaxFileNameLength.Name = "txtMaxFileNameLength"
        Me.txtMaxFileNameLength.Size = New System.Drawing.Size(51, 20)
        Me.txtMaxFileNameLength.TabIndex = 16
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(24, 66)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(120, 13)
        Me.Label11.TabIndex = 15
        Me.Label11.Text = "Illegal Folder Characters"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtIllegalFolderCharacters
        '
        Me.txtIllegalFolderCharacters.Location = New System.Drawing.Point(155, 63)
        Me.txtIllegalFolderCharacters.Name = "txtIllegalFolderCharacters"
        Me.txtIllegalFolderCharacters.Size = New System.Drawing.Size(135, 20)
        Me.txtIllegalFolderCharacters.TabIndex = 14
        '
        'chkIllegalCharacters
        '
        Me.chkIllegalCharacters.AutoSize = True
        Me.chkIllegalCharacters.Location = New System.Drawing.Point(8, 17)
        Me.chkIllegalCharacters.Name = "chkIllegalCharacters"
        Me.chkIllegalCharacters.Size = New System.Drawing.Size(156, 17)
        Me.chkIllegalCharacters.TabIndex = 13
        Me.chkIllegalCharacters.Text = "Check for Illegal Characters"
        Me.chkIllegalCharacters.UseVisualStyleBackColor = True
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(37, 41)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(107, 13)
        Me.Label10.TabIndex = 12
        Me.Label10.Text = "Illegal File Characters"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtIllegalFileCharacters
        '
        Me.txtIllegalFileCharacters.Location = New System.Drawing.Point(155, 38)
        Me.txtIllegalFileCharacters.Name = "txtIllegalFileCharacters"
        Me.txtIllegalFileCharacters.Size = New System.Drawing.Size(135, 20)
        Me.txtIllegalFileCharacters.TabIndex = 11
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(202, 18)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(115, 13)
        Me.Label9.TabIndex = 10
        Me.Label9.Text = "Illegal Characters Color"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'gbStartingFolder
        '
        Me.gbStartingFolder.Controls.Add(Me.btnSelectFolder)
        Me.gbStartingFolder.Controls.Add(Me.txtStartingFolder)
        Me.gbStartingFolder.Location = New System.Drawing.Point(8, 4)
        Me.gbStartingFolder.Name = "gbStartingFolder"
        Me.gbStartingFolder.Size = New System.Drawing.Size(359, 49)
        Me.gbStartingFolder.TabIndex = 10
        Me.gbStartingFolder.TabStop = False
        Me.gbStartingFolder.Text = "Starting Folder"
        '
        'btnSelectFolder
        '
        Me.btnSelectFolder.Location = New System.Drawing.Point(298, 17)
        Me.btnSelectFolder.Name = "btnSelectFolder"
        Me.btnSelectFolder.Size = New System.Drawing.Size(53, 23)
        Me.btnSelectFolder.TabIndex = 1
        Me.btnSelectFolder.Text = "Pick"
        Me.btnSelectFolder.UseVisualStyleBackColor = True
        '
        'txtStartingFolder
        '
        Me.txtStartingFolder.Location = New System.Drawing.Point(8, 19)
        Me.txtStartingFolder.Name = "txtStartingFolder"
        Me.txtStartingFolder.Size = New System.Drawing.Size(278, 20)
        Me.txtStartingFolder.TabIndex = 0
        Me.ToolTips.SetToolTip(Me.txtStartingFolder, "Enter folder or click Pick.  Pick starts from what is entered.")
        '
        'Excel_FolderMaps_Form
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(375, 618)
        Me.Controls.Add(Me.gbStartingFolder)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.gbContents)
        Me.Controls.Add(Me.cmdCreateFolderMap)
        Me.Controls.Add(Me.cmdCancel)
        Me.Name = "Excel_FolderMaps_Form"
        Me.Text = "Excel_FolderMaps"
        Me.gbContents.ResumeLayout(False)
        Me.gbContents.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.gbStartingFolder.ResumeLayout(False)
        Me.gbStartingFolder.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private m_blnCancel As Boolean = False

    Private Sub chkGroupResults_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkGroupResults.Click
        If chkGroupResults.Checked Then
            Me.txtGroupLevel.Enabled = True
            Me.spnGroupLevel.Enabled = True
        Else
            Me.txtGroupLevel.Enabled = False
            Me.spnGroupLevel.Enabled = False
        End If
    End Sub

    Private Sub chkLimitLevels_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkLimitLevels.Click
        If chkLimitLevels.Checked Then
            Me.txtLimitLevel.Enabled = True
            Me.spnLimitLevel.Enabled = True
        Else
            Me.txtLimitLevel.Enabled = False
            Me.spnLimitLevel.Enabled = False
        End If
    End Sub

    Private Sub chkColorCodeDates_CheckedChange(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkColorCodeDates.CheckedChanged
        If chkColorCodeDates.Checked Then
            Me.pnlDefaultColor.Enabled = True
            Me.pnlMonthCreatedColor.Enabled = True
            Me.pnlMonthAccessedColor.Enabled = True
            Me.pnlMonthWrittenColor.Enabled = True

            Me.txtMonthsSinceCreated.Enabled = True
            Me.txtMonthsSinceWritten.Enabled = True
            Me.txtMonthsSinceAccessed.Enabled = True

            Me.spnMonthsSinceCreated.Enabled = True
            Me.spnMonthsSinceWritten.Enabled = True
            Me.spnMonthsSinceAccessed.Enabled = True
        Else
            Me.pnlDefaultColor.Enabled = False
            Me.pnlMonthCreatedColor.Enabled = False
            Me.pnlMonthAccessedColor.Enabled = False
            Me.pnlMonthWrittenColor.Enabled = False

            Me.txtMonthsSinceCreated.Enabled = False
            Me.txtMonthsSinceWritten.Enabled = False
            Me.txtMonthsSinceAccessed.Enabled = False

            Me.spnMonthsSinceCreated.Enabled = False
            Me.spnMonthsSinceWritten.Enabled = False
            Me.spnMonthsSinceAccessed.Enabled = False
        End If
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        'Me.Hide()
        m_blnCancel = True
        Me.Close()
    End Sub

    Private Sub cmdCreateFolderMap_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdCreateFolderMap.Click
        ' Just hide the form until we get the values off the controls.  See Excel_FolderMaps.
        ' This is probably not optimal.

        If Me.txtStartingFolder.Text.Length > 0 Then
            'Me.Hide()
            Me.DialogResult = Windows.Forms.DialogResult.OK
            Me.Close()
        Else
            MessageBox.Show("Must select starting folder")
            Me.txtStartingFolder.Focus()
            'Me.Show()
        End If
    End Sub

    Private Sub spnGroupLevel_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles spnGroupLevel.ValueChanged
        Me.txtGroupLevel.Text = Me.spnGroupLevel.Value
    End Sub

    Private Sub spnLimitLevel_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles spnLimitLevel.ValueChanged
        Me.txtLimitLevel.Text = Me.spnLimitLevel.Value
    End Sub

    Private Sub spnMonthsSinceCreated_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles spnMonthsSinceCreated.ValueChanged
        Me.txtMonthsSinceCreated.Text = Me.spnMonthsSinceCreated.Value
    End Sub

    Private Sub spnMonthsSinceUpdated_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles spnMonthsSinceWritten.ValueChanged
        Me.txtMonthsSinceWritten.Text = Me.spnMonthsSinceWritten.Value
    End Sub

    Private Sub spnMonthsSinceAccessed_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles spnMonthsSinceAccessed.ValueChanged
        Me.txtMonthsSinceAccessed.Text = Me.spnMonthsSinceAccessed.Value
    End Sub

    Private Sub Excel_FolderMaps_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        With Me.txtLimitLevel
            .Text = 1
            .Enabled = False
        End With

        With Me.spnLimitLevel
            .Minimum = 1
            .Value = 1
            .Enabled = False
        End With

        With Me.txtGroupLevel
            .Text = 3
            .Enabled = False
        End With

        With Me.spnGroupLevel
            .Minimum = 3
            .Value = 3
            .Enabled = False
        End With

        With Me.txtMonthsSinceCreated
            .Text = 24
            .Enabled = False
        End With

        With Me.txtMonthsSinceAccessed
            .Text = 24
            .Enabled = False
        End With

        With Me.txtMonthsSinceWritten
            .Text = 24
            .Enabled = False
        End With

        With Me.spnMonthsSinceCreated
            .Minimum = 1
            .Value = 24
            .Enabled = False
        End With

        With Me.spnMonthsSinceWritten
            .Minimum = 1
            .Value = 24
            .Enabled = False
        End With

        With Me.spnMonthsSinceAccessed
            .Minimum = 1
            .Value = 24
            .Enabled = False
        End With

        With Me.txtIllegalFileCharacters
            .Text = Regex.cIllegalFileCharacters
            .Enabled = False
        End With

        With Me.txtIllegalFolderCharacters
            .Text = Regex.cIllegalFolderCharacters
            .Enabled = False
        End With

        With Me.txtMaxFileNameLength
            .Text = Common.cMaxFileNameLength
            .Enabled = False
        End With

    End Sub

    Private Sub chkIllegalCharacters_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkIllegalCharacters.CheckedChanged
        If chkIllegalCharacters.Checked Then
            Me.pnlIllegalCharactersColor.Enabled = True
            Me.txtIllegalFileCharacters.Enabled = True
            Me.txtIllegalFolderCharacters.Enabled = True
        Else
            Me.pnlIllegalCharactersColor.Enabled = False
            Me.txtIllegalFileCharacters.Enabled = False
            Me.txtIllegalFolderCharacters.Enabled = False
        End If
    End Sub

    Private Sub chkFileNameLength_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkFileNameLength.CheckedChanged
        If chkFileNameLength.Checked Then
            Me.txtMaxFileNameLength.Enabled = True
            Me.pnlIllegalFileNameLengthColor.Enabled = True
        Else
            Me.txtMaxFileNameLength.Enabled = False
            Me.pnlIllegalFileNameLengthColor.Enabled = False
        End If
    End Sub

    Private Sub ColorBox_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) _
    Handles pnlDefaultColor.DoubleClick, pnlPatternMatchFileColor.DoubleClick, pnlFolderHighlightColor.DoubleClick, _
            pnlMonthAccessedColor.DoubleClick, pnlMonthCreatedColor.DoubleClick, pnlMonthWrittenColor.DoubleClick, _
            pnlIllegalCharactersColor.DoubleClick, pnlPathTooLongColor.DoubleClick, pnlIllegalFileNameLengthColor.DoubleClick, pnlNoAccessColor.DoubleClick
        ColorDialog1.Color = sender.BackColor
        Dim dlgResult As DialogResult = ColorDialog1.ShowDialog()
        If Not (dlgResult = Windows.Forms.DialogResult.Cancel) Then
            sender.BackColor = ColorDialog1.Color
        End If
    End Sub

    Private Sub chkShowFiles_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkShowFiles.Click
        If chkShowFiles.Checked Then
            'chkShowFolders.Checked = True   ' Must display folders if displaying files

            chkPatternMatchFileOutput.Enabled = True
            txtFileMatchPattern.Enabled = True
            chkSkipFoldersWithNoFiles.Enabled = True
        Else
            chkPatternMatchFileOutput.Enabled = False
            chkPatternMatchFileOutput.CheckState = CheckState.Unchecked
            txtFileMatchPattern.Enabled = False
            chkSkipFoldersWithNoFiles.Enabled = False
        End If
    End Sub

    Private Sub chkShowFolders_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkShowFolders.Click
        'If Not chkShowFolders.Checked Then
        '    chkShowFiles.Checked = False    ' Do not display files if not displaying folders
        'End If
    End Sub

    Private Sub btnSelectFolder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSelectFolder.Click
        Me.FolderBrowserDialog1.ShowNewFolderButton = False

        If txtStartingFolder.Text.Length > 0 Then
            Me.FolderBrowserDialog1.SelectedPath = txtStartingFolder.Text
        End If

        Me.FolderBrowserDialog1.ShowDialog()

        txtStartingFolder.Text = Me.FolderBrowserDialog1.SelectedPath()
    End Sub

    Private Sub txtFileMatchPattern_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFileMatchPattern.TextChanged

    End Sub
End Class
