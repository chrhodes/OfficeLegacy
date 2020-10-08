<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TaskPane_Webs
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
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.btnGetAllSubWebCollection = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.btnGetWebCollection = New System.Windows.Forms.Button
        Me.txtURL = New System.Windows.Forms.TextBox
        Me.lblURL = New System.Windows.Forms.Label
        Me.cbOnTimeTeams = New System.Windows.Forms.ComboBox
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnGetAllSubWebCollection
        '
        Me.btnGetAllSubWebCollection.Location = New System.Drawing.Point(13, 52)
        Me.btnGetAllSubWebCollection.Name = "btnGetAllSubWebCollection"
        Me.btnGetAllSubWebCollection.Size = New System.Drawing.Size(262, 25)
        Me.btnGetAllSubWebCollection.TabIndex = 12
        Me.btnGetAllSubWebCollection.Text = "Get All Sub Web Collection"
        Me.btnGetAllSubWebCollection.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.btnGetWebCollection)
        Me.GroupBox2.Controls.Add(Me.txtURL)
        Me.GroupBox2.Controls.Add(Me.lblURL)
        Me.GroupBox2.Controls.Add(Me.btnGetAllSubWebCollection)
        Me.GroupBox2.Controls.Add(Me.cbOnTimeTeams)
        Me.GroupBox2.Controls.Add(Me.GroupBox3)
        Me.GroupBox2.Location = New System.Drawing.Point(8, 11)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(281, 403)
        Me.GroupBox2.TabIndex = 10
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Webs (Webs.asmx)"
        '
        'btnGetWebCollection
        '
        Me.btnGetWebCollection.Location = New System.Drawing.Point(13, 94)
        Me.btnGetWebCollection.Name = "btnGetWebCollection"
        Me.btnGetWebCollection.Size = New System.Drawing.Size(262, 25)
        Me.btnGetWebCollection.TabIndex = 21
        Me.btnGetWebCollection.Text = "Get Web Collection"
        Me.btnGetWebCollection.UseVisualStyleBackColor = True
        '
        'txtURL
        '
        Me.txtURL.Location = New System.Drawing.Point(43, 16)
        Me.txtURL.Name = "txtURL"
        Me.txtURL.Size = New System.Drawing.Size(232, 20)
        Me.txtURL.TabIndex = 20
        '
        'lblURL
        '
        Me.lblURL.AutoSize = True
        Me.lblURL.Location = New System.Drawing.Point(10, 19)
        Me.lblURL.Name = "lblURL"
        Me.lblURL.Size = New System.Drawing.Size(29, 13)
        Me.lblURL.TabIndex = 19
        Me.lblURL.Text = "URL"
        '
        'cbOnTimeTeams
        '
        Me.cbOnTimeTeams.FormattingEnabled = True
        Me.cbOnTimeTeams.Location = New System.Drawing.Point(13, 152)
        Me.cbOnTimeTeams.Name = "cbOnTimeTeams"
        Me.cbOnTimeTeams.Size = New System.Drawing.Size(153, 21)
        Me.cbOnTimeTeams.TabIndex = 18
        '
        'GroupBox3
        '
        Me.GroupBox3.Location = New System.Drawing.Point(6, 243)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(167, 154)
        Me.GroupBox3.TabIndex = 15
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Debug"
        '
        'TaskPane_Webs
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.GroupBox2)
        Me.Name = "TaskPane_Webs"
        Me.Size = New System.Drawing.Size(300, 500)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents btnGetAllSubWebCollection As System.Windows.Forms.Button
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents cbOnTimeTeams As System.Windows.Forms.ComboBox
    Friend WithEvents btnGetWebCollection As System.Windows.Forms.Button
    Friend WithEvents txtURL As System.Windows.Forms.TextBox
    Friend WithEvents lblURL As System.Windows.Forms.Label

End Class
