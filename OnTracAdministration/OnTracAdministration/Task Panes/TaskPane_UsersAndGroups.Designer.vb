<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TaskPane_UsersAndGroups
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
        Me.btnGetUserCollectionFromSite = New System.Windows.Forms.Button
        Me.btnGetGroupCollectionFromSite = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.btnGetUserCollectionFromWeb = New System.Windows.Forms.Button
        Me.txtURL = New System.Windows.Forms.TextBox
        Me.lblURL = New System.Windows.Forms.Label
        Me.btnGetGroupCollectionFromWeb = New System.Windows.Forms.Button
        Me.cbOnTimeTeams = New System.Windows.Forms.ComboBox
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnGetUserCollectionFromSite
        '
        Me.btnGetUserCollectionFromSite.Location = New System.Drawing.Point(13, 156)
        Me.btnGetUserCollectionFromSite.Name = "btnGetUserCollectionFromSite"
        Me.btnGetUserCollectionFromSite.Size = New System.Drawing.Size(252, 25)
        Me.btnGetUserCollectionFromSite.TabIndex = 19
        Me.btnGetUserCollectionFromSite.Text = "GetUserCollectionFromSite()"
        Me.ToolTip1.SetToolTip(Me.btnGetUserCollectionFromSite, "Validate selected file(s) contain valid On-Time data")
        Me.btnGetUserCollectionFromSite.UseVisualStyleBackColor = True
        '
        'btnGetGroupCollectionFromSite
        '
        Me.btnGetGroupCollectionFromSite.Location = New System.Drawing.Point(13, 48)
        Me.btnGetGroupCollectionFromSite.Name = "btnGetGroupCollectionFromSite"
        Me.btnGetGroupCollectionFromSite.Size = New System.Drawing.Size(252, 25)
        Me.btnGetGroupCollectionFromSite.TabIndex = 12
        Me.btnGetGroupCollectionFromSite.Text = "GetGroupCollectionFromSite()"
        Me.btnGetGroupCollectionFromSite.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.btnGetUserCollectionFromWeb)
        Me.GroupBox2.Controls.Add(Me.txtURL)
        Me.GroupBox2.Controls.Add(Me.lblURL)
        Me.GroupBox2.Controls.Add(Me.btnGetGroupCollectionFromWeb)
        Me.GroupBox2.Controls.Add(Me.btnGetUserCollectionFromSite)
        Me.GroupBox2.Controls.Add(Me.btnGetGroupCollectionFromSite)
        Me.GroupBox2.Controls.Add(Me.cbOnTimeTeams)
        Me.GroupBox2.Controls.Add(Me.GroupBox3)
        Me.GroupBox2.Location = New System.Drawing.Point(8, 11)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(271, 457)
        Me.GroupBox2.TabIndex = 10
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Users and Groups (UserGroup.asmx)"
        '
        'btnGetUserCollectionFromWeb
        '
        Me.btnGetUserCollectionFromWeb.Location = New System.Drawing.Point(13, 187)
        Me.btnGetUserCollectionFromWeb.Name = "btnGetUserCollectionFromWeb"
        Me.btnGetUserCollectionFromWeb.Size = New System.Drawing.Size(252, 25)
        Me.btnGetUserCollectionFromWeb.TabIndex = 23
        Me.btnGetUserCollectionFromWeb.Text = "GetUserCollectionFromWeb()"
        Me.btnGetUserCollectionFromWeb.UseVisualStyleBackColor = True
        '
        'txtURL
        '
        Me.txtURL.Location = New System.Drawing.Point(46, 19)
        Me.txtURL.Name = "txtURL"
        Me.txtURL.Size = New System.Drawing.Size(219, 20)
        Me.txtURL.TabIndex = 22
        '
        'lblURL
        '
        Me.lblURL.AutoSize = True
        Me.lblURL.Location = New System.Drawing.Point(13, 22)
        Me.lblURL.Name = "lblURL"
        Me.lblURL.Size = New System.Drawing.Size(29, 13)
        Me.lblURL.TabIndex = 21
        Me.lblURL.Text = "URL"
        '
        'btnGetGroupCollectionFromWeb
        '
        Me.btnGetGroupCollectionFromWeb.Location = New System.Drawing.Point(13, 79)
        Me.btnGetGroupCollectionFromWeb.Name = "btnGetGroupCollectionFromWeb"
        Me.btnGetGroupCollectionFromWeb.Size = New System.Drawing.Size(252, 25)
        Me.btnGetGroupCollectionFromWeb.TabIndex = 20
        Me.btnGetGroupCollectionFromWeb.Text = "GetGroupCollectionFromWeb()"
        Me.btnGetGroupCollectionFromWeb.UseVisualStyleBackColor = True
        '
        'cbOnTimeTeams
        '
        Me.cbOnTimeTeams.FormattingEnabled = True
        Me.cbOnTimeTeams.Location = New System.Drawing.Point(13, 250)
        Me.cbOnTimeTeams.Name = "cbOnTimeTeams"
        Me.cbOnTimeTeams.Size = New System.Drawing.Size(153, 21)
        Me.cbOnTimeTeams.TabIndex = 18
        '
        'GroupBox3
        '
        Me.GroupBox3.Location = New System.Drawing.Point(6, 308)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(259, 89)
        Me.GroupBox3.TabIndex = 15
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Debug"
        '
        'TaskPane_UsersAndGroups
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.GroupBox2)
        Me.Name = "TaskPane_UsersAndGroups"
        Me.Size = New System.Drawing.Size(300, 500)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents btnGetGroupCollectionFromSite As System.Windows.Forms.Button
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents cbOnTimeTeams As System.Windows.Forms.ComboBox
    Friend WithEvents btnGetUserCollectionFromSite As System.Windows.Forms.Button
    Friend WithEvents btnGetGroupCollectionFromWeb As System.Windows.Forms.Button
    Friend WithEvents txtURL As System.Windows.Forms.TextBox
    Friend WithEvents lblURL As System.Windows.Forms.Label
    Friend WithEvents btnGetUserCollectionFromWeb As System.Windows.Forms.Button

End Class
