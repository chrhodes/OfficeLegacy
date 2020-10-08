<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TaskPane_UserAdmin
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
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.txtWebUserID = New System.Windows.Forms.TextBox
        Me.txtWebUserName = New System.Windows.Forms.TextBox
        Me.cbWebUsers = New System.Windows.Forms.ComboBox
        Me.btnGetWebUsers = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.txtSiteCollectionUserID = New System.Windows.Forms.TextBox
        Me.txtSiteCollectionUserName = New System.Windows.Forms.TextBox
        Me.cbSiteCollectionUsers = New System.Windows.Forms.ComboBox
        Me.btnGetSiteCollectionUsers = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtTitle = New System.Windows.Forms.TextBox
        Me.txtWebURL = New System.Windows.Forms.TextBox
        Me.btnGetAllSubWebs = New System.Windows.Forms.Button
        Me.cbWebs = New System.Windows.Forms.ComboBox
        Me.btnFindSitesWithUser = New System.Windows.Forms.Button
        Me.txtURL = New System.Windows.Forms.TextBox
        Me.lblURL = New System.Windows.Forms.Label
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.GroupBox3)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.txtTitle)
        Me.GroupBox1.Controls.Add(Me.txtWebURL)
        Me.GroupBox1.Controls.Add(Me.btnGetAllSubWebs)
        Me.GroupBox1.Controls.Add(Me.cbWebs)
        Me.GroupBox1.Controls.Add(Me.btnFindSitesWithUser)
        Me.GroupBox1.Controls.Add(Me.txtURL)
        Me.GroupBox1.Controls.Add(Me.lblURL)
        Me.GroupBox1.Location = New System.Drawing.Point(3, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(294, 438)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "User Admin"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.txtWebUserID)
        Me.GroupBox3.Controls.Add(Me.txtWebUserName)
        Me.GroupBox3.Controls.Add(Me.cbWebUsers)
        Me.GroupBox3.Controls.Add(Me.btnGetWebUsers)
        Me.GroupBox3.Location = New System.Drawing.Point(9, 254)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(278, 100)
        Me.GroupBox3.TabIndex = 35
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Web Users"
        '
        'txtWebUserID
        '
        Me.txtWebUserID.Location = New System.Drawing.Point(228, 71)
        Me.txtWebUserID.Name = "txtWebUserID"
        Me.txtWebUserID.Size = New System.Drawing.Size(42, 20)
        Me.txtWebUserID.TabIndex = 27
        '
        'txtWebUserName
        '
        Me.txtWebUserName.Location = New System.Drawing.Point(9, 71)
        Me.txtWebUserName.Name = "txtWebUserName"
        Me.txtWebUserName.Size = New System.Drawing.Size(211, 20)
        Me.txtWebUserName.TabIndex = 26
        '
        'cbWebUsers
        '
        Me.cbWebUsers.FormattingEnabled = True
        Me.cbWebUsers.Location = New System.Drawing.Point(9, 44)
        Me.cbWebUsers.Name = "cbWebUsers"
        Me.cbWebUsers.Size = New System.Drawing.Size(261, 21)
        Me.cbWebUsers.TabIndex = 2
        '
        'btnGetWebUsers
        '
        Me.btnGetWebUsers.Location = New System.Drawing.Point(9, 16)
        Me.btnGetWebUsers.Name = "btnGetWebUsers"
        Me.btnGetWebUsers.Size = New System.Drawing.Size(261, 23)
        Me.btnGetWebUsers.TabIndex = 1
        Me.btnGetWebUsers.Text = "Get Web Users"
        Me.btnGetWebUsers.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.txtSiteCollectionUserID)
        Me.GroupBox2.Controls.Add(Me.txtSiteCollectionUserName)
        Me.GroupBox2.Controls.Add(Me.cbSiteCollectionUsers)
        Me.GroupBox2.Controls.Add(Me.btnGetSiteCollectionUsers)
        Me.GroupBox2.Location = New System.Drawing.Point(9, 148)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(278, 100)
        Me.GroupBox2.TabIndex = 34
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Site Collection Users"
        '
        'txtSiteCollectionUserID
        '
        Me.txtSiteCollectionUserID.Location = New System.Drawing.Point(228, 71)
        Me.txtSiteCollectionUserID.Name = "txtSiteCollectionUserID"
        Me.txtSiteCollectionUserID.Size = New System.Drawing.Size(42, 20)
        Me.txtSiteCollectionUserID.TabIndex = 27
        '
        'txtSiteCollectionUserName
        '
        Me.txtSiteCollectionUserName.Location = New System.Drawing.Point(9, 71)
        Me.txtSiteCollectionUserName.Name = "txtSiteCollectionUserName"
        Me.txtSiteCollectionUserName.Size = New System.Drawing.Size(211, 20)
        Me.txtSiteCollectionUserName.TabIndex = 26
        '
        'cbSiteCollectionUsers
        '
        Me.cbSiteCollectionUsers.FormattingEnabled = True
        Me.cbSiteCollectionUsers.Location = New System.Drawing.Point(9, 44)
        Me.cbSiteCollectionUsers.Name = "cbSiteCollectionUsers"
        Me.cbSiteCollectionUsers.Size = New System.Drawing.Size(261, 21)
        Me.cbSiteCollectionUsers.TabIndex = 2
        '
        'btnGetSiteCollectionUsers
        '
        Me.btnGetSiteCollectionUsers.Location = New System.Drawing.Point(9, 16)
        Me.btnGetSiteCollectionUsers.Name = "btnGetSiteCollectionUsers"
        Me.btnGetSiteCollectionUsers.Size = New System.Drawing.Size(261, 23)
        Me.btnGetSiteCollectionUsers.TabIndex = 1
        Me.btnGetSiteCollectionUsers.Text = "Get Site Collection Users"
        Me.btnGetSiteCollectionUsers.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 124)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(27, 13)
        Me.Label2.TabIndex = 33
        Me.Label2.Text = "Title"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 97)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(29, 13)
        Me.Label1.TabIndex = 32
        Me.Label1.Text = "URL"
        '
        'txtTitle
        '
        Me.txtTitle.Location = New System.Drawing.Point(38, 121)
        Me.txtTitle.Name = "txtTitle"
        Me.txtTitle.Size = New System.Drawing.Size(250, 20)
        Me.txtTitle.TabIndex = 31
        '
        'txtWebURL
        '
        Me.txtWebURL.Location = New System.Drawing.Point(38, 94)
        Me.txtWebURL.Name = "txtWebURL"
        Me.txtWebURL.Size = New System.Drawing.Size(250, 20)
        Me.txtWebURL.TabIndex = 30
        '
        'btnGetAllSubWebs
        '
        Me.btnGetAllSubWebs.Location = New System.Drawing.Point(6, 40)
        Me.btnGetAllSubWebs.Name = "btnGetAllSubWebs"
        Me.btnGetAllSubWebs.Size = New System.Drawing.Size(282, 23)
        Me.btnGetAllSubWebs.TabIndex = 29
        Me.btnGetAllSubWebs.Text = "Get All Sub Webs"
        Me.btnGetAllSubWebs.UseVisualStyleBackColor = True
        '
        'cbWebs
        '
        Me.cbWebs.FormattingEnabled = True
        Me.cbWebs.Location = New System.Drawing.Point(6, 67)
        Me.cbWebs.Name = "cbWebs"
        Me.cbWebs.Size = New System.Drawing.Size(282, 21)
        Me.cbWebs.TabIndex = 28
        '
        'btnFindSitesWithUser
        '
        Me.btnFindSitesWithUser.Location = New System.Drawing.Point(6, 374)
        Me.btnFindSitesWithUser.Name = "btnFindSitesWithUser"
        Me.btnFindSitesWithUser.Size = New System.Drawing.Size(282, 23)
        Me.btnFindSitesWithUser.TabIndex = 25
        Me.btnFindSitesWithUser.Text = "Find Sites with User"
        Me.btnFindSitesWithUser.UseVisualStyleBackColor = True
        '
        'txtURL
        '
        Me.txtURL.Location = New System.Drawing.Point(48, 16)
        Me.txtURL.Name = "txtURL"
        Me.txtURL.Size = New System.Drawing.Size(240, 20)
        Me.txtURL.TabIndex = 24
        Me.txtURL.Text = "http://ontrac"
        '
        'lblURL
        '
        Me.lblURL.AutoSize = True
        Me.lblURL.Location = New System.Drawing.Point(15, 19)
        Me.lblURL.Name = "lblURL"
        Me.lblURL.Size = New System.Drawing.Size(29, 13)
        Me.lblURL.TabIndex = 23
        Me.lblURL.Text = "URL"
        '
        'TaskPane_UserAdmin
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "TaskPane_UserAdmin"
        Me.Size = New System.Drawing.Size(300, 500)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cbSiteCollectionUsers As System.Windows.Forms.ComboBox
    Friend WithEvents btnGetSiteCollectionUsers As System.Windows.Forms.Button
    Friend WithEvents txtURL As System.Windows.Forms.TextBox
    Friend WithEvents lblURL As System.Windows.Forms.Label
    Friend WithEvents btnFindSitesWithUser As System.Windows.Forms.Button
    Friend WithEvents txtSiteCollectionUserID As System.Windows.Forms.TextBox
    Friend WithEvents txtSiteCollectionUserName As System.Windows.Forms.TextBox
    Friend WithEvents cbWebs As System.Windows.Forms.ComboBox
    Friend WithEvents btnGetAllSubWebs As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtTitle As System.Windows.Forms.TextBox
    Friend WithEvents txtWebURL As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents txtWebUserID As System.Windows.Forms.TextBox
    Friend WithEvents txtWebUserName As System.Windows.Forms.TextBox
    Friend WithEvents cbWebUsers As System.Windows.Forms.ComboBox
    Friend WithEvents btnGetWebUsers As System.Windows.Forms.Button

End Class
