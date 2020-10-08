<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDebugWindow
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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
        Me.txtOutput = New System.Windows.Forms.TextBox()
        Me.chkDebugLevel2 = New System.Windows.Forms.CheckBox()
        Me.chkDebugLevel1 = New System.Windows.Forms.CheckBox()
        Me.chkDebugSQL = New System.Windows.Forms.CheckBox()
        Me.btnClearOutput = New System.Windows.Forms.Button()
        Me.gbDebugOptions = New System.Windows.Forms.GroupBox()
        Me.gbDebugOptions.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtOutput
        '
        Me.txtOutput.Location = New System.Drawing.Point(196, 12)
        Me.txtOutput.Multiline = True
        Me.txtOutput.Name = "txtOutput"
        Me.txtOutput.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtOutput.Size = New System.Drawing.Size(576, 538)
        Me.txtOutput.TabIndex = 0
        '
        'chkDebugLevel2
        '
        Me.chkDebugLevel2.AutoSize = True
        Me.chkDebugLevel2.Location = New System.Drawing.Point(16, 71)
        Me.chkDebugLevel2.Name = "chkDebugLevel2"
        Me.chkDebugLevel2.Size = New System.Drawing.Size(96, 17)
        Me.chkDebugLevel2.TabIndex = 1
        Me.chkDebugLevel2.Text = "Debug Level 2"
        Me.chkDebugLevel2.UseVisualStyleBackColor = True
        '
        'chkDebugLevel1
        '
        Me.chkDebugLevel1.AutoSize = True
        Me.chkDebugLevel1.Location = New System.Drawing.Point(16, 48)
        Me.chkDebugLevel1.Name = "chkDebugLevel1"
        Me.chkDebugLevel1.Size = New System.Drawing.Size(96, 17)
        Me.chkDebugLevel1.TabIndex = 2
        Me.chkDebugLevel1.Text = "Debug Level 1"
        Me.chkDebugLevel1.UseVisualStyleBackColor = True
        '
        'chkDebugSQL
        '
        Me.chkDebugSQL.AutoSize = True
        Me.chkDebugSQL.Location = New System.Drawing.Point(16, 25)
        Me.chkDebugSQL.Name = "chkDebugSQL"
        Me.chkDebugSQL.Size = New System.Drawing.Size(82, 17)
        Me.chkDebugSQL.TabIndex = 3
        Me.chkDebugSQL.Text = "Debug SQL"
        Me.chkDebugSQL.UseVisualStyleBackColor = True
        '
        'btnClearOutput
        '
        Me.btnClearOutput.Location = New System.Drawing.Point(12, 12)
        Me.btnClearOutput.Name = "btnClearOutput"
        Me.btnClearOutput.Size = New System.Drawing.Size(178, 23)
        Me.btnClearOutput.TabIndex = 4
        Me.btnClearOutput.Text = "Clear Output"
        Me.btnClearOutput.UseVisualStyleBackColor = True
        '
        'gbDebugOptions
        '
        Me.gbDebugOptions.Controls.Add(Me.chkDebugLevel1)
        Me.gbDebugOptions.Controls.Add(Me.chkDebugLevel2)
        Me.gbDebugOptions.Controls.Add(Me.chkDebugSQL)
        Me.gbDebugOptions.Location = New System.Drawing.Point(12, 53)
        Me.gbDebugOptions.Name = "gbDebugOptions"
        Me.gbDebugOptions.Size = New System.Drawing.Size(178, 130)
        Me.gbDebugOptions.TabIndex = 5
        Me.gbDebugOptions.TabStop = False
        Me.gbDebugOptions.Text = "Debug Options"
        '
        'frmDebugWindow
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(784, 562)
        Me.Controls.Add(Me.gbDebugOptions)
        Me.Controls.Add(Me.btnClearOutput)
        Me.Controls.Add(Me.txtOutput)
        Me.Name = "frmDebugWindow"
        Me.Text = "frmDebugWindow"
        Me.gbDebugOptions.ResumeLayout(False)
        Me.gbDebugOptions.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtOutput As System.Windows.Forms.TextBox
    Friend WithEvents chkDebugLevel2 As System.Windows.Forms.CheckBox
    Friend WithEvents chkDebugLevel1 As System.Windows.Forms.CheckBox
    Friend WithEvents chkDebugSQL As System.Windows.Forms.CheckBox
    Friend WithEvents btnClearOutput As System.Windows.Forms.Button
    Friend WithEvents gbDebugOptions As System.Windows.Forms.GroupBox
End Class
