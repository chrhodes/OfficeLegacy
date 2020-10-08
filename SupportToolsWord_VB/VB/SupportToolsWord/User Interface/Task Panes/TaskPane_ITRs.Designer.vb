<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TaskPane_ITRs
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
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.btnGetITRInformation = New System.Windows.Forms.Button()
        Me.btnDisplayITRDetail = New System.Windows.Forms.Button()
        Me.gbITRWork = New System.Windows.Forms.GroupBox()
        Me.gbITRWork.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnGetITRInformation
        '
        Me.btnGetITRInformation.Location = New System.Drawing.Point(15, 48)
        Me.btnGetITRInformation.Name = "btnGetITRInformation"
        Me.btnGetITRInformation.Size = New System.Drawing.Size(155, 23)
        Me.btnGetITRInformation.TabIndex = 31
        Me.btnGetITRInformation.Text = "Get ITR Information"
        Me.btnGetITRInformation.UseVisualStyleBackColor = True
        '
        'btnDisplayITRDetail
        '
        Me.btnDisplayITRDetail.Location = New System.Drawing.Point(15, 77)
        Me.btnDisplayITRDetail.Name = "btnDisplayITRDetail"
        Me.btnDisplayITRDetail.Size = New System.Drawing.Size(155, 23)
        Me.btnDisplayITRDetail.TabIndex = 32
        Me.btnDisplayITRDetail.Text = "Display ITR Detail"
        Me.btnDisplayITRDetail.UseVisualStyleBackColor = True
        '
        'gbITRWork
        '
        Me.gbITRWork.Controls.Add(Me.btnDisplayITRDetail)
        Me.gbITRWork.Controls.Add(Me.btnGetITRInformation)
        Me.gbITRWork.Location = New System.Drawing.Point(13, 13)
        Me.gbITRWork.Name = "gbITRWork"
        Me.gbITRWork.Size = New System.Drawing.Size(182, 288)
        Me.gbITRWork.TabIndex = 22
        Me.gbITRWork.TabStop = False
        Me.gbITRWork.Text = "ITR Work"
        '
        'TaskPane_ITRs
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.gbITRWork)
        Me.Name = "TaskPane_ITRs"
        Me.Size = New System.Drawing.Size(211, 646)
        Me.gbITRWork.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents gbITRWork As System.Windows.Forms.GroupBox
    Friend WithEvents btnGetITRInformation As System.Windows.Forms.Button
    Friend WithEvents btnDisplayITRDetail As System.Windows.Forms.Button

End Class
