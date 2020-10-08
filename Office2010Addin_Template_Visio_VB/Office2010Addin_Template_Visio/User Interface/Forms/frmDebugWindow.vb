Public Class frmDebugWindow

    Private Sub btnClearOutput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearOutput.Click
        Me.txtOutput.Clear()
    End Sub

    Private Sub chkDebugSQL_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDebugSQL.CheckedChanged
        Common.DebugSQL = DirectCast(sender, System.Windows.Forms.CheckBox).Checked
    End Sub

    Private Sub chkDebugLevel1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDebugLevel1.CheckedChanged
        Common.DebugLevel1 = DirectCast(sender, System.Windows.Forms.CheckBox).Checked
    End Sub

    Private Sub chkDebugLevel2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDebugLevel2.CheckedChanged
        Common.DebugLevel2 = DirectCast(sender, System.Windows.Forms.CheckBox).Checked
    End Sub

    Private Sub frmDebugWindow_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
        Common.DeveloperMode = False
        e.Cancel = True
    End Sub
End Class