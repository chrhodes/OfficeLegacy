Public Class frmITRDetail

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        System.Diagnostics.Process.Start(String.Format("http://lifedart/default.asp?type=incident&name={0}", txtID.Text))
    End Sub
End Class