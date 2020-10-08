Public Class TaskPane_CreateSheets

    Private Sub btnCreateWorksheet_X_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateWorksheet_X.Click
        XWorkSheet.CreateWorkSheet(Common.cSN_X)
    End Sub

    Private Sub btnCreateWorksheet_Y_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateWorksheet_Y.Click
        YWorkSheet.CreateWorkSheet(Common.cSN_Y)
    End Sub

    Private Sub btnCreateWorksheet_Z_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateWorksheet_Z.Click
        ZWorkSheet.CreateWorksheet(Common.cSN_Z)
    End Sub

End Class
