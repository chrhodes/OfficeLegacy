﻿
Public Class ThisWorkbook

    Private Sub ThisWorkbook_Startup(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Startup
        ' Many routines in Util-Excel require access to the current application context.
        Globals.ExcelApp = ThisApplication
    End Sub

    Private Sub ThisWorkbook_Shutdown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shutdown

    End Sub

End Class
