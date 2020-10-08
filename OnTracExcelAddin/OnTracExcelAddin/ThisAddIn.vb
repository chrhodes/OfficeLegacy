Imports PacificLife.Life

Public Class ThisAddIn
    Private _CmdBars As CmdBars

    Private Sub ThisAddIn_Startup(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Startup
        'PLLog.Trace("Enter", "OnTracExcelAddin")

        Try
            _CmdBars = New CmdBars
            _CmdBars.CreateCommandBars()

            '_AppEvents = New AppEvents
            '_AppEvents.Initialize()
        Catch ex As Exception
            'PLLog.Error(ex, "OnTracExcelAddin")
        End Try

        'PLLog.Trace("Exit", "OnTracExcelAddin")
    End Sub

    Private Sub ThisAddIn_Shutdown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shutdown
        'PLLog.Trace("Enter", "OnTracExcelAddin")

        Try
            _CmdBars.RemoveCommandBars()
            _CmdBars = Nothing
            '_AppEvents = Nothing
        Catch ex As Exception
            'PLLog.Error(ex, "OnTracExcelAddin")
        End Try

        'PLLog.Trace("Exit", "OnTracExcelAddin")
    End Sub

End Class
