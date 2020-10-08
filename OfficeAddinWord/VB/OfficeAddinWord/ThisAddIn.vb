Imports PacificLife.Life

Public Class ThisAddIn
    Private _CmdBars As CmdBars
    Private _AppEvents As AppEvents

    Public AppEventsWatchWindow As AddinHelper.WatchWindow

    Private Sub ThisAddIn_Startup(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Startup
        PLLog.Trace("Enter", Globals.PROJECT_NAME)

        Try
            _CmdBars = New CmdBars
            _CmdBars.CreateCommandBars()

            _AppEvents = New AppEvents
            _AppEvents.Initialize()
        Catch ex As Exception
            PLLog.Error(ex, Globals.PROJECT_NAME)
        End Try

        PLLog.Trace("Exit", Globals.PROJECT_NAME)
    End Sub

    Private Sub ThisAddIn_Shutdown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shutdown
        PLLog.Trace("Enter", Globals.PROJECT_NAME)

        Try
            _CmdBars.RemoveCommandBars()
            _CmdBars = Nothing
            _AppEvents = Nothing
        Catch ex As Exception
            PLLog.Error(ex, Globals.PROJECT_NAME)
        End Try

        PLLog.Trace("Exit", Globals.PROJECT_NAME)
    End Sub
End Class

