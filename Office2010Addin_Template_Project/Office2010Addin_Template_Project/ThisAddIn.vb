Imports PacificLife.Life


Public Class ThisAddIn

    Public CustomTaskPanes As Microsoft.Office.Tools.CustomTaskPaneCollection

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        PLLog.Trace("Enter", Common.PROJECT_NAME)

        Globals.Ribbons.Ribbon.cbDisplayEvents.Checked = Common.DisplayEvents
        Globals.Ribbons.Ribbon.cbEnableAppEvents.Checked = Common.HAS_APP_EVENTS

        Try
            If Common.HAS_APP_EVENTS Then
                Common.AppEvents = New AppEvents
                Common.AppEvents.Initialize()
            End If

            ' Set the context for the ProejctHelper code to this application
            'Common.ProjectHelper.Application = Globals.ThisAddIn.Application

            ' Have to work a little harder to create custom task panes in Project.

            CustomTaskPanes = Globals.Factory.CreateCustomTaskPaneCollection( _
                Nothing, Nothing, "CustomTaskPanes", "CustomTaskPanes", Me)
            'Common.TaskPaneHelp = myCustomTaskPaneCollection.Add(TaskPane_Help, "TaskPane_Help")



        Catch ex As Exception
            PLLog.Error(ex, Common.PROJECT_NAME)
        End Try

        PLLog.Trace("Exit", Common.PROJECT_NAME)
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

End Class
