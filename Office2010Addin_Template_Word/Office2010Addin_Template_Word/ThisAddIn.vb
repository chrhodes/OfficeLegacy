Imports PacificLife.Life
Imports Word = Microsoft.Office.Interop.Word
Imports Office = Microsoft.Office.Core
Imports Microsoft.Office.Tools.Word
Imports VSTO = Microsoft.Office.Tools.Word

Public Class ThisAddIn

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        PLLog.Trace("Enter", Common.PROJECT_NAME)

        Globals.Ribbons.Ribbon.cbDisplayEvents.Checked = Common.DisplayEvents
        Globals.Ribbons.Ribbon.cbEnableAppEvents.Checked = Common.HAS_APP_EVENTS

        Try
            If Common.HAS_APP_EVENTS Then
                Common.AppEvents = New AppEvents
                Common.AppEvents.Initialize()
            End If

            ' Set the context for the WordHelper code to this application
            Common.WordHelper.Application = Globals.ThisAddIn.Application

        Catch ex As Exception
            PLLog.Error(ex, Common.PROJECT_NAME)
        End Try

        PLLog.Trace("Exit", Common.PROJECT_NAME)
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        PLLog.Trace("Enter", Common.PROJECT_NAME)

        Try
            If Common.HAS_APP_EVENTS Then
                Common.AppEvents = Nothing
            End If
        Catch ex As Exception
            PLLog.Error(ex, Common.PROJECT_NAME)
        End Try

        PLLog.Trace("Exit", Common.PROJECT_NAME)
    End Sub


End Class
