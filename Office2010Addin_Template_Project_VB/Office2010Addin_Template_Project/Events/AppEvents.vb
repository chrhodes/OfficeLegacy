Public Class AppEvents
    ' Handles all application generated events.  Code is in <Application>AppEvents class.

    Private _ProjectAppEvents As ProjectAppEvents

    Public Sub Initialize()
        If Common.HAS_APP_EVENTS Then
            _ProjectAppEvents = New ProjectAppEvents
            _ProjectAppEvents.ProjectApplication = Globals.ThisAddIn.Application
        End If
    End Sub
End Class
