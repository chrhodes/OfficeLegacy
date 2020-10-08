Public Class AppEvents
    ' Handles all application generated events.  Code is in <Application>AppEvents class.
    Private _VisioAppEvents As VisioAppEvents

    Public Sub Initialize()
        If Globals.HAS_VISIO_APP_EVENTS Then
            _VisioAppEvents = New VisioAppEvents
            _VisioAppEvents.VisioApplication = Globals.ThisAddIn.Application
        End If
    End Sub
End Class
