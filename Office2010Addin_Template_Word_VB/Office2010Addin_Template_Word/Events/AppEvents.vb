Public Class AppEvents
    ' Handles all application generated events.  Code is in <Application>AppEvents class.

    Private _WordAppEvents As WordAppEvents

    Public Sub Initialize()
        If Common.HAS_APP_EVENTS Then
            _WordAppEvents = New WordAppEvents
            _WordAppEvents.WordApplication = Globals.ThisAddIn.Application
        End If
    End Sub
End Class
