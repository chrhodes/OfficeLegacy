Public Class AppEvents
    ' Handles all application generated events.  Code is in <Application>AppEvents class.
    Private _ExcelAppEvents As ExcelAppEvents

    Public Sub Initialize()
        If Globals.HAS_EXCEL_APP_EVENTS Then
            _ExcelAppEvents = New ExcelAppEvents
            _ExcelAppEvents.ExcelApplication = Globals.ThisAddIn.Application
        End If
    End Sub
End Class
