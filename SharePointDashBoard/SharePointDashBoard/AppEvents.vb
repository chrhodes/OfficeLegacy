Public Class AppEvents
    ' Handles all application generated events.  Code is in <Application>AppEvents class.
    ' Handles all application command bar generated events.  Code is in <Application>CBarEvents class.

    Private _ExcelAppEvents As ExcelAppEvents
    'Private _OutlookAppEvents As OutlookAppEvents
    'Private _PowerPointAppEvents As PowerPointAppEvents
    'Private _ProjectAppEvents As ProjectAppEvents
    'Private _VisioAppEvents As VisioAppEvents
    'Private _WordAppEvents As WordAppEvents

    'Public Sub Initialize()
    '    Select Case Globals.ThisAddIn.Application.Name
    '        'Case "Microsoft Access"
    '        '    ' g_Access = application

    '        Case "Microsoft Excel"
    '            'Globals.Excel = DirectCast(Globals.HostApp, Microsoft.Office.Interop.Excel.Application)
    '            'Globals.AppName = Globals.Excel.Name
    '            'Globals.AppVersion = Globals.Excel.Version

    '            If Globals.HAS_EXCEL_APP_EVENTS Then
    '                _ExcelAppEvents = New ExcelAppEvents
    '                '_ExcelAppEvents.ExcelAppEvent = DirectCast(Globals.HostApp, Microsoft.Office.Interop.Excel.Application)
    '                _ExcelAppEvents.ExcelAppEvent = Globals.ThisAddIn.Application
    '            End If

    '            'Case "Microsoft PowerPoint"
    '            '    Globals.PowerPoint = DirectCast(Globals.HostApp, Microsoft.Office.Interop.PowerPoint.Application)
    '            '    Globals.AppName = Globals.PowerPoint.Name
    '            '    Globals.AppVersion = Globals.PowerPoint.Version

    '            '    If Globals.HAS_POWERPOINT_APP_EVENTS Then
    '            '        _PowerPointAppEvents = New PowerPointAppEvents
    '            '        _PowerPointAppEvents.PowerPointAppEvent = DirectCast(Globals.HostApp, Microsoft.Office.Interop.PowerPoint.Application)
    '            '    End If

    '            'Case "Microsoft Project"
    '            '    Globals.Project = DirectCast(Globals.HostApp, Microsoft.Office.Interop.MSProject.Application)
    '            '    Globals.AppName = Globals.Project.Name
    '            '    Globals.AppVersion = Globals.Project.Version

    '            '    If Globals.HAS_PROJECT_APP_EVENTS Then
    '            '        _ProjectAppEvents = New ProjectAppEvents
    '            '        _ProjectAppEvents.ProjectAppEvent = DirectCast(Globals.HostApp, Microsoft.Office.Interop.MSProject.Application)
    '            '    End If

    '            'Case "Microsoft Visio"
    '            '    Globals.Visio = DirectCast(Globals.HostApp, Microsoft.Office.Interop.Visio.Application)
    '            '    Globals.AppName = "Microsoft Visio" ' No name is present in .Name property??
    '            '    Globals.AppVersion = Globals.Visio.Version

    '            '    If Globals.HAS_VISIO_APP_EVENTS Then
    '            '        _VisioAppEvents = New VisioAppEvents
    '            '        _VisioAppEvents.VisioAppEvent = DirectCast(Globals.HostApp, Microsoft.Office.Interop.Visio.Application)
    '            '    End If

    '            'Case "Microsoft Word"
    '            '    Globals.Word = DirectCast(Globals.HostApp, Microsoft.Office.Interop.Word.Application)
    '            '    Globals.AppName = Globals.Word.Name
    '            '    Globals.AppVersion = Globals.Word.Version

    '            '    If Globals.HAS_WORD_APP_EVENTS Then
    '            '        _WordAppEvents = New WordAppEvents
    '            '        _WordAppEvents.WordAppEvent = DirectCast(Globals.HostApp, Microsoft.Office.Interop.Word.Application)
    '            '    End If

    '            'Case "Outlook"
    '            '    Globals.Outlook = DirectCast(Globals.HostApp, Microsoft.Office.Interop.Outlook.Application)
    '            '    Globals.AppName = Globals.Outlook.Name
    '            '    Globals.AppVersion = Globals.Outlook.Version

    '            '    If Globals.HAS_OUTLOOK_APP_EVENTS Then
    '            '        _OutlookAppEvents = New OutlookAppEvents
    '            '        _OutlookAppEvents.OutlookAppEvent = DirectCast(Globals.HostApp, Microsoft.Office.Interop.Outlook.Application)
    '            '    End If

    '        Case Else
    '            MsgBox("Unknown AppName:" & Globals.ThisAddIn.Application.Name)
    '    End Select
    'End Sub
End Class
