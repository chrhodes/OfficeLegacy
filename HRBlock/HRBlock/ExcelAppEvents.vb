Option Strict Off
Option Explicit On

Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop
Imports PacificLife.Life

Public Class ExcelAppEvents
    ' This catches the events from the application
    Public WithEvents ExcelApplication As Microsoft.Office.Interop.Excel.Application

    Private Const cMODULE_NAME As String = Globals.PROJECT_NAME & ".ExcelAppEvents"

    Private Sub ExcelApplication_AfterCalculate() Handles ExcelApplication.AfterCalculate
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "AfterCalculate")
    End Sub

    Private Sub ExcelAppEvent_NewWorkbook(ByVal Wb As Microsoft.Office.Interop.Excel.Workbook) Handles ExcelApplication.NewWorkbook
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "NewWorkbook")
    End Sub

    Private Sub ExcelAppEvent_SheetActivate(ByVal Sh As Object) Handles ExcelApplication.SheetActivate
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "SheetActivate")
    End Sub

    Private Sub ExcelAppEvent_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Microsoft.Office.Interop.Excel.Range, ByRef Cancel As Boolean) Handles ExcelApplication.SheetBeforeDoubleClick
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "SheetBeforeDoubleClick")
    End Sub

    Private Sub ExcelAppEvent_SheetBeforeRightClick(ByVal Sh As Object, ByVal Target As Microsoft.Office.Interop.Excel.Range, ByRef Cancel As Boolean) Handles ExcelApplication.SheetBeforeRightClick
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "SheetBeforeRightClick")
    End Sub

    Private Sub ExcelAppEvent_SheetCalculate(ByVal Sh As Object) Handles ExcelApplication.SheetCalculate
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "SheetCalculate")
    End Sub

    Private Sub ExcelAppEvent_SheetChange(ByVal Sh As Object, ByVal Target As Microsoft.Office.Interop.Excel.Range) Handles ExcelApplication.SheetChange
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "SheetChange")
    End Sub

    Private Sub ExcelAppEvent_SheetDeactivate(ByVal Sh As Object) Handles ExcelApplication.SheetDeactivate
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "SheetDeactivate")
    End Sub

    Private Sub ExcelAppEvent_SheetFollowHyperlink(ByVal Sh As Object, ByVal Target As Microsoft.Office.Interop.Excel.Hyperlink) Handles ExcelApplication.SheetFollowHyperlink
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "SheetFollowHyperlink")
    End Sub

    Private Sub ExcelAppEvent_SheetPivotTableUpdate(ByVal Sh As Object, ByVal Target As Microsoft.Office.Interop.Excel.PivotTable) Handles ExcelApplication.SheetPivotTableUpdate
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "SheetPivotTableUpdate")
    End Sub

    Private Sub ExcelAppEvent_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Microsoft.Office.Interop.Excel.Range) Handles ExcelApplication.SheetSelectionChange
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "SheetSelectionChange")
    End Sub

    Private Sub ExcelAppEvent_WindowActivate(ByVal Wb As Microsoft.Office.Interop.Excel.Workbook, ByVal Wn As Microsoft.Office.Interop.Excel.Window) Handles ExcelApplication.WindowActivate
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "WindowActivate")
    End Sub

    Private Sub ExcelAppEvent_WindowDeactivate(ByVal Wb As Microsoft.Office.Interop.Excel.Workbook, ByVal Wn As Microsoft.Office.Interop.Excel.Window) Handles ExcelApplication.WindowDeactivate
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "WindowDeactivate")
    End Sub

    Private Sub ExcelAppEvent_WindowResize(ByVal Wb As Microsoft.Office.Interop.Excel.Workbook, ByVal Wn As Microsoft.Office.Interop.Excel.Window) Handles ExcelApplication.WindowResize
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "WindowResize")
    End Sub

    Private Sub ExcelAppEvent_WorkbookActivate(ByVal Wb As Microsoft.Office.Interop.Excel.Workbook) Handles ExcelApplication.WorkbookActivate
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "WorkbookActivate")
    End Sub

    Private Sub ExcelAppEvent_WorkbookAddinInstall(ByVal Wb As Microsoft.Office.Interop.Excel.Workbook) Handles ExcelApplication.WorkbookAddinInstall
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "WorkbookBeforePrint")
    End Sub

    Private Sub ExcelAppEvent_WorkbookAddinUninstall(ByVal Wb As Microsoft.Office.Interop.Excel.Workbook) Handles ExcelApplication.WorkbookAddinUninstall
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "WorkbookAddinUninstall")
    End Sub

    Private Sub ExcelAppEvent_WorkbookBeforeClose(ByVal Wb As Microsoft.Office.Interop.Excel.Workbook, ByRef Cancel As Boolean) Handles ExcelApplication.WorkbookBeforeClose
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "WorkbookBeforeClose")
    End Sub

    Private Sub ExcelAppEvent_WorkbookBeforePrint(ByVal Wb As Microsoft.Office.Interop.Excel.Workbook, ByRef Cancel As Boolean) Handles ExcelApplication.WorkbookBeforePrint
        Static i As Short
        i = i + 1
        'PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "WorkbookBeforePrint")
    End Sub

    Private Sub ExcelAppEvent_WorkbookBeforeSave(ByVal Wb As Microsoft.Office.Interop.Excel.Workbook, ByVal SaveAsUI As Boolean, ByRef Cancel As Boolean) Handles ExcelApplication.WorkbookBeforeSave
        Static i As Short
        i = i + 1
        '    PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "WorkbookBeforeSave")
    End Sub

    Private Sub ExcelAppEvent_WorkbookDeactivate(ByVal Wb As Microsoft.Office.Interop.Excel.Workbook) Handles ExcelApplication.WorkbookDeactivate
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "WorkbookDeactivate")
    End Sub

    Private Sub ExcelAppEvent_WorkbookNewSheet(ByVal Wb As Microsoft.Office.Interop.Excel.Workbook, ByVal Sh As Object) Handles ExcelApplication.WorkbookNewSheet
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "WorkbookNewSheet")
    End Sub

    Private Sub ExcelAppEvent_WorkbookOpen(ByVal Wb As Microsoft.Office.Interop.Excel.Workbook) Handles ExcelApplication.WorkbookOpen
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "WorkbookOpen")
    End Sub

    Private Sub ExcelAppEvent_WorkbookPivotTableCloseConnection(ByVal Wb As Microsoft.Office.Interop.Excel.Workbook, ByVal Target As Microsoft.Office.Interop.Excel.PivotTable) Handles ExcelApplication.WorkbookPivotTableCloseConnection
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "WorkbookPivotTableCloseConnection")
    End Sub

    Private Sub ExcelAppEvent_WorkbookPivotTableOpenConnection(ByVal Wb As Microsoft.Office.Interop.Excel.Workbook, ByVal Target As Microsoft.Office.Interop.Excel.PivotTable) Handles ExcelApplication.WorkbookPivotTableOpenConnection
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "WorkbookPivotTableOpenConnection")
    End Sub

    Private Sub ExcelAppEvent_WorkbookAfterXmlExport(ByVal Wb As Microsoft.Office.Interop.Excel.Workbook, ByVal Map As Microsoft.Office.Interop.Excel.XmlMap, ByVal Url As String, ByVal Result As Microsoft.Office.Interop.Excel.XlXmlExportResult) Handles ExcelApplication.WorkbookAfterXmlExport
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "WorkbookAfterXmlExport")
    End Sub

    Private Sub ExcelAppEvent_WorkbookAfterXmlImport(ByVal Wb As Microsoft.Office.Interop.Excel.Workbook, ByVal Map As Microsoft.Office.Interop.Excel.XmlMap, ByVal IsRefresh As Boolean, ByVal Result As Microsoft.Office.Interop.Excel.XlXmlImportResult) Handles ExcelApplication.WorkbookAfterXmlImport
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "WorkbookAfterXmlImport")
    End Sub

    Private Sub ExcelAppEvent_WorkbookBeforeXmlExport(ByVal Wb As Microsoft.Office.Interop.Excel.Workbook, ByVal Map As Microsoft.Office.Interop.Excel.XmlMap, ByVal Url As String, ByRef Cancel As Boolean) Handles ExcelApplication.WorkbookBeforeXmlExport
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "WorkbookBeforeXmlExport")
    End Sub

    Private Sub ExcelAppEvent_WorkbookBeforeXmlImport(ByVal Wb As Microsoft.Office.Interop.Excel.Workbook, ByVal Map As Microsoft.Office.Interop.Excel.XmlMap, ByVal Url As String, ByVal IsRefresh As Boolean, ByRef Cancel As Boolean) Handles ExcelApplication.WorkbookBeforeXmlImport
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "WorkbookBeforeXmlImport")
    End Sub

    Private Sub ExcelApplication_WorkbookRowsetComplete(ByVal Wb As Microsoft.Office.Interop.Excel.Workbook, ByVal Description As String, ByVal Sheet As String, ByVal Success As Boolean) Handles ExcelApplication.WorkbookRowsetComplete
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "WorkbookRowsetComplete")
    End Sub

    Private Sub ExcelApplication_WorkbookSync(ByVal Wb As Microsoft.Office.Interop.Excel.Workbook, ByVal SyncEventType As Microsoft.Office.Core.MsoSyncEventType) Handles ExcelApplication.WorkbookSync
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "WorkbookSync")
    End Sub

    Private Shared Sub DisplayInWatchWindow(ByVal i As Short, ByVal outputLine As String)
        Dim watchWindow As AddinHelper.WatchWindow = Globals.ThisAddIn.AppEventsWatchWindow
        If Not watchWindow Is Nothing Then
            watchWindow.AddOutputLine(String.Format("{0}:{1}", outputLine, i))
        End If
    End Sub
End Class