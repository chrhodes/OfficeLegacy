Option Strict Off
Option Explicit On

Imports Microsoft.Office.Interop
Imports PacificLife.Life

Public Class VisioAppEvents
    ' This catches the events from the application
    ' ToDo: Does this need to be Public?
    Public WithEvents VisioApplication As Microsoft.Office.Interop.Visio.Application

    Private Const cMODULE_NAME As String = Common.PROJECT_NAME & ".VisioAppEvents"

    Private Shared Sub DisplayInWatchWindow(ByVal i As Short, ByVal outputLine As String)
        If Common.DisplayEvents Then
            AddinHelper.WatchWindow.AddOutputLine(String.Format("{0}:{1}", outputLine, i))
        End If
    End Sub

    Private Sub VisioApplication_AfterModal(ByVal app As Microsoft.Office.Interop.Visio.Application) Handles VisioApplication.AfterModal
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_AfterRemoveHiddenInformation(ByVal doc As Microsoft.Office.Interop.Visio.Document) Handles VisioApplication.AfterRemoveHiddenInformation
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_AfterResume(ByVal app As Microsoft.Office.Interop.Visio.Application) Handles VisioApplication.AfterResume
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_AfterResumeEvents(ByVal app As Microsoft.Office.Interop.Visio.Application) Handles VisioApplication.AfterResumeEvents
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_AppActivated(ByVal app As Microsoft.Office.Interop.Visio.Application) Handles VisioApplication.AppActivated
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_AppDeactivated(ByVal app As Microsoft.Office.Interop.Visio.Application) Handles VisioApplication.AppDeactivated
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_AppObjActivated(ByVal app As Microsoft.Office.Interop.Visio.Application) Handles VisioApplication.AppObjActivated
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_AppObjDeactivated(ByVal app As Microsoft.Office.Interop.Visio.Application) Handles VisioApplication.AppObjDeactivated
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_BeforeDataRecordsetDelete(ByVal DataRecordset As Microsoft.Office.Interop.Visio.DataRecordset) Handles VisioApplication.BeforeDataRecordsetDelete
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_BeforeDocumentClose(ByVal doc As Microsoft.Office.Interop.Visio.Document) Handles VisioApplication.BeforeDocumentClose
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_BeforeDocumentSave(ByVal doc As Microsoft.Office.Interop.Visio.Document) Handles VisioApplication.BeforeDocumentSave
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_BeforeDocumentSaveAs(ByVal doc As Microsoft.Office.Interop.Visio.Document) Handles VisioApplication.BeforeDocumentSaveAs
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_BeforeMasterDelete(ByVal Master As Microsoft.Office.Interop.Visio.Master) Handles VisioApplication.BeforeMasterDelete
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_BeforeModal(ByVal app As Microsoft.Office.Interop.Visio.Application) Handles VisioApplication.BeforeModal
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_BeforePageDelete(ByVal Page As Microsoft.Office.Interop.Visio.Page) Handles VisioApplication.BeforePageDelete
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_BeforeQuit(ByVal app As Microsoft.Office.Interop.Visio.Application) Handles VisioApplication.BeforeQuit
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_BeforeSelectionDelete(ByVal Selection As Microsoft.Office.Interop.Visio.Selection) Handles VisioApplication.BeforeSelectionDelete
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_BeforeShapeDelete(ByVal Shape As Microsoft.Office.Interop.Visio.Shape) Handles VisioApplication.BeforeShapeDelete
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_BeforeShapeTextEdit(ByVal Shape As Microsoft.Office.Interop.Visio.Shape) Handles VisioApplication.BeforeShapeTextEdit
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_BeforeStyleDelete(ByVal Style As Microsoft.Office.Interop.Visio.Style) Handles VisioApplication.BeforeStyleDelete
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_BeforeSuspend(ByVal app As Microsoft.Office.Interop.Visio.Application) Handles VisioApplication.BeforeSuspend
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_BeforeSuspendEvents(ByVal app As Microsoft.Office.Interop.Visio.Application) Handles VisioApplication.BeforeSuspendEvents
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_BeforeWindowClosed(ByVal Window As Microsoft.Office.Interop.Visio.Window) Handles VisioApplication.BeforeWindowClosed
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_BeforeWindowPageTurn(ByVal Window As Microsoft.Office.Interop.Visio.Window) Handles VisioApplication.BeforeWindowPageTurn
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_BeforeWindowSelDelete(ByVal Window As Microsoft.Office.Interop.Visio.Window) Handles VisioApplication.BeforeWindowSelDelete
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_CellChanged(ByVal Cell As Microsoft.Office.Interop.Visio.Cell) Handles VisioApplication.CellChanged
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_ConnectionsAdded(ByVal Connects As Microsoft.Office.Interop.Visio.Connects) Handles VisioApplication.ConnectionsAdded
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_ConnectionsDeleted(ByVal Connects As Microsoft.Office.Interop.Visio.Connects) Handles VisioApplication.ConnectionsDeleted
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_ConvertToGroupCanceled(ByVal Selection As Microsoft.Office.Interop.Visio.Selection) Handles VisioApplication.ConvertToGroupCanceled
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_DataRecordsetAdded(ByVal DataRecordset As Microsoft.Office.Interop.Visio.DataRecordset) Handles VisioApplication.DataRecordsetAdded
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_DataRecordsetChanged(ByVal DataRecordsetChanged As Microsoft.Office.Interop.Visio.DataRecordsetChangedEvent) Handles VisioApplication.DataRecordsetChanged
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_DesignModeEntered(ByVal doc As Microsoft.Office.Interop.Visio.Document) Handles VisioApplication.DesignModeEntered
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_DocumentChanged(ByVal doc As Microsoft.Office.Interop.Visio.Document) Handles VisioApplication.DocumentChanged
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_DocumentCloseCanceled(ByVal doc As Microsoft.Office.Interop.Visio.Document) Handles VisioApplication.DocumentCloseCanceled
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_DocumentCreated(ByVal doc As Microsoft.Office.Interop.Visio.Document) Handles VisioApplication.DocumentCreated
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_DocumentOpened(ByVal doc As Microsoft.Office.Interop.Visio.Document) Handles VisioApplication.DocumentOpened
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_DocumentSaved(ByVal doc As Microsoft.Office.Interop.Visio.Document) Handles VisioApplication.DocumentSaved
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_DocumentSavedAs(ByVal doc As Microsoft.Office.Interop.Visio.Document) Handles VisioApplication.DocumentSavedAs
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_EnterScope(ByVal app As Microsoft.Office.Interop.Visio.Application, ByVal nScopeID As Integer, ByVal bstrDescription As String) Handles VisioApplication.EnterScope
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_ExitScope(ByVal app As Microsoft.Office.Interop.Visio.Application, ByVal nScopeID As Integer, ByVal bstrDescription As String, ByVal bErrOrCancelled As Boolean) Handles VisioApplication.ExitScope
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_FormulaChanged(ByVal Cell As Microsoft.Office.Interop.Visio.Cell) Handles VisioApplication.FormulaChanged
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_GroupCanceled(ByVal Selection As Microsoft.Office.Interop.Visio.Selection) Handles VisioApplication.GroupCanceled
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    'Private Sub VisioApplication_KeyDown(ByVal KeyCode As Integer, ByVal KeyButtonState As Integer, ByRef CancelDefault As Boolean) Handles VisioApplication.KeyDown
    '    Static i As Short
    '    i = i + 1
        'DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    'End Sub

    Private Sub VisioApplication_KeyPress(ByVal KeyAscii As Integer, ByRef CancelDefault As Boolean) Handles VisioApplication.KeyPress
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    'Private Sub VisioApplication_KeyUp(ByVal KeyCode As Integer, ByVal KeyButtonState As Integer, ByRef CancelDefault As Boolean) Handles VisioApplication.KeyUp
    '    Static i As Short
    '    i = i + 1
        'DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    'End Sub

    Private Sub VisioApplication_MarkerEvent(ByVal app As Microsoft.Office.Interop.Visio.Application, ByVal SequenceNum As Integer, ByVal ContextString As String) Handles VisioApplication.MarkerEvent
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_MasterAdded(ByVal Master As Microsoft.Office.Interop.Visio.Master) Handles VisioApplication.MasterAdded
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_MasterChanged(ByVal Master As Microsoft.Office.Interop.Visio.Master) Handles VisioApplication.MasterChanged
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_MasterDeleteCanceled(ByVal Master As Microsoft.Office.Interop.Visio.Master) Handles VisioApplication.MasterDeleteCanceled
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    'Private Sub VisioApplication_MouseDown(ByVal Button As Integer, ByVal KeyButtonState As Integer, ByVal x As Double, ByVal y As Double, ByRef CancelDefault As Boolean) Handles VisioApplication.MouseDown
    '    Static i As Short
    '    i = i + 1
        'DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    'End Sub

    ' This is chatty

    'Private Sub VisioApplication_MouseMove(ByVal Button As Integer, ByVal KeyButtonState As Integer, ByVal x As Double, ByVal y As Double, ByRef CancelDefault As Boolean) Handles VisioApplication.MouseMove
    '    Static i As Short
    '    i = i + 1
        'DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    'End Sub

    'Private Sub VisioApplication_MouseUp(ByVal Button As Integer, ByVal KeyButtonState As Integer, ByVal x As Double, ByVal y As Double, ByRef CancelDefault As Boolean) Handles VisioApplication.MouseUp
    '    Static i As Short
    '    i = i + 1
        'DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    'End Sub

    ' This is very chatty.

    'Private Sub VisioApplication_MustFlushScopeBeginning(ByVal app As Microsoft.Office.Interop.Visio.Application) Handles VisioApplication.MustFlushScopeBeginning
    '    Static i As Short
    '    i = i + 1
        'DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    'End Sub

    ' This is very chatty.

    'Private Sub VisioApplication_MustFlushScopeEnded(ByVal app As Microsoft.Office.Interop.Visio.Application) Handles VisioApplication.MustFlushScopeEnded
    '    Static i As Short
    '    i = i + 1
        'DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    'End Sub

    ' This is very chatty.

    'Private Sub VisioApplication_NoEventsPending(ByVal app As Microsoft.Office.Interop.Visio.Application) Handles VisioApplication.NoEventsPending
    '    Static i As Short
    '    i = i + 1
        'DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    'End Sub

    Private Function VisioApplication_OnKeystrokeMessageForAddon(ByVal MSG As Microsoft.Office.Interop.Visio.MSGWrap) As Boolean Handles VisioApplication.OnKeystrokeMessageForAddon
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Function

    Private Sub VisioApplication_PageAdded(ByVal Page As Microsoft.Office.Interop.Visio.Page) Handles VisioApplication.PageAdded
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_PageChanged(ByVal Page As Microsoft.Office.Interop.Visio.Page) Handles VisioApplication.PageChanged
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_PageDeleteCanceled(ByVal Page As Microsoft.Office.Interop.Visio.Page) Handles VisioApplication.PageDeleteCanceled
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Function VisioApplication_QueryCancelConvertToGroup(ByVal Selection As Microsoft.Office.Interop.Visio.Selection) As Boolean Handles VisioApplication.QueryCancelConvertToGroup
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Function

    Private Function VisioApplication_QueryCancelDocumentClose(ByVal doc As Microsoft.Office.Interop.Visio.Document) As Boolean Handles VisioApplication.QueryCancelDocumentClose
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Function

    Private Function VisioApplication_QueryCancelGroup(ByVal Selection As Microsoft.Office.Interop.Visio.Selection) As Boolean Handles VisioApplication.QueryCancelGroup
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Function

    Private Function VisioApplication_QueryCancelMasterDelete(ByVal Master As Microsoft.Office.Interop.Visio.Master) As Boolean Handles VisioApplication.QueryCancelMasterDelete
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Function

    Private Function VisioApplication_QueryCancelPageDelete(ByVal Page As Microsoft.Office.Interop.Visio.Page) As Boolean Handles VisioApplication.QueryCancelPageDelete
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Function

    Private Function VisioApplication_QueryCancelQuit(ByVal app As Microsoft.Office.Interop.Visio.Application) As Boolean Handles VisioApplication.QueryCancelQuit
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Function

    Private Function VisioApplication_QueryCancelSelectionDelete(ByVal Selection As Microsoft.Office.Interop.Visio.Selection) As Boolean Handles VisioApplication.QueryCancelSelectionDelete
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Function

    Private Function VisioApplication_QueryCancelStyleDelete(ByVal Style As Microsoft.Office.Interop.Visio.Style) As Boolean Handles VisioApplication.QueryCancelStyleDelete
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Function

    Private Function VisioApplication_QueryCancelSuspend(ByVal app As Microsoft.Office.Interop.Visio.Application) As Boolean Handles VisioApplication.QueryCancelSuspend
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Function

    Private Function VisioApplication_QueryCancelSuspendEvents(ByVal app As Microsoft.Office.Interop.Visio.Application) As Boolean Handles VisioApplication.QueryCancelSuspendEvents
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Function

    Private Function VisioApplication_QueryCancelUngroup(ByVal Selection As Microsoft.Office.Interop.Visio.Selection) As Boolean Handles VisioApplication.QueryCancelUngroup
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Function

    Private Function VisioApplication_QueryCancelWindowClose(ByVal Window As Microsoft.Office.Interop.Visio.Window) As Boolean Handles VisioApplication.QueryCancelWindowClose
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Function

    Private Sub VisioApplication_QuitCanceled(ByVal app As Microsoft.Office.Interop.Visio.Application) Handles VisioApplication.QuitCanceled
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_RunModeEntered(ByVal doc As Microsoft.Office.Interop.Visio.Document) Handles VisioApplication.RunModeEntered
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_SelectionAdded(ByVal Selection As Microsoft.Office.Interop.Visio.Selection) Handles VisioApplication.SelectionAdded
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)

        Dim shapeIDs As System.Array

        Selection.GetIDs(shapeIDs)

        For Each shape As Visio.Shape In shapeIDs
        DisplayInWatchWindow(i, String.Format("{0}:{1}:{2}", System.Reflection.MethodInfo.GetCurrentMethod().Name, Shape.ID, Shape.Text))
        Next
    End Sub

    Private Sub VisioApplication_SelectionChanged(ByVal Window As Microsoft.Office.Interop.Visio.Window) Handles VisioApplication.SelectionChanged
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_SelectionDeleteCanceled(ByVal Selection As Microsoft.Office.Interop.Visio.Selection) Handles VisioApplication.SelectionDeleteCanceled
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_ShapeAdded(ByVal Shape As Microsoft.Office.Interop.Visio.Shape) Handles VisioApplication.ShapeAdded
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, String.Format("{0}:{1}", Shape.ID, System.Reflection.MethodInfo.GetCurrentMethod().Name))
    End Sub

    Private Sub VisioApplication_ShapeChanged(ByVal Shape As Microsoft.Office.Interop.Visio.Shape) Handles VisioApplication.ShapeChanged
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_ShapeDataGraphicChanged(ByVal Shape As Microsoft.Office.Interop.Visio.Shape) Handles VisioApplication.ShapeDataGraphicChanged
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_ShapeExitedTextEdit(ByVal Shape As Microsoft.Office.Interop.Visio.Shape) Handles VisioApplication.ShapeExitedTextEdit
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_ShapeLinkAdded(ByVal Shape As Microsoft.Office.Interop.Visio.Shape, ByVal DataRecordsetID As Integer, ByVal DataRowID As Integer) Handles VisioApplication.ShapeLinkAdded
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_ShapeLinkDeleted(ByVal Shape As Microsoft.Office.Interop.Visio.Shape, ByVal DataRecordsetID As Integer, ByVal DataRowID As Integer) Handles VisioApplication.ShapeLinkDeleted
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_ShapeParentChanged(ByVal Shape As Microsoft.Office.Interop.Visio.Shape) Handles VisioApplication.ShapeParentChanged
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_StyleAdded(ByVal Style As Microsoft.Office.Interop.Visio.Style) Handles VisioApplication.StyleAdded
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_StyleChanged(ByVal Style As Microsoft.Office.Interop.Visio.Style) Handles VisioApplication.StyleChanged
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_StyleDeleteCanceled(ByVal Style As Microsoft.Office.Interop.Visio.Style) Handles VisioApplication.StyleDeleteCanceled
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_SuspendCanceled(ByVal app As Microsoft.Office.Interop.Visio.Application) Handles VisioApplication.SuspendCanceled
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_SuspendEventsCanceled(ByVal app As Microsoft.Office.Interop.Visio.Application) Handles VisioApplication.SuspendEventsCanceled
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_TextChanged(ByVal Shape As Microsoft.Office.Interop.Visio.Shape) Handles VisioApplication.TextChanged
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_UngroupCanceled(ByVal Selection As Microsoft.Office.Interop.Visio.Selection) Handles VisioApplication.UngroupCanceled
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_ViewChanged(ByVal Window As Microsoft.Office.Interop.Visio.Window) Handles VisioApplication.ViewChanged
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    ' This happens all the time.  Probably don't want to catch it :)

    'Private Sub VisioApplication_VisioIsIdle(ByVal app As Microsoft.Office.Interop.Visio.Application) Handles VisioApplication.VisioIsIdle
    '    Static i As Short
    '    i = i + 1
    Public Sub DisplayInWatchWindow()
        
    End Sub
    'End Sub

    Private Sub VisioApplication_WindowActivated(ByVal Window As Microsoft.Office.Interop.Visio.Window) Handles VisioApplication.WindowActivated
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_WindowChanged(ByVal Window As Microsoft.Office.Interop.Visio.Window) Handles VisioApplication.WindowChanged
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_WindowCloseCanceled(ByVal Window As Microsoft.Office.Interop.Visio.Window) Handles VisioApplication.WindowCloseCanceled
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_WindowOpened(ByVal Window As Microsoft.Office.Interop.Visio.Window) Handles VisioApplication.WindowOpened
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub VisioApplication_WindowTurnedToPage(ByVal Window As Microsoft.Office.Interop.Visio.Window) Handles VisioApplication.WindowTurnedToPage
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub
End Class