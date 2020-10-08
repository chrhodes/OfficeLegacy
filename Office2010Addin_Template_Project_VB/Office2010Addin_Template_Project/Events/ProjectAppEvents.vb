Option Strict Off
Option Explicit On

Public Class ProjectAppEvents
    ' This catches the events from the application
    Public WithEvents ProjectApplication As Microsoft.Office.Interop.MSProject.Application

    Private Const cMODULE_NAME As String = Common.PROJECT_NAME & ".ProjectAppEvents"

    Private Sub ProjectApplication_AfterCubeBuilt(ByRef CubeFileName As String) Handles ProjectApplication.AfterCubeBuilt
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_ApplicationBeforeClose(ByVal Info As Microsoft.Office.Interop.MSProject.EventInfo) Handles ProjectApplication.ApplicationBeforeClose
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_ConnectionStatusChanged(ByVal online As Boolean) Handles ProjectApplication.ConnectionStatusChanged
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_IsFunctionalitySupported(ByVal bstrFunctionality As String, ByVal Info As Microsoft.Office.Interop.MSProject.EventInfo) Handles ProjectApplication.IsFunctionalitySupported
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_JobCompleted(ByVal bstrName As String, ByVal bstrprojGuid As String, ByVal bstrjobGuid As String, ByVal jobType As Integer, ByVal lResult As Integer) Handles ProjectApplication.JobCompleted
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_JobStart(ByVal bstrName As String, ByVal bstrprojGuid As String, ByVal bstrjobGuid As String, ByVal jobType As Integer, ByVal lResult As Integer) Handles ProjectApplication.JobStart
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_LoadWebPage(ByVal Window As Microsoft.Office.Interop.MSProject.Window, ByRef TargetPage As String) Handles ProjectApplication.LoadWebPage
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_LoadWebPane(ByVal Window As Microsoft.Office.Interop.MSProject.Window, ByRef TargetPage As String) Handles ProjectApplication.LoadWebPane
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_NewProject(ByVal pj As Microsoft.Office.Interop.MSProject.Project) Handles ProjectApplication.NewProject
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_OnUndoOrRedo(ByVal bstrLabel As String, ByVal bstrGUID As String, ByVal fUndo As Boolean) Handles ProjectApplication.OnUndoOrRedo
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_PaneActivate() Handles ProjectApplication.PaneActivate
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_ProjectAfterSave() Handles ProjectApplication.ProjectAfterSave
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_ProjectAssignmentNew(ByVal pj As Microsoft.Office.Interop.MSProject.Project, ByVal ID As Integer) Handles ProjectApplication.ProjectAssignmentNew
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_ProjectBeforeAssignmentChange(ByVal asg As Microsoft.Office.Interop.MSProject.Assignment, ByVal Field As Microsoft.Office.Interop.MSProject.PjAssignmentField, ByVal NewVal As Object, ByRef Cancel As Boolean) Handles ProjectApplication.ProjectBeforeAssignmentChange
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_ProjectBeforeAssignmentChange2(ByVal asg As Microsoft.Office.Interop.MSProject.Assignment, ByVal Field As Microsoft.Office.Interop.MSProject.PjAssignmentField, ByVal NewVal As Object, ByVal Info As Microsoft.Office.Interop.MSProject.EventInfo) Handles ProjectApplication.ProjectBeforeAssignmentChange2
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_ProjectBeforeAssignmentDelete(ByVal asg As Microsoft.Office.Interop.MSProject.Assignment, ByRef Cancel As Boolean) Handles ProjectApplication.ProjectBeforeAssignmentDelete
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_ProjectBeforeAssignmentDelete2(ByVal asg As Microsoft.Office.Interop.MSProject.Assignment, ByVal Info As Microsoft.Office.Interop.MSProject.EventInfo) Handles ProjectApplication.ProjectBeforeAssignmentDelete2
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_ProjectBeforeAssignmentNew(ByVal pj As Microsoft.Office.Interop.MSProject.Project, ByRef Cancel As Boolean) Handles ProjectApplication.ProjectBeforeAssignmentNew
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_ProjectBeforeAssignmentNew2(ByVal pj As Microsoft.Office.Interop.MSProject.Project, ByVal Info As Microsoft.Office.Interop.MSProject.EventInfo) Handles ProjectApplication.ProjectBeforeAssignmentNew2
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_ProjectBeforeClearBaseline(ByVal pj As Microsoft.Office.Interop.MSProject.Project, ByVal Interim As Boolean, ByVal bl As Microsoft.Office.Interop.MSProject.PjBaselines, ByVal InterimFrom As Microsoft.Office.Interop.MSProject.PjSaveBaselineTo, ByVal AllTasks As Boolean, ByVal Info As Microsoft.Office.Interop.MSProject.EventInfo) Handles ProjectApplication.ProjectBeforeClearBaseline
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_ProjectBeforeClose(ByVal pj As Microsoft.Office.Interop.MSProject.Project, ByRef Cancel As Boolean) Handles ProjectApplication.ProjectBeforeClose
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_ProjectBeforeClose2(ByVal pj As Microsoft.Office.Interop.MSProject.Project, ByVal Info As Microsoft.Office.Interop.MSProject.EventInfo) Handles ProjectApplication.ProjectBeforeClose2
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_ProjectBeforePrint(ByVal pj As Microsoft.Office.Interop.MSProject.Project, ByRef Cancel As Boolean) Handles ProjectApplication.ProjectBeforePrint
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_ProjectBeforePrint2(ByVal pj As Microsoft.Office.Interop.MSProject.Project, ByVal Info As Microsoft.Office.Interop.MSProject.EventInfo) Handles ProjectApplication.ProjectBeforePrint2
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_ProjectBeforePublish(ByVal pj As Microsoft.Office.Interop.MSProject.Project, ByRef Cancel As Boolean) Handles ProjectApplication.ProjectBeforePublish
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_ProjectBeforeResourceChange(ByVal res As Microsoft.Office.Interop.MSProject.Resource, ByVal Field As Microsoft.Office.Interop.MSProject.PjField, ByVal NewVal As Object, ByRef Cancel As Boolean) Handles ProjectApplication.ProjectBeforeResourceChange
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_ProjectBeforeResourceChange2(ByVal res As Microsoft.Office.Interop.MSProject.Resource, ByVal Field As Microsoft.Office.Interop.MSProject.PjField, ByVal NewVal As Object, ByVal Info As Microsoft.Office.Interop.MSProject.EventInfo) Handles ProjectApplication.ProjectBeforeResourceChange2
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_ProjectBeforeResourceDelete(ByVal res As Microsoft.Office.Interop.MSProject.Resource, ByRef Cancel As Boolean) Handles ProjectApplication.ProjectBeforeResourceDelete
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_ProjectBeforeResourceDelete2(ByVal res As Microsoft.Office.Interop.MSProject.Resource, ByVal Info As Microsoft.Office.Interop.MSProject.EventInfo) Handles ProjectApplication.ProjectBeforeResourceDelete2
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_ProjectBeforeResourceNew(ByVal pj As Microsoft.Office.Interop.MSProject.Project, ByRef Cancel As Boolean) Handles ProjectApplication.ProjectBeforeResourceNew
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_ProjectBeforeResourceNew2(ByVal pj As Microsoft.Office.Interop.MSProject.Project, ByVal Info As Microsoft.Office.Interop.MSProject.EventInfo) Handles ProjectApplication.ProjectBeforeResourceNew2
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_ProjectBeforeSave(ByVal pj As Microsoft.Office.Interop.MSProject.Project, ByVal SaveAsUi As Boolean, ByRef Cancel As Boolean) Handles ProjectApplication.ProjectBeforeSave
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_ProjectBeforeSave2(ByVal pj As Microsoft.Office.Interop.MSProject.Project, ByVal SaveAsUi As Boolean, ByVal Info As Microsoft.Office.Interop.MSProject.EventInfo) Handles ProjectApplication.ProjectBeforeSave2
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_ProjectBeforeSaveBaseline(ByVal pj As Microsoft.Office.Interop.MSProject.Project, ByVal Interim As Boolean, ByVal bl As Microsoft.Office.Interop.MSProject.PjBaselines, ByVal InterimCopy As Microsoft.Office.Interop.MSProject.PjSaveBaselineFrom, ByVal InterimInto As Microsoft.Office.Interop.MSProject.PjSaveBaselineTo, ByVal AllTasks As Boolean, ByVal RollupToSummaryTasks As Boolean, ByVal RollupFromSubtasks As Boolean, ByVal Info As Microsoft.Office.Interop.MSProject.EventInfo) Handles ProjectApplication.ProjectBeforeSaveBaseline
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_ProjectBeforeTaskChange(ByVal tsk As Microsoft.Office.Interop.MSProject.Task, ByVal Field As Microsoft.Office.Interop.MSProject.PjField, ByVal NewVal As Object, ByRef Cancel As Boolean) Handles ProjectApplication.ProjectBeforeTaskChange
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_ProjectBeforeTaskChange2(ByVal tsk As Microsoft.Office.Interop.MSProject.Task, ByVal Field As Microsoft.Office.Interop.MSProject.PjField, ByVal NewVal As Object, ByVal Info As Microsoft.Office.Interop.MSProject.EventInfo) Handles ProjectApplication.ProjectBeforeTaskChange2
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_ProjectBeforeTaskDelete(ByVal tsk As Microsoft.Office.Interop.MSProject.Task, ByRef Cancel As Boolean) Handles ProjectApplication.ProjectBeforeTaskDelete
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_ProjectBeforeTaskDelete2(ByVal tsk As Microsoft.Office.Interop.MSProject.Task, ByVal Info As Microsoft.Office.Interop.MSProject.EventInfo) Handles ProjectApplication.ProjectBeforeTaskDelete2
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_ProjectBeforeTaskNew(ByVal pj As Microsoft.Office.Interop.MSProject.Project, ByRef Cancel As Boolean) Handles ProjectApplication.ProjectBeforeTaskNew
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_ProjectBeforeTaskNew2(ByVal pj As Microsoft.Office.Interop.MSProject.Project, ByVal Info As Microsoft.Office.Interop.MSProject.EventInfo) Handles ProjectApplication.ProjectBeforeTaskNew2
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_ProjectCalculate(ByVal pj As Microsoft.Office.Interop.MSProject.Project) Handles ProjectApplication.ProjectCalculate
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_ProjectResourceNew(ByVal pj As Microsoft.Office.Interop.MSProject.Project, ByVal ID As Integer) Handles ProjectApplication.ProjectResourceNew
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_ProjectTaskNew(ByVal pj As Microsoft.Office.Interop.MSProject.Project, ByVal ID As Integer) Handles ProjectApplication.ProjectTaskNew
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_SaveCompletedToServer(ByVal bstrName As String, ByVal bstrprojGuid As String) Handles ProjectApplication.SaveCompletedToServer
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_SaveStartingToServer(ByVal bstrName As String, ByVal bstrprojGuid As String) Handles ProjectApplication.SaveStartingToServer
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_SecondaryViewChange(ByVal Window As Microsoft.Office.Interop.MSProject.Window, ByVal prevView As Microsoft.Office.Interop.MSProject.View, ByVal newView As Microsoft.Office.Interop.MSProject.View, ByVal success As Boolean) Handles ProjectApplication.SecondaryViewChange
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_WindowActivate(ByVal activatedWindow As Microsoft.Office.Interop.MSProject.Window) Handles ProjectApplication.WindowActivate
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_WindowBeforeViewChange(ByVal Window As Microsoft.Office.Interop.MSProject.Window, ByVal prevView As Microsoft.Office.Interop.MSProject.View, ByVal newView As Microsoft.Office.Interop.MSProject.View, ByVal projectHasViewWindow As Boolean, ByVal Info As Microsoft.Office.Interop.MSProject.EventInfo) Handles ProjectApplication.WindowBeforeViewChange
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_WindowDeactivate(ByVal deactivatedWindow As Microsoft.Office.Interop.MSProject.Window) Handles ProjectApplication.WindowDeactivate
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_WindowGoalAreaChange(ByVal Window As Microsoft.Office.Interop.MSProject.Window, ByVal goalArea As Integer) Handles ProjectApplication.WindowGoalAreaChange
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_WindowSelectionChange(ByVal Window As Microsoft.Office.Interop.MSProject.Window, ByVal sel As Microsoft.Office.Interop.MSProject.Selection, ByVal selType As Object) Handles ProjectApplication.WindowSelectionChange
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_WindowSidepaneDisplayChange(ByVal Window As Microsoft.Office.Interop.MSProject.Window, ByVal Close As Boolean) Handles ProjectApplication.WindowSidepaneDisplayChange
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_WindowSidepaneTaskChange(ByVal Window As Microsoft.Office.Interop.MSProject.Window, ByVal ID As Integer, ByVal IsGoalArea As Boolean) Handles ProjectApplication.WindowSidepaneTaskChange
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_WindowViewChange(ByVal Window As Microsoft.Office.Interop.MSProject.Window, ByVal prevView As Microsoft.Office.Interop.MSProject.View, ByVal newView As Microsoft.Office.Interop.MSProject.View, ByVal success As Boolean) Handles ProjectApplication.WindowViewChange
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Sub ProjectApplication_WorkpaneDisplayChange(ByVal DisplayState As Boolean) Handles ProjectApplication.WorkpaneDisplayChange
        Static i As Short
        i = i + 1
        DisplayInWatchWindow(i, System.Reflection.MethodInfo.GetCurrentMethod().Name)
    End Sub

    Private Shared Sub DisplayInWatchWindow(ByVal i As Short, ByVal outputLine As String)
        If Common.DisplayEvents Then
            AddinHelper.WatchWindow.AddOutputLine(String.Format("{0}:{1}", outputLine, i))
        End If
    End Sub
End Class