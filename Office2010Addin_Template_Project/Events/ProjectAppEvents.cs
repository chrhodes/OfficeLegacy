using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Office2010Addin_Template_Project.Events
{
    class ProjectAppEvents
    {
        private Microsoft.Office.Interop.MSProject.Application _ProjectApplication;
        public Microsoft.Office.Interop.MSProject.Application ProjectApplication
        {
            get
            {
                return _ProjectApplication;
            }
            set
            {
                if (_ProjectApplication != null)
                {
                	// Should remove all the event handlers;
                }

                _ProjectApplication = value;

                if (_ProjectApplication != null)
                {
                    _ProjectApplication.AfterCubeBuilt += new Microsoft.Office.Interop.MSProject._EProjectApp2_AfterCubeBuiltEventHandler(_ProjectApplication_AfterCubeBuilt);
                    _ProjectApplication.ApplicationBeforeClose += new Microsoft.Office.Interop.MSProject._EProjectApp2_ApplicationBeforeCloseEventHandler(_ProjectApplication_ApplicationBeforeClose);
                    _ProjectApplication.ConnectionStatusChanged += new Microsoft.Office.Interop.MSProject._EProjectApp2_ConnectionStatusChangedEventHandler(_ProjectApplication_ConnectionStatusChanged);
                    _ProjectApplication.IsFunctionalitySupported += new Microsoft.Office.Interop.MSProject._EProjectApp2_IsFunctionalitySupportedEventHandler(_ProjectApplication_IsFunctionalitySupported);
                    _ProjectApplication.JobCompleted += new Microsoft.Office.Interop.MSProject._EProjectApp2_JobCompletedEventHandler(_ProjectApplication_JobCompleted);
                    _ProjectApplication.JobStart += new Microsoft.Office.Interop.MSProject._EProjectApp2_JobStartEventHandler(_ProjectApplication_JobStart);
                    _ProjectApplication.LoadWebPage += new Microsoft.Office.Interop.MSProject._EProjectApp2_LoadWebPageEventHandler(_ProjectApplication_LoadWebPage);
                    _ProjectApplication.LoadWebPane += new Microsoft.Office.Interop.MSProject._EProjectApp2_LoadWebPaneEventHandler(_ProjectApplication_LoadWebPane);
                    _ProjectApplication.NewProject += new Microsoft.Office.Interop.MSProject._EProjectApp2_NewProjectEventHandler(_ProjectApplication_NewProject);
                    _ProjectApplication.OnUndoOrRedo += new Microsoft.Office.Interop.MSProject._EProjectApp2_OnUndoOrRedoEventHandler(_ProjectApplication_OnUndoOrRedo);
                    _ProjectApplication.PaneActivate += new Microsoft.Office.Interop.MSProject._EProjectApp2_PaneActivateEventHandler(_ProjectApplication_PaneActivate);
                    _ProjectApplication.ProjectAfterSave += new Microsoft.Office.Interop.MSProject._EProjectApp2_ProjectAfterSaveEventHandler(_ProjectApplication_ProjectAfterSave);
                    _ProjectApplication.ProjectAssignmentNew += new Microsoft.Office.Interop.MSProject._EProjectApp2_ProjectAssignmentNewEventHandler(_ProjectApplication_ProjectAssignmentNew);
                    _ProjectApplication.ProjectBeforeAssignmentChange += new Microsoft.Office.Interop.MSProject._EProjectApp2_ProjectBeforeAssignmentChangeEventHandler(_ProjectApplication_ProjectBeforeAssignmentChange);
                    _ProjectApplication.ProjectBeforeAssignmentChange2 += new Microsoft.Office.Interop.MSProject._EProjectApp2_ProjectBeforeAssignmentChange2EventHandler(_ProjectApplication_ProjectBeforeAssignmentChange2);
                    _ProjectApplication.ProjectBeforeAssignmentDelete += new Microsoft.Office.Interop.MSProject._EProjectApp2_ProjectBeforeAssignmentDeleteEventHandler(_ProjectApplication_ProjectBeforeAssignmentDelete);
                    _ProjectApplication.ProjectBeforeAssignmentDelete2 += new Microsoft.Office.Interop.MSProject._EProjectApp2_ProjectBeforeAssignmentDelete2EventHandler(_ProjectApplication_ProjectBeforeAssignmentDelete2);
                    _ProjectApplication.ProjectBeforeAssignmentNew += new Microsoft.Office.Interop.MSProject._EProjectApp2_ProjectBeforeAssignmentNewEventHandler(_ProjectApplication_ProjectBeforeAssignmentNew);
                    _ProjectApplication.ProjectBeforeAssignmentNew2 += new Microsoft.Office.Interop.MSProject._EProjectApp2_ProjectBeforeAssignmentNew2EventHandler(_ProjectApplication_ProjectBeforeAssignmentNew2);
                    _ProjectApplication.ProjectBeforeClearBaseline += new Microsoft.Office.Interop.MSProject._EProjectApp2_ProjectBeforeClearBaselineEventHandler(_ProjectApplication_ProjectBeforeClearBaseline);
                    _ProjectApplication.ProjectBeforeClose += new Microsoft.Office.Interop.MSProject._EProjectApp2_ProjectBeforeCloseEventHandler(_ProjectApplication_ProjectBeforeClose);
                    _ProjectApplication.ProjectBeforeClose2 += new Microsoft.Office.Interop.MSProject._EProjectApp2_ProjectBeforeClose2EventHandler(_ProjectApplication_ProjectBeforeClose2);
                    _ProjectApplication.ProjectBeforePrint += new Microsoft.Office.Interop.MSProject._EProjectApp2_ProjectBeforePrintEventHandler(_ProjectApplication_ProjectBeforePrint);
                    _ProjectApplication.ProjectBeforePrint2 += new Microsoft.Office.Interop.MSProject._EProjectApp2_ProjectBeforePrint2EventHandler(_ProjectApplication_ProjectBeforePrint2);
                    _ProjectApplication.ProjectBeforePublish += new Microsoft.Office.Interop.MSProject._EProjectApp2_ProjectBeforePublishEventHandler(_ProjectApplication_ProjectBeforePublish);
                    _ProjectApplication.ProjectBeforeResourceChange += new Microsoft.Office.Interop.MSProject._EProjectApp2_ProjectBeforeResourceChangeEventHandler(_ProjectApplication_ProjectBeforeResourceChange);
                    _ProjectApplication.ProjectBeforeResourceChange2 += new Microsoft.Office.Interop.MSProject._EProjectApp2_ProjectBeforeResourceChange2EventHandler(_ProjectApplication_ProjectBeforeResourceChange2);
                    _ProjectApplication.ProjectBeforeResourceDelete += new Microsoft.Office.Interop.MSProject._EProjectApp2_ProjectBeforeResourceDeleteEventHandler(_ProjectApplication_ProjectBeforeResourceDelete);
                    _ProjectApplication.ProjectBeforeResourceDelete2 += new Microsoft.Office.Interop.MSProject._EProjectApp2_ProjectBeforeResourceDelete2EventHandler(_ProjectApplication_ProjectBeforeResourceDelete2);
                    _ProjectApplication.ProjectBeforeResourceNew += new Microsoft.Office.Interop.MSProject._EProjectApp2_ProjectBeforeResourceNewEventHandler(_ProjectApplication_ProjectBeforeResourceNew);
                    _ProjectApplication.ProjectBeforeResourceNew2 += new Microsoft.Office.Interop.MSProject._EProjectApp2_ProjectBeforeResourceNew2EventHandler(_ProjectApplication_ProjectBeforeResourceNew2);
                    _ProjectApplication.ProjectBeforeSave += new Microsoft.Office.Interop.MSProject._EProjectApp2_ProjectBeforeSaveEventHandler(_ProjectApplication_ProjectBeforeSave);
                    _ProjectApplication.ProjectBeforeSave2 += new Microsoft.Office.Interop.MSProject._EProjectApp2_ProjectBeforeSave2EventHandler(_ProjectApplication_ProjectBeforeSave2);
                    _ProjectApplication.ProjectBeforeSaveBaseline += new Microsoft.Office.Interop.MSProject._EProjectApp2_ProjectBeforeSaveBaselineEventHandler(_ProjectApplication_ProjectBeforeSaveBaseline);
                    _ProjectApplication.ProjectBeforeTaskChange += new Microsoft.Office.Interop.MSProject._EProjectApp2_ProjectBeforeTaskChangeEventHandler(_ProjectApplication_ProjectBeforeTaskChange);
                    _ProjectApplication.ProjectBeforeTaskChange2 += new Microsoft.Office.Interop.MSProject._EProjectApp2_ProjectBeforeTaskChange2EventHandler(_ProjectApplication_ProjectBeforeTaskChange2);
                    _ProjectApplication.ProjectBeforeTaskDelete += new Microsoft.Office.Interop.MSProject._EProjectApp2_ProjectBeforeTaskDeleteEventHandler(_ProjectApplication_ProjectBeforeTaskDelete);
                    _ProjectApplication.ProjectBeforeTaskDelete2 += new Microsoft.Office.Interop.MSProject._EProjectApp2_ProjectBeforeTaskDelete2EventHandler(_ProjectApplication_ProjectBeforeTaskDelete2);
                    _ProjectApplication.ProjectBeforeTaskNew += new Microsoft.Office.Interop.MSProject._EProjectApp2_ProjectBeforeTaskNewEventHandler(_ProjectApplication_ProjectBeforeTaskNew);
                    _ProjectApplication.ProjectBeforeTaskNew2 += new Microsoft.Office.Interop.MSProject._EProjectApp2_ProjectBeforeTaskNew2EventHandler(_ProjectApplication_ProjectBeforeTaskNew2);
                    _ProjectApplication.ProjectCalculate += new Microsoft.Office.Interop.MSProject._EProjectApp2_ProjectCalculateEventHandler(_ProjectApplication_ProjectCalculate);
                    _ProjectApplication.ProjectResourceNew += new Microsoft.Office.Interop.MSProject._EProjectApp2_ProjectResourceNewEventHandler(_ProjectApplication_ProjectResourceNew);
                    _ProjectApplication.ProjectTaskNew += new Microsoft.Office.Interop.MSProject._EProjectApp2_ProjectTaskNewEventHandler(_ProjectApplication_ProjectTaskNew);
                    _ProjectApplication.SaveCompletedToServer += new Microsoft.Office.Interop.MSProject._EProjectApp2_SaveCompletedToServerEventHandler(_ProjectApplication_SaveCompletedToServer);
                    _ProjectApplication.SaveStartingToServer += new Microsoft.Office.Interop.MSProject._EProjectApp2_SaveStartingToServerEventHandler(_ProjectApplication_SaveStartingToServer);
                    _ProjectApplication.SecondaryViewChange += new Microsoft.Office.Interop.MSProject._EProjectApp2_SecondaryViewChangeEventHandler(_ProjectApplication_SecondaryViewChange);
                    _ProjectApplication.WindowBeforeViewChange += new Microsoft.Office.Interop.MSProject._EProjectApp2_WindowBeforeViewChangeEventHandler(_ProjectApplication_WindowBeforeViewChange);
                    _ProjectApplication.WindowDeactivate += new Microsoft.Office.Interop.MSProject._EProjectApp2_WindowDeactivateEventHandler(_ProjectApplication_WindowDeactivate);
                    _ProjectApplication.WindowGoalAreaChange += new Microsoft.Office.Interop.MSProject._EProjectApp2_WindowGoalAreaChangeEventHandler(_ProjectApplication_WindowGoalAreaChange);
                    _ProjectApplication.WindowSelectionChange += new Microsoft.Office.Interop.MSProject._EProjectApp2_WindowSelectionChangeEventHandler(_ProjectApplication_WindowSelectionChange);
                    _ProjectApplication.WindowSidepaneDisplayChange += new Microsoft.Office.Interop.MSProject._EProjectApp2_WindowSidepaneDisplayChangeEventHandler(_ProjectApplication_WindowSidepaneDisplayChange);
                    _ProjectApplication.WindowSidepaneTaskChange += new Microsoft.Office.Interop.MSProject._EProjectApp2_WindowSidepaneTaskChangeEventHandler(_ProjectApplication_WindowSidepaneTaskChange);
                    _ProjectApplication.WindowViewChange += new Microsoft.Office.Interop.MSProject._EProjectApp2_WindowViewChangeEventHandler(_ProjectApplication_WindowViewChange);
                    _ProjectApplication.WorkpaneDisplayChange += new Microsoft.Office.Interop.MSProject._EProjectApp2_WorkpaneDisplayChangeEventHandler(_ProjectApplication_WorkpaneDisplayChange);
                }
            }
        }

        short WorkpaneDisplayChange;
        void _ProjectApplication_WorkpaneDisplayChange(bool DisplayState)
        {
            DisplayInWatchWindow(WorkpaneDisplayChange++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short WindowViewChange;
        void _ProjectApplication_WindowViewChange(Microsoft.Office.Interop.MSProject.Window Window, Microsoft.Office.Interop.MSProject.View prevView, Microsoft.Office.Interop.MSProject.View newView, bool success)
        {
            DisplayInWatchWindow(WindowViewChange++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short WindowSidepaneTaskChange;
        void _ProjectApplication_WindowSidepaneTaskChange(Microsoft.Office.Interop.MSProject.Window Window, int ID, bool IsGoalArea)
        {
            DisplayInWatchWindow(WindowSidepaneTaskChange++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short WindowSidepaneDisplayChange;
        void _ProjectApplication_WindowSidepaneDisplayChange(Microsoft.Office.Interop.MSProject.Window Window, bool Close)
        {
            DisplayInWatchWindow(WindowSidepaneDisplayChange++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short WindowSelectionChange;
        void _ProjectApplication_WindowSelectionChange(Microsoft.Office.Interop.MSProject.Window Window, Microsoft.Office.Interop.MSProject.Selection sel, object selType)
        {
            DisplayInWatchWindow(WindowSelectionChange++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short WindowGoalAreaChange;
        void _ProjectApplication_WindowGoalAreaChange(Microsoft.Office.Interop.MSProject.Window Window, int goalArea)
        {
            DisplayInWatchWindow(WindowGoalAreaChange++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short WindowDeactivate;
        void _ProjectApplication_WindowDeactivate(Microsoft.Office.Interop.MSProject.Window deactivatedWindow)
        {
            DisplayInWatchWindow(WindowDeactivate++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short WindowBeforeViewChange;
        void _ProjectApplication_WindowBeforeViewChange(Microsoft.Office.Interop.MSProject.Window Window, Microsoft.Office.Interop.MSProject.View prevView, Microsoft.Office.Interop.MSProject.View newView, bool projectHasViewWindow, Microsoft.Office.Interop.MSProject.EventInfo Info)
        {
            DisplayInWatchWindow(WindowBeforeViewChange++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short SecondaryViewChange;
        void _ProjectApplication_SecondaryViewChange(Microsoft.Office.Interop.MSProject.Window Window, Microsoft.Office.Interop.MSProject.View prevView, Microsoft.Office.Interop.MSProject.View newView, bool success)
        {
            DisplayInWatchWindow(SecondaryViewChange++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short SaveStartingToServer;
        void _ProjectApplication_SaveStartingToServer(string bstrName, string bstrprojGuid)
        {
            DisplayInWatchWindow(SaveStartingToServer++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short SaveCompletedToServer;
        void _ProjectApplication_SaveCompletedToServer(string bstrName, string bstrprojGuid)
        {
            DisplayInWatchWindow(SaveCompletedToServer++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProjectTaskNew;
        void _ProjectApplication_ProjectTaskNew(Microsoft.Office.Interop.MSProject.Project pj, int ID)
        {
            DisplayInWatchWindow(ProjectTaskNew++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProjectResourceNew;
        void _ProjectApplication_ProjectResourceNew(Microsoft.Office.Interop.MSProject.Project pj, int ID)
        {
            DisplayInWatchWindow(ProjectResourceNew++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProjectCalculate;
        void _ProjectApplication_ProjectCalculate(Microsoft.Office.Interop.MSProject.Project pj)
        {
            DisplayInWatchWindow(ProjectCalculate++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProjectBeforeTaskNew2;
        void _ProjectApplication_ProjectBeforeTaskNew2(Microsoft.Office.Interop.MSProject.Project pj, Microsoft.Office.Interop.MSProject.EventInfo Info)
        {
            DisplayInWatchWindow(ProjectBeforeTaskNew2++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProjectBeforeTaskNew;
        void _ProjectApplication_ProjectBeforeTaskNew(Microsoft.Office.Interop.MSProject.Project pj, ref bool Cancel)
        {
            DisplayInWatchWindow(ProjectBeforeTaskNew++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProjectBeforeTaskDelete2;
        void _ProjectApplication_ProjectBeforeTaskDelete2(Microsoft.Office.Interop.MSProject.Task tsk, Microsoft.Office.Interop.MSProject.EventInfo Info)
        {
            DisplayInWatchWindow(ProjectBeforeTaskDelete2++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProjectBeforeTaskDelete;
        void _ProjectApplication_ProjectBeforeTaskDelete(Microsoft.Office.Interop.MSProject.Task tsk, ref bool Cancel)
        {
            DisplayInWatchWindow(ProjectBeforeTaskDelete++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProjectBeforeTaskChange2;
        void _ProjectApplication_ProjectBeforeTaskChange2(Microsoft.Office.Interop.MSProject.Task tsk, Microsoft.Office.Interop.MSProject.PjField Field, object NewVal, Microsoft.Office.Interop.MSProject.EventInfo Info)
        {
            DisplayInWatchWindow(ProjectBeforeTaskChange2++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProjectBeforeTaskChange;
        void _ProjectApplication_ProjectBeforeTaskChange(Microsoft.Office.Interop.MSProject.Task tsk, Microsoft.Office.Interop.MSProject.PjField Field, object NewVal, ref bool Cancel)
        {
            DisplayInWatchWindow(ProjectBeforeTaskChange++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProjectBeforeSaveBaseline;
        void _ProjectApplication_ProjectBeforeSaveBaseline(Microsoft.Office.Interop.MSProject.Project pj, bool Interim, Microsoft.Office.Interop.MSProject.PjBaselines bl, Microsoft.Office.Interop.MSProject.PjSaveBaselineFrom InterimCopy, Microsoft.Office.Interop.MSProject.PjSaveBaselineTo InterimInto, bool AllTasks, bool RollupToSummaryTasks, bool RollupFromSubtasks, Microsoft.Office.Interop.MSProject.EventInfo Info)
        {
            DisplayInWatchWindow(ProjectBeforeSaveBaseline++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProjectBeforeSave2;
        void _ProjectApplication_ProjectBeforeSave2(Microsoft.Office.Interop.MSProject.Project pj, bool SaveAsUi, Microsoft.Office.Interop.MSProject.EventInfo Info)
        {
            DisplayInWatchWindow(ProjectBeforeSave2++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProjectBeforeSave;
        void _ProjectApplication_ProjectBeforeSave(Microsoft.Office.Interop.MSProject.Project pj, bool SaveAsUi, ref bool Cancel)
        {
            DisplayInWatchWindow(ProjectBeforeSave++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProjectBeforeResourceNew2;
        void _ProjectApplication_ProjectBeforeResourceNew2(Microsoft.Office.Interop.MSProject.Project pj, Microsoft.Office.Interop.MSProject.EventInfo Info)
        {
            DisplayInWatchWindow(ProjectBeforeResourceNew2++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProjectBeforeResourceNew;
        void _ProjectApplication_ProjectBeforeResourceNew(Microsoft.Office.Interop.MSProject.Project pj, ref bool Cancel)
        {
            DisplayInWatchWindow(ProjectBeforeResourceNew++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProjectBeforeResourceDelete2;
        void _ProjectApplication_ProjectBeforeResourceDelete2(Microsoft.Office.Interop.MSProject.Resource res, Microsoft.Office.Interop.MSProject.EventInfo Info)
        {
            DisplayInWatchWindow(ProjectBeforeResourceDelete2++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProjectBeforeResourceDelete;
        void _ProjectApplication_ProjectBeforeResourceDelete(Microsoft.Office.Interop.MSProject.Resource res, ref bool Cancel)
        {
            DisplayInWatchWindow(ProjectBeforeResourceDelete++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProjectBeforeResourceChange2;
        void _ProjectApplication_ProjectBeforeResourceChange2(Microsoft.Office.Interop.MSProject.Resource res, Microsoft.Office.Interop.MSProject.PjField Field, object NewVal, Microsoft.Office.Interop.MSProject.EventInfo Info)
        {
            DisplayInWatchWindow(ProjectBeforeResourceChange2++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProjectBeforeResourceChange;
        void _ProjectApplication_ProjectBeforeResourceChange(Microsoft.Office.Interop.MSProject.Resource res, Microsoft.Office.Interop.MSProject.PjField Field, object NewVal, ref bool Cancel)
        {
            DisplayInWatchWindow(ProjectBeforeResourceChange++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProjectBeforePublish;
        void _ProjectApplication_ProjectBeforePublish(Microsoft.Office.Interop.MSProject.Project pj, ref bool Cancel)
        {
            DisplayInWatchWindow(ProjectBeforePublish++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProjectBeforePrint2;
        void _ProjectApplication_ProjectBeforePrint2(Microsoft.Office.Interop.MSProject.Project pj, Microsoft.Office.Interop.MSProject.EventInfo Info)
        {
            DisplayInWatchWindow(ProjectBeforePrint2++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProjectBeforePrint;
        void _ProjectApplication_ProjectBeforePrint(Microsoft.Office.Interop.MSProject.Project pj, ref bool Cancel)
        {
            DisplayInWatchWindow(ProjectBeforePrint++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProjectBeforeClose;
        void _ProjectApplication_ProjectBeforeClose(Microsoft.Office.Interop.MSProject.Project pj, ref bool Cancel)
        {
            DisplayInWatchWindow(ProjectBeforeClose++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProjectBeforeClose2;
        void _ProjectApplication_ProjectBeforeClose2(Microsoft.Office.Interop.MSProject.Project pj, Microsoft.Office.Interop.MSProject.EventInfo Info)
        {
            DisplayInWatchWindow(ProjectBeforeClose2++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProjectBeforeClearBaseline;
        void _ProjectApplication_ProjectBeforeClearBaseline(Microsoft.Office.Interop.MSProject.Project pj, bool Interim, Microsoft.Office.Interop.MSProject.PjBaselines bl, Microsoft.Office.Interop.MSProject.PjSaveBaselineTo InterimFrom, bool AllTasks, Microsoft.Office.Interop.MSProject.EventInfo Info)
        {
                DisplayInWatchWindow(ProjectBeforeClearBaseline++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProjectBeforeAssignmentNew2;
        void _ProjectApplication_ProjectBeforeAssignmentNew2(Microsoft.Office.Interop.MSProject.Project pj, Microsoft.Office.Interop.MSProject.EventInfo Info)
        {
            DisplayInWatchWindow(ProjectBeforeAssignmentNew2++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProjectBeforeAssignmentNew;
        void _ProjectApplication_ProjectBeforeAssignmentNew(Microsoft.Office.Interop.MSProject.Project pj, ref bool Cancel)
        {
            DisplayInWatchWindow(ProjectBeforeAssignmentNew++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProjectBeforeAssignmentDelete2;
        void _ProjectApplication_ProjectBeforeAssignmentDelete2(Microsoft.Office.Interop.MSProject.Assignment asg, Microsoft.Office.Interop.MSProject.EventInfo Info)
        {
            DisplayInWatchWindow(ProjectBeforeAssignmentDelete2++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProjectBeforeAssignmentDelete;
        void _ProjectApplication_ProjectBeforeAssignmentDelete(Microsoft.Office.Interop.MSProject.Assignment asg, ref bool Cancel)
        {
            DisplayInWatchWindow(ProjectBeforeAssignmentDelete++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProjectBeforeAssignmentChange2;
        void _ProjectApplication_ProjectBeforeAssignmentChange2(Microsoft.Office.Interop.MSProject.Assignment asg, Microsoft.Office.Interop.MSProject.PjAssignmentField Field, object NewVal, Microsoft.Office.Interop.MSProject.EventInfo Info)
        {
            DisplayInWatchWindow(ProjectBeforeAssignmentChange2++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProjectBeforeAssignmentChange;
        void _ProjectApplication_ProjectBeforeAssignmentChange(Microsoft.Office.Interop.MSProject.Assignment asg, Microsoft.Office.Interop.MSProject.PjAssignmentField Field, object NewVal, ref bool Cancel)
        {
            DisplayInWatchWindow(ProjectBeforeAssignmentChange++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProjectAssignmentNew;
        void _ProjectApplication_ProjectAssignmentNew(Microsoft.Office.Interop.MSProject.Project pj, int ID)
        {
            DisplayInWatchWindow(ProjectAssignmentNew++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProjectAfterSave;
        void _ProjectApplication_ProjectAfterSave()
        {
            DisplayInWatchWindow(ProjectAfterSave++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short PaneActivate;
        void _ProjectApplication_PaneActivate()
        {
            DisplayInWatchWindow(PaneActivate++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short OnUndoOrRedo;
        void _ProjectApplication_OnUndoOrRedo(string bstrLabel, string bstrGUID, bool fUndo)
        {
            DisplayInWatchWindow(OnUndoOrRedo++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short NewProject;
        void _ProjectApplication_NewProject(Microsoft.Office.Interop.MSProject.Project pj)
        {
            DisplayInWatchWindow(NewProject++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short LoadWebPane;
        void _ProjectApplication_LoadWebPane(Microsoft.Office.Interop.MSProject.Window Window, ref string TargetPage)
        {
            DisplayInWatchWindow(LoadWebPane++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short LoadWebPage;
        void _ProjectApplication_LoadWebPage(Microsoft.Office.Interop.MSProject.Window Window, ref string TargetPage)
        {
            DisplayInWatchWindow(LoadWebPage++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short JobStart;
        void _ProjectApplication_JobStart(string bstrName, string bstrprojGuid, string bstrjobGuid, int jobType, int lResult)
        {
            DisplayInWatchWindow(JobStart++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short JobCompleted;
        void _ProjectApplication_JobCompleted(string bstrName, string bstrprojGuid, string bstrjobGuid, int jobType, int lResult)
        {
            DisplayInWatchWindow(JobCompleted++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short IsFunctionalitySupported;
        void _ProjectApplication_IsFunctionalitySupported(string bstrFunctionality, Microsoft.Office.Interop.MSProject.EventInfo Info)
        {
            DisplayInWatchWindow(IsFunctionalitySupported++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ConnectionStatusChanged;
        void _ProjectApplication_ConnectionStatusChanged(bool online)
        {
            DisplayInWatchWindow(ConnectionStatusChanged++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ApplicationBeforeClose;
        void _ProjectApplication_ApplicationBeforeClose(Microsoft.Office.Interop.MSProject.EventInfo Info)
        {
            DisplayInWatchWindow(ApplicationBeforeClose++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short AfterCubeBuilt;
        void _ProjectApplication_AfterCubeBuilt(ref string CubeFileName)
        {
            DisplayInWatchWindow(AfterCubeBuilt++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        private void DisplayInWatchWindow(short i, string outputLine)
        {
            if (Common.DisplayEvents)
            {
                AddinHelper.Common.WriteToWatchWindow(string.Format("{0}:{1}", outputLine, i));
            }
        }
    }
}
