using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Office2010Addin_Template_Visio.Events
{
    class VisioAppEvents
    {
        private Microsoft.Office.Interop.Visio.Application _VisioApplication;
        public Microsoft.Office.Interop.Visio.Application VisioApplication
        {
            get
            {
                return _VisioApplication;
            }
            set
            {
                if (_VisioApplication != null)
                {
                	// Should remove all the event handlers;
                }

                _VisioApplication = value;

                if (_VisioApplication != null)
                {
                    _VisioApplication.AfterModal += new Microsoft.Office.Interop.Visio.EApplication_AfterModalEventHandler(_VisioApplication_AfterModal);
                    _VisioApplication.AfterRemoveHiddenInformation += new Microsoft.Office.Interop.Visio.EApplication_AfterRemoveHiddenInformationEventHandler(_VisioApplication_AfterRemoveHiddenInformation);
                    _VisioApplication.AfterResume += new Microsoft.Office.Interop.Visio.EApplication_AfterResumeEventHandler(_VisioApplication_AfterResume);
                    _VisioApplication.AfterResumeEvents += new Microsoft.Office.Interop.Visio.EApplication_AfterResumeEventsEventHandler(_VisioApplication_AfterResumeEvents);
                    _VisioApplication.AppActivated += new Microsoft.Office.Interop.Visio.EApplication_AppActivatedEventHandler(_VisioApplication_AppActivated);
                    _VisioApplication.AppDeactivated += new Microsoft.Office.Interop.Visio.EApplication_AppDeactivatedEventHandler(_VisioApplication_AppDeactivated);
                    _VisioApplication.AppObjActivated += new Microsoft.Office.Interop.Visio.EApplication_AppObjActivatedEventHandler(_VisioApplication_AppObjActivated);
                    _VisioApplication.AppObjDeactivated += new Microsoft.Office.Interop.Visio.EApplication_AppObjDeactivatedEventHandler(_VisioApplication_AppObjDeactivated);
                    _VisioApplication.BeforeDataRecordsetDelete += new Microsoft.Office.Interop.Visio.EApplication_BeforeDataRecordsetDeleteEventHandler(_VisioApplication_BeforeDataRecordsetDelete);
                    _VisioApplication.BeforeDocumentClose += new Microsoft.Office.Interop.Visio.EApplication_BeforeDocumentCloseEventHandler(_VisioApplication_BeforeDocumentClose);
                    _VisioApplication.BeforeDocumentSave += new Microsoft.Office.Interop.Visio.EApplication_BeforeDocumentSaveEventHandler(_VisioApplication_BeforeDocumentSave);
                    _VisioApplication.BeforeDocumentSaveAs += new Microsoft.Office.Interop.Visio.EApplication_BeforeDocumentSaveAsEventHandler(_VisioApplication_BeforeDocumentSaveAs);
                    _VisioApplication.BeforeMasterDelete += new Microsoft.Office.Interop.Visio.EApplication_BeforeMasterDeleteEventHandler(_VisioApplication_BeforeMasterDelete);   
                    _VisioApplication.BeforeModal += new Microsoft.Office.Interop.Visio.EApplication_BeforeModalEventHandler(_VisioApplication_BeforeModal);
                    _VisioApplication.BeforePageDelete += new Microsoft.Office.Interop.Visio.EApplication_BeforePageDeleteEventHandler(_VisioApplication_BeforePageDelete);
                    _VisioApplication.BeforeQuit += new Microsoft.Office.Interop.Visio.EApplication_BeforeQuitEventHandler(_VisioApplication_BeforeQuit);
                    _VisioApplication.BeforeSelectionDelete += new Microsoft.Office.Interop.Visio.EApplication_BeforeSelectionDeleteEventHandler(_VisioApplication_BeforeSelectionDelete);
                    _VisioApplication.BeforeShapeDelete += new Microsoft.Office.Interop.Visio.EApplication_BeforeShapeDeleteEventHandler(_VisioApplication_BeforeShapeDelete);
                    _VisioApplication.BeforeShapeTextEdit += new Microsoft.Office.Interop.Visio.EApplication_BeforeShapeTextEditEventHandler(_VisioApplication_BeforeShapeTextEdit);
                    _VisioApplication.BeforeStyleDelete += new Microsoft.Office.Interop.Visio.EApplication_BeforeStyleDeleteEventHandler(_VisioApplication_BeforeStyleDelete);
                    _VisioApplication.BeforeSuspend += new Microsoft.Office.Interop.Visio.EApplication_BeforeSuspendEventHandler(_VisioApplication_BeforeSuspend);
                    _VisioApplication.BeforeSuspendEvents += new Microsoft.Office.Interop.Visio.EApplication_BeforeSuspendEventsEventHandler(_VisioApplication_BeforeSuspendEvents);
                    _VisioApplication.BeforeWindowClosed += new Microsoft.Office.Interop.Visio.EApplication_BeforeWindowClosedEventHandler(_VisioApplication_BeforeWindowClosed);
                    _VisioApplication.BeforeWindowPageTurn += new Microsoft.Office.Interop.Visio.EApplication_BeforeWindowPageTurnEventHandler(_VisioApplication_BeforeWindowPageTurn);
                    _VisioApplication.BeforeWindowSelDelete += new Microsoft.Office.Interop.Visio.EApplication_BeforeWindowSelDeleteEventHandler(_VisioApplication_BeforeWindowSelDelete);
                    _VisioApplication.CalloutRelationshipAdded += new Microsoft.Office.Interop.Visio.EApplication_CalloutRelationshipAddedEventHandler(_VisioApplication_CalloutRelationshipAdded);
                    _VisioApplication.CalloutRelationshipDeleted += new Microsoft.Office.Interop.Visio.EApplication_CalloutRelationshipDeletedEventHandler(_VisioApplication_CalloutRelationshipDeleted);
                    _VisioApplication.CellChanged += new Microsoft.Office.Interop.Visio.EApplication_CellChangedEventHandler(_VisioApplication_CellChanged);
                    _VisioApplication.ConnectionsAdded += new Microsoft.Office.Interop.Visio.EApplication_ConnectionsAddedEventHandler(_VisioApplication_ConnectionsAdded);
                    _VisioApplication.ConnectionsDeleted += new Microsoft.Office.Interop.Visio.EApplication_ConnectionsDeletedEventHandler(_VisioApplication_ConnectionsDeleted);
                    _VisioApplication.ContainerRelationshipAdded += new Microsoft.Office.Interop.Visio.EApplication_ContainerRelationshipAddedEventHandler(_VisioApplication_ContainerRelationshipAdded);
                    _VisioApplication.ContainerRelationshipDeleted += new Microsoft.Office.Interop.Visio.EApplication_ContainerRelationshipDeletedEventHandler(_VisioApplication_ContainerRelationshipDeleted);
                    _VisioApplication.ConvertToGroupCanceled += new Microsoft.Office.Interop.Visio.EApplication_ConvertToGroupCanceledEventHandler(_VisioApplication_ConvertToGroupCanceled);
                    _VisioApplication.DataRecordsetAdded += new Microsoft.Office.Interop.Visio.EApplication_DataRecordsetAddedEventHandler(_VisioApplication_DataRecordsetAdded);
                    _VisioApplication.DataRecordsetChanged += new Microsoft.Office.Interop.Visio.EApplication_DataRecordsetChangedEventHandler(_VisioApplication_DataRecordsetChanged);
                    _VisioApplication.DesignModeEntered += new Microsoft.Office.Interop.Visio.EApplication_DesignModeEnteredEventHandler(_VisioApplication_DesignModeEntered);
                    _VisioApplication.DocumentChanged += new Microsoft.Office.Interop.Visio.EApplication_DocumentChangedEventHandler(_VisioApplication_DocumentChanged);
                    _VisioApplication.DocumentCloseCanceled += new Microsoft.Office.Interop.Visio.EApplication_DocumentCloseCanceledEventHandler(_VisioApplication_DocumentCloseCanceled);
                    _VisioApplication.DocumentCreated += new Microsoft.Office.Interop.Visio.EApplication_DocumentCreatedEventHandler(_VisioApplication_DocumentCreated);
                    _VisioApplication.DocumentOpened += new Microsoft.Office.Interop.Visio.EApplication_DocumentOpenedEventHandler(_VisioApplication_DocumentOpened);
                    _VisioApplication.DocumentSaved += new Microsoft.Office.Interop.Visio.EApplication_DocumentSavedEventHandler(_VisioApplication_DocumentSaved);
                    _VisioApplication.DocumentSavedAs += new Microsoft.Office.Interop.Visio.EApplication_DocumentSavedAsEventHandler(_VisioApplication_DocumentSavedAs);
                    _VisioApplication.EnterScope += new Microsoft.Office.Interop.Visio.EApplication_EnterScopeEventHandler(_VisioApplication_EnterScope);
                    _VisioApplication.ExitScope += new Microsoft.Office.Interop.Visio.EApplication_ExitScopeEventHandler(_VisioApplication_ExitScope);
                    _VisioApplication.FormulaChanged += new Microsoft.Office.Interop.Visio.EApplication_FormulaChangedEventHandler(_VisioApplication_FormulaChanged);
                    _VisioApplication.GroupCanceled += new Microsoft.Office.Interop.Visio.EApplication_GroupCanceledEventHandler(_VisioApplication_GroupCanceled);
                    _VisioApplication.KeyDown += new Microsoft.Office.Interop.Visio.EApplication_KeyDownEventHandler(_VisioApplication_KeyDown);
                    _VisioApplication.KeyPress += new Microsoft.Office.Interop.Visio.EApplication_KeyPressEventHandler(_VisioApplication_KeyPress);
                    _VisioApplication.KeyUp += new Microsoft.Office.Interop.Visio.EApplication_KeyUpEventHandler(_VisioApplication_KeyUp);
                    _VisioApplication.MarkerEvent += new Microsoft.Office.Interop.Visio.EApplication_MarkerEventEventHandler(_VisioApplication_MarkerEvent);
                    _VisioApplication.MasterAdded += new Microsoft.Office.Interop.Visio.EApplication_MasterAddedEventHandler(_VisioApplication_MasterAdded);
                    _VisioApplication.MasterChanged += new Microsoft.Office.Interop.Visio.EApplication_MasterChangedEventHandler(_VisioApplication_MasterChanged);
                    _VisioApplication.MasterDeleteCanceled += new Microsoft.Office.Interop.Visio.EApplication_MasterDeleteCanceledEventHandler(_VisioApplication_MasterDeleteCanceled);
                    _VisioApplication.MouseDown += new Microsoft.Office.Interop.Visio.EApplication_MouseDownEventHandler(_VisioApplication_MouseDown);
                    _VisioApplication.MouseMove += new Microsoft.Office.Interop.Visio.EApplication_MouseMoveEventHandler(_VisioApplication_MouseMove);
                    _VisioApplication.MouseUp += new Microsoft.Office.Interop.Visio.EApplication_MouseUpEventHandler(_VisioApplication_MouseUp);
                    _VisioApplication.MustFlushScopeBeginning += new Microsoft.Office.Interop.Visio.EApplication_MustFlushScopeBeginningEventHandler(_VisioApplication_MustFlushScopeBeginning);
                    _VisioApplication.MustFlushScopeEnded += new Microsoft.Office.Interop.Visio.EApplication_MustFlushScopeEndedEventHandler(_VisioApplication_MustFlushScopeEnded);
                    _VisioApplication.NoEventsPending += new Microsoft.Office.Interop.Visio.EApplication_NoEventsPendingEventHandler(_VisioApplication_NoEventsPending);
                    _VisioApplication.OnKeystrokeMessageForAddon += new Microsoft.Office.Interop.Visio.EApplication_OnKeystrokeMessageForAddonEventHandler(_VisioApplication_OnKeystrokeMessageForAddon);
                    _VisioApplication.PageAdded += new Microsoft.Office.Interop.Visio.EApplication_PageAddedEventHandler(_VisioApplication_PageAdded);
                    _VisioApplication.PageChanged += new Microsoft.Office.Interop.Visio.EApplication_PageChangedEventHandler(_VisioApplication_PageChanged);
                    _VisioApplication.PageDeleteCanceled += new Microsoft.Office.Interop.Visio.EApplication_PageDeleteCanceledEventHandler(_VisioApplication_PageDeleteCanceled);
                    _VisioApplication.QueryCancelConvertToGroup += new Microsoft.Office.Interop.Visio.EApplication_QueryCancelConvertToGroupEventHandler(_VisioApplication_QueryCancelConvertToGroup);
                    _VisioApplication.QueryCancelDocumentClose += new Microsoft.Office.Interop.Visio.EApplication_QueryCancelDocumentCloseEventHandler(_VisioApplication_QueryCancelDocumentClose);
                    _VisioApplication.QueryCancelGroup += new Microsoft.Office.Interop.Visio.EApplication_QueryCancelGroupEventHandler(_VisioApplication_QueryCancelGroup);
                    _VisioApplication.QueryCancelMasterDelete += new Microsoft.Office.Interop.Visio.EApplication_QueryCancelMasterDeleteEventHandler(_VisioApplication_QueryCancelMasterDelete);
                    _VisioApplication.QueryCancelPageDelete += new Microsoft.Office.Interop.Visio.EApplication_QueryCancelPageDeleteEventHandler(_VisioApplication_QueryCancelPageDelete);
                    _VisioApplication.QueryCancelQuit += new Microsoft.Office.Interop.Visio.EApplication_QueryCancelQuitEventHandler(_VisioApplication_QueryCancelQuit);
                    _VisioApplication.QueryCancelSelectionDelete += new Microsoft.Office.Interop.Visio.EApplication_QueryCancelSelectionDeleteEventHandler(_VisioApplication_QueryCancelSelectionDelete);
                    _VisioApplication.QueryCancelStyleDelete += new Microsoft.Office.Interop.Visio.EApplication_QueryCancelStyleDeleteEventHandler(_VisioApplication_QueryCancelStyleDelete);
                    _VisioApplication.QueryCancelSuspend += new Microsoft.Office.Interop.Visio.EApplication_QueryCancelSuspendEventHandler(_VisioApplication_QueryCancelSuspend);
                    _VisioApplication.QueryCancelSuspendEvents += new Microsoft.Office.Interop.Visio.EApplication_QueryCancelSuspendEventsEventHandler(_VisioApplication_QueryCancelSuspendEvents);
                    _VisioApplication.QueryCancelUngroup += new Microsoft.Office.Interop.Visio.EApplication_QueryCancelUngroupEventHandler(_VisioApplication_QueryCancelUngroup);
                    _VisioApplication.QueryCancelWindowClose += new Microsoft.Office.Interop.Visio.EApplication_QueryCancelWindowCloseEventHandler(_VisioApplication_QueryCancelWindowClose);
                    _VisioApplication.QuitCanceled += new Microsoft.Office.Interop.Visio.EApplication_QuitCanceledEventHandler(_VisioApplication_QuitCanceled);
                    _VisioApplication.RuleSetValidated += new Microsoft.Office.Interop.Visio.EApplication_RuleSetValidatedEventHandler(_VisioApplication_RuleSetValidated);
                    _VisioApplication.RunModeEntered += new Microsoft.Office.Interop.Visio.EApplication_RunModeEnteredEventHandler(_VisioApplication_RunModeEntered);
                    _VisioApplication.SelectionAdded += new Microsoft.Office.Interop.Visio.EApplication_SelectionAddedEventHandler(_VisioApplication_SelectionAdded);
                    _VisioApplication.SelectionChanged += new Microsoft.Office.Interop.Visio.EApplication_SelectionChangedEventHandler(_VisioApplication_SelectionChanged);
                    _VisioApplication.SelectionDeleteCanceled += new Microsoft.Office.Interop.Visio.EApplication_SelectionDeleteCanceledEventHandler(_VisioApplication_SelectionDeleteCanceled);
                    _VisioApplication.ShapeAdded += new Microsoft.Office.Interop.Visio.EApplication_ShapeAddedEventHandler(_VisioApplication_ShapeAdded);
                    _VisioApplication.ShapeChanged += new Microsoft.Office.Interop.Visio.EApplication_ShapeChangedEventHandler(_VisioApplication_ShapeChanged);
                    _VisioApplication.ShapeDataGraphicChanged += new Microsoft.Office.Interop.Visio.EApplication_ShapeDataGraphicChangedEventHandler(_VisioApplication_ShapeDataGraphicChanged);
                    _VisioApplication.ShapeExitedTextEdit += new Microsoft.Office.Interop.Visio.EApplication_ShapeExitedTextEditEventHandler(_VisioApplication_ShapeExitedTextEdit);
                    _VisioApplication.ShapeLinkAdded += new Microsoft.Office.Interop.Visio.EApplication_ShapeLinkAddedEventHandler(_VisioApplication_ShapeLinkAdded);
                    _VisioApplication.ShapeLinkDeleted += new Microsoft.Office.Interop.Visio.EApplication_ShapeLinkDeletedEventHandler(_VisioApplication_ShapeLinkDeleted);
                    _VisioApplication.ShapeParentChanged += new Microsoft.Office.Interop.Visio.EApplication_ShapeParentChangedEventHandler(_VisioApplication_ShapeParentChanged);
                    _VisioApplication.StyleAdded += new Microsoft.Office.Interop.Visio.EApplication_StyleAddedEventHandler(_VisioApplication_StyleAdded);
                    _VisioApplication.StyleChanged += new Microsoft.Office.Interop.Visio.EApplication_StyleChangedEventHandler(_VisioApplication_StyleChanged);
                    _VisioApplication.StyleDeleteCanceled += new Microsoft.Office.Interop.Visio.EApplication_StyleDeleteCanceledEventHandler(_VisioApplication_StyleDeleteCanceled);
                    _VisioApplication.SuspendCanceled += new Microsoft.Office.Interop.Visio.EApplication_SuspendCanceledEventHandler(_VisioApplication_SuspendCanceled);
                    _VisioApplication.SuspendEventsCanceled += new Microsoft.Office.Interop.Visio.EApplication_SuspendEventsCanceledEventHandler(_VisioApplication_SuspendEventsCanceled);
                    _VisioApplication.TextChanged += new Microsoft.Office.Interop.Visio.EApplication_TextChangedEventHandler(_VisioApplication_TextChanged);
                    _VisioApplication.UngroupCanceled += new Microsoft.Office.Interop.Visio.EApplication_UngroupCanceledEventHandler(_VisioApplication_UngroupCanceled);
                    _VisioApplication.ViewChanged += new Microsoft.Office.Interop.Visio.EApplication_ViewChangedEventHandler(_VisioApplication_ViewChanged);
                    _VisioApplication.WindowActivated += new Microsoft.Office.Interop.Visio.EApplication_WindowActivatedEventHandler(_VisioApplication_WindowActivated);
                    _VisioApplication.WindowChanged += new Microsoft.Office.Interop.Visio.EApplication_WindowChangedEventHandler(_VisioApplication_WindowChanged);
                    _VisioApplication.WindowCloseCanceled += new Microsoft.Office.Interop.Visio.EApplication_WindowCloseCanceledEventHandler(_VisioApplication_WindowCloseCanceled);
                    _VisioApplication.WindowOpened += new Microsoft.Office.Interop.Visio.EApplication_WindowOpenedEventHandler(_VisioApplication_WindowOpened);
                    _VisioApplication.WindowTurnedToPage += new Microsoft.Office.Interop.Visio.EApplication_WindowTurnedToPageEventHandler(_VisioApplication_WindowTurnedToPage);
                }
            }
        }

        short countBeforeMasterDelete;
        void _VisioApplication_BeforeMasterDelete(Microsoft.Office.Interop.Visio.Master Master)
        {
            DisplayInWatchWindow(countBeforeMasterDelete++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countTextChange;
        void _VisioApplication_TextChanged(Microsoft.Office.Interop.Visio.Shape Shape)
        {
            DisplayInWatchWindow(countTextChange++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countWindowTurnedToPage;
        void _VisioApplication_WindowTurnedToPage(Microsoft.Office.Interop.Visio.Window Window)
        {
            DisplayInWatchWindow(countWindowTurnedToPage++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countWindowOpened;
        void _VisioApplication_WindowOpened(Microsoft.Office.Interop.Visio.Window Window)
        {
            DisplayInWatchWindow(countWindowOpened++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countWindowCloseCanceled;
        void _VisioApplication_WindowCloseCanceled(Microsoft.Office.Interop.Visio.Window Window)
        {
            DisplayInWatchWindow(countWindowCloseCanceled++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countWindowChanged;
        void _VisioApplication_WindowChanged(Microsoft.Office.Interop.Visio.Window Window)
        {
            DisplayInWatchWindow(countWindowChanged++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countWindowActivated;
        void _VisioApplication_WindowActivated(Microsoft.Office.Interop.Visio.Window Window)
        {
            DisplayInWatchWindow(countWindowActivated++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countViewChanged;
        void _VisioApplication_ViewChanged(Microsoft.Office.Interop.Visio.Window Window)
        {
            DisplayInWatchWindow(countViewChanged++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countUngroupCanceled;
        void _VisioApplication_UngroupCanceled(Microsoft.Office.Interop.Visio.Selection Selection)
        {
            DisplayInWatchWindow(countUngroupCanceled++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countSuspendEventsCanceled;
        void _VisioApplication_SuspendEventsCanceled(Microsoft.Office.Interop.Visio.Application app)
        {
            DisplayInWatchWindow(countSuspendEventsCanceled++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countSuspendCanceled;
        void _VisioApplication_SuspendCanceled(Microsoft.Office.Interop.Visio.Application app)
        {
            DisplayInWatchWindow(countSuspendCanceled++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countStyleDeleteCanceled;
        void _VisioApplication_StyleDeleteCanceled(Microsoft.Office.Interop.Visio.Style Style)
        {
            DisplayInWatchWindow(countStyleDeleteCanceled++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countStyleChanged;
        void _VisioApplication_StyleChanged(Microsoft.Office.Interop.Visio.Style Style)
        {
            DisplayInWatchWindow(countStyleChanged++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countStyleAdded;
        void _VisioApplication_StyleAdded(Microsoft.Office.Interop.Visio.Style Style)
        {
            DisplayInWatchWindow(countStyleAdded++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countShapeParentChanged;
        void _VisioApplication_ShapeParentChanged(Microsoft.Office.Interop.Visio.Shape Shape)
        {
            DisplayInWatchWindow(countShapeParentChanged++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countShapeLinkDeleted;
        void _VisioApplication_ShapeLinkDeleted(Microsoft.Office.Interop.Visio.Shape Shape, int DataRecordsetID, int DataRowID)
        {
            DisplayInWatchWindow(countShapeLinkDeleted++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countShapeLinkAdded;
        void _VisioApplication_ShapeLinkAdded(Microsoft.Office.Interop.Visio.Shape Shape, int DataRecordsetID, int DataRowID)
        {
            DisplayInWatchWindow(countShapeLinkAdded++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countShapeExitedTextEdit;
        void _VisioApplication_ShapeExitedTextEdit(Microsoft.Office.Interop.Visio.Shape Shape)
        {
            DisplayInWatchWindow(countShapeExitedTextEdit++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countShapeDataGraphicChanged;
        void _VisioApplication_ShapeDataGraphicChanged(Microsoft.Office.Interop.Visio.Shape Shape)
        {
            DisplayInWatchWindow(countShapeDataGraphicChanged++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countShapeChanged;
        void _VisioApplication_ShapeChanged(Microsoft.Office.Interop.Visio.Shape Shape)
        {
            DisplayInWatchWindow(countShapeChanged++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countShapeAdded;
        void _VisioApplication_ShapeAdded(Microsoft.Office.Interop.Visio.Shape Shape)
        {
            DisplayInWatchWindow(countShapeAdded++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countSelectionDeleteCanceled;
        void _VisioApplication_SelectionDeleteCanceled(Microsoft.Office.Interop.Visio.Selection Selection)
        {
            DisplayInWatchWindow(countSelectionDeleteCanceled++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countSelectionChanged;
        void _VisioApplication_SelectionChanged(Microsoft.Office.Interop.Visio.Window Window)
        {
            DisplayInWatchWindow(countSelectionChanged++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countSelectionAdded;
        void _VisioApplication_SelectionAdded(Microsoft.Office.Interop.Visio.Selection Selection)
        {
            DisplayInWatchWindow(countSelectionAdded++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countRunModeEntered;
        void _VisioApplication_RunModeEntered(Microsoft.Office.Interop.Visio.Document Doc)
        {
            DisplayInWatchWindow(countRunModeEntered++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countRuleSetValidated;
        void _VisioApplication_RuleSetValidated(Microsoft.Office.Interop.Visio.ValidationRuleSet RuleSet)
        {
            DisplayInWatchWindow(countRuleSetValidated++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countQuitCanceled;
        void _VisioApplication_QuitCanceled(Microsoft.Office.Interop.Visio.Application app)
        {
            DisplayInWatchWindow(countQuitCanceled++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countQueryCancelWindowClose;
        bool _VisioApplication_QueryCancelWindowClose(Microsoft.Office.Interop.Visio.Window Window)
        {
            DisplayInWatchWindow(countQueryCancelWindowClose++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
            return false;
        }
        
        short countQueryCancelUngroup;
        bool _VisioApplication_QueryCancelUngroup(Microsoft.Office.Interop.Visio.Selection Selection)
        {
            DisplayInWatchWindow(countQueryCancelUngroup++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
            return false;
        }
        
        short countQueryCancelSuspendEvents;
        bool _VisioApplication_QueryCancelSuspendEvents(Microsoft.Office.Interop.Visio.Application app)
        {
            DisplayInWatchWindow(countQueryCancelSuspendEvents++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
            return false;
        }
        
        short countQueryCancelSuspend;
        bool _VisioApplication_QueryCancelSuspend(Microsoft.Office.Interop.Visio.Application app)
        {
            DisplayInWatchWindow(countQueryCancelSuspend++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
            return false;
        }
        
        short countQueryCancelStyleDelete;
        bool _VisioApplication_QueryCancelStyleDelete(Microsoft.Office.Interop.Visio.Style Style)
        {
            DisplayInWatchWindow(countQueryCancelStyleDelete++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
            return false;
        }
        
        short countQueryCancelSelectionDelete;
        bool _VisioApplication_QueryCancelSelectionDelete(Microsoft.Office.Interop.Visio.Selection Selection)
        {
            DisplayInWatchWindow(countQueryCancelSelectionDelete++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
            return false;
        }
        
        short countQueryCancelQuit;
        bool _VisioApplication_QueryCancelQuit(Microsoft.Office.Interop.Visio.Application app)
        {
            DisplayInWatchWindow(countQueryCancelQuit++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
            return false;
        }
        
        short countQueryCancelPageDelete;
        bool _VisioApplication_QueryCancelPageDelete(Microsoft.Office.Interop.Visio.Page Page)
        {
            DisplayInWatchWindow(countQueryCancelPageDelete++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
            return false;
        }
        
        short countQueryCancelMasterDelete;
        bool _VisioApplication_QueryCancelMasterDelete(Microsoft.Office.Interop.Visio.Master Master)
        {
            DisplayInWatchWindow(countQueryCancelMasterDelete++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
            return false;
        }
        
        short countQueryCancelGroup;
        bool _VisioApplication_QueryCancelGroup(Microsoft.Office.Interop.Visio.Selection Selection)
        {
            DisplayInWatchWindow(countQueryCancelGroup++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
            return false;
        }
        
        short countQueryCancelDocumentClose;
        bool _VisioApplication_QueryCancelDocumentClose(Microsoft.Office.Interop.Visio.Document Doc)
        {
            DisplayInWatchWindow(countQueryCancelDocumentClose++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
            return false;
        }
        
        short countQueryCancelConvertToGroup;
        bool _VisioApplication_QueryCancelConvertToGroup(Microsoft.Office.Interop.Visio.Selection Selection)
        {
            DisplayInWatchWindow(countQueryCancelConvertToGroup++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
            return false;
        }
        
        short countPageDeleteCanceled;
        void _VisioApplication_PageDeleteCanceled(Microsoft.Office.Interop.Visio.Page Page)
        {
            DisplayInWatchWindow(countPageDeleteCanceled++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countPageChanged;
        void _VisioApplication_PageChanged(Microsoft.Office.Interop.Visio.Page Page)
        {
            DisplayInWatchWindow(countPageChanged++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countPageAdded;
        void _VisioApplication_PageAdded(Microsoft.Office.Interop.Visio.Page Page)
        {
            DisplayInWatchWindow(countPageAdded++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countOnKeystrokeMessageForAddon;
        bool _VisioApplication_OnKeystrokeMessageForAddon(Microsoft.Office.Interop.Visio.MSGWrap MSG)
        {
            DisplayInWatchWindow(countOnKeystrokeMessageForAddon++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
            return false;
        }
        
        short countNoEventsPending;
        void _VisioApplication_NoEventsPending(Microsoft.Office.Interop.Visio.Application app)
        {
            if (Common.DisplayChattyEvents)
            {
                DisplayInWatchWindow(countNoEventsPending++, System.Reflection.MethodInfo.GetCurrentMethod().Name);            	;
            }
            else
            {
                countNoEventsPending++;
            }
        }
        
        short countMustFlushScopeEnded;
        void _VisioApplication_MustFlushScopeEnded(Microsoft.Office.Interop.Visio.Application app)
        {
            DisplayInWatchWindow(countMustFlushScopeEnded++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countMustFlushScopeBeginning;
        void _VisioApplication_MustFlushScopeBeginning(Microsoft.Office.Interop.Visio.Application app)
        {
            DisplayInWatchWindow(countMustFlushScopeBeginning++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countMouseDown;
        void _VisioApplication_MouseDown(int Button, int KeyButtonState, double x, double y, ref bool CancelDefault)
        {
            if (Common.DisplayChattyEvents)
            {
                DisplayInWatchWindow(countMouseDown++, System.Reflection.MethodInfo.GetCurrentMethod().Name);            	;
            }
            else
            {
                countMouseDown++;
            }
        }
        
        short countMouseUp;
        void _VisioApplication_MouseUp(int Button, int KeyButtonState, double x, double y, ref bool CancelDefault)
        {
            if (Common.DisplayChattyEvents)
            {
                DisplayInWatchWindow(countMouseUp++, System.Reflection.MethodInfo.GetCurrentMethod().Name);            	;
            }
            else
            {
                countMouseUp++;
            }
        }
        
        short countMouseMove;
        void _VisioApplication_MouseMove(int Button, int KeyButtonState, double x, double y, ref bool CancelDefault)
        {
            if (Common.DisplayChattyEvents)
            {
                DisplayInWatchWindow(countMouseMove++, System.Reflection.MethodInfo.GetCurrentMethod().Name);            	;
            }
            else
            {
                countMouseMove++;
            }
        }
        
        short countMasterDeleteCanceled;
        void _VisioApplication_MasterDeleteCanceled(Microsoft.Office.Interop.Visio.Master Master)
        {
            DisplayInWatchWindow(countMasterDeleteCanceled++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countMasterChanged;
        void _VisioApplication_MasterChanged(Microsoft.Office.Interop.Visio.Master Master)
        {
            DisplayInWatchWindow(countMasterChanged++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countMasterAdded;
        void _VisioApplication_MasterAdded(Microsoft.Office.Interop.Visio.Master Master)
        {
            DisplayInWatchWindow(countMasterAdded++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countMarkerEvent;
        void _VisioApplication_MarkerEvent(Microsoft.Office.Interop.Visio.Application app, int SequenceNum, string ContextString)
        {
            DisplayInWatchWindow(countMarkerEvent++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countKeyUp;
        void _VisioApplication_KeyUp(int KeyCode, int KeyButtonState, ref bool CancelDefault)
        {
            DisplayInWatchWindow(countKeyUp++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countKeyPress;
        void _VisioApplication_KeyPress(int KeyAscii, ref bool CancelDefault)
        {
            DisplayInWatchWindow(countKeyPress++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countKeyDown;
        void _VisioApplication_KeyDown(int KeyCode, int KeyButtonState, ref bool CancelDefault)
        {
            DisplayInWatchWindow(countKeyDown++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countGroupCanceled;
        void _VisioApplication_GroupCanceled(Microsoft.Office.Interop.Visio.Selection Selection)
        {
            DisplayInWatchWindow(countGroupCanceled++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countFormulaChanged;
        void _VisioApplication_FormulaChanged(Microsoft.Office.Interop.Visio.Cell Cell)
        {
            DisplayInWatchWindow(countFormulaChanged++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countExitScope;
        void _VisioApplication_ExitScope(Microsoft.Office.Interop.Visio.Application app, int nScopeID, string bstrDescription, bool bErrOrCancelled)
        {
            DisplayInWatchWindow(countExitScope++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countEnterScope;
        void _VisioApplication_EnterScope(Microsoft.Office.Interop.Visio.Application app, int nScopeID, string bstrDescription)
        {
            DisplayInWatchWindow(countEnterScope++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countDocumentSavedAs;
        void _VisioApplication_DocumentSavedAs(Microsoft.Office.Interop.Visio.Document Doc)
        {
            DisplayInWatchWindow(countDocumentSavedAs++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countDocumentSaved;
        void _VisioApplication_DocumentSaved(Microsoft.Office.Interop.Visio.Document Doc)
        {
            DisplayInWatchWindow(countDocumentSaved++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countDocumentOpened;
        void _VisioApplication_DocumentOpened(Microsoft.Office.Interop.Visio.Document Doc)
        {
            DisplayInWatchWindow(countDocumentOpened++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countDocumentCreated;
        void _VisioApplication_DocumentCreated(Microsoft.Office.Interop.Visio.Document Doc)
        {
            DisplayInWatchWindow(countDocumentCreated++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countDocumentCloseCanceled;
        void _VisioApplication_DocumentCloseCanceled(Microsoft.Office.Interop.Visio.Document Doc)
        {
            DisplayInWatchWindow(countDocumentCloseCanceled++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countDocumentChanged;
        void _VisioApplication_DocumentChanged(Microsoft.Office.Interop.Visio.Document Doc)
        {
            DisplayInWatchWindow(countDocumentChanged++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countDesignModeEntered;
        void _VisioApplication_DesignModeEntered(Microsoft.Office.Interop.Visio.Document Doc)
        {
            DisplayInWatchWindow(countDesignModeEntered++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countDataRecordsetChanged;
        void _VisioApplication_DataRecordsetChanged(Microsoft.Office.Interop.Visio.DataRecordsetChangedEvent DataRecordsetChanged)
        {
            DisplayInWatchWindow(countDataRecordsetChanged++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countDataRecordsetAdded;
        void _VisioApplication_DataRecordsetAdded(Microsoft.Office.Interop.Visio.DataRecordset DataRecordset)
        {
            DisplayInWatchWindow(countDataRecordsetAdded++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countConvertToGroupCanceled;
        void _VisioApplication_ConvertToGroupCanceled(Microsoft.Office.Interop.Visio.Selection Selection)
        {
            DisplayInWatchWindow(countConvertToGroupCanceled++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countContainerRelationshipDeleted;
        void _VisioApplication_ContainerRelationshipDeleted(Microsoft.Office.Interop.Visio.RelatedShapePairEvent ShapePair)
        {
            DisplayInWatchWindow(countContainerRelationshipDeleted++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countContainerRelationshipAdded;
        void _VisioApplication_ContainerRelationshipAdded(Microsoft.Office.Interop.Visio.RelatedShapePairEvent ShapePair)
        {
            DisplayInWatchWindow(countContainerRelationshipAdded++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countConnectionsDeleted;
        void _VisioApplication_ConnectionsDeleted(Microsoft.Office.Interop.Visio.Connects Connects)
        {
            DisplayInWatchWindow(countConnectionsDeleted++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countConnectionsAdded;
        void _VisioApplication_ConnectionsAdded(Microsoft.Office.Interop.Visio.Connects Connects)
        {
            DisplayInWatchWindow(countConnectionsAdded++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countCellChanged;
        void _VisioApplication_CellChanged(Microsoft.Office.Interop.Visio.Cell Cell)
        {
            DisplayInWatchWindow(countCellChanged++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countCalloutRelationshipDeleted;
        void _VisioApplication_CalloutRelationshipDeleted(Microsoft.Office.Interop.Visio.RelatedShapePairEvent ShapePair)
        {
            DisplayInWatchWindow(countCalloutRelationshipDeleted++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countCalloutRelationshipAdded;
        void _VisioApplication_CalloutRelationshipAdded(Microsoft.Office.Interop.Visio.RelatedShapePairEvent ShapePair)
        {
            DisplayInWatchWindow(countCalloutRelationshipAdded++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countBeforeWindowSelDelete;
        void _VisioApplication_BeforeWindowSelDelete(Microsoft.Office.Interop.Visio.Window Window)
        {
            DisplayInWatchWindow(countBeforeWindowSelDelete++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countBeforeWindowPageTurn;
        void _VisioApplication_BeforeWindowPageTurn(Microsoft.Office.Interop.Visio.Window Window)
        {
            DisplayInWatchWindow(countBeforeWindowPageTurn++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countBeforeWindowClosed;
        void _VisioApplication_BeforeWindowClosed(Microsoft.Office.Interop.Visio.Window Window)
        {
            DisplayInWatchWindow(countBeforeWindowClosed++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countBeforeSuspendEvents;
        void _VisioApplication_BeforeSuspendEvents(Microsoft.Office.Interop.Visio.Application app)
        {
            DisplayInWatchWindow(countBeforeSuspendEvents++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countBeforeSuspend;
        void _VisioApplication_BeforeSuspend(Microsoft.Office.Interop.Visio.Application app)
        {
            DisplayInWatchWindow(countBeforeSuspend++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countBeforeStyleDelete;
        void _VisioApplication_BeforeStyleDelete(Microsoft.Office.Interop.Visio.Style Style)
        {
            DisplayInWatchWindow(countBeforeStyleDelete++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countBeforeShapeTextEdit;
        void _VisioApplication_BeforeShapeTextEdit(Microsoft.Office.Interop.Visio.Shape Shape)
        {
            DisplayInWatchWindow(countBeforeShapeTextEdit++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countBeforeShapeDelete;
        void _VisioApplication_BeforeShapeDelete(Microsoft.Office.Interop.Visio.Shape Shape)
        {
            DisplayInWatchWindow(countBeforeShapeDelete++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countBeforeSelectionDelete;
        void _VisioApplication_BeforeSelectionDelete(Microsoft.Office.Interop.Visio.Selection Selection)
        {
            DisplayInWatchWindow(countBeforeSelectionDelete++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countBeforeQuit;
        void _VisioApplication_BeforeQuit(Microsoft.Office.Interop.Visio.Application app)
        {
            DisplayInWatchWindow(countBeforeQuit++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countBeforePageDelete;
        void _VisioApplication_BeforePageDelete(Microsoft.Office.Interop.Visio.Page Page)
        {
            DisplayInWatchWindow(countBeforePageDelete++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countBeforeModal;
        void _VisioApplication_BeforeModal(Microsoft.Office.Interop.Visio.Application app)
        {
            DisplayInWatchWindow(countBeforeModal++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countBeforeDocumentSaveAs;
        void _VisioApplication_BeforeDocumentSaveAs(Microsoft.Office.Interop.Visio.Document Doc)
        {
            DisplayInWatchWindow(countBeforeDocumentSaveAs++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countBeforeDocumentSave;
        void _VisioApplication_BeforeDocumentSave(Microsoft.Office.Interop.Visio.Document Doc)
        {
            DisplayInWatchWindow(countBeforeDocumentSave++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countBeforeDocumentClose;
        void _VisioApplication_BeforeDocumentClose(Microsoft.Office.Interop.Visio.Document Doc)
        {
            DisplayInWatchWindow(countBeforeDocumentClose++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countBeforeDataRecordsetDelete;
        void _VisioApplication_BeforeDataRecordsetDelete(Microsoft.Office.Interop.Visio.DataRecordset DataRecordset)
        {
            DisplayInWatchWindow(countBeforeDataRecordsetDelete++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countAppObjDeactivated;
        void _VisioApplication_AppObjDeactivated(Microsoft.Office.Interop.Visio.Application app)
        {
            DisplayInWatchWindow(countAppObjDeactivated++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countAppObjActivated;
        void _VisioApplication_AppObjActivated(Microsoft.Office.Interop.Visio.Application app)
        {
            DisplayInWatchWindow(countAppObjActivated++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countAppDeactivated;
        void _VisioApplication_AppDeactivated(Microsoft.Office.Interop.Visio.Application app)
        {
            DisplayInWatchWindow(countAppDeactivated++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countAppActivated;
        void _VisioApplication_AppActivated(Microsoft.Office.Interop.Visio.Application app)
        {
            DisplayInWatchWindow(countAppActivated++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countAfterResumeEvents;
        void _VisioApplication_AfterResumeEvents(Microsoft.Office.Interop.Visio.Application app)
        {
            DisplayInWatchWindow(countAfterResumeEvents++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countAfterResume;
        void _VisioApplication_AfterResume(Microsoft.Office.Interop.Visio.Application app)
        {
            DisplayInWatchWindow(countAfterResume++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countAfterRemoveHiddenInformation;
        void _VisioApplication_AfterRemoveHiddenInformation(Microsoft.Office.Interop.Visio.Document Doc)
        {
            DisplayInWatchWindow(countAfterRemoveHiddenInformation++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        
        short countAfterModal;
        void _VisioApplication_AfterModal(Microsoft.Office.Interop.Visio.Application app)
        {
            DisplayInWatchWindow(countAfterModal++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
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
