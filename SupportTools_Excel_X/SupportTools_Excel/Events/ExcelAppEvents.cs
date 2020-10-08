using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelHlp = AddinHelper.Excel;

namespace SupportTools_Excel.Events
{
    class ExcelAppEvents
    {
        private Microsoft.Office.Interop.Excel.Application _ExcelApplication;
        public Microsoft.Office.Interop.Excel.Application ExcelApplication
        {
            get
            {
                return _ExcelApplication;
            }
            set
            {
                if(_ExcelApplication != null)
                {
                    // Should remove all the event handlers;
                }

                _ExcelApplication = value;

                if(_ExcelApplication != null)
                {
                    _ExcelApplication.AfterCalculate += new Microsoft.Office.Interop.Excel.AppEvents_AfterCalculateEventHandler(_ExcelApplication_AfterCalculate);
                    _ExcelApplication.ProtectedViewWindowActivate += new Microsoft.Office.Interop.Excel.AppEvents_ProtectedViewWindowActivateEventHandler(_ExcelApplication_ProtectedViewWindowActivate);
                    _ExcelApplication.ProtectedViewWindowBeforeClose += new Microsoft.Office.Interop.Excel.AppEvents_ProtectedViewWindowBeforeCloseEventHandler(_ExcelApplication_ProtectedViewWindowBeforeClose);
                    _ExcelApplication.ProtectedViewWindowBeforeEdit += new Microsoft.Office.Interop.Excel.AppEvents_ProtectedViewWindowBeforeEditEventHandler(_ExcelApplication_ProtectedViewWindowBeforeEdit);
                    _ExcelApplication.ProtectedViewWindowDeactivate += new Microsoft.Office.Interop.Excel.AppEvents_ProtectedViewWindowDeactivateEventHandler(_ExcelApplication_ProtectedViewWindowDeactivate);
                    _ExcelApplication.ProtectedViewWindowOpen += new Microsoft.Office.Interop.Excel.AppEvents_ProtectedViewWindowOpenEventHandler(_ExcelApplication_ProtectedViewWindowOpen);
                    _ExcelApplication.ProtectedViewWindowResize += new Microsoft.Office.Interop.Excel.AppEvents_ProtectedViewWindowResizeEventHandler(_ExcelApplication_ProtectedViewWindowResize);
                    _ExcelApplication.SheetActivate += new Microsoft.Office.Interop.Excel.AppEvents_SheetActivateEventHandler(_ExcelApplication_SheetActivate);
                    _ExcelApplication.SheetBeforeDoubleClick += new Microsoft.Office.Interop.Excel.AppEvents_SheetBeforeDoubleClickEventHandler(_ExcelApplication_SheetBeforeDoubleClick);
                    _ExcelApplication.SheetBeforeRightClick += new Microsoft.Office.Interop.Excel.AppEvents_SheetBeforeRightClickEventHandler(_ExcelApplication_SheetBeforeRightClick);
                    _ExcelApplication.SheetCalculate += new Microsoft.Office.Interop.Excel.AppEvents_SheetCalculateEventHandler(_ExcelApplication_SheetCalculate);
                    _ExcelApplication.SheetChange += new Microsoft.Office.Interop.Excel.AppEvents_SheetChangeEventHandler(_ExcelApplication_SheetChange);
                    _ExcelApplication.SheetDeactivate += new Microsoft.Office.Interop.Excel.AppEvents_SheetDeactivateEventHandler(_ExcelApplication_SheetDeactivate);
                    _ExcelApplication.SheetFollowHyperlink += new Microsoft.Office.Interop.Excel.AppEvents_SheetFollowHyperlinkEventHandler(_ExcelApplication_SheetFollowHyperlink);
                    _ExcelApplication.SheetPivotTableAfterValueChange += new Microsoft.Office.Interop.Excel.AppEvents_SheetPivotTableAfterValueChangeEventHandler(_ExcelApplication_SheetPivotTableAfterValueChange);
                    _ExcelApplication.SheetPivotTableBeforeAllocateChanges += new Microsoft.Office.Interop.Excel.AppEvents_SheetPivotTableBeforeAllocateChangesEventHandler(_ExcelApplication_SheetPivotTableBeforeAllocateChanges);
                    _ExcelApplication.SheetPivotTableBeforeCommitChanges += new Microsoft.Office.Interop.Excel.AppEvents_SheetPivotTableBeforeCommitChangesEventHandler(_ExcelApplication_SheetPivotTableBeforeCommitChanges);
                    _ExcelApplication.SheetPivotTableBeforeDiscardChanges += new Microsoft.Office.Interop.Excel.AppEvents_SheetPivotTableBeforeDiscardChangesEventHandler(_ExcelApplication_SheetPivotTableBeforeDiscardChanges);
                    _ExcelApplication.SheetPivotTableUpdate += new Microsoft.Office.Interop.Excel.AppEvents_SheetPivotTableUpdateEventHandler(_ExcelApplication_SheetPivotTableUpdate);
                    _ExcelApplication.SheetSelectionChange += new Microsoft.Office.Interop.Excel.AppEvents_SheetSelectionChangeEventHandler(_ExcelApplication_SheetSelectionChange);
                    _ExcelApplication.WindowActivate += new Microsoft.Office.Interop.Excel.AppEvents_WindowActivateEventHandler(_ExcelApplication_WindowActivate);
                    _ExcelApplication.WindowDeactivate += new Microsoft.Office.Interop.Excel.AppEvents_WindowDeactivateEventHandler(_ExcelApplication_WindowDeactivate);
                    _ExcelApplication.WindowResize += new Microsoft.Office.Interop.Excel.AppEvents_WindowResizeEventHandler(_ExcelApplication_WindowResize);
                    _ExcelApplication.WorkbookActivate += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookActivateEventHandler(_ExcelApplication_WorkbookActivate);
                    _ExcelApplication.WorkbookAddinInstall += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookAddinInstallEventHandler(_ExcelApplication_WorkbookAddinInstall);
                    _ExcelApplication.WorkbookAddinUninstall += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookAddinUninstallEventHandler(_ExcelApplication_WorkbookAddinUninstall);
                    _ExcelApplication.WorkbookAfterSave += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookAfterSaveEventHandler(_ExcelApplication_WorkbookAfterSave);
                    _ExcelApplication.WorkbookAfterXmlExport += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookAfterXmlExportEventHandler(_ExcelApplication_WorkbookAfterXmlExport);
                    _ExcelApplication.WorkbookAfterXmlImport += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookAfterXmlImportEventHandler(_ExcelApplication_WorkbookAfterXmlImport);
                    _ExcelApplication.WorkbookBeforeClose += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookBeforeCloseEventHandler(_ExcelApplication_WorkbookBeforeClose);
                    _ExcelApplication.WorkbookBeforePrint += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookBeforePrintEventHandler(_ExcelApplication_WorkbookBeforePrint);
                    _ExcelApplication.WorkbookBeforeSave += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookBeforeSaveEventHandler(_ExcelApplication_WorkbookBeforeSave);
                    _ExcelApplication.WorkbookBeforeXmlExport += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookBeforeXmlExportEventHandler(_ExcelApplication_WorkbookBeforeXmlExport);
                    _ExcelApplication.WorkbookBeforeXmlImport += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookBeforeXmlImportEventHandler(_ExcelApplication_WorkbookBeforeXmlImport);
                    _ExcelApplication.WorkbookDeactivate += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookDeactivateEventHandler(_ExcelApplication_WorkbookDeactivate);
                    _ExcelApplication.WorkbookNewChart += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookNewChartEventHandler(_ExcelApplication_WorkbookNewChart);
                    _ExcelApplication.WorkbookNewSheet += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookNewSheetEventHandler(_ExcelApplication_WorkbookNewSheet);
                    _ExcelApplication.WorkbookOpen += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookOpenEventHandler(_ExcelApplication_WorkbookOpen);
                    _ExcelApplication.WorkbookPivotTableCloseConnection += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookPivotTableCloseConnectionEventHandler(_ExcelApplication_WorkbookPivotTableCloseConnection);
                    _ExcelApplication.WorkbookPivotTableOpenConnection += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookPivotTableOpenConnectionEventHandler(_ExcelApplication_WorkbookPivotTableOpenConnection);
                    _ExcelApplication.WorkbookRowsetComplete += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookRowsetCompleteEventHandler(_ExcelApplication_WorkbookRowsetComplete);
                    _ExcelApplication.WorkbookSync += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookSyncEventHandler(_ExcelApplication_WorkbookSync);
                }
            }
        }

        short WorkbookSync;
        void _ExcelApplication_WorkbookSync(Microsoft.Office.Interop.Excel.Workbook Wb, Microsoft.Office.Core.MsoSyncEventType SyncEventType)
        {
            DisplayInWatchWindow(WorkbookSync++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short WorkbookRowsetComplete;
        void _ExcelApplication_WorkbookRowsetComplete(Microsoft.Office.Interop.Excel.Workbook Wb, string Description, string Sheet, bool Success)
        {
            DisplayInWatchWindow(WorkbookRowsetComplete++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short WorkbookPivotTableOpenConnection;
        void _ExcelApplication_WorkbookPivotTableOpenConnection(Microsoft.Office.Interop.Excel.Workbook Wb, Microsoft.Office.Interop.Excel.PivotTable Target)
        {
            DisplayInWatchWindow(WorkbookPivotTableOpenConnection++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short WorkbookPivotTableCloseConnection;
        void _ExcelApplication_WorkbookPivotTableCloseConnection(Microsoft.Office.Interop.Excel.Workbook Wb, Microsoft.Office.Interop.Excel.PivotTable Target)
        {
            DisplayInWatchWindow(WorkbookPivotTableCloseConnection++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short WorkbookOpen;
        void _ExcelApplication_WorkbookOpen(Microsoft.Office.Interop.Excel.Workbook Wb)
        {
            DisplayInWatchWindow(WorkbookOpen++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short WorkbookNewSheet;
        void _ExcelApplication_WorkbookNewSheet(Microsoft.Office.Interop.Excel.Workbook Wb, object Sh)
        {
            DisplayInWatchWindow(WorkbookNewSheet++, System.Reflection.MethodInfo.GetCurrentMethod().Name);

            if(ExcelHlp.HasCustomTableOfContents())
            {
                DisplayInWatchWindow(1, System.Reflection.MethodInfo.GetCurrentMethod().Name + " - AddTableOfContents()");
                Actions.Excel_TableOfContents.CreateTableOfContents();
            }
        }

        short WorkbookNewChart;
        void _ExcelApplication_WorkbookNewChart(Microsoft.Office.Interop.Excel.Workbook Wb, Microsoft.Office.Interop.Excel.Chart Ch)
        {
            DisplayInWatchWindow(WorkbookNewChart++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short WorkbookDeactivate;
        void _ExcelApplication_WorkbookDeactivate(Microsoft.Office.Interop.Excel.Workbook Wb)
        {
            DisplayInWatchWindow(WorkbookDeactivate++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short WorkbookBeforeXmlImport;
        void _ExcelApplication_WorkbookBeforeXmlImport(Microsoft.Office.Interop.Excel.Workbook Wb, Microsoft.Office.Interop.Excel.XmlMap Map, string Url, bool IsRefresh, ref bool Cancel)
        {
            DisplayInWatchWindow(WorkbookBeforeXmlImport++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short WorkbookBeforeXmlExport;
        void _ExcelApplication_WorkbookBeforeXmlExport(Microsoft.Office.Interop.Excel.Workbook Wb, Microsoft.Office.Interop.Excel.XmlMap Map, string Url, ref bool Cancel)
        {
            DisplayInWatchWindow(WorkbookBeforeXmlExport++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short WorkbookBeforeSave;
        void _ExcelApplication_WorkbookBeforeSave(Microsoft.Office.Interop.Excel.Workbook Wb, bool SaveAsUI, ref bool Cancel)
        {
            DisplayInWatchWindow(WorkbookBeforeSave++, System.Reflection.MethodInfo.GetCurrentMethod().Name);

            if (ExcelHlp.HasCustomFooter())
            {
            	Actions.Excel_PageFormatting.AddFooter();
            }
        }

        short WorkbookBeforePrint;
        void _ExcelApplication_WorkbookBeforePrint(Microsoft.Office.Interop.Excel.Workbook Wb, ref bool Cancel)
        {
            DisplayInWatchWindow(WorkbookBeforePrint++, System.Reflection.MethodInfo.GetCurrentMethod().Name);

            if (ExcelHlp.HasCustomFooter())
            {
                Actions.Excel_PageFormatting.AddFooter();
            }
        }

        short WorkbookBeforeClose;
        void _ExcelApplication_WorkbookBeforeClose(Microsoft.Office.Interop.Excel.Workbook Wb, ref bool Cancel)
        {
            DisplayInWatchWindow(WorkbookBeforeClose++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short WorkbookAfterXmlImport;
        void _ExcelApplication_WorkbookAfterXmlImport(Microsoft.Office.Interop.Excel.Workbook Wb, Microsoft.Office.Interop.Excel.XmlMap Map, bool IsRefresh, Microsoft.Office.Interop.Excel.XlXmlImportResult Result)
        {
            DisplayInWatchWindow(WorkbookAfterXmlImport++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short WorkbookAfterXmlExport;
        void _ExcelApplication_WorkbookAfterXmlExport(Microsoft.Office.Interop.Excel.Workbook Wb, Microsoft.Office.Interop.Excel.XmlMap Map, string Url, Microsoft.Office.Interop.Excel.XlXmlExportResult Result)
        {
            DisplayInWatchWindow(WorkbookAfterXmlExport++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short WorkbookAfterSave;
        void _ExcelApplication_WorkbookAfterSave(Microsoft.Office.Interop.Excel.Workbook Wb, bool Success)
        {
            DisplayInWatchWindow(WorkbookAfterSave++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short WorkbookAddinUninstall;
        void _ExcelApplication_WorkbookAddinUninstall(Microsoft.Office.Interop.Excel.Workbook Wb)
        {
            DisplayInWatchWindow(WorkbookAddinUninstall++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short WorkbookAddinInstall;
        void _ExcelApplication_WorkbookAddinInstall(Microsoft.Office.Interop.Excel.Workbook Wb)
        {
            DisplayInWatchWindow(WorkbookAddinInstall++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short WorkbookActivate;
        void _ExcelApplication_WorkbookActivate(Microsoft.Office.Interop.Excel.Workbook Wb)
        {
            DisplayInWatchWindow(WorkbookActivate++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short WindowResize;
        void _ExcelApplication_WindowResize(Microsoft.Office.Interop.Excel.Workbook Wb, Microsoft.Office.Interop.Excel.Window Wn)
        {
            DisplayInWatchWindow(WindowResize++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short WindowDeactivate;
        void _ExcelApplication_WindowDeactivate(Microsoft.Office.Interop.Excel.Workbook Wb, Microsoft.Office.Interop.Excel.Window Wn)
        {
            DisplayInWatchWindow(WindowDeactivate++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short WindowActivate;
        void _ExcelApplication_WindowActivate(Microsoft.Office.Interop.Excel.Workbook Wb, Microsoft.Office.Interop.Excel.Window Wn)
        {
            DisplayInWatchWindow(WindowActivate++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short SheetSelectionChange;
        void _ExcelApplication_SheetSelectionChange(object Sh, Microsoft.Office.Interop.Excel.Range Target)
        {
            DisplayInWatchWindow(SheetSelectionChange++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short SheetPivotTableUpdate;
        void _ExcelApplication_SheetPivotTableUpdate(object Sh, Microsoft.Office.Interop.Excel.PivotTable Target)
        {
            DisplayInWatchWindow(SheetPivotTableUpdate++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short SheetPivotTableBeforeDiscardChanges;
        void _ExcelApplication_SheetPivotTableBeforeDiscardChanges(object Sh, Microsoft.Office.Interop.Excel.PivotTable TargetPivotTable, int ValueChangeStart, int ValueChangeEnd)
        {
            DisplayInWatchWindow(SheetPivotTableBeforeDiscardChanges++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short SheetPivotTableBeforeCommitChanges;
        void _ExcelApplication_SheetPivotTableBeforeCommitChanges(object Sh, Microsoft.Office.Interop.Excel.PivotTable TargetPivotTable, int ValueChangeStart, int ValueChangeEnd, ref bool Cancel)
        {
            DisplayInWatchWindow(SheetPivotTableBeforeCommitChanges++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short SheetPivotTableBeforeAllocateChanges;
        void _ExcelApplication_SheetPivotTableBeforeAllocateChanges(object Sh, Microsoft.Office.Interop.Excel.PivotTable TargetPivotTable, int ValueChangeStart, int ValueChangeEnd, ref bool Cancel)
        {
            DisplayInWatchWindow(SheetPivotTableBeforeAllocateChanges++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short SheetPivotTableAfterValueChange;
        void _ExcelApplication_SheetPivotTableAfterValueChange(object Sh, Microsoft.Office.Interop.Excel.PivotTable TargetPivotTable, Microsoft.Office.Interop.Excel.Range TargetRange)
        {
            DisplayInWatchWindow(SheetPivotTableAfterValueChange++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short SheetFollowHyperlink;
        void _ExcelApplication_SheetFollowHyperlink(object Sh, Microsoft.Office.Interop.Excel.Hyperlink Target)
        {
            DisplayInWatchWindow(SheetFollowHyperlink++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short SheetDeactivate;
        void _ExcelApplication_SheetDeactivate(object Sh)
        {
            DisplayInWatchWindow(SheetDeactivate++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short SheetChange;
        void _ExcelApplication_SheetChange(object Sh, Microsoft.Office.Interop.Excel.Range Target)
        {
            DisplayInWatchWindow(SheetChange++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short SheetCalculate;
        void _ExcelApplication_SheetCalculate(object Sh)
        {
            DisplayInWatchWindow(SheetCalculate++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short SheetBeforeRightClick;
        void _ExcelApplication_SheetBeforeRightClick(object Sh, Microsoft.Office.Interop.Excel.Range Target, ref bool Cancel)
        {
            DisplayInWatchWindow(SheetBeforeRightClick++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short SheetBeforeDoubleClick;
        void _ExcelApplication_SheetBeforeDoubleClick(object Sh, Microsoft.Office.Interop.Excel.Range Target, ref bool Cancel)
        {
            DisplayInWatchWindow(SheetBeforeDoubleClick++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short SheetActivate;
        void _ExcelApplication_SheetActivate(object Sh)
        {
            DisplayInWatchWindow(SheetActivate++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProtectedViewWindowResize;
        void _ExcelApplication_ProtectedViewWindowResize(Microsoft.Office.Interop.Excel.ProtectedViewWindow Pvw)
        {
            DisplayInWatchWindow(ProtectedViewWindowResize++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProtectedViewWindowOpen;
        void _ExcelApplication_ProtectedViewWindowOpen(Microsoft.Office.Interop.Excel.ProtectedViewWindow Pvw)
        {
            DisplayInWatchWindow(ProtectedViewWindowOpen++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProtectedViewWindowDeactivate;
        void _ExcelApplication_ProtectedViewWindowDeactivate(Microsoft.Office.Interop.Excel.ProtectedViewWindow Pvw)
        {
            DisplayInWatchWindow(ProtectedViewWindowDeactivate++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProtectedViewWindowBeforeEdit;
        void _ExcelApplication_ProtectedViewWindowBeforeEdit(Microsoft.Office.Interop.Excel.ProtectedViewWindow Pvw, ref bool Cancel)
        {
            DisplayInWatchWindow(ProtectedViewWindowBeforeEdit++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProtectedViewWindowBeforeClose;
        void _ExcelApplication_ProtectedViewWindowBeforeClose(Microsoft.Office.Interop.Excel.ProtectedViewWindow Pvw, Microsoft.Office.Interop.Excel.XlProtectedViewCloseReason Reason, ref bool Cancel)
        {
            DisplayInWatchWindow(ProtectedViewWindowBeforeClose++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ProtectedViewWindowActivate;
        void _ExcelApplication_ProtectedViewWindowActivate(Microsoft.Office.Interop.Excel.ProtectedViewWindow Pvw)
        {
            DisplayInWatchWindow(ProtectedViewWindowActivate++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short AfterCalculate;
        void _ExcelApplication_AfterCalculate()
        {
            DisplayInWatchWindow(AfterCalculate++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        private void DisplayInWatchWindow(short i, string outputLine)
        {
            if(Common.DisplayEvents)
            {
                AddinHelper.Common.WriteToWatchWindow(string.Format("{0}:{1}", outputLine, i));
            }
        }
    }
}
