using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Office2010Addin_Template_Word.Events
{
    class WordAppEvents
    {
        private Microsoft.Office.Interop.Word.Application _WordApplication;
        public Microsoft.Office.Interop.Word.Application WordApplication
        {
            get
            {
                return _WordApplication;
            }
            set
            {
                if (_WordApplication != null)
                {
                	// Should remove all the event handlers;
                }

                _WordApplication = value;

                if (_WordApplication != null)
                {
                    _WordApplication.DocumentBeforeClose += new Microsoft.Office.Interop.Word.ApplicationEvents4_DocumentBeforeCloseEventHandler(_WordApplication_DocumentBeforeClose);
                    _WordApplication.DocumentBeforePrint += new Microsoft.Office.Interop.Word.ApplicationEvents4_DocumentBeforePrintEventHandler(_WordApplication_DocumentBeforePrint);
                    _WordApplication.DocumentBeforeSave += new Microsoft.Office.Interop.Word.ApplicationEvents4_DocumentBeforeSaveEventHandler(_WordApplication_DocumentBeforeSave);
                    _WordApplication.DocumentChange += new Microsoft.Office.Interop.Word.ApplicationEvents4_DocumentChangeEventHandler(_WordApplication_DocumentChange);
                    _WordApplication.DocumentOpen += new Microsoft.Office.Interop.Word.ApplicationEvents4_DocumentOpenEventHandler(_WordApplication_DocumentOpen);
                    _WordApplication.DocumentSync += new Microsoft.Office.Interop.Word.ApplicationEvents4_DocumentSyncEventHandler(_WordApplication_DocumentSync);
                    _WordApplication.EPostageInsert += new Microsoft.Office.Interop.Word.ApplicationEvents4_EPostageInsertEventHandler(_WordApplication_EPostageInsert);
                    _WordApplication.EPostageInsertEx += new Microsoft.Office.Interop.Word.ApplicationEvents4_EPostageInsertExEventHandler(_WordApplication_EPostageInsertEx);
                    _WordApplication.EPostagePropertyDialog += new Microsoft.Office.Interop.Word.ApplicationEvents4_EPostagePropertyDialogEventHandler(_WordApplication_EPostagePropertyDialog);
                    _WordApplication.MailMergeAfterMerge += new Microsoft.Office.Interop.Word.ApplicationEvents4_MailMergeAfterMergeEventHandler(_WordApplication_MailMergeAfterMerge);
                    _WordApplication.MailMergeAfterRecordMerge += new Microsoft.Office.Interop.Word.ApplicationEvents4_MailMergeAfterRecordMergeEventHandler(_WordApplication_MailMergeAfterRecordMerge);
                    _WordApplication.MailMergeBeforeMerge += new Microsoft.Office.Interop.Word.ApplicationEvents4_MailMergeBeforeMergeEventHandler(_WordApplication_MailMergeBeforeMerge);
                    _WordApplication.MailMergeBeforeRecordMerge += new Microsoft.Office.Interop.Word.ApplicationEvents4_MailMergeBeforeRecordMergeEventHandler(_WordApplication_MailMergeBeforeRecordMerge);
                    _WordApplication.MailMergeDataSourceLoad += new Microsoft.Office.Interop.Word.ApplicationEvents4_MailMergeDataSourceLoadEventHandler(_WordApplication_MailMergeDataSourceLoad);
                    _WordApplication.MailMergeDataSourceValidate += new Microsoft.Office.Interop.Word.ApplicationEvents4_MailMergeDataSourceValidateEventHandler(_WordApplication_MailMergeDataSourceValidate);
                    _WordApplication.MailMergeDataSourceValidate2 += new Microsoft.Office.Interop.Word.ApplicationEvents4_MailMergeDataSourceValidate2EventHandler(_WordApplication_MailMergeDataSourceValidate2);
                    _WordApplication.MailMergeWizardSendToCustom += new Microsoft.Office.Interop.Word.ApplicationEvents4_MailMergeWizardSendToCustomEventHandler(_WordApplication_MailMergeWizardSendToCustom);
                    _WordApplication.MailMergeWizardStateChange += new Microsoft.Office.Interop.Word.ApplicationEvents4_MailMergeWizardStateChangeEventHandler(_WordApplication_MailMergeWizardStateChange);
                    _WordApplication.ProtectedViewWindowActivate += new Microsoft.Office.Interop.Word.ApplicationEvents4_ProtectedViewWindowActivateEventHandler(_WordApplication_ProtectedViewWindowActivate);
                    _WordApplication.ProtectedViewWindowBeforeClose += new Microsoft.Office.Interop.Word.ApplicationEvents4_ProtectedViewWindowBeforeCloseEventHandler(_WordApplication_ProtectedViewWindowBeforeClose);
                    _WordApplication.ProtectedViewWindowBeforeEdit += new Microsoft.Office.Interop.Word.ApplicationEvents4_ProtectedViewWindowBeforeEditEventHandler(_WordApplication_ProtectedViewWindowBeforeEdit);
                    _WordApplication.ProtectedViewWindowDeactivate += new Microsoft.Office.Interop.Word.ApplicationEvents4_ProtectedViewWindowDeactivateEventHandler(_WordApplication_ProtectedViewWindowDeactivate);
                    _WordApplication.ProtectedViewWindowOpen += new Microsoft.Office.Interop.Word.ApplicationEvents4_ProtectedViewWindowOpenEventHandler(_WordApplication_ProtectedViewWindowOpen);
                    _WordApplication.ProtectedViewWindowSize += new Microsoft.Office.Interop.Word.ApplicationEvents4_ProtectedViewWindowSizeEventHandler(_WordApplication_ProtectedViewWindowSize);
                    _WordApplication.Startup += new Microsoft.Office.Interop.Word.ApplicationEvents4_StartupEventHandler(_WordApplication_Startup);
                    _WordApplication.WindowActivate += new Microsoft.Office.Interop.Word.ApplicationEvents4_WindowActivateEventHandler(_WordApplication_WindowActivate);
                    _WordApplication.WindowBeforeDoubleClick += new Microsoft.Office.Interop.Word.ApplicationEvents4_WindowBeforeDoubleClickEventHandler(_WordApplication_WindowBeforeDoubleClick);
                    _WordApplication.WindowBeforeRightClick += new Microsoft.Office.Interop.Word.ApplicationEvents4_WindowBeforeRightClickEventHandler(_WordApplication_WindowBeforeRightClick);
                    _WordApplication.WindowDeactivate += new Microsoft.Office.Interop.Word.ApplicationEvents4_WindowDeactivateEventHandler(_WordApplication_WindowDeactivate);
                    _WordApplication.WindowSelectionChange += new Microsoft.Office.Interop.Word.ApplicationEvents4_WindowSelectionChangeEventHandler(_WordApplication_WindowSelectionChange);
                    _WordApplication.WindowSize += new Microsoft.Office.Interop.Word.ApplicationEvents4_WindowSizeEventHandler(_WordApplication_WindowSize);
                    _WordApplication.XMLSelectionChange += new Microsoft.Office.Interop.Word.ApplicationEvents4_XMLSelectionChangeEventHandler(_WordApplication_XMLSelectionChange);
                    _WordApplication.XMLValidationError += new Microsoft.Office.Interop.Word.ApplicationEvents4_XMLValidationErrorEventHandler(_WordApplication_XMLValidationError);
                }
            }
        }

        short countXMLValidationError;
        void _WordApplication_XMLValidationError(Microsoft.Office.Interop.Word.XMLNode XMLNode)
        {
            DisplayInWatchWindow(countXMLValidationError++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countXMLSelectionChange;
        void _WordApplication_XMLSelectionChange(Microsoft.Office.Interop.Word.Selection Sel, Microsoft.Office.Interop.Word.XMLNode OldXMLNode, Microsoft.Office.Interop.Word.XMLNode NewXMLNode, ref int Reason)
        {
            DisplayInWatchWindow(countXMLSelectionChange++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countWindowSize;
        void _WordApplication_WindowSize(Microsoft.Office.Interop.Word.Document Doc, Microsoft.Office.Interop.Word.Window Wn)
        {
            DisplayInWatchWindow(countWindowSize++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countWindowSelectionChange;
        void _WordApplication_WindowSelectionChange(Microsoft.Office.Interop.Word.Selection Sel)
        {
            DisplayInWatchWindow(countWindowSelectionChange++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countWindowDeactivate;
        void _WordApplication_WindowDeactivate(Microsoft.Office.Interop.Word.Document Doc, Microsoft.Office.Interop.Word.Window Wn)
        {
            DisplayInWatchWindow(countWindowDeactivate++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countWindowBeforeRightClick;
        void _WordApplication_WindowBeforeRightClick(Microsoft.Office.Interop.Word.Selection Sel, ref bool Cancel)
        {
            DisplayInWatchWindow(countWindowBeforeRightClick++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short chountWindowBeforeDoubleClick;
        void _WordApplication_WindowBeforeDoubleClick(Microsoft.Office.Interop.Word.Selection Sel, ref bool Cancel)
        {
            DisplayInWatchWindow(chountWindowBeforeDoubleClick++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countWindowActivate;
        void _WordApplication_WindowActivate(Microsoft.Office.Interop.Word.Document Doc, Microsoft.Office.Interop.Word.Window Wn)
        {
            DisplayInWatchWindow(countWindowActivate++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countStartup;
        void _WordApplication_Startup()
        {
            DisplayInWatchWindow(countStartup++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countProtectedViewWindowSize;
        void _WordApplication_ProtectedViewWindowSize(Microsoft.Office.Interop.Word.ProtectedViewWindow PvWindow)
        {
            DisplayInWatchWindow(countProtectedViewWindowSize++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }
        short countProtectedViewWindowOpen;
        void _WordApplication_ProtectedViewWindowOpen(Microsoft.Office.Interop.Word.ProtectedViewWindow PvWindow)
        {
            DisplayInWatchWindow(countProtectedViewWindowOpen++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countProtectedViewWindowDeactivate;
        void _WordApplication_ProtectedViewWindowDeactivate(Microsoft.Office.Interop.Word.ProtectedViewWindow PvWindow)
        {
            DisplayInWatchWindow(countProtectedViewWindowDeactivate++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countProtectedViewWindowBeforeEdit;
        void _WordApplication_ProtectedViewWindowBeforeEdit(Microsoft.Office.Interop.Word.ProtectedViewWindow PvWindow, ref bool Cancel)
        {
            DisplayInWatchWindow(countProtectedViewWindowBeforeEdit++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countProtectedViewWindowBeforeClose;
        void _WordApplication_ProtectedViewWindowBeforeClose(Microsoft.Office.Interop.Word.ProtectedViewWindow PvWindow, int CloseReason, ref bool Cancel)
        {
            DisplayInWatchWindow(countProtectedViewWindowBeforeClose++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countProtectedViewWindowActivate;
        void _WordApplication_ProtectedViewWindowActivate(Microsoft.Office.Interop.Word.ProtectedViewWindow PvWindow)
        {
            DisplayInWatchWindow(countProtectedViewWindowActivate++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countMailMergeWizardStateChange;
        void _WordApplication_MailMergeWizardStateChange(Microsoft.Office.Interop.Word.Document Doc, ref int FromState, ref int ToState, ref bool Handled)
        {
            DisplayInWatchWindow(countMailMergeWizardStateChange++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countMailMergeWizardSendToCustom;
        void _WordApplication_MailMergeWizardSendToCustom(Microsoft.Office.Interop.Word.Document Doc)
        {
            DisplayInWatchWindow(countMailMergeWizardSendToCustom++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countMailMergeDataSourceValidate2;
        void _WordApplication_MailMergeDataSourceValidate2(Microsoft.Office.Interop.Word.Document Doc, ref bool Handled)
        {
            DisplayInWatchWindow(countMailMergeDataSourceValidate2++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countMailMergeDataSourceValidate;
        void _WordApplication_MailMergeDataSourceValidate(Microsoft.Office.Interop.Word.Document Doc, ref bool Handled)
        {
            DisplayInWatchWindow(countMailMergeDataSourceValidate++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countMailMergeDataSourceLoad;
        void _WordApplication_MailMergeDataSourceLoad(Microsoft.Office.Interop.Word.Document Doc)
        {
            DisplayInWatchWindow(countMailMergeDataSourceLoad++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countMailMergeBeforeRecordMerge;
        void _WordApplication_MailMergeBeforeRecordMerge(Microsoft.Office.Interop.Word.Document Doc, ref bool Cancel)
        {
            DisplayInWatchWindow(countMailMergeBeforeRecordMerge++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countMailMergeBeforeMerg;
        void _WordApplication_MailMergeBeforeMerge(Microsoft.Office.Interop.Word.Document Doc, int StartRecord, int EndRecord, ref bool Cancel)
        {
            DisplayInWatchWindow(countMailMergeBeforeMerg++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countMailMergeAfterRecordMerge;
        void _WordApplication_MailMergeAfterRecordMerge(Microsoft.Office.Interop.Word.Document Doc)
        {
            DisplayInWatchWindow(countMailMergeAfterRecordMerge++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countMailMergeAfterMerge;
        void _WordApplication_MailMergeAfterMerge(Microsoft.Office.Interop.Word.Document Doc, Microsoft.Office.Interop.Word.Document DocResult)
        {
            DisplayInWatchWindow(countMailMergeAfterMerge++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countEPostagePropertyDialog;
        void _WordApplication_EPostagePropertyDialog(Microsoft.Office.Interop.Word.Document Doc)
        {
            DisplayInWatchWindow(countEPostagePropertyDialog++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countEPostageInsertEx;
        void _WordApplication_EPostageInsertEx(Microsoft.Office.Interop.Word.Document Doc, int cpDeliveryAddrStart, int cpDeliveryAddrEnd, int cpReturnAddrStart, int cpReturnAddrEnd, int xaWidth, int yaHeight, string bstrPrinterName, string bstrPaperFeed, bool fPrint, ref bool fCancel)
        {
            DisplayInWatchWindow(countEPostageInsertEx++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countEPostageInsert;
        void _WordApplication_EPostageInsert(Microsoft.Office.Interop.Word.Document Doc)
        {
            DisplayInWatchWindow(countEPostageInsert++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countDocumentSync;
        void _WordApplication_DocumentSync(Microsoft.Office.Interop.Word.Document Doc, Microsoft.Office.Core.MsoSyncEventType SyncEventType)
        {
            DisplayInWatchWindow(countDocumentSync++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countDocumentOpen;
        void _WordApplication_DocumentOpen(Microsoft.Office.Interop.Word.Document Doc)
        {
            DisplayInWatchWindow(countDocumentOpen++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countDocumentChange;
        void _WordApplication_DocumentChange()
        {
            DisplayInWatchWindow(countDocumentChange++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countDocumentBeforeSave;
        void _WordApplication_DocumentBeforeSave(Microsoft.Office.Interop.Word.Document Doc, ref bool SaveAsUI, ref bool Cancel)
        {
            DisplayInWatchWindow(countDocumentBeforeSave++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countDocumentBeforePrint;
        void _WordApplication_DocumentBeforePrint(Microsoft.Office.Interop.Word.Document Doc, ref bool Cancel)
        {
            DisplayInWatchWindow(countDocumentBeforePrint++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countDocumentBeforeClose;
        void _WordApplication_DocumentBeforeClose(Microsoft.Office.Interop.Word.Document Doc, ref bool Cancel)
        {
            DisplayInWatchWindow(countDocumentBeforeClose++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
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
