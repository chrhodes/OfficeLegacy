using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SupportTools_Outlook.Events
{
    class OutlookAppEvents
    {
        private Microsoft.Office.Interop.Outlook.Application _OutlookApplication;
        public Microsoft.Office.Interop.Outlook.Application OutlookApplication
        {
            get
            {
                return _OutlookApplication;
            }
            set
            {
                if (_OutlookApplication != null)
                {
                    // Should remove all the event handlers;
                }

                _OutlookApplication = value;

                if (_OutlookApplication != null)
                {
                    _OutlookApplication.AdvancedSearchComplete += new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_AdvancedSearchCompleteEventHandler(_OutlookApplication_AdvancedSearchComplete);
                    _OutlookApplication.AdvancedSearchStopped += new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_AdvancedSearchStoppedEventHandler(_OutlookApplication_AdvancedSearchStopped);
                    _OutlookApplication.AttachmentContextMenuDisplay += new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_AttachmentContextMenuDisplayEventHandler(_OutlookApplication_AttachmentContextMenuDisplay);
                    _OutlookApplication.BeforeFolderSharingDialog += new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_BeforeFolderSharingDialogEventHandler(_OutlookApplication_BeforeFolderSharingDialog);
                    _OutlookApplication.ContextMenuClose += new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_ContextMenuCloseEventHandler(_OutlookApplication_ContextMenuClose);
                    _OutlookApplication.FolderContextMenuDisplay += new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_FolderContextMenuDisplayEventHandler(_OutlookApplication_FolderContextMenuDisplay);
                    _OutlookApplication.ItemContextMenuDisplay += new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_ItemContextMenuDisplayEventHandler(_OutlookApplication_ItemContextMenuDisplay);
                    _OutlookApplication.ItemLoad += new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_ItemLoadEventHandler(_OutlookApplication_ItemLoad);
                    _OutlookApplication.ItemSend += new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_ItemSendEventHandler(_OutlookApplication_ItemSend);
                    _OutlookApplication.MAPILogonComplete += new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_MAPILogonCompleteEventHandler(_OutlookApplication_MAPILogonComplete);
                    _OutlookApplication.NewMail += new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_NewMailEventHandler(_OutlookApplication_NewMail);
                    _OutlookApplication.NewMailEx += new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_NewMailExEventHandler(_OutlookApplication_NewMailEx);
                    _OutlookApplication.OptionsPagesAdd += new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_OptionsPagesAddEventHandler(_OutlookApplication_OptionsPagesAdd);
                    _OutlookApplication.Reminder += new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_ReminderEventHandler(_OutlookApplication_Reminder);
                    _OutlookApplication.ShortcutContextMenuDisplay += new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_ShortcutContextMenuDisplayEventHandler(_OutlookApplication_ShortcutContextMenuDisplay);
                    _OutlookApplication.Startup += new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_StartupEventHandler(_OutlookApplication_Startup);
                    _OutlookApplication.StoreContextMenuDisplay += new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_StoreContextMenuDisplayEventHandler(_OutlookApplication_StoreContextMenuDisplay);
                    _OutlookApplication.ViewContextMenuDisplay += new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_ViewContextMenuDisplayEventHandler(_OutlookApplication_ViewContextMenuDisplay);
                }
            }
        }

        short Reminder;
        void _OutlookApplication_Reminder(object Item)
        {
            DisplayInWatchWindow(Reminder++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ViewContextMenuDisplay;
        void _OutlookApplication_ViewContextMenuDisplay(Microsoft.Office.Core.CommandBar CommandBar, Microsoft.Office.Interop.Outlook.View View)
        {
            DisplayInWatchWindow(ViewContextMenuDisplay++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short StoreContextMenuDisplay;
        void _OutlookApplication_StoreContextMenuDisplay(Microsoft.Office.Core.CommandBar CommandBar, Microsoft.Office.Interop.Outlook.Store Store)
        {
            DisplayInWatchWindow(StoreContextMenuDisplay++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short Startup;
        void _OutlookApplication_Startup()
        {
            DisplayInWatchWindow(Startup++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ShortcutContextMenuDisplay;
        void _OutlookApplication_ShortcutContextMenuDisplay(Microsoft.Office.Core.CommandBar CommandBar, Microsoft.Office.Interop.Outlook.OutlookBarShortcut Shortcut)
        {
            DisplayInWatchWindow(ShortcutContextMenuDisplay++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short OptionsPagesAdd;
        void _OutlookApplication_OptionsPagesAdd(Microsoft.Office.Interop.Outlook.PropertyPages Pages)
        {
            DisplayInWatchWindow(OptionsPagesAdd++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short NewMailEx;
        void _OutlookApplication_NewMailEx(string EntryIDCollection)
        {
            DisplayInWatchWindow(NewMailEx++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short NewMail;
        void _OutlookApplication_NewMail()
        {
            DisplayInWatchWindow(NewMail++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short MAPILogonComplete;
        void _OutlookApplication_MAPILogonComplete()
        {
            DisplayInWatchWindow(MAPILogonComplete++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ItemSend;
        void _OutlookApplication_ItemSend(object Item, ref bool Cancel)
        {
            DisplayInWatchWindow(ItemSend++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ItemLoad;
        void _OutlookApplication_ItemLoad(object Item)
        {
            DisplayInWatchWindow(ItemLoad++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ItemContextMenuDisplay;
        void _OutlookApplication_ItemContextMenuDisplay(Microsoft.Office.Core.CommandBar CommandBar, Microsoft.Office.Interop.Outlook.Selection Selection)
        {
            DisplayInWatchWindow(ItemContextMenuDisplay++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short FolderContextMenuDisplay;
        void _OutlookApplication_FolderContextMenuDisplay(Microsoft.Office.Core.CommandBar CommandBar, Microsoft.Office.Interop.Outlook.MAPIFolder Folder)
        {
            DisplayInWatchWindow(FolderContextMenuDisplay++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short ContextMenuClose;
        void _OutlookApplication_ContextMenuClose(Microsoft.Office.Interop.Outlook.OlContextMenu ContextMenu)
        {
            DisplayInWatchWindow(ContextMenuClose++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short BeforeFolderSharingDialog;
        void _OutlookApplication_BeforeFolderSharingDialog(Microsoft.Office.Interop.Outlook.MAPIFolder FolderToShare, ref bool Cancel)
        {
            DisplayInWatchWindow(BeforeFolderSharingDialog++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short AttachmentContextMenuDisplay;
        void _OutlookApplication_AttachmentContextMenuDisplay(Microsoft.Office.Core.CommandBar CommandBar, Microsoft.Office.Interop.Outlook.AttachmentSelection Attachments)
        {
            DisplayInWatchWindow(AttachmentContextMenuDisplay++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short AdvancedSearchStopped;
        void _OutlookApplication_AdvancedSearchStopped(Microsoft.Office.Interop.Outlook.Search SearchObject)
        {
            DisplayInWatchWindow(AdvancedSearchStopped++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short AdvancedSearchComplete;
        void _OutlookApplication_AdvancedSearchComplete(Microsoft.Office.Interop.Outlook.Search SearchObject)
        {
            DisplayInWatchWindow(AdvancedSearchComplete++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
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
