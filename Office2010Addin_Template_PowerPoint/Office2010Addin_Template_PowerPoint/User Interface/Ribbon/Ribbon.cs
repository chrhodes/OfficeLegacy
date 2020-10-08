using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace Office2010Addin_Template_PowerPoint
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            
        }

        #region Event Handlers

        private void btnAddInInfo_Click(object sender, RibbonControlEventArgs e)
        {
            DisplayAddInInfo();
        }

        private void btnDebugWindow_Click(object sender, RibbonControlEventArgs e)
        {
            DisplayDebugWindow();
        }

        private void btnDeveloperMode_Click(object sender, RibbonControlEventArgs e)
        {
            ToggleDeveloperMode();
        }

        private void btnAppUtilities_Click(object sender, RibbonControlEventArgs e)
        {
            if (Common.TaskPaneAppUtil == null)
            {
            	Common.TaskPaneAppUtil = AddinHelper.TaskPaneUtil.AddTaskPane(new User_Interface.Task_Panes.TaskPane_AppUtil(), "App Utilities", Globals.ThisAddIn.CustomTaskPanes);
            }
            else
            {
                Common.TaskPaneAppUtil.Visible = !Common.TaskPaneAppUtil.Visible;
            }
        }

        private void btnWatchWindow_Click(object sender, RibbonControlEventArgs e)
        {
            DisplayWatchWindow();
        }

        private void chkDisplayEvents_Click(object sender, RibbonControlEventArgs e)
        {
            Common.DisplayEvents = chkDisplayEvents.Checked;
        }

        private void chkEnableAppEvents_Click(object sender, RibbonControlEventArgs e)
        {
            Common.HasAppEvents = chkEnableAppEvents.Checked;

            if (Common.HasAppEvents)
            {
            	if (Common.AppEvents == null)
                {
                	Common.AppEvents = new Events.PowerPointAppEvents();
                    Common.AppEvents.PowerPointApplication = Globals.ThisAddIn.Application;
                }
            }
            else
            {
                Common.AppEvents = null;
            }
        }

        #endregion

        #region Main Function Routines


        private void DisplayAddInInfo()
        {
            AddinHelper.AddInInfo.DisplayInfo();
        }

        private void DisplayDebugWindow()
        {
            if (AddinHelper.Common.DebugWindow.Visible)
            {
            	AddinHelper.Common.DebugWindow.Visible = false;
            }
            else
            {
                AddinHelper.Common.DebugWindow.Visible = true;
            }
        }

        private void DisplayWatchWindow()
        {
            AddinHelper.Common.WatchWindow.Visible = !AddinHelper.Common.WatchWindow.Visible;
        }

        private void ToggleDeveloperMode()
        {
            AddinHelper.Common.DeveloperMode = !AddinHelper.Common.DeveloperMode;
            Globals.Ribbons.Ribbon.grpDebug.Visible = AddinHelper.Common.DeveloperMode;
        }

#endregion



    }
}
