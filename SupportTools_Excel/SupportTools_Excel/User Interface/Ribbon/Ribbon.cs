using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using ExcelHlp = AddinHelper.Excel;

namespace SupportTools_Excel
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
            if(Common.TaskPaneExcelUtil == null)
            {
                Common.TaskPaneExcelUtil = AddinHelper.TaskPaneUtil.AddTaskPane(new User_Interface.Task_Panes.TaskPane_ExcelUtil(), "Excel Utilities", Globals.ThisAddIn.CustomTaskPanes);
                Common.TaskPaneExcelUtil.Width = Common.TaskPaneExcelUtil.Control.Width;
            }
            else
            {
                Common.TaskPaneExcelUtil.Visible = ! Common.TaskPaneExcelUtil.Visible;
            }
        }

        private void btnITRs_Click(object sender, RibbonControlEventArgs e)
        {
            if(Common.TaskPaneITRs == null)
            {
                Common.TaskPaneITRs = AddinHelper.TaskPaneUtil.AddTaskPane(new User_Interface.Task_Panes.TaskPane_ITRs(), "ITRs", Globals.ThisAddIn.CustomTaskPanes);
                Common.TaskPaneITRs.Width = Common.TaskPaneITRs.Control.Width;
            }
            else
            {
                Common.TaskPaneITRs.Visible = ! Common.TaskPaneITRs.Visible;
            }
        }

        private void btnLogParser_Click(object sender, RibbonControlEventArgs e)
        {
            if(Common.TaskPaneLogParser == null)
            {
                Common.TaskPaneLogParser = AddinHelper.TaskPaneUtil.AddTaskPane(new User_Interface.Task_Panes.TaskPane_LogParser(), "Log Parser", Globals.ThisAddIn.CustomTaskPanes);
                Common.TaskPaneLogParser.Width = Common.TaskPaneLogParser.Control.Width;
            }
            else
            {
                Common.TaskPaneLogParser.Visible = ! Common.TaskPaneLogParser.Visible;
            }
        }

        private void btnLTC_Click(object sender, RibbonControlEventArgs e)
        {
            if (Common.TaskPaneLTC == null)
            {
                Common.TaskPaneLTC = AddinHelper.TaskPaneUtil.AddTaskPane(new User_Interface.Task_Panes.TaskPane_LTC(), "LTC", Globals.ThisAddIn.CustomTaskPanes);
                Common.TaskPaneLTC.Width = Common.TaskPaneLTC.Control.Width;
            }
            else
            {
                Common.TaskPaneLTC.Visible = !Common.TaskPaneLTC.Visible;
            }
        }

        //private void btnMTreaty_Click(object sender, RibbonControlEventArgs e)
        //{
        //    if(Common.TaskPaneMTreaty == null)
        //    {
        //        Common.TaskPaneMTreaty = AddinHelper.TaskPaneUtil.AddTaskPane(new User_Interface.Task_Panes.TaskPane_MTreaty(), "MTreaty", Globals.ThisAddIn.CustomTaskPanes);
        //        Common.TaskPaneMTreaty.Width = Common.TaskPaneMTreaty.Control.Width;
        //    }
        //    else
        //    {
        //        Common.TaskPaneMTreaty.Visible = ! Common.TaskPaneMTreaty.Visible;
        //    }
        //}

        private void btnNetworkTraces_Click(object sender, RibbonControlEventArgs e)
        {
            if(Common.TaskPaneNetworkTrace == null)
            {
                Common.TaskPaneNetworkTrace = AddinHelper.TaskPaneUtil.AddTaskPane(new User_Interface.Task_Panes.TaskPane_NetworkTrace(), "Network Traces", Globals.ThisAddIn.CustomTaskPanes);
                Common.TaskPaneNetworkTrace.Width = Common.TaskPaneNetworkTrace.Control.Width;
            }
            else
            {
                Common.TaskPaneNetworkTrace.Visible = ! Common.TaskPaneNetworkTrace.Visible;
            }
        }

        private void btnSQLSMO_Click(object sender, RibbonControlEventArgs e)
        {
            if(Common.TaskPaneSQLSMO == null)
            {
                Common.TaskPaneSQLSMO = AddinHelper.TaskPaneUtil.AddTaskPane(new User_Interface.Task_Panes.TaskPane_SQLSMO(), "SQL SMO", Globals.ThisAddIn.CustomTaskPanes);
                // This throws an exception
                //Globals.ThisAddIn.Application.CommandBars["SQL SMO"].Width = Common.TaskPaneSQLSMO.Width;
                foreach(Microsoft.Office.Core.CommandBar bar in Globals.ThisAddIn.Application.CommandBars)
                {
                    string foo = bar.Name;

                    if (foo == "SQL SMO")
                    {
                        // Which is curious as the bar is found!
                        //Globals.ThisAddIn.Application.CommandBars["SQL SMO"].Width = Common.TaskPaneSQLSMO.Width;
                    }
                }

                // This works if the minimum size for the control has been set.
                Common.TaskPaneSQLSMO.Width = Common.TaskPaneSQLSMO.Control.Width;
            }
            else
            {
                Common.TaskPaneSQLSMO.Visible = !Common.TaskPaneSQLSMO.Visible;
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

            if(Common.HasAppEvents)
            {
                if(Common.AppEvents == null)
                {
                    Common.AppEvents = new Events.ExcelAppEvents();
                    Common.AppEvents.ExcelApplication = Globals.ThisAddIn.Application;
                }
            }
            else
            {
                Common.AppEvents = null;
                Common.AppEvents.ExcelApplication = null;
            }
        }

        private void chkScreenUpdates_Click(object sender, RibbonControlEventArgs e)
        {
            ExcelHlp.DisplayScreenUpdates = chkScreenUpdates.Checked;
        }
        #endregion

        #region Main Function Routines


        private void DisplayAddInInfo()
        {
            AddinHelper.AddInInfo.DisplayInfo();
        }

        private void DisplayDebugWindow()
        {
            if(AddinHelper.Common.DebugWindow.Visible)
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
