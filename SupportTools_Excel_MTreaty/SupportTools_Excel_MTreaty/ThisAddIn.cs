using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

using PacificLife.Life;
using ExcelHlp = AddinHelper.Excel;

namespace SupportTools_Excel_MTreaty
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //PLLog.Trace("Enter", Common.PROJECT_NAME);
            Common.DeveloperMode = true;
            Common.WriteToDebugWindow("ThisAddIn_Startup()");
            Common.DeveloperMode = false;

            Globals.Ribbons.Ribbon.chkDisplayEvents.Checked = Common.DisplayEvents;
            Globals.Ribbons.Ribbon.chkEnableAppEvents.Checked = Common.HasAppEvents;

            try
            {
                if (Common.HasAppEvents)
                {
                    Common.AppEvents = new Events.ExcelAppEvents();
                    Common.AppEvents.ExcelApplication = Globals.ThisAddIn.Application;
                }

                ExcelHlp.ExcelApplication = Globals.ThisAddIn.Application;
            }
            catch (Exception ex)
            {
                //PLLog.Error(ex, Common.PROJECT_NAME);
                Common.DeveloperMode = true;
                Common.WriteToDebugWindow(ex.ToString());
                Common.DeveloperMode = false;
            }

            //PLLog.Trace("Exit", Common.PROJECT_NAME);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            //PLLog.Trace("Enter", Common.PROJECT_NAME);
            Common.DeveloperMode = true;
            Common.WriteToDebugWindow("ThisAddIn_Shutdown()");
            Common.DeveloperMode = false;

            try
            {
                if (Common.HasAppEvents)
                {
                    Common.AppEvents = null;
                }
            }
            catch (Exception ex)
            {
                //PLLog.Error(ex, Common.PROJECT_NAME);
                Common.DeveloperMode = true;
                Common.WriteToDebugWindow(ex.ToString());
                Common.DeveloperMode = false;
            }

            //PLLog.Trace("Exit", Common.PROJECT_NAME);
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += ThisAddIn_Startup;
            this.Shutdown += ThisAddIn_Shutdown;
        }

        #endregion
    }
}
