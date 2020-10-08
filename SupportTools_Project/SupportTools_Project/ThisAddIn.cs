using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using MSProject = Microsoft.Office.Interop.MSProject;
using Office = Microsoft.Office.Core;

using PacificLife.Life;

namespace SupportTools_Project
{
    public partial class ThisAddIn
    {
        // Need to do a bit more work to use CustomTask Panes in Project.  (Handled by Designer normally)

        internal Microsoft.Office.Tools.CustomTaskPaneCollection CustomTaskPanes;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            PLLog.Trace("Enter", Common.PROJECT_NAME);

            Globals.Ribbons.Ribbon.chkDisplayEvents.Checked = Common.DisplayEvents;
            Globals.Ribbons.Ribbon.chkEnableAppEvents.Checked = Common.HasAppEvents;

            try
            {
                if (Common.HasAppEvents)
                {
                    Common.AppEvents = new Events.ProjectAppEvents();
                    Common.AppEvents.ProjectApplication = Globals.ThisAddIn.Application;
                }

                Common.ProjectHelper.ProjectApplication = Globals.ThisAddIn.Application;

                // Need to do a bit more work to use CustomTask Panes in Project.  (Handled by Designer normally)

                CustomTaskPanes = Globals.Factory.CreateCustomTaskPaneCollection(null, null, "CustomTaskPanes", "CustomTaskPanes", this);
            }
            catch (Exception ex)
            {
                PLLog.Error(ex, Common.PROJECT_NAME);
            }

            PLLog.Trace("Exit", Common.PROJECT_NAME);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            PLLog.Trace("Enter", Common.PROJECT_NAME);

            try
            {
                if (Common.HasAppEvents)
                {
                    Common.AppEvents = null;
                }

                // Need to do a bit more work to use CustomTask Panes in Project.  (Handled by Designer normally)
                CustomTaskPanes.Dispose();
            }
            catch (Exception ex)
            {
                PLLog.Error(ex, Common.PROJECT_NAME);
            }

            PLLog.Trace("Exit", Common.PROJECT_NAME);
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
