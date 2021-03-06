﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;

using PacificLife.Life;

namespace Compliance_Office2010Addin_Word
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            PLLog.Trace("Enter", Common.PROJECT_NAME);

            Globals.Ribbons.Ribbon.chkDisplayEvents.Checked = Common.DisplayEvents;
            Globals.Ribbons.Ribbon.chkEnableAppEvents.Checked = Common.HasAppEvents;

            try
            {
                if(Common.HasAppEvents)
                {
                    Common.AppEvents = new Events.WordAppEvents();
                    Common.AppEvents.WordApplication = Globals.ThisAddIn.Application;
                }

                Common.WordHelper.WordApplication = Globals.ThisAddIn.Application;
            }
            catch(Exception ex)
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
                if(Common.HasAppEvents)
                {
                    Common.AppEvents = null;
                }
            }
            catch(Exception ex)
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
