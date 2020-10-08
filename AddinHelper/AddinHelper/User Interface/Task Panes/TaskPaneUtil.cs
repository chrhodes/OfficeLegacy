using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Office.Tools;

namespace AddinHelper
{
    public class TaskPaneUtil
    {
        public static Microsoft.Office.Tools.CustomTaskPane AddTaskPane(System.Windows.Forms.UserControl taskPane, string name, CustomTaskPaneCollection customTaskPanes)
        {
            //PLLog.Trace3("Enter", Common.PROJECT_NAME);

	        CustomTaskPane ctp = default(CustomTaskPane);
	        ctp = customTaskPanes.Add(taskPane, name);
	        ctp.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
	        ctp.Visible = true;

            //PLLog.Trace3("Exit", System.Data.Common.PROJECT_NAME);
	        return ctp;

        }

        public static void RemoveTaskPane(CustomTaskPane taskPane, CustomTaskPaneCollection customTaskPanes)
        {
            //PLLog.Trace3("Enter", System.Data.Common.PROJECT_NAME);

	        customTaskPanes.Remove(taskPane);

            //PLLog.Trace3("Exit", System.Data.Common.PROJECT_NAME);
        }

    }
}
