using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SupportTools_Outlook
{
    class Common : AddinHelper.Common
    {
        new public const string PROJECT_NAME = "SupportTools_Outlook";

        public static AddinHelper.Outlook OutlookHelper = new AddinHelper.Outlook();
        public static Events.OutlookAppEvents AppEvents;

        //public static Microsoft.Office.Tools.CustomTaskPane TaskPaneConfig;

        //public static Microsoft.Office.Tools.CustomTaskPane TaskPaneHelp;

        public static Microsoft.Office.Tools.CustomTaskPane TaskPaneAppUtil;
    }
}
