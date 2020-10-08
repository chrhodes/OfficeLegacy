using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Office2010Addin_Template_Outlook
{
    class Common : AddinHelper.Common
    {
        new public const string PROJECT_NAME = "Office2010Addin_Template_Outlook";

        public static AddinHelper.Outlook OutlookHelper = new AddinHelper.Outlook();
        public static Events.OutlookAppEvents AppEvents;

        //public static Microsoft.Office.Tools.CustomTaskPane TaskPaneConfig;

        //public static Microsoft.Office.Tools.CustomTaskPane TaskPaneHelp;

        public static Microsoft.Office.Tools.CustomTaskPane TaskPaneAppUtil;
    }
}
