using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Office2010Addin_Template_Visio
{
    class Common : AddinHelper.Common
    {
        new public const string PROJECT_NAME = "Office2010Addin_Template_Visio";

        public static AddinHelper.Visio VisioHelper = new AddinHelper.Visio();
        public static Events.VisioAppEvents AppEvents;
        public static bool DisplayChattyEvents = false;

        //public static Microsoft.Office.Tools.CustomTaskPane TaskPaneConfig;

        //public static Microsoft.Office.Tools.CustomTaskPane TaskPaneHelp;

        public static Microsoft.Office.Tools.CustomTaskPane TaskPaneAppUtil;
    }
}
