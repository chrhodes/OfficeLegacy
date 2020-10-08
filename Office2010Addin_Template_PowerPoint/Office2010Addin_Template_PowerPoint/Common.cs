using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Office2010Addin_Template_PowerPoint
{
    class Common : AddinHelper.Common
    {
        new public const string PROJECT_NAME = "Office2010Addin_Template_PowerPoint";

        public static AddinHelper.PowerPoint PowerPointHelper = new AddinHelper.PowerPoint();
        public static Events.PowerPointAppEvents AppEvents;

        //public static Microsoft.Office.Tools.CustomTaskPane TaskPaneConfig;

        //public static Microsoft.Office.Tools.CustomTaskPane TaskPaneHelp;

        public static Microsoft.Office.Tools.CustomTaskPane TaskPaneAppUtil;
    }
}
