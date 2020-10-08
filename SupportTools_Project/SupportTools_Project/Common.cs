using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SupportTools_Project
{
    class Common : AddinHelper.Common
    {
        new public const string PROJECT_NAME = "SupportTools_Project";

        public static AddinHelper.Project ProjectHelper = new AddinHelper.Project();
        public static Events.ProjectAppEvents AppEvents;

        //public static Microsoft.Office.Tools.CustomTaskPane TaskPaneConfig;

        //public static Microsoft.Office.Tools.CustomTaskPane TaskPaneHelp;

        public static Microsoft.Office.Tools.CustomTaskPane TaskPaneAppUtil;
    }
}
