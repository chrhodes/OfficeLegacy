using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Compliance_Office2010Addin_Word
{
    class Common : AddinHelper.Common
    {
        new public const string PROJECT_NAME = "Compliance_Office2010Addin_Word";

        public static AddinHelper.Word WordHelper = new AddinHelper.Word();
        public static Events.WordAppEvents AppEvents;

        //public static Microsoft.Office.Tools.CustomTaskPane TaskPaneConfig;

        //public static Microsoft.Office.Tools.CustomTaskPane TaskPaneHelp;

        public static Microsoft.Office.Tools.CustomTaskPane TaskPaneComplianceUtil;

        public static Microsoft.Office.Tools.CustomTaskPane TaskPaneAppUtil;
    }
}
