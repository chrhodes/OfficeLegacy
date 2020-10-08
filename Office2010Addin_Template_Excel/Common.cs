using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Office2010Addin_Template_Excel
{
    class Common : AddinHelper.Common
    {
        new public const string PROJECT_NAME = "Office2010Addin_Template_Excel";

        public static AddinHelper.Excel ExcelHelper = new AddinHelper.Excel();
        public static Events.ExcelAppEvents AppEvents;

        public const int cMaxFileNameLength = 128;
        private static Data.ApplicationDS _ApplicationDS;
        public static Data.ApplicationDS ApplicationDS
        {
            get
            {
                if (_ApplicationDS == null)
                {
                    _ApplicationDS = new Data.ApplicationDS();
                }
                return _ApplicationDS;
            }
            set
            {
                _ApplicationDS = value;
            }
        }

        //public static Microsoft.Office.Tools.CustomTaskPane TaskPaneConfig;

        //public static Microsoft.Office.Tools.CustomTaskPane TaskPaneHelp;

        public static Microsoft.Office.Tools.CustomTaskPane TaskPaneAppUtil;

    }
}
