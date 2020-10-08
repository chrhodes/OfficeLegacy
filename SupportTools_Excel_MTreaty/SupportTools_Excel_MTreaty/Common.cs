using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SupportTools_Excel_MTreaty
{
    class Common : AddinHelper.Common
    {
        new public const string PROJECT_NAME = "SupportTools_Excel_MTreaty";

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


        #region MTreaty

        public static Microsoft.Office.Tools.CustomTaskPane TaskPaneMTreaty;

        public const string cMTREATY_FOLDER_PROD = @"\\lifenas115\DataServices\Production\M_Treaty_Reporting";
        public const string cMTREATY_FOLDER_STAGING = @"\\lifenas215\DataServices\QA_Staging\M_Treaty_Reporting";
        public const string cMTREATY_FUND_SERVICE_FEES_SHEETNAME = "Stan Tucker";
        public const string cMTREATY_FUND_ADVISORY_FEES_SHEETNAME = "Combined Fees";
        public const string cMTREATY_CASH_MANAGEMENT_FEES_SHEETNAME = "Analysis";
        public const string cMTREATY_VITS_FEES_SHEETNAME = "??";

        #endregion

    }
}
