using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace AddinHelper
{
    public class Outlook
    {
        private Microsoft.Office.Interop.Outlook.Application _OutlookApplication;
        public Microsoft.Office.Interop.Outlook.Application OutlookApplication
        {
	        get { return _OutlookApplication; }
	        set { _OutlookApplication = value; }
        }


        private bool _enableScreenUpdatesToggle = true;
        public bool EnableScreenUpdatesToggle
        {
	        get { return _enableScreenUpdatesToggle; }
	        set { _enableScreenUpdatesToggle = value; }
        }

    }
}
