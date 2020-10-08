using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Visio;
using Office = Microsoft.Office.Core;

namespace AddinHelper
{
    public class Visio
    {
        private Microsoft.Office.Interop.Visio.Application _VisioApplication;
        public Microsoft.Office.Interop.Visio.Application VisioApplication
        {
	        get { return _VisioApplication; }
	        set { _VisioApplication = value; }
        }


        private bool _enableScreenUpdatesToggle = true;
        public bool EnableScreenUpdatesToggle
        {
	        get { return _enableScreenUpdatesToggle; }
	        set { _enableScreenUpdatesToggle = value; }
        }

    }
}
