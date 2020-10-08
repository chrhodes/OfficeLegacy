using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.MSProject;
using Office = Microsoft.Office.Core;

namespace AddinHelper
{
    public class Project
    {
        private Microsoft.Office.Interop.MSProject.Application _ProjectApplication;
        public Microsoft.Office.Interop.MSProject.Application ProjectApplication
        {
	        get { return _ProjectApplication; }
	        set { _ProjectApplication = value; }
        }


        private bool _enableScreenUpdatesToggle = true;
        public bool EnableScreenUpdatesToggle
        {
	        get { return _enableScreenUpdatesToggle; }
	        set { _enableScreenUpdatesToggle = value; }
        }

    }
}
