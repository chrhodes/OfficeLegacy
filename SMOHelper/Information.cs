using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SMO = Microsoft.SqlServer.Management.Smo;

namespace SMOHelper
{
    public class Information
    {
        /// <summary>
        /// Initializes a new instance of the Information class.
        /// </summary>
        public Information(SMO.Information information)
        {
            AddinHelper.Common.WriteToDebugWindow(string.Format("SMOH.{0}({1})", "Information", information));
            // Save the real table in case we need to get something from it.

            _information = information;    
        }

        public SMO.Information _information;
    }
}
