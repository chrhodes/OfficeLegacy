using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SMO = Microsoft.SqlServer.Management.Smo;

namespace SMOHelper
{
    public class LogFile
    {
        /// <summary>
        /// Initializes a new instance of the LogFile class.
        /// </summary>
        public LogFile(SMO.LogFile logFile)
        {
            AddinHelper.Common.WriteToDebugWindow(string.Format("SMOH.{0}({1})", "LogFile", logFile));
            // Save the real table in case we need to get something from it.

            _logFile = logFile;


            // NB. Not all properties are available depending on login credentials
            // Wrap in TC blocks.

            try { Name = logFile.Name; }                           catch (Exception) { Name = "<No Access>"; }
        }

        public SMO.LogFile _logFile;

        public string Name;
    }
}
