using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SMO = Microsoft.SqlServer.Management.Smo;

namespace SMOHelper
{
    class Configuration
    {
        /// <summary>
        /// Initializes a new instance of the Configuration class.
        /// </summary>
        public Configuration(SMO.Configuration configuration)
        {
            AddinHelper.Common.WriteToDebugWindow(string.Format("SMOH.{0}({1})", "Configuration", configuration));
            // Save the real thing in case we need to get something from it.

            _configuration = configuration;                
        }

        public SMO.Configuration _configuration;
    }
}
