using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SMO = Microsoft.SqlServer.Management.Smo;

namespace SMOHelper
{
    public class ServerAdapter
    {
        /// <summary>
        /// Initializes a new instance of the ServerAdapter class.
        /// </summary>
        public ServerAdapter(SMO.ServerAdapter serverAdapter)
        {
            AddinHelper.Common.WriteToDebugWindow(string.Format("SMOH.{0}({1})", "ServerAdapter", serverAdapter));
            // Save the real thing in case we need to get something from it.

            _serverAdapter = serverAdapter;  
            
            // NB. Not all properties are available depending on login credentials
            // Wrap in TC blocks.

            try { InstanceName = serverAdapter.InstanceName; }                           catch (Exception) { InstanceName = "<No Access>"; }
        }

        public SMO.ServerAdapter _serverAdapter;

        public string InstanceName;

    }
}
