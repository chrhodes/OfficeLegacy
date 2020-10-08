using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SMO = Microsoft.SqlServer.Management.Smo;

namespace SMOHelper
{
    public class ServerRole
    {
        #region Initialization
        /// <summary>
        /// Initializes a new instance of the ServerRole class.
        /// </summary>
        public ServerRole(SMO.ServerRole serverRole)
        {
            AddinHelper.Common.WriteToDebugWindow(string.Format("SMOH.{0}({1})", "ServerRole", serverRole));

            // Save the real thing in case we need to get something from it.

            _serverRole = serverRole; 

            // NB. Not all properties are available depending on login credentials
            // Wrap in TC blocks.

            try { Name = serverRole.Name; }                           catch (Exception) { Name = "<No Access>"; }
        }

        #endregion

        SMO.ServerRole _serverRole;

        public string Name;
    }
}
