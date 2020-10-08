using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SMO = Microsoft.SqlServer.Management.Smo;

namespace SMOHelper
{
    public class PhysicalPartition
    {
        /// <summary>
        /// Initializes a new instance of the PhysicalPartition class.
        /// </summary>
        public PhysicalPartition(SMO.PhysicalPartition physicalPartition)
        {
            AddinHelper.Common.WriteToDebugWindow(string.Format("SMOH.{0}({1})", "ServerAdapter", physicalPartition));
            // Save the real thing in case we need to get something from it.

            _physicalPartition = physicalPartition;  
            
            // NB. Not all properties are available depending on login credentials
            // Wrap in TC blocks.

            try { FileGroupName = physicalPartition.FileGroupName; }           catch (Exception) { FileGroupName = "<No Access>"; }
        }

        public SMO.PhysicalPartition _physicalPartition;

        public string FileGroupName;
    }
}
