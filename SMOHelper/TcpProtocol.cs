using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SMO = Microsoft.SqlServer.Management.Smo;

namespace SMOHelper
{
    public class TcpProtocol
    {
        /// <summary>
        /// Initializes a new instance of the TcpProtocol class.
        /// </summary>
        public TcpProtocol(SMO.TcpProtocol tcpProtocol)
        {
            AddinHelper.Common.WriteToDebugWindow(string.Format("SMOH.{0}({1})", "TcpProtocol", tcpProtocol));
            // Save the real thing in case we need to get something from it.

            _tcpProtocol = tcpProtocol;      
        
            // NB. Not all properties are available depending on login credentials
            // Wrap in TC blocks.

            try { IsDynamicPort = tcpProtocol.IsDynamicPort.ToString(); }               catch (Exception) { IsDynamicPort = "<No Access>"; }
            try { IsSystemObject = tcpProtocol.IsSystemObject.ToString(); }             catch (Exception) { IsSystemObject = "<No Access>"; }
            try { ListenerIPAddress = tcpProtocol.ListenerIPAddress.ToString(); }       catch (Exception) { ListenerIPAddress = "<No Access>"; }
            try { ListenerPort = tcpProtocol.ListenerPort.ToString(); }                 catch (Exception) { ListenerPort = "<No Access>"; }
        }

        public SMO.TcpProtocol _tcpProtocol;

        string IsDynamicPort;
        string IsSystemObject;
        string ListenerIPAddress;
        string ListenerPort;
    }
}
