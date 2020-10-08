using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SMO = Microsoft.SqlServer.Management.Smo;

namespace SMOHelper
{
    public class Endpoint
    {
        /// <summary>
        /// Initializes a new instance of the Endpoint class.
        /// </summary>
        public Endpoint(SMO.Endpoint endPoint)
        {
            AddinHelper.Common.WriteToDebugWindow(string.Format("SMOH.{0}({1})", "EndPoint", endPoint));

            // Save the real thing in case we need to get something from it.

            _endPoint = endPoint; 

            // NB. Not all properties are available depending on login credentials
            // Wrap in TC blocks.

            try { EndpointState = endPoint.EndpointState.ToString(); }      catch (Exception) { EndpointState = "<No Access>"; }
            try { EndpointType = endPoint.EndpointType.ToString(); }        catch (Exception) { EndpointType = "<No Access>"; }
            try { ID = endPoint.ID.ToString(); }                            catch (Exception) { ID = "<No Access>"; }
            try { IsAdminEndpoint = endPoint.IsAdminEndpoint.ToString(); }  catch (Exception) { IsAdminEndpoint = "<No Access>"; }
            try { IsSystemObject = endPoint.IsSystemObject.ToString(); }    catch (Exception) { IsSystemObject = "<No Access>"; }
            try { Name = endPoint.Name; }                                   catch (Exception) { Name = "<No Access>"; }
            try { Owner = endPoint.Owner;}                                  catch (Exception) { Owner = "<No Access>"; }
            try { Protocol = endPoint.Protocol.ToString(); }                catch (Exception) { Protocol = "<No Access>"; }
            try { ProtocolType = endPoint.ProtocolType.ToString(); }        catch (Exception) { ProtocolType = "<No Access>"; }
            try { Urn = endPoint.Urn;}                                      catch (Exception) { Urn = "<No Access>"; }
        }

        SMO.Endpoint _endPoint;

        public string EndpointState;
        public string EndpointType;
        public string ID;
        public string IsAdminEndpoint;
        public string IsSystemObject;
        public string Name;
        public string Owner;
        public string Protocol;
        public string ProtocolType;
        public string Urn;
    }
}
