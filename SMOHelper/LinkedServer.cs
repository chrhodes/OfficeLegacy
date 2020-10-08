using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SMO = Microsoft.SqlServer.Management.Smo;

namespace SMOHelper
{
    public class LinkedServer
    {
        //public string BackupDirectory;
        //public string Collation;
        //public string Edition;
        //public string MasterDBPath;
        //public string MasterDBLogPath;
        public string Catalog;
        public string Name;
        //public string OSVersion;
        //public string ProcessorUsage;
        //public string Processors;
        //public string Product;
        //public string ProductLevel;
        //public string RootDirectory;
        //public string ServiceAccount;
        //public string ServiceStartMode;
        //public string SqlCharSetName;
        //public string SqlDomainGroup;
        //public string SqlSortOrderName;
        //public string VersionString;

         /// <summary>
        /// Initializes a new instance of the Server class.
        /// </summary>
        public LinkedServer(SMO.LinkedServer linkedServer)
        {
            AddinHelper.Common.WriteToDebugWindow(string.Format("SMOH.{0}({1})", "LinkedServer", linkedServer));
            // NB. Not all properties are available depending on login credentials
            // Wrap in TC blocks.

            //try
            //{
            //     BackupDirectory = server.BackupDirectory;
            //}
            //catch(Exception)
            //{
            //    BackupDirectory = "<No Access>";
            //}

            //try
            //{
            //    Catalog = server.Catalog;
            //}
            //catch (Exception)
            //{
            //    Catalog = "<No Access>";
            //}

            //try
            //{
            //    Edition = server.Edition;
            //}
            //catch(Exception)
            //{
            //    Edition = "<No Access>";
            //}

            //try
            //{
            //    MasterDBPath = server.MasterDBPath;
            //}
            //catch (Exception)
            //{
            //    MasterDBPath = "<No Access>";
            //}

            //try
            //{
            //    MasterDBLogPath = server.MasterDBLogPath;
            //}
            //catch (Exception)
            //{
            //    MasterDBLogPath = "<No Access>";
            //}

            Name = linkedServer.Name;

            //try
            //{
            //    OSVersion = server.OSVersion;
            //}
            //catch (Exception)
            //{
            //    OSVersion = "<No Access>";
            //}

            //try
            //{
            //    Processors = server.Processors.ToString();
            //}
            //catch (Exception)
            //{
            //    Processors = "<No Access>";
            //}

            //try
            //{
            //    ProcessorUsage = server.ProcessorUsage.ToString();
            //}
            //catch (Exception)
            //{
            //    ProcessorUsage = "<No Access>";
            //}

            //try
            //{
            //    Product = server.Product;
            //}
            //catch (Exception)
            //{
            //    Product = "<No Access>";
            //}

            //try
            //{
            //    ProductLevel = server.ProductLevel;
            //}
            //catch (Exception)
            //{
            //    ProductLevel = "<No Access>";
            //}

            //try
            //{
            //    RootDirectory = server.RootDirectory;
            //}
            //catch (Exception)
            //{
            //    RootDirectory = "<No Access>";
            //}

            //try
            //{
            //    ServiceAccount = server.ServiceAccount;
            //}
            //catch (Exception)
            //{
            //    ServiceAccount = "<No Access>";
            //}

            //try
            //{
            //    switch (server.ServiceStartMode)
            //    {
            //        case SMO.ServiceStartMode.Auto:
            //            ServiceStartMode = "Auto";
            //            break;
            //        case SMO.ServiceStartMode.Boot:
            //            ServiceStartMode = "Boot";
            //            break;
            //        case SMO.ServiceStartMode.Disabled:
            //            ServiceStartMode = "Disabled";
            //            break;
            //        case SMO.ServiceStartMode.Manual:
            //            ServiceStartMode = "Manual";
            //            break;
            //        case SMO.ServiceStartMode.System:
            //            ServiceStartMode = "System";
            //            break;
            //    }
            //}
            //catch (Exception)
            //{
            //    ServiceStartMode = "<No Access>";
            //}

            //try
            //{
            //    SqlCharSetName = server.SqlCharSetName;
            //}
            //catch (Exception)
            //{
            //    SqlCharSetName = "<No Access>";
            //}

            //try
            //{
            //    SqlDomainGroup = server.SqlDomainGroup;
            //}
            //catch (Exception)
            //{
            //    SqlDomainGroup = "<No Access>";
            //}

            //try
            //{
            //    SqlSortOrderName = server.SqlSortOrderName;
            //}
            //catch (Exception)
            //{
            //    SqlSortOrderName = "<No Access>";
            //}

            //try
            //{
            //    VersionString = server.VersionString;
            //}
            //catch(Exception)
            //{
            //    VersionString = "<No Access>";
            //}
            
            //Databases = new Dictionary<string, Database>();

            //foreach (SMO.Database database in server.Databases)
            //{
            //    SMOHelper.Database db = new SMOHelper.Database(database);

            //    Databases.Add(database.Name, db);
            //}

            //Logins = new Dictionary<string, Login>();

            //foreach (SMO.Login login in server.Logins)
            //{
            //    SMOHelper.Login li = new SMOHelper.Login(login);

            //    Logins.Add(login.Name, li);
            //}

            //ServerRoles = new Dictionary<string, ServerRole>();

            //foreach (SMO.ServerRole role in server.Roles)
            //{
            //    SMOHelper.ServerRole serverRole = new SMOHelper.ServerRole(role);

            //    ServerRoles.Add(role.Name, serverRole);
            //}
        }
    }
}
