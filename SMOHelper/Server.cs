using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SMO=Microsoft.SqlServer.Management.Smo;

namespace SMOHelper
{
    public class Server
    {

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the Server class.
        /// </summary>
        public Server(SMO.Server server)
        {
            //AddinHelper.Common.WriteToDebugWindow(string.Format("SMOH.{0}({1})", "Server", server));
            _Server = server;

            // NB. Not all properties are available depending on login credentials
            // Wrap in TC blocks.

            try { BackupDirectory = server.BackupDirectory; }                           catch (Exception) { BackupDirectory = "<No Access>"; }
            try { BrowserServiceAccount = server.BrowserServiceAccount; }               catch (Exception) { BrowserServiceAccount = "<No Access>"; }
            try { BrowserStartMode = server.BrowserStartMode.ToString(); }              catch (Exception) { BrowserStartMode = "<No Access>"; }
            try { BuildClrVersionString = server.BuildClrVersionString; }               catch (Exception) { BuildClrVersionString = "<No Access>"; }
            try { BuildNumber = server.BuildNumber.ToString(); }                        catch (Exception) { BuildNumber = "<No Access>"; }
            try { Collation = server.Collation; }                                       catch (Exception) { Collation = "<No Access>"; }
            try { ComparisonStyle = server.ComparisonStyle.ToString(); }                catch (Exception) { ComparisonStyle = "<No Access>"; }
            try { ComputerNamePhysicalNetBIOS = server.ComputerNamePhysicalNetBIOS; }   catch (Exception) { ComputerNamePhysicalNetBIOS = "<No Access>"; }
            try { DefaultFile = server.DefaultFile; }                                   catch (Exception) { DefaultFile = "<No Access>"; }
            try { DefaultLog = server.DefaultLog; }                                     catch (Exception) { DefaultLog = "<No Access>"; }
            try { Edition = server.Edition; }                                           catch (Exception) { Edition = "<No Access>";  }
            try { EngineEdition = server.EngineEdition.ToString(); }                    catch (Exception) { EngineEdition = "<No Access>";  }
            try { ErrorLogPath = server.ErrorLogPath; }                                 catch (Exception) { ErrorLogPath = "<No Access>"; }
            try { FilestreamLevel = server.FilestreamLevel.ToString(); }                catch (Exception) { FilestreamLevel = "<No Access>"; }
            try { FilestreamShareName = server.FilestreamShareName; }                   catch (Exception) { FilestreamShareName = "<No Access>"; }
            try { InstallDataDirectory = server.InstallDataDirectory; }                 catch (Exception) { InstallDataDirectory = "<No Access>"; }
            try { InstallSharedDirectory = server.InstallSharedDirectory; }             catch (Exception) { InstallSharedDirectory = "<No Access>"; }
            try { InstanceName = server.InstanceName; }                                 catch (Exception) { InstanceName = "<No Access>"; }
            try { IsCaseSensitive = server.IsCaseSensitive.ToString(); }                catch (Exception) { IsCaseSensitive = "<No Access>"; }
            try { IsClustered = server.IsClustered.ToString(); }                        catch (Exception) { IsClustered = "<No Access>"; }
            try { IsFullTextInstalled = server.IsFullTextInstalled.ToString(); }        catch (Exception) { IsFullTextInstalled = "<No Access>"; }
            try { IsSingleUser = server.IsSingleUser.ToString(); }                      catch (Exception) { IsSingleUser = "<No Access>"; }
            try { LoginMode = server.LoginMode.ToString(); }                            catch (Exception) { LoginMode = "<No Access>"; }
            try { MasterDBPath = server.MasterDBPath; }                                 catch (Exception) { MasterDBPath = "<No Access>"; }
            try { MasterDBLogPath = server.MasterDBLogPath; }                           catch (Exception) { MasterDBLogPath = "<No Access>"; }
            try { MaxPrecision = server.MaxPrecision.ToString(); }                      catch (Exception) { MaxPrecision = "<No Access>"; }
            try { Name = server.Name; }                                                 catch (Exception) { Name = "<No Access>"; }
            try { NamedPipesEnabled = server.NamedPipesEnabled.ToString(); }            catch (Exception) { NamedPipesEnabled = "<No Access>"; }
            try { NetName = server.NetName; }                                           catch (Exception) { NetName = "<No Access>"; }
            try { NumberOfLogFiles = server.NumberOfLogFiles.ToString(); }              catch (Exception) { NumberOfLogFiles = "<No Access>"; }
            try { OSVersion = server.OSVersion; }                                       catch (Exception) { OSVersion = "<No Access>"; }
            try { PerfMonMode = server.PerfMonMode.ToString(); }                        catch (Exception) { PerfMonMode = "<No Access>"; }
            try { PhysicalMemory = server.PhysicalMemory.ToString(); }                  catch (Exception) { PhysicalMemory = "<No Access>"; }
            try { PhysicalMemoryUsageInKB = server.PhysicalMemoryUsageInKB.ToString(); } catch (Exception) { PhysicalMemoryUsageInKB = "<No Access>"; }
            try { Platform = server.Platform; }                                         catch (Exception) { Platform = "<No Access>"; }
            try { Processors = server.Processors.ToString(); }                          catch (Exception) { Processors = "<No Access>"; }
            try { ProcessorUsage = server.ProcessorUsage.ToString(); }                  catch (Exception) { ProcessorUsage = "<No Access>"; }
            try { Product = server.Product; }                                           catch (Exception) { Product = "<No Access>"; }
            try { ProductLevel = server.ProductLevel; }                                 catch (Exception) { ProductLevel = "<No Access>"; }
            try { ResourceVersionString = server.ResourceVersionString; }               catch (Exception) { ResourceVersionString = "<No Access>"; }
            try { RootDirectory = server.RootDirectory; }                               catch (Exception) { RootDirectory = "<No Access>"; }
            try { ServerType = server.ServerType.ToString(); }                          catch (Exception) { ServerType = "<No Access>"; }
            try { ServiceAccount = server.ServiceAccount.ToString(); }                  catch (Exception) { ServiceAccount = "<No Access>"; }
            try { ServiceInstanceId = server.ServiceInstanceId.ToString(); }            catch (Exception) { ServiceInstanceId = "<No Access>"; }
            try { ServiceName = server.ServiceName.ToString(); }                        catch (Exception) { ServiceName = "<No Access>"; }

            try
            {
                switch (server.ServiceStartMode)
                {
                    case SMO.ServiceStartMode.Auto:
                        ServiceStartMode = "Auto";
                        break;
                    case SMO.ServiceStartMode.Boot:
                        ServiceStartMode = "Boot";
                        break;
                    case SMO.ServiceStartMode.Disabled:
                        ServiceStartMode = "Disabled";
                        break;
                    case SMO.ServiceStartMode.Manual:
                        ServiceStartMode = "Manual";
                        break;
                    case SMO.ServiceStartMode.System:
                        ServiceStartMode = "System";
                        break;
                }
            }
            catch (Exception) { ServiceStartMode = "<No Access>";  }

            try { SqlCharSetName = server.SqlCharSetName; }                             catch (Exception) { SqlCharSetName = "<No Access>"; }
            try { SqlDomainGroup = server.SqlDomainGroup; }                             catch (Exception) { SqlDomainGroup = "<No Access>"; }
            try { SqlSortOrderName = server.SqlSortOrderName; }                         catch (Exception) { SqlSortOrderName = "<No Access>"; }
            try { Status = server.Status.ToString(); }                                  catch (Exception) { Status = "<No Access>"; }
            try { TcpEnabled = server.TcpEnabled.ToString(); }                          catch (Exception) { TcpEnabled = "<No Access>"; }
            try { VersionString = server.VersionString; }                               catch (Exception) { VersionString = "<No Access>"; }

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

        #endregion
        
        #region Fields and Properties

        SMO.Server _Server;

        public string BackupDirectory;
        public string BrowserServiceAccount;
        public string BrowserStartMode;
        public string BuildClrVersionString;
        public string BuildNumber;
        public string Collation;
        public string ComparisonStyle;
        public string ComputerNamePhysicalNetBIOS;
        public string DefaultFile;
        public string DefaultLog;
        public string Edition;
        public string EngineEdition;
        public string ErrorLogPath;
        public string FilestreamLevel;
        public string FilestreamShareName;
        public string InstallDataDirectory;
        public string InstallSharedDirectory;
        public string InstanceName;
        public string IsCaseSensitive;
        public string IsClustered;
        public string IsFullTextInstalled;
        public string IsSingleUser;
        public string LoginMode;
        public string MasterDBPath;
        public string MasterDBLogPath;
        public string MaxPrecision;
        public string Name;
        public string NamedPipesEnabled;
        public string NetName;
        public string NumberOfLogFiles;
        public string OSVersion;
        public string PerfMonMode;
        public string PhysicalMemory;
        public string PhysicalMemoryUsageInKB;
        public string Platform;
        public string Processors;
        public string ProcessorUsage;
        public string Product;
        public string ProductLevel;
        public string ResourceVersionString;
        public string RootDirectory;
        public string ServerType;
        public string ServiceAccount;
        public string ServiceInstanceId;
        public string ServiceName;
        public string ServiceStartMode;
        public string SqlCharSetName;
        public string SqlDomainGroup;
        public string SqlSortOrderName;
        public string Status;
        public string TcpEnabled;
        public string VersionString;

        private Dictionary<string, Database> _Databases;
        public Dictionary<string, Database> Databases
        {
            get
            {
                if (null == _Databases)
                {
                    _Databases = new Dictionary<string, Database>();

                    foreach (SMO.Database database in _Server.Databases)
                    {
                        SMOHelper.Database db = new SMOHelper.Database(database);
                        _Databases.Add(database.Name, db);
                    }
                }

                return _Databases;
            }
            set
            {
                _Databases = value;
            }
        }

        private Dictionary<string, Endpoint> _Endpoints;
        public Dictionary<string, Endpoint> Endpoints
        {
            get
            {
                if (null == _Endpoints)
                {
                    _Endpoints = new Dictionary<string, Endpoint>();

                    foreach (SMO.Endpoint endPoint in _Server.Endpoints)
                    {
                        SMOHelper.Endpoint ep = new SMOHelper.Endpoint(endPoint);
                        _Endpoints.Add(endPoint.Name, ep);
                    }
                }

                return _Endpoints;
            }
            set
            {
                _Endpoints = value;
            }
        }

        private Dictionary<string, LinkedServer> _LinkedServers;
        public Dictionary<string, LinkedServer> LinkedServers
        {
            get
            {
                if (null == _LinkedServers)
                {
                    _LinkedServers = new Dictionary<string, LinkedServer>();

                    foreach (SMO.LinkedServer linkedServer in _Server.LinkedServers)
                    {
                        SMOHelper.LinkedServer ls = new SMOHelper.LinkedServer(linkedServer);
                        _LinkedServers.Add(linkedServer.Name, ls);
                    }
                }

                return _LinkedServers;
            }
            set
            {
                _LinkedServers = value;
            }
        }

        private Dictionary<string, Login> _Logins;
        public Dictionary<string, Login> Logins
        {
            get
            {
                if (null == _Logins)
                {
                    _Logins = new Dictionary<string, Login>();

                    foreach (SMO.Login login in _Server.Logins)
                    {
                        SMOHelper.Login li = new SMOHelper.Login(login);
                        _Logins.Add(login.Name, li);
                    }
                }

                return _Logins;
            }
            set
            {
                _Logins = value;
            }
        }

        private Dictionary<string, ServerRole> _ServerRoles;
        public Dictionary<string, ServerRole> ServerRoles
        {
            get
            {
                if (null == _ServerRoles)
                {
                    _ServerRoles = new Dictionary<string, ServerRole>();

                    foreach (SMO.ServerRole role in _Server.Roles)
                    {
                        SMOHelper.ServerRole r = new SMOHelper.ServerRole(role);
                        _ServerRoles.Add(role.Name, r);
                    }
                }

                return _ServerRoles;
            }
            set
            {
                _ServerRoles = value;
            }
        }

        #endregion
    }
}
