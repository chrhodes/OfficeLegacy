using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SMO = Microsoft.SqlServer.Management.Smo;

namespace SMOHelper
{
    public class Database
    {
        #region Initialization

        /// <summary>
        /// Initializes a new instance of the Database class.
        /// </summary>
        public Database(SMO.Database database)
        {
            AddinHelper.Common.WriteToDebugWindow(string.Format("SMOH.{0}({1})", "Database", database));
            // NB. Not all properties are available depending on login credentials
            // Wrap in TC blocks.

            _Database = database;


            try { ActiveConnections = database.ActiveConnections.ToString(); }    catch(Exception) { ActiveConnections = "<No Access>"; }
            try { AnsiNullDefault = database.AnsiNullDefault.ToString(); }    catch(Exception) { AnsiNullDefault = "<No Access>"; }
            try { AnsiNullsEnabled = database.AnsiNullsEnabled.ToString(); }    catch(Exception) { AnsiNullsEnabled = "<No Access>"; }
            try { AnsiPaddingEnabled = database.AnsiPaddingEnabled.ToString(); }    catch(Exception) { AnsiPaddingEnabled = "<No Access>"; }
            try { AnsiWarningsEnabled = database.AnsiWarningsEnabled.ToString(); }    catch(Exception) { AnsiWarningsEnabled = "<No Access>"; }
            try { AutoClose = database.AutoClose.ToString(); }    catch(Exception) { AutoClose = "<No Access>"; }
            try { AutoCreateStatisticsEnabled = database.AutoCreateStatisticsEnabled.ToString(); }    catch(Exception) { AutoCreateStatisticsEnabled = "<No Access>"; }
            try { AutoShrink = database.AutoShrink.ToString(); }    catch(Exception) { AutoShrink = "<No Access>"; }
            try { AutoUpdateStatisticsAsync = database.AutoUpdateStatisticsAsync.ToString();    }    catch(Exception) { AutoUpdateStatisticsAsync = "<No Access>"; }
            try { AutoUpdateStatisticsEnabled = database.AutoUpdateStatisticsEnabled.ToString();    }    catch(Exception) { AutoUpdateStatisticsEnabled = "<No Access>"; }
            try { BrokerEnabled = database.BrokerEnabled.ToString();    }    catch(Exception) { BrokerEnabled = "<No Access>"; }
            try { CaseSensitive = database.CaseSensitive.ToString();    }    catch(Exception) { CaseSensitive = "<No Access>"; }
            try { ChangeTrackingAutoCleanUp = database.ChangeTrackingAutoCleanUp.ToString();    }    catch(Exception) { ChangeTrackingAutoCleanUp = "<No Access>"; }
            try { ChangeTrackingEnabled = database.ChangeTrackingEnabled.ToString();    }    catch(Exception) { ChangeTrackingEnabled = "<No Access>"; }
            try { ChangeTrackingRetentionPeriod = database.ChangeTrackingRetentionPeriod.ToString();    }    catch(Exception) { ChangeTrackingRetentionPeriod = "<No Access>"; }
            try { ChangeTrackingRetentionPeriodUnits = database.ChangeTrackingRetentionPeriodUnits.ToString();    }    catch(Exception) { ChangeTrackingRetentionPeriodUnits = "<No Access>"; }
            try { CloseCursorsOnCommitEnabled = database.CloseCursorsOnCommitEnabled.ToString();    }    catch(Exception) { CloseCursorsOnCommitEnabled = "<No Access>"; }
            try { Collation = database.Collation.ToString();    }    catch(Exception) { Collation = "<No Access>"; }

            try
            {
                switch (database.CompatibilityLevel)
                {
                    case Microsoft.SqlServer.Management.Smo.CompatibilityLevel.Version60:
                        CompatibilityLevel = "Version60";
                        break;
                    case Microsoft.SqlServer.Management.Smo.CompatibilityLevel.Version65:
                        CompatibilityLevel = "Version65";
                        break;
                    case Microsoft.SqlServer.Management.Smo.CompatibilityLevel.Version70:
                        CompatibilityLevel = "Version70";
                        break;
                    case Microsoft.SqlServer.Management.Smo.CompatibilityLevel.Version80:
                        CompatibilityLevel = "Version80";
                        break;
                    case Microsoft.SqlServer.Management.Smo.CompatibilityLevel.Version90:
                        CompatibilityLevel = "Version90";
                        break;
                    case Microsoft.SqlServer.Management.Smo.CompatibilityLevel.Version100:
                        CompatibilityLevel = "Version100";
                        break;
                }
            }
            catch (Exception)
            {
                CompatibilityLevel = "<No Access>";
            }

            try { ConcatenateNullYieldsNull = database.ConcatenateNullYieldsNull.ToString(); }      catch(Exception) { ConcatenateNullYieldsNull = "<No Access>"; }
            try { CreateDate = database.CreateDate.ToString("yyyy-MM-dd hh:mm:ss"); }               catch(Exception) { CreateDate = "<No Access>"; }
            try { DatabaseGuid = database.DatabaseGuid.ToString(); }                                catch(Exception) { DatabaseGuid = "<No Access>"; }
            try { DatabaseSnapshotBaseName = database.DatabaseSnapshotBaseName; }                   catch(Exception) { DatabaseSnapshotBaseName = "<No Access>"; }
            try { DataSpaceUsage = database.DataSpaceUsage.ToString(); }                            catch(Exception) { DataSpaceUsage = "<No Access>"; }
            try { DateCorrelationOptimization = database.DateCorrelationOptimization.ToString(); }  catch(Exception) { DateCorrelationOptimization = "<No Access>"; }
            try { DboLogin = database.DboLogin.ToString(); }                                        catch(Exception) { DboLogin = "<No Access>"; }
            try { DefaultFileGroup = database.DefaultFileGroup; }                                   catch(Exception) { DefaultFileGroup = "<No Access>";      }
            try { DefaultFileStreamFileGroup = database.DefaultFileStreamFileGroup.ToString(); }    catch(Exception) { DefaultFileStreamFileGroup = "<No Access>"; }
            try { DefaultFullTextCatalog = database.DefaultFullTextCatalog.ToString();    }         catch(Exception) { DefaultFullTextCatalog = "<No Access>"; }
            try { DefaultSchema = database.DefaultSchema.ToString(); }                              catch(Exception) { DefaultSchema = "<No Access>"; }
            try { EncryptionEnabled = database.EncryptionEnabled.ToString(); }                      catch(Exception) { EncryptionEnabled = "<No Access>"; }
            try { HonorBrokerPriority = database.HonorBrokerPriority.ToString(); }                  catch(Exception) { HonorBrokerPriority = "<No Access>"; }
            try { ID = database.ID.ToString(); }                                                    catch(Exception) { ID = "<No Access>";  }
            try { IndexSpaceUsage = database.IndexSpaceUsage.ToString(); }                          catch(Exception) { IndexSpaceUsage = "<No Access>";  }
            try { IsAccessible = database.IsAccessible.ToString(); }                                catch(Exception) { IsAccessible = "<No Access>";  }
            try { IsDatabaseSnapshot = database.IsDatabaseSnapshot.ToString(); }                    catch(Exception) { IsDatabaseSnapshot = "<No Access>";  }
            try { IsDatabaseSnapshotBase = database.IsDatabaseSnapshotBase.ToString(); }            catch(Exception) { IsDatabaseSnapshotBase = "<No Access>";  }
            try { IsDbAccessAdmin = database.IsDbAccessAdmin.ToString(); }                          catch(Exception) { IsDbAccessAdmin = "<No Access>";  }
            try { IsDbBackupOperator = database.IsDbBackupOperator.ToString(); }                    catch(Exception) { IsDbBackupOperator = "<No Access>";  }
            try { IsDbDatareader = database.IsDbDatareader.ToString(); }                            catch(Exception) { IsDbDatareader = "<No Access>";  }
            try { IsDbDatawriter = database.IsDbDatawriter.ToString(); }                            catch(Exception) { IsDbDatawriter = "<No Access>";  }
            try { IsDbDdlAdmin = database.IsDbDdlAdmin.ToString(); }                                catch(Exception) { IsDbDdlAdmin = "<No Access>";  }
            try { IsDbDenyDatareader = database.ID.ToString(); }                                    catch(Exception) { IsDbDenyDatareader = "<No Access>";  }
            try { IsDbDenyDatawriter = database.IsDbDenyDatawriter.ToString(); }                    catch(Exception) { IsDbDenyDatawriter = "<No Access>";  }
            try { IsDbOwner = database.IsDbOwner.ToString(); }                                      catch(Exception) { IsDbOwner = "<No Access>";  }
            try { IsDbSecurityAdmin = database.IsDbSecurityAdmin.ToString(); }                      catch(Exception) { IsDbSecurityAdmin = "<No Access>";  }
            try { IsFullTextEnabled = database.IsFullTextEnabled.ToString(); }                      catch(Exception) { IsFullTextEnabled = "<No Access>";  }
            try { IsMailHost = database.IsMailHost.ToString(); }                                    catch(Exception) { IsMailHost = "<No Access>";  }
            try { IsManagementDataWarehouse = database.IsManagementDataWarehouse.ToString(); }      catch(Exception) { IsManagementDataWarehouse = "<No Access>";  }
            try { IsMirroringEnabled = database.IsMirroringEnabled.ToString(); }                    catch(Exception) { IsMirroringEnabled = "<No Access>";  }
            try { IsParameterizationForced = database.IsParameterizationForced.ToString(); }        catch(Exception) { IsParameterizationForced = "<No Access>";  }
            try { IsReadCommittedSnapshotOn = database.IsReadCommittedSnapshotOn.ToString(); }      catch(Exception) { IsReadCommittedSnapshotOn = "<No Access>";  }
            try { IsSystemObject = database.IsSystemObject.ToString(); }                            catch(Exception) { IsSystemObject = "<No Access>";  }
            //try { IsTouched = database.IsTouched.ToString(); }                                    catch(Exception) { IsTouched = "<No Access>";  }
            try { IsUpdateable = database.IsUpdateable.ToString(); }                                catch(Exception) { IsUpdateable = "<No Access>";  }
            try { IsVarDecimalStorageFormatEnabled = database.IsVarDecimalStorageFormatEnabled.ToString(); }  catch(Exception) { IsVarDecimalStorageFormatEnabled = "<No Access>";  }
            try { LastBackupDate = database.LastBackupDate.ToString(); }                            catch(Exception) { LastBackupDate = "<No Access>";  }
            try { LastDifferentialBackupDate = database.LastDifferentialBackupDate.ToString(); }    catch(Exception) { LastDifferentialBackupDate = "<No Access>";  }
            try { LastLogBackupDate = database.LastLogBackupDate.ToString(); }                      catch(Exception) { LastLogBackupDate = "<No Access>";  }
            try { LocalCursorsDefault = database.LocalCursorsDefault.ToString(); }                  catch(Exception) { LocalCursorsDefault = "<No Access>";  }
            try { LogReuseWaitStatus = database.LogReuseWaitStatus.ToString(); }                    catch(Exception) { LogReuseWaitStatus = "<No Access>";  }
            try { MirroringFailoverLogSequenceNumber = database.MirroringFailoverLogSequenceNumber.ToString(); }  catch(Exception) { MirroringFailoverLogSequenceNumber = "<No Access>";  }
            try { MirroringID = database.MirroringID.ToString(); }                                  catch(Exception) { MirroringID = "<No Access>";  }
            try { MirroringPartner = database.MirroringPartner.ToString(); }                        catch(Exception) { MirroringPartner = "<No Access>";  }
            try { MirroringPartnerInstance = database.MirroringPartnerInstance.ToString(); }        catch(Exception) { MirroringPartnerInstance = "<No Access>";  }
            try { MirroringRedoQueueMaxSize = database.MirroringRedoQueueMaxSize.ToString(); }      catch(Exception) { MirroringRedoQueueMaxSize = "<No Access>";  }
            try { MirroringRoleSequence = database.MirroringRoleSequence.ToString(); }              catch(Exception) { MirroringRoleSequence = "<No Access>";  }
            try { MirroringSafetyLevel = database.MirroringSafetyLevel.ToString(); }                catch(Exception) { MirroringSafetyLevel = "<No Access>";  }
            try { MirroringSafetySequence = database.MirroringSafetySequence.ToString(); }          catch(Exception) { MirroringSafetySequence = "<No Access>";  }
            try { MirroringStatus = database.MirroringStatus.ToString(); }                          catch(Exception) { MirroringStatus = "<No Access>";  }
            try { MirroringTimeout = database.MirroringTimeout.ToString(); }                        catch(Exception) { MirroringTimeout = "<No Access>";  }
            try { MirroringWitness = database.MirroringWitness.ToString(); }                        catch(Exception) { MirroringWitness = "<No Access>";  }
            try { MirroringWitnessStatus = database.MirroringWitnessStatus.ToString(); }            catch(Exception) { MirroringWitnessStatus = "<No Access>";  }
            try { Name = database.Name; }                                                           catch(Exception) { Name = "<No Access>"; }
            try { NumericRoundAbortEnabled = database.NumericRoundAbortEnabled.ToString(); }        catch(Exception) { NumericRoundAbortEnabled = "<No Access>"; }
            try { Owner = database.Owner; }                                                         catch(Exception) { Owner = "<No Access>"; }
            try { PageVerify = database.PageVerify.ToString(); }                                    catch(Exception) { PageVerify = "<No Access>"; }
            try { Parent = database.Parent.Name; }                                                  catch(Exception) { Parent = "<No Access>"; }
            try { PrimaryFilePath = database.PrimaryFilePath.ToString(); }                          catch(Exception) { PrimaryFilePath = "<No Access>"; }
            try { QuotedIdentifiersEnabled = database.QuotedIdentifiersEnabled.ToString(); }        catch(Exception) { QuotedIdentifiersEnabled = "<No Access>"; }
            try { ReadOnly = database.ReadOnly.ToString(); }                                        catch(Exception) { ReadOnly = "<No Access>"; }
            try { RecoveryForkGuid = database.RecoveryForkGuid.ToString(); }                        catch(Exception) { RecoveryForkGuid = "<No Access>"; }
            try { RecoveryModel = database.RecoveryModel.ToString(); }                              catch(Exception) { RecoveryModel = "<No Access>"; }
            try { RecursiveTriggersEnabled = database.RecursiveTriggersEnabled.ToString(); }        catch(Exception) { RecursiveTriggersEnabled = "<No Access>"; }
            try { Size = database.Size.ToString(); }                                                catch(Exception) { Size = "<No Access>"; }
            try { SnapshotIsolationState = database.SnapshotIsolationState.ToString(); }            catch(Exception) { SnapshotIsolationState = "<No Access>"; }
            try { SpaceAvailable = database.SpaceAvailable.ToString(); }                            catch(Exception) { SpaceAvailable = "<No Access>"; }
            try { State = database.State.ToString(); }                                              catch(Exception) { State = "<No Access>"; }
            try { Status = database.Status.ToString(); }                                            catch(Exception) { Status = "<No Access>"; }
            try { UserName = database.UserName.ToString(); }                                        catch(Exception) { UserName = "<No Access>"; }
            try { Version = database.Version.ToString(); }                                          catch(Exception) { Version = "<No Access>"; }

            try { ExtendedProperties = database.ExtendedProperties; }           catch (Exception)  {       }
        }

        #endregion

        public SMO.Database _Database;

        public string ActiveConnections;
        public string AnsiNullDefault;
        public string AnsiNullsEnabled;
        public string AnsiPaddingEnabled;
        public string AnsiWarningsEnabled;
        public string AutoClose;
        public string AutoCreateStatisticsEnabled;
        public string AutoShrink;
        public string AutoUpdateStatisticsAsync;
        public string AutoUpdateStatisticsEnabled;
        public string BrokerEnabled;
        public string CaseSensitive;
        public string ChangeTrackingAutoCleanUp;
        public string ChangeTrackingEnabled;
        public string ChangeTrackingRetentionPeriod;
        public string ChangeTrackingRetentionPeriodUnits;
        public string CloseCursorsOnCommitEnabled;
        public string Collation;
        public string CompatibilityLevel;
        public string ConcatenateNullYieldsNull;
        public string CreateDate;
        public string DatabaseGuid;
        public string DatabaseSnapshotBaseName;
        public string DataSpaceUsage;
        public string DateCorrelationOptimization;
        public string DboLogin;
        public string DefaultFileGroup;
        public string DefaultFileStreamFileGroup;
        public string DefaultFullTextCatalog;
        public string DefaultSchema;
        public string EncryptionEnabled;
        public string HonorBrokerPriority;
        public string ID;
        public string IndexSpaceUsage;
        public string IsAccessible;
        public string IsDatabaseSnapshot;
        public string IsDatabaseSnapshotBase;
        public string IsDbAccessAdmin;
        public string IsDbBackupOperator;
        public string IsDbDatareader;
        public string IsDbDatawriter;
        public string IsDbDdlAdmin;
        public string IsDbDenyDatareader;
        public string IsDbDenyDatawriter;
        public string IsDbOwner;
        public string IsDbSecurityAdmin;
        public string IsFullTextEnabled;
        public string IsMailHost;
        public string IsManagementDataWarehouse;
        public string IsMirroringEnabled;
        public string IsParameterizationForced;
        public string IsReadCommittedSnapshotOn;
        public string IsSystemObject;
        //public string IsTouched;
        public string IsUpdateable;
        public string IsVarDecimalStorageFormatEnabled;
        public string LastBackupDate;
        public string LastDifferentialBackupDate;
        public string LastLogBackupDate;
        public string LocalCursorsDefault;
        public string LogReuseWaitStatus;
        public string MirroringFailoverLogSequenceNumber;
        public string MirroringID;
        public string MirroringPartner;
        public string MirroringPartnerInstance;
        public string MirroringRedoQueueMaxSize;
        public string MirroringRoleSequence;
        public string MirroringSafetyLevel;
        public string MirroringSafetySequence;
        public string MirroringStatus;
        public string MirroringTimeout;
        public string MirroringWitness;
        public string MirroringWitnessStatus;
        public string Name;
        public string NumericRoundAbortEnabled;
        public string ObjectInSpace;
        public string Owner;
        public string PageVerify;
        public string Parent;
        public string PrimaryFilePath;
        public string QuotedIdentifiersEnabled;
        public string ReadOnly;
        public string RecoveryForkGuid;
        public string RecoveryModel;
        public string RecursiveTriggersEnabled;
        public string Size;
        public string SnapshotIsolationState;
        public string SpaceAvailable;
        public string State;
        public string Status;
        public string UserName;
        public string Version;

        public bool IncludeSystemViews = false;
        public bool IncludeSystemStoredProcedures = false;

        private SMO.ExtendedPropertyCollection _ExtendedProperties;
        public SMO.ExtendedPropertyCollection ExtendedProperties
        {
	        get {	return _ExtendedProperties; }
	        set { _ExtendedProperties = value;	}
        }
   
        private Dictionary<string, FileGroup> _FileGroups;
        public Dictionary<string, FileGroup> FileGroups
        {
            get
            {
                if (null == _FileGroups)
                {
                    _FileGroups = new Dictionary<string, FileGroup>();

                    foreach (SMO.FileGroup fileGroup in _Database.FileGroups)
                    {
                        SMOHelper.FileGroup fg = new SMOHelper.FileGroup(fileGroup);
                        _FileGroups.Add(fileGroup.Name, fg);
                    }
                }

                return _FileGroups;
            }
            set
            {
                _FileGroups = value;
            }
        }

        private Dictionary<string, StoredProcedure> _StoredProcedures;
        public Dictionary<string, StoredProcedure> StoredProcedures
        {
            get
            {
                if (null == _StoredProcedures)
                {
                    _StoredProcedures = new Dictionary<string, StoredProcedure>();

                    foreach (SMO.StoredProcedure storedProcedure in _Database.StoredProcedures)
                    {
                        if (storedProcedure.IsSystemObject && !IncludeSystemStoredProcedures)
                        {
                            continue;
                        }

                        SMOHelper.StoredProcedure sp = new SMOHelper.StoredProcedure(storedProcedure);
                        _StoredProcedures.Add(storedProcedure.Name, sp);
                    }
                }

                return _StoredProcedures;
            }
            set
            {
                _StoredProcedures = value;
            }
        }

        private Dictionary<string, Table> _Tables;
        public Dictionary<string, Table> Tables
        {
            get
            {
                if (null == _Tables)
                {
                    _Tables = new Dictionary<string, Table>();

                    foreach (SMO.Table table in _Database.Tables)
                    {
                        SMOHelper.Table tbl = new SMOHelper.Table(table);

                        _Tables.Add(table.Name, tbl);
                    }
                }

                return _Tables;
            }
            set
            {
                _Tables = value;
            }
        }

        private Dictionary<string, View> _Views;
        public Dictionary<string, View> Views
        {
            get
            {
                if (null == _Views)
                {
                    _Views = new Dictionary<string, View>();

                    foreach (SMO.View View in _Database.Views)
                    {
                        if (View.IsSystemObject && ! IncludeSystemViews)
                        {
                        	continue;
                        }

                        SMOHelper.View tbl = new SMOHelper.View(View);

                        _Views.Add(View.Name, tbl);
                    }
                }

                return _Views;
            }
            set
            {
                _Views = value;
            }
        }
    }
}
