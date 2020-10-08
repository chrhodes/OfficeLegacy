using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using Microsoft.Office.Interop.Excel;

using Microsoft.SqlServer.Management.Common;
using SMO=Microsoft.SqlServer.Management.Smo;
using SMOWMI=Microsoft.SqlServer.Management.Smo.Wmi;

using SMOH=SMOHelper;

using ExcelHlp=AddinHelper.Excel;


// TODO
// Consider passing in range instead of startingRow

namespace SupportTools_Excel.User_Interface.Task_Panes
{
    public partial class TaskPane_SQLSMO : UserControl
    {
        SMO.Server _SMOServer;      //  This is the real one
        SMOH.Server _SMOHServer;    // This is the one that hides the access restrictions

        public TaskPane_SQLSMO()
        {
            InitializeComponent();
        }

        #region Event Handlers

        private void btnCreateDatabase_Click(object sender, EventArgs e)
        {
            CreateDatabase();
        }

        private void btnCreateDatabaseInfoWorkSheet_Click(object sender, EventArgs e)
        {
            try
            {
                ExcelHlp.ScreenUpdatesOff();
                SMOH.Database dataBase = _SMOHServer.Databases[cbDatabases.Text];
                CreateWorksheet_DatabaseInfo(dataBase);
            }
            catch(Exception ex)
            {
                 MessageBox.Show(ex.ToString());  
            }
            finally
            {
                ExcelHlp.ScreenUpdatesOn(true);                
            }
        }

        private void btnCreateDatabaseInfoWorksheets_Click(object sender, EventArgs e)
        {
            try
            {
                ExcelHlp.ScreenUpdatesOff();
                CreateAllWorksheetsOf_DatabaseInfo(_SMOHServer);
            }
            catch(Exception ex)
            {
                 MessageBox.Show(ex.ToString());               
            }
            finally
            {
                ExcelHlp.ScreenUpdatesOn(true);                
            }
        }

        private void btnCreateInstanceInfoWorkSheet_Click(object sender, EventArgs e)
        {
            try
            {
                ExcelHlp.ScreenUpdatesOff();
                CreateWorksheet_InstanceInfo(_SMOHServer, chkListInstanceDetails.Checked);
            }
            catch(Exception ex)
            {
                 MessageBox.Show(ex.ToString());                 
            }
            finally
            {
                 ExcelHlp.ScreenUpdatesOn(true);               
            }
        }

        private void btnCreateStoredProcedureInfoWorksheet_Click(object sender, EventArgs e)
        {
            try
            {
                ExcelHlp.ScreenUpdatesOff();
                SMOH.Database dataBase = _SMOHServer.Databases[cbDatabases.Text];
                SMOH.StoredProcedure storedProcedure = dataBase.StoredProcedures[cbStoredProcedures.Text];
                CreateWorksheet_StoredProcedureInfo(storedProcedure);
            }
            catch(Exception ex)
            {
                 MessageBox.Show(ex.ToString());                 
            }
            finally
            {
                ExcelHlp.ScreenUpdatesOn(true);                
            }
        }

        private void btnCreateStoredProcedureInfoWorksheets_Click(object sender, EventArgs e)
        {
            try
            {
                ExcelHlp.ScreenUpdatesOff();
                SMOH.Database dataBase = _SMOHServer.Databases[cbDatabases.Text];
                CreateAllWorksheeetsOf_StoredProcedureInfo(dataBase);
            }
            catch(Exception ex)
            {
                 MessageBox.Show(ex.ToString());                 
            }
            finally
            {
                ExcelHlp.ScreenUpdatesOn(true);                
            }
        }

        private void btnCreateTable_Click(object sender, EventArgs e)
        {
            CreateTable();
        }

        private void btnCreateTableInfoWorksheet_Click(object sender, EventArgs e)
        {
            try
            {
                ExcelHlp.ScreenUpdatesOff();
                SMOH.Database dataBase = _SMOHServer.Databases[cbDatabases.Text];
                SMOH.Table table = dataBase.Tables[cbTables.Text];
                CreateWorksheet_TableInfo(table);
            }
            catch(Exception ex)
            {
                 MessageBox.Show(ex.ToString());              
            }
            finally
            {
                 ExcelHlp.ScreenUpdatesOn(true);                   
            }
        }

        private void btnCreateTableInfoWorkSheets_Click(object sender, EventArgs e)
        {
            try
            {
                ExcelHlp.ScreenUpdatesOff();
                SMOH.Database dataBase = _SMOHServer.Databases[cbDatabases.Text];
                CreateAllWorkSheetsOf_TableInfo(dataBase);
            }
            catch(Exception ex)
            {
                 MessageBox.Show(ex.ToString());                 
            }
            finally
            {
                ExcelHlp.ScreenUpdatesOn(true);                 
            }
        }

        private void btnCreateViewInfoWorkSheet_Click(object sender, EventArgs e)
        {
            try
            {
                ExcelHlp.ScreenUpdatesOff();
                SMOH.Database dataBase = _SMOHServer.Databases[cbDatabases.Text];
                SMOH.View view = dataBase.Views[cbViews.Text];
                CreateWorksheet_ViewInfo(view);
            }
            catch(Exception ex)
            {
                 MessageBox.Show(ex.ToString());                
            }
            finally
            {
                ExcelHlp.ScreenUpdatesOn(true);                  
            }
        }

        private void btnCreateViewInfoWorksheets_Click(object sender, EventArgs e)
        {
            try
            {
                ExcelHlp.ScreenUpdatesOff();
                SMOH.Database dataBase = _SMOHServer.Databases[cbDatabases.Text];
                CreateAllWorksheetsOf_ViewInfo(dataBase);
            }
            catch(Exception ex)
            {
                 MessageBox.Show(ex.ToString());             
            }
            finally
            {
                 ExcelHlp.ScreenUpdatesOn(true);                    
            }
        }

        private void btnEnumAvailableSqlServers_Click(object sender, EventArgs e)
        {
            EnumAvailableSqlServers();
        }

        private void btnGetServerInfo_Click(object sender, EventArgs e)
        {
            GetServerInfo();
        }

        private void btnListDatabases_Click(object sender, EventArgs e)
        {
            //ExcelHlp.ScreenUpdatesOff();
            //Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet;
            //int currentRow = Globals.ThisAddIn.Application.ActiveCell.Row;

            //ListDataBases(ws, _SMOServer, currentRow);
            //ExcelHlp.ScreenUpdatesOn(true);
        }

        private void btnListRows_Table_Click(object sender, EventArgs e)
        {
            try
            {
                SMOH.Database dataBase = _SMOHServer.Databases[cbDatabases.Text];
                SMOH.Table table = dataBase.Tables[cbTables.Text];
                ListRows(table._table);
            }
            catch(Exception ex)
            {
                 MessageBox.Show(ex.ToString());               
            }
            finally
            {
                ExcelHlp.ScreenUpdatesOn(true);                   
            }
        }

        private void btnListRows_View_Click(object sender, EventArgs e)
        {
            try
            {
                SMOH.Database dataBase = _SMOHServer.Databases[cbDatabases.Text];
                SMOH.View view = dataBase.Views[cbViews.Text];
                ListRows(view._view);
            }
            catch(Exception ex)
            {
                 MessageBox.Show(ex.ToString());                 
            }
            finally
            {
                ExcelHlp.ScreenUpdatesOn(true);                 
            }
        }

        private void btnListStoredProcedures_Click(object sender, EventArgs e)
        {
            //ExcelHlp.ScreenUpdatesOff();
            //Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet;
            //int currentRow = Globals.ThisAddIn.Application.ActiveCell.Row;
            //SMO.Database dataBase = _SMOServer.Databases[cbDatabases.Text];

            //ListStoredProcedures(ws, dataBase, currentRow);
            //ExcelHlp.ScreenUpdatesOn(true);
        }

        private void btnListTables_Click(object sender, EventArgs e)
        {
            //ExcelHlp.ScreenUpdatesOff();
            //Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet;
            //int currentRow = Globals.ThisAddIn.Application.ActiveCell.Row;
            //SMO.Database dataBase = _SMOServer.Databases[cbDatabases.Text];

            //ListTables(ws, dataBase, currentRow);
            //ExcelHlp.ScreenUpdatesOn(true);
        }

        private void btnLogoff_Click(object sender, EventArgs e)
        {
            Logoff();
            btnLogon.Enabled = true;
            btnLogon.BackColor = SystemColors.Control;
            btnLogoff.Enabled = false;
            btnLogoff.BackColor = SystemColors.Control;
            lblInstancName.Text = "";
            gbInstanceOperations.Visible = false;
        }

        private void btnLogon_Click(object sender, EventArgs e)
        {
            if (Logon() == true)
            {
                btnLogoff.Enabled = true;
                btnLogoff.BackColor = Color.Green;
                btnLogon.Enabled = false;
                btnLogon.BackColor = Color.Green;
                lblInstancName.Text = ucDBInstanceList.InstanceName;
                gbInstanceOperations.Visible = true;
            }
        }

        private void cbDatabases_SelectedIndexChanged(object sender, EventArgs e)
        {
            cbTables.Items.Clear();

            foreach(string name in _SMOHServer.Databases[cbDatabases.Text].Tables.Keys)
            {
                cbTables.Items.Add(name);
            }

            cbViews.Items.Clear();

            foreach(SMOH.View view in _SMOHServer.Databases[cbDatabases.Text].Views.Values)
            {
                if(view.IsSystemObject && !ckIncludeSystemStoredProcedures.Checked)
                {
                    continue;
                }

                cbViews.Items.Add(view.Name);
            }

            cbStoredProcedures.Items.Clear();

            foreach(SMOH.StoredProcedure sp in _SMOHServer.Databases[cbDatabases.Text].StoredProcedures.Values)
            {
                if(sp.IsSystemObject && !ckIncludeSystemStoredProcedures.Checked)
                {
                    continue;
                }

                cbStoredProcedures.Items.Add(sp.Name);
            }
        }

        private void ckIncludeDBStoredProcedures_CheckedChanged(object sender, EventArgs e)
        {
            gbStoredProceduresOperations.Visible = ckIncludeDBStoredProcedures.Checked;
        }

        private void ckIncludeDBTables_CheckedChanged(object sender, EventArgs e)
        {
            gbTableOperations.Visible = ckIncludeDBTables.Checked;
        }

        private void ckIncludeDBViews_CheckedChanged(object sender, EventArgs e)
        {
            gbViewOperations.Visible = ckIncludeDBViews.Checked;
        }

        private void ckIntegratedSecurity_CheckedChanged(object sender, EventArgs e)
        {
            if (ckIntegratedSecurity.Checked)
            {
                lblUserName.Visible = false;
                cbUserName.Visible = false;

                lblPassword.Visible = false;
                txtPassword.Visible = false;
            }
            else
            {
                lblUserName.Visible = true;
                cbUserName.Visible = true;

                lblPassword.Visible = true;
                txtPassword.Visible = true;
            }
        }

        private void TaskPane_SQLSMO_Load(object sender, EventArgs e)
        {
            ucDBInstanceList.PopulateInstanceListFromFile(@"C:\Temp\DbInstanceList.xml");
        }

        private void ucDBInstanceList_InstanceSelectionChanged_Event()
        {
            txtServerName.Text = ucDBInstanceList.ServerName;

            // A new instance has been selected.  Decide what to do.
            //MessageBox.Show("TODO: InstanceChanged.  How to handle");
        }

        #endregion

        #region Main Function Routines

        private int AddSection_ColumnInfo(Worksheet ws, SMOH.Table table, int startingRow)
        {
            ExcelHlp.AddContentToCell(ws.Cells[startingRow++, 1], "Columns");

            int currentRow = startingRow;
            int col = 1;
            int rowsAdded = 0;

            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 35, currentRow, "Name", 12);
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 10, currentRow, "DataType", 12);
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 10, currentRow, "Maximum\nLength", 12);
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 10, currentRow, "Numeric\nPrecision", 12);
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 10, currentRow, "Numeric\nScale", 12);
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 10, currentRow, "Default", 12);
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++,  5, currentRow, "ID", 12);
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 12, currentRow, "Identity", 12);
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 12, currentRow, "Is\nPrimaryKey", 12);
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 12, currentRow, "Is\nForeignKey", 12);
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 12, currentRow, "Nullable", 12);

            rowsAdded++;

            rowsAdded += DisplayListOf_Columns(ws, table, ++currentRow);

            // -2 because we are starting from 'A' and we have already incremented above
            Char endColumn = (char)('A' + (col - 2));   

            string groupRange = string.Format("{0}:{1}", "C", endColumn);
            ws.Columns[groupRange].Group();
            ws.Columns[groupRange].Hidden = true;

            return ++rowsAdded;
        }

        private int AddSection_ColumnInfo(Worksheet ws, SMOH.View view, int startingRow)
        {
            ExcelHlp.AddContentToCell(ws.Cells[startingRow++, 1], "Columns");

            int currentRow = startingRow;
            int col = 1;
            int rowsAdded = 0;

            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 35, currentRow, "Name", 12);
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 10, currentRow, "DataType", 12);
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 10, currentRow, "Maximum\nLength", 12);
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 10, currentRow, "Numeric\nPrecision", 12);
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 10, currentRow, "Numeric\nScale", 12);
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 10, currentRow, "Default", 12);
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++,  5, currentRow, "ID", 12);
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 12, currentRow, "Identity", 12);
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 12, currentRow, "Is\nPrimaryKey", 12);
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 12, currentRow, "Is\nForeignKey", 12);
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 12, currentRow, "Nullable", 12);

            rowsAdded++;

            rowsAdded += DisplayListOf_Columns(ws, view, ++currentRow);

            Char endColumn = (char)('A' + (col - 2));  // -2 because we are starting from 'A' and we have already incremented above

            string groupRange = string.Format("{0}:{1}", "C", endColumn);
            ws.Columns[groupRange].Group();
            ws.Columns[groupRange].Hidden = true;

            return rowsAdded;
        }
  
        private int AddSection_DatabaseInfo(Worksheet ws, SMOH.Database dataBase, int startingRow)
        {
            Range rng = ws.Cells[startingRow, 1];

            int rowsAdded = 0;

            AddTitledInfo(rng.Offset[rowsAdded++, 0], "As of:", DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss"));
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "DB Name:", dataBase.Name);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "Instance:", dataBase.Parent);

            AddTitledInfo(rng.Offset[rowsAdded++, 0], "ActiveConnections:", dataBase.ActiveConnections);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "AnsiNullDefault:", dataBase.AnsiNullDefault);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "AnsiNullsEnabled:", dataBase.AnsiNullsEnabled);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "AnsiPaddingEnabled:", dataBase.AnsiPaddingEnabled);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "AnsiWarningsEnabled:", dataBase.AnsiWarningsEnabled);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "AutoClose:", dataBase.AutoClose);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "AutoCreateStatisticsEnabled:", dataBase.AutoCreateStatisticsEnabled);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "AutoShrink:", dataBase.AutoShrink);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "AutoUpdateStatisticsAsync:", dataBase.AutoUpdateStatisticsAsync);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "AutoUpdateStatisticsEnabled:", dataBase.AutoUpdateStatisticsEnabled);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "BrokerEnabled:", dataBase.BrokerEnabled);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "CaseSensitive:", dataBase.CaseSensitive);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "ChangeTrackingAutoCleanUp:", dataBase.ChangeTrackingAutoCleanUp);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "ChangeTrackingEnabled:", dataBase.ChangeTrackingEnabled);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "ChangeTrackingRetentionPeriod:", dataBase.ChangeTrackingRetentionPeriod);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "ChangeTrackingRetentionPeriodUnits:", dataBase.ChangeTrackingRetentionPeriodUnits);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "CloseCursorsOnCommitEnabled:", dataBase.CloseCursorsOnCommitEnabled);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "Collation:", dataBase.Collation);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "CompatibilityLevel:", dataBase.CompatibilityLevel);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "ConcatenateNullYieldsNull:", dataBase.ConcatenateNullYieldsNull);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "CreateDate:", dataBase.CreateDate);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "DatabaseGuid:", dataBase.DatabaseGuid);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "DatabaseSnapshotBaseName:", dataBase.DatabaseSnapshotBaseName);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "DataSpaceUsage:", dataBase.DataSpaceUsage);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "DateCorrelationOptimization:", dataBase.DateCorrelationOptimization);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "DboLogin:", dataBase.DboLogin);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "DefaultFileGroup:", dataBase.DefaultFileGroup);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "DefaultFileStreamFileGroup:", dataBase.DefaultFileStreamFileGroup);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "DefaultFullTextCatalog:", dataBase.DefaultFullTextCatalog);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "DefaultSchema:", dataBase.DefaultSchema);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "HonorBrokerPriority:", dataBase.HonorBrokerPriority);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "ID:", dataBase.ID);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "IndexSpaceUsage:", dataBase.IndexSpaceUsage);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "IsAccessible:", dataBase.IsAccessible);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "IsDatabaseSnapshot:", dataBase.IsDatabaseSnapshot);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "IsDatabaseSnapshotBase:", dataBase.IsDatabaseSnapshotBase);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "IsDbAccessAdmin:", dataBase.IsDbAccessAdmin);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "IsDbBackupOperator:", dataBase.IsDbBackupOperator);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "IsDbDatareader:", dataBase.IsDbDatareader);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "IsDbDatawriter:", dataBase.IsDbDatawriter);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "IsDbDdlAdmin:", dataBase.IsDbDdlAdmin);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "IsDbDenyDatareader:", dataBase.IsDbDenyDatareader);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "IsDbDenyDatawriter:", dataBase.IsDbDenyDatawriter);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "IsDbOwner:", dataBase.IsDbOwner);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "IsDbSecurityAdmin:", dataBase.IsDbSecurityAdmin);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "IsFullTextEnabled:", dataBase.IsFullTextEnabled);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "IsMailHost:", dataBase.IsMailHost);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "IsManagementDataWarehouse:", dataBase.IsManagementDataWarehouse);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "IsMirroringEnabled:", dataBase.IsMirroringEnabled);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "IsParameterizationForced:", dataBase.IsParameterizationForced);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "IsReadCommittedSnapshotOn:", dataBase.IsReadCommittedSnapshotOn);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "IsSystemObject:", dataBase.IsSystemObject);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "IsUpdateable:", dataBase.IsUpdateable);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "IsVarDecimalStorageFormatEnabled:", dataBase.IsVarDecimalStorageFormatEnabled);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "LastBackupDate:", dataBase.LastBackupDate);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "LastDifferentialBackupDate:", dataBase.LastDifferentialBackupDate);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "LastLogBackupDate:", dataBase.LastLogBackupDate);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "LocalCursorsDefault:", dataBase.LocalCursorsDefault);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "LogReuseWaitStatus:", dataBase.LogReuseWaitStatus);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "MirroringFailoverLogSequenceNumber:", dataBase.MirroringFailoverLogSequenceNumber);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "MirroringID:", dataBase.MirroringID);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "MirroringPartner:", dataBase.IsParameterizationForced);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "MirroringRedoQueueMaxSize:", dataBase.MirroringRedoQueueMaxSize);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "MirroringRoleSequence:", dataBase.MirroringRoleSequence);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "MirroringSafetyLevel:", dataBase.MirroringSafetyLevel);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "MirroringSafetySequence:", dataBase.MirroringSafetySequence);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "MirroringStatus:", dataBase.MirroringStatus);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "MirroringTimeout:", dataBase.MirroringTimeout);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "MirroringWitness:", dataBase.MirroringWitness);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "NumericRoundAbortEnabled:", dataBase.NumericRoundAbortEnabled);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "Owner:", dataBase.Owner);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "PageVerify:", dataBase.PageVerify);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "PrimaryFilePath:", dataBase.PrimaryFilePath);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "QuotedIdentifiersEnabled:", dataBase.QuotedIdentifiersEnabled);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "ReadOnly:", dataBase.ReadOnly);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "RecoveryForkGuid:", dataBase.RecoveryForkGuid);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "RecoveryModel:", dataBase.RecoveryModel);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "RecursiveTriggersEnabled:", dataBase.RecursiveTriggersEnabled);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "Size:", dataBase.Size);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "SnapshotIsolationState:", dataBase.SnapshotIsolationState);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "SpaceAvailable:", dataBase.SpaceAvailable);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "State:", dataBase.State);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "Status:", dataBase.Status);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "UserName:", dataBase.UserName);
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "Version:", dataBase.Version);

            rowsAdded++;

            return rowsAdded;
        }

        private int AddSection_Databases(Worksheet ws, SMOH.Server serverInstance, int startingRow)
        {
            ExcelHlp.AddContentToCell(ws.Cells[startingRow++, 1], "Databases", 14, ExcelHlp.MakeBold.Yes);

            int currentRow = startingRow;
            int rowsAdded = 0;
            int col = 1;

            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 25, currentRow, "Name", 12);
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 20, currentRow, "Owner", 12);
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 15, currentRow, "Create\nDate", 12);
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 15, currentRow, "DataSpace\nUsage", 12);
            //ExcelHlp.AddColumnToSheet(ref ws, col++, 5,  row, "ID");
            //ExcelHlp.AddColumnToSheet(ref ws, col++, 35, row, "DatabaseGuid");

            rowsAdded++;

            rowsAdded += DisplayListOf_DataBases(ws, serverInstance, ++currentRow);

            // Create a Table to structure the output
            ws.ListObjects.Add(XlListObjectSourceType.xlSrcRange, ws.Range[ws.Cells[startingRow, 1], ws.Cells[startingRow + rowsAdded - 1, col - 1]], Type.Missing, XlYesNoGuess.xlYes).Name = string.Format("tblDatabases_{0}", ws.Name);
            string groupRange = string.Format("{0}:{1}", startingRow, startingRow + rowsAdded - 1);
            ws.Rows[groupRange].Group();
            ws.Rows[groupRange].Hidden = true;

            return rowsAdded;
        }

        private int AddSection_DataFileInfo(Worksheet ws, SMOH.FileGroup fileGroup, int startingRow)
        {
            int rowsAdded = 0;
            int currentRow = startingRow;

            ExcelHlp.AddContentToCell(ws.Cells[currentRow++, 1], "DataFiles", 14, ExcelHlp.MakeBold.Yes);
            rowsAdded++;

            int col = 1;

            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 45, currentRow, "FileName");
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 45, currentRow, "AvailableSpace");
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 45, currentRow, "Name");
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 45, currentRow, "Growth");
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 45, currentRow, "GrowthType");
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 10, currentRow, "MaxSize");
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 13, currentRow, "NumberOf\nDiskReads");
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 13, currentRow, "NumberOf\nDiskWrites");
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 10, currentRow, "Size");
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 10, currentRow, "State");
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 11, currentRow, "UsedSpace");
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 10, currentRow, "Volume\nFreeSpace");

            rowsAdded++;

            rowsAdded += DisplayListOf_DataFiles(ws, fileGroup, ++currentRow);

            //// Create a Table to structure the output
            //ws.ListObjects.Add(XlListObjectSourceType.xlSrcRange, ws.Range[ws.Cells[startingRow, 1], ws.Cells[startingRow + rowsAdded - 1, col - 1]], Type.Missing, XlYesNoGuess.xlYes).Name = string.Format("tbl_{0}", ws.Name);
            string groupRange = string.Format("{0}:{1}", startingRow + 1, startingRow + rowsAdded - 1);
            ws.Rows[groupRange].Group();
            ws.Rows[groupRange].Hidden = true;

            return rowsAdded;
        }

        private int AddSection_EndPoints(Worksheet ws, SMOH.Server server, int startingRow)
        {
            int rowsAdded = 0;
            int currentRow = startingRow;

            ExcelHlp.AddContentToCell(ws.Cells[currentRow++, 1], "Endpoints", 14, ExcelHlp.MakeBold.Yes);
            rowsAdded++;

            int col = 1;

            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 45, currentRow, "EndpointState");
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 45, currentRow, "EndpointType");
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 45, currentRow, "ID");
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 45, currentRow, "IsAdminEndpoint");
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 45, currentRow, "IsSystemObject");
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 10, currentRow, "Name");
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 13, currentRow, "Owner");
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 13, currentRow, "Protocol");
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 10, currentRow, "ProtocolType");
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 10, currentRow, "Urn");

            rowsAdded++;

            rowsAdded += DisplayListOf_Endpoints(ws, server, ++currentRow);

            //// Create a Table to structure the output
            //ws.ListObjects.Add(XlListObjectSourceType.xlSrcRange, ws.Range[ws.Cells[startingRow, 1], ws.Cells[startingRow + rowsAdded - 1, col - 1]], Type.Missing, XlYesNoGuess.xlYes).Name = string.Format("tbl_{0}", ws.Name);
            string groupRange = string.Format("{0}:{1}", startingRow + 1, startingRow + rowsAdded - 1);
            ws.Rows[groupRange].Group();
            ws.Rows[groupRange].Hidden = true;

            return rowsAdded;
        }
        private int AddSection_ExtendedPropertyInfo(Worksheet ws, SMO.ExtendedPropertyCollection extendedProperties, int startingRow)
        {
            int rowsAdded = 0;
            int currentRow = startingRow;

            AddTitledInfo(ws.Cells[currentRow++, 1], "ExtendedProperties:", extendedProperties.Count.ToString());

            rowsAdded++;

            int col = 1;

            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 45, currentRow, "Name");
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 10, currentRow, "Value");

            rowsAdded++;

            rowsAdded += DisplayListOf_ExtendedProperties(ws, extendedProperties, ++currentRow);

            // Create a Table to structure the output
            //ws.ListObjects.Add(XlListObjectSourceType.xlSrcRange, ws.Range[ws.Cells[startingRow, 1], ws.Cells[startingRow + rowsAdded - 1, col - 1]], Type.Missing, XlYesNoGuess.xlYes).Name = string.Format("tblFileGroups_{0}", ws.Name);
            string groupRange = string.Format("{0}:{1}", startingRow + 1, startingRow + rowsAdded - 1);
            ws.Rows[groupRange].Group();
            ws.Rows[groupRange].Hidden = true;

            return rowsAdded;
        }

        private int AddSection_FileGroupInfo(Worksheet ws, SMOH.Database dataBase,int startingRow)
        {
            int rowsAdded = 0;
            int currentRow = startingRow;

            AddTitledInfo(ws.Cells[currentRow++, 1], "FileGroups:", dataBase.FileGroups.Count.ToString());

            //ExcelHlp.AddContentToCell(ws.Cells[currentRow++, 1], "FileGroups", 14, ExcelHlp.MakeBold.Yes);
            rowsAdded++;

            int col = 1;

            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 45, currentRow, "Name");
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 10, currentRow, "IsDefault");

            rowsAdded++;

            rowsAdded += DisplayListOf_FileGroups(ws, dataBase, ++currentRow);

            // Create a Table to structure the output
            //ws.ListObjects.Add(XlListObjectSourceType.xlSrcRange, ws.Range[ws.Cells[startingRow, 1], ws.Cells[startingRow + rowsAdded - 1, col - 1]], Type.Missing, XlYesNoGuess.xlYes).Name = string.Format("tblFileGroups_{0}", ws.Name);
            string groupRange = string.Format("{0}:{1}", startingRow + 1, startingRow + rowsAdded - 1);
            ws.Rows[groupRange].Group();
            ws.Rows[groupRange].Hidden = true;

            return rowsAdded;
        }

        private int AddSection_InstanceInfo(Worksheet ws, SMOH.Server serverInstance, int startingRow, bool showDetail)
        {
            Range rng = ws.Cells[startingRow, 1];

            int rowsAdded = 0;

            AddTitledInfo(rng.Offset[rowsAdded++, 0], "As of:", DateTime.Now.ToString());
            AddTitledInfo(rng.Offset[rowsAdded++, 0], "Instance Name:", serverInstance.Name);

            if (showDetail)
            {
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "Backup\nDirectory:", serverInstance.BackupDirectory);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "Browser\nService Account:", serverInstance.BrowserServiceAccount);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "Browser\nStartMode:", serverInstance.BrowserStartMode);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "Build\nClrVersion:", serverInstance.BuildClrVersionString);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "BuildNumber:", serverInstance.BuildNumber);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "Collation:", serverInstance.Collation);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "ComparisonStyle:", serverInstance.ComparisonStyle);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "ComputerNamePhysicalNetBIOS:", serverInstance.ComputerNamePhysicalNetBIOS);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "DefaultFile:", serverInstance.DefaultFile);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "DefaultLog:", serverInstance.DefaultLog);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "Edition:", serverInstance.Edition);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "ErrorLogPath:", serverInstance.ErrorLogPath);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "FilestreamLevel:", serverInstance.FilestreamLevel);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "FilestreamShareName:", serverInstance.FilestreamShareName);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "InstallDataDirectory:", serverInstance.InstallDataDirectory);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "InstallSharedDirectory:", serverInstance.InstallSharedDirectory);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "InstanceName:", serverInstance.InstanceName);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "IsCaseSensitive:", serverInstance.IsCaseSensitive);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "IsClustered:", serverInstance.IsClustered);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "IsFullTextInstalled:", serverInstance.IsFullTextInstalled);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "IsSingleUser:", serverInstance.IsSingleUser);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "LoginMode:", serverInstance.LoginMode);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "MasterDBPath:", serverInstance.MasterDBPath);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "MasterDBLogPath:", serverInstance.MasterDBLogPath);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "MaxPrecision:", serverInstance.MaxPrecision);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "NamedPipesEnabled:", serverInstance.NamedPipesEnabled);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "NetName:", serverInstance.NetName);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "NumberOfLogFiles:", serverInstance.NumberOfLogFiles);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "OSVersion:", serverInstance.OSVersion);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "PerfMonMode:", serverInstance.PerfMonMode);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "PhysicalMemory:", serverInstance.PhysicalMemory);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "PhysicalMemoryUsageInKB:", serverInstance.PhysicalMemoryUsageInKB);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "Platform:", serverInstance.Platform);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "Processors:", serverInstance.Processors);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "ProcessorUsage:", serverInstance.ProcessorUsage);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "Product:", serverInstance.Product);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "ProductLevel:", serverInstance.ProductLevel);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "ResourceVersion:", serverInstance.ResourceVersionString);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "Root\nDirectory:", serverInstance.RootDirectory);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "ServerType:", serverInstance.ServerType);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "Service\nAccount:", serverInstance.ServiceAccount);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "ServiceInstanceId:", serverInstance.ServiceInstanceId);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "ServiceName:", serverInstance.ServiceName);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "Service\nStartMode:", serverInstance.ServiceStartMode);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "SqlCharSet\nName:", serverInstance.SqlCharSetName);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "SqlDomain\nGroup:", serverInstance.SqlDomainGroup);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "SqlSortOrder\nName:", serverInstance.SqlSortOrderName);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "Status:", serverInstance.Status);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "TcpEnabled:", serverInstance.TcpEnabled);
                AddTitledInfo(rng.Offset[rowsAdded++, 0], "Version:", serverInstance.VersionString);
            }

            return rowsAdded;
        }

        private int AddSection_LinkedServers(Worksheet ws, SMOHelper.Server serverInstance, int startingRow)
        {
            ExcelHlp.AddContentToCell(ws.Cells[startingRow++, 1], "LinkedServers", 14, ExcelHlp.MakeBold.Yes);

            int currentRow = startingRow;
            int rowsAdded = 0;
            int col = 1;

            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 25, currentRow, "Name");
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 25, currentRow, "Catalog");

            rowsAdded++;

            rowsAdded += DisplayListOf_LinkedServers(ws, serverInstance, ++currentRow);

            // Create a Table to structure the output
            ws.ListObjects.Add(XlListObjectSourceType.xlSrcRange, ws.Range[ws.Cells[startingRow, 1], ws.Cells[startingRow + rowsAdded - 1, col - 1]], Type.Missing, XlYesNoGuess.xlYes).Name = string.Format("tblLinkedServers_{0}", ws.Name);
            string groupRange = string.Format("{0}:{1}", startingRow, startingRow + rowsAdded - 1);
            ws.Rows[groupRange].Group();
            ws.Rows[groupRange].Hidden = true;

            return rowsAdded;
        }

        private int AddSection_Logins(Worksheet ws, SMOH.Server serverInstance, int startingRow)
        {
            ExcelHlp.AddContentToCell(ws.Cells[startingRow++, 1], "Logins", 14, ExcelHlp.MakeBold.Yes);

            int currentRow = startingRow;
            int rowsAdded = 0;
            int col = 1;

            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 25, currentRow, "Name");

            rowsAdded++;

            rowsAdded += DisplayListOf_Logins(ws, serverInstance, ++currentRow);

            // Create a Table to structure the output
            ws.ListObjects.Add(XlListObjectSourceType.xlSrcRange, ws.Range[ws.Cells[startingRow, 1], ws.Cells[startingRow + rowsAdded - 1, col - 1]], Type.Missing, XlYesNoGuess.xlYes).Name = string.Format("tblLogins_{0}", ws.Name);
            string groupRange = string.Format("{0}:{1}", startingRow, startingRow + rowsAdded - 1);
            ws.Rows[groupRange].Group();
            ws.Rows[groupRange].Hidden = true;

            return rowsAdded;
        }

        private int AddSection_ServerRoles(Worksheet ws, SMOH.Server serverInstance, int startingRow)
        {
            ExcelHlp.AddContentToCell(ws.Cells[startingRow++, 1], "ServerRoles", 14, ExcelHlp.MakeBold.Yes);

            int currentRow = startingRow;
            int rowsAdded = 0;
            int col = 1;

            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 25, currentRow, "Name");

            rowsAdded++;

            rowsAdded += DisplayListOf_ServerRoles(ws, serverInstance, ++currentRow);

            // Create a Table to structure the output
            ws.ListObjects.Add(XlListObjectSourceType.xlSrcRange, ws.Range[ws.Cells[startingRow, 1], ws.Cells[startingRow + rowsAdded - 1, col - 1]], Type.Missing, XlYesNoGuess.xlYes).Name = string.Format("tblServerRoles_{0}", ws.Name);
            string groupRange = string.Format("{0}:{1}", startingRow, startingRow + rowsAdded - 1);
            ws.Rows[groupRange].Group();
            ws.Rows[groupRange].Hidden = true;

            return rowsAdded;
        }

        private int AddSection_StoredProcedure(Worksheet ws, SMOHelper.Database dataBase, int startingRow)
        {
            AddTitledInfo(ws.Cells[startingRow++, 1], "StoredProcedures:", dataBase.StoredProcedures.Count.ToString());
            //ExcelHlp.AddContentToCell(ws.Cells[startingRow++, 1], "StoredProcedures", 14, ExcelHlp.MakeBold.Yes);

            int currentRow = startingRow;
            int rowsAdded = 0;
            int col = 1;

            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 45, currentRow, "Name");
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 10, currentRow, "Owner");
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 15, currentRow, "Create\nDate");
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 13, currentRow, "Date\nLastModified");
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 11, currentRow, "ID");
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 12, currentRow, "MethodName");

            rowsAdded++;

            rowsAdded += DisplayListOf_StoredProcedures(ws, dataBase, ++currentRow);

            // Create a Table to structure the output
            ws.ListObjects.Add(XlListObjectSourceType.xlSrcRange, ws.Range[ws.Cells[startingRow, 1], ws.Cells[startingRow + rowsAdded - 1, col - 1]], Type.Missing, XlYesNoGuess.xlYes).Name = string.Format("tblStoredProcedures_{0}", ws.Name);
            string groupRange = string.Format("{0}:{1}", startingRow, startingRow + rowsAdded - 1);
            ws.Rows[groupRange].Group();
            ws.Rows[groupRange].Hidden = true;

            return rowsAdded;
        }

        private int AddSection_TableInfo(Worksheet ws, SMOH.Database dataBase, int startingRow)
        {
            AddTitledInfo(ws.Cells[startingRow++, 1], "Tables:", dataBase.Tables.Count.ToString());

            //ExcelHlp.AddContentToCell(ws.Cells[startingRow++, 1], "Tables", 14, ExcelHlp.MakeBold.Yes);

            int currentRow = startingRow;
            int rowsAdded = 0;
            int col = 1;

            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 45, currentRow, "Name");
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 10, currentRow, "Owner");
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 15, currentRow, "CreateDate");
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 13, currentRow, "Date\nLastModified");
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 11, currentRow, "DataSpace\nUsed");;
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 11, currentRow, "ID");
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 11, currentRow, "RowCount");

            rowsAdded++;

            rowsAdded += DisplayListOf_Tables(ws, dataBase, ++currentRow);

            // Create a Table to structure the output
            ws.ListObjects.Add(XlListObjectSourceType.xlSrcRange, ws.Range[ws.Cells[startingRow, 1], ws.Cells[startingRow + rowsAdded - 1, col - 1]], Type.Missing, XlYesNoGuess.xlYes).Name = string.Format("tblTables_{0}", ws.Name);
            string groupRange = string.Format("{0}:{1}", startingRow, startingRow + rowsAdded - 1);
            ws.Rows[groupRange].Group();
            ws.Rows[groupRange].Hidden = true;

            return rowsAdded;
        }

        private int AddSection_ViewInfo(Worksheet ws, SMOH.Database dataBase, int startingRow)
        {
            AddTitledInfo(ws.Cells[startingRow++, 1], "Views:", dataBase.Views.Count.ToString());
            //ExcelHlp.AddContentToCell(ws.Cells[startingRow++, 1], "Views", 14, ExcelHlp.MakeBold.Yes);

            int currentRow = startingRow;
            int rowsAdded = 0;
            int col = 1;

            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 45, currentRow, "Name");
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 10, currentRow, "Owner");
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 15, currentRow, "CreateDate");
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 13, currentRow, "Date\nLastModified");
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 10, currentRow, "ID");

            rowsAdded++;

            rowsAdded += DisplayListOf_Views(ws, dataBase, ++currentRow);

            // Create a Table to structure the output
            ws.ListObjects.Add(XlListObjectSourceType.xlSrcRange, ws.Range[ws.Cells[startingRow, 1], ws.Cells[startingRow + rowsAdded - 1, col - 1]], Type.Missing, XlYesNoGuess.xlYes).Name = string.Format("tblViews_{0}", ws.Name);
            string groupRange = string.Format("{0}:{1}", startingRow, startingRow + rowsAdded - 1);
            ws.Rows[groupRange].Group();
            ws.Rows[groupRange].Hidden = true;

            return rowsAdded;
        }

        private void AddTitledInfo(Range outputStartingRange, string title, string info)
        {
            ExcelHlp.AddContentToCell(outputStartingRange.Offset[0, 0], title, 14, ExcelHlp.MakeBold.Yes);
            ExcelHlp.AddContentToCell(outputStartingRange.Offset[0, 1], info, 14);
        }

        private void CreateAllWorksheetsOf_DatabaseInfo(SMOH.Server serverInstance)
        {
            foreach (SMOH.Database database in serverInstance.Databases.Values)
            {
                try
                {
                    CreateWorksheet_DatabaseInfo(database);
                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }
        
        private void CreateAllWorksheeetsOf_StoredProcedureInfo(SMOHelper.Database dataBase)
        {
            foreach(SMOH.StoredProcedure storedProcedure in dataBase.StoredProcedures.Values)
            {
                try
                {
                    CreateWorksheet_StoredProcedureInfo(storedProcedure);
                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }

        private void CreateAllWorkSheetsOf_TableInfo(SMOH.Database database)
        {
            foreach (SMOH.Table table in database.Tables.Values)
            {
                try
                {
                    CreateWorksheet_TableInfo(table);
                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }

        private void CreateAllWorksheetsOf_ViewInfo(SMOHelper.Database dataBase)
        {
            foreach(SMOH.View view in dataBase.Views.Values)
            {
                try
                {
                    CreateWorksheet_ViewInfo(view);
                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }   

        private void CreateDatabase()
        {
            if (_SMOServer.Databases[txtNewDatabase.Text] != null)
            {
            	if (DialogResult.OK != MessageBox.Show("Drop Database?", "Database Exists Warning", MessageBoxButtons.OKCancel))
                {
                    return;
                }
                	
                _SMOServer.Databases[txtNewDatabase.Text].Drop();
            }

            SMO.Database newDatabase = new SMO.Database(_SMOServer, txtNewDatabase.Text);
            newDatabase.Create();
            // TODO: Need to update SMOH object list and comboboxes
        }

        private void CreateTable()
        {
            SMO.Database database = _SMOServer.Databases[cbDatabases.Text];

            SMO.Table newTable = new SMO.Table(database, txtNewTable.Text);

            //SMO.DataType varChar = new SMO.DataType(SMO.SqlDataType.VarChar, 100);

            SMO.Column col1 = new SMO.Column(newTable, "col1", SMO.DataType.VarChar(100));
            
            SMO.Column col2 = new SMO.Column(newTable, "col2", SMO.DataType.DateTime);
            SMO.Column col3 = new SMO.Column(newTable, "col3", SMO.DataType.Int);

            newTable.Columns.Add(col1);
            newTable.Columns.Add(col2);
            newTable.Columns.Add(col3);
            newTable.Create();

            //database.Tables.Add(newTable);

        }

        private void CreateWorksheet_DatabaseInfo(SMOH.Database dataBase)
        {
            long startTicks = Common.WriteToDebugWindow("CreateDatabaseInfoWorkSheet(Start)");

            string sheetName = ExcelHlp.SafeSheetName("D>" + dataBase.Name);
            Worksheet ws = ExcelHlp.NewWorksheet(sheetName, afterSheetName: "LAST");

            // Output starts here.  Each Display method returns the output end point.

            int startingRow = 2;

            startingRow += AddSection_DatabaseInfo(ws, dataBase, startingRow);

            startingRow += AddSection_ExtendedPropertyInfo(ws, dataBase.ExtendedProperties, startingRow);

            startingRow += AddSection_FileGroupInfo(ws, dataBase, startingRow);

            if (ckIncludeDBTables.Checked)
            {
                startingRow += 2;
                startingRow += AddSection_TableInfo(ws, dataBase, startingRow);
            }

            if (ckIncludeDBViews.Checked)
            {
                startingRow += 2;
                startingRow += AddSection_ViewInfo(ws, dataBase, startingRow);
            }

            if (ckIncludeDBStoredProcedures.Checked)
            {
                startingRow += 2;
                startingRow += AddSection_StoredProcedure(ws, dataBase, startingRow);
            }

            Common.WriteToDebugWindow("CreateDatabaseInfoWorkSheet(Start)", startTicks);
        }

        private void CreateWorksheet_InstanceInfo(SMOH.Server serverInstance, bool showDetails)
        {
            long startTicks = Common.WriteToDebugWindow("CreateInstanceInfoWorksheet(Start)");

            string sheetName = ExcelHlp.SafeSheetName("I>" + serverInstance.Name);
            Worksheet ws = ExcelHlp.NewWorksheet(sheetName, beforeSheetName: "FIRST");

            // Output starts here.  Each Display method returns the output end point.

            int startingRow = 2;

            startingRow += AddSection_InstanceInfo(ws, serverInstance, startingRow, showDetails);
            startingRow++;
            startingRow += AddSection_Databases(ws, serverInstance, startingRow);
            startingRow += 2;
            startingRow += +AddSection_Logins(ws, serverInstance, startingRow);
            startingRow += 2;
            startingRow += AddSection_ServerRoles(ws, serverInstance, startingRow);
            startingRow += 2;
            startingRow += AddSection_LinkedServers(ws, serverInstance, startingRow);
            startingRow += 2;
            startingRow += AddSection_EndPoints(ws, serverInstance, startingRow);
            startingRow += 2;

            Common.WriteToDebugWindow("CreateInstanceInfoWorksheet(End)", startTicks);
        }

        private void CreateWorksheet_StoredProcedureInfo(SMOH.StoredProcedure storedProcedure)
        {
            long startTicks = Common.WriteToDebugWindow("CreateStoredProcedureInfoWorksheet(Start)");

            string sheetName = ExcelHlp.SafeSheetName("S>" + storedProcedure.Name);
            Worksheet ws = ExcelHlp.NewWorksheet(sheetName, afterSheetName: "LAST");
            int fontSize = 8;

            AddTitledInfo(ws.Cells[2, 1], "As of:", DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss"));
            AddTitledInfo(ws.Cells[3, 1], "DB Name:", storedProcedure.Name);


            ExcelHlp.AddContentToCell(ws.Cells[6, 1], "Parameters", 14, ExcelHlp.MakeBold.Yes);
            int row = 7;
            int col = 1;

            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 25, row, "Name", 12);
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 20, row, "DataType", 12);
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 10, row, "Maximum\nLength", 12);
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 10, row, "Numeric\nPrecision", 12);
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 10, row, "Numeric\nScale", 12);
            ExcelHlp.AddColumnHeaderToSheet(ref ws, col++, 10, row, "Default\nValue", 12);

            row += DisplayListOf_StoredProcedureParameters(ws, storedProcedure, ++row);

            row++;

            AddTitledInfo(ws.Cells[row++, 1], "Header:", "");
            ExcelHlp.AddContentToCell(ws.Cells[row++, 2], storedProcedure.TextHeader, fontSize);
            AddTitledInfo(ws.Cells[row++, 1], "Body:", "");
            ExcelHlp.AddContentToCell(ws.Cells[row++, 2], storedProcedure.TextBody, fontSize);

            Common.WriteToDebugWindow("CreateStoredProcedureInfoWorksheet(End)", startTicks);
        }

        private void CreateWorksheet_TableInfo(SMOH.Table table)
        {
            long startTicks = Common.WriteToDebugWindow("CreateTableInfoWorkSheet(Start)");

            string sheetName = ExcelHlp.SafeSheetName("T>" + table.Name);
            Worksheet ws = ExcelHlp.NewWorksheet(sheetName, afterSheetName: "LAST");

            AddTitledInfo(ws.Cells[2, 1], "As of:", DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss"));
            AddTitledInfo(ws.Cells[3, 1], "Name:", table.Name);

            int startingRow = 6;
            startingRow += AddSection_ExtendedPropertyInfo(ws, table.ExtendedProperties, startingRow);

            AddSection_ColumnInfo(ws, table,startingRow);

            Common.WriteToDebugWindow("CreateTableInfoWorkSheet(End)", startTicks);
        }

        private void CreateWorksheet_ViewInfo(SMOH.View view)
        {
            long startTicks = Common.WriteToDebugWindow("CreateViewInfoWorkSheet(Start)");

            string sheetName = ExcelHlp.SafeSheetName("V>" + view.Name);
            Worksheet ws = ExcelHlp.NewWorksheet(sheetName, afterSheetName: "LAST");

            AddTitledInfo(ws.Cells[2, 1], "As of:", DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss"));
            AddTitledInfo(ws.Cells[3, 1], "Name:", view.Name);

            int startingRow = 6;

            startingRow += AddSection_ExtendedPropertyInfo(ws, view.ExtendedProperties, startingRow);
            AddSection_ColumnInfo(ws, view, startingRow);

            Common.WriteToDebugWindow("CreateViewInfoWorkSheet(End)", startTicks);
        }

        private int DisplayListOf_Columns(Worksheet ws, SMOH.Table table, int startingRow)
        {
            Range rng = ws.Cells[startingRow, 1];
            int rowsAdded = 0;

            foreach (SMOH.Column column in table.Columns.Values)
            {
                int col = 0;

                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], column.Name);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], column.DataType);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], column.MaximumLength);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], column.NumericPrecision);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], column.NumericScale);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], column.Default);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], column.ID);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], column.Identity);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], column.InPrimaryKey);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], column.IsForeignKey);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], column.Nullable);

                rowsAdded++;
            }

            return rowsAdded;
        }

        private int DisplayListOf_Columns(Worksheet ws, SMOH.View view, int startingRow)
        {
            Range rng = ws.Cells[startingRow, 1];
            int rowsAdded = 0;

            foreach (SMOH.Column column in view.Columns.Values)
            {
                int col = 0;

                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], column.Name);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], column.DataType);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], column.MaximumLength);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], column.NumericPrecision);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], column.NumericScale);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], column.Default);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], column.ID);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], column.Identity);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], column.InPrimaryKey);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], column.IsForeignKey);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], column.Nullable);

                rowsAdded++;
            }

            return rowsAdded;
        }

        private int DisplayListOf_DataBases(Worksheet ws, SMOH.Server serverInstance, int startingRow)
        {
            // The columns in this method need to be kept in sync with CreateInstanceInfoWorksheet()

            Range rng = ws.Cells[startingRow, 1];
            int rowsAdded = 0;

            foreach (SMOH.Database dataBase in serverInstance.Databases.Values)
            {
                int col = 0;

                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], dataBase.Name);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], dataBase.Owner);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], dataBase.CreateDate);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], dataBase.DataSpaceUsage);

                //ExcelHlp.AddContentToCell(rng.Offset[row, col++], dataBase.ID.ToString());
                //ExcelHlp.AddContentToCell(rng.Offset[row, col++], dataBase.DatabaseGuid.ToString());

                rowsAdded++;
            }

            return rowsAdded;
        }

        private int DisplayListOf_DataFiles(Worksheet ws, SMOH.FileGroup fileGroup, int startingRow)
        {
            // The columns in this method need to be kept in sync with CreateInstanceInfoWorksheet()

            Range rng = ws.Cells[startingRow, 1];
            int rowsAdded = 0;

            foreach (SMOH.DataFile dataFile in fileGroup.DataFiles.Values)
            {
                int col = 0;

                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], dataFile.FileName);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], dataFile.AvailableSpace);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], dataFile.Name);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], dataFile.Growth);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], dataFile.GrowthType);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], dataFile.MaxSize);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], dataFile.NumberOfDiskReads);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], dataFile.NumberOfDiskWrites);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], dataFile.Size);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], dataFile.State);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], dataFile.UsedSpace);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], dataFile.VolumeFreeSpace);

                rowsAdded++;
            }

            return rowsAdded;
        }

        private int DisplayListOf_Endpoints(Worksheet ws, SMOH.Server server, int startingRow)
        {
            // The columns in this method need to be kept in sync with CreateInstanceInfoWorksheet()

            Range rng = ws.Cells[startingRow, 1];
            int rowsAdded = 0;

            foreach (SMOH.Endpoint endPoint in server.Endpoints.Values)
            {
                int col = 0;

                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], endPoint.EndpointState);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], endPoint.EndpointType);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], endPoint.ID);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], endPoint.IsAdminEndpoint);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], endPoint.IsSystemObject);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], endPoint.Name);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], endPoint.Owner);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], endPoint.Protocol);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], endPoint.ProtocolType);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], endPoint.Urn);

                rowsAdded++;
            }

            return rowsAdded;
        }

        private int DisplayListOf_ExtendedProperties(Worksheet ws, SMO.ExtendedPropertyCollection extendedProperties, int startingRow)
        {
            // The columns in this method need to be kept in sync with CreateInstanceInfoWorksheet()

            Range rng = ws.Cells[startingRow, 1];
            int rowsAdded = 0;

            foreach (SMO.ExtendedProperty extendedProperty in extendedProperties)
            {
                //int row = 0;
                int col = 0;

                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], extendedProperty.Name);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], extendedProperty.Value.ToString());

                rowsAdded++;
            }

            return rowsAdded;
        }

        private int DisplayListOf_FileGroups(Worksheet ws, SMOH.Database dataBase, int startingRow)
        {
            // The columns in this method need to be kept in sync with CreateInstanceInfoWorksheet()

            Range rng = ws.Cells[startingRow, 1];
            int rowsAdded = 0;

            foreach (SMOH.FileGroup fileGroup in dataBase.FileGroups.Values)
            {
                //int row = 0;
                int col = 0;

                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], fileGroup.Name);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], fileGroup.IsDefault);
                
                rowsAdded++;

                rowsAdded += AddSection_DataFileInfo(ws, fileGroup, startingRow + rowsAdded);
                rng = ws.Cells[startingRow + rowsAdded, 1];
            }

            return rowsAdded;
        }

        private int DisplayListOf_LinkedServers(Worksheet ws, SMOH.Server serverInstance, int startingRow)
        {
            // The columns in this method need to be kept in sync with CreateInstanceInfoWorksheet()

            Range rng = ws.Cells[startingRow, 1];
            int rowsAdded = 0;

            foreach (SMOH.LinkedServer linkedServer in serverInstance.LinkedServers.Values)
            {
                int col = 0;

                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], linkedServer.Name);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], linkedServer.Catalog);

                rowsAdded++;
            }

            return rowsAdded;
        }

        private int DisplayListOf_Logins(Worksheet ws, SMOH.Server serverInstance, int startingRow)
        {
            // The columns in this method need to be kept in sync with CreateInstanceInfoWorksheet()

            Range rng = ws.Cells[startingRow, 1];
            int rowsAdded = 0;

            foreach (SMOH.Login login in serverInstance.Logins.Values)
            {
                int col = 0;

                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], login.Name);

                rowsAdded++;
            }

            return rowsAdded;
        }

        private int DisplayListOf_ServerRoles(Worksheet ws, SMOH.Server serverInstance, int startingRow)
        {
            // The columns in this method need to be kept in sync with CreateInstanceInfoWorksheet()

            Range rng = ws.Cells[startingRow, 1];
            int rowsAdded = 0;

            foreach (SMOH.ServerRole serverRole in serverInstance.ServerRoles.Values)
            {
                int col = 0;

                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], serverRole.Name);

                rowsAdded++;
            }

            return rowsAdded;
        }

        private int DisplayListOf_StoredProcedureParameters(Worksheet ws, SMOHelper.StoredProcedure storedProcedure, int startingRow)
        {
            Range rng = ws.Cells[startingRow, 1];
            int rowsAdded = 0;

            foreach(SMOH.StoredProcedureParameter parameter in storedProcedure.Parameters.Values)
            {
                int col = 0;

                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], parameter.Name);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], parameter.DataType);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], parameter.MaximumLength);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], parameter.NumericPrecision);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], parameter.NumericScale);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], parameter.DefaultValue);

                rowsAdded++;
            }

            return rowsAdded;
        }

        private int DisplayListOf_StoredProcedures(Worksheet ws, SMOH.Database dataBase, int startingRow)
        {
            Range rng = ws.Cells[startingRow, 1];
            int rowsAdded = 0;

            foreach (SMOH.StoredProcedure storedProcedure in dataBase.StoredProcedures.Values)
            {
                if (storedProcedure.IsSystemObject && ! ckIncludeSystemStoredProcedures.Checked)
                {
                    continue;
                }

                int col = 0;

                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], storedProcedure.Name);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], storedProcedure.Owner);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], storedProcedure.CreateDate);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], storedProcedure.DateLastModified);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], storedProcedure.ID);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], storedProcedure.MethodName);

                rowsAdded++;
            }

            return rowsAdded;
        }

        private int DisplayListOf_Tables(Worksheet ws, SMOH.Database dataBase, int startingRow)
        {
            // The columns in this method need to be kept in sync with CreateDatabaseInfoWorkSheet()

            Range rng = ws.Cells[startingRow, 1];
            int rowsAdded = 0;

            foreach (SMOH.Table table in dataBase.Tables.Values)
            {
                int col = 0;

                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], table.Name);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], table.Owner);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], table.CreateDate);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], table.DateLastModified);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], table.DataSpaceUsed);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], table.ID.ToString());
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], table.RowCount.ToString());

                rowsAdded++;
            }

            return rowsAdded;
        }

        private int DisplayListOf_Views(Worksheet ws, SMOH.Database dataBase, int startingRow)
        {
            // The columns in this method need to be kept in sync with CreateDatabaseInfoWorkSheet()

            Range rng = ws.Cells[startingRow, 1];
            int rowsAdded = 0;

            foreach(SMOH.View view in dataBase.Views.Values)
            {
                int col = 0;

                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], view.Name);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], view.Owner);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], view.CreateDate);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], view.DateLastModified);
                ExcelHlp.AddContentToCell(rng.Offset[rowsAdded, col++], view.ID.ToString());

                rowsAdded++;
            }

            return rowsAdded;
        }

        private void EnumAvailableSqlServers()
        {
            SMO.SmoApplication smoApplication = new SMO.SmoApplication();

            foreach (var sqlServer in SMO.SmoApplication.EnumAvailableSqlServers().AsEnumerable())
            {
                sqlServer.Field<string>("Name");
                sqlServer.Field<string>("Server");
                sqlServer.Field<string>("Instance");
                sqlServer.Field<bool>("IsClustered");
                sqlServer.Field<string>("Version");
                sqlServer.Field<bool>("IsLocal");
            }
        }

        private void GetServerInfo()
        {
            SMOWMI.ManagedComputer computer = new SMOWMI.ManagedComputer(txtServerName.Text);
            Common.WriteToDebugWindow(String.Format("Name:{0}", computer.Name));


        }

        private void ListRows(Microsoft.SqlServer.Management.Smo.View view)
        {
            throw new NotImplementedException();
        }

        private void ListRows(Microsoft.SqlServer.Management.Smo.Table table)
        {
            throw new NotImplementedException();
        }

        private void Logoff()
        {
            if (_SMOServer != null)
            {
            	_SMOServer.ConnectionContext.Disconnect();

                // May want to keep these around
                _SMOHServer = null;
            }

            cbDatabases.Items.Clear();
        }

        private bool Logon()
        {
            bool result = false;

            _SMOServer = new SMO.Server(ucDBInstanceList.InstanceName);

            if (ckIntegratedSecurity.Checked)
            {
                _SMOServer.ConnectionContext.LoginSecure = true;
            }
            else
            {
                _SMOServer.ConnectionContext.LoginSecure = false;
                //_SMOServer.ConnectionContext.Login = txtUserName.Text;
                _SMOServer.ConnectionContext.Login = cbUserName.Text;
                _SMOServer.ConnectionContext.Password = txtPassword.Text;
            }

            try
            {
                // Don't have to explicitly connect.  Connection is established when needed and pooling is used.
                // If connect, no pooling.

                //_SMOServer.ConnectionContext.Connect();

                // Load the SMOHelper Server object with information.
                // This allows us to not worry about access privileges in this code.
                // The values will be available or marked as "<No Access>"
                _SMOHServer = new SMOH.Server(_SMOServer);

                cbDatabases.Items.Clear();

                foreach (string name in _SMOHServer.Databases.Keys)
                {
                    cbDatabases.Items.Add(name);
                }

                result = true;
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            return result;
        }

        #endregion

    }
}
