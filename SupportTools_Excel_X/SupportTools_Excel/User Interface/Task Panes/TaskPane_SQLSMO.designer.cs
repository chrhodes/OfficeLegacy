namespace SupportTools_Excel.User_Interface.Task_Panes
{
    partial class TaskPane_SQLSMO
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if(disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.gbConnect = new System.Windows.Forms.GroupBox();
            this.btnEnumAvailableSqlServers = new System.Windows.Forms.Button();
            this.cbUserName = new System.Windows.Forms.ComboBox();
            this.lblInstancName = new System.Windows.Forms.Label();
            this.gbInstanceOperations = new System.Windows.Forms.GroupBox();
            this.chkListInstanceDetails = new System.Windows.Forms.CheckBox();
            this.btnCreateDatabase = new System.Windows.Forms.Button();
            this.txtNewDatabase = new System.Windows.Forms.TextBox();
            this.gbDatabaseOperations = new System.Windows.Forms.GroupBox();
            this.ckIncludeDBStoredProcedures = new System.Windows.Forms.CheckBox();
            this.txtNewTable = new System.Windows.Forms.TextBox();
            this.ckIncludeDBViews = new System.Windows.Forms.CheckBox();
            this.ckIncludeDBTables = new System.Windows.Forms.CheckBox();
            this.btnCreateDatabaseInfoWorksheets = new System.Windows.Forms.Button();
            this.btnCreateTable = new System.Windows.Forms.Button();
            this.gbStoredProceduresOperations = new System.Windows.Forms.GroupBox();
            this.btnCreateStoredProcedureInfoWorksheet = new System.Windows.Forms.Button();
            this.cbStoredProcedures = new System.Windows.Forms.ComboBox();
            this.btnCreateStoredProcedureInfoWorksheets = new System.Windows.Forms.Button();
            this.ckIncludeSystemStoredProcedures = new System.Windows.Forms.CheckBox();
            this.cbDatabases = new System.Windows.Forms.ComboBox();
            this.gbViewOperations = new System.Windows.Forms.GroupBox();
            this.btnListRows_View = new System.Windows.Forms.Button();
            this.ckIncludeSystemViews = new System.Windows.Forms.CheckBox();
            this.btnCreateViewInfoWorkSheet = new System.Windows.Forms.Button();
            this.btnCreateViewInfoWorksheets = new System.Windows.Forms.Button();
            this.cbViews = new System.Windows.Forms.ComboBox();
            this.gbTableOperations = new System.Windows.Forms.GroupBox();
            this.btnListRows_Table = new System.Windows.Forms.Button();
            this.btnCreateTableInfoWorksheet = new System.Windows.Forms.Button();
            this.cbTables = new System.Windows.Forms.ComboBox();
            this.btnCreateTableInfoWorkSheets = new System.Windows.Forms.Button();
            this.btnCreateDatabaseInfoWorkSheet = new System.Windows.Forms.Button();
            this.btnCreateInstanceInfoWorkSheet = new System.Windows.Forms.Button();
            this.btnLogoff = new System.Windows.Forms.Button();
            this.btnLogon = new System.Windows.Forms.Button();
            this.lblPassword = new System.Windows.Forms.Label();
            this.lblUserName = new System.Windows.Forms.Label();
            this.txtPassword = new System.Windows.Forms.TextBox();
            this.ckIntegratedSecurity = new System.Windows.Forms.CheckBox();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.gbServerOperations = new System.Windows.Forms.GroupBox();
            this.txtServerName = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.ucDBInstanceList = new SupportTools_Excel.User_Interface.User_Controls.ucDBInstanceList();
            this.btnGetServerInfo = new System.Windows.Forms.Button();
            this.gbConnect.SuspendLayout();
            this.gbInstanceOperations.SuspendLayout();
            this.gbDatabaseOperations.SuspendLayout();
            this.gbStoredProceduresOperations.SuspendLayout();
            this.gbViewOperations.SuspendLayout();
            this.gbTableOperations.SuspendLayout();
            this.gbServerOperations.SuspendLayout();
            this.SuspendLayout();
            // 
            // gbConnect
            // 
            this.gbConnect.BackColor = System.Drawing.Color.DarkGray;
            this.gbConnect.Controls.Add(this.gbServerOperations);
            this.gbConnect.Controls.Add(this.btnEnumAvailableSqlServers);
            this.gbConnect.Controls.Add(this.cbUserName);
            this.gbConnect.Controls.Add(this.lblInstancName);
            this.gbConnect.Controls.Add(this.ucDBInstanceList);
            this.gbConnect.Controls.Add(this.gbInstanceOperations);
            this.gbConnect.Controls.Add(this.btnLogoff);
            this.gbConnect.Controls.Add(this.btnLogon);
            this.gbConnect.Controls.Add(this.lblPassword);
            this.gbConnect.Controls.Add(this.lblUserName);
            this.gbConnect.Controls.Add(this.txtPassword);
            this.gbConnect.Controls.Add(this.ckIntegratedSecurity);
            this.gbConnect.Location = new System.Drawing.Point(3, 3);
            this.gbConnect.Name = "gbConnect";
            this.gbConnect.Size = new System.Drawing.Size(512, 903);
            this.gbConnect.TabIndex = 0;
            this.gbConnect.TabStop = false;
            this.gbConnect.Text = "Connect";
            // 
            // btnEnumAvailableSqlServers
            // 
            this.btnEnumAvailableSqlServers.Location = new System.Drawing.Point(331, 19);
            this.btnEnumAvailableSqlServers.Name = "btnEnumAvailableSqlServers";
            this.btnEnumAvailableSqlServers.Size = new System.Drawing.Size(162, 23);
            this.btnEnumAvailableSqlServers.TabIndex = 24;
            this.btnEnumAvailableSqlServers.Text = "Enum Available SQL Servers";
            this.btnEnumAvailableSqlServers.UseVisualStyleBackColor = true;
            this.btnEnumAvailableSqlServers.Click += new System.EventHandler(this.btnEnumAvailableSqlServers_Click);
            // 
            // cbUserName
            // 
            this.cbUserName.FormattingEnabled = true;
            this.cbUserName.Items.AddRange(new object[] {
            "btalk_user",
            "dataservices_select",
            "DCMAdmin",
            "dcmdev",
            "DOI_DBConnector",
            "EAC_SQLUser",
            "EAI_dbconnector",
            "inbound_admin",
            "mcreai",
            "naveditor",
            "navreader",
            "reportmodel_dbconnector",
            "RiskClass_Admin",
            "sa",
            "sradmin",
            "SSIS_DBConnector",
            "surplusreliefreporting_admin",
            "wseditor"});
            this.cbUserName.Location = new System.Drawing.Point(68, 86);
            this.cbUserName.Name = "cbUserName";
            this.cbUserName.Size = new System.Drawing.Size(171, 21);
            this.cbUserName.TabIndex = 23;
            this.cbUserName.Visible = false;
            // 
            // lblInstancName
            // 
            this.lblInstancName.AutoSize = true;
            this.lblInstancName.Location = new System.Drawing.Point(325, 29);
            this.lblInstancName.Name = "lblInstancName";
            this.lblInstancName.Size = new System.Drawing.Size(0, 13);
            this.lblInstancName.TabIndex = 13;
            // 
            // gbInstanceOperations
            // 
            this.gbInstanceOperations.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gbInstanceOperations.BackColor = System.Drawing.Color.LightGreen;
            this.gbInstanceOperations.Controls.Add(this.chkListInstanceDetails);
            this.gbInstanceOperations.Controls.Add(this.btnCreateDatabase);
            this.gbInstanceOperations.Controls.Add(this.txtNewDatabase);
            this.gbInstanceOperations.Controls.Add(this.gbDatabaseOperations);
            this.gbInstanceOperations.Controls.Add(this.btnCreateInstanceInfoWorkSheet);
            this.gbInstanceOperations.Location = new System.Drawing.Point(6, 137);
            this.gbInstanceOperations.Name = "gbInstanceOperations";
            this.gbInstanceOperations.Size = new System.Drawing.Size(498, 570);
            this.gbInstanceOperations.TabIndex = 1;
            this.gbInstanceOperations.TabStop = false;
            this.gbInstanceOperations.Text = "Instance Operations";
            // 
            // chkListInstanceDetails
            // 
            this.chkListInstanceDetails.AutoSize = true;
            this.chkListInstanceDetails.Checked = true;
            this.chkListInstanceDetails.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkListInstanceDetails.Location = new System.Drawing.Point(239, 23);
            this.chkListInstanceDetails.Name = "chkListInstanceDetails";
            this.chkListInstanceDetails.Size = new System.Drawing.Size(121, 17);
            this.chkListInstanceDetails.TabIndex = 7;
            this.chkListInstanceDetails.Text = "List Instance Details";
            this.chkListInstanceDetails.UseVisualStyleBackColor = true;
            // 
            // btnCreateDatabase
            // 
            this.btnCreateDatabase.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnCreateDatabase.Location = new System.Drawing.Point(366, 49);
            this.btnCreateDatabase.Name = "btnCreateDatabase";
            this.btnCreateDatabase.Size = new System.Drawing.Size(121, 23);
            this.btnCreateDatabase.TabIndex = 18;
            this.btnCreateDatabase.Text = "Create Database";
            this.btnCreateDatabase.UseVisualStyleBackColor = false;
            this.btnCreateDatabase.Click += new System.EventHandler(this.btnCreateDatabase_Click);
            // 
            // txtNewDatabase
            // 
            this.txtNewDatabase.Location = new System.Drawing.Point(366, 23);
            this.txtNewDatabase.Name = "txtNewDatabase";
            this.txtNewDatabase.Size = new System.Drawing.Size(121, 20);
            this.txtNewDatabase.TabIndex = 19;
            // 
            // gbDatabaseOperations
            // 
            this.gbDatabaseOperations.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gbDatabaseOperations.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.gbDatabaseOperations.Controls.Add(this.ckIncludeDBStoredProcedures);
            this.gbDatabaseOperations.Controls.Add(this.txtNewTable);
            this.gbDatabaseOperations.Controls.Add(this.ckIncludeDBViews);
            this.gbDatabaseOperations.Controls.Add(this.ckIncludeDBTables);
            this.gbDatabaseOperations.Controls.Add(this.btnCreateDatabaseInfoWorksheets);
            this.gbDatabaseOperations.Controls.Add(this.btnCreateTable);
            this.gbDatabaseOperations.Controls.Add(this.gbStoredProceduresOperations);
            this.gbDatabaseOperations.Controls.Add(this.cbDatabases);
            this.gbDatabaseOperations.Controls.Add(this.gbViewOperations);
            this.gbDatabaseOperations.Controls.Add(this.gbTableOperations);
            this.gbDatabaseOperations.Controls.Add(this.btnCreateDatabaseInfoWorkSheet);
            this.gbDatabaseOperations.Location = new System.Drawing.Point(6, 78);
            this.gbDatabaseOperations.Name = "gbDatabaseOperations";
            this.gbDatabaseOperations.Size = new System.Drawing.Size(485, 486);
            this.gbDatabaseOperations.TabIndex = 6;
            this.gbDatabaseOperations.TabStop = false;
            this.gbDatabaseOperations.Text = "Database Operations";
            // 
            // ckIncludeDBStoredProcedures
            // 
            this.ckIncludeDBStoredProcedures.AutoSize = true;
            this.ckIncludeDBStoredProcedures.Checked = true;
            this.ckIncludeDBStoredProcedures.CheckState = System.Windows.Forms.CheckState.Checked;
            this.ckIncludeDBStoredProcedures.Location = new System.Drawing.Point(258, 77);
            this.ckIncludeDBStoredProcedures.Name = "ckIncludeDBStoredProcedures";
            this.ckIncludeDBStoredProcedures.Size = new System.Drawing.Size(87, 17);
            this.ckIncludeDBStoredProcedures.TabIndex = 22;
            this.ckIncludeDBStoredProcedures.Text = "Stored Procs";
            this.ckIncludeDBStoredProcedures.UseVisualStyleBackColor = true;
            this.ckIncludeDBStoredProcedures.CheckedChanged += new System.EventHandler(this.ckIncludeDBStoredProcedures_CheckedChanged);
            // 
            // txtNewTable
            // 
            this.txtNewTable.Location = new System.Drawing.Point(355, 46);
            this.txtNewTable.Name = "txtNewTable";
            this.txtNewTable.Size = new System.Drawing.Size(121, 20);
            this.txtNewTable.TabIndex = 20;
            // 
            // ckIncludeDBViews
            // 
            this.ckIncludeDBViews.AutoSize = true;
            this.ckIncludeDBViews.Checked = true;
            this.ckIncludeDBViews.CheckState = System.Windows.Forms.CheckState.Checked;
            this.ckIncludeDBViews.Location = new System.Drawing.Point(258, 62);
            this.ckIncludeDBViews.Name = "ckIncludeDBViews";
            this.ckIncludeDBViews.Size = new System.Drawing.Size(54, 17);
            this.ckIncludeDBViews.TabIndex = 21;
            this.ckIncludeDBViews.Text = "Views";
            this.ckIncludeDBViews.UseVisualStyleBackColor = true;
            this.ckIncludeDBViews.CheckedChanged += new System.EventHandler(this.ckIncludeDBViews_CheckedChanged);
            // 
            // ckIncludeDBTables
            // 
            this.ckIncludeDBTables.AutoSize = true;
            this.ckIncludeDBTables.Checked = true;
            this.ckIncludeDBTables.CheckState = System.Windows.Forms.CheckState.Checked;
            this.ckIncludeDBTables.Location = new System.Drawing.Point(258, 47);
            this.ckIncludeDBTables.Name = "ckIncludeDBTables";
            this.ckIncludeDBTables.Size = new System.Drawing.Size(58, 17);
            this.ckIncludeDBTables.TabIndex = 20;
            this.ckIncludeDBTables.Text = "Tables";
            this.ckIncludeDBTables.UseVisualStyleBackColor = true;
            this.ckIncludeDBTables.CheckedChanged += new System.EventHandler(this.ckIncludeDBTables_CheckedChanged);
            // 
            // btnCreateDatabaseInfoWorksheets
            // 
            this.btnCreateDatabaseInfoWorksheets.BackColor = System.Drawing.Color.LightGreen;
            this.btnCreateDatabaseInfoWorksheets.Location = new System.Drawing.Point(6, 17);
            this.btnCreateDatabaseInfoWorksheets.Name = "btnCreateDatabaseInfoWorksheets";
            this.btnCreateDatabaseInfoWorksheets.Size = new System.Drawing.Size(246, 22);
            this.btnCreateDatabaseInfoWorksheets.TabIndex = 0;
            this.btnCreateDatabaseInfoWorksheets.Text = "DatabaseInfo Worksheets";
            this.toolTip1.SetToolTip(this.btnCreateDatabaseInfoWorksheets, "Create a worksheet for each Database in Instance");
            this.btnCreateDatabaseInfoWorksheets.UseVisualStyleBackColor = false;
            this.btnCreateDatabaseInfoWorksheets.Click += new System.EventHandler(this.btnCreateDatabaseInfoWorksheets_Click);
            // 
            // btnCreateTable
            // 
            this.btnCreateTable.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnCreateTable.Location = new System.Drawing.Point(355, 72);
            this.btnCreateTable.Name = "btnCreateTable";
            this.btnCreateTable.Size = new System.Drawing.Size(121, 23);
            this.btnCreateTable.TabIndex = 14;
            this.btnCreateTable.Text = "Create Table";
            this.btnCreateTable.UseVisualStyleBackColor = false;
            this.btnCreateTable.Click += new System.EventHandler(this.btnCreateTable_Click);
            // 
            // gbStoredProceduresOperations
            // 
            this.gbStoredProceduresOperations.BackColor = System.Drawing.Color.White;
            this.gbStoredProceduresOperations.Controls.Add(this.btnCreateStoredProcedureInfoWorksheet);
            this.gbStoredProceduresOperations.Controls.Add(this.cbStoredProcedures);
            this.gbStoredProceduresOperations.Controls.Add(this.btnCreateStoredProcedureInfoWorksheets);
            this.gbStoredProceduresOperations.Controls.Add(this.ckIncludeSystemStoredProcedures);
            this.gbStoredProceduresOperations.Location = new System.Drawing.Point(6, 351);
            this.gbStoredProceduresOperations.Name = "gbStoredProceduresOperations";
            this.gbStoredProceduresOperations.Size = new System.Drawing.Size(357, 130);
            this.gbStoredProceduresOperations.TabIndex = 16;
            this.gbStoredProceduresOperations.TabStop = false;
            this.gbStoredProceduresOperations.Text = "StoredProcedure Operations";
            // 
            // btnCreateStoredProcedureInfoWorksheet
            // 
            this.btnCreateStoredProcedureInfoWorksheet.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnCreateStoredProcedureInfoWorksheet.Location = new System.Drawing.Point(5, 98);
            this.btnCreateStoredProcedureInfoWorksheet.Name = "btnCreateStoredProcedureInfoWorksheet";
            this.btnCreateStoredProcedureInfoWorksheet.Size = new System.Drawing.Size(199, 23);
            this.btnCreateStoredProcedureInfoWorksheet.TabIndex = 15;
            this.btnCreateStoredProcedureInfoWorksheet.Text = "StoredProcedureInfo";
            this.btnCreateStoredProcedureInfoWorksheet.UseVisualStyleBackColor = false;
            this.btnCreateStoredProcedureInfoWorksheet.Click += new System.EventHandler(this.btnCreateStoredProcedureInfoWorksheet_Click);
            // 
            // cbStoredProcedures
            // 
            this.cbStoredProcedures.FormattingEnabled = true;
            this.cbStoredProcedures.Location = new System.Drawing.Point(6, 48);
            this.cbStoredProcedures.Name = "cbStoredProcedures";
            this.cbStoredProcedures.Size = new System.Drawing.Size(345, 21);
            this.cbStoredProcedures.TabIndex = 5;
            // 
            // btnCreateStoredProcedureInfoWorksheets
            // 
            this.btnCreateStoredProcedureInfoWorksheets.BackColor = System.Drawing.Color.LightGreen;
            this.btnCreateStoredProcedureInfoWorksheets.Location = new System.Drawing.Point(7, 19);
            this.btnCreateStoredProcedureInfoWorksheets.Name = "btnCreateStoredProcedureInfoWorksheets";
            this.btnCreateStoredProcedureInfoWorksheets.Size = new System.Drawing.Size(197, 23);
            this.btnCreateStoredProcedureInfoWorksheets.TabIndex = 17;
            this.btnCreateStoredProcedureInfoWorksheets.Text = "StoredProcedureInfo WorkSheets";
            this.btnCreateStoredProcedureInfoWorksheets.UseVisualStyleBackColor = false;
            this.btnCreateStoredProcedureInfoWorksheets.Click += new System.EventHandler(this.btnCreateStoredProcedureInfoWorksheets_Click);
            // 
            // ckIncludeSystemStoredProcedures
            // 
            this.ckIncludeSystemStoredProcedures.AutoSize = true;
            this.ckIncludeSystemStoredProcedures.Location = new System.Drawing.Point(7, 75);
            this.ckIncludeSystemStoredProcedures.Name = "ckIncludeSystemStoredProcedures";
            this.ckIncludeSystemStoredProcedures.Size = new System.Drawing.Size(186, 17);
            this.ckIncludeSystemStoredProcedures.TabIndex = 10;
            this.ckIncludeSystemStoredProcedures.Text = "Include System StoredProcedures";
            this.ckIncludeSystemStoredProcedures.UseVisualStyleBackColor = true;
            // 
            // cbDatabases
            // 
            this.cbDatabases.FormattingEnabled = true;
            this.cbDatabases.Location = new System.Drawing.Point(6, 45);
            this.cbDatabases.Name = "cbDatabases";
            this.cbDatabases.Size = new System.Drawing.Size(246, 21);
            this.cbDatabases.TabIndex = 2;
            this.cbDatabases.SelectedIndexChanged += new System.EventHandler(this.cbDatabases_SelectedIndexChanged);
            // 
            // gbViewOperations
            // 
            this.gbViewOperations.BackColor = System.Drawing.Color.White;
            this.gbViewOperations.Controls.Add(this.btnListRows_View);
            this.gbViewOperations.Controls.Add(this.ckIncludeSystemViews);
            this.gbViewOperations.Controls.Add(this.btnCreateViewInfoWorkSheet);
            this.gbViewOperations.Controls.Add(this.btnCreateViewInfoWorksheets);
            this.gbViewOperations.Controls.Add(this.cbViews);
            this.gbViewOperations.Location = new System.Drawing.Point(6, 217);
            this.gbViewOperations.Name = "gbViewOperations";
            this.gbViewOperations.Size = new System.Drawing.Size(357, 128);
            this.gbViewOperations.TabIndex = 13;
            this.gbViewOperations.TabStop = false;
            this.gbViewOperations.Text = "View Operations";
            // 
            // btnListRows_View
            // 
            this.btnListRows_View.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnListRows_View.Location = new System.Drawing.Point(230, 99);
            this.btnListRows_View.Name = "btnListRows_View";
            this.btnListRows_View.Size = new System.Drawing.Size(121, 23);
            this.btnListRows_View.TabIndex = 22;
            this.btnListRows_View.Text = "List Rows";
            this.btnListRows_View.UseVisualStyleBackColor = false;
            this.btnListRows_View.Click += new System.EventHandler(this.btnListRows_View_Click);
            // 
            // ckIncludeSystemViews
            // 
            this.ckIncludeSystemViews.AutoSize = true;
            this.ckIncludeSystemViews.Location = new System.Drawing.Point(7, 76);
            this.ckIncludeSystemViews.Name = "ckIncludeSystemViews";
            this.ckIncludeSystemViews.Size = new System.Drawing.Size(129, 17);
            this.ckIncludeSystemViews.TabIndex = 16;
            this.ckIncludeSystemViews.Text = "Include System Views";
            this.ckIncludeSystemViews.UseVisualStyleBackColor = true;
            // 
            // btnCreateViewInfoWorkSheet
            // 
            this.btnCreateViewInfoWorkSheet.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnCreateViewInfoWorkSheet.Location = new System.Drawing.Point(5, 99);
            this.btnCreateViewInfoWorkSheet.Name = "btnCreateViewInfoWorkSheet";
            this.btnCreateViewInfoWorkSheet.Size = new System.Drawing.Size(199, 23);
            this.btnCreateViewInfoWorkSheet.TabIndex = 15;
            this.btnCreateViewInfoWorkSheet.Text = "ViewInfo";
            this.btnCreateViewInfoWorkSheet.UseVisualStyleBackColor = false;
            this.btnCreateViewInfoWorkSheet.Click += new System.EventHandler(this.btnCreateViewInfoWorkSheet_Click);
            // 
            // btnCreateViewInfoWorksheets
            // 
            this.btnCreateViewInfoWorksheets.BackColor = System.Drawing.Color.LightGreen;
            this.btnCreateViewInfoWorksheets.Location = new System.Drawing.Point(7, 20);
            this.btnCreateViewInfoWorksheets.Name = "btnCreateViewInfoWorksheets";
            this.btnCreateViewInfoWorksheets.Size = new System.Drawing.Size(197, 23);
            this.btnCreateViewInfoWorksheets.TabIndex = 5;
            this.btnCreateViewInfoWorksheets.Text = "ViewInfo WorkSheets";
            this.btnCreateViewInfoWorksheets.UseVisualStyleBackColor = false;
            this.btnCreateViewInfoWorksheets.Click += new System.EventHandler(this.btnCreateViewInfoWorksheets_Click);
            // 
            // cbViews
            // 
            this.cbViews.FormattingEnabled = true;
            this.cbViews.Location = new System.Drawing.Point(7, 49);
            this.cbViews.Name = "cbViews";
            this.cbViews.Size = new System.Drawing.Size(344, 21);
            this.cbViews.TabIndex = 4;
            // 
            // gbTableOperations
            // 
            this.gbTableOperations.BackColor = System.Drawing.Color.White;
            this.gbTableOperations.Controls.Add(this.btnListRows_Table);
            this.gbTableOperations.Controls.Add(this.btnCreateTableInfoWorksheet);
            this.gbTableOperations.Controls.Add(this.cbTables);
            this.gbTableOperations.Controls.Add(this.btnCreateTableInfoWorkSheets);
            this.gbTableOperations.Location = new System.Drawing.Point(6, 101);
            this.gbTableOperations.Name = "gbTableOperations";
            this.gbTableOperations.Size = new System.Drawing.Size(357, 110);
            this.gbTableOperations.TabIndex = 12;
            this.gbTableOperations.TabStop = false;
            this.gbTableOperations.Text = "Table Operations";
            // 
            // btnListRows_Table
            // 
            this.btnListRows_Table.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnListRows_Table.Location = new System.Drawing.Point(230, 75);
            this.btnListRows_Table.Name = "btnListRows_Table";
            this.btnListRows_Table.Size = new System.Drawing.Size(121, 23);
            this.btnListRows_Table.TabIndex = 21;
            this.btnListRows_Table.Text = "List Rows";
            this.btnListRows_Table.UseVisualStyleBackColor = false;
            this.btnListRows_Table.Click += new System.EventHandler(this.btnListRows_Table_Click);
            // 
            // btnCreateTableInfoWorksheet
            // 
            this.btnCreateTableInfoWorksheet.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnCreateTableInfoWorksheet.Location = new System.Drawing.Point(5, 75);
            this.btnCreateTableInfoWorksheet.Name = "btnCreateTableInfoWorksheet";
            this.btnCreateTableInfoWorksheet.Size = new System.Drawing.Size(199, 23);
            this.btnCreateTableInfoWorksheet.TabIndex = 14;
            this.btnCreateTableInfoWorksheet.Text = "TableInfo";
            this.toolTip1.SetToolTip(this.btnCreateTableInfoWorksheet, "Create TableInfo Worksheet");
            this.btnCreateTableInfoWorksheet.UseVisualStyleBackColor = false;
            this.btnCreateTableInfoWorksheet.Click += new System.EventHandler(this.btnCreateTableInfoWorksheet_Click);
            // 
            // cbTables
            // 
            this.cbTables.FormattingEnabled = true;
            this.cbTables.Location = new System.Drawing.Point(5, 48);
            this.cbTables.Name = "cbTables";
            this.cbTables.Size = new System.Drawing.Size(346, 21);
            this.cbTables.TabIndex = 4;
            // 
            // btnCreateTableInfoWorkSheets
            // 
            this.btnCreateTableInfoWorkSheets.BackColor = System.Drawing.Color.LightGreen;
            this.btnCreateTableInfoWorkSheets.Location = new System.Drawing.Point(5, 19);
            this.btnCreateTableInfoWorkSheets.Name = "btnCreateTableInfoWorkSheets";
            this.btnCreateTableInfoWorkSheets.Size = new System.Drawing.Size(199, 23);
            this.btnCreateTableInfoWorkSheets.TabIndex = 3;
            this.btnCreateTableInfoWorkSheets.Text = "TableInfo WorkSheets";
            this.btnCreateTableInfoWorkSheets.UseVisualStyleBackColor = false;
            this.btnCreateTableInfoWorkSheets.Click += new System.EventHandler(this.btnCreateTableInfoWorkSheets_Click);
            // 
            // btnCreateDatabaseInfoWorkSheet
            // 
            this.btnCreateDatabaseInfoWorkSheet.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnCreateDatabaseInfoWorkSheet.Location = new System.Drawing.Point(6, 72);
            this.btnCreateDatabaseInfoWorkSheet.Name = "btnCreateDatabaseInfoWorkSheet";
            this.btnCreateDatabaseInfoWorkSheet.Size = new System.Drawing.Size(246, 23);
            this.btnCreateDatabaseInfoWorkSheet.TabIndex = 4;
            this.btnCreateDatabaseInfoWorkSheet.Text = "DatabaseInfo";
            this.btnCreateDatabaseInfoWorkSheet.UseVisualStyleBackColor = false;
            this.btnCreateDatabaseInfoWorkSheet.Click += new System.EventHandler(this.btnCreateDatabaseInfoWorkSheet_Click);
            // 
            // btnCreateInstanceInfoWorkSheet
            // 
            this.btnCreateInstanceInfoWorkSheet.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnCreateInstanceInfoWorkSheet.Location = new System.Drawing.Point(7, 19);
            this.btnCreateInstanceInfoWorkSheet.Name = "btnCreateInstanceInfoWorkSheet";
            this.btnCreateInstanceInfoWorkSheet.Size = new System.Drawing.Size(226, 23);
            this.btnCreateInstanceInfoWorkSheet.TabIndex = 2;
            this.btnCreateInstanceInfoWorkSheet.Text = "InstanceInfo";
            this.btnCreateInstanceInfoWorkSheet.UseVisualStyleBackColor = false;
            this.btnCreateInstanceInfoWorkSheet.Click += new System.EventHandler(this.btnCreateInstanceInfoWorkSheet_Click);
            // 
            // btnLogoff
            // 
            this.btnLogoff.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnLogoff.Enabled = false;
            this.btnLogoff.Location = new System.Drawing.Point(249, 89);
            this.btnLogoff.Name = "btnLogoff";
            this.btnLogoff.Size = new System.Drawing.Size(75, 23);
            this.btnLogoff.TabIndex = 8;
            this.btnLogoff.Text = "Logoff";
            this.btnLogoff.UseVisualStyleBackColor = false;
            this.btnLogoff.Click += new System.EventHandler(this.btnLogoff_Click);
            // 
            // btnLogon
            // 
            this.btnLogon.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnLogon.Location = new System.Drawing.Point(248, 63);
            this.btnLogon.Name = "btnLogon";
            this.btnLogon.Size = new System.Drawing.Size(75, 23);
            this.btnLogon.TabIndex = 7;
            this.btnLogon.Text = "Logon";
            this.btnLogon.UseVisualStyleBackColor = false;
            this.btnLogon.Click += new System.EventHandler(this.btnLogon_Click);
            // 
            // lblPassword
            // 
            this.lblPassword.AutoSize = true;
            this.lblPassword.Location = new System.Drawing.Point(9, 113);
            this.lblPassword.Name = "lblPassword";
            this.lblPassword.Size = new System.Drawing.Size(53, 13);
            this.lblPassword.TabIndex = 6;
            this.lblPassword.Text = "Password";
            this.lblPassword.Visible = false;
            // 
            // lblUserName
            // 
            this.lblUserName.AutoSize = true;
            this.lblUserName.Location = new System.Drawing.Point(5, 89);
            this.lblUserName.Name = "lblUserName";
            this.lblUserName.Size = new System.Drawing.Size(57, 13);
            this.lblUserName.TabIndex = 5;
            this.lblUserName.Text = "UserName";
            this.lblUserName.Visible = false;
            // 
            // txtPassword
            // 
            this.txtPassword.Location = new System.Drawing.Point(68, 110);
            this.txtPassword.Name = "txtPassword";
            this.txtPassword.PasswordChar = '*';
            this.txtPassword.Size = new System.Drawing.Size(171, 20);
            this.txtPassword.TabIndex = 4;
            this.txtPassword.Visible = false;
            // 
            // ckIntegratedSecurity
            // 
            this.ckIntegratedSecurity.AutoSize = true;
            this.ckIntegratedSecurity.Checked = true;
            this.ckIntegratedSecurity.CheckState = System.Windows.Forms.CheckState.Checked;
            this.ckIntegratedSecurity.Location = new System.Drawing.Point(6, 63);
            this.ckIntegratedSecurity.Name = "ckIntegratedSecurity";
            this.ckIntegratedSecurity.Size = new System.Drawing.Size(137, 17);
            this.ckIntegratedSecurity.TabIndex = 0;
            this.ckIntegratedSecurity.Text = "Use Integrated Security";
            this.ckIntegratedSecurity.UseVisualStyleBackColor = true;
            this.ckIntegratedSecurity.CheckedChanged += new System.EventHandler(this.ckIntegratedSecurity_CheckedChanged);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // gbServerOperations
            // 
            this.gbServerOperations.BackColor = System.Drawing.Color.LightSkyBlue;
            this.gbServerOperations.Controls.Add(this.btnGetServerInfo);
            this.gbServerOperations.Controls.Add(this.label1);
            this.gbServerOperations.Controls.Add(this.txtServerName);
            this.gbServerOperations.Location = new System.Drawing.Point(6, 713);
            this.gbServerOperations.Name = "gbServerOperations";
            this.gbServerOperations.Size = new System.Drawing.Size(498, 184);
            this.gbServerOperations.TabIndex = 25;
            this.gbServerOperations.TabStop = false;
            this.gbServerOperations.Text = "Server Operations";
            // 
            // txtServerName
            // 
            this.txtServerName.Location = new System.Drawing.Point(88, 19);
            this.txtServerName.Name = "txtServerName";
            this.txtServerName.Size = new System.Drawing.Size(145, 20);
            this.txtServerName.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(66, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "ServerName";
            this.label1.Visible = false;
            // 
            // ucDBInstanceList
            // 
            this.ucDBInstanceList.BackColor = System.Drawing.SystemColors.ControlDark;
            this.ucDBInstanceList.Location = new System.Drawing.Point(6, 15);
            this.ucDBInstanceList.Name = "ucDBInstanceList";
            this.ucDBInstanceList.Size = new System.Drawing.Size(313, 42);
            this.ucDBInstanceList.TabIndex = 12;
            this.ucDBInstanceList.InstanceSelectionChanged_Event += new SupportTools_Excel.User_Interface.User_Controls.ucDBInstanceList.InstanceSelectionChanged(this.ucDBInstanceList_InstanceSelectionChanged_Event);
            // 
            // btnGetServerInfo
            // 
            this.btnGetServerInfo.Location = new System.Drawing.Point(247, 19);
            this.btnGetServerInfo.Name = "btnGetServerInfo";
            this.btnGetServerInfo.Size = new System.Drawing.Size(122, 23);
            this.btnGetServerInfo.TabIndex = 26;
            this.btnGetServerInfo.Text = "GetServer Info";
            this.btnGetServerInfo.UseVisualStyleBackColor = true;
            this.btnGetServerInfo.Click += new System.EventHandler(this.btnGetServerInfo_Click);
            // 
            // TaskPane_SQLSMO
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.Controls.Add(this.gbConnect);
            this.MinimumSize = new System.Drawing.Size(515, 0);
            this.Name = "TaskPane_SQLSMO";
            this.Size = new System.Drawing.Size(515, 909);
            this.Load += new System.EventHandler(this.TaskPane_SQLSMO_Load);
            this.gbConnect.ResumeLayout(false);
            this.gbConnect.PerformLayout();
            this.gbInstanceOperations.ResumeLayout(false);
            this.gbInstanceOperations.PerformLayout();
            this.gbDatabaseOperations.ResumeLayout(false);
            this.gbDatabaseOperations.PerformLayout();
            this.gbStoredProceduresOperations.ResumeLayout(false);
            this.gbStoredProceduresOperations.PerformLayout();
            this.gbViewOperations.ResumeLayout(false);
            this.gbViewOperations.PerformLayout();
            this.gbTableOperations.ResumeLayout(false);
            this.gbServerOperations.ResumeLayout(false);
            this.gbServerOperations.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox gbConnect;
        private System.Windows.Forms.Button btnLogoff;
        private System.Windows.Forms.Button btnLogon;
        private System.Windows.Forms.Label lblPassword;
        private System.Windows.Forms.Label lblUserName;
        private System.Windows.Forms.TextBox txtPassword;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.CheckBox ckIntegratedSecurity;
        private System.Windows.Forms.GroupBox gbInstanceOperations;
        private System.Windows.Forms.ComboBox cbDatabases;
        private System.Windows.Forms.Button btnCreateDatabaseInfoWorksheets;
        private System.Windows.Forms.ComboBox cbStoredProcedures;
        private System.Windows.Forms.ComboBox cbTables;
        private System.Windows.Forms.CheckBox ckIncludeSystemStoredProcedures;
        private System.Windows.Forms.Button btnCreateInstanceInfoWorkSheet;
        private System.Windows.Forms.Button btnCreateTableInfoWorkSheets;
        private System.Windows.Forms.Button btnCreateDatabaseInfoWorkSheet;
        private System.Windows.Forms.Button btnCreateViewInfoWorksheets;
        private System.Windows.Forms.GroupBox gbDatabaseOperations;
        private System.Windows.Forms.GroupBox gbViewOperations;
        private System.Windows.Forms.ComboBox cbViews;
        private System.Windows.Forms.GroupBox gbTableOperations;
        private System.Windows.Forms.Button btnCreateViewInfoWorkSheet;
        private System.Windows.Forms.Button btnCreateTableInfoWorksheet;
        private System.Windows.Forms.Button btnCreateStoredProcedureInfoWorksheets;
        private System.Windows.Forms.GroupBox gbStoredProceduresOperations;
        private System.Windows.Forms.Button btnCreateStoredProcedureInfoWorksheet;
        private System.Windows.Forms.CheckBox ckIncludeSystemViews;
        private System.Windows.Forms.CheckBox chkListInstanceDetails;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private User_Controls.ucDBInstanceList ucDBInstanceList;
        private System.Windows.Forms.Label lblInstancName;
        private System.Windows.Forms.Button btnCreateTable;
        private System.Windows.Forms.Button btnCreateDatabase;
        private System.Windows.Forms.TextBox txtNewTable;
        private System.Windows.Forms.TextBox txtNewDatabase;
        private System.Windows.Forms.Button btnListRows_View;
        private System.Windows.Forms.Button btnListRows_Table;
        private System.Windows.Forms.CheckBox ckIncludeDBStoredProcedures;
        private System.Windows.Forms.CheckBox ckIncludeDBViews;
        private System.Windows.Forms.CheckBox ckIncludeDBTables;
        private System.Windows.Forms.ComboBox cbUserName;
        private System.Windows.Forms.Button btnEnumAvailableSqlServers;
        private System.Windows.Forms.GroupBox gbServerOperations;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtServerName;
        private System.Windows.Forms.Button btnGetServerInfo;

    }
}
