namespace SupportTools_Excel.User_Interface.Task_Panes
{
    partial class TaskPane_ITRs
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
            this.btnAddPageBreaks = new System.Windows.Forms.Button();
            this.gbITRWork = new System.Windows.Forms.GroupBox();
            this.btnDisplayITRDetail = new System.Windows.Forms.Button();
            this.btnGetITRInformation = new System.Windows.Forms.Button();
            this.btnMergeDuplicateRows = new System.Windows.Forms.Button();
            this.btnFormatSourceITRs = new System.Windows.Forms.Button();
            this.btnAddPivotTables = new System.Windows.Forms.Button();
            this.btnAddListObjects = new System.Windows.Forms.Button();
            this.ToolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.btnDuplicateInputSheet = new System.Windows.Forms.Button();
            this.btnProcessDuplicateRows = new System.Windows.Forms.Button();
            this.btnProcessDynamicOutput = new System.Windows.Forms.Button();
            this.btnProcessDARTOutput = new System.Windows.Forms.Button();
            this.gbDebug = new System.Windows.Forms.GroupBox();
            this.cmbTeamName = new System.Windows.Forms.ComboBox();
            this.Label1 = new System.Windows.Forms.Label();
            this.gbDARTReport = new System.Windows.Forms.GroupBox();
            this.gbITRWork.SuspendLayout();
            this.gbDebug.SuspendLayout();
            this.gbDARTReport.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnAddPageBreaks
            // 
            this.btnAddPageBreaks.Location = new System.Drawing.Point(10, 173);
            this.btnAddPageBreaks.Name = "btnAddPageBreaks";
            this.btnAddPageBreaks.Size = new System.Drawing.Size(142, 23);
            this.btnAddPageBreaks.TabIndex = 20;
            this.btnAddPageBreaks.Text = "Add Page Breaks";
            this.btnAddPageBreaks.UseVisualStyleBackColor = true;
            this.btnAddPageBreaks.Click += new System.EventHandler(this.btnAddPageBreaks_Click);
            // 
            // gbITRWork
            // 
            this.gbITRWork.Controls.Add(this.btnDisplayITRDetail);
            this.gbITRWork.Controls.Add(this.btnGetITRInformation);
            this.gbITRWork.Controls.Add(this.btnMergeDuplicateRows);
            this.gbITRWork.Location = new System.Drawing.Point(3, 346);
            this.gbITRWork.Name = "gbITRWork";
            this.gbITRWork.Size = new System.Drawing.Size(182, 288);
            this.gbITRWork.TabIndex = 24;
            this.gbITRWork.TabStop = false;
            this.gbITRWork.Text = "ITR Work";
            // 
            // btnDisplayITRDetail
            // 
            this.btnDisplayITRDetail.Location = new System.Drawing.Point(15, 77);
            this.btnDisplayITRDetail.Name = "btnDisplayITRDetail";
            this.btnDisplayITRDetail.Size = new System.Drawing.Size(155, 23);
            this.btnDisplayITRDetail.TabIndex = 32;
            this.btnDisplayITRDetail.Text = "Display ITR Detail";
            this.btnDisplayITRDetail.UseVisualStyleBackColor = true;
            this.btnDisplayITRDetail.Click += new System.EventHandler(this.btnDisplayITRDetail_Click);
            // 
            // btnGetITRInformation
            // 
            this.btnGetITRInformation.Location = new System.Drawing.Point(15, 48);
            this.btnGetITRInformation.Name = "btnGetITRInformation";
            this.btnGetITRInformation.Size = new System.Drawing.Size(155, 23);
            this.btnGetITRInformation.TabIndex = 31;
            this.btnGetITRInformation.Text = "Get ITR Information";
            this.btnGetITRInformation.UseVisualStyleBackColor = true;
            this.btnGetITRInformation.Click += new System.EventHandler(this.btnGetITRInformation_Click);
            // 
            // btnMergeDuplicateRows
            // 
            this.btnMergeDuplicateRows.Location = new System.Drawing.Point(15, 19);
            this.btnMergeDuplicateRows.Name = "btnMergeDuplicateRows";
            this.btnMergeDuplicateRows.Size = new System.Drawing.Size(155, 23);
            this.btnMergeDuplicateRows.TabIndex = 30;
            this.btnMergeDuplicateRows.Text = "Merge Duplicate Rows";
            this.btnMergeDuplicateRows.UseVisualStyleBackColor = true;
            this.btnMergeDuplicateRows.Click += new System.EventHandler(this.btnMergeDuplicateRows_Click);
            // 
            // btnFormatSourceITRs
            // 
            this.btnFormatSourceITRs.Location = new System.Drawing.Point(10, 150);
            this.btnFormatSourceITRs.Name = "btnFormatSourceITRs";
            this.btnFormatSourceITRs.Size = new System.Drawing.Size(142, 23);
            this.btnFormatSourceITRs.TabIndex = 19;
            this.btnFormatSourceITRs.Text = "Format SourceITRs";
            this.btnFormatSourceITRs.UseVisualStyleBackColor = true;
            this.btnFormatSourceITRs.Click += new System.EventHandler(this.btnFormatSourceITRs_Click);
            // 
            // btnAddPivotTables
            // 
            this.btnAddPivotTables.Location = new System.Drawing.Point(10, 111);
            this.btnAddPivotTables.Name = "btnAddPivotTables";
            this.btnAddPivotTables.Size = new System.Drawing.Size(142, 23);
            this.btnAddPivotTables.TabIndex = 16;
            this.btnAddPivotTables.Text = "Add Pivot Tables";
            this.btnAddPivotTables.UseVisualStyleBackColor = true;
            this.btnAddPivotTables.Click += new System.EventHandler(this.btnAddPivotTables_Click);
            // 
            // btnAddListObjects
            // 
            this.btnAddListObjects.Location = new System.Drawing.Point(10, 88);
            this.btnAddListObjects.Name = "btnAddListObjects";
            this.btnAddListObjects.Size = new System.Drawing.Size(142, 23);
            this.btnAddListObjects.TabIndex = 6;
            this.btnAddListObjects.Text = "Add ListObjects";
            this.btnAddListObjects.UseVisualStyleBackColor = true;
            this.btnAddListObjects.Click += new System.EventHandler(this.btnAddListObjects_Click);
            // 
            // btnDuplicateInputSheet
            // 
            this.btnDuplicateInputSheet.Location = new System.Drawing.Point(10, 42);
            this.btnDuplicateInputSheet.Name = "btnDuplicateInputSheet";
            this.btnDuplicateInputSheet.Size = new System.Drawing.Size(142, 23);
            this.btnDuplicateInputSheet.TabIndex = 18;
            this.btnDuplicateInputSheet.Text = "Duplicate Input Sheet";
            this.ToolTip1.SetToolTip(this.btnDuplicateInputSheet, "Step One");
            this.btnDuplicateInputSheet.UseVisualStyleBackColor = true;
            this.btnDuplicateInputSheet.Click += new System.EventHandler(this.btnDuplicateInputSheet_Click);
            // 
            // btnProcessDuplicateRows
            // 
            this.btnProcessDuplicateRows.Location = new System.Drawing.Point(10, 65);
            this.btnProcessDuplicateRows.Name = "btnProcessDuplicateRows";
            this.btnProcessDuplicateRows.Size = new System.Drawing.Size(142, 23);
            this.btnProcessDuplicateRows.TabIndex = 17;
            this.btnProcessDuplicateRows.Text = "Process Duplicate Rows";
            this.ToolTip1.SetToolTip(this.btnProcessDuplicateRows, "Step One");
            this.btnProcessDuplicateRows.UseVisualStyleBackColor = true;
            this.btnProcessDuplicateRows.Click += new System.EventHandler(this.btnProcessDuplicateRow_Click);
            // 
            // btnProcessDynamicOutput
            // 
            this.btnProcessDynamicOutput.Location = new System.Drawing.Point(10, 19);
            this.btnProcessDynamicOutput.Name = "btnProcessDynamicOutput";
            this.btnProcessDynamicOutput.Size = new System.Drawing.Size(142, 23);
            this.btnProcessDynamicOutput.TabIndex = 3;
            this.btnProcessDynamicOutput.Text = "Process Dynamic Output";
            this.ToolTip1.SetToolTip(this.btnProcessDynamicOutput, "Step One");
            this.btnProcessDynamicOutput.UseVisualStyleBackColor = true;
            this.btnProcessDynamicOutput.Click += new System.EventHandler(this.btnProcessDynamicOutput_Click);
            // 
            // btnProcessDARTOutput
            // 
            this.btnProcessDARTOutput.Location = new System.Drawing.Point(8, 60);
            this.btnProcessDARTOutput.Name = "btnProcessDARTOutput";
            this.btnProcessDARTOutput.Size = new System.Drawing.Size(162, 23);
            this.btnProcessDARTOutput.TabIndex = 21;
            this.btnProcessDARTOutput.Text = "Process DART Output";
            this.ToolTip1.SetToolTip(this.btnProcessDARTOutput, "Step One");
            this.btnProcessDARTOutput.UseVisualStyleBackColor = true;
            this.btnProcessDARTOutput.Click += new System.EventHandler(this.btnProcessDARTOutput_Click);
            // 
            // gbDebug
            // 
            this.gbDebug.Controls.Add(this.btnAddPageBreaks);
            this.gbDebug.Controls.Add(this.btnFormatSourceITRs);
            this.gbDebug.Controls.Add(this.btnDuplicateInputSheet);
            this.gbDebug.Controls.Add(this.btnProcessDuplicateRows);
            this.gbDebug.Controls.Add(this.btnAddPivotTables);
            this.gbDebug.Controls.Add(this.btnAddListObjects);
            this.gbDebug.Controls.Add(this.btnProcessDynamicOutput);
            this.gbDebug.Location = new System.Drawing.Point(9, 106);
            this.gbDebug.Name = "gbDebug";
            this.gbDebug.Size = new System.Drawing.Size(162, 207);
            this.gbDebug.TabIndex = 24;
            this.gbDebug.TabStop = false;
            this.gbDebug.Text = "Debug";
            // 
            // cmbTeamName
            // 
            this.cmbTeamName.FormattingEnabled = true;
            this.cmbTeamName.Items.AddRange(new object[] {
            "Data Services",
            "Integration Services",
            "Reporting Services"});
            this.cmbTeamName.Location = new System.Drawing.Point(9, 32);
            this.cmbTeamName.Name = "cmbTeamName";
            this.cmbTeamName.Size = new System.Drawing.Size(161, 21);
            this.cmbTeamName.TabIndex = 23;
            // 
            // Label1
            // 
            this.Label1.AutoSize = true;
            this.Label1.Location = new System.Drawing.Point(6, 16);
            this.Label1.Name = "Label1";
            this.Label1.Size = new System.Drawing.Size(65, 13);
            this.Label1.TabIndex = 22;
            this.Label1.Text = "Team Name";
            // 
            // gbDARTReport
            // 
            this.gbDARTReport.Controls.Add(this.gbDebug);
            this.gbDARTReport.Controls.Add(this.cmbTeamName);
            this.gbDARTReport.Controls.Add(this.Label1);
            this.gbDARTReport.Controls.Add(this.btnProcessDARTOutput);
            this.gbDARTReport.Location = new System.Drawing.Point(3, 15);
            this.gbDARTReport.Name = "gbDARTReport";
            this.gbDARTReport.Size = new System.Drawing.Size(182, 325);
            this.gbDARTReport.TabIndex = 23;
            this.gbDARTReport.TabStop = false;
            this.gbDARTReport.Text = "DART Report";
            // 
            // TaskPane_ITRs
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.Controls.Add(this.gbITRWork);
            this.Controls.Add(this.gbDARTReport);
            this.MinimumSize = new System.Drawing.Size(200, 0);
            this.Name = "TaskPane_ITRs";
            this.Size = new System.Drawing.Size(200, 670);
            this.gbITRWork.ResumeLayout(false);
            this.gbDebug.ResumeLayout(false);
            this.gbDARTReport.ResumeLayout(false);
            this.gbDARTReport.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal System.Windows.Forms.Button btnAddPageBreaks;
        internal System.Windows.Forms.GroupBox gbITRWork;
        internal System.Windows.Forms.Button btnDisplayITRDetail;
        internal System.Windows.Forms.Button btnGetITRInformation;
        internal System.Windows.Forms.Button btnMergeDuplicateRows;
        internal System.Windows.Forms.Button btnFormatSourceITRs;
        internal System.Windows.Forms.Button btnAddPivotTables;
        internal System.Windows.Forms.Button btnAddListObjects;
        internal System.Windows.Forms.ToolTip ToolTip1;
        internal System.Windows.Forms.Button btnDuplicateInputSheet;
        internal System.Windows.Forms.Button btnProcessDuplicateRows;
        internal System.Windows.Forms.Button btnProcessDynamicOutput;
        internal System.Windows.Forms.Button btnProcessDARTOutput;
        internal System.Windows.Forms.GroupBox gbDebug;
        internal System.Windows.Forms.ComboBox cmbTeamName;
        internal System.Windows.Forms.Label Label1;
        internal System.Windows.Forms.GroupBox gbDARTReport;
    }
}
