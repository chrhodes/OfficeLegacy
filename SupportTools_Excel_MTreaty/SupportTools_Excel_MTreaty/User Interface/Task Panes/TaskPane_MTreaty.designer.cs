namespace SupportTools_Excel_MTreaty.User_Interface.Task_Panes
{
    partial class TaskPane_MTreaty
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
            this.gbTask_SelectEnvironment = new System.Windows.Forms.GroupBox();
            this.pnlTask_SelectEnvironment = new System.Windows.Forms.Panel();
            this.gbTask_SelectFileType = new System.Windows.Forms.GroupBox();
            this.pnlTask_SelectFileType = new System.Windows.Forms.Panel();
            this.btnOpenFile = new System.Windows.Forms.Button();
            this.gbTask_ProcessFile = new System.Windows.Forms.GroupBox();
            this.pnlTask_ProcessFile = new System.Windows.Forms.Panel();
            this.btnProcessFile = new System.Windows.Forms.Button();
            this.dtpAsOfDate = new System.Windows.Forms.DateTimePicker();
            this.btnSaveOutputFiles = new System.Windows.Forms.Button();
            this.gbTask_OpenFile = new System.Windows.Forms.GroupBox();
            this.pnlTask_OpenFile = new System.Windows.Forms.Panel();
            this.gbTask_SaveFiles = new System.Windows.Forms.GroupBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.pnlTask_SaveControlFile = new System.Windows.Forms.Panel();
            this.pnlTask_SaveDetailFile = new System.Windows.Forms.Panel();
            this.gbTask_ConfirmAsOfDate = new System.Windows.Forms.GroupBox();
            this.pnlTask_ConfirmAsOfDate = new System.Windows.Forms.Panel();
            this.gbTask_ReviewOutput = new System.Windows.Forms.GroupBox();
            this.btnOpenOutputFolder = new System.Windows.Forms.Button();
            this.ucFileTypeList1 = new SupportTools_Excel_MTreaty.User_Interface.User_Controls.ucFileTypeList();
            this.ucEnvironmentList1 = new SupportTools_Excel_MTreaty.User_Interface.User_Controls.ucEnvironmentList();
            this.btnPlayAgain = new System.Windows.Forms.Button();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.gbTask_SelectEnvironment.SuspendLayout();
            this.gbTask_SelectFileType.SuspendLayout();
            this.gbTask_ProcessFile.SuspendLayout();
            this.gbTask_OpenFile.SuspendLayout();
            this.gbTask_SaveFiles.SuspendLayout();
            this.gbTask_ConfirmAsOfDate.SuspendLayout();
            this.gbTask_ReviewOutput.SuspendLayout();
            this.SuspendLayout();
            // 
            // gbTask_SelectEnvironment
            // 
            this.gbTask_SelectEnvironment.Controls.Add(this.pnlTask_SelectEnvironment);
            this.gbTask_SelectEnvironment.Controls.Add(this.ucEnvironmentList1);
            this.gbTask_SelectEnvironment.Location = new System.Drawing.Point(3, 3);
            this.gbTask_SelectEnvironment.Name = "gbTask_SelectEnvironment";
            this.gbTask_SelectEnvironment.Size = new System.Drawing.Size(256, 71);
            this.gbTask_SelectEnvironment.TabIndex = 0;
            this.gbTask_SelectEnvironment.TabStop = false;
            this.gbTask_SelectEnvironment.Text = "1. Select Environment";
            // 
            // pnlTask_SelectEnvironment
            // 
            this.pnlTask_SelectEnvironment.Location = new System.Drawing.Point(212, 40);
            this.pnlTask_SelectEnvironment.Name = "pnlTask_SelectEnvironment";
            this.pnlTask_SelectEnvironment.Size = new System.Drawing.Size(38, 20);
            this.pnlTask_SelectEnvironment.TabIndex = 1;
            // 
            // gbTask_SelectFileType
            // 
            this.gbTask_SelectFileType.Controls.Add(this.pnlTask_SelectFileType);
            this.gbTask_SelectFileType.Controls.Add(this.ucFileTypeList1);
            this.gbTask_SelectFileType.Location = new System.Drawing.Point(3, 80);
            this.gbTask_SelectFileType.Name = "gbTask_SelectFileType";
            this.gbTask_SelectFileType.Size = new System.Drawing.Size(256, 71);
            this.gbTask_SelectFileType.TabIndex = 1;
            this.gbTask_SelectFileType.TabStop = false;
            this.gbTask_SelectFileType.Text = "2. Select Manual FileType";
            // 
            // pnlTask_SelectFileType
            // 
            this.pnlTask_SelectFileType.Location = new System.Drawing.Point(212, 40);
            this.pnlTask_SelectFileType.Name = "pnlTask_SelectFileType";
            this.pnlTask_SelectFileType.Size = new System.Drawing.Size(38, 20);
            this.pnlTask_SelectFileType.TabIndex = 2;
            // 
            // btnOpenFile
            // 
            this.btnOpenFile.Location = new System.Drawing.Point(6, 19);
            this.btnOpenFile.Name = "btnOpenFile";
            this.btnOpenFile.Size = new System.Drawing.Size(200, 23);
            this.btnOpenFile.TabIndex = 2;
            this.btnOpenFile.Text = "Open File";
            this.btnOpenFile.UseVisualStyleBackColor = true;
            this.btnOpenFile.Click += new System.EventHandler(this.btnOpenFile_Click);
            // 
            // gbTask_ProcessFile
            // 
            this.gbTask_ProcessFile.Controls.Add(this.pnlTask_ProcessFile);
            this.gbTask_ProcessFile.Controls.Add(this.btnProcessFile);
            this.gbTask_ProcessFile.Location = new System.Drawing.Point(3, 275);
            this.gbTask_ProcessFile.Name = "gbTask_ProcessFile";
            this.gbTask_ProcessFile.Size = new System.Drawing.Size(256, 53);
            this.gbTask_ProcessFile.TabIndex = 3;
            this.gbTask_ProcessFile.TabStop = false;
            this.gbTask_ProcessFile.Text = "5. Process Input File";
            // 
            // pnlTask_ProcessFile
            // 
            this.pnlTask_ProcessFile.Location = new System.Drawing.Point(212, 21);
            this.pnlTask_ProcessFile.Name = "pnlTask_ProcessFile";
            this.pnlTask_ProcessFile.Size = new System.Drawing.Size(38, 20);
            this.pnlTask_ProcessFile.TabIndex = 6;
            // 
            // btnProcessFile
            // 
            this.btnProcessFile.Location = new System.Drawing.Point(6, 19);
            this.btnProcessFile.Name = "btnProcessFile";
            this.btnProcessFile.Size = new System.Drawing.Size(199, 23);
            this.btnProcessFile.TabIndex = 4;
            this.btnProcessFile.Text = "Process File";
            this.btnProcessFile.UseVisualStyleBackColor = true;
            this.btnProcessFile.Click += new System.EventHandler(this.btnProcessFile_Click);
            // 
            // dtpAsOfDate
            // 
            this.dtpAsOfDate.Location = new System.Drawing.Point(7, 19);
            this.dtpAsOfDate.Name = "dtpAsOfDate";
            this.dtpAsOfDate.Size = new System.Drawing.Size(200, 20);
            this.dtpAsOfDate.TabIndex = 3;
            this.dtpAsOfDate.Value = new System.DateTime(2012, 5, 18, 14, 7, 1, 0);
            this.dtpAsOfDate.ValueChanged += new System.EventHandler(this.dtpAsOfDate_ValueChanged);
            this.dtpAsOfDate.MouseUp += new System.Windows.Forms.MouseEventHandler(this.dtpAsOfDate_MouseUp);
            // 
            // btnSaveOutputFiles
            // 
            this.btnSaveOutputFiles.Location = new System.Drawing.Point(6, 32);
            this.btnSaveOutputFiles.Name = "btnSaveOutputFiles";
            this.btnSaveOutputFiles.Size = new System.Drawing.Size(200, 53);
            this.btnSaveOutputFiles.TabIndex = 5;
            this.btnSaveOutputFiles.Text = "Save Output Files";
            this.btnSaveOutputFiles.UseVisualStyleBackColor = true;
            this.btnSaveOutputFiles.Click += new System.EventHandler(this.btnSaveOutputFiles_Click);
            // 
            // gbTask_OpenFile
            // 
            this.gbTask_OpenFile.Controls.Add(this.pnlTask_OpenFile);
            this.gbTask_OpenFile.Controls.Add(this.btnOpenFile);
            this.gbTask_OpenFile.Location = new System.Drawing.Point(3, 157);
            this.gbTask_OpenFile.Name = "gbTask_OpenFile";
            this.gbTask_OpenFile.Size = new System.Drawing.Size(256, 52);
            this.gbTask_OpenFile.TabIndex = 4;
            this.gbTask_OpenFile.TabStop = false;
            this.gbTask_OpenFile.Text = "3. OpenFile";
            // 
            // pnlTask_OpenFile
            // 
            this.pnlTask_OpenFile.Location = new System.Drawing.Point(212, 21);
            this.pnlTask_OpenFile.Name = "pnlTask_OpenFile";
            this.pnlTask_OpenFile.Size = new System.Drawing.Size(38, 20);
            this.pnlTask_OpenFile.TabIndex = 3;
            // 
            // gbTask_SaveFiles
            // 
            this.gbTask_SaveFiles.Controls.Add(this.label2);
            this.gbTask_SaveFiles.Controls.Add(this.label1);
            this.gbTask_SaveFiles.Controls.Add(this.pnlTask_SaveControlFile);
            this.gbTask_SaveFiles.Controls.Add(this.pnlTask_SaveDetailFile);
            this.gbTask_SaveFiles.Controls.Add(this.btnSaveOutputFiles);
            this.gbTask_SaveFiles.Location = new System.Drawing.Point(3, 334);
            this.gbTask_SaveFiles.Name = "gbTask_SaveFiles";
            this.gbTask_SaveFiles.Size = new System.Drawing.Size(256, 92);
            this.gbTask_SaveFiles.TabIndex = 5;
            this.gbTask_SaveFiles.TabStop = false;
            this.gbTask_SaveFiles.Text = "6. Save Output Detail and Control Files";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(210, 52);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(40, 13);
            this.label2.TabIndex = 10;
            this.label2.Text = "Control";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(216, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(34, 13);
            this.label1.TabIndex = 9;
            this.label1.Text = "Detail";
            // 
            // pnlTask_SaveControlFile
            // 
            this.pnlTask_SaveControlFile.Location = new System.Drawing.Point(212, 68);
            this.pnlTask_SaveControlFile.Name = "pnlTask_SaveControlFile";
            this.pnlTask_SaveControlFile.Size = new System.Drawing.Size(38, 17);
            this.pnlTask_SaveControlFile.TabIndex = 8;
            // 
            // pnlTask_SaveDetailFile
            // 
            this.pnlTask_SaveDetailFile.Location = new System.Drawing.Point(212, 32);
            this.pnlTask_SaveDetailFile.Name = "pnlTask_SaveDetailFile";
            this.pnlTask_SaveDetailFile.Size = new System.Drawing.Size(38, 17);
            this.pnlTask_SaveDetailFile.TabIndex = 7;
            // 
            // gbTask_ConfirmAsOfDate
            // 
            this.gbTask_ConfirmAsOfDate.Controls.Add(this.pnlTask_ConfirmAsOfDate);
            this.gbTask_ConfirmAsOfDate.Controls.Add(this.dtpAsOfDate);
            this.gbTask_ConfirmAsOfDate.Location = new System.Drawing.Point(3, 215);
            this.gbTask_ConfirmAsOfDate.Name = "gbTask_ConfirmAsOfDate";
            this.gbTask_ConfirmAsOfDate.Size = new System.Drawing.Size(256, 54);
            this.gbTask_ConfirmAsOfDate.TabIndex = 6;
            this.gbTask_ConfirmAsOfDate.TabStop = false;
            this.gbTask_ConfirmAsOfDate.Text = "4. Confirm As of Date";
            // 
            // pnlTask_ConfirmAsOfDate
            // 
            this.pnlTask_ConfirmAsOfDate.Location = new System.Drawing.Point(212, 19);
            this.pnlTask_ConfirmAsOfDate.Name = "pnlTask_ConfirmAsOfDate";
            this.pnlTask_ConfirmAsOfDate.Size = new System.Drawing.Size(38, 20);
            this.pnlTask_ConfirmAsOfDate.TabIndex = 5;
            // 
            // gbTask_ReviewOutput
            // 
            this.gbTask_ReviewOutput.Controls.Add(this.btnOpenOutputFolder);
            this.gbTask_ReviewOutput.Location = new System.Drawing.Point(3, 432);
            this.gbTask_ReviewOutput.Name = "gbTask_ReviewOutput";
            this.gbTask_ReviewOutput.Size = new System.Drawing.Size(256, 52);
            this.gbTask_ReviewOutput.TabIndex = 7;
            this.gbTask_ReviewOutput.TabStop = false;
            this.gbTask_ReviewOutput.Text = "7. Review Output";
            // 
            // btnOpenOutputFolder
            // 
            this.btnOpenOutputFolder.Location = new System.Drawing.Point(7, 19);
            this.btnOpenOutputFolder.Name = "btnOpenOutputFolder";
            this.btnOpenOutputFolder.Size = new System.Drawing.Size(198, 23);
            this.btnOpenOutputFolder.TabIndex = 6;
            this.btnOpenOutputFolder.Text = "Open Output Folder";
            this.btnOpenOutputFolder.UseVisualStyleBackColor = true;
            this.btnOpenOutputFolder.Click += new System.EventHandler(this.btnReviewOutput_Click);
            // 
            // ucFileTypeList1
            // 
            this.ucFileTypeList1.Location = new System.Drawing.Point(6, 19);
            this.ucFileTypeList1.Name = "ucFileTypeList1";
            this.ucFileTypeList1.Size = new System.Drawing.Size(199, 45);
            this.ucFileTypeList1.TabIndex = 1;
            this.ucFileTypeList1.FileTypeSelectionChanged_Event += new SupportTools_Excel_MTreaty.User_Interface.User_Controls.ucFileTypeList.FileTypeSelectionChanged(this.ucFileTypeList1_FileTypeSelectionChanged_Event);
            // 
            // ucEnvironmentList1
            // 
            this.ucEnvironmentList1.Location = new System.Drawing.Point(6, 19);
            this.ucEnvironmentList1.Name = "ucEnvironmentList1";
            this.ucEnvironmentList1.Size = new System.Drawing.Size(199, 45);
            this.ucEnvironmentList1.TabIndex = 0;
            this.ucEnvironmentList1.EnvironmentselectionChanged_Event += new SupportTools_Excel_MTreaty.User_Interface.User_Controls.ucEnvironmentList.EnvironmentselectionChanged(this.ucEnvironmentList1_EnvironmentselectionChanged_Event);
            // 
            // btnPlayAgain
            // 
            this.btnPlayAgain.Location = new System.Drawing.Point(10, 490);
            this.btnPlayAgain.Name = "btnPlayAgain";
            this.btnPlayAgain.Size = new System.Drawing.Size(82, 23);
            this.btnPlayAgain.TabIndex = 8;
            this.btnPlayAgain.Text = "Play Again";
            this.toolTip1.SetToolTip(this.btnPlayAgain, "Reset all Controls to initial state.");
            this.btnPlayAgain.UseVisualStyleBackColor = true;
            this.btnPlayAgain.Click += new System.EventHandler(this.btnPlayAgain_Click);
            // 
            // TaskPane_MTreaty
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.btnPlayAgain);
            this.Controls.Add(this.gbTask_ReviewOutput);
            this.Controls.Add(this.gbTask_ConfirmAsOfDate);
            this.Controls.Add(this.gbTask_SaveFiles);
            this.Controls.Add(this.gbTask_OpenFile);
            this.Controls.Add(this.gbTask_ProcessFile);
            this.Controls.Add(this.gbTask_SelectFileType);
            this.Controls.Add(this.gbTask_SelectEnvironment);
            this.MinimumSize = new System.Drawing.Size(266, 0);
            this.Name = "TaskPane_MTreaty";
            this.Size = new System.Drawing.Size(266, 525);
            this.Load += new System.EventHandler(this.TaskPane_MTreaty_Load);
            this.gbTask_SelectEnvironment.ResumeLayout(false);
            this.gbTask_SelectFileType.ResumeLayout(false);
            this.gbTask_ProcessFile.ResumeLayout(false);
            this.gbTask_OpenFile.ResumeLayout(false);
            this.gbTask_SaveFiles.ResumeLayout(false);
            this.gbTask_SaveFiles.PerformLayout();
            this.gbTask_ConfirmAsOfDate.ResumeLayout(false);
            this.gbTask_ReviewOutput.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox gbTask_SelectEnvironment;
        private System.Windows.Forms.GroupBox gbTask_SelectFileType;
        private System.Windows.Forms.Button btnOpenFile;
        private System.Windows.Forms.GroupBox gbTask_ProcessFile;
        private System.Windows.Forms.Button btnProcessFile;
        private System.Windows.Forms.DateTimePicker dtpAsOfDate;
        private System.Windows.Forms.Button btnSaveOutputFiles;
        private System.Windows.Forms.GroupBox gbTask_OpenFile;
        private System.Windows.Forms.GroupBox gbTask_SaveFiles;
        private User_Controls.ucEnvironmentList ucEnvironmentList1;
        private User_Controls.ucFileTypeList ucFileTypeList1;
        private System.Windows.Forms.GroupBox gbTask_ConfirmAsOfDate;
        private System.Windows.Forms.Panel pnlTask_SelectEnvironment;
        private System.Windows.Forms.Panel pnlTask_SelectFileType;
        private System.Windows.Forms.Panel pnlTask_ProcessFile;
        private System.Windows.Forms.Panel pnlTask_OpenFile;
        private System.Windows.Forms.Panel pnlTask_SaveDetailFile;
        private System.Windows.Forms.Panel pnlTask_ConfirmAsOfDate;
        private System.Windows.Forms.Panel pnlTask_SaveControlFile;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox gbTask_ReviewOutput;
        private System.Windows.Forms.Button btnOpenOutputFolder;
        private System.Windows.Forms.Button btnPlayAgain;
        private System.Windows.Forms.ToolTip toolTip1;
        //private System.Windows.Forms.GroupBox groupBox1;
    }
}
