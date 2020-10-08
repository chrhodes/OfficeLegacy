namespace SupportTools_PowerPoint.User_Interface.Task_Panes
{
    partial class TaskPane_ComplianceUtil
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
            this.colorDialog1 = new System.Windows.Forms.ColorDialog();
            this.btnUpdateStatesFromFile = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.cbSourceFiles = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // btnUpdateStatesFromFile
            // 
            this.btnUpdateStatesFromFile.Location = new System.Drawing.Point(20, 138);
            this.btnUpdateStatesFromFile.Name = "btnUpdateStatesFromFile";
            this.btnUpdateStatesFromFile.Size = new System.Drawing.Size(201, 25);
            this.btnUpdateStatesFromFile.TabIndex = 29;
            this.btnUpdateStatesFromFile.Text = "Update States From File";
            this.btnUpdateStatesFromFile.UseVisualStyleBackColor = true;
            this.btnUpdateStatesFromFile.Click += new System.EventHandler(this.btnUpdateStatesFromFile_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(13, 603);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(201, 25);
            this.button1.TabIndex = 30;
            this.button1.Text = "Update States From SharePoint List";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // cbSourceFiles
            // 
            this.cbSourceFiles.FormattingEnabled = true;
            this.cbSourceFiles.Location = new System.Drawing.Point(20, 100);
            this.cbSourceFiles.Name = "cbSourceFiles";
            this.cbSourceFiles.Size = new System.Drawing.Size(201, 21);
            this.cbSourceFiles.TabIndex = 31;
            // 
            // TaskPane_ComplianceUtil
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.cbSourceFiles);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.btnUpdateStatesFromFile);
            this.MinimumSize = new System.Drawing.Size(350, 0);
            this.Name = "TaskPane_ComplianceUtil";
            this.Size = new System.Drawing.Size(350, 660);
            this.Load += new System.EventHandler(this.TaskPane_ComplianceUtil_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ColorDialog colorDialog1;
        private System.Windows.Forms.Button btnUpdateStatesFromFile;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ComboBox cbSourceFiles;
    }
}
