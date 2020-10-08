namespace SupportTools_Outlook.User_Interface.Task_Panes
{
    partial class TaskPane_AppUtil
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
            if (disposing && (components != null))
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
            this.btnAddRules = new System.Windows.Forms.Button();
            this.btnListFolders = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnAddRules
            // 
            this.btnAddRules.Location = new System.Drawing.Point(58, 26);
            this.btnAddRules.Name = "btnAddRules";
            this.btnAddRules.Size = new System.Drawing.Size(75, 23);
            this.btnAddRules.TabIndex = 0;
            this.btnAddRules.Text = "Add Rules";
            this.btnAddRules.UseVisualStyleBackColor = true;
            this.btnAddRules.Click += new System.EventHandler(this.btnAddRules_Click);
            // 
            // btnListFolders
            // 
            this.btnListFolders.Location = new System.Drawing.Point(58, 73);
            this.btnListFolders.Name = "btnListFolders";
            this.btnListFolders.Size = new System.Drawing.Size(75, 23);
            this.btnListFolders.TabIndex = 1;
            this.btnListFolders.Text = "List Folders";
            this.btnListFolders.UseVisualStyleBackColor = true;
            this.btnListFolders.Click += new System.EventHandler(this.btnListFolders_Click);
            // 
            // TaskPane_AppUtil
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.btnListFolders);
            this.Controls.Add(this.btnAddRules);
            this.Name = "TaskPane_AppUtil";
            this.Size = new System.Drawing.Size(200, 400);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnAddRules;
        private System.Windows.Forms.Button btnListFolders;
    }
}
