namespace SupportTools_Excel.User_Interface.User_Controls
{
    partial class ucDBInstanceList
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
            this.cbInstance = new System.Windows.Forms.ComboBox();
            this.lblInstance = new System.Windows.Forms.Label();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.SuspendLayout();
            // 
            // cbInstance
            // 
            this.cbInstance.DisplayMember = "ExpandedName";
            this.cbInstance.FormattingEnabled = true;
            this.cbInstance.Location = new System.Drawing.Point(3, 16);
            this.cbInstance.Name = "cbInstance";
            this.cbInstance.Size = new System.Drawing.Size(295, 21);
            this.cbInstance.TabIndex = 11;
            this.cbInstance.ValueMember = "FullName";
            this.cbInstance.SelectedIndexChanged += new System.EventHandler(this.cbInstance_SelectedIndexChanged);
            // 
            // lblInstance
            // 
            this.lblInstance.AutoSize = true;
            this.lblInstance.Location = new System.Drawing.Point(3, 0);
            this.lblInstance.Name = "lblInstance";
            this.lblInstance.Size = new System.Drawing.Size(48, 13);
            this.lblInstance.TabIndex = 10;
            this.lblInstance.Text = "Instance";
            this.toolTip1.SetToolTip(this.lblInstance, "Double Click to Load New Instance List");
            this.lblInstance.Click += new System.EventHandler(this.lblInstance_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // ucDBInstanceList
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlDark;
            this.Controls.Add(this.cbInstance);
            this.Controls.Add(this.lblInstance);
            this.Name = "ucDBInstanceList";
            this.Size = new System.Drawing.Size(301, 43);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox cbInstance;
        private System.Windows.Forms.Label lblInstance;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
    }
}
