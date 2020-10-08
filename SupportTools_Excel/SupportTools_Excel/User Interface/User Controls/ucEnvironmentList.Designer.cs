namespace SupportTools_Excel.User_Interface.User_Controls
{
    partial class ucEnvironmentList
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
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.lblEnvironments = new System.Windows.Forms.Label();
            this.cbEnvironments = new System.Windows.Forms.ComboBox();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.SuspendLayout();
            // 
            // lblEnvironments
            // 
            this.lblEnvironments.AutoSize = true;
            this.lblEnvironments.Location = new System.Drawing.Point(1, 5);
            this.lblEnvironments.Name = "lblEnvironments";
            this.lblEnvironments.Size = new System.Drawing.Size(71, 13);
            this.lblEnvironments.TabIndex = 12;
            this.lblEnvironments.Text = "Environments";
            this.toolTip1.SetToolTip(this.lblEnvironments, "Double Click to Load New List from File");
            this.lblEnvironments.DoubleClick += new System.EventHandler(this.lblEnvironments_DoubleClick);
            // 
            // cbEnvironments
            // 
            this.cbEnvironments.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cbEnvironments.DisplayMember = "Name";
            this.cbEnvironments.FormattingEnabled = true;
            this.cbEnvironments.Location = new System.Drawing.Point(1, 21);
            this.cbEnvironments.Name = "cbEnvironments";
            this.cbEnvironments.Size = new System.Drawing.Size(295, 21);
            this.cbEnvironments.TabIndex = 13;
            this.cbEnvironments.ValueMember = "Path";
            this.cbEnvironments.SelectedIndexChanged += new System.EventHandler(this.cbEnvironments_SelectedIndexChanged);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // ucEnvironmentList
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.cbEnvironments);
            this.Controls.Add(this.lblEnvironments);
            this.Name = "ucEnvironmentList";
            this.Size = new System.Drawing.Size(297, 45);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Label lblEnvironments;
        private System.Windows.Forms.ComboBox cbEnvironments;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
    }
}
