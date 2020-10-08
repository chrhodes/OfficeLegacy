namespace SupportTools_PowerPoint.User_Interface.User_Controls
{
    partial class ucStateList
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
            this.components = new System.ComponentModel.Container();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.lblLabel = new System.Windows.Forms.Label();
            this.cbComboBox = new System.Windows.Forms.ComboBox();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.SuspendLayout();
            // 
            // lblLabel
            // 
            this.lblLabel.AutoSize = true;
            this.lblLabel.Location = new System.Drawing.Point(1, 5);
            this.lblLabel.Name = "lblLabel";
            this.lblLabel.Size = new System.Drawing.Size(37, 13);
            this.lblLabel.TabIndex = 12;
            this.lblLabel.Text = "States";
            this.toolTip1.SetToolTip(this.lblLabel, "Double Click to Load New List from File");
            this.lblLabel.DoubleClick += new System.EventHandler(this.lblEnvironments_DoubleClick);
            // 
            // cbComboBox
            // 
            this.cbComboBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cbComboBox.DisplayMember = "StateName";
            this.cbComboBox.FormattingEnabled = true;
            this.cbComboBox.Location = new System.Drawing.Point(1, 21);
            this.cbComboBox.Name = "cbComboBox";
            this.cbComboBox.Size = new System.Drawing.Size(295, 21);
            this.cbComboBox.TabIndex = 13;
            this.cbComboBox.ValueMember = "ShapeName";
            this.cbComboBox.SelectedIndexChanged += new System.EventHandler(this.cbComboBox_SelectedIndexChanged);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // ucStateList
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.cbComboBox);
            this.Controls.Add(this.lblLabel);
            this.Name = "ucStateList";
            this.Size = new System.Drawing.Size(297, 45);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Label lblLabel;
        private System.Windows.Forms.ComboBox cbComboBox;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
    }
}
