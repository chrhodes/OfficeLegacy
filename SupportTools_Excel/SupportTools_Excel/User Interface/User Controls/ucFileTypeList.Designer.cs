namespace SupportTools_Excel.User_Interface.User_Controls
{
    partial class ucFileTypeList
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
            this.lblFileTypes = new System.Windows.Forms.Label();
            this.cbFileTypes = new System.Windows.Forms.ComboBox();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.SuspendLayout();
            // 
            // lblFileTypes
            // 
            this.lblFileTypes.AutoSize = true;
            this.lblFileTypes.Location = new System.Drawing.Point(1, 5);
            this.lblFileTypes.Name = "lblFileTypes";
            this.lblFileTypes.Size = new System.Drawing.Size(52, 13);
            this.lblFileTypes.TabIndex = 12;
            this.lblFileTypes.Text = "FileTypes";
            this.toolTip1.SetToolTip(this.lblFileTypes, "Double Click to Load New Instance List");
            this.lblFileTypes.DoubleClick += new System.EventHandler(this.lblFileTypes_DoubleClick);
            // 
            // cbFileTypes
            // 
            this.cbFileTypes.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cbFileTypes.DisplayMember = "DisplayValue";
            this.cbFileTypes.FormattingEnabled = true;
            this.cbFileTypes.Location = new System.Drawing.Point(1, 21);
            this.cbFileTypes.Name = "cbFileTypes";
            this.cbFileTypes.Size = new System.Drawing.Size(295, 21);
            this.cbFileTypes.TabIndex = 13;
            this.cbFileTypes.ValueMember = "RelativePath";
            this.cbFileTypes.SelectedIndexChanged += new System.EventHandler(this.cbFileTypes_SelectedIndexChanged);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // ucFileTypeList
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.cbFileTypes);
            this.Controls.Add(this.lblFileTypes);
            this.Name = "ucFileTypeList";
            this.Size = new System.Drawing.Size(297, 45);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Label lblFileTypes;
        private System.Windows.Forms.ComboBox cbFileTypes;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
    }
}
