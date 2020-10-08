namespace SupportTools_PowerPoint.User_Interface.User_Controls
{
    partial class ucColorList
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
            this.lblColors = new System.Windows.Forms.Label();
            this.cbColors = new System.Windows.Forms.ComboBox();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.SuspendLayout();
            // 
            // lblColors
            // 
            this.lblColors.AutoSize = true;
            this.lblColors.Location = new System.Drawing.Point(1, 4);
            this.lblColors.Name = "lblColors";
            this.lblColors.Size = new System.Drawing.Size(36, 13);
            this.lblColors.TabIndex = 12;
            this.lblColors.Text = "Colors";
            this.toolTip1.SetToolTip(this.lblColors, "Double Click to Load New List from File");
            this.lblColors.DoubleClick += new System.EventHandler(this.lblColors_DoubleClick);
            // 
            // cbColors
            // 
            this.cbColors.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cbColors.DisplayMember = "ColorName";
            this.cbColors.FormattingEnabled = true;
            this.cbColors.Location = new System.Drawing.Point(1, 21);
            this.cbColors.Name = "cbColors";
            this.cbColors.Size = new System.Drawing.Size(295, 21);
            this.cbColors.TabIndex = 13;
            this.cbColors.ValueMember = "Value";
            this.cbColors.SelectedIndexChanged += new System.EventHandler(this.cbColors_SelectedIndexChanged);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // ucColorList
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.cbColors);
            this.Controls.Add(this.lblColors);
            this.Name = "ucColorList";
            this.Size = new System.Drawing.Size(297, 45);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Label lblColors;
        private System.Windows.Forms.ComboBox cbColors;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
    }
}
