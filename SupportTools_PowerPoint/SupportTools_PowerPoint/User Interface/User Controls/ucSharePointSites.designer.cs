namespace SupportTools_PowerPoint.User_Interface.User_Controls
{
    partial class ucSharePointSites
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
            this.lblListName = new System.Windows.Forms.Label();
            this.cbListElements = new System.Windows.Forms.ComboBox();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.SuspendLayout();
            // 
            // lblListName
            // 
            this.lblListName.AutoSize = true;
            this.lblListName.Location = new System.Drawing.Point(1, 5);
            this.lblListName.Name = "lblListName";
            this.lblListName.Size = new System.Drawing.Size(85, 13);
            this.lblListName.TabIndex = 12;
            this.lblListName.Text = "SharePoint Sites";
            this.toolTip1.SetToolTip(this.lblListName, "Double Click to Load New List from File");
            this.lblListName.DoubleClick += new System.EventHandler(this.lblListName_DoubleClick);
            // 
            // cbListElements
            // 
            this.cbListElements.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cbListElements.DisplayMember = "Url";
            this.cbListElements.FormattingEnabled = true;
            this.cbListElements.Location = new System.Drawing.Point(1, 21);
            this.cbListElements.Name = "cbListElements";
            this.cbListElements.Size = new System.Drawing.Size(295, 21);
            this.cbListElements.TabIndex = 13;
            this.cbListElements.ValueMember = "Url";
            this.cbListElements.SelectedIndexChanged += new System.EventHandler(this.cbListElements_SelectedIndexChanged);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // ucSharePointSites
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.cbListElements);
            this.Controls.Add(this.lblListName);
            this.Name = "ucSharePointSites";
            this.Size = new System.Drawing.Size(297, 45);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Label lblListName;
        private System.Windows.Forms.ComboBox cbListElements;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
    }
}
