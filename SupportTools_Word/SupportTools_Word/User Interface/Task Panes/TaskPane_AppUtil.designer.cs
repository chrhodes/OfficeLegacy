namespace SupportTools_Word.User_Interface.Task_Panes
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
            this.btnAddFooter = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnAddFooter
            // 
            this.btnAddFooter.BackgroundImage = global::SupportTools_Word.Properties.Resources.add_footer;
            this.btnAddFooter.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnAddFooter.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnAddFooter.Location = new System.Drawing.Point(12, 15);
            this.btnAddFooter.Name = "btnAddFooter";
            this.btnAddFooter.Size = new System.Drawing.Size(83, 41);
            this.btnAddFooter.TabIndex = 30;
            this.btnAddFooter.Text = "Add\r\nFooter";
            this.btnAddFooter.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnAddFooter.UseVisualStyleBackColor = true;
            this.btnAddFooter.Click += new System.EventHandler(this.btnAddFooter_Click);
            // 
            // TaskPane_AppUtil
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.btnAddFooter);
            this.Name = "TaskPane_AppUtil";
            this.Size = new System.Drawing.Size(200, 400);
            this.ResumeLayout(false);

        }

        #endregion

        internal System.Windows.Forms.Button btnAddFooter;
    }
}
