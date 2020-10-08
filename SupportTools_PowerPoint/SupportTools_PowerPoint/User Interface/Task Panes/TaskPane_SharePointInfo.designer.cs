namespace SupportTools_PowerPoint.User_Interface.Task_Panes
{
    partial class TaskPane_SharePointInfo
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
            this.gbListExplorer = new System.Windows.Forms.GroupBox();
            this.btnGetColumns = new System.Windows.Forms.Button();
            this.cbColumns = new System.Windows.Forms.ComboBox();
            this.btnGetListViews = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.cbViews = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.cbItems = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnGetItems = new System.Windows.Forms.Button();
            this.cbSharePointLists = new System.Windows.Forms.ComboBox();
            this.btnGetSiteLists = new System.Windows.Forms.Button();
            this.gbSharePointSite = new System.Windows.Forms.GroupBox();
            this.btnGetSitePages = new System.Windows.Forms.Button();
            this.ucSharePointSites1 = new SupportTools_PowerPoint.User_Interface.User_Controls.ucSharePointSites();
            this.label4 = new System.Windows.Forms.Label();
            this.gbListExplorer.SuspendLayout();
            this.gbSharePointSite.SuspendLayout();
            this.SuspendLayout();
            // 
            // gbListExplorer
            // 
            this.gbListExplorer.Controls.Add(this.label4);
            this.gbListExplorer.Controls.Add(this.btnGetColumns);
            this.gbListExplorer.Controls.Add(this.cbColumns);
            this.gbListExplorer.Controls.Add(this.btnGetListViews);
            this.gbListExplorer.Controls.Add(this.label3);
            this.gbListExplorer.Controls.Add(this.cbViews);
            this.gbListExplorer.Controls.Add(this.label2);
            this.gbListExplorer.Controls.Add(this.cbItems);
            this.gbListExplorer.Controls.Add(this.label1);
            this.gbListExplorer.Controls.Add(this.btnGetItems);
            this.gbListExplorer.Controls.Add(this.cbSharePointLists);
            this.gbListExplorer.Location = new System.Drawing.Point(3, 128);
            this.gbListExplorer.Name = "gbListExplorer";
            this.gbListExplorer.Size = new System.Drawing.Size(344, 270);
            this.gbListExplorer.TabIndex = 0;
            this.gbListExplorer.TabStop = false;
            this.gbListExplorer.Text = "List Explorer";
            // 
            // btnGetColumns
            // 
            this.btnGetColumns.Location = new System.Drawing.Point(258, 82);
            this.btnGetColumns.Name = "btnGetColumns";
            this.btnGetColumns.Size = new System.Drawing.Size(75, 23);
            this.btnGetColumns.TabIndex = 56;
            this.btnGetColumns.Text = "Get Columns";
            this.btnGetColumns.UseVisualStyleBackColor = true;
            // 
            // cbColumns
            // 
            this.cbColumns.FormattingEnabled = true;
            this.cbColumns.Location = new System.Drawing.Point(6, 137);
            this.cbColumns.Name = "cbColumns";
            this.cbColumns.Size = new System.Drawing.Size(246, 21);
            this.cbColumns.TabIndex = 55;
            // 
            // btnGetListViews
            // 
            this.btnGetListViews.Location = new System.Drawing.Point(258, 36);
            this.btnGetListViews.Name = "btnGetListViews";
            this.btnGetListViews.Size = new System.Drawing.Size(75, 23);
            this.btnGetListViews.TabIndex = 54;
            this.btnGetListViews.Text = "Get Views";
            this.btnGetListViews.UseVisualStyleBackColor = true;
            this.btnGetListViews.Click += new System.EventHandler(this.btnGetListViews_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 68);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(35, 13);
            this.label3.TabIndex = 53;
            this.label3.Text = "Views";
            // 
            // cbViews
            // 
            this.cbViews.FormattingEnabled = true;
            this.cbViews.Location = new System.Drawing.Point(6, 84);
            this.cbViews.Name = "cbViews";
            this.cbViews.Size = new System.Drawing.Size(246, 21);
            this.cbViews.TabIndex = 52;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 121);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(47, 13);
            this.label2.TabIndex = 51;
            this.label2.Text = "Columns";
            // 
            // cbItems
            // 
            this.cbItems.FormattingEnabled = true;
            this.cbItems.Location = new System.Drawing.Point(6, 180);
            this.cbItems.Name = "cbItems";
            this.cbItems.Size = new System.Drawing.Size(246, 21);
            this.cbItems.TabIndex = 50;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(28, 13);
            this.label1.TabIndex = 49;
            this.label1.Text = "Lists";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // btnGetItems
            // 
            this.btnGetItems.Location = new System.Drawing.Point(258, 137);
            this.btnGetItems.Name = "btnGetItems";
            this.btnGetItems.Size = new System.Drawing.Size(75, 23);
            this.btnGetItems.TabIndex = 48;
            this.btnGetItems.Text = "Get Items";
            this.btnGetItems.UseVisualStyleBackColor = true;
            // 
            // cbSharePointLists
            // 
            this.cbSharePointLists.FormattingEnabled = true;
            this.cbSharePointLists.Location = new System.Drawing.Point(6, 36);
            this.cbSharePointLists.Name = "cbSharePointLists";
            this.cbSharePointLists.Size = new System.Drawing.Size(246, 21);
            this.cbSharePointLists.TabIndex = 47;
            // 
            // btnGetSiteLists
            // 
            this.btnGetSiteLists.Location = new System.Drawing.Point(108, 70);
            this.btnGetSiteLists.Name = "btnGetSiteLists";
            this.btnGetSiteLists.Size = new System.Drawing.Size(93, 23);
            this.btnGetSiteLists.TabIndex = 45;
            this.btnGetSiteLists.Text = "Get Site Lists";
            this.btnGetSiteLists.UseVisualStyleBackColor = true;
            this.btnGetSiteLists.Click += new System.EventHandler(this.btnGetSiteLists_Click);
            // 
            // gbSharePointSite
            // 
            this.gbSharePointSite.Controls.Add(this.btnGetSitePages);
            this.gbSharePointSite.Controls.Add(this.ucSharePointSites1);
            this.gbSharePointSite.Controls.Add(this.btnGetSiteLists);
            this.gbSharePointSite.Location = new System.Drawing.Point(3, 3);
            this.gbSharePointSite.Name = "gbSharePointSite";
            this.gbSharePointSite.Size = new System.Drawing.Size(344, 119);
            this.gbSharePointSite.TabIndex = 1;
            this.gbSharePointSite.TabStop = false;
            this.gbSharePointSite.Text = "Select SharePoint Site";
            // 
            // btnGetSitePages
            // 
            this.btnGetSitePages.Location = new System.Drawing.Point(9, 70);
            this.btnGetSitePages.Name = "btnGetSitePages";
            this.btnGetSitePages.Size = new System.Drawing.Size(93, 23);
            this.btnGetSitePages.TabIndex = 48;
            this.btnGetSitePages.Text = "Get Site Pages";
            this.btnGetSitePages.UseVisualStyleBackColor = true;
            this.btnGetSitePages.Click += new System.EventHandler(this.btnGetSitePages_Click_1);
            // 
            // ucSharePointSites1
            // 
            this.ucSharePointSites1.Location = new System.Drawing.Point(6, 19);
            this.ucSharePointSites1.Name = "ucSharePointSites1";
            this.ucSharePointSites1.Size = new System.Drawing.Size(246, 45);
            this.ucSharePointSites1.TabIndex = 46;
            this.ucSharePointSites1.Url = null;
            this.ucSharePointSites1.ListElementsSelectionChanged_Event += new SupportTools_PowerPoint.User_Interface.User_Controls.ucSharePointSites.ListElementsSelectionChanged(this.ucSharePointSites1_ListElementsSelectionChanged_Event);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(6, 164);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(32, 13);
            this.label4.TabIndex = 57;
            this.label4.Text = "Items";
            // 
            // TaskPane_SharePointInfo
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.gbSharePointSite);
            this.Controls.Add(this.gbListExplorer);
            this.MinimumSize = new System.Drawing.Size(350, 0);
            this.Name = "TaskPane_SharePointInfo";
            this.Size = new System.Drawing.Size(350, 646);
            this.Load += new System.EventHandler(this.TaskPane_SharePointInfo_Load);
            this.gbListExplorer.ResumeLayout(false);
            this.gbListExplorer.PerformLayout();
            this.gbSharePointSite.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox gbListExplorer;
        private System.Windows.Forms.Button btnGetColumns;
        private System.Windows.Forms.ComboBox cbColumns;
        private System.Windows.Forms.Button btnGetListViews;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cbViews;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cbItems;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnGetItems;
        private System.Windows.Forms.ComboBox cbSharePointLists;
        private System.Windows.Forms.Button btnGetSiteLists;
        private User_Controls.ucSharePointSites ucSharePointSites1;
        private System.Windows.Forms.GroupBox gbSharePointSite;
        private System.Windows.Forms.Button btnGetSitePages;
        private System.Windows.Forms.Label label4;
    }
}
