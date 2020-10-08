namespace Compliance_Office2010Addin_Word.User_Interface.Task_Panes
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
            this.components = new System.ComponentModel.Container();
            this.gbCreateIndexStyles = new System.Windows.Forms.GroupBox();
            this.btnCreateIndexStyles = new System.Windows.Forms.Button();
            this.gbMarkIndexWords = new System.Windows.Forms.GroupBox();
            this.btnMarkIndexWords = new System.Windows.Forms.Button();
            this.gbTagIndexWords = new System.Windows.Forms.GroupBox();
            this.btnTagIndexHeadingStyle = new System.Windows.Forms.Button();
            this.btnFindIndexHeadingStyle = new System.Windows.Forms.Button();
            this.btnUpdateIndex = new System.Windows.Forms.Button();
            this.btnTagIndexWordStyle = new System.Windows.Forms.Button();
            this.btnFindIndexWordStyle = new System.Windows.Forms.Button();
            this.gbImproveReadability = new System.Windows.Forms.GroupBox();
            this.txtReadabilityStatistics = new System.Windows.Forms.TextBox();
            this.btnDisplayReadability = new System.Windows.Forms.Button();
            this.ckIndexWordsOnly = new System.Windows.Forms.CheckBox();
            this.txtReplacementWord = new System.Windows.Forms.TextBox();
            this.btnSaveReplacementWords = new System.Windows.Forms.Button();
            this.lblReplacementWord = new System.Windows.Forms.Label();
            this.btnZapReplacementWords = new System.Windows.Forms.Button();
            this.gbSpellCheck = new System.Windows.Forms.GroupBox();
            this.btnResetCheck = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.btnClearStylesFromWords = new System.Windows.Forms.Button();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.btnDeleteIndexFields = new System.Windows.Forms.Button();
            this.btnFindIndexFields = new System.Windows.Forms.Button();
            this.gbCreateIndexStyles.SuspendLayout();
            this.gbMarkIndexWords.SuspendLayout();
            this.gbTagIndexWords.SuspendLayout();
            this.gbImproveReadability.SuspendLayout();
            this.gbSpellCheck.SuspendLayout();
            this.SuspendLayout();
            // 
            // gbCreateIndexStyles
            // 
            this.gbCreateIndexStyles.Controls.Add(this.btnCreateIndexStyles);
            this.gbCreateIndexStyles.Location = new System.Drawing.Point(15, 3);
            this.gbCreateIndexStyles.Name = "gbCreateIndexStyles";
            this.gbCreateIndexStyles.Size = new System.Drawing.Size(289, 45);
            this.gbCreateIndexStyles.TabIndex = 11;
            this.gbCreateIndexStyles.TabStop = false;
            this.gbCreateIndexStyles.Text = "Create Index Styles";
            // 
            // btnCreateIndexStyles
            // 
            this.btnCreateIndexStyles.Location = new System.Drawing.Point(6, 16);
            this.btnCreateIndexStyles.Name = "btnCreateIndexStyles";
            this.btnCreateIndexStyles.Size = new System.Drawing.Size(277, 23);
            this.btnCreateIndexStyles.TabIndex = 9;
            this.btnCreateIndexStyles.Text = "Create Index Styles";
            this.btnCreateIndexStyles.UseVisualStyleBackColor = true;
            this.btnCreateIndexStyles.Click += new System.EventHandler(this.btnCreateIndexStyles_Click);
            // 
            // gbMarkIndexWords
            // 
            this.gbMarkIndexWords.Controls.Add(this.btnClearStylesFromWords);
            this.gbMarkIndexWords.Controls.Add(this.btnMarkIndexWords);
            this.gbMarkIndexWords.Location = new System.Drawing.Point(15, 50);
            this.gbMarkIndexWords.Name = "gbMarkIndexWords";
            this.gbMarkIndexWords.Size = new System.Drawing.Size(283, 49);
            this.gbMarkIndexWords.TabIndex = 13;
            this.gbMarkIndexWords.TabStop = false;
            this.gbMarkIndexWords.Text = "Mark Index Words";
            // 
            // btnMarkIndexWords
            // 
            this.btnMarkIndexWords.Location = new System.Drawing.Point(6, 19);
            this.btnMarkIndexWords.Name = "btnMarkIndexWords";
            this.btnMarkIndexWords.Size = new System.Drawing.Size(131, 23);
            this.btnMarkIndexWords.TabIndex = 8;
            this.btnMarkIndexWords.Text = "Mark Index Words";
            this.btnMarkIndexWords.UseVisualStyleBackColor = true;
            this.btnMarkIndexWords.Click += new System.EventHandler(this.btnMarkIndexWords_Click);
            // 
            // gbTagIndexWords
            // 
            this.gbTagIndexWords.Controls.Add(this.btnFindIndexFields);
            this.gbTagIndexWords.Controls.Add(this.btnDeleteIndexFields);
            this.gbTagIndexWords.Controls.Add(this.btnTagIndexHeadingStyle);
            this.gbTagIndexWords.Controls.Add(this.btnFindIndexHeadingStyle);
            this.gbTagIndexWords.Controls.Add(this.btnUpdateIndex);
            this.gbTagIndexWords.Controls.Add(this.btnTagIndexWordStyle);
            this.gbTagIndexWords.Controls.Add(this.btnFindIndexWordStyle);
            this.gbTagIndexWords.Location = new System.Drawing.Point(15, 101);
            this.gbTagIndexWords.Name = "gbTagIndexWords";
            this.gbTagIndexWords.Size = new System.Drawing.Size(283, 156);
            this.gbTagIndexWords.TabIndex = 14;
            this.gbTagIndexWords.TabStop = false;
            this.gbTagIndexWords.Text = "Tag Index Words";
            // 
            // btnTagIndexHeadingStyle
            // 
            this.btnTagIndexHeadingStyle.Location = new System.Drawing.Point(143, 48);
            this.btnTagIndexHeadingStyle.Name = "btnTagIndexHeadingStyle";
            this.btnTagIndexHeadingStyle.Size = new System.Drawing.Size(131, 23);
            this.btnTagIndexHeadingStyle.TabIndex = 10;
            this.btnTagIndexHeadingStyle.Text = "Tag IndexHeading Style";
            this.btnTagIndexHeadingStyle.UseVisualStyleBackColor = true;
            this.btnTagIndexHeadingStyle.Click += new System.EventHandler(this.btnTagIndexHeadingStyleWords_Click);
            // 
            // btnFindIndexHeadingStyle
            // 
            this.btnFindIndexHeadingStyle.Location = new System.Drawing.Point(6, 48);
            this.btnFindIndexHeadingStyle.Name = "btnFindIndexHeadingStyle";
            this.btnFindIndexHeadingStyle.Size = new System.Drawing.Size(131, 23);
            this.btnFindIndexHeadingStyle.TabIndex = 9;
            this.btnFindIndexHeadingStyle.Text = "Find IndexHeading Style";
            this.btnFindIndexHeadingStyle.UseVisualStyleBackColor = true;
            this.btnFindIndexHeadingStyle.Click += new System.EventHandler(this.btnFindIndexHeadingStyle_Click);
            // 
            // btnUpdateIndex
            // 
            this.btnUpdateIndex.Location = new System.Drawing.Point(6, 127);
            this.btnUpdateIndex.Name = "btnUpdateIndex";
            this.btnUpdateIndex.Size = new System.Drawing.Size(268, 23);
            this.btnUpdateIndex.TabIndex = 2;
            this.btnUpdateIndex.Text = "Update Index";
            this.btnUpdateIndex.UseVisualStyleBackColor = true;
            this.btnUpdateIndex.Click += new System.EventHandler(this.btnUpdateIndex_Click);
            // 
            // btnTagIndexWordStyle
            // 
            this.btnTagIndexWordStyle.Location = new System.Drawing.Point(143, 19);
            this.btnTagIndexWordStyle.Name = "btnTagIndexWordStyle";
            this.btnTagIndexWordStyle.Size = new System.Drawing.Size(131, 23);
            this.btnTagIndexWordStyle.TabIndex = 1;
            this.btnTagIndexWordStyle.Text = "Tag IndexWord Style";
            this.btnTagIndexWordStyle.UseVisualStyleBackColor = true;
            this.btnTagIndexWordStyle.Click += new System.EventHandler(this.btnTagIndexWordStyleWords_Click);
            // 
            // btnFindIndexWordStyle
            // 
            this.btnFindIndexWordStyle.Location = new System.Drawing.Point(6, 18);
            this.btnFindIndexWordStyle.Name = "btnFindIndexWordStyle";
            this.btnFindIndexWordStyle.Size = new System.Drawing.Size(131, 23);
            this.btnFindIndexWordStyle.TabIndex = 0;
            this.btnFindIndexWordStyle.Text = "Find IndexWord Style";
            this.btnFindIndexWordStyle.UseVisualStyleBackColor = true;
            this.btnFindIndexWordStyle.Click += new System.EventHandler(this.btnFindIndexWordStyle_Click);
            // 
            // gbImproveReadability
            // 
            this.gbImproveReadability.Controls.Add(this.txtReadabilityStatistics);
            this.gbImproveReadability.Controls.Add(this.btnDisplayReadability);
            this.gbImproveReadability.Controls.Add(this.ckIndexWordsOnly);
            this.gbImproveReadability.Controls.Add(this.txtReplacementWord);
            this.gbImproveReadability.Controls.Add(this.btnSaveReplacementWords);
            this.gbImproveReadability.Controls.Add(this.lblReplacementWord);
            this.gbImproveReadability.Controls.Add(this.btnZapReplacementWords);
            this.gbImproveReadability.Location = new System.Drawing.Point(15, 259);
            this.gbImproveReadability.Name = "gbImproveReadability";
            this.gbImproveReadability.Size = new System.Drawing.Size(200, 282);
            this.gbImproveReadability.TabIndex = 15;
            this.gbImproveReadability.TabStop = false;
            this.gbImproveReadability.Text = "Improve Readability";
            // 
            // txtReadabilityStatistics
            // 
            this.txtReadabilityStatistics.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtReadabilityStatistics.Location = new System.Drawing.Point(10, 137);
            this.txtReadabilityStatistics.Multiline = true;
            this.txtReadabilityStatistics.Name = "txtReadabilityStatistics";
            this.txtReadabilityStatistics.Size = new System.Drawing.Size(179, 140);
            this.txtReadabilityStatistics.TabIndex = 9;
            // 
            // btnDisplayReadability
            // 
            this.btnDisplayReadability.Location = new System.Drawing.Point(15, 111);
            this.btnDisplayReadability.Name = "btnDisplayReadability";
            this.btnDisplayReadability.Size = new System.Drawing.Size(171, 23);
            this.btnDisplayReadability.TabIndex = 8;
            this.btnDisplayReadability.Text = "Display Readability Statistics";
            this.btnDisplayReadability.UseVisualStyleBackColor = true;
            this.btnDisplayReadability.Click += new System.EventHandler(this.btnDisplayReadability_Click);
            // 
            // ckIndexWordsOnly
            // 
            this.ckIndexWordsOnly.AutoSize = true;
            this.ckIndexWordsOnly.Location = new System.Drawing.Point(15, 92);
            this.ckIndexWordsOnly.Name = "ckIndexWordsOnly";
            this.ckIndexWordsOnly.Size = new System.Drawing.Size(110, 17);
            this.ckIndexWordsOnly.TabIndex = 6;
            this.ckIndexWordsOnly.Text = "Index Words Only";
            this.ckIndexWordsOnly.UseVisualStyleBackColor = true;
            // 
            // txtReplacementWord
            // 
            this.txtReplacementWord.Location = new System.Drawing.Point(117, 42);
            this.txtReplacementWord.Name = "txtReplacementWord";
            this.txtReplacementWord.Size = new System.Drawing.Size(69, 20);
            this.txtReplacementWord.TabIndex = 5;
            this.txtReplacementWord.Text = " ";
            // 
            // btnSaveReplacementWords
            // 
            this.btnSaveReplacementWords.Location = new System.Drawing.Point(15, 16);
            this.btnSaveReplacementWords.Name = "btnSaveReplacementWords";
            this.btnSaveReplacementWords.Size = new System.Drawing.Size(171, 23);
            this.btnSaveReplacementWords.TabIndex = 2;
            this.btnSaveReplacementWords.Text = "Save Replacement Words";
            this.btnSaveReplacementWords.UseVisualStyleBackColor = true;
            this.btnSaveReplacementWords.Click += new System.EventHandler(this.btnSaveReplacementWords_Click);
            // 
            // lblReplacementWord
            // 
            this.lblReplacementWord.AutoSize = true;
            this.lblReplacementWord.Location = new System.Drawing.Point(12, 45);
            this.lblReplacementWord.Name = "lblReplacementWord";
            this.lblReplacementWord.Size = new System.Drawing.Size(99, 13);
            this.lblReplacementWord.TabIndex = 4;
            this.lblReplacementWord.Text = "Replacement Word";
            // 
            // btnZapReplacementWords
            // 
            this.btnZapReplacementWords.Location = new System.Drawing.Point(15, 65);
            this.btnZapReplacementWords.Name = "btnZapReplacementWords";
            this.btnZapReplacementWords.Size = new System.Drawing.Size(171, 23);
            this.btnZapReplacementWords.TabIndex = 3;
            this.btnZapReplacementWords.Text = "Zap Replacement Words";
            this.btnZapReplacementWords.UseVisualStyleBackColor = true;
            this.btnZapReplacementWords.Click += new System.EventHandler(this.btnZapReplacementWords_Click);
            // 
            // gbSpellCheck
            // 
            this.gbSpellCheck.Controls.Add(this.btnResetCheck);
            this.gbSpellCheck.Location = new System.Drawing.Point(15, 543);
            this.gbSpellCheck.Name = "gbSpellCheck";
            this.gbSpellCheck.Size = new System.Drawing.Size(200, 50);
            this.gbSpellCheck.TabIndex = 16;
            this.gbSpellCheck.TabStop = false;
            this.gbSpellCheck.Text = "Spell Check";
            // 
            // btnResetCheck
            // 
            this.btnResetCheck.Location = new System.Drawing.Point(15, 19);
            this.btnResetCheck.Name = "btnResetCheck";
            this.btnResetCheck.Size = new System.Drawing.Size(171, 23);
            this.btnResetCheck.TabIndex = 2;
            this.btnResetCheck.Text = "Reset Spelling and Grammer";
            this.btnResetCheck.UseVisualStyleBackColor = true;
            this.btnResetCheck.Click += new System.EventHandler(this.btnResetCheck_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // btnClearStylesFromWords
            // 
            this.btnClearStylesFromWords.Location = new System.Drawing.Point(143, 19);
            this.btnClearStylesFromWords.Name = "btnClearStylesFromWords";
            this.btnClearStylesFromWords.Size = new System.Drawing.Size(131, 23);
            this.btnClearStylesFromWords.TabIndex = 9;
            this.btnClearStylesFromWords.Text = "Clear Styles from Words";
            this.toolTip1.SetToolTip(this.btnClearStylesFromWords, "Clear Styles from Words");
            this.btnClearStylesFromWords.UseVisualStyleBackColor = true;
            this.btnClearStylesFromWords.Click += new System.EventHandler(this.btnClearStylesFromWords_Click);
            // 
            // btnDeleteIndexFields
            // 
            this.btnDeleteIndexFields.Location = new System.Drawing.Point(143, 77);
            this.btnDeleteIndexFields.Name = "btnDeleteIndexFields";
            this.btnDeleteIndexFields.Size = new System.Drawing.Size(131, 23);
            this.btnDeleteIndexFields.TabIndex = 11;
            this.btnDeleteIndexFields.Text = "Delete Index Fields";
            this.btnDeleteIndexFields.UseVisualStyleBackColor = true;
            this.btnDeleteIndexFields.Click += new System.EventHandler(this.btnDeleteIndexFields_Click);
            // 
            // btnFindIndexFields
            // 
            this.btnFindIndexFields.Location = new System.Drawing.Point(6, 77);
            this.btnFindIndexFields.Name = "btnFindIndexFields";
            this.btnFindIndexFields.Size = new System.Drawing.Size(131, 23);
            this.btnFindIndexFields.TabIndex = 12;
            this.btnFindIndexFields.Text = "Find Index Fields";
            this.btnFindIndexFields.UseVisualStyleBackColor = true;
            this.btnFindIndexFields.Click += new System.EventHandler(this.btnFindIndexFields_Click);
            // 
            // TaskPane_ComplianceUtil
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.gbSpellCheck);
            this.Controls.Add(this.gbImproveReadability);
            this.Controls.Add(this.gbTagIndexWords);
            this.Controls.Add(this.gbMarkIndexWords);
            this.Controls.Add(this.gbCreateIndexStyles);
            this.Name = "TaskPane_ComplianceUtil";
            this.Size = new System.Drawing.Size(311, 600);
            this.gbCreateIndexStyles.ResumeLayout(false);
            this.gbMarkIndexWords.ResumeLayout(false);
            this.gbTagIndexWords.ResumeLayout(false);
            this.gbImproveReadability.ResumeLayout(false);
            this.gbImproveReadability.PerformLayout();
            this.gbSpellCheck.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        internal System.Windows.Forms.GroupBox gbCreateIndexStyles;
        internal System.Windows.Forms.Button btnCreateIndexStyles;
        internal System.Windows.Forms.GroupBox gbMarkIndexWords;
        internal System.Windows.Forms.Button btnMarkIndexWords;
        internal System.Windows.Forms.GroupBox gbTagIndexWords;
        internal System.Windows.Forms.Button btnTagIndexHeadingStyle;
        internal System.Windows.Forms.Button btnFindIndexHeadingStyle;
        internal System.Windows.Forms.Button btnUpdateIndex;
        internal System.Windows.Forms.Button btnTagIndexWordStyle;
        internal System.Windows.Forms.Button btnFindIndexWordStyle;
        internal System.Windows.Forms.GroupBox gbImproveReadability;
        internal System.Windows.Forms.TextBox txtReadabilityStatistics;
        internal System.Windows.Forms.Button btnDisplayReadability;
        internal System.Windows.Forms.CheckBox ckIndexWordsOnly;
        internal System.Windows.Forms.TextBox txtReplacementWord;
        internal System.Windows.Forms.Button btnSaveReplacementWords;
        internal System.Windows.Forms.Label lblReplacementWord;
        internal System.Windows.Forms.Button btnZapReplacementWords;
        internal System.Windows.Forms.GroupBox gbSpellCheck;
        internal System.Windows.Forms.Button btnResetCheck;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        internal System.Windows.Forms.Button btnClearStylesFromWords;
        private System.Windows.Forms.ToolTip toolTip1;
        internal System.Windows.Forms.Button btnDeleteIndexFields;
        internal System.Windows.Forms.Button btnFindIndexFields;
    }
}
