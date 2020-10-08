using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SupportTools_Excel.User_Interface.Forms
{
    public partial class frmExcel_FolderMaps : Form
    {
        Boolean _cancel = false;

        public class RegEx
        {
            // SharePoint Folder/File/Document Libraries may not contain any of the following characters
            //   / \ : * ? " < > | <TAB> { } % ~ &
            // nor may they end in periods or contain embedded double periods.
            // The following regular expressions capture these rules.  

            internal const string cIllegalFileCharacters = "[/\\\\:\\*\\?\"<>\\|#\\{}%~&]|\\.\\.";

            internal const string cIllegalFolderCharacters = "[:\\*\\?\"<>\\|#\\{}%~&]";

        }

        #region Initialization
        
        public frmExcel_FolderMaps()
        {
            InitializeComponent();
        }

        #endregion

        #region Event Handlers

        private void btnSelectFolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog1.ShowNewFolderButton = false;

            if (txtStartingFolder.Text.Length > 0)
            {
            	FolderBrowserDialog1.SelectedPath = txtStartingFolder.Text;
            }

            if (FolderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
            	txtStartingFolder.Text = FolderBrowserDialog1.SelectedPath;
            }
        }

        private void chkColorCodeDates_CheckedChanged(object sender, EventArgs e)
        {
            if (chkColorCodeDates.Checked)
            {
            	pnlDefaultColor.Enabled = true;
                pnlMonthCreatedColor.Enabled = true;
                pnlMonthAccessedColor.Enabled =  true;
                pnlMonthWrittenColor.Enabled =  true;

                txtMonthsSinceCreated.Enabled = true;
                txtMonthsSinceWritten.Enabled =  true;
                txtMonthsSinceAccessed.Enabled =  true;

                spnMonthsSinceCreated.Enabled =  true;
                spnMonthsSinceWritten.Enabled =  true;
                spnMonthsSinceAccessed.Enabled =  true;
            }
            else
            {
            	pnlDefaultColor.Enabled = false;
                pnlMonthCreatedColor.Enabled = false;
                pnlMonthAccessedColor.Enabled =  false;
                pnlMonthWrittenColor.Enabled =  false;

                txtMonthsSinceCreated.Enabled = false;
                txtMonthsSinceWritten.Enabled =  false;
                txtMonthsSinceAccessed.Enabled =  false;

                spnMonthsSinceCreated.Enabled =  false;
                spnMonthsSinceWritten.Enabled =  false;
                spnMonthsSinceAccessed.Enabled =  false;
            }
        }

        private void chkFileNameLength_CheckedChanged(object sender, EventArgs e)
        {
            if (chkFileNameLength.Checked)
            {
            	txtMaxFileNameLength.Enabled = true;
                pnlIllegalFileNameLengthColor.Enabled = true;
            }
            else
            {
            	txtMaxFileNameLength.Enabled = false;
                pnlIllegalFileNameLengthColor.Enabled = false;
            }
        }

        private void chkGroupResults_CheckedChanged(object sender, EventArgs e)
        {
            if (chkGroupResults.Checked)
            {
            	txtGroupLevel.Enabled = true;
                spnGroupLevel.Enabled = true;
            }
            else
            {
            	txtGroupLevel.Enabled = false;
                spnGroupLevel.Enabled = false;
            }
        }

        private void chkIllegalCharacters_CheckedChanged(object sender, EventArgs e)
        {
            if (chkIllegalCharacters.Checked)
            {
            	pnlIllegalCharactersColor.Enabled = true;
                txtIllegalFileCharacters.Enabled = true;
                txtIllegalFolderCharacters.Enabled = true;
            }
            else
            {
            	pnlIllegalCharactersColor.Enabled = false;
                txtIllegalFileCharacters.Enabled = false;
                txtIllegalFolderCharacters.Enabled = false;
            }
        }

        private void chkLimitLevels_CheckedChanged(object sender, EventArgs e)
        {
            if (chkLimitLevels.Checked)
            {
            	txtLimitLevel.Enabled = true;
                spnLimitLevel.Enabled = true;
            }
            else
            {
            	txtLimitLevel.Enabled = false;
                spnLimitLevel.Enabled = false;
            }
        }

        private void chkPatternMatchFileOutput_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void chkPatternMatchFolderHighlight_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void chkShowFiles_CheckedChanged(object sender, EventArgs e)
        {
            if (chkShowFiles.Checked)
            {
            	chkPatternMatchFileOutput.Enabled = true;
                txtFileMatchPattern.Enabled = true;
                chkSkipFoldersWithNoFiles.Enabled = true;
            }
            else
            {
            	chkPatternMatchFileOutput.Enabled = false;
                chkPatternMatchFileOutput.CheckState = CheckState.Unchecked;
                txtFileMatchPattern.Enabled = false;
                chkSkipFoldersWithNoFiles.Enabled = false;                
            }
        }

        private void chkShowFolders_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void chkSkipFoldersWithNoFiles_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void cmdCancel_Click(object sender, EventArgs e)
        {
            _cancel = true;
            Close();
        }

        private void cmdCreateFolderMap_Click(object sender, EventArgs e)
        {
            if (txtStartingFolder.Text.Length > 0)
            {
            	DialogResult = System.Windows.Forms.DialogResult.OK;
                Close();
            }
            else
            {
                MessageBox.Show("Must select starting folder");
                txtStartingFolder.Focus();
            }
        }

        private void frmExcel_FolderMaps_Load(object sender, EventArgs e)
        {
            txtLimitLevel.Text = "1";
            txtLimitLevel.Enabled = false;
            spnLimitLevel.Minimum = 1;
            spnLimitLevel.Value = 1;
            spnLimitLevel.Enabled = false;

            txtGroupLevel.Text = "3";
            txtGroupLevel.Enabled = false;
            spnGroupLevel.Minimum = 3;
            spnGroupLevel.Value = 3;
            spnGroupLevel.Enabled = false;

            txtMonthsSinceCreated.Text = "24";
            txtMonthsSinceCreated.Enabled = false;
            spnMonthsSinceCreated.Minimum = 1;
            spnMonthsSinceCreated.Value = 24;
            spnMonthsSinceCreated.Enabled = false;

            txtMonthsSinceAccessed.Text = "24";
            txtMonthsSinceAccessed.Enabled = false;
            spnMonthsSinceAccessed.Minimum = 1;
            spnMonthsSinceAccessed.Value = 24;
            spnMonthsSinceAccessed.Enabled = false;

            txtMonthsSinceWritten.Text = "24";
            txtMonthsSinceWritten.Enabled = false;
            spnMonthsSinceWritten.Minimum = 1;
            spnMonthsSinceWritten.Value = 24;
            spnMonthsSinceWritten.Enabled = false;

            txtIllegalFileCharacters.Text = RegEx.cIllegalFileCharacters;
            txtIllegalFileCharacters.Enabled = false;

            txtIllegalFolderCharacters.Text = RegEx.cIllegalFolderCharacters;
            txtIllegalFolderCharacters.Enabled = false;

            txtMaxFileNameLength.Text = Common.cMaxFileNameLength.ToString();
            txtMaxFileNameLength.Enabled = false;
        }

        private void ColorBox_DoubleClick(object sender, EventArgs e)
        {
            ColorDialog1.Color = ((Panel)sender).BackColor;

            if (ColorDialog1.ShowDialog() != DialogResult.Cancel)
            {
            	((Panel)sender).BackColor = ColorDialog1.Color;
            }
        }

        private void spnGroupLevel_ValueChanged(object sender, EventArgs e)
        {
            txtGroupLevel.Text = spnGroupLevel.Value.ToString();
        }

        private void spnLimitLevel_ValueChanged(object sender, EventArgs e)
        {
            txtLimitLevel.Text = spnLimitLevel.Value.ToString();
        }

        private void spnMonthsSinceAccessed_ValueChanged(object sender, EventArgs e)
        {
            txtMonthsSinceAccessed.Text = spnMonthsSinceAccessed.Value.ToString();
        }

        private void spnMonthsSinceCreated_ValueChanged(object sender, EventArgs e)
        {
            txtMonthsSinceCreated.Text = spnMonthsSinceCreated.Value.ToString();
        }

        private void spnMonthsSinceWritten_ValueChanged(object sender, EventArgs e)
        {
            txtMonthsSinceWritten.Text = spnMonthsSinceWritten.Value.ToString();
        }

        #endregion

    }
}
