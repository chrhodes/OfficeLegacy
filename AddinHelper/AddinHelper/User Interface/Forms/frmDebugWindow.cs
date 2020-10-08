using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AddinHelper.User_Interface.Forms
{
    public partial class frmDebugWindow : Form
    {
        public frmDebugWindow()
        {
            InitializeComponent();
        }

        #region Event Handlers

        private void btnClearOutput_Click(object sender, EventArgs e)
        {
            txtOutput.Clear();
        }

        //private void frmDebugWindow_FormClosing(object sender, FormClosingEventArgs e)
        //{
        //    Hide();
        //    e.Cancel = true;
        //}

        private void chkDebugSQL_CheckedChanged(object sender, EventArgs e)
        {
            Common.DebugSQL = ((CheckBox)sender).Checked;
        }

        private void chkDebugLevel1_CheckedChanged(object sender, EventArgs e)
        {
            Common.DebugLevel1 = ((CheckBox)sender).Checked;
        }

        private void chkDebugLevel2_CheckedChanged(object sender, EventArgs e)
        {
            Common.DebugLevel2 = ((CheckBox)sender).Checked;
        }

        private void frmDebugWindow_FormClosed(object sender, FormClosedEventArgs e)
        {
            Common.DebugWindow = null;
        }

        #endregion

        #region Main Function Routines

        public void AddOutputLine(string outputLine)
        {
            txtOutput.AppendText(outputLine + Environment.NewLine);
        }

        #endregion
    }
}
