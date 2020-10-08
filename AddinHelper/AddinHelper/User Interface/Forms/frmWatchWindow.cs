using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace AddinHelper.User_Interface.Forms
{
    public partial class frmWatchWindow : Form
    {
        public frmWatchWindow()
        {
            InitializeComponent();
        }

        internal void AddOutputLine(string outputLine)
        {
            txtOutput.AppendText(outputLine + Environment.NewLine);
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            txtOutput.Clear();
        }

        private void btnSaveToFile_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Not Implemented Yet");
        }

        private void frmWatchWindow_FormClosed(object sender, FormClosedEventArgs e)
        {
            Common.WatchWindow = null;
        }
    }
}
