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
    public partial class frmITRDetail : Form
    {
        public frmITRDetail()
        {
            InitializeComponent();
        }

        private void LinkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(string.Format("http://lifedart/default.asp?type=incident&name={0}", txtID.Text));
        }
    }
}
