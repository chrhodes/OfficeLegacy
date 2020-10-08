using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SupportTools_Word.User_Interface.Task_Panes
{
    public partial class TaskPane_AppUtil : UserControl
    {
        public TaskPane_AppUtil()
        {
            InitializeComponent();
        }

        private void btnAddFooter_Click(object sender, EventArgs e)
        {
            Actions.Word_PageFormatting.AddFooter();
        }
    }
}
