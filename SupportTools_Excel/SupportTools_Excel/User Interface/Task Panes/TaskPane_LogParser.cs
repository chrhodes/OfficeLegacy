using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SupportTools_Excel.User_Interface.Task_Panes
{
    public partial class TaskPane_LogParser : UserControl
    {
        public TaskPane_LogParser()
        {
            InitializeComponent();
            label1.Text = Width.ToString();
            label2.Text = Height.ToString();
        }
    }
}
