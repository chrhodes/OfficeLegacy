using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Xml.Linq;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using PacificLife.Life;

namespace SupportTools_Visio.User_Interface.Windows
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //private readonly static int CLASS_BASE_ERRORNUMBER = HIS.ErrorNumbers.HIS_ADMINISTRATION;
        //private const string PLLOG_APPNAME = "HIS";

        public MainWindow()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            Actions.Class1.DoSomething();
        }


        #region Main Function Routines


        #endregion

    }
}
