using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace SupportTools_Visio.User_Interface.User_Controls_WPF
{
    /// <summary>
    /// Interaction logic for UserControl1.xaml
    /// </summary>
    public partial class wucRetrieveShape : UserControl
    {
        public wucRetrieveShape()
        {
            InitializeComponent();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            var row = gcItems.View.FocusedRowData.Row;
            
            MessageBox.Show("Button_Click_1");
        }

        private void RadioButton_Click_1(object sender, RoutedEventArgs e)
        {
            switch (((RadioButton)sender).Tag.ToString())
            {
                case "SelectAll":
                    
                    break;

                case "ClearAll":
                    
                    break;
            }
            ;
        }
    }
}
