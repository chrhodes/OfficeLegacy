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
using System.Xml;
using System.Xml.Linq;

namespace SupportTools_Project.User_Interface_WPF
{
    /// <summary>
    /// Interaction logic for UserControl1.xaml
    /// </summary>
    public partial class UserControl1 : UserControl
    {
        public UserControl1()
        {
            InitializeComponent();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;

            var project = app.ActiveProject;

            var resource = project.Resources.Add("chr");

            resource.Group = "Christopher Rhodes";
            resource.Type = Microsoft.Office.Interop.MSProject.PjResourceTypes.pjResourceTypeWork;
            resource.Initials = "chr";
            resource.Hyperlink = "http://microsoft.com";

            XDocument xdoc = XDocument.Load(@"C:\temp\Resources2.xml");
            foreach (var r in xdoc.Elements("Resources").Elements("Resource"))
            {
                var r2 = project.Resources.Add(r.Attribute("Name").Value);
                r2.Type = Microsoft.Office.Interop.MSProject.PjResourceTypes.pjResourceTypeWork;
                r2.Initials = r.Attribute("Initials").Value;
                r2.Group = r.Attribute("Group").Value;
                
            }
            MessageBox.Show("WPF Rocks");
        }
    }
}
