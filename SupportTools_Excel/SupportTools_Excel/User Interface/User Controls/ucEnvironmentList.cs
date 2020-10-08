using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using System.Collections;

namespace SupportTools_Excel.User_Interface.User_Controls
{
    public partial class ucEnvironmentList : UserControl
    {
        #region Initialization

        public ucEnvironmentList()
        {
            InitializeComponent();
        }

        #endregion

        // TODO: Update these

        private const string cXMLRootElement = "ExcelMTreatyAdmin";
        private const string cListElements = "EnvironmentList";
        private const string cElement = "Environment";

        private string _RawXML;

        public string Environment;
        public string Path;

        private IEnumerable<ListTypeInfo> _Environments = null;
        public IEnumerable<ListTypeInfo> Environments
        {
            get
            {
                if(null == _Environments)
                {
                    _Environments = GetList(_RawXML, cXMLRootElement);
                }

                return _Environments;
            }
            set
            {
                _Environments = value;
            }
        }

        public delegate void EnvironmentselectionChanged();
        public event EnvironmentselectionChanged EnvironmentselectionChanged_Event;

        #region Event Handlers

        private void cbEnvironments_SelectedIndexChanged(object sender, EventArgs e)
        {
            //((instanceInfo)cbInstance.SelectedItem).FullName;
            //string a = cbInstance.SelectedValue.ToString();
            //string b = cbInstance.SelectedText.ToString();
            //string c = cbInstance.Text;

            Environment = ((ListTypeInfo)cbEnvironments.SelectedItem).Name;
            Path = ((ListTypeInfo)cbEnvironments.SelectedItem).Path;

            EnvironmentselectionChanged temp = Interlocked.CompareExchange(ref EnvironmentselectionChanged_Event, null, null);

            if(temp != null)
            {
                temp();
            }
        }

        private void lblEnvironments_DoubleClick(object sender, EventArgs e)
        {
            LoadNewListFromFile();
        }

        #endregion

        #region Main Function Routines

        private static IEnumerable<ListTypeInfo> GetList(string root, string param)
        {
            IEnumerable<ListTypeInfo> listItems = null;

            listItems =
                from item in XDocument.Parse(root).Descendants(cListElements).Elements(cElement)
                select new ListTypeInfo(
                                item.Attribute("Name").Value,
                                item.Attribute("Path").Value);

            return listItems;
        }

        private void LoadNewListFromFile()
        {
            openFileDialog1.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*";
            openFileDialog1.FileName = "";

            if(DialogResult.OK == openFileDialog1.ShowDialog())
            {
                string fileName = openFileDialog1.FileName;

                PopulateListFromFile(fileName);
            }
        }

        public void PopulateListFromFile(string fileName)
        {
            using(StreamReader streamReader = new StreamReader(fileName))
            {
                cbEnvironments.Items.Clear();
                Environments = null;

                _RawXML = streamReader.ReadToEnd();

                foreach(ListTypeInfo fileType in Environments)
                {
                    cbEnvironments.Items.Add(fileType);
                }
            }
        }

        #endregion

        public class ListTypeInfo
        {
            public ListTypeInfo(string name, string path)
            {
                _Name = name;
                _Path = path;
            }

            private string _Name;
            public string Name
            {
                get
                {
                    return _Name;
                }
                set
                {
                    _Name = value;
                }
            }

            private string _Path;
            public string Path
            {
                get
                {
                    return _Path;
                }
                set
                {
                    _Path = value;
                }
            }

            //private string _ExpandedName;
            //public string ExpandedName
            //{
            //    get
            //    {
            //        if (name.Length > 0 || port.Length > 0)
            //        {
            //            return string.Format(@"{0} ({1}) \ {2} ({3})", 
            //                server.Length > 0 ? server : "??", 
            //                ipv4Address.Length > 0 ? ipv4Address : "??", 
            //                name.Length > 0 ? name : "??", 
            //                port.Length > 0 ? port : "??");
            //        }
            //        else
            //        {
            //            return string.Format(@"{0} ({1}) \ <Default>",
            //                server.Length > 0 ? server : "??",
            //                ipv4Address.Length > 0 ? ipv4Address : "??");
            //        }

            //    }
            //    set
            //    {

            //        _ExpandedName = value;
            //    }
            //}

            //private string _FullName;
            //public string FullName
            //{
            //    get
            //    {
            //        string validName;

            //        // Prefer HostName over IPs and InstanceNames over Port#s

            //        if (name.Length > 0 || port.Length > 0)  // Have Instance Info           
            //        {   
            //            validName = string.Format(@"{0}{1}",
            //                server.Length > 0 ? server : ipv4Address,
            //                name.Length > 0   ? @"\" + name : "," + port
            //                );
            //        }
            //        else                                                
            //        {
            //            validName = server.Length > 0 ? server : ipv4Address;
            //        }

            //        return validName;
            //    }
            //    set
            //    {
            //        _FullName = value;
            //    }
            //}

        }
    }
}
