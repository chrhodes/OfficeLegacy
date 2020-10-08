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

namespace SupportTools_PowerPoint.User_Interface.User_Controls
{
    public partial class ucStateList : UserControl
    {
        #region Initialization

        public ucStateList()
        {
            InitializeComponent();
        }

        #endregion

        // TODO: Update these

        private const string cXMLRootElement = "UnitedStatesMap";
        private const string cListElements = "States";
        private const string cElement = "State";

        private string _RawXML;

        public string StateName;
        public string ShapeName;
        public string TextLine2;
        public string ForeColor;

        private IEnumerable<ListTypeInfo> _ListElements = null;
        public IEnumerable<ListTypeInfo> ListElements
        {
            get
            {
                if (null == _ListElements)
                {
                    _ListElements = GetList(_RawXML, cXMLRootElement);
                }

                return _ListElements;
            }
            set
            {
                _ListElements = value;
            }
        }

        public delegate void ListElementsSelectionChanged();
        public event ListElementsSelectionChanged ListElementsSelectionChanged_Event;

        #region Event Handlers

        private void cbComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            //((instanceInfo)cbInstance.SelectedItem).FullName;
            //string a = cbInstance.SelectedValue.ToString();
            //string b = cbInstance.SelectedText.ToString();
            //string c = cbInstance.Text;

            ShapeName = ((ListTypeInfo)cbComboBox.SelectedItem).ShapeName;
            StateName = ((ListTypeInfo)cbComboBox.SelectedItem).StateName;
            TextLine2 = ((ListTypeInfo)cbComboBox.SelectedItem).TextLine2;
            ForeColor = ((ListTypeInfo)cbComboBox.SelectedItem).ForeColor;

            ListElementsSelectionChanged temp = Interlocked.CompareExchange(ref ListElementsSelectionChanged_Event, null, null);

            if (temp != null)
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
                                item.Attribute("ShapeName").Value,
                                item.Attribute("StateName").Value,
                                item.Attribute("TextLine2").Value,
                                item.Attribute("ForeColor").Value);

            return listItems;
        }

        private void LoadNewListFromFile()
        {
            openFileDialog1.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*";
            openFileDialog1.FileName = "";

            if (DialogResult.OK == openFileDialog1.ShowDialog())
            {
                string fileName = openFileDialog1.FileName;

                PopulateListFromFile(fileName);
            }
        }

        public void PopulateListFromFile(string fileName)
        {
            using (StreamReader streamReader = new StreamReader(fileName))
            {
                cbComboBox.Items.Clear();
                ListElements = null;

                _RawXML = streamReader.ReadToEnd();

                foreach (ListTypeInfo lti in ListElements)
                {
                    cbComboBox.Items.Add(lti);
                }
            }
        }

        #endregion

        public class ListTypeInfo
        {
            public ListTypeInfo(string shapeName, string stateName, string textLine2, string foreColor)
            {
                ShapeName = shapeName;
                StateName = stateName;
                TextLine2 = textLine2;
                ForeColor = foreColor;
            }

            public string ShapeName { get; set; }
            public string StateName { get; set; }
            public string TextLine2 { get; set; }
            public string ForeColor { get; set; }

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
