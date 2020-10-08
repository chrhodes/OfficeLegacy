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
    public partial class ucColorList : UserControl
    {
        #region Initialization

        public ucColorList()
        {
            InitializeComponent();
        }

        #endregion

        // TODO: Update these

        private const string cXMLRootElement = "UnitedStatesMap";
        private const string cListElements = "Colors";
        private const string cElement = "Color";

        private string _RawXML;

        public string ColorName;
        public string ColorValue;

        private IEnumerable<ListTypeInfo> _Colors = null;
        public IEnumerable<ListTypeInfo> Colors
        {
            get
            {
                if (null == _Colors)
                {
                    _Colors = GetList(_RawXML, cXMLRootElement);
                }

                return _Colors;
            }
            set
            {
                _Colors = value;
            }
        }

        public delegate void ColorsSelectionChanged();
        public event ColorsSelectionChanged ColorsSelectionChanged_Event;

        #region Event Handlers

        private void cbColors_SelectedIndexChanged(object sender, EventArgs e)
        {
            //((instanceInfo)cbInstance.SelectedItem).FullName;
            //string a = cbInstance.SelectedValue.ToString();
            //string b = cbInstance.SelectedText.ToString();
            //string c = cbInstance.Text;

            ColorName = ((ListTypeInfo)cbColors.SelectedItem).ColorName;
            ColorValue = ((ListTypeInfo)cbColors.SelectedItem).ColorValue;

            ColorsSelectionChanged temp = Interlocked.CompareExchange(ref ColorsSelectionChanged_Event, null, null);

            if (temp != null)
            {
                temp();
            }
        }

        private void lblColors_DoubleClick(object sender, EventArgs e)
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
                                item.Attribute("Value").Value);

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
                cbColors.Items.Clear();
                Colors = null;

                _RawXML = streamReader.ReadToEnd();

                foreach (ListTypeInfo color in Colors)
                {
                    cbColors.Items.Add(color);
                }
            }
        }

        #endregion

        public class ListTypeInfo
        {
            public ListTypeInfo(string colorName, string colorValue)
            {
                ColorName = colorName;
                ColorValue = colorValue;
            }

            public string ColorName { get; set; }
            public string ColorValue { get; set; }

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
