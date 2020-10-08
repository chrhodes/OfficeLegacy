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
    public partial class ucSharePointSites : UserControl
    {
        #region Initialization

        public ucSharePointSites()
        {
            InitializeComponent();
        }

        #endregion

        // TODO: Update these

        private const string cXMLRootElement = "SupportTools_PowerPoint";
        private const string cListElements = "SharePointSites";
        private const string cElement = "Site";

        private string _RawXML;

        private IEnumerable<ListTypeInfo> _ListElements = null;
        public IEnumerable<ListTypeInfo> ListElements
        {
            get
            {
                if(null == _ListElements)
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

        private string _Url;
        public string Url
        {
            get
            {
                return _Url;
            }
            set
            {
                _Url = value;
            }
        }
        
        public delegate void ListElementsSelectionChanged();
        public event ListElementsSelectionChanged ListElementsSelectionChanged_Event;

        #region Event Handlers

        private void cbListElements_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Update as needed to surface info

            Url = ((ListTypeInfo)cbListElements.SelectedItem).Url;

            ListElementsSelectionChanged temp = Interlocked.CompareExchange(ref ListElementsSelectionChanged_Event, null, null);

            if(temp != null)
            {
                temp();
            }
        }

        private void lblListName_DoubleClick(object sender, EventArgs e)
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
                                item.Attribute("Url").Value);

            return listItems;
        }

        private void LoadNewListFromFile()
        {
            openFileDialog1.Filter = @"XML files (*.xml)|*.xml|All files (*.*)|*.*";
            openFileDialog1.FileName = "";
            openFileDialog1.InitialDirectory = @"C:\temp";

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
                cbListElements.Items.Clear();
                ListElements = null;

                _RawXML = streamReader.ReadToEnd();

                foreach(ListTypeInfo fileType in ListElements)
                {
                    cbListElements.Items.Add(fileType);
                }
            }
        }

        #endregion

        public class ListTypeInfo
        {
            // TODO: Add elements to constructor
            // See GetList() for who calls.  Need to match constructor to select new ListTypeInfo(...)

            public ListTypeInfo(string url)
            {
                _Url = url;
            }

            private string _Url;
            public string Url
            {
                get
                {
                    return _Url;
                }
                set
                {
                    _Url = value;
                }
            }

            //private string _Path;
            //public string Path
            //{
            //    get
            //    {
            //        return _Path;
            //    }
            //    set
            //    {
            //        _Path = value;
            //    }
            //}

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
