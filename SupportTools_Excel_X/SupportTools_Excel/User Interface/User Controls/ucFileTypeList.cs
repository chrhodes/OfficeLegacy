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
    public partial class ucFileTypeList : UserControl
    {
        #region Initialization

        public ucFileTypeList()
        {
            InitializeComponent();
        }

        #endregion

        // TODO: Update these

        private const string cXMLRootElement = "ExcelMTreatyAdmin";
        private const string cListElements = "FileTypeList";
        private const string cElement = "FileType";

        private string _RawXML;

        public string ListDisplayValue;
        public string FileType;
        public string RelativePath;

        private IEnumerable<FileTypeInfo> _FileTypes = null;
        public IEnumerable<FileTypeInfo> FileTypes
        {
            get
            {
                if (null == _FileTypes)
                {
                	_FileTypes = GetList(_RawXML, cXMLRootElement);
                }
                
                return _FileTypes;
            }
            set
            {
                _FileTypes = value;
            }
        }

        public delegate void FileTypeSelectionChanged();
        public event FileTypeSelectionChanged FileTypeSelectionChanged_Event;

        #region Event Handlers

        private void cbFileTypes_SelectedIndexChanged(object sender, EventArgs e)
        {
            //((instanceInfo)cbInstance.SelectedItem).FullName;
            //string a = cbInstance.SelectedValue.ToString();
            //string b = cbInstance.SelectedText.ToString();
            //string c = cbInstance.Text;

            FileType = ((FileTypeInfo)cbFileTypes.SelectedItem).Name;
            RelativePath = ((FileTypeInfo)cbFileTypes.SelectedItem).RelativePath;
            ListDisplayValue = ((FileTypeInfo)cbFileTypes.SelectedItem).DisplayValue;

            FileTypeSelectionChanged temp = Interlocked.CompareExchange(ref FileTypeSelectionChanged_Event, null, null);

            if (temp != null)
            {
            	temp();
            }
        }

        private void lblFileTypes_DoubleClick(object sender, EventArgs e)
        {
            LoadNewListFromFile();
        }

        #endregion

        #region Main Function Routines

        private static IEnumerable<FileTypeInfo> GetList(string root, string param)
        {
            IEnumerable<FileTypeInfo> listItems = 
                from item in XDocument.Parse(root).Descendants(cListElements).Elements(cElement)
                select new FileTypeInfo(
                     item.Attribute("Name").Value,
                     item.Attribute("Frequency").Value,
                     item.Attribute("RelativePath").Value);
            
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
                cbFileTypes.Items.Clear();
                FileTypes = null;

                _RawXML = streamReader.ReadToEnd();

                foreach(FileTypeInfo fileType in FileTypes)
                {
                    cbFileTypes.Items.Add(fileType);
                }
            }
        }

        #endregion

        public class FileTypeInfo
        {
            public FileTypeInfo(string name, string frequency, string relativePath)
            {
                _Name = name;
                _Frequency = frequency;
                _RelativePath = relativePath;
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

            private string _Frequency;
            public string Frequency
            {
            	get	{ return _Frequency; }
            	set
            	{
            		_Frequency = value;
            	}
            }
            
            private string _RelativePath;
            public string RelativePath
            {
                get
                {
                    return _RelativePath;
                }
                set
                {
                    _RelativePath = value;
                }
            }

            private string _DisplayValue;
            public string DisplayValue
            {
                get
                {
                    return String.Format("{0} ({1})", Name, Frequency);
                }
                set
                {
                    _DisplayValue = value;
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
