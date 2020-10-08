using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;


namespace SupportTools_PowerPoint.User_Interface.Task_Panes
{
    public partial class TaskPane_SharePointInfo : UserControl
    {

        #region Initialization

        public TaskPane_SharePointInfo()
        {
            InitializeComponent();
        }

        private void TaskPane_SharePointInfo_Load(object sender, EventArgs e)
        {
            ucSharePointSites1.PopulateListFromFile(@"C:\Temp\SharePointSites.xml");
        }

        #endregion

        #region Event Handlers

        private void btnGetItems_Click(object sender, EventArgs e)
        {
            FillComboBoxWithListItems(cbItems, cbSharePointLists.Text);
        }

        private void btnGetSiteLists_Click(object sender, EventArgs e)
        {
            LoadListsFromSite();
        }

        private void btnGetSitePages_Click(object sender, EventArgs e)
        {

        }

        private void btnGetSitePages_Click_1(object sender, EventArgs e)
        {
            MessageBox.Show("Not Yet");
        }

        private void btnGetSiteViews_Click(object sender, EventArgs e)
        {

        }


        private void btnGetSiteViews_Click_1(object sender, EventArgs e)
        {
            SharePoint.LoadViewsFromSite(false);
        }
        #endregion

        public void LoadListsFromSite()
        {
            // We may get called to reload the information.  Clear any existing stuff.
            Common.ApplicationDS.dtLists.Clear(); 
            
            using(SharePointWS_Lists.Lists listService = new SharePointWS_Lists.Lists())
            {
                listService.Credentials = System.Net.CredentialCache.DefaultCredentials;
           
                listService.Url = string.Format("{0}/_vti_bin/Lists.asmx", ucSharePointSites1.Url);;

                XElement listCollectionNode = null;

                cbSharePointLists.Items.Clear();

                try
                {
                    listCollectionNode = listService.GetListCollection().GetXElement();
                
                    foreach(XElement node in listCollectionNode.DescendantNodes())
                    {
                        Data.ApplicationDS.dtListsRow dtListRow = Common.ApplicationDS.dtLists.NewdtListsRow();
                        PopulateListRow(node, dtListRow);
                        cbSharePointLists.Items.Add(dtListRow.Title);
                        Common.ApplicationDS.dtLists.AdddtListsRow(dtListRow);
                        Common.WriteToDebugWindow(String.Format("Title:{0}   DefaultViewUrl:{1}", dtListRow.Title, dtListRow.DefaultViewUrl));
                    }
                }
                catch(Exception ex)
                {
                    MessageBox.Show(string.Format("Exception: {0}.{1}() - {2}",
                                 System.Reflection.Assembly.GetExecutingAssembly().FullName,
                                 System.Reflection.MethodInfo.GetCurrentMethod().Name,
                                 ex.ToString()
                                 ));      
                }
            }
        }

        public void FillComboBoxWithListItems(System.Windows.Forms.ComboBox comboBox, string listName)
        {
            // Empty the list first in case already populated.
            comboBox.Items.Clear();

            using(SharePointWS_Lists.Lists listsService = new SharePointWS_Lists.Lists())
            {
                listsService.Credentials = System.Net.CredentialCache.DefaultCredentials;
                listsService.Url = string.Format("{0}/_vti_bin/Lists.asmx", ucSharePointSites1.Url);;

                //XElement listItems = GetAllListItems(listsService, listName);

                XElement foo = GetAllListItems(listsService, listName);
                var elements = foo.Elements(XName.Get("row", "#RowsetSchema"));
                var descendents = foo.Descendants(XName.Get("row", "#RowsetSchema"));

                //foreach(XElement node in GetAllListItems(listsService, listName).Elements(XName.Get("row", "#RowsetSchema")))
                //{
                //    comboBox.Items.Add((string)node.Value);
                //}

                foreach(XElement node in descendents)
                {
                    Common.WriteToDebugWindow(String.Format("Title:{0}   Created:{1}   IIPRC Filing?:{2}", 
                        (string)node.Attribute("ows_Title"), (string)node.Attribute("ows_Created")));
                    comboBox.Items.Add((string)node.Attribute("ows_Title"));
                }
            }
        }

        public XElement GetAllListItems(SharePointWS_Lists.Lists listsService, string listName)
        {
            string viewName = null;

            XElement query = new XElement("Query");
            XElement viewFields = new XElement("ViewFields");

            // TODO: This is a hack.  Not sure what to do if the list has more than 1000 rows.
            // If you don't specify the rowLimit it defaults to what the "default" view allows unless you
            // specify a different view.

            string rowLimit = "1000";

            XElement queryOptions = new XElement("QueryOptions");
            string webID = null;
            
            try
            {
                XmlNode listItems = listsService.GetListItems(listName, viewName, query.GetXmlNode(), viewFields.GetXmlNode(), rowLimit, queryOptions.GetXmlNode(), webID);

                Common.WriteToDebugWindow(listItems.OuterXml);
                return listsService.GetListItems(listName, viewName, query.GetXmlNode(), viewFields.GetXmlNode(), rowLimit, queryOptions.GetXmlNode(), webID).GetXElement();
            }
            catch(Exception ex)
            {
                MessageBox.Show(string.Format("Exception: {0}.{1}() - {2}",
                    System.Reflection.Assembly.GetExecutingAssembly().FullName,
                    System.Reflection.MethodInfo.GetCurrentMethod().Name,
                    ex.ToString()
                    ));
                return null;
            }
        }

        #region Main Function Routines
        public void PopulateListRow(XElement list, Data.ApplicationDS.dtListsRow listRow)
        {
            try
            {
                listRow.Title = (string)list.Attribute("Title");
                listRow.ID = (string)list.Attribute("ID");
                listRow.DocTemplateUrl = (string)list.Attribute("DocTemplateUrl");
                listRow.DefaultViewUrl = (string)list.Attribute("DefaultViewUrl");
                listRow.Description = (string)list.Attribute("Description");
                listRow.ImageUrl = (string)list.Attribute("ImageUrl");
                listRow.Name = (string)list.Attribute("Name");
                listRow.BaseType = (string)list.Attribute("BaseType");
                listRow.ServerTemplate = (string)list.Attribute("ServerTemplate");
                listRow.Created = (string)list.Attribute("Created");
                listRow.Modified = (string)list.Attribute("Modified");
                listRow.LastDeleted = (string)list.Attribute("LastDeleted");
                listRow.Version = (string)list.Attribute("Version");
                listRow.Direction = (string)list.Attribute("Direction");
                listRow.ThumbnailSize = (string)list.Attribute("ThumbnailSize");
                listRow.WebImageHeight = (string)list.Attribute("WebImageHeight");
                listRow.WebImageWidth = (string)list.Attribute("WebImageWidth");
                listRow.Flags = (string)list.Attribute("Flags");
                listRow.ItemCount = (string)list.Attribute("ItemCount");
                listRow.AnonymousPermsMask = (string)list.Attribute("AnonymousPermMask");
                listRow.RootFolder = (string)list.Attribute("RootFolder");
                listRow.ReadSecurity = (string)list.Attribute("ReadSecurity");
                listRow.WriteSecurity = (string)list.Attribute("WriteSecurity");
                listRow.Author = (string)list.Attribute("Author");
                listRow.EventSinkAssembly = (string)list.Attribute("EventSinkAssembly");
                listRow.EventSinkClass = (string)list.Attribute("EventSinkClass");
                listRow.EventSinkData = (string)list.Attribute("EventSinkData");
                listRow.EmailInsertsFolder = (string)list.Attribute("EmailInsertsFolder");
                listRow.AllowDeletion = (string)list.Attribute("AllowDeletion");
                listRow.AllowMultiResponses = (string)list.Attribute("AllowMultiResponses");
                listRow.EnableAttachments = (string)list.Attribute("EnableAttachments");
                listRow.EnableModeration = (string)list.Attribute("EnableModeration");
                listRow.EnableVersioning = (string)list.Attribute("EnableVersioning");
                listRow.Hidden = (string)list.Attribute("Hidden");
                listRow.MultipleDataList = (string)list.Attribute("MultipleDataList");
                listRow.Ordered = (string)list.Attribute("Ordered");
                listRow.ShowUser = (string)list.Attribute("ShowUser");

                // Views have not been loaded for this list, yet.
                listRow.ViewsLoaded = false;
            }
            catch(Exception ex)
            {
                MessageBox.Show(string.Format("Exception: {0}.{1}() - {2}",
                    System.Reflection.Assembly.GetExecutingAssembly().FullName,
                    System.Reflection.MethodInfo.GetCurrentMethod().Name,
                    ex.ToString()
                    ));
            }
        }

        #endregion

        private void btnGetListViews_Click(object sender, EventArgs e)
        {
            GetViews();
        }

        private void GetViews()
        {
            if (cbSharePointLists.Text.Length > 0)
            {
                cbViews.Items.Clear();

                // Get view details (true)
            	SharePoint.LoadListViewsFromSite(cbSharePointLists.Text, true);

                foreach(Data.ApplicationDS.dtViewsRow view in Common.ApplicationDS.dtViews)
                {

                    cbViews.Items.Add(view.DisplayName);
                    Common.WriteToDebugWindow(String.Format("DisplayName:{0}  ListName:{1}  Name:{2}  Query:{3}  Url:{4}", 
                        view.DisplayName, view.ListName, view.Name, view.Query, view.Url));
                             
                }
            }
        }

        private void ucSharePointSites1_ListElementsSelectionChanged_Event()
        {
            SharePoint.SiteURL = ucSharePointSites1.Url;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
