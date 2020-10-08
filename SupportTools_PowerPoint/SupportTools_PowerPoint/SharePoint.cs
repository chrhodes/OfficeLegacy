using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

namespace SupportTools_PowerPoint
{
    static class SharePoint
    {
        const string SCHEMA_V2_WEBPART = "http://schemas.microsoft.com/WebPart/v2";
        const string SCHEMA_V3_WEBPART = "http://schemas.microsoft.com/WebPart/v3";

        public static string SiteURL
        {
            get;
            set;
        }

        // TODO: This is not called from anywhere.  Need to test.
        public static string GetListID(string listName)
        {

            var query = from o in Common.ApplicationDS.dtLists
                        where o.Name == listName
                        select new {o.ID};


            foreach (var row in query)
            {
                return row.ID;
            }

            //foreach (Data.ApplicationDS.dtListsRow row in Common.ApplicationDS.dtLists)
            //{
            //    if (listName == row.Name) {
            //        return row.ID;
            //    }
            //}

	        return "<Not Found>";
        }

        public static void CheckInFile(string pageUrl, string comments, string checkInType)
        {
	        using (SharePointWS_Lists.Lists listsService = new SharePointWS_Lists.Lists())
            {
		        listsService.Credentials = System.Net.CredentialCache.DefaultCredentials;
		        listsService.Url = string.Format("{0}/_vti_bin/Lists.asmx", SiteURL);

		        try
                {
			        listsService.CheckInFile(pageUrl, comments, checkInType);
		        }
                catch (Exception ex)
                {
                    MessageBox.Show(string.Format("Exception: {0}.{1}() - {2}",
                                 System.Reflection.Assembly.GetExecutingAssembly().FullName,
                                 System.Reflection.MethodInfo.GetCurrentMethod().Name,
                                 ex.ToString()
                                 )); 
		        }
	        }
        }

        // TODO: Consider passing additional arguments to support LocalCheckOut and LastModified time args

        public static void CheckOutFile(string pageUrl)
        {
	        using (SharePointWS_Lists.Lists listsService = new SharePointWS_Lists.Lists())
            {
		        listsService.Credentials = System.Net.CredentialCache.DefaultCredentials;
		        listsService.Url = string.Format("{0}/_vti_bin/Lists.asmx", SiteURL);

		        listsService.CheckOutFile(pageUrl, "false", null);
	        }
        }

        public static void DeleteView(string listName, string viewName)
        {
            using(SharePointWS_Views.Views viewsService = new SharePointWS_Views.Views())
            {
                viewsService.Credentials = System.Net.CredentialCache.DefaultCredentials;
                viewsService.Url = string.Format("{0}/_vti_bin/Views.asmx", SiteURL);


                viewsService.DeleteView(listName, viewName);
            }            
        }

        public static void FillComboBoxWithListItems(System.Windows.Forms.ComboBox comboBox, string listName)
        {
            // Empty the list first in case already populated.
            comboBox.Items.Clear();

            using(SharePointWS_Lists.Lists listsService = new SharePointWS_Lists.Lists())
            {
                listsService.Credentials = System.Net.CredentialCache.DefaultCredentials;
                listsService.Url = string.Format("{0}/_vti_bin/Lists.asmx", SiteURL);;

                //XElement listItems = GetAllListItems(listsService, listName);

                foreach(XElement node in GetAllListItems(listsService, listName).Elements(XName.Get("row", "#RowsetSchema")))
                {
                    comboBox.Items.Add((string)node.Value);
                }
            }
        }

        public static XmlNode GetAllListContentTypes(SharePointWS_Lists.Lists listsService, string listName)
        {
            try
            {
                return listsService.GetListContentTypes(listName, "0x01");
            }
            catch(Exception)
            {
                return null;
            }
        }

        public static XElement GetAllListItems(SharePointWS_Lists.Lists listsService, string listName)
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

        // TODO: This is not called from anywhere.  Need to test.  Create a test then switch to LINQ
        public static XmlNode GetContentTypeInfo(SharePointWS_Lists.Lists listService, string listName, string contentTypeName)
        {
            
            XmlNode contentTypes = GetAllListContentTypes(listService, listName);
            XElement contentTypes2 = contentTypes.GetXElement();

            string contentTypeID = "";

            foreach(XmlNode node in contentTypes)
            {
                //Common.WriteToDebugWindow(string.Format("Name:{0} ID:{1}", node.Attributes["Name"].Value, node.Attributes["ID"].Value));

                if(node.Attributes["Name"].Value == contentTypeName)
                {
                    contentTypeID = node.Attributes["ID"].Value;
                    break;
                }
            }

            try
            {
                return listService.GetListContentType(listName, contentTypeID);
            }
            catch(Exception)
            {
                return null;
            }

            //XmlNode contentTypes = GetAllListContentTypes(listService, listName);
            //string contentTypeID = "";

            //foreach(XmlNode node in contentTypes)
            //{
            //    //Common.WriteToDebugWindow(string.Format("Name:{0} ID:{1}", node.Attributes["Name"].Value, node.Attributes["ID"].Value));

            //    if(node.Attributes["Name"].Value == contentTypeName)
            //    {
            //        contentTypeID = node.Attributes["ID"].Value;
            //        break;
            //    }
            //}

            //try
            //{
            //    return listService.GetListContentType(listName, contentTypeID);
            //}
            //catch(Exception)
            //{
            //    return null;
            //}
        }

        public static void LoadListsFromSite()
        {
            // We may get called to reload the information.  Clear any existing stuff.
            Common.ApplicationDS.dtLists.Clear(); 
            
            using(SharePointWS_Lists.Lists listService = new SharePointWS_Lists.Lists())
            {
                listService.Credentials = System.Net.CredentialCache.DefaultCredentials;
                listService.Url = string.Format("{0}/_vti_bin/Lists.asmx", SiteURL);;

                XElement listCollectionNode = null;

                try
                {
                    listCollectionNode = listService.GetListCollection().GetXElement();
                
                    foreach(XElement node in listCollectionNode.DescendantNodes())
                    {
                        Data.ApplicationDS.dtListsRow dtListRow = Common.ApplicationDS.dtLists.NewdtListsRow();
                        PopulateListRow(node, dtListRow);
                        Common.ApplicationDS.dtLists.AdddtListsRow(dtListRow);
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

        public static void LoadListViewDetailsFromSite(string listName, string viewName, Data.ApplicationDS.dtViewsRow viewRow)
        {
            using (SharePointWS_Views.Views viewsService = new SharePointWS_Views.Views())
            {
		        viewsService.Credentials = System.Net.CredentialCache.DefaultCredentials;
		        viewsService.Url = string.Format("{0}/_vti_bin/Views.asmx", SiteURL);

			    // You get more if you get the individual View (by name which is the GUID)
                GetListViewDetails(listName, viewsService, viewName, viewRow);
	        }   
        }

        public static void LoadListViewsFromSite(string listName, bool loadViewDetails)
        {
            // TODO: Figure out why this doesn't work.
        //    var query = from list in Common.ApplicationDS.dtLists.AsEnumerable()
        //                where list.Name == listName
        //                select new
        //                {
        //                    list
        //                };

        //    MessageBox.Show(query.Count().ToString());

        //    Data.ApplicationDS.dtListsRow listRow = null;
        //    foreach(var row in query)
        //    {
        //        //listRow = row;
        //        row.ToString();
        //    }

            string searchExpression = string.Format("Title = '{0}'", listName);
            DataRow[] foundRows = Common.ApplicationDS.dtLists.Select(searchExpression);

            // We should only ever find one row.
            if (foundRows.GetLength(0) > 1)
            {
                throw new ApplicationException("LoadListViewsFromSite Fatal Error");
            }

            Data.ApplicationDS.dtListsRow listRow = (Data.ApplicationDS.dtListsRow)foundRows[0];

            if (listRow.ViewsLoaded == true)
            {
                if(!loadViewDetails)
                {
                    return;
                }
                else
                {
                    // TODO SOON: Walk the views for this list and see if the details have been loaded.
                    // Load them if not.
                    ;
                }
            }
            else
            {
                using (SharePointWS_Views.Views viewsService = new SharePointWS_Views.Views())
                {
		            viewsService.Credentials = System.Net.CredentialCache.DefaultCredentials;
		            viewsService.Url = string.Format("{0}/_vti_bin/Views.asmx", SiteURL);

                    System.Diagnostics.Stopwatch stopWatch = new System.Diagnostics.Stopwatch();
                    stopWatch.Start();

                    // Get the collection of views for the specified list

		            XElement viewCollectionNode = viewsService.GetViewCollection(listName).GetXElement();

                    stopWatch.Stop();

                    if (Common.DebugLevel2)
                    {
        	            Common.WriteToDebugWindow(string.Format("LoadListViewsFromSite(): {0} - {1}", listName, stopWatch.ElapsedMilliseconds));
                    }

		            int count = 0;

		            foreach (XElement view in viewCollectionNode.DescendantNodes())
                    {
			            Data.ApplicationDS.dtViewsRow viewRow = Common.ApplicationDS.Views.NewdtViewsRow();

			            viewRow.ListName = listName;

			            // You only get this information from GetViewCollection()

			            viewRow.DisplayName = (string)view.Attribute("DisplayName");
			            viewRow.Name = (string)view.Attribute("Name");
			            viewRow.Url = (string)view.Attribute("Url");

			            // TODO: Find the last modification time if exist and track. Hum, unfortunately this does not exist

			            Common.ApplicationDS.Views.AdddtViewsRow(viewRow);


			            if (loadViewDetails)
                        {
			                string viewName = viewRow.Name;

                            stopWatch.Reset();
                            stopWatch.Start();

			                // You get more if you get the individual View by name (which is the GUID)
				            GetListViewDetails(listName, viewsService, viewName, viewRow);

                            stopWatch.Stop();

                            if (Common.DebugLevel2)
                            {
        	                    Common.WriteToDebugWindow(string.Format("List GetView(Loop): {0} - {1} - {2}", listName, count, stopWatch.ElapsedMilliseconds ));
                            }
			            }

			            count += 1;
		            }
	            }

	            // Indicate that the views for this list have been added.
                listRow.ViewsLoaded = true;
            }
        }

        public static void LoadPagesFromSite()
        {
            // We may get called to reload the information.  Clear any existing stuff.
            Common.ApplicationDS.dtPages.Clear(); 
            
            using(SharePointWS_Lists.Lists listService = new SharePointWS_Lists.Lists())
            {
                listService.Credentials = System.Net.CredentialCache.DefaultCredentials;
                listService.Url = string.Format("{0}/_vti_bin/Lists.asmx", SiteURL);

                XElement pagesXElement = null;

                try
                {
                    pagesXElement = GetAllListItems(listService, "Pages");

                    foreach(XElement page in pagesXElement.Descendants(XName.Get("row", "#RowsetSchema")))
                    {
                        Data.ApplicationDS.dtPagesRow pageRow = Common.ApplicationDS.dtPages.NewdtPagesRow();
                        PopulatePageRow(page, pageRow);
                        Common.ApplicationDS.dtPages.AdddtPagesRow(pageRow);
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

        public static void LoadViewsFromSite(bool loadViewDetails)
        {
            foreach(XElement list in ConfigData.DefaultLists)
            {
                if(Common.DebugLevel1)
                {
                    Common.WriteToDebugWindow(list.Attribute("Name").Value);
                }

                SharePoint.LoadListViewsFromSite(list.Attribute("Name").Value, loadViewDetails);
            }
        }

        //public static void LoadWebPartsFromPage(string pageUrl, string linkFileName)
        //{
        //    // TODO: Check to see if WebParts have already been added.  If not, add them.

        //    string searchExpression = string.Format("LinkFileName = '{0}'", linkFileName);
        //    DataRow[] foundRows = Common.ApplicationDS.dtPages.Select(searchExpression);

        //    // We should only ever find one row.
        //    if (foundRows.GetLength(0) == 1)
        //    {
        //        if ((bool)foundRows[0]["WebPartsLoaded"] == true)
        //        {
        //            return;
        //        }
        //    }
        //    else if (foundRows.GetLength(0) > 1 | foundRows.GetLength(0) == 0)
        //    {
        //        throw new ApplicationException("LoadWebPartsFromPage Fatal Error");
        //    }

        //    Data.ApplicationDS.dtPagesRow pageRow = (Data.ApplicationDS.dtPagesRow)foundRows[0];

        //    using (SharePointWS_WebPartPages.WebPartPagesWebService webPartPageService = new SharePointWS_WebPartPages.WebPartPagesWebService())
        //    {
        //        webPartPageService.Credentials = System.Net.CredentialCache.DefaultCredentials;
        //        webPartPageService.Url = string.Format("{0}/_vti_bin/WebPartPages.asmx ", SiteURL);

        //        XElement webPartsXml = null;

        //        try
        //        {
        //            // GetWebPartProperties has been replaced with GetWebPartProperties2

        //            webPartsXml = webPartPageService.GetWebPartProperties2(
        //                pageUrl, 
        //                SharePointWS_WebPartPages.Storage.Shared, 
        //                SharePointWS_WebPartPages.SPWebServiceBehavior.Version3).GetXElement();

        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show(string.Format("Exception: {0}.{1}() - {2}",
        //                         System.Reflection.Assembly.GetExecutingAssembly().FullName,
        //                         System.Reflection.MethodInfo.GetCurrentMethod().Name,
        //                         ex.ToString()
        //                         ));    
        //        }

        //        PopulateWebPartsTable(webPartsXml, pageRow.Title, pageRow.EncodedAbsUrl);

        //        pageRow.WebPartsLoaded = true;

        //    }
        //}

        //private static void PopulateWebPartsTable(XElement webPartsXml, string title, string encodedAbsUrl)
        //{
        //    // CHR NOTE: All of the following work.

        //    // 1.

        //    // This doesn't work as the v2 stuff have a different default namespace.  The v3 parts have no namespace
        //    var webParts = webPartsXml.Elements(XName.Get("WebPart", "http://microsoft.com/sharepoint/webpartpages"));

        //    webParts = webPartsXml.Elements();

        //    // 2. Note the overloaded "+" operator

        //    XNamespace wpp = "http://microsoft.com/sharepoint/webpartpages";
        //    var webParts2 =  webPartsXml.Elements(wpp + "WebPart");

        //    // 3. 

        //    XNamespace wpp2 = webPartsXml.GetDefaultNamespace();

        //    var webParts3 = webPartsXml.Elements(wpp2 + "WebPart");

        //    XName name = XName.Get("WebPart", wpp2.NamespaceName);

        //    var webParts4 = webPartsXml.Elements(XName.Get("WebPart", webPartsXml.GetDefaultNamespace().NamespaceName));


        //    //Debug.Print(string.Format("LocalName <{0}> NameSpace <{1}> NameSpaceName <{2}> ToString <{3}>",
        //    //    name.LocalName, name.Namespace, name.NamespaceName, name.ToString()));

        //    foreach (XElement node in webParts)
        //    {
        //        Data.ApplicationDS.dtWebPartsRow dtWebPartRow = Common.ApplicationDS.dtWebParts.NewdtWebPartsRow();

        //        dtWebPartRow.PageTitle = title;
        //        dtWebPartRow.PageEncodedAbsUrl = encodedAbsUrl;

        //        //PopulateWebPartRow(node, dtWebPartRow, nsMgr);
        //        PopulateWebPartRow(node, dtWebPartRow);
        //        Common.ApplicationDS.dtWebParts.AdddtWebPartsRow(dtWebPartRow);
        //    }
        //}

        //private static void PopulateWebPartRow(XElement webPartNode, Data.ApplicationDS.dtWebPartsRow webPartRow)
        //{
        //    try
        //    {
        //        // All WebParts have an ID Attribute
        //        webPartRow.ID = (string)webPartNode.Attribute("ID");

        //        // There seem to be two different types of WebParts.
        //        // The v2 WebParts have all the elements directly under the <WebPart> element

        //        if (webPartNode.GetDefaultNamespace().NamespaceName.Contains("v2"))
        //        {
        //            webPartRow.WebPartType = "v2";

        //            Data.ApplicationDS.dtWebPartV2Row v2Row = Common.ApplicationDS.dtWebPartV2.NewdtWebPartV2Row();
        //            PopulateWebPartRowFromV2WebPart(webPartNode, v2Row);
        //            Common.ApplicationDS.dtWebPartV2.AdddtWebPartV2Row(v2Row);

        //            // Go grab a few things of interest
        //            webPartRow.Assembly = v2Row.Assembly;
        //            webPartRow.Type = v2Row.TypeName;
        //        }
        //        else 
        //        {
        //            // The v3 WebParts have a child element (in a different namespace) that contains the stuff of interest.
        //            if ( webPartNode.Elements().First().GetDefaultNamespace().NamespaceName.Contains("v3"))
        //            {
        //                webPartRow.WebPartType = "v3";

        //                Data.ApplicationDS.dtWebPartV3Row v3Row = Common.ApplicationDS.dtWebPartV3.NewdtWebPartV3Row();
        //                PopulateWebPartRowFromV3WebPart(webPartNode, v3Row);
        //                Common.ApplicationDS.dtWebPartV3.AdddtWebPartV3Row(v3Row);

        //                // Go grab a few things of interest
        //                char[] splitChars = new char[] { ',' };

        //                webPartRow.Assembly = (string)XElement.Parse(v3Row.metaData).Element(XName.Get("type", SCHEMA_V3_WEBPART)).Attribute("name").Value.Split(splitChars)[0];
        //                webPartRow.Type = (string)XElement.Parse(v3Row.metaData).Element(XName.Get("type", SCHEMA_V3_WEBPART)).Attribute("name").Value.Split(splitChars)[1];
        //            } 
        //            else
        //            {
        //                throw new ApplicationException("Unexpected Web Part Type");
        //            }
        //        }

        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(string.Format("Exception: {0}.{1}() - {2}",
        //            System.Reflection.Assembly.GetExecutingAssembly().FullName,
        //            System.Reflection.MethodInfo.GetCurrentMethod().Name,
        //            ex.ToString()
        //            ));
        //    }
        //}

        //private static void PopulateWebPartRowFromV2WebPart(XElement webPartNode, Data.ApplicationDS.dtWebPartV2Row webPartRow)
        //{
        //    XNamespace xmlns = webPartNode.GetDefaultNamespace();

        //    try
        //    {
        //        // Loop through the elements and populate the DataTable
        //        // This saves having to write ugly code like this for each element
        //        //webPartRow.Title = (string)webPartNode.Element(xmlns + "Title");
        //        //webPartRow.FrameState = (string)webPartNode.Element(xmlns + "FrameState");

        //        DataRow dataRow = (DataRow)webPartRow;

        //        foreach(XElement item in webPartNode.Elements())
        //        {
        //            //Debug.Print(string.Format(">{0}< - >{1}<",
        //            //    item.Name.LocalName, item.Value));
        //            //Debug.Print(string.Format("{0} - {1} - {2} - {3} - {4} - {5}",
        //            //    item.Name,
        //            //    item.Name.LocalName,
        //            //    item.Name.Namespace,
        //            //    item.Name.NamespaceName,
        //            //    item.GetType(),
        //            //    item.GetDefaultNamespace().NamespaceName));

        //            dataRow[item.Name.LocalName] = item.Value;
        //        }

        //        //DataRow dr = new DataRow();


        //        // Don't think we are going to need this switch logic anymore as long as
        //        // the dtWebParts gets extended with new columns as needed.

        //        //switch (webPartRow.TypeName)
        //        //{
        //        //    case "Microsoft.SharePoint.WebPartPages.ContentEditorWebPart":

        //        //        break;

        //        //    case "Microsoft.SharePoint.WebPartPages.ListViewWebPart":

        //        //        webPartRow.ListViewXml = (string)webPartNode.Element(XName.Get("ListViewXml", "http://schemas.microsoft.com/WebPart/v2/ListView"));
        //        //        break;

        //        //    default:

        //        //        break;
        //        //}

        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(string.Format("Exception: {0}.{1}() - {2}",
        //            System.Reflection.Assembly.GetExecutingAssembly().FullName,
        //            System.Reflection.MethodInfo.GetCurrentMethod().Name,
        //            ex.ToString()
        //            ));
        //    }

        //}

        //private static void PopulateWebPartRowFromV3WebPart(XElement webPartNode, Data.ApplicationDS.dtWebPartV3Row webPartRow)
        //{
        //    // v3 WebParts (at least the ones that have been examined so far :)) have the following structure
        //    // <WebPart>
        //    //   <webPart>
        //    //     <metaData />
        //    //     <data>
        //    //       <properties>
        //    //         <property />
        //    //         ...
        //    //       </properties>
        //    //     </data>
        //    //   </webPart>
        //    // </WebPart>

        //        //        webPartRow.ListViewXml = (string)webPartNode.Element(XName.Get("ListViewXml", SCHEMA_V3_WEBPART));

        //    char[] splitChars = new char[] { ',' };

        //    try
        //    {

        //        string foo1 = (string)webPartNode;
        //        string foo2 = (string)webPartNode.Value;
        //        string foo3 = (string)webPartNode.ToString();
        //        string foo4 = (string)webPartNode.Descendants().First().ToString();
        //        string foo5 = (string)webPartNode.Elements().First().ToString();
        //        string foo6 = (string)webPartNode.Descendants(XName.Get("metaData", SCHEMA_V3_WEBPART)).First().ToString();
        //        string foo7 = (string)webPartNode.Descendants(XName.Get("webPart", SCHEMA_V3_WEBPART)).First().ToString();

        //        try
        //        {
        //            string foo8 = (string)webPartNode.Elements(XName.Get("metaData", SCHEMA_V3_WEBPART)).First().ToString();
        //        }
        //        catch(Exception ex)
        //        {
                    
        //        }
        //        try
        //        {
        //            string foo9 = (string)webPartNode.Elements(XName.Get("webPart", SCHEMA_V3_WEBPART)).First().ToString();
        //        }
        //        catch(Exception ex)
        //        {
                    
        //        }

        //        string foo10 = (string)webPartNode.FirstNode.ToString();
        //        string foo11 = (string)webPartNode.Element(XName.Get("webPart", SCHEMA_V3_WEBPART)).FirstNode.ToString();

        //        webPartRow.metaData = (string)webPartNode.Descendants(XName.Get("metaData", SCHEMA_V3_WEBPART)).First().ToString();
        //        webPartRow.data = (string)webPartNode.Descendants(XName.Get("data", SCHEMA_V3_WEBPART)).First().ToString();  
        //        //XmlNode metaDataElement = webPartNode["metaData"];
        //        //Common.WriteToDebugWindow(string.Format("name:{0}", metaDataElement["type"].Attributes["name"].Value));
        //        //// TODO: Trim the last three parts of the assembly name
        //        //webPartRow.TypeName = metaDataElement["type"].Attributes["name"].Value.Split(splitChars)[0];
        //        //XmlNode dataElement = webPartNode["data"];

        //        //XmlNodeList propertyElements = dataElement.SelectNodes("//v3:property[@name]", nsMgr);
        //        //Common.WriteToDebugWindow(propertyElements.Count.ToString());

        //        //foreach (XmlNode node in propertyElements)
        //        //{
        //        //    if (node.ChildNodes.Count > 0)
        //        //    {
        //        //        Common.WriteToDebugWindow(string.Format("{0}  {1}  {2}", node.Attributes["name"].Value, node.NodeType.ToString(), node.InnerXml.ToString()));
        //        //    } 
        //        //    else
        //        //    {
        //        //        Common.WriteToDebugWindow(string.Format("{0}  {1}", node.Attributes["name"].Value, node.NodeType.ToString()));
        //        //    }

        //        //}
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(string.Format("Exception: {0}.{1}() - {2}",
        //            System.Reflection.Assembly.GetExecutingAssembly().FullName,
        //            System.Reflection.MethodInfo.GetCurrentMethod().Name,
        //            ex.ToString()
        //            ));
        //    }
        //}

        public static void PopulateListRow(XElement list, Data.ApplicationDS.dtListsRow listRow)
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

        public static void PopulatePageRow(XElement page, Data.ApplicationDS.dtPagesRow pageRow)
        {
            //// TODO: Until we figure out how to determine if page is checked out, just default this to false
            //pageRow.CheckedOut = false;

            //pageRow.ContentTypeId = (string)page.Attribute("ows_ContentTypeId");
            //pageRow.ContentTypeId = (string)page.Attribute("ows_ContentTypeId");
            //pageRow.FileLeafRef = (string)page.Attribute("ows_FileLeafRef");
            //pageRow.ModifiedBy = (string)page.Attribute("ows_Modified_x0020_By");
            //pageRow.CreatedBy = (string)page.Attribute("ows_Created_x0020_By");
            //pageRow.FileType = (string)page.Attribute("ows_File_x0020_Type");
            //pageRow.Title = (string)page.Attribute("ows_Title");
            //pageRow.PublishingContact = (string)page.Attribute("ows_PublishingContact");
            //pageRow.PublishingPageLayout = (string)page.Attribute("ows_PublishingPageLayout");
            //pageRow.ContentType = (string)page.Attribute("ows_ContentType");

            //// Not all values are populated
            //// TODO: Read about casting versus .Value.  Casting seems to quietly deal with the not present problem.
            //try
            //{
            //    pageRow.PageType = (string)page.Attribute("ows_PageType");
            //}
            //catch
            //{
            //    pageRow.PageType = "<none>";
            //}
            //try
            //{
            //    pageRow.AppName = (string)page.Attribute("ows_AppName");

            //}
            //catch
            //{
            //    pageRow.AppName = "";
            //}
            //try
            //{
            //    pageRow.Project = (string)page.Attribute("ows_Project");

            //}
            //catch
            //{
            //    pageRow.Project = "";
            //}
            //try
            //{
            //    pageRow.Release = (string)page.Attribute("ows_Release");

            //}
            //catch
            //{
            //    pageRow.Release = "";
            //}
            //try
            //{
            //    pageRow.TeamName = (string)page.Attribute("ows_TeamName");

            //}
            //catch
            //{
            //    pageRow.TeamName = "";
            //}

            //pageRow.BusinessOwner = (string)page.Attribute("ows_Business_x0020_Owner");
            //pageRow.PageState = (string)page.Attribute("ows_Page_x0020_State");
            //pageRow.ID = (string)page.Attribute("ows_ID");
            //pageRow.Created = (string)page.Attribute("ows_Created");
            //pageRow.Author = (string)page.Attribute("ows_Author");
            //pageRow.Modified = (string)page.Attribute("ows_Modified");
            //pageRow.Editor = (string)page.Attribute("ows_Editor");
            //pageRow.ModerationStatus = (string)page.Attribute("ows__ModerationStatus");
            //pageRow.FileRef = (string)page.Attribute("ows_FileRef");
            //pageRow.FileDirRef = (string)page.Attribute("ows_FileDirRef");
            //pageRow.LastModified = (string)page.Attribute("ows_Last_x0020_Modified");
            //pageRow.CreatedDate = (string)page.Attribute("ows_Created_x0020_Date");
            //pageRow.FileSize = (string)page.Attribute("ows_File_x0020_Size");
            //pageRow.FSObjType = (string)page.Attribute("ows_FSObjType");
            //pageRow.PermMask = (string)page.Attribute("ows_PermMask");
            //pageRow.CheckedOutUserId = (string)page.Attribute("ows_CheckedOutUserId");
            //pageRow.IsCheckedoutToLocal = (string)page.Attribute("ows_IsCheckedoutToLocal");
            //pageRow.UniqueId = (string)page.Attribute("ows_UniqueId");
            //pageRow.ProgId = (string)page.Attribute("ows_ProgId");
            //pageRow.ScopeId = (string)page.Attribute("ows_ScopeId");
            //pageRow.VirusStatus = (string)page.Attribute("ows_VirusStatus");
            //pageRow.CheckedOutTitle = (string)page.Attribute("ows_CheckedOutTitle");
            //pageRow.CheckinComment = (string)page.Attribute("ows__CheckinComment");
            //pageRow.EditMenuTableStart = (string)page.Attribute("ows__EditMenuTableStart");
            //pageRow.EditMenuTableEnd = (string)page.Attribute("ows__EditMenuTableEnd");
            //pageRow.LinkFilenameNoMenu = (string)page.Attribute("ows_LinkFilenameNoMenu");
            //pageRow.LinkFilename = (string)page.Attribute("ows_LinkFilename");
            //pageRow.DocIcon = (string)page.Attribute("ows_DocIcon");
            //pageRow.ServerUrl = (string)page.Attribute("ows_ServerUrl");
            //pageRow.EncodedAbsUrl = (string)page.Attribute("ows_EncodedAbsUrl");
            //pageRow.BaseName = (string)page.Attribute("ows_BaseName");
            //pageRow.FileSizeDisplay = (string)page.Attribute("ows_FileSizeDisplay");
            //pageRow.MetaInfo = (string)page.Attribute("ows_MetaInfo");
            //pageRow.Level = (string)page.Attribute("ows__Level");
            //pageRow.IsCurrentVersion = (string)page.Attribute("ows__IsCurrentVersion");
            //pageRow.SelectTitle = (string)page.Attribute("ows_SelectTitle");
            //pageRow.SelectFilename = (string)page.Attribute("ows_SelectFilename");
            //pageRow.owshiddenversion = (string)page.Attribute("ows_owshiddenversion");
            //pageRow.UIVersion = (string)page.Attribute("ows__UIVersion");
            //pageRow.UIVersionString = (string)page.Attribute("ows__UIVersionString");
            //pageRow.Order = (string)page.Attribute("ows_Order");
            //pageRow.GUID = (string)page.Attribute("ows_GUID");
            //pageRow.WorkflowVersion = (string)page.Attribute("ows_WorkflowVersion");
            //pageRow.ParentVersionString = (string)page.Attribute("ows_ParentVersionString");
            //pageRow.ParentLeafName = (string)page.Attribute("ows_ParentLeafName");
            //pageRow.Combine = (string)page.Attribute("ows_Combine");
            //pageRow.RepairDocument = (string)page.Attribute("ows_RepairDocument");

            // Not sure how this could have every worked as there is no xmlns:z attribute??
            //try
            //{
            //    pageRow.xmlnsz = node.Attributes["xmlns:z"].Value;               
            //}
            //catch(Exception)
            //{
            //    pageRow.xmlnsz = "";
            //}



            //try
            //{
            //    Common.WriteToDebugWindow(String.Format("  CheckedOutUserId:{0}", node.Attributes["ows_CheckedOutUserId"].Value));
            //    try
            //    {
            //        Common.WriteToDebugWindow(string.Format("  CheckoutUser:{0}", node.Attributes["ows_CheckoutUser"].Value));
            //        pageRow.CheckedOut = true;
            //    }
            //    catch (Exception)
            //    {
            //        pageRow.CheckedOut = false;
            //    }
            //}
        }

        public static void UpdateView(string listName, string viewID, string viewQuery, string viewViewFields, string viewAggregations, string viewRowLimit)
        {
	        using (SharePointWS_Views.Views viewsService = new SharePointWS_Views.Views())
            {
		        viewsService.Credentials = System.Net.CredentialCache.DefaultCredentials;
		        string viewsWebServiceUrl = string.Format("{0}/_vti_bin/Views.asmx", SiteURL);

		        viewsService.Url = viewsWebServiceUrl;

		        XmlNode viewXmlNode = default(XmlNode);

                XmlNode viewProperties = null;
                XmlNode formats = null;

		        try
                {
			        viewXmlNode = viewsService.UpdateView(
                        listName, 
                        viewID, 
                        viewProperties, 
                        Util.ConvertToXmlNode(viewQuery), 
                        Util.ConvertToXmlNode(viewViewFields), 
                        Util.ConvertToXmlNode(viewAggregations), 
                        formats, 
                        Util.ConvertToXmlNode(viewRowLimit));
		        }
                catch (Exception ex)
                {
			        // TODO: Handle this better.  Should not display from Util class.
			        MessageBox.Show(string.Format("Error Updating {0} Views{1}{2}", listName, "\n", ex));
		        }
	        }
        }

        #region Private Methods

        /// <summary>
        /// This method is private, there is a public method that can be called, LoadListViewDetailsFromSite()
        /// </summary>
        /// <param name="listName"></param>
        /// <param name="viewsService"></param>
        /// <param name="viewName"></param>
        /// <param name="viewRow"></param>
        internal static void GetListViewDetails(string listName, SharePointWS_Views.Views viewsService, string viewName, Data.ApplicationDS.dtViewsRow viewRow)
        {

            XmlNode xmlViewNode = default(XmlNode);

            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();

            xmlViewNode = viewsService.GetView(listName, viewName);

            stopWatch.Stop();

            if(Common.DebugLevel2)
            {
                Common.WriteToDebugWindow(string.Format("GetListViewDetails(): {0} - {1}", listName, stopWatch.ElapsedMilliseconds));
            }

            viewRow.Name = xmlViewNode.Attributes["Name"].Value;
            viewRow.Type = xmlViewNode.Attributes["Type"].Value;
            viewRow.DisplayName = xmlViewNode.Attributes["DisplayName"].Value;
            // For some reason the URL returned by GetView() does not include the full SiteUrl.
            // We have already populated it so just skip.  TODO NB.  The previous program used this one.  Need
            // to decide which to use.
            viewRow.Url = xmlViewNode.Attributes["Url"].Value;
            viewRow.Level = xmlViewNode.Attributes["Level"].Value;
            viewRow.BaseViewID = xmlViewNode.Attributes["BaseViewID"].Value;
            viewRow.ContentTypeID = xmlViewNode.Attributes["ContentTypeID"].Value;

            try
            {
                viewRow.ImageUrl = xmlViewNode.Attributes["ImageUrl"].Value;
            }
            catch(Exception)
            {
                viewRow.ImageUrl = "<none>";
            }

            // Get the OuterXml to make later use easier.

            try
            {
                viewRow.Query = Util.FormatXml(xmlViewNode["Query"].OuterXml);
            }
            catch(Exception)
            {
                viewRow.Query = "<none>";
            }

            viewRow.ViewFields = Util.FormatXml(xmlViewNode["ViewFields"].OuterXml);

            try
            {
                viewRow.RowLimit = Util.FormatXml(xmlViewNode["RowLimit"].OuterXml);
            }
            catch(Exception)
            {
                viewRow.RowLimit = "<none>";
            }

            try
            {
                viewRow.Aggregations = Util.FormatXml(xmlViewNode["Aggregations"].OuterXml);
            }
            catch(Exception)
            {
                viewRow.Aggregations = "<none>";
            }

            viewRow.OuterXml = xmlViewNode.OuterXml;

            viewRow.DetailsLoadTime = DateTime.Now;
            viewRow.ViewDetailsLoaded = true;
        }

        #endregion
    }
}
