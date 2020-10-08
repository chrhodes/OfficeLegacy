using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

namespace SupportTools_Excel_MTreaty
{
    /// <summary>
    /// ConfigData
    /// 
    /// This class holds all configuration data read from the configuration information passed
    /// in the Reload() method when the initial UserControl is loaded.
    /// 
    /// To use this class, 
    /// 1. Add code in the Reload() method of the initial UserControl to store the ConfigData
    ///    using the RawXML property.  This is done automatically in ucBase.
    /// 2. Add properties to surface the configuration data.  Mark as static.
    ///    Initialize backing store variables!
    /// </summary>
    /// <remarks>Place only static items in this class.  This Class cannot be instantiated.</remarks>
    static class Config
    {
        // TODO: Update this with the root element from the configuration data.
        private static string _RootElementName = "AppConfigData";

        #region Initialization

        public static string RawXML
        {
            get;
            set;
        }

        #endregion

        #region Public Configuration Data

        // TODO: Add static Properties to expose configuration information here

        #region Template Samples

        // Cat, Dog, NorthWindDBConnectionString, and DataBaseOneConnection
        // are examples and may be deleted.

        private static string _Cat = "";
        public static string Cat
        {
            get
            {
                return PopulateConfigValue("Cat", ref _Cat);
            }
            set
            {
                _Cat = value;
            }
        }


        private static string _Dog;
        public static string Dog
        {
            get
            {
                return PopulateConfigValue("Dog", ref _Dog);
            }
            set
            {
                _Dog = value;
            }
        }

        private static string _DataBaseOneConnection = "";
        public static string DataBaseOneConnection
        {
            get
            {
                return PopulateConfigValue("DatabaseOneConnection", ref _DataBaseOneConnection);
            }
            set
            {
                _DataBaseOneConnection = value;
            }
        }

        private static string _northWindDBConnectionString = "";
        public static string NorthWindDBConnectionString
        {
            get
            {
                return PopulateConfigValue("NorthWindDBConnection", ref _northWindDBConnectionString);
            }
            set
            {
                _northWindDBConnectionString = value;
            }
        }

        #region Complex Property example

        // This shows how to return a more complex type from the config data.

        // 1. First Define the type

        public class urlinfo
        {
            public string envCode;
            public string url;
            public string description;

            public urlinfo(string EnvCode, string Url, string Description)
            {
                envCode = EnvCode;
                url = Url;
                description = Description;
            }
        }

        // 2. Declare a property to expose the type.

        private static IEnumerable<urlinfo> _URLs;
        public static IEnumerable<urlinfo> URLs
        {
            get
            {
                _URLs = GetUrlList("AppConfigData", "Url");
                return _URLs;
            }
            set
            {
                _URLs = value;
            }
        }

        // 3. Declare a method to extract the type from the config data.

        private static IEnumerable<urlinfo> GetUrlList(string root, string param)
        {
            IEnumerable<urlinfo> urls = null;

            if (RawXML == null)
            {
                MessageBox.Show("Error GetUrlList(). ConfigData has not been loaded in Reload() yet.", "Config Data Error", MessageBoxButtons.OK);
                return urls;
            }

            // Don't think we even need to wrap this in a try catch as it is not evaluated here.

            try
            {
                urls =
                    from item in XDocument.Parse(RawXML).Descendants(root).Descendants(param)
                    select new urlinfo(item.Attribute("EnvCode").Value, item.Value, item.Attribute("Description").Value);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            return urls;
        }

        #endregion

        #endregion

        #endregion

        public static string GetAllConfigInfo()
        {
            StringBuilder info = new StringBuilder(4096);

            Type ty = typeof(Config);

            object[] methodArgs = new object[1];

            MethodInfo[] methods = ty.GetMethods();

            foreach (MethodInfo method in methods)
            {
                //Common.WriteToDebugWindow(string.Format("Name: {0} ReturnType: {1}", method.Name, method.ReturnType));

                if (method.Name.StartsWith("get"))
                {
                    string value;

                    switch (method.ReturnType.ToString())
                    {
                        case "System.Boolean":
                            try
                            {
                                value = ((Boolean)method.Invoke(ty, null)).ToString();
                            }
                            catch (Exception)
                            {
                                value = "<NULL>";
                            }
                            break;

                        case "System.String":
                            try
                            {
                                value = (string)method.Invoke(ty, null);
                            }
                            catch (Exception)
                            {
                                value = "<NULL>";
                            }

                            break;

                        case "System.Int32":
                            try
                            {
                                value = ((Int32)method.Invoke(ty, null)).ToString();
                            }
                            catch (Exception)
                            {
                                value = "<NULL>";
                            }
                            break;

                        default:
                            value = "Unknown Return Type:" + method.ReturnType.ToString();
                            break;
                    }

                    // Display get_X() as ReturnType Config.X >Value<
                    info.AppendLine(string.Format("{0,16} {1,50} >{2}<",
                        method.ReturnType, "ConfigData." + method.Name.Substring(4), value));
                }
            }

            return info.ToString();
        }

        #region Internal Methods

        internal static string PopulateConfigValue(string configElement, ref string backingVariable)
        {
            if (null == backingVariable)
            {
                MessageBox.Show(
                    string.Format("{0}.  BackingVariable for {1} not initialized",
                        System.Reflection.MethodInfo.GetCurrentMethod().Name, configElement),
                    "ConfigData Error", MessageBoxButtons.OK);
                return "";
            }

            if (backingVariable.Length == 0)
            {
                backingVariable = GetElementValue(_RootElementName, configElement);

                if (0 == backingVariable.Length)
                {
                    MessageBox.Show(
                        string.Format("{0}.  Could not find {1} in <{2}>.",
                            System.Reflection.MethodInfo.GetCurrentMethod().Name, configElement, _RootElementName),
                        "ConfigData Error", MessageBoxButtons.OK);
                }
            }

            return backingVariable;
        }

        /// <summary>
        /// Looks up an value in an XML <element></element> contained by an ancestor <element></element>
        /// </summary>
        /// <param name="descendantName"></param>
        /// <param name="elementName"></param>
        /// <returns></returns>
        internal static XElement GetDescendantElement(string descendantName, string elementName)
        {
            if (RawXML == null)
            {
                MessageBox.Show("Error GetParameter(). ConfigData has not been loaded in Reload() yet.", "Config Data Error", MessageBoxButtons.OK);
                return null;
            }

            try
            {
                // Using cool LINQ types.
                return XDocument.Parse(RawXML).Descendants(descendantName).Elements(elementName).ElementAt(0);
            }
            catch (Exception)
            {
                // We can land here because the param in the <element>(root) did not exist in the config data.
                return null;
            }
        }

        internal static IEnumerable<XElement> GetDescendantElements(string descendantName, string elementName)
        {
            if (RawXML == null)
            {
                MessageBox.Show("Error GetParameter(). ConfigData has not been loaded in Reload() yet.", "Config Data Error", MessageBoxButtons.OK);
                return null;
            }

            try
            {
                return XDocument.Parse(RawXML).Descendants(descendantName).Elements(elementName);
            }
            catch (Exception)
            {
                // We can land here because the param in the <element>(root) did not exist in the config data.
                return null;
            }
        }

        internal static IEnumerable<XElement> GetDescendants(string descendantName)
        {
            if (RawXML == null)
            {
                MessageBox.Show("Error GetParameter(). ConfigData has not been loaded in Reload() yet.", "Config Data Error", MessageBoxButtons.OK);
                return null;
            }

            try
            {
                return XDocument.Parse(RawXML).Descendants(descendantName);
            }
            catch (Exception)
            {
                // We can land here because the param in the <element>(root) did not exist in the config data.
                return null;
            }
        }

        /// <summary>
        /// Looks up an value in an XML <element></element> contained by an ancestor <element></element>
        /// </summary>
        /// <param name="descendantName"></param>
        /// <param name="elementName"></param>
        /// <returns></returns>
        internal static string GetElementValue(string descendantName, string elementName)
        {
            if (RawXML == null)
            {
                MessageBox.Show("Error GetParameter(). ConfigData has not been loaded in Reload() yet.", "Config Data Error", MessageBoxButtons.OK);
                return "";
            }

            try
            {
                // Using cool LINQ types.
                return XDocument.Parse(RawXML).Descendants(descendantName).Elements(elementName).ElementAt(0).Value;

                // Old school.
                //    xmlDoc.LoadXml(RawXML);
                //    xmlRoot = (XmlElement)xmlDoc.GetElementsByTagName(descendantName).Item(0);
                //    xml = xmlRoot.GetElementsByTagName(elementName).Item(0).InnerXml;
            }
            catch (Exception)
            {
                // We can land here because the param in the <element>(root) did not exist in the config data.
                return "";
            }
        }

        #endregion

    }
}

