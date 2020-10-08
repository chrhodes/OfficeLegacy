using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace SupportTools_PowerPoint
{
    class Util
    {
        static string _xmlnsPattern = " xmlns=.http://schemas.microsoft.com/sharepoint/soap/.";

        public static XmlNode ConvertToXmlNode(string inputString)
        {
	        XmlDocument outputXml = new XmlDocument();

	        //Dim s As String = RegularExpressions.Regex.Replace(inputString, "\r\n", "")

	        //Common.WriteToDebugWindow(String.Format("inputstring:{0}  s:{1}", inputString.Length, s.Length))

            //outputXml.InnerXml = System.Text.RegularExpressions.Regex.Replace(inputString, "\\r\\n", "");
	        outputXml.InnerXml = System.Text.RegularExpressions.Regex.Replace(inputString, Environment.NewLine, "");

	        return outputXml;
        }

        public static string FormatXml(string inputXml)
        {
            string result = RemoveXmlNs(inputXml);
            return System.Text.RegularExpressions.Regex.Replace(result, ">", String.Format(">{0}", Environment.NewLine));
        }

        public static string RemoveIllegalCharacters(string str)
        {
	        str = str.Replace(" ", "");
	        str = str.Replace("&", "");

	        return str;
        }

        private static string RemoveXmlNs(string inputXml)
        {
	        return System.Text.RegularExpressions.Regex.Replace(inputXml, _xmlnsPattern, "");
        }

        public static string ReplaceIllegalCharacters(string str)
        {
	        //str = str.Replace(" ", "")
	        str = str.Replace("&", "&amp;");

	        return str;
        }

    }
}
