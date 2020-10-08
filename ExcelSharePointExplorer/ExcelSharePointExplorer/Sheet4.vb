Imports System.Xml

Public Class Sheet4

    Private Sub Sheet4_Startup(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Startup

    End Sub

    Private Sub Sheet4_Shutdown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shutdown

    End Sub

    Private Sub btnGetLists_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetLists.Click
        Using listsService As New learnsharepoint_lists.Lists()
            listsService.Credentials = System.Net.CredentialCache.DefaultCredentials
            Dim listsWebServiceUrl As String = "http://learnsharepoint/teams/ITTechSvc/_vti_bin/Lists.asmx"

            'Dim cache As System.Net.CredentialCache = New System.Net.CredentialCache()

            'cache.Add(New Uri(listsWebServiceUrl), "NTLM", New System.Net.NetworkCredential("crhodes", "HappyH0jnacki8", "PACIFICMUTUAL"))
            'listsService.Credentials = cache

            listsService.Url = listsWebServiceUrl

            Dim listCollectionNode As System.Xml.XmlNode = listsService.GetListCollection()
            Dim xmlNode As System.Xml.XmlNode

            Dim viewName As String = ""
            Dim listName As String = ""
            Dim i As Integer = 5

            For Each xmlNode In listCollectionNode
                Cells(i, 1).Value = xmlNode.Attributes("Title").Value
                i += 1
            Next
        End Using
    End Sub

    Private Sub btnGetData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetData.Click
        Using listsService As New learnsharepoint_lists.Lists()
            listsService.Credentials = System.Net.CredentialCache.DefaultCredentials
            Dim listsWebServiceUrl As String = "http://learnsharepoint/teams/ITTechSvc/_vti_bin/Lists.asmx"

            'Dim cache As System.Net.CredentialCache = New System.Net.CredentialCache()

            'cache.Add(New Uri(listsWebServiceUrl), "NTLM", New System.Net.NetworkCredential("crhodes", "HappyH0jnacki8", "PACIFICMUTUAL"))
            'listsService.Credentials = cache

            listsService.Url = listsWebServiceUrl

            Dim xmlDoc As XmlDocument = New XmlDocument

            Dim listName As String = Application.ActiveCell.Value
            Dim viewName As String = Nothing
            Dim query As XmlElement = xmlDoc.CreateElement("Query")
            Dim viewFields As XmlElement = xmlDoc.CreateElement("ViewFields")
            Dim rowLimit As String = Nothing
            Dim queryOptions As XmlElement = xmlDoc.CreateElement("QueryOptions")
            Dim webID As String = Nothing

            Dim listItemsXml As XmlNode = listsService.GetListItems(listName, viewName, query, viewFields, rowLimit, queryOptions, webID)

            'Dim listCollectionNode As System.Xml.XmlNode = listsService.GetListCollection()
            'Dim xmlNode As System.Xml.XmlNode

            Dim i As Integer = 5

            Dim xmlDocResults = New XmlDataDocument
            xmlDocResults.LoadXml(listItemsXml.InnerXml)
            Dim resultRows As XmlNodeList = xmlDocResults.GetElementsByTagName("z:row")

            For Each xmlNode As XmlNode In resultRows
                Cells(i, 5).Value = xmlNode.Attributes("ows_Title").Value
                i += 1
            Next
        End Using
    End Sub

    'Function ConvertToXmlNode(ByVal inputString As String) As XmlNode
    '    Dim outputXml As XmlDocument = New XmlDocument

    '    outputXml.InnerXml = inputString

    '    ' TODO: HOw to?

    '    Return outputXml
    'End Function
End Class
