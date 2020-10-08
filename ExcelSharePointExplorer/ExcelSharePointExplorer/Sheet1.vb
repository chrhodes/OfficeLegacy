Imports System.Xml

Public Class Sheet1

    Private Sub Sheet1_Startup(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Startup

    End Sub

    Private Sub Sheet1_Shutdown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shutdown

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

    Private Sub btnGetViews_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetViews.Click
        Using viewsService As New learnsharepoint_ITTechSvc.Views()
            viewsService.Credentials = System.Net.CredentialCache.DefaultCredentials
            Dim viewsWebServiceUrl As String = "http://learnsharepoint/teams/ITTechSvc/_vti_bin/Views.asmx"

            'Dim cache As System.Net.CredentialCache = New System.Net.CredentialCache()

            'cache.Add(New Uri(viewsWebServiceUrl), "NTLM", New System.Net.NetworkCredential("crhodes", "HappyH0jnacki8", "PACIFICMUTUAL"))
            'viewsService.Credentials = cache

            viewsService.Url = viewsWebServiceUrl

            Dim listName As String = ""
            listName = Application.ActiveCell.Value
            Application.Cells(4, 2).Value = listName

            Dim viewCollectionNode As System.Xml.XmlNode = viewsService.GetViewCollection(listName)
            Dim xmlNode As System.Xml.XmlNode
            Dim xmlViewNode As System.Xml.XmlNode
            Dim xmlViewHtmlNode As System.Xml.XmlNode
            Dim viewName As String = ""

            Dim i As Integer = 5

            Application.Cells(4, 2).Value = "Name"
            Application.Cells(4, 3).Value = "Type"
            Application.Cells(4, 4).Value = "DisplayName"
            Application.Cells(4, 5).Value = "Url"
            Application.Cells(4, 6).Value = "Level"
            Application.Cells(4, 7).Value = "BaseViewID"
            Application.Cells(4, 8).Value = "ContentTypeID"
            Application.Cells(4, 9).Value = "ImageURL"
            Application.Cells(4, 10).Value = "Query"
            Application.Cells(4, 11).Value = "ViewFields"
            Application.Cells(4, 12).Value = "RowLimit"
            Application.Cells(4, 13).Value = "Aggregations"
            Application.Cells(4, 14).Value = "OuterXML"

            For Each xmlNode In viewCollectionNode
                System.Diagnostics.Debug.Print(xmlNode.Attributes("DisplayName").Value)
                ' You only get this information from the viewCollecitonNode

                'Cells(i, 2).Value = xmlNode.Attributes("DisplayName").Value
                'Cells(i, 3).Value = xmlNode.Attributes("Name").Value
                'Cells(i, 4).Value = xmlNode.Attributes("Url").Value

                ' You get more if you get the individual View (by name which is the GUID)

                viewName = xmlNode.Attributes("Name").Value
                xmlViewNode = viewsService.GetView(listName, viewName)

                Cells(i, 2).Value = xmlViewNode.Attributes("Name").Value
                Cells(i, 3).Value = xmlViewNode.Attributes("Type").Value
                Cells(i, 4).Value = xmlViewNode.Attributes("DisplayName").Value
                Cells(i, 5).Value = xmlViewNode.Attributes("Url").Value
                Cells(i, 6).Value = xmlViewNode.Attributes("Level").Value
                Cells(i, 7).Value = xmlViewNode.Attributes("BaseViewID").Value
                Cells(i, 8).Value = xmlViewNode.Attributes("ContentTypeID").Value
                Try
                    Cells(i, 9).Value = xmlViewNode.Attributes("ImageUrl").Value
                Catch ex As Exception

                End Try

                Cells(i, 10).Value = xmlViewNode.Item("Query").InnerXml
                Cells(i, 11).Value = xmlViewNode.Item("ViewFields").InnerXml

                Try
                    Cells(i, 12).Value = xmlViewNode.Item("RowLimit").OuterXml
                Catch ex As Exception
                    
                End Try

                Try
                    Cells(i, 13).Value = xmlViewNode.Item("Aggregations").OuterXml
                Catch ex As Exception

                End Try

                Cells(i, 14).Value = xmlViewNode.OuterXml

                'If "DART" = xmlNode.Attributes("DisplayName").Value Then
                '    i += 1
                '    xmlViewHtmlNode = viewsService.GetViewHtml(listName, viewName)
                '    Cells(i, 6).Value = xmlViewHtmlNode.OuterXml
                'End If

                i += 1
            Next
        End Using
    End Sub

    Private Sub btnDeleteView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteView.Click
        Using viewsService As New learnsharepoint_ITTechSvc.Views()
            viewsService.Credentials = System.Net.CredentialCache.DefaultCredentials
            Dim viewsWebServiceUrl As String = "http://learnsharepoint/teams/ITTechSvc/_vti_bin/Views.asmx"

            'Dim cache As System.Net.CredentialCache = New System.Net.CredentialCache()

            'cache.Add(New Uri(viewsWebServiceUrl), "NTLM", New System.Net.NetworkCredential("crhodes", "HappyH0jnacki8", "PACIFICMUTUAL"))
            'viewsService.Credentials = cache

            viewsService.Url = viewsWebServiceUrl

            Dim listName As String
            listName = Application.Cells(4, 2).Value

            Dim viewName As String
            viewName = Application.ActiveCell.Offset(0, 1).Value

            viewsService.DeleteView(listName, viewName)
        End Using
    End Sub
End Class
