Imports System.Xml

Public Class Sheet3

    Private Sub Sheet3_Startup(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Startup

    End Sub

    Private Sub Sheet3_Shutdown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shutdown

    End Sub

    Private Sub btnUpdateView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdateView.Click
        Using viewsService As New learnsharepoint_ITTechSvc.Views()
            viewsService.Credentials = System.Net.CredentialCache.DefaultCredentials
            Dim viewsWebServiceUrl As String = "http://learnsharepoint/teams/ITTechSvc/_vti_bin/Views.asmx"
            Dim currentRow As Integer

            'Dim cache As System.Net.CredentialCache = New System.Net.CredentialCache()

            'cache.Add(New Uri(viewsWebServiceUrl), "NTLM", New System.Net.NetworkCredential("crhodes", "HappyH0jnacki8", "PACIFICMUTUAL"))
            'viewsService.Credentials = cache

            viewsService.Url = viewsWebServiceUrl

            currentRow = Application.ActiveCell.Row

            Dim listName As String = Application.Cells(currentRow, 1).Value
            Dim viewName As String = Application.Cells(currentRow, 2).Value
            Dim viewProperties As XmlNode = Nothing ' ConvertToXmlNode(Application.Cells(currentRow, 3).Value)
            Dim viewQuery As XmlNode = ConvertToXmlNode(Application.Cells(currentRow, 4).Value)
            Dim viewFields As XmlNode = ConvertToXmlNode(Application.Cells(currentRow, 5).Value)
            Dim viewAggregations As XmlNode = Nothing ' ConvertToXmlNode(Application.Cells(currentRow, 6).Value)
            Dim viewFormats As XmlNode = Nothing ' ConvertToXmlNode(Application.Cells(currentRow, 7).Value)
            Dim viewRowLimit As XmlNode = ConvertToXmlNode(Application.Cells(currentRow, 8).Value)

            Dim viewXmlNode As XmlNode

            viewXmlNode = viewsService.UpdateView(listName, viewName, viewProperties, viewQuery, viewFields, viewAggregations, viewFormats, viewRowLimit)

        End Using
    End Sub

    Function ConvertToXmlNode(ByVal inputString As String) As XmlNode
        Dim outputXml As XmlDocument = New XmlDocument

        outputXml.InnerXml = inputString

        ' TODO: HOw to?

        Return outputXml
    End Function
End Class
