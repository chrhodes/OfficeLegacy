Imports System.Xml

Public Class Sheet2

    Private Sub Sheet2_Startup(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Startup

    End Sub

    Private Sub Sheet2_Shutdown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shutdown

    End Sub

    Private Sub btnAddView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddView.Click
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
            Dim viewFields As XmlNode = ConvertToXmlNode(Application.Cells(currentRow, 3).Value)
            Dim viewQuery As XmlNode = ConvertToXmlNode(Application.Cells(currentRow, 4).Value)
            Dim rowLimit As XmlNode = ConvertToXmlNode(Application.Cells(currentRow, 5).Value)
            Dim viewType As String = Application.Cells(currentRow, 6).Value()
            Dim makeDefaultView As Boolean = CBool(Application.Cells(currentRow, 7).Value())

            viewsService.AddView(listName, viewName, viewFields, viewQuery, rowLimit, viewType, makeDefaultView)
        End Using
    End Sub

    Function ConvertToXmlNode(ByVal inputString As String) As XmlNode
        Dim outputXml As XmlDocument = New XmlDocument

        outputXml.InnerXml = inputString

        ' TODO: HOw to?

        Return outputXml
    End Function
End Class
