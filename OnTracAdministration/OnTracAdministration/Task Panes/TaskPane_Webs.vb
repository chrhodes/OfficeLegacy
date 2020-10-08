Imports System.ComponentModel
Imports System.Reflection
Imports System.Runtime

Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel

Public Class TaskPane_Webs
    'Private _teams As Data.DataSet

    Private _inputFilePath As String = Globals.cDEFAULT_ONTIMEDATA_FOLDER
    Private _paneWidth As Integer = 300
    Public Property PaneWidth() As Integer 
        Get
            Return _paneWidth
        End Get
        Set(ByVal Value As Integer )
            _paneWidth = Value
        End Set
    End Property


    Private Sub TaskPane_Webs_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '' Ensure that any config data we need is available.  Ok to call multiple times.
        ''Config.IntializeApplication()

        'For Each dataTable As Data.DataTable In Config.ConfigInfo.Tables
        '    'Debug.Print(dataTable.TableName)

        '    'For Each dataColumn As Data.DataColumn In dataTable.Columns
        '    '    Debug.Print(dataColumn.ColumnName)
        '    'Next

        '    Select Case dataTable.TableName
        '        Case "team"
        '            For Each dataRow As Data.DataRow In dataTable.Rows
        '                Me.clbOnTimeTeams.Items.Add(dataRow.Item("name")).ToString()
        '                Me.cbOnTimeTeams.Items.Add(dataRow.Item("name")).ToString()
        '                'Debug.Print(dataRow.Item("name").ToString())
        '                'Debug.Print(dataRow.Item("id").ToString())
        '                'Debug.Print(dataRow.Item("team_Id").ToString())
        '            Next

        '            'Case "manager"
        '            '    For Each dataRow As Data.DataRow In dataTable.Rows
        '            '        Debug.Print(dataRow.Item("manager_Text").ToString())
        '            '        Debug.Print(dataRow.Item("ext").ToString())
        '            '        Debug.Print(dataRow.Item("team_Id").ToString())
        '            '    Next

        '    End Select
        'Next
    End Sub


    Private Sub btnGetAllSubWebCollection_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetAllSubWebCollection.Click
        Dim rng As Microsoft.Office.Interop.Excel.Range
        Dim inputPathAndFileName As String
        Dim inputSheetName As String
        Dim inputWorkbook As Excel.Workbook
        Dim errorMessage As String = ""
        Dim ws As Excel.Worksheet = Util.NewWorksheet("All Sub Webs")

        Dim OnTracService As New ontrac1.Webs()
        Dim webServiceURL As String

        If txtURL.TextLength > 0 Then
            webServiceURL = txtURL.Text & "/_vti_bin/webs.asmx"
        Else
            MsgBox("Must provide URL")
            Return
        End If

        ' Use credentials of logged on user

        OnTracService.Credentials = System.Net.CredentialCache.DefaultCredentials

        ' or use specific credentials

        'Dim cache As System.Net.CredentialCache = New System.Net.CredentialCache()
        'cache.Add(New Uri(webService), "NTLM", New System.Net.NetworkCredential("pspappca", "Production2007", "PACIFICMUTUAL"))

        'OnTracService.Credentials = cache

        OnTracService.Url = webServiceURL

        Util.ScreenUpdatesOff()

        Util.AddColumnToSheet(ws, 1, 50, True, 5, "Title")
        Util.AddColumnToSheet(ws, 2, 75, True, 5, "Url")

        Dim websNode As System.Xml.XmlNode

        Try
            websNode = OnTracService.GetAllSubWebCollection() ' This throws and exception :(

            Dim i As Integer = 6

            For Each webNode As System.Xml.XmlNode In websNode
                ws.Cells(i, 1).Value = webNode.Attributes("Title").Value
                ws.Cells(i, 2).Value = webNode.Attributes("Url").Value
                i += 1
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        Util.ScreenUpdatesOn()

    End Sub

    Private Sub btnGetWebCollection_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetWebCollection.Click
        Dim rng As Microsoft.Office.Interop.Excel.Range
        Dim inputPathAndFileName As String
        Dim inputSheetName As String
        Dim inputWorkbook As Excel.Workbook
        Dim errorMessage As String = ""
        Dim ws As Excel.Worksheet = Util.NewWorksheet("Direct Sub Webs")

        Dim OnTracService As New ontrac1.Webs()
        Dim webServiceURL As String

        If txtURL.TextLength > 0 Then
            webServiceURL = txtURL.Text & "/_vti_bin/webs.asmx"
        Else
            MsgBox("Must provide URL")
            Return
        End If

        ' Use credentials of logged on user

        OnTracService.Credentials = System.Net.CredentialCache.DefaultCredentials

        ' or use specific credentials

        'Dim cache As System.Net.CredentialCache = New System.Net.CredentialCache()
        'cache.Add(New Uri(webService), "NTLM", New System.Net.NetworkCredential("pspappca", "Production2007", "PACIFICMUTUAL"))

        'OnTracService.Credentials = cache

        OnTracService.Url = webServiceURL

        Util.ScreenUpdatesOff()

        Util.AddColumnToSheet(ws, 1, 50, True, 5, "Title")
        Util.AddColumnToSheet(ws, 2, 75, True, 5, "Url")

        Dim websNode As System.Xml.XmlNode

        Try
            websNode = OnTracService.GetWebCollection() ' This throws an exception on http://ontrac :(

            Dim i As Integer = 6

            For Each webNode As System.Xml.XmlNode In websNode
                ws.Cells(i, 1).Value = webNode.Attributes("Title").Value
                ws.Cells(i, 2).Value = webNode.Attributes("Url").Value
                i += 1
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        Util.ScreenUpdatesOn()
    End Sub
End Class
