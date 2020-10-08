Public Class TaskPane_UserAdmin
    Private _paneWidth As Integer = 300

    Public Property PaneWidth() As Integer
        Get
            Return _paneWidth
        End Get
        Set(ByVal Value As Integer)
            _paneWidth = Value
        End Set
    End Property

    'Public IncludeCharts As Boolean = True
    'Public IncludeFeedback As Boolean = True

    'Private Sub btnClearDestinationSheets_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearDestinationSheets.Click
    '    Dim AllTeamsScoreCardWS As Excel.Worksheet = Globals.ThisAddIn.Application.Sheets(Globals.cSN_ScoreCards_AllTeams)
    '    Dim AllTeamsPartnerSurveyWS As Excel.Worksheet = Globals.ThisAddIn.Application.Sheets(Globals.cSN_PartnerSurvey_AllTeams)
    '    Dim AllTeamsBusinessSurveyWS As Excel.Worksheet = Globals.ThisAddIn.Application.Sheets(Globals.cSN_BusinessSurvey_AllTeams)
    '    Dim AllTeamsITSurveyWS As Excel.Worksheet = Globals.ThisAddIn.Application.Sheets(Globals.cSN_ITSurvey_AllTeams)

    '    AllTeamsScoreCardWS.Range(Globals.cSCAT_ResultsCells).ClearContents()

    '    AllTeamsPartnerSurveyWS.Range(Globals.cSRAT_SurveyResultsCells).ClearContents()
    '    AllTeamsBusinessSurveyWS.Range(Globals.cSRAT_SurveyResultsCells).ClearContents()
    '    AllTeamsITSurveyWS.Range(Globals.cSRAT_SurveyResultsCells).ClearContents()
    'End Sub

    'Private Sub btnCopyValues_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCopyValues.Click
    '    Dim scorecardWS As Excel.Worksheet = Globals.ThisAddIn.Application.Sheets(Globals.cSN_ScoreCards_IndividualTeam)

    '    'Util.ScreenUpdatesOff()

    '    For Each teamRow As Data.DataRow In Config.Teams.Rows
    '        ' TODO: Should store the currently selected team

    '        Debug.Print(teamRow("name"))

    '        scorecardWS.Range(Globals.cSC_TeamNameCell).Value = teamRow("name")

    '        Util.CopyScorecardValuesToAllTeamsScorecardWorksheet()
    '        Util.CopySurveyValuesToAllTeamsSurveyWorksheets()

    '    Next

    '    'Util.ScreenUpdatesOn()
    'End Sub

    'Private Sub cbTeams_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbTeams.SelectedIndexChanged
    '    Dim scorecardWS As Excel.Worksheet = Globals.ThisAddIn.Application.Sheets(Globals.cSN_ScoreCards_IndividualTeam)
    '    Dim currentWS As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet

    '    ' Update the Team selected on the Individual Team Scorecard worksheet.  
    '    ' The other worksheets use this cell to determine which team is active.

    '    scorecardWS.Range(Globals.cSC_TeamNameCell).Value = Me.cbTeams.SelectedItem.ToString()
    'End Sub

    'Private Sub TaskPane_Results_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    '    ' Ensure that any config data we need is available.  Ok to call multiple times.
    '    'Config.IntializeApplication()

    '    For Each dataTable As Data.DataTable In Config.ConfigInfo.Tables
    '        Select Case dataTable.TableName
    '            Case "team"
    '                For Each dataRow As Data.DataRow In dataTable.Rows
    '                    Me.cbTeams.Items.Add(dataRow.Item("name")).ToString()
    '                Next
    '        End Select
    '    Next
    'End Sub

    'Private Sub btnProduceAllTeamsScorecards_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnProduceAllTeamsScorecards.Click
    '    'Results.ProduceAllTeams_Scorecards()
    'End Sub

    'Private Sub btnProduceIndividualTeamScorecard_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnProduceIndividualTeamScorecard.Click
    '    'Results.ProduceIndividualTeamScorecard()
    'End Sub

    'Private Sub btnAddSurveyResultsToPowerPointIndividualTeam_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddSurveyResultsToPowerPointIndividualTeam.Click
    '    'Results.ProducePowerPointOutputOneTeam(IncludeCharts, IncludeFeedback)
    'End Sub

    'Private Sub btnAddSurveyResultsToWordIndividualTeam_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddSurveyResultsToWordIndividualTeam.Click
    '    'Results.ProduceWordOutputOneTeam(IncludeCharts, IncludeFeedback)
    'End Sub

    'Private Sub btnAddSurveyResultsToPowerPointAllTeams_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddSurveyResultsToPowerPointAllTeams.Click
    '    'Results.ProducePowerPointOutputAllTeams(IncludeCharts, IncludeFeedback)
    'End Sub

    'Private Sub btnAllSurveyResultsToWordAllTeams_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAllSurveyResultsToWordAllTeams.Click
    '    'Results.ProduceWordOutputAllTeams(IncludeCharts, IncludeFeedback)
    'End Sub

    'Private Sub chkIncludeFeedback_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkIncludeFeedback.CheckedChanged
    '    IncludeFeedback = Me.chkIncludeFeedback.Checked
    'End Sub

    'Private Sub chkIncludeCharts_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkIncludeCharts.CheckedChanged
    '    IncludeCharts = Me.chkIncludeCharts.Checked
    'End Sub

    'Private Sub btnFormatAllSurveySheets_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFormatAllSurveySheets.Click
    '    'Results.FormatAllSurveySheets()
    'End Sub

    'Private Sub btnProduceAllTeamsScorecard_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnProduceAllTeamsScorecard.Click
    '    'Results.ProduceAllTeams_Scorecard()
    'End Sub

    Private Sub btnGetSiteCollectionUsers_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetSiteCollectionUsers.Click
        Dim rng As Microsoft.Office.Interop.Excel.Range
        Dim inputPathAndFileName As String
        Dim inputSheetName As String
        Dim inputWorkbook As Excel.Workbook
        Dim errorMessage As String = ""
        Dim ws As Excel.Worksheet = Util.NewWorksheet("Users from Site")
        ws.Cells.Clear()

        Dim OnTracService As New ontrac.UserGroup()

        Dim webServiceURL As String

        If txtURL.TextLength > 0 Then
            webServiceURL = txtURL.Text & "/_vti_bin/usergroup.asmx"
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



        Try
            Dim node As System.Xml.XmlNode = OnTracService.GetUserCollectionFromSite()
            Dim innerNode As System.Xml.XmlNode = node.FirstChild()

            Dim i As Integer = 6

            For Each userNode As System.Xml.XmlNode In innerNode
                'Me.cbUsers.Items.Add(New DictionaryEntry(userNode.Attributes("ID").Value(), userNode.Attributes("Name").Value))
                cbSiteCollectionUsers.Items.Add(New KeyValuePair(userNode.Attributes("ID").Value, userNode.Attributes("Name").Value))
            Next

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        Util.ScreenUpdatesOn()
    End Sub

    Private Sub cbSiteCollectionUsers_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSiteCollectionUsers.SelectedIndexChanged
        Dim kvp As KeyValuePair = cbSiteCollectionUsers.SelectedItem

        txtSiteCollectionUserName.Text = kvp.m_value
        txtSiteCollectionUserID.Text = kvp.m_key
    End Sub

    Private Sub btnFindSitesWithUser_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindSitesWithUser.Click

    End Sub

    Private Sub TaskPane_UserAdmin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
       
    End Sub

    Private Sub btnGetAllSubWebs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetAllSubWebs.Click
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
                cbWebs.Items.Add(New KeyValuePair(webNode.Attributes("Url").Value, webNode.Attributes("Title").Value))

                ws.Cells(i, 1).Value = webNode.Attributes("Title").Value
                ws.Cells(i, 2).Value = webNode.Attributes("Url").Value
                i += 1
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        Util.ScreenUpdatesOn()
    End Sub

    Private Sub cbWebs_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbWebs.SelectedIndexChanged
        Dim kvp As KeyValuePair = cbWebs.SelectedItem

        txtTitle.Text = kvp.m_value
        txtWebURL.Text = kvp.m_key
    End Sub

    Private Sub btnGetWebUsers_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetWebUsers.Click
        Dim rng As Microsoft.Office.Interop.Excel.Range
        Dim inputPathAndFileName As String
        Dim inputSheetName As String
        Dim inputWorkbook As Excel.Workbook
        Dim errorMessage As String = ""
        'Dim ws As Excel.Worksheet = Util.NewWorksheet("Users from Site")
        'ws.Cells.Clear()

        Dim OnTracService As New ontrac.UserGroup()

        Dim webServiceURL As String

        If txtURL.TextLength > 0 Then
            webServiceURL = txtURL.Text & "/_vti_bin/usergroup.asmx"
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

        Try
            Dim node As System.Xml.XmlNode = OnTracService.GetUserCollectionFromWeb()
            Dim innerNode As System.Xml.XmlNode = node.FirstChild()

            For Each userNode As System.Xml.XmlNode In innerNode
                cbSiteCollectionUsers.Items.Add(New KeyValuePair(userNode.Attributes("ID").Value, userNode.Attributes("Name").Value))
            Next

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        Util.ScreenUpdatesOn()
    End Sub

    Private Sub cbWebUsers_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbWebUsers.SelectedIndexChanged
        Dim kvp As KeyValuePair = cbWebUsers.SelectedItem

        txtWebUserName.Text = kvp.m_value
        txtWebUserID.Text = kvp.m_key
    End Sub
End Class
