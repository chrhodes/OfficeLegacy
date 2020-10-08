Public Class TaskPane_WordUtil

    'Public IncludeCharts As Boolean = True
    'Public IncludeFeedback As Boolean = True

    'Private Sub btnClearDestinationSheets_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearDestinationSheets.Click
    '    'Dim AllTeamsScoreCardWS As Excel.Worksheet = Globals.ThisAddIn.Application.Sheets(Globals.cSN_ScoreCards_AllTeams)
    '    'Dim AllTeamsPartnerSurveyWS As Excel.Worksheet = Globals.ThisAddIn.Application.Sheets(Globals.cSN_PartnerSurvey_AllTeams)
    '    'Dim AllTeamsBusinessSurveyWS As Excel.Worksheet = Globals.ThisAddIn.Application.Sheets(Globals.cSN_BusinessSurvey_AllTeams)
    '    'Dim AllTeamsITSurveyWS As Excel.Worksheet = Globals.ThisAddIn.Application.Sheets(Globals.cSN_ITSurvey_AllTeams)

    '    'AllTeamsScoreCardWS.Range(Globals.cSCAT_ResultsCells).ClearContents()

    '    'AllTeamsPartnerSurveyWS.Range(Globals.cSRAT_SurveyResultsCells).ClearContents()
    '    'AllTeamsBusinessSurveyWS.Range(Globals.cSRAT_SurveyResultsCells).ClearContents()
    '    'AllTeamsITSurveyWS.Range(Globals.cSRAT_SurveyResultsCells).ClearContents()
    'End Sub

    'Private Sub btnCopyValues_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCopyValues.Click
    '    'Dim scorecardWS As Excel.Worksheet = Globals.ThisAddIn.Application.Sheets(Globals.cSN_ScoreCards_IndividualTeam)

    '    ''Util.ScreenUpdatesOff()

    '    'For Each teamRow As Data.DataRow In Config.Teams.Rows
    '    '    ' TODO: Should store the currently selected team

    '    '    Debug.Print(teamRow("name"))

    '    '    scorecardWS.Range(Globals.cSC_TeamNameCell).Value = teamRow("name")

    '    '    'Globals.ThisAddIn.ExcelUtil.CopyScorecardValuesToAllTeamsScorecardWorksheet()
    '    '    'Globals.ThisAddIn.ExcelUtil.CopySurveyValuesToAllTeamsSurveyWorksheets()

    '    'Next

    '    ''Util.ScreenUpdatesOn()
    'End Sub



    'Private Sub cbTeams_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbTeams.SelectedIndexChanged
    '    'Dim scorecardWS As Excel.Worksheet = Globals.ThisAddIn.Application.Sheets(Globals.cSN_ScoreCards_IndividualTeam)
    '    'Dim currentWS As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet

    '    '' Update the Team selected on the Individual Team Scorecard worksheet.  
    '    '' The other worksheets use this cell to determine which team is active.

    '    'scorecardWS.Range(Globals.cSC_TeamNameCell).Value = Me.cbTeams.SelectedItem.ToString()
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
End Class
