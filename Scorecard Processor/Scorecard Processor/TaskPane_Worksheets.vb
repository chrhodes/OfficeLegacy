Public Class TaskPane_Worksheets

    Private Sub btnCreatePage_OnTimeDeliveryData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateOnTimeDeliveryWorksheet.Click
        CreateSheet.NewSheet(Globals.cSN_OnTimeDeliveryData)
    End Sub

    Private Sub btnCreatePage_BudgetVariance_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateBudgetVarianceWorksheet.Click
        CreateSheet.NewSheet(Globals.cSN_BudgetVarianceData)
    End Sub

    Private Sub btnCreatePage_ITRProcessing_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateITRProcessingWorksheet.Click
        CreateSheet.NewSheet(Globals.cSN_ITRProcessingData)
    End Sub

    'Private Sub cbTeams_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Dim scorecardWS As Excel.Worksheet = Globals.ThisAddIn.Application.Sheets(Globals.cSN_ScoreCards_IndividualTeam)
    '    Dim currentWS As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet

    '    ' Update the Team selected on the Scorecard worksheet.  
    '    ' The other worksheets use this cell to determine which team is active.

    '    scorecardWS.Range(Globals.cSC_TeamNameCell).Value = Me.cbTeams.SelectedItem.ToString()
    'End Sub

    Private Sub btnCreateScoreCards_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateScoreCardWorksheets.Click
        CreateSheet.NewSheet(Globals.cSN_ScoreCards_IndividualTeam)
        CreateSheet.NewSheet(Globals.cSN_ScoreCards_AllTeams)
        CreateSheet.NewSheet(Globals.cSN_ScoreCard_PreviousPeriod)
    End Sub

    'Private Sub btnCopyValues_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
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

    'Private Sub TaskPane_Worksheets_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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

    'Private Sub btnClearDestinationSheets_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearDestinationSheets.Click
    '    Dim AllTeamsScoreCardWS As Excel.Worksheet = Globals.ThisAddIn.Application.Sheets(Globals.cSN_ScoreCards_AllTeams)
    '    Dim AllTeamsPartnerSurveyWS As Excel.Worksheet = Globals.ThisAddIn.Application.Sheets(Globals.cSN_PartnerSurveyAllTeams)
    '    Dim AllTeamsBusinessSurveyWS As Excel.Worksheet = Globals.ThisAddIn.Application.Sheets(Globals.cSN_BusinessSurveyAllTeams)
    '    Dim AllTeamsITSurveyWS As Excel.Worksheet = Globals.ThisAddIn.Application.Sheets(Globals.cSN_ITSurveyAllTeams)

    '    AllTeamsScoreCardWS.Range(Globals.cSCAT_ResultsCells).ClearContents()

    '    AllTeamsPartnerSurveyWS.Range(Globals.cSR_SurveyResultsCells).ClearContents()
    '    AllTeamsBusinessSurveyWS.Range(Globals.cSR_SurveyResultsCells).ClearContents()
    '    AllTeamsITSurveyWS.Range(Globals.cSR_SurveyResultsCells).ClearContents()
    'End Sub

    'Private Sub btnCreateTeamScorecards_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateTeamScorecards.Click
    '    Dim scorecardWS As Excel.Worksheet = Globals.ThisAddIn.Application.Sheets(Globals.cSN_ScoreCards_IndividualTeam)

    '    'Util.ScreenUpdatesOff()
    '    Globals.ThisAddIn.Application.DisplayStatusBar = True

    '    For Each teamRow As Data.DataRow In Config.Teams.Rows
    '        ' Set the scorecard to the team so values update
    '        scorecardWS.Range(Globals.cSC_TeamNameCell).Value = teamRow("name")

    '        SaveTeamScorecard(scorecardWS)
    '        Globals.ThisAddIn.Application.StatusBar = String.Format("Saved {0} team Scorecard", teamRow("name"))
    '    Next

    '    ' Return control of the status bar to Excel
    '    Globals.ThisAddIn.Application.StatusBar = False
    'End Sub

    Private Sub btnCreateRollupWorksheets_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateRollupWorksheets.Click
        CreateSheet.NewSheet(Globals.cSN_ScoreCards_BudgetRollUp)
        CreateSheet.NewSheet(Globals.cSN_ScoreCards_SurveyRollUp)
    End Sub

    Private Sub btnCreateSurveyMappingWorksheet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateSurveyMappingWorksheet.Click
        CreateSheet.NewSheet(Globals.cSN_ScoreCards_SurveyMapping)
    End Sub

    Private Sub btnCreateSurveyWorksheets_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateSurveyWorksheets.Click
        CreateSheet.NewSheet(Globals.cSN_PartnerSurvey_AllTeams)
        CreateSheet.NewSheet(Globals.cSN_BusinessSurvey_AllTeams)
        CreateSheet.NewSheet(Globals.cSN_ITSurvey_AllTeams)
    End Sub
End Class
