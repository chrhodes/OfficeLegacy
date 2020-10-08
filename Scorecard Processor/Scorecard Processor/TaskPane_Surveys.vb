Imports PacificLife.Life

' Most of the work is done in SurveyDataWorksheet and SurveyQuestionsWorksheet

Public Class TaskPane_Surveys

    Public IncludeCharts As Boolean = True
    Public IncludeFeedback As Boolean = True

    Private Sub btnLoadSurveyQuestions_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoadSurveyQuestions.Click
        SurveyQuestionsWorkSheet.LoadQuestions()
    End Sub

    Private Sub btnAddFormulas_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddFormulas.Click
        SurveyDataWorkSheet.AddFormulas()
    End Sub

    Private Sub btnCreateDataWorksheet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateDataWorksheet.Click
        SurveyDataWorkSheet.CreateWorksheet(Globals.ThisAddIn.Application.ActiveSheet.Name)
    End Sub

    Private Sub btnFormatWorksheet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFormatWorksheet.Click
        SurveyDataWorkSheet.FormatWorksheet()
    End Sub

    Private Sub btnLoadSurveyData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoadSurveyData.Click
        ' TODO: Push this method into
        Util.NewWorksheet(Me.cmbSurveyName.SelectedItem.ToString & " Data")
        ' This method so the user doesn't have to pick the survey, ugh.
        SurveyDataWorkSheet.LoadSurveyDataFromMDBFile()
    End Sub

    Private Sub btnAddCharts_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddCharts.Click
        Charts.DeleteChartsFromActiveWorkSheet()
        Charts.AddCharts()
    End Sub

    'Private Sub btnAddSurveyResultsToPowerPoint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddSurveyResultsToPowerPoint.Click
    '    PowerPointIntegration.AddSurveyResults(IncludeCharts, IncludeFeedback)
    'End Sub

    'Private Sub cbTeams_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Dim scorecardWS As Excel.Worksheet = Globals.ThisAddIn.Application.Sheets(Globals.cSN_ScoreCards_IndividualTeam)
    '    Dim currentWS As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet

    '    ' Update the Team selected on the Scorecard worksheet.  
    '    ' The other worksheets use this cell to determine which team is active.

    '    scorecardWS.Range(Globals.cSC_TeamNameCell).Value = Me.cbTeams.SelectedItem.ToString()

    '    ' Now ensure the Charts have the right info on them.

    '    Charts.UpdateChartTitles(currentWS.Name)
    'End Sub

    'Private Sub TaskPane_Data_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    '    For Each dataTable As Data.DataTable In Config.ConfigInfo.Tables
    '        Select Case dataTable.TableName
    '            Case "team"
    '                For Each dataRow As Data.DataRow In dataTable.Rows
    '                    Me.cbTeams.Items.Add(dataRow.Item("name")).ToString()
    '                Next
    '        End Select
    '    Next
    'End Sub

    'Private Sub btnAddSurveyResultsToWord_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddSurveyResultsToWord.Click
    '    Dim wdIntegration As WordIntegration = New WordIntegration

    '    wdIntegration.AddSurveyResults(IncludeCharts, IncludeFeedback)
    'End Sub

    'Private Sub chkIncludeCharts_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkIncludeCharts.CheckedChanged
    '    IncludeCharts = Me.chkIncludeCharts.Checked
    'End Sub

    'Private Sub chkIncludeFeedback_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkIncludeFeedback.CheckedChanged
    '    IncludeFeedback = Me.chkIncludeFeedback.Checked
    'End Sub

    'Private Sub btnProduceWordOutput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnProduceWordOutput.Click
    '    PLLog.Trace("Enter", "Scorecard")

    '    Dim wdIntegration As WordIntegration = New WordIntegration

    '    Dim scorecardWS As Excel.Worksheet = Globals.ThisAddIn.Application.Sheets(Globals.cSN_ScoreCards_IndividualTeam)
    '    Dim teamName As String

    '    For Each teamRow As Data.DataRow In Config.Teams.Rows
    '        teamName = teamRow("name")

    '        PLLog.Info("Producing Word output for team: " & teamName, "Scorecard")

    '        scorecardWS.Range(Globals.cSC_TeamNameCell).Value = teamRow("name")

    '        ' Not sure if we need to do this but are trying to ensure that the data
    '        ' has been refreshed to reflect the newly selected team.

    '        Globals.ThisAddIn.Application.CalculateFullRebuild()

    '        wdIntegration.AddSurveyResults(IncludeCharts, IncludeFeedback)
    '    Next

    '    PLLog.Trace("Exit", "Scorecard")
    'End Sub

    'Private Sub btnProducePowerPointOutput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnProducePowerPointOutput.Click
    '    PLLog.Trace("Enter", "Scorecard")

    '    Dim scorecardWS As Excel.Worksheet = Globals.ThisAddIn.Application.Sheets(Globals.cSN_ScoreCards_IndividualTeam)
    '    Dim teamName As String

    '    For Each teamRow As Data.DataRow In Config.Teams.Rows
    '        teamName = teamRow("name")

    '        PLLog.Info("Producing Word output for team: " & teamName, "Scorecard")

    '        scorecardWS.Range(Globals.cSC_TeamNameCell).Value = teamRow("name")

    '        ' Not sure if we need to do this but are trying to ensure that the data
    '        ' has been refreshed to reflect the newly selected team.

    '        Globals.ThisAddIn.Application.CalculateFullRebuild()

    '        PowerPointIntegration.AddSurveyResults(IncludeCharts, IncludeFeedback)
    '    Next

    '    PLLog.Trace("Exit", "Scorecard")
    'End Sub

    'Private Sub btnFormatAllSurveySheets_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFormatAllSurveySheets.Click
    '    PLLog.Trace("Enter", "Scorecard")

    '    Dim ws As Excel.Worksheet

    '    ws = Globals.ThisAddIn.Application.Sheets(Globals.cSN_PartnerSurvey_AllTeams)
    '    ConditionalyFormatRange(ws.Range(Globals.cSR_SurveyResultsCells))

    '    ws = Globals.ThisAddIn.Application.Sheets(Globals.cSN_BusinessSurvey_AllTeams)
    '    ConditionalyFormatRange(ws.Range(Globals.cSR_SurveyResultsCells))

    '    ws = Globals.ThisAddIn.Application.Sheets(Globals.cSN_ITSurvey_AllTeams)
    '    ConditionalyFormatRange(ws.Range(Globals.cSR_SurveyResultsCells))

    '    PLLog.Trace("Enter", "Scorecard")
    'End Sub

    'Private Sub ConditionalyFormatRange(ByRef rng As Excel.Range)
    '    Util.DisplayExcelRange(rng)
    '    rng.FormatConditions.AddColorScale(ColorScaleType:=3)
    '    rng.FormatConditions(rng.FormatConditions.Count).SetFirstPriority()

    '    With rng.FormatConditions(1)
    '        With .ColorScaleCriteria(1)
    '            .Type = Excel.XlConditionValueTypes.xlConditionValueLowestValue

    '            With .FormatColor()
    '                .Color = 7039480
    '                .TintAndShade = 0
    '            End With
    '        End With

    '        With .ColorScaleCriteria(2)
    '            .Type = Excel.XlConditionValueTypes.xlConditionValuePercentile
    '            .Value = 50

    '            With .FormatColor
    '                .Color = 8711167
    '                .TintAndShade = 0
    '            End With
    '        End With

    '        With .ColorScaleCriteria(3)
    '            .Type = Excel.XlConditionValueTypes.xlConditionValueHighestValue

    '            With .FormatColor
    '                .Color = 8109667
    '                .TintAndShade = 0
    '            End With
    '        End With
    '    End With
    'End Sub

End Class
