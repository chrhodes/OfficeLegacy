Imports PacificLife.Life

Public Class Results
    Public Shared Sub ProduceAllTeams_Scorecard()
        Dim resultsWB As Excel.Workbook
        Dim scorecardAllTeamsWS As Excel.Worksheet
        Dim partnerSurveyAllTeamsWS As Excel.Worksheet
        Dim businessSurveyAllTeamsWS As Excel.Worksheet
        Dim itSurveyAllTeamsWS As Excel.Worksheet

        Dim scorecardWB As Excel.Workbook
        Dim allTeamsWS As Excel.Worksheet

        Dim fileName As String

        Dim scorecardWS As Excel.Worksheet = Globals.ThisAddIn.Application.Sheets(Globals.cSN_ScoreCards_IndividualTeam)

        fileName = "IT Scorecard" & "-" & scorecardWS.Range(Globals.cSCIT_SurveyPeriodCell).Value

        ' TODO: Perhaps a list of common paths from config file.

        Dim workbookPath As String = Globals.ThisAddIn.Application.ActiveWorkbook.Path

        fileName = workbookPath & "\Team Scorecards\" & fileName

        resultsWB = Globals.ThisAddIn.Application.ActiveWorkbook
        scorecardAllTeamsWS = Globals.ThisAddIn.Application.Sheets(Globals.cSN_ScoreCards_AllTeams)
        partnerSurveyAllTeamsWS = Globals.ThisAddIn.Application.Sheets(Globals.cSN_PartnerSurvey_AllTeams)
        businessSurveyAllTeamsWS = Globals.ThisAddIn.Application.Sheets(Globals.cSN_BusinessSurvey_AllTeams)
        itSurveyAllTeamsWS = Globals.ThisAddIn.Application.Sheets(Globals.cSN_ITSurvey_AllTeams)

        ' First get the scorecard results for all teams

        scorecardAllTeamsWS.Range(Globals.cSCAT_Scorecard_Cells).Copy()

        scorecardWB = Globals.ThisAddIn.Application.Workbooks.Add()

        ' TODO: Figure out how to get rid of any sheets that get created.

        'allTeamsWS = scorecardWB.ActiveSheet

        allTeamsWS = Util.NewWorksheet("IT Scorecard - All Teams")

        ' and add it to the new workbook

        allTeamsWS.Range("$A$1").PasteSpecial( _
                            Paste:=Excel.XlPasteType.xlPasteValuesAndNumberFormats, _
                            Operation:=Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, _
                            SkipBlanks:=False, _
                            Transpose:=False)

        ' then paste the formats so things look nice

        allTeamsWS.Range("$A$1").PasteSpecial( _
                                    Paste:=Excel.XlPasteType.xlPasteFormats, _
                                    Operation:=Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, _
                                    SkipBlanks:=False, _
                                    Transpose:=False)

        ' then the column widths

        allTeamsWS.Range("$A$1").PasteSpecial( _
                                    Paste:=Excel.XlPasteType.xlPasteColumnWidths, _
                                    Operation:=Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, _
                                    SkipBlanks:=False, _
                                    Transpose:=False)

        ' Get the Partner Survey results

        partnerSurveyAllTeamsWS.Range(Globals.cSRAT_Scorecard_Cells).Copy()

        'With scorecardWB
        '    .Sheets.Add(After:=.Sheets(.Sheets.Count))
        'End With

        allTeamsWS = Util.NewWorksheet("Partner Survey")

        ' and add it to the new workbook

        allTeamsWS.Range("$A$1").PasteSpecial( _
                            Paste:=Excel.XlPasteType.xlPasteValuesAndNumberFormats, _
                            Operation:=Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, _
                            SkipBlanks:=False, _
                            Transpose:=False)

        ' then paste the formats so things look nice

        allTeamsWS.Range("$A$1").PasteSpecial( _
                                    Paste:=Excel.XlPasteType.xlPasteFormats, _
                                    Operation:=Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, _
                                    SkipBlanks:=False, _
                                    Transpose:=False)

        ' then the column widths

        allTeamsWS.Range("$A$1").PasteSpecial( _
                                    Paste:=Excel.XlPasteType.xlPasteColumnWidths, _
                                    Operation:=Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, _
                                    SkipBlanks:=False, _
                                    Transpose:=False)

        ' Get the Business Survey results

        businessSurveyAllTeamsWS.Range(Globals.cSRAT_Scorecard_Cells).Copy()

        'With scorecardWB
        '    .Sheets.Add(After:=.Sheets(.Sheets.Count))
        'End With

        allTeamsWS = Util.NewWorksheet("Business Survey")

        ' and add it to the new workbook

        allTeamsWS.Range("$A$1").PasteSpecial( _
                            Paste:=Excel.XlPasteType.xlPasteValuesAndNumberFormats, _
                            Operation:=Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, _
                            SkipBlanks:=False, _
                            Transpose:=False)

        ' then paste the formats so things look nice

        allTeamsWS.Range("$A$1").PasteSpecial( _
                                    Paste:=Excel.XlPasteType.xlPasteFormats, _
                                    Operation:=Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, _
                                    SkipBlanks:=False, _
                                    Transpose:=False)

        ' then the column widths

        allTeamsWS.Range("$A$1").PasteSpecial( _
                                    Paste:=Excel.XlPasteType.xlPasteColumnWidths, _
                                    Operation:=Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, _
                                    SkipBlanks:=False, _
                                    Transpose:=False)

        itSurveyAllTeamsWS.Range(Globals.cSRAT_Scorecard_Cells).Copy()

        'With scorecardWB
        '    .Sheets.Add(After:=.Sheets(.Sheets.Count))
        'End With

        allTeamsWS = Util.NewWorksheet("IT Survey")

        ' and add it to the new workbook

        allTeamsWS.Range("$A$1").PasteSpecial( _
                            Paste:=Excel.XlPasteType.xlPasteValuesAndNumberFormats, _
                            Operation:=Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, _
                            SkipBlanks:=False, _
                            Transpose:=False)

        ' then paste the formats so things look nice

        allTeamsWS.Range("$A$1").PasteSpecial( _
                                    Paste:=Excel.XlPasteType.xlPasteFormats, _
                                    Operation:=Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, _
                                    SkipBlanks:=False, _
                                    Transpose:=False)

        ' then the column widths

        allTeamsWS.Range("$A$1").PasteSpecial( _
                                    Paste:=Excel.XlPasteType.xlPasteColumnWidths, _
                                    Operation:=Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, _
                                    SkipBlanks:=False, _
                                    Transpose:=False)

        ' Excel8 is 2003 format.

        scorecardWB.SaveAs(Filename:=fileName.ToString, FileFormat:=Excel.XlFileFormat.xlExcel8)
        scorecardWB.Close(SaveChanges:=Excel.XlSaveAction.xlSaveChanges)
        resultsWB.Activate()
    End Sub

    Public Shared Sub ProduceAllTeams_Scorecards()
        Dim scorecardWS As Excel.Worksheet = Globals.ThisAddIn.Application.Sheets(Globals.cSN_ScoreCards_IndividualTeam)

        'Util.ScreenUpdatesOff()
        Globals.ThisAddIn.Application.DisplayStatusBar = True

        For Each teamRow As Data.DataRow In Config.Teams.Rows
            ' Set the scorecard to the team so values update
            scorecardWS.Range(Globals.cSC_TeamNameCell).Value = teamRow("name")

            SaveTeamScorecard(scorecardWS)
            Globals.ThisAddIn.Application.StatusBar = String.Format("Saved {0} team Scorecard", teamRow("name"))
        Next

        ' Return control of the status bar to Excel
        Globals.ThisAddIn.Application.StatusBar = False
    End Sub

    Public Shared Sub ProduceIndividualTeamScorecard()
        Dim scorecardWS As Excel.Worksheet = Globals.ThisAddIn.Application.Sheets(Globals.cSN_ScoreCards_IndividualTeam)

        'Util.ScreenUpdatesOff()
        Globals.ThisAddIn.Application.DisplayStatusBar = True

        ' The scorecardWS should have the correct team selected already.  This is handled
        ' in the cbTeams_SelectedIndexChanged() method.

        Results.SaveTeamScorecard(scorecardWS)
        Globals.ThisAddIn.Application.StatusBar = _
            String.Format("Saved {0} team Scorecard", scorecardWS.Range(Globals.cSC_TeamNameCell).Value)

        ' Return control of the status bar to Excel
        Globals.ThisAddIn.Application.StatusBar = False
    End Sub

    Public Shared Sub ProducePowerPointOutputOneTeam(ByVal includeCharts As Boolean, ByVal includeFeedback As Boolean)
        Dim wdIntegration As WordIntegration = New WordIntegration

        wdIntegration.AddSurveyResults(includeCharts, includeFeedback)
    End Sub

    Public Shared Sub ProducePowerPointOutputAllTeams(ByVal includeCharts As Boolean, ByVal includeFeedback As Boolean)
        PLLog.Trace("Enter", "Scorecard")

        Dim scorecardWS As Excel.Worksheet = Globals.ThisAddIn.Application.Sheets(Globals.cSN_ScoreCards_IndividualTeam)
        Dim teamName As String

        For Each teamRow As Data.DataRow In Config.Teams.Rows
            teamName = teamRow("name")

            PLLog.Info("Producing Word output for team: " & teamName, "Scorecard")

            scorecardWS.Range(Globals.cSC_TeamNameCell).Value = teamRow("name")

            ' Not sure if we need to do this but are trying to ensure that the data
            ' has been refreshed to reflect the newly selected team.

            Globals.ThisAddIn.Application.CalculateFullRebuild()

            PowerPointIntegration.AddSurveyResults(IncludeCharts, IncludeFeedback)
        Next

        PLLog.Trace("Exit", "Scorecard")
    End Sub

    Public Shared Sub ProduceWordOutputOneTeam(ByVal includeCharts As Boolean, ByVal includeFeedback As Boolean)
        Dim wdIntegration As WordIntegration = New WordIntegration

        wdIntegration.AddSurveyResults(includeCharts, includeFeedback)
    End Sub

    Public Shared Sub ProduceWordOutputAllTeams(ByVal includeCharts As Boolean, ByVal includeFeedback As Boolean)
        PLLog.Trace("Enter", "Scorecard")

        Dim wdIntegration As WordIntegration = New WordIntegration

        Dim scorecardWS As Excel.Worksheet = Globals.ThisAddIn.Application.Sheets(Globals.cSN_ScoreCards_IndividualTeam)
        Dim teamName As String

        For Each teamRow As Data.DataRow In Config.Teams.Rows
            teamName = teamRow("name")

            PLLog.Info("Producing Word output for team: " & teamName, "Scorecard")

            scorecardWS.Range(Globals.cSC_TeamNameCell).Value = teamRow("name")

            ' Not sure if we need to do this but are trying to ensure that the data
            ' has been refreshed to reflect the newly selected team.

            Globals.ThisAddIn.Application.CalculateFullRebuild()

            wdIntegration.AddSurveyResults(includeCharts, IncludeFeedback)
        Next

        PLLog.Trace("Exit", "Scorecard")
    End Sub

    Public Shared Sub SaveTeamScorecard(ByVal scorecardWS As Excel.Worksheet)
        Dim teamWB As Excel.Workbook
        Dim teamWS As Excel.Worksheet

        Debug.Print(scorecardWS.Name)
        Dim fileName As String

        fileName = "TeamScorecard" _
            & "-" & scorecardWS.Range(Globals.cSCIT_SurveyPeriodCell).Value _
            & " - " & scorecardWS.Range(Globals.cSCIT_TeamNameCell).Value

        ' TODO: Perhaps a list of common paths from config file.

        Dim workbookPath As String = Globals.ThisAddIn.Application.ActiveWorkbook.Path

        fileName = workbookPath & "\Team Scorecards\" & fileName

        'scorecardWS.Range("A2:L42").Select()
        scorecardWS.Range("A2:L42").Copy()
        'Range("K42").Activate()
        'Selection.Copy()
        teamWB = Globals.ThisAddIn.Application.Workbooks.Add()
        teamWS = teamWB.ActiveSheet

        ' First paste the data

        teamWS.Range("$A$1").PasteSpecial( _
                                    Paste:=Excel.XlPasteType.xlPasteValuesAndNumberFormats, _
                                    Operation:=Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, _
                                    SkipBlanks:=False, _
                                    Transpose:=False)

        ' then paste the formats so things look nice

        teamWS.Range("$A$1").PasteSpecial( _
                                    Paste:=Excel.XlPasteType.xlPasteFormats, _
                                    Operation:=Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, _
                                    SkipBlanks:=False, _
                                    Transpose:=False)

        ' then the column widths

        teamWS.Range("$A$1").PasteSpecial( _
                                    Paste:=Excel.XlPasteType.xlPasteColumnWidths, _
                                    Operation:=Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, _
                                    SkipBlanks:=False, _
                                    Transpose:=False)

        Util.ScreenUpdatesOff()

        With teamWS.PageSetup
            .PrintArea = "$A$1:$L$41"
            .LeftHeader = ""
            .CenterHeader = ""
            .RightHeader = ""
            .LeftFooter = ""
            .CenterFooter = ""
            .RightFooter = ""
            .LeftMargin = Globals.ThisAddIn.Application.InchesToPoints(0.7)
            .RightMargin = Globals.ThisAddIn.Application.InchesToPoints(0.7)
            .TopMargin = Globals.ThisAddIn.Application.InchesToPoints(0.75)
            .BottomMargin = Globals.ThisAddIn.Application.InchesToPoints(0.75)
            .HeaderMargin = Globals.ThisAddIn.Application.InchesToPoints(0.3)
            .FooterMargin = Globals.ThisAddIn.Application.InchesToPoints(0.3)
            .PrintHeadings = False
            .PrintGridlines = False
            .PrintComments = Excel.XlPrintLocation.xlPrintNoComments
            .PrintQuality = 600
            .CenterHorizontally = False
            .CenterVertically = False
            .Orientation = Excel.XlPageOrientation.xlPortrait
            .Draft = False
            .PaperSize = Excel.XlPaperSize.xlPaperLetter
            .FirstPageNumber = Excel.Constants.xlAutomatic
            .Order = Excel.XlOrder.xlDownThenOver
            .BlackAndWhite = False
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 1
            .PrintErrors = Excel.XlPrintErrors.xlPrintErrorsDisplayed
            .OddAndEvenPagesHeaderFooter = False
            .DifferentFirstPageHeaderFooter = False
            .ScaleWithDocHeaderFooter = True
            .AlignMarginsHeaderFooter = True
            .EvenPage.LeftHeader.Text = ""
            .EvenPage.CenterHeader.Text = ""
            .EvenPage.RightHeader.Text = ""
            .EvenPage.LeftFooter.Text = ""
            .EvenPage.CenterFooter.Text = ""
            .EvenPage.RightFooter.Text = ""
            .FirstPage.LeftHeader.Text = ""
            .FirstPage.CenterHeader.Text = ""
            .FirstPage.RightHeader.Text = ""
            .FirstPage.LeftFooter.Text = ""
            .FirstPage.CenterFooter.Text = ""
            .FirstPage.RightFooter.Text = ""
        End With

        ' Excel8 is 2003 format.

        teamWB.SaveAs(Filename:=fileName.ToString, FileFormat:=Excel.XlFileFormat.xlExcel8)
        teamWB.Close(SaveChanges:=Excel.XlSaveAction.xlSaveChanges)

        Util.ScreenUpdatesOn()
    End Sub

    Public Shared Sub FormatAllSurveySheets()
        PLLog.Trace("Enter", "Scorecard")

        Dim ws As Excel.Worksheet

        ws = Globals.ThisAddIn.Application.Sheets(Globals.cSN_PartnerSurvey_AllTeams)
        ConditionalyFormatRange(ws.Range(Globals.cSRAT_SurveyResultsCells))

        ws = Globals.ThisAddIn.Application.Sheets(Globals.cSN_BusinessSurvey_AllTeams)
        ConditionalyFormatRange(ws.Range(Globals.cSRAT_SurveyResultsCells))

        ws = Globals.ThisAddIn.Application.Sheets(Globals.cSN_ITSurvey_AllTeams)
        ConditionalyFormatRange(ws.Range(Globals.cSRAT_SurveyResultsCells))

        PLLog.Trace("Enter", "Scorecard")
    End Sub

    Private Shared Sub ConditionalyFormatRange(ByRef rng As Excel.Range)
        Util.DisplayExcelRange(rng)
        rng.FormatConditions.AddColorScale(ColorScaleType:=3)
        rng.FormatConditions(rng.FormatConditions.Count).SetFirstPriority()

        With rng.FormatConditions(1)
            With .ColorScaleCriteria(1)
                .Type = Excel.XlConditionValueTypes.xlConditionValueLowestValue

                With .FormatColor()
                    .Color = 7039480
                    .TintAndShade = 0
                End With
            End With

            With .ColorScaleCriteria(2)
                .Type = Excel.XlConditionValueTypes.xlConditionValuePercentile
                .Value = 50

                With .FormatColor
                    .Color = 8711167
                    .TintAndShade = 0
                End With
            End With

            With .ColorScaleCriteria(3)
                .Type = Excel.XlConditionValueTypes.xlConditionValueHighestValue

                With .FormatColor
                    .Color = 8109667
                    .TintAndShade = 0
                End With
            End With
        End With
    End Sub
End Class
