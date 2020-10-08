Imports PacificLife.Life

Public Class CreateSheet

#Region "Public Methods"

    Public Shared Function NewSheet(ByVal sheetType As String) As Excel.Worksheet
        PLLog.Trace1("Enter", "Scorecard")

        Dim ws As Excel.Worksheet = Nothing

        Select Case sheetType
            Case Globals.cSN_ScoreCards_IndividualTeam
                MessageBox.Show("Not implemented yet")

            Case Globals.cSN_ScoreCards_AllTeams
                MessageBox.Show("Not implemented yet")

            Case Globals.cSN_ScoreCard_PreviousPeriod
                MessageBox.Show("Not implemented yet")

            Case Globals.cSN_ScoreCards_BudgetRollUp
                MessageBox.Show("Not implemented yet")

            Case Globals.cSN_ScoreCards_SurveyRollUp
                MessageBox.Show("Not implemented yet")

            Case Globals.cSN_ScoreCards_SurveyMapping
                MessageBox.Show("Not implemented yet")

            Case Globals.cSN_PartnerSurvey_AllTeams
                MessageBox.Show("Not implemented yet")

            Case Globals.cSN_BusinessSurvey_AllTeams
                MessageBox.Show("Not implemented yet")

            Case Globals.cSN_ITSurvey_AllTeams
                MessageBox.Show("Not implemented yet")

            Case Globals.cSN_PartnerSurveyQuestions
                'ws = CreatePartnerSurveyQuestionsSheet(Globals.cPartnerSurveyQuestionsSheetName)

            Case Globals.cSN_PartnerSurveyData
                ws = SurveyDataWorkSheet.CreateWorksheet(Globals.cSN_PartnerSurveyData)

            Case Globals.cSN_BusinessSurveyQuestions
                'ws = CreateBusinessSurveyQuestionsSheet(Globals.cBusinessSurveyQuestionsSheetName)

            Case Globals.cSN_BusinessSurveyData
                ws = SurveyDataWorkSheet.CreateWorksheet(Globals.cSN_BusinessSurveyData)

            Case Globals.cSN_ITSurveyQuestions
                'ws = CreateITSurveyQuestionsSheet(Globals.cITSurveyQuestionsSheetName)

            Case Globals.cSN_ITSurveyData
                ws = SurveyDataWorkSheet.CreateWorksheet(Globals.cSN_ITSurveyData)

            Case Globals.cSN_HelpDeskSurveyQuestions
                'ws = CreateHelpDeskSurveyQuestionsSheet(Globals.cHelpDeskSurveyQuestionsSheetName)

            Case Globals.cSN_BudgetVarianceData
                ws = BudgetVarianceWorkSheet.CreateWorkSheet(Globals.cSN_BudgetVarianceData)

            Case Globals.cSN_ITRProcessingData
                ws = ITRProcessingWorkSheet.CreateWorkSheet(Globals.cSN_ITRProcessingData)

            Case Globals.cSN_OnTimeDeliveryData
                ws = OnTimeDeliveryWorkSheet.CreateWorkSheet(Globals.cSN_OnTimeDeliveryData)

            Case Else
                MessageBox.Show("Unknown SheetType")
                ws = Nothing

        End Select

        ' Todo: Validate we got a good ws back

        PLLog.Trace1("Exit", "Scorecard")

        Return ws
    End Function

#End Region

End Class
