Partial Friend NotInheritable Class Globals
    '**********************************************************************
    '   P u b l i c    C o n s t a n t s
    '**********************************************************************

    Public Const cPROJECT_NAME As String = "ScoreCard Processing"
    Public Const cPROJECT_VERSION As String = "1.0.0"
    Public Const cDATA_VERSION As String = "1.0.0"
    Public Const cCHART_VERSION As String = "1.0.0"
    Public Const cPLLOG_NAME As String = "NEWAPPNAME"

    ' TODO: Perhaps load these from Config File.

    Public Const cDEFAULT_SURVEY_FOLDER As String = "M:\Technology Office\IT Balanced Scorecard\Surveys\V2 Surveys\"
    Public Const cDEFAULT_ONTIMEDATA_FOLDER As String = "J:\LDM"
    Public Const cNumberTeams As Integer = 10

    Public Const cKB As Double = 1024
    Public Const cMB As Double = cKB * cKB
    Public Const cGB As Double = cMB * cKB

    'Public Const cDataEntryCell As Integer = cGreen

    ' Charting Colors

    'Public Const cBlack As Integer = 1
    'Public Const cColor2 As Integer = 2
    'Public Const cColor3 As Integer = 6
    'Public Const cRed As Integer = 3
    'Public Const cGreen As Integer = 4
    'Public Const cBlue As Integer = 5
    'Public Const cPink As Integer = 7

    'Public Const cOrange As Integer = 46

    'Public Const cLT_TURQUOISE As Integer = 34
    'Public Const cLT_GREEN As Integer = 35
    'Public Const cROSE As Integer = 38
    'Public Const cLT_YELLOW As Integer = 36
    'Public Const cTAN As Integer = 40
    'Public Const cGOLD As Integer = 44

#Region "Debug constants"

    Public Shared cScreenUpdatesOff As Boolean = True

#End Region

    '------------------------------------------------------------
    '   Constants to control the Worksheets
    '------------------------------------------------------------

#Region "Scorecard Constants"
    ' Scorecard constants start with "cSC_"
    Public Const cSC_TeamNameCell As String = "$B$2"

#End Region

#Region "On-Time Delivery worksheet (cOTD_) constants"

    Public Const cOTD_MetricName As String = " OT Data"

    ' Cell addresses

    Public Const cOTD_MetricName_Cell As String = "$C$4"
    Public Const cOTD_SurveyPeriod_Cell As String = "$C$5"
    Public Const cOTD_InputFile_Cell As String = "$C$6"
    Public Const cOTD_InputSheet_Cell As String = "$C$7"

    Public Const cOTD_ManagerName_Cell As String = "$C$11"
    Public Const cOTD_TeamName_Cell As String = "$C$12"

    Public Const cOTD_StartDataRow_Cell As String = "$C$13"
    Public Const cOTD_EndDataRow_Cell As String = "$C$14"
    Public Const cOTD_StartDataColumn_Cell As String = "$C$15"
    Public Const cOTD_EndDataColumn_Cell As String = "$C$16"

    Public Const cOTD_WeightedScheduledOnTimePercentage_Cell As String = "$C$24"
    Public Const cOTD_WeightedActualOnTimePercentage_Cell As String = "$C$25"
    Public Const cOTD_OnTimePercentage_Cell As String = "$C$26"
    Public Const cOTD_NumberReleases_Cell As String = "$C$27"
    Public Const cOTD_TotalActualReleaseDays_Cell As String = "$C$28"
    Public Const cOTD_TotalScheduledReleaseDays_Cell As String = "$C$29"

    Public Const cOTD_DataColumn As Integer = 7

    ' Offsets from file name on On-Time Data worksheet

    Public Const cOTD_TeamName_Offset As Integer = -8
    Public Const cOTD_Score_Offset As Integer = -7
    Public Const cOTD_WeightedScheduledOnTimePercentage_Offset As Integer = -6
    Public Const cOTD_WeightedActualOnTimePercentage_Offset As Integer = -5
    Public Const cOTD_OnTimePercentage_Offset As Integer = -4
    Public Const cOTD_NumberReleases_Offset As Integer = -3
    Public Const cOTD_Manager_Offset As Integer = -2
    Public Const cOTD_Extension_Offset As Integer = -1

    Public Const cOTD_SheetName_Offset As Integer = 1
    Public Const cOTD_DataSheet_Offset As Integer = 3

    Public Const cOTD_ChartTop As Integer = 10
    Public Const cOTD_ChartLeft As Integer = 600

#End Region

#Region "All Sheet constants"

    Public Const cRawDataCell As String = "$A$30"
    Public Const cAllITString As String = "All IT"

#End Region

#Region "Sheet Names (cSN_) constants"

    '------------------------------------------------------------
    '   Constants to access the various sheets with names we
    '   want to keep constant because they are referenced in
    '   the code.
    '------------------------------------------------------------

    Public Const cSN_ScoreCards_IndividualTeam As String = "Scorecard - Individual Team"
    Public Const cSN_ScoreCards_AllTeams As String = "Scorecard - All Teams"
    Public Const cSN_ScoreCard_PreviousPeriod As String = "Scorecard - Previous Period"

    Public Const cSN_ScoreCards_BudgetRollUp As String = "Budget Rollup"
    Public Const cSN_ScoreCards_SurveyRollUp As String = "Survey Rollup"
    Public Const cSN_ScoreCards_SurveyMapping As String = "Survey Mapping"

    Public Const cSN_PartnerSurvey_AllTeams As String = "Partner Survey - All Teams"
    Public Const cSN_BusinessSurvey_AllTeams As String = "Business Survey - All Teams"
    Public Const cSN_ITSurvey_AllTeams As String = "IT Survey - All Teams"

    Public Const cSN_Lookups As String = "Lookups"
    Public Const cSN_Teams As String = "Teams"

    Public Const cSN_PartnerSurveyQuestions As String = "Partner Survey Questions"
    Public Const cSN_PartnerSurveyData As String = "Partner Survey Data"

    Public Const cSN_BusinessSurveyQuestions As String = "Business Survey Questions"
    Public Const cSN_BusinessSurveyData As String = "Business Survey Data"

    Public Const cSN_ITSurveyQuestions As String = "IT Survey Questions"
    Public Const cSN_ITSurveyData As String = "IT Survey Data"

    Public Const cSN_HelpDeskSurveyQuestions As String = "Help desk Survey Questions"

    Public Const cSN_BudgetVarianceData As String = "Budget Variance Data"
    Public Const cSN_ITRProcessingData As String = "ITR Processing Data"
    Public Const cSN_OnTimeDeliveryData As String = "On-Time Delivery Data"

#End Region

#Region "Scorecard - Individual Team worksheet (cSCIT_) constants "

    Public Const cSCIT_TeamNameCell As String = "$B$2"
    Public Const cSCIT_SurveyPeriodCell As String = "$E$2"

    Public Const cSCIT_OpenedITRsCell As String = "$A$42"
    Public Const cSCIT_ClosedITRsCell As String = "$B$42"
    Public Const cSCIT_ActiveITRsCell As String = "$C$42"

    Public Const cSCIT_TeamScoreCells As String = "$F$3:$F$33"

#End Region

#Region "Scorecard - All Teams worksheet (cSCAT_) constants"

    Public Const cSCAT_ResultsCells As String = "$F$3:$T$36"
    Public Const cSCAT_Scorecard_Cells As String = "$A$2:$S$36"

#End Region

#Region "Survey Mapping worksheet (cSM_) constants"

    Public Const cSM_ITSurveyCells As String = "$B$4:$B$23"
    Public Const cSM_ITSurveyResponsesCell As String = "$D$2"

    Public Const cSM_BusinessSurveyCells As String = "$J$4:$J$17"
    Public Const cSM_BusinessSurveyResponsesCell As String = "$L$2"

    Public Const cSM_PartnerSurveyCells As String = "$R$4:$R$11"
    Public Const cSM_PartnerSurveyResponsesCell As String = "$U$2"

    Public Const cSM_ResponseCountOffset As Integer = 22

#End Region

#Region "Survey Results - All Teams worksheet(s) (cSRAT_) constants"

    Public Const cSRAT_SurveyResultsCells As String = "$F$3:$T$24"
    Public Const cSRAT_Scorecard_Cells As String = "$A$2:$S$25"

#End Region

#Region "Lookups worksheet (cLU_) constants "

    ' Used on Lookups worksheet and Teams worksheet.

    Public Const cLU_TeamsInfoCell As String = "$A$5"
    Public Const cLU_ManagerInfoCell As String = "$A$29"

#End Region

#Region "Survey Data worksheet(s) constants (cSD_)"

    ' Survey Data constants start with "cSD_"

    Public Const cSD_SurveyNameCell As String = "$C$4"
    Public Const cSD_SurveyPeriodCell As String = "$C$5"
    Public Const cSD_QuestionsSheetCell As String = "$C$6"
    Public Const cSD_QuestionsLocationCell As String = "$C$7"
    Public Const cSD_ColumnWidthLocationCell As String = "$C$8"
    Public Const cSD_StatisticsLocationCell As String = "$C$9"
    Public Const cSD_PrimaryQuestionsLocationCell As String = "$C$10"
    Public Const cSD_FollowUpQuestionsLocationCell As String = "$C$11"
    Public Const cSD_TeamNameCell As String = "$C$12"
    Public Const cSD_TeamNameCell_RC As String = "R12C3"    ' Keep in sync with cSD_TeamNameCell

    Public Const cSD_StartDataRowCell As String = "$C$13"
    Public Const cSD_EndDataRowCell As String = "$C$14"
    Public Const cSD_StartDataColumnCell As String = "$C$15"
    Public Const cSD_EndDataColumnCell As String = "$C$16"

    Public Const cSD_ChartCountCell As String = "$C$17"
    Public Const cSD_QuestionCountCell As String = "$C$18"

    Public Const cSD_ResponseCountCell As String = "$C$23"

    Public Const cSD_AverageDeviationCell As String = "$C$25"
    Public Const cSD_OverallAverageCell As String = "$C$26"

    ' Row and Column addresses

    Public Const cSD_RawDataRow As Integer = 31
    Public Const cSD_RawDataColumn As Integer = 1

    ' This is where the Survey Data lives once it is in final processing position

    Public Const cSD_StartDataRow As Integer = 31
    Public Const cSD_StartDataColumn As Integer = 6

    Public Const cSD_ResponseLabelRowStart As Integer = 16
    Public Const cSD_ResponseLabelRowEnd As Integer = 21    ' Used to be 22 when had Not Answered
    Public Const cSD_ResponseLabelColumn As Integer = 5
    Public Const cSD_ResponseValueRowStart As Integer = cSD_ResponseLabelRowStart
    Public Const cSD_ResponseValueRowEnd As Integer = cSD_ResponseLabelRowEnd
    Public Const cSD_ResponseValueColumn As Integer = 6

    Public Const cSD_ResponseCountRow As Integer = 23
    Public Const cSD_AverageDeviationRow As Integer = 25
    Public Const cSD_OverallAverageRow As Integer = 26
    Public Const cSD_QuestionTextRow As Integer = 28
    Public Const cSD_QuestionIDRow As Integer = 30
    Public Const cSD_HeaderIDRow As Integer = 30

    ' Column numbers

    Public Const cSD_ID_Column As Integer = 1
    Public Const cSD_SurveyPeriod_Column As Integer = 2
    Public Const cSD_RespondentName_Column As Integer = 3
    Public Const cSD_TeamName_Column As Integer = 4
    Public Const cSD_TeamNumber_Column As Integer = 5

    ' Offsets from cQuestionIDRow

    Public Const cSD_ColumnWidthOffset As Integer = -29
    Public Const cSD_IsStatisticsColumnOffset As Integer = -28
    Public Const cSD_ISPrimarQuestionColumnOffset As Integer = -27
    Public Const cSD_IsFollowUpQuestionColumnOffset As Integer = -26

    ' Columns

    Public Const cSD_ColumnWidth_Row As Integer = 1
    Public Const cSD_IsStatisticsColumn_Row As Integer = 2
    Public Const cSD_IsPrimaryQuestionColumn_Row As Integer = 3
    Public Const cSD_IsFollowUpQuestionColumns_Row As Integer = 4

    ' Dimensions and starting location of Charts in Excel
    ' Need to keep height/width ration constanst across the applications.

    Public Shared cSD_SurveyChartStartingOffset As Integer = 590
    Public Shared cSD_SurveyChartSpacing As Integer = 611
    Public Shared cSD_SurveyChartHeight As Integer = 270
    Public Shared cSD_SurveyChartWidth As Integer = 540

#End Region

#Region "Survey Questions worksheet(s)(cSQ_) constants"

    ' Survey Questions constants start with "cSQ_"

    ' These locations indicate the location of the cell on the "Lookups" sheet
    ' that point to the

    Public Const cSQ_QuestionsLocationCell As String = "$A$4"
    Public Const cSQ_ColumnWidthsLocationCell As String = "$A$5"
    Public Const cSQ_StatisticsLocationCell As String = "$A$6"
    Public Const cSQ_PrimaryQuestionsLocationCell As String = "$A$7"
    Public Const cSQ_FollowUpQuestionsLocationCell As String = "$A$8"

#End Region

#Region "Chart Formatting (cCH_) constants"
    ' Chart Formating constants start with "cCH_"

    Public Shared cCH_TickLabelFontSize As Integer = 12
    Public Shared cCH_DataLabelFontSize As Integer = 12
    Public Shared cCH_SurveyValueAxisLabel As String = "# of Responses"
    Public Shared cCH_OnTimeDeliveryValueAxisLabel As String = "% of Estimated Schedule Time used"

#End Region

#Region "PowerPoint (cPP_) constants"

    ' PowerPoint constants start with "cPP_"

    ' Dimensions and starting location of Charts in PowerPoint.
    ' Need to keep height/width ration constanst across the applications.

    Public Shared cPP_SurveyChartTop As Integer = 150
    Public Shared cPP_SurveyChartLeft As Integer = 50
    Public Shared cPP_SurveyChartHeight As Integer = 310
    Public Shared cPP_SurveyChartWidth As Integer = 620

    Public Shared cPP_OnTimeChartTop As Integer = 150
    Public Shared cPP_OnTimeChartLeft As Integer = 50
    Public Shared cPP_OnTimeChartHeight As Integer = 310
    Public Shared cPP_OnTimeChartWidth As Integer = 620

    ' Starting location of survey results in PowerPoint

    Public Shared cPP_SurveyFeedbackTop As Integer = 150
    Public Shared cPP_SurveyFeedbackLeft As Integer = 50
    Public Shared cPP_SurveyFeedbackHeight As Integer = 400
    Public Shared cPP_SurveyFeedbackWidth As Integer = 620

    Public Shared cPP_TitleFontSize As Integer = 18
    Public Shared cPP_SubQuestionFontSize As Integer = 16    ' This is not correct name.
    Public Shared cPP_ResponseFontSize As Integer = 10
    Public Shared cPP_ResponseLeftMargin As Integer = 144

    Public Shared cPP_MaxResponseLengthPerPage As Integer = 2000

#End Region

    Public Const cCOLUMN_WIDTH As Single = 5.0#
    Public Const cRowHeight As Single = 12.75

    Public Const cIDColumnWidth As Integer = 5
    Public Const cReleaseNameColumnWidth As Integer = 30
    Public Const cDescriptionColumnWidth As Integer = 50
    Public Const cDateColumnWidth As Integer = 12
    Public Const cPercentColumnWidth As Integer = 10

    Public Const cMaxSheetNameLen As Integer = 40

    Public Const cHeaderID_RowShort As Integer = 10  ' For sheets with no charts
    Public Const cHeaderID_Column As Integer = 1

    ' This is where the raw data is read into.  We then shove some columns over
    ' to get things where we need them to keep all the survey sheets consistent.

    ' Control the formatting of each sheet.

    Public Const cSurveyResultsRawColumn1Width As Integer = 5
    Public Const cSurveyResultsRawColumn2Width As Integer = 25
    Public Const cSurveyResultsRawColumn3Width As Integer = 25
    Public Const cSurveyResultsRawColumn4Width As Integer = 25
    Public Const cSurveyResultsRawColumn5Width As Integer = 25
    Public Const cSurveyResultsRawQuestionColumnWidth As Integer = 5
    Public Const cSurveyResultsRawQuestionResponseColumnWidth As Integer = 110
    ' The row numbers for each type of information.

    Public Const cChart_HeaderRow As Integer = 1
    Public Const cChart_StatisticsRow As Integer = 2
    Public Const cChart_DataInfoRow As Integer = 3

    Public Const cChart_ProjectVersionColumn As Integer = 8
    Public Const cChart_DataVersionColumn As Integer = 9
    Public Const cChart_ChartVersionColumn As Integer = 10

    Public Const cChart_NbrVisibleDataRows As Integer = 3    ' Like at least 2.

    Public Const cHeaderFontSize As Integer = 12
    Public Const cHeaderFontSizeMedium As Integer = 10
    Public Const cHeaderFontSizeSmall As Integer = 8

    Public m_vntPriorCalculationState As Object
    Public m_vntPriorScreenUpdatingState As Object

    Public Const cPercentageFormatString As String = "[Green]0%;[Red]-0%;[Black]General;[Black]General"
    Public Const cNumericFormatString As String = "[Green]0.00;[Red]-0.00;[Black]General;[Black]General"

#Region "Enumerations"

    Public Enum WrapText As Byte
        Yes = 1
        No = 0
    End Enum

    Public Enum MakeBold As Byte
        Yes = 1
        No = 0
    End Enum

    Public Enum UnderLine As Byte
        Yes = 1
        No = 0
    End Enum

#End Region

End Class
