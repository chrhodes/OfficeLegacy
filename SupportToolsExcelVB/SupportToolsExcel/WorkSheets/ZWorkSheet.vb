Option Explicit On

Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Core
Imports PacificLife.Life

Public Class ZWorkSheet

    ' This routine has been factored out of CreateSurveyDataSheet() so data ranges
    ' can be changed and then the forumlas that depend on the values updated.  This
    ' was doen with INDIRECT functions before, however, there is a limit (bug) of 256 characters
    ' per programmatically assigned FormulaArray.  This requires us to hand craft the formulas 
    ' that are used.  See statisticsArrayFomula and countQuestionsArrayFormula

    Public Shared Sub AddFormulas()
        'PLLog.Trace1("Enter", Common.PROJECT_NAME)

        'Dim ws As Excel.Worksheet
        'Dim startRow As Integer = Common.cHeaderID_RowShort
        'Dim startCol As Integer = Common.cHeaderID_Column
        'Dim headerFontSize As Integer = Common.cHeaderFontSize
        'Dim headerBold As Boolean = True
        'Dim headerUnderline As Boolean = True
        'Dim headerWrapText As Boolean = True
        'Dim headerHorizontalAlignment As Excel.XlHAlign = Excel.XlHAlign.xlHAlignCenter
        ''Dim row As Integer
        ''Dim column As Integer
        'Dim questionsRange As Excel.Range
        'Dim startDataRow As Long
        'Dim startDataColumn As Long
        'Dim endDataRow As Long
        'Dim endDataColumn As Long
        ''Dim val As Integer
        'Dim questionsRow As Long = Common.cSD_QuestionIDRow
        'Dim currentProtectionMode As Boolean

        'ws = Globals.ThisAddIn.Application.ActiveSheet
        'Common.ExcelUtil.ProtectSheet(ws, False)
        'Common.ExcelUtil.ScreenUpdatesOff()

        'With ws
        '    ' Get the current values off the sheet.

        '    startDataRow = ws.Range(Common.cSD_StartDataRowCell).Value
        '    startDataColumn = ws.Range(Common.cSD_StartDataColumnCell).Value
        '    endDataRow = ws.Range(Common.cSD_EndDataRowCell).Value
        '    endDataColumn = ws.Range(Common.cSD_EndDataColumnCell).Value

        '    ' Add the formulas that calculate the average

        '    ' TODO: Grrr.  Excel sucks.  Cannot enter array formula longer than 255 characters from code???
        '    ' So, make the formula shorter by ripping out all the INDIRECT stuff.  This was only
        '    ' really needed during debugging.  Can always put this code in a method we call if
        '    ' the referenced cells are updated.  Oh, yeah, that is this routine now :)

        '    ' This generates data for All responses always.

        '    Dim statisticsArrayFormulaAllIT As String

        '    statisticsArrayFormulaAllIT = _
        '        "(" & _
        '        "   IF(R" & startDataRow & "C:R" & endDataRow & "C < 6, " & _
        '        "      R" & startDataRow & "C:R" & endDataRow & "C, " & _
        '        "      """"" & _
        '        "    )" & _
        '        ")"

        '    ' This generates data filtered by what Team is selected (including "All IT")

        '    Dim statisticsArrayFormula As String

        '    ' TODO: Use StringBuilder.  Also, this does not handle single row Not Applicable responses well.

        '    statisticsArrayFormula = _
        '        "(" & _
        '        "  IF(" & Common.cSD_TeamNameCell_RC & " = """ & Common.cAllITString & """," & _
        '        "    IF(R" & startDataRow & "C:R" & endDataRow & "C < 6, " & _
        '        "      R" & startDataRow & "C:R" & endDataRow & "C, " & _
        '        "      """"" & _
        '        "    )," & _
        '        "    IF(" & Common.cSD_TeamNameCell_RC & " = R" & startDataRow & "C" & Common.cSD_TeamName_Column & ":R" & endDataRow & "C" & Common.cSD_TeamName_Column & "," & _
        '        "      IF(R" & startDataRow & "C:R" & endDataRow & "C < 6, " & _
        '        "        R" & startDataRow & "C:R" & endDataRow & "C, " & _
        '        "        """"" & _
        '        "      )" & _
        '        "    )" & _
        '        "  )" & _
        '        ")"

        '    Dim countQuestionsArrayFormula As String

        '    countQuestionsArrayFormula = _
        '    "=SUM(" & _
        '    "  IF(" & Common.cSD_TeamNameCell_RC & " = """ & Common.cAllITString & """," & _
        '    "    IF(R" & startDataRow & "C:R" & endDataRow & "C = RC" & Common.cSD_TeamName_Column & ", " & _
        '    "      1," & _
        '    "      0)," & _
        '    "    IF(R" & startDataRow & "C" & Common.cSD_TeamName_Column & ":R" & endDataRow & "C" & Common.cSD_TeamName_Column & " = " & Common.cSD_TeamNameCell_RC & "," & _
        '    "      IF(R" & startDataRow & "C:R" & endDataRow & "C = RC" & Common.cSD_TeamName_Column & ", " & _
        '    "        1," & _
        '    "        0)," & _
        '    "    )" & _
        '    "  )" & _
        '    ")"
        '    '
        '    '=SUM(IF(R12C3="All",IF(R31C:R37C=R16C4,1,0),IF(R31C:R37C=R9C3,IF(R31C:R37C=R16C4,1,0),)))

        '    Dim countQuestionsArrayFormulaIndirect As String

        '    Dim teamNameLookupFormula As String
        '    teamNameLookupFormula = "=LOOKUP(RC[1],INDIRECT(Team_Name_Lookup))"

        '    ' TODO: This cannot be used unless the numbers that need to be relative are adjusted, ex. 16.

        '    countQuestionsArrayFormulaIndirect = _
        '    "=SUM(IF($C$12=""" & Common.cAllITString & """,IF(INDIRECT(""R"" & $C$13 & ""C:R"" & $C$14 & ""C"",FALSE)=$D16,1,0),IF(INDIRECT(""R"" & $C$13 & ""C:R"" & $C$14 & ""C4"",FALSE)=$C$12,IF(INDIRECT(""R"" & $C$13 & ""C:R"" & $C$14 & ""C"",FALSE) = $D16,1,0),0)))"

        '    ' This is what we want to put in.  If Cntrl-Shift-Enter from application this works.
        '    '.Range(.Cells(Common.cQuestionIDRow - 6, startDataColumn), .Cells(Common.cQuestionIDRow - 6, endDataColumn)).Formula = _
        '    '    "=AVERAGE(IF($C$12=""All"",IF(INDIRECT(""R""&$C$13&""C:R""&$C$14&""C"",FALSE)<6,INDIRECT(""R""&$C$13&""C:R""&$C$14&""C"",FALSE),""""),IF(INDIRECT(""R""&$C$13&""C4:R""&$C$14&""C4"",FALSE)=$C$12,(IF(INDIRECT(""R""&$C$13&""C:R""&$C$14&""C"",FALSE)<6,INDIRECT(""R""&$C$13&""C:R""&$C$14&""C"",FALSE),"""")))))"

        '    ' Add the formulas that calculate the stdevp

        '    '.Range(.Cells(startDataRow - 5, startDataColumn), .Cells(startDataRow - 5, endDataColumn)).FormulaArray = _
        '    '    "=AVERAGE(IF($C$12=""All"",IF(INDIRECT(""R""&$C$13&""C:R""&$C$14&""C"",FALSE)<6,INDIRECT(""R""&$C$13&""C:R""&$C$14&""C"",FALSE),""""),IF(INDIRECT(""R""&$C$13&""C4:R""&$C$14&""C4"",FALSE)=$C$12,(IF(INDIRECT(""R""&$C$13&""C:R""&$C$14&""C"",FALSE)<6,INDIRECT(""R""&$C$13&""C:R""&$C$14&""C"",FALSE),"""")))))"

        '    ' TODO: May want to move this back to CreateSurveyDataWorkSheet as it is all indirect.

        '    ' Add the team name lookup formulas

        '    .Range(.Cells(startDataRow, Common.cSD_TeamName_Column), .Cells(endDataRow, Common.cSD_TeamName_Column)).FormulaR1C1 = teamNameLookupFormula

        '    ' Walk the questions adding the statistics Array Formulas when appropriate (Statistics Row = "X")

        '    questionsRange = .Range(.Cells(Common.cSD_QuestionIDRow, startDataColumn), .Cells(Common.cSD_QuestionIDRow, endDataColumn))

        '    'Dim val As Integer

        '    Globals.ThisAddIn.Application.CalculateFull()

        '    ' HACK: The values in the indirect references from the quesitons sheets are not reliably
        '    ' reflected.  Sometimes the rng.Offset cells have 0 sometimes the values we expect.  Need
        '    ' to investigate what to do here.

        '    Dim averageAggregationFormula As String = ""

        '    Dim questionCountFormula As String = "=SUM(" & _
        '        "R" & Common.cSD_IsPrimaryQuestionColumn_Row & "C" & startDataColumn & ":" & _
        '        "R" & Common.cSD_IsPrimaryQuestionColumn_Row & "C" & endDataColumn & ")"

        '    Dim chartCountFormula As String = "=SUM(" & _
        '        "R" & Common.cSD_IsStatisticsColumn_Row & "C" & startDataColumn & ":" & _
        '        "R" & Common.cSD_IsStatisticsColumn_Row & "C" & endDataColumn & ")"

        '    ' To make tweaking the chart size easier show how many questions and charts have
        '    ' been indicated on the Question format sheet, but give the user a cell they can
        '    ' override.

        '    .Range(Common.cSD_ChartCountCell).FormulaR1C1 = "=RC[-2]"
        '    .Range(Common.cSD_ChartCountCell).Offset(0, -2).FormulaR1C1 = chartCountFormula
        '    .Range(Common.cSD_QuestionCountCell).FormulaR1C1 = "=RC[-2]"
        '    .Range(Common.cSD_QuestionCountCell).Offset(0, -2).FormulaR1C1 = questionCountFormula

        '    For Each rng As Excel.Range In questionsRange
        '        'MessageBox.Show(rng.Value & "R" & rng.Row & "C" & rng.Column & ":>" & rng.Offset(-29, 0).Value & ":" & rng.Offset(-28, 0).Value & ":" & rng.Offset(-27, 0).Value & "<")

        '        'val = rng.Offset(-28, 0).Value

        '        If rng.Offset(Common.cSD_IsStatisticsColumnOffset, 0).Value > 0 Then

        '            ' Add Statistics Array Formulas

        '            With rng.Offset(-6, 0)
        '                .FormulaArray = "=IFERROR(AVERAGE" & statisticsArrayFormulaAllIT & ", ""No Data"")"
        '                .NumberFormat = "0.0"
        '            End With

        '            With rng.Offset(-5, 0)
        '                .FormulaArray = "=IFERROR(STDEVP" & statisticsArrayFormula & ", ""No Data"")"
        '                .NumberFormat = "0.0"
        '            End With

        '            With rng.Offset(-4, 0)
        '                .FormulaArray = "=IFERROR(AVERAGE" & statisticsArrayFormula & ", ""No Data"")"
        '                .NumberFormat = "0.0"
        '            End With

        '            ' Add the formulas that perform the Response range hack
        '            ' This is no longer needed as the questions have been updated.

        '            'With rng.Offset(-3, 0)
        '            '    .Formula = "=6-F26"
        '            '    .NumberFormat = "0.0"
        '            'End With

        '            ' Add formulas that count the number of each type of response.

        '            rng.Offset(-14, 0).FormulaArray = countQuestionsArrayFormula
        '            rng.Offset(-13, 0).FormulaArray = countQuestionsArrayFormula
        '            rng.Offset(-12, 0).FormulaArray = countQuestionsArrayFormula
        '            rng.Offset(-11, 0).FormulaArray = countQuestionsArrayFormula
        '            rng.Offset(-10, 0).FormulaArray = countQuestionsArrayFormula
        '            rng.Offset(-9, 0).FormulaArray = countQuestionsArrayFormula
        '            rng.Offset(-8, 0).FormulaArray = countQuestionsArrayFormula

        '            ' And sums the results

        '            rng.Offset(-7, 0).Formula = "=SUM(F16:F22)"

        '            ' TODO: Build the formula that aggregates the results.

        '            If averageAggregationFormula = "" Then
        '                averageAggregationFormula = "RC" & rng.Offset(-7, 0).Column
        '            Else
        '                averageAggregationFormula = averageAggregationFormula & ", RC" & rng.Offset(-7, 0).Column
        '            End If

        '        End If

        '        If rng.Offset(Common.cSD_ISPrimarQuestionColumnOffset, 0).Value > 0 Then
        '            ' Add question lookup formula

        '            rng.Offset(-2, 0).FormulaR1C1 = _
        '                "=LOOKUP(R30C, INDIRECT(""'"" & R6C3 & ""'!"" & INDIRECT(""'"" & R6C3 & ""'!"" & R7C3)))"

        '            ' Merge the Primary Question Cell with the next adjacent cell to improve formatting.
        '            ' Only do this if the column width is < 10 (This is a HACK)

        '            Debug.Print(rng.Offset(Common.cSD_ColumnWidthOffset, 0).Value & " " & rng.Column())

        '            If rng.Offset(Common.cSD_ColumnWidthOffset, 0).Value < 10 Then
        '                .Range(ws.Cells(Common.cSD_QuestionTextRow, rng.Column()), ws.Cells(Common.cSD_QuestionTextRow, rng.Column() + 1)).Merge()
        '            End If
        '        End If

        '        If rng.Offset(Common.cSD_IsFollowUpQuestionColumnOffset, 0).Value > 0 Then
        '            ' Add follow-up question lookup formula

        '            rng.Offset(-1, 0).FormulaR1C1 = _
        '                "=LOOKUP(R30C, INDIRECT(""'"" & R6C3 & ""'!"" & INDIRECT(""'"" & R6C3 & ""'!"" & R7C3)))"
        '        End If
        '    Next

        '    ' The formula used below is built in the For Each rng loop above.
        '    ' The AVERAGE formula is wrapped to handle the situation where there
        '    ' are no results for the selected team which would return a DIV error.

        '    .Range("$B$24").Value = "Overall Average (All IT)"

        '    With .Range("$C$24")
        '        .FormulaR1C1 = "=IFERROR(AVERAGE(" & averageAggregationFormula & "), ""No Data"")"
        '        .NumberFormat = "0.0"
        '    End With

        '    .Range("$B$25").Value = "Average Deviation (Team)"

        '    With .Range("$C$25")
        '        .FormulaR1C1 = "=IFERROR(AVERAGE(" & averageAggregationFormula & "), ""No Data"")"
        '        .NumberFormat = "0.0"
        '    End With

        '    .Range("$B$26").Value = "Overall Average (Team)"

        '    With .Range("$C$26")
        '        .FormulaR1C1 = "=IFERROR(AVERAGE(" & averageAggregationFormula & "), ""No Data"")"
        '        .NumberFormat = "0.0"
        '    End With

        '    .PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape
        'End With

        'Common.ExcelUtil.ProtectSheet(ws, currentProtectionMode)
        'Common.ExcelUtil.ScreenUpdatesOn()

        'PLLog.Trace1("Exit", Common.PROJECT_NAME)
    End Sub

    Public Shared Function CreateWorksheet(ByVal sheetName As String) As Excel.Worksheet
        'PLLog.Trace1("Enter", Common.PROJECT_NAME)

        'Dim ws As Excel.Worksheet
        'Dim startRow As Integer = Common.cHeaderID_RowShort
        'Dim startCol As Integer = Common.cHeaderID_Column
        'Dim headerFontSize As Integer = Common.cHeaderFontSize
        'Dim headerBold As Boolean = True
        'Dim headerUnderline As Boolean = True
        'Dim headerWrapText As Boolean = True
        'Dim headerHorizontalAlignment As Excel.XlHAlign = Excel.XlHAlign.xlHAlignCenter
        ''Dim row As Integer
        ''Dim column As Integer
        ''Dim questionsRange As Excel.Range
        'Dim startDataRow As Long = 31
        'Dim startDataColumn As Long = 6
        'Dim endDataRow As Long
        'Dim endDataColumn As Long
        ''Dim val As Integer
        'Dim questionsRow As Long = Common.cSD_QuestionIDRow
        'Dim currentProtectionMode As Boolean
        'Dim surveyName As String = ""

        '' Sheet should already have data on it from LoadMDBFile()

        'Globals.ThisAddIn.Application.Sheets(sheetName).Activate()
        'ws = Globals.ThisAddIn.Application.ActiveSheet

        'currentProtectionMode = Common.ExcelUtil.ProtectSheet(ws, False)
        'Common.ExcelUtil.ScreenUpdatesOff()

        '' Shove the data into the location we need based on the type of survey.
        '' Then add any information that was missing.

        'Select Case sheetName
        '    Case Common.cSN_PartnerSurveyData
        '        surveyName = "Partner Survey"
        '        ws.Columns("D:D").Select()

        '        With Globals.ThisAddIn.Application.Selection
        '            .Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove)
        '            .Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove)
        '        End With

        '        With ws.Range("$D$30")
        '            .Value = "Team Name"
        '            .AddComment("Need to add this to the survey")
        '        End With

        '        With ws.Range("$E$30")
        '            .Value = "<blank>"
        '        End With

        '    Case Common.cSN_BusinessSurveyData
        '        surveyName = "Business Survey"
        '        ws.Columns("D:D").Select()

        '        With Globals.ThisAddIn.Application.Selection
        '            .Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove)
        '        End With

        '        With ws.Range("$D$30")
        '            .Value = "Team Name"
        '        End With

        '    Case Common.cSN_ITSurveyData
        '        surveyName = "IT Survey"
        '        ws.Columns("C:C").Select()

        '        With Globals.ThisAddIn.Application.Selection
        '            .Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove)
        '            .Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove)
        '        End With

        '        With ws.Range("$C$30")
        '            .Value = "<blank>"
        '        End With

        '        With ws.Range("$D$30")
        '            .Value = "Team Name"
        '        End With

        '    Case Else
        '        MessageBox.Show("Unsupported Survey Data: " & sheetName)

        'End Select

        'ws.Range("A1").Activate()

        '' This information is present on all the Survey Data Sheets

        'With ws
        '    .Range("$B$4").Value = "Survey Name"
        '    .Range(Common.cSD_SurveyNameCell).Value = surveyName

        '    .Range("$B$5").Value = "Survey Period"
        '    .Range(Common.cSD_SurveyPeriodCell).Formula = "=Survey_Period"

        '    .Range("$B$6").Value = "Questions Sheet"
        '    .Range(Common.cSD_QuestionsSheetCell).Formula = "=$C$4 & "" Questions"""

        '    .Range("$B$7").Value = "Questions Location"
        '    .Range(Common.cSD_QuestionsLocationCell).Formula = "=Questions_Location"

        '    .Range("$B$8").Value = "Row Width Location"
        '    .Range(Common.cSD_ColumnWidthLocationCell).Value = "=Row_Width_Location"

        '    .Range("$B$9").Value = "Statistics Location"
        '    .Range(Common.cSD_StatisticsLocationCell).Value = "=Statistics_Location"

        '    .Range("$B$10").Value = "Primary Questions Location"
        '    .Range(Common.cSD_PrimaryQuestionsLocationCell).Value = "=Primary_Questions_Location"

        '    .Range("$B$11").Value = "FollowUp Questions Location"
        '    .Range(Common.cSD_FollowUpQuestionsLocationCell).Value = "=FollowUp_Questions_Location"

        '    .Range("$B$12").Value = "Team Name"
        '    .Range(Common.cSD_TeamNameCell).Value = "=Team_Name"

        '    .Range("$B$13").Value = "Start Data Row"
        '    .Range(Common.cSD_StartDataRowCell).Value = startDataRow

        '    endDataRow = Common.ExcelUtil.FindLastRow(.Range("$A$30"))

        '    .Range("$B$14").Value = "End Data Row"
        '    .Range(Common.cSD_EndDataRowCell).Value = endDataRow
        '    ' End Data Row filled in when data loaded

        '    .Range("$B$15").Value = "Start Data Column"
        '    .Range(Common.cSD_StartDataColumnCell).Value = startDataColumn

        '    endDataColumn = Common.ExcelUtil.FindLastColumn(.Range("$A$30"))

        '    .Range("$B$16").Value = "End Data Column"
        '    .Range(Common.cSD_EndDataColumnCell).Value = endDataColumn
        '    ' End Data Column filled in when data loaded

        '    With .Range("$B$17")
        '        .Value = "Chart Count"
        '        ' Chart count entered by user
        '        .AddComment("Enter number of charts to produce.  Only chart questions with required answers")
        '    End With

        '    With .Range("$B$18")
        '        .Value = "Question Count"
        '        .AddComment("Enter the number of questions")
        '    End With

        '    .Range("$B$23").Value = "Response Count"
        '    ' TODO: This needs to account for how many rows of results we have.
        '    .Range("$C$23").FormulaArray = "=IF($C$12=""" & Common.cAllITString & """,$C$14-$C$13+1,SUM(IF(INDIRECT(""R""&$C$13&""C4:R""&$C$14&""C4"",FALSE) = $C$12,1,0)))"

        '    .Range("$E$1").Value = "Width"
        '    .Range("$E$2").Value = "Statistics"
        '    .Range("$E$3").Value = "Primary Question"
        '    .Range("$E$4").Value = "Secondary Question"

        '    With .Range("$D$16")
        '        .Value = "5"
        '        '.AddComment("Change this to 5 to 1 when fixed in questions")
        '    End With

        '    .Range("$D$17").Value = "4"
        '    .Range("$D$18").Value = "3"
        '    .Range("$D$19").Value = "2"
        '    .Range("$D$20").Value = "1"
        '    .Range("$D$21").Value = "6"
        '    .Range("$D$22").Value = "0"

        '    .Range("$E$16").Value = "Strongly Agree"
        '    .Range("$E$17").Value = "Agree"
        '    .Range("$E$18").Value = "Neutral"
        '    .Range("$E$19").Value = "Disagree"
        '    .Range("$E$20").Value = "Strongly Disagree"
        '    .Range("$E$21").Value = "Not Applicable"
        '    .Range("$E$22").Value = "Not Answered"

        '    .Range("$E$24").Value = "Average (All IT)"
        '    .Range("$E$25").Value = "StdDevP"
        '    .Range("$E$26").Value = "Average (Team)"

        '    With .Range("$E$27")
        '        .Value = "Average Hack"
        '        .AddComment("Remove this calculation when fixed in questions")
        '    End With

        '    .Range("$E$28").Value = "Question"
        '    .Range("$E$29").Value = "Follow-up Question"

        '    ' Indicate the parts of the worksheet that need data entered by the user.  Protect the rest.

        '    ' TODO: Finish this section.

        '    'With .Range(.Cells(startRow + 1, startCol), .Cells(startRow + Common.cNumberTeams, startCol + 6))
        '    '    .Locked = False

        '    '    With .Interior
        '    '        .Pattern = Excel.Constants.xlSolid
        '    '        .PatternColorIndex = Excel.Constants.xlAutomatic
        '    '        .ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3
        '    '        .TintAndShade = 0.599993896298105
        '    '        .PatternTintAndShade = 0
        '    '    End With
        '    'End With

        '    ' Add the formulas that show the width of this column

        '    .Range(.Cells(1, startDataColumn), .Cells(1, endDataColumn)).Formula = _
        '        "=LOOKUP(F$30,INDIRECT(""'"" & $C$6 & ""'!"" & INDIRECT(""'"" & $C$6 & ""'!"" & $C$8)))"

        '    ' Add the formulas that show if statistics are being generated for this column

        '    .Range(.Cells(2, startDataColumn), .Cells(2, endDataColumn)).Formula = _
        '        "=LOOKUP(F$30,INDIRECT(""'"" & $C$6 & ""'!"" & INDIRECT(""'"" & $C$6 & ""'!"" & $C$9)))"

        '    ' Add the formulas that show if primary question in this column

        '    .Range(.Cells(3, startDataColumn), .Cells(3, endDataColumn)).Formula = _
        '        "=LOOKUP(F$30,INDIRECT(""'"" & $C$6 & ""'!"" & INDIRECT(""'"" & $C$6 & ""'!"" & $C$10)))"

        '    ' Add the formulas that show if follow-up question in this column

        '    .Range(.Cells(4, startDataColumn), .Cells(4, endDataColumn)).Formula = _
        '        "=LOOKUP(F$30,INDIRECT(""'"" & $C$6 & ""'!"" & INDIRECT(""'"" & $C$6 & ""'!"" & $C$11)))"

        'End With

        'Common.ExcelUtil.ProtectSheet(ws, True)
        'Common.ExcelUtil.ScreenUpdatesOn()

        'PLLog.Trace1("Exit", Common.PROJECT_NAME)

        '' TODO: Determine if we really need to return the sheet.  We do not create it anymore.
        'Return ws
        Return Nothing
    End Function

    Public Shared Sub FormatWorksheet()
        'PLLog.Trace1("Enter", Common.PROJECT_NAME)

        'Dim questionCount As Integer
        'Dim chartCount As Integer
        'Dim startDataColumn As Integer
        'Dim endDataColumn As Integer
        'Dim i As Long
        'Dim currentProtectionMode As Boolean
        'Dim sht As Excel.Worksheet

        'sht = Globals.ThisAddIn.Application.ActiveSheet
        'currentProtectionMode = Common.ExcelUtil.ProtectSheet(sht, False)
        'Common.ExcelUtil.ScreenUpdatesOff()

        'With sht
        '    ' TODO: Some column name constants please.

        '    .Columns(1).ColumnWidth = Common.cSurveyResultsRawColumn1Width
        '    .Columns(2).ColumnWidth = Common.cSurveyResultsRawColumn2Width
        '    .Columns(3).ColumnWidth = Common.cSurveyResultsRawColumn3Width
        '    .Columns(4).ColumnWidth = Common.cSurveyResultsRawColumn4Width
        '    .Columns(5).ColumnWidth = Common.cSurveyResultsRawColumn5Width

        '    questionCount = .Range(Common.cSD_QuestionCountCell).Value
        '    chartCount = .Range(Common.cSD_ChartCountCell).Value
        '    startDataColumn = .Range(Common.cSD_StartDataColumnCell).Value
        '    endDataColumn = .Range(Common.cSD_EndDataColumnCell).Value

        '    For i = startDataColumn To endDataColumn
        '        With .Columns(i)
        '            Debug.Print(sht.Cells(1, i).value)
        '            .ColumnWidth = sht.Cells(1, i).value
        '            .WrapText = True
        '        End With
        '    Next i

        '    ' We just made some things wider, so adjust the height of the cells
        '    ' to reflect the new room.

        '    .Cells.Select()
        '    .Application.Selection.Rows.AutoFit()
        '    .Range("A1").Select()

        'End With

        'Common.ExcelUtil.ProtectSheet(sht, currentProtectionMode)
        'Common.ExcelUtil.ScreenUpdatesOn()

        'PLLog.Trace1("Exit", Common.PROJECT_NAME)

    End Sub

    Public Shared Sub LoadSurveyDataFromMDBFile()
        'PLLog.Trace1("Enter", Common.PROJECT_NAME)

        'Dim ws As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet

        'Dim inputFile As String = Common.ExcelUtil.GetFile(Common.cDEFAULT_SURVEY_FOLDER, "Select File Containing Survey Data", "Survey Data Files (*.MDB)|*.MDB")

        'If "" = inputFile Then
        '    Return
        'End If

        'Dim inputPath As String = System.IO.Path.GetDirectoryName(inputFile)

        'ws.Unprotect()

        'With ws.ListObjects.Add( _
        '    SourceType:=0, _
        '    Source:= _
        '        "ODBC;DSN=MS Access Database" _
        '        & ";DBQ=" & inputFile _
        '        & ";DefaultDir=" & inputPath _
        '        & ";DriverId=25;FIL=MS Access;MaxBufferSize=2048;PageTimeout=5;", _
        '    Destination:=ws.Range("$A$30")).QueryTable
        '    .CommandText = "SELECT * FROM Answers0 Answers0"
        '    .RowNumbers = False
        '    .FillAdjacentFormulas = False
        '    .PreserveFormatting = True
        '    .RefreshOnFileOpen = False
        '    .BackgroundQuery = True
        '    .RefreshStyle = Excel.XlCellInsertionMode.xlInsertDeleteCells
        '    .SavePassword = False
        '    .SaveData = True
        '    .AdjustColumnWidth = True
        '    .RefreshPeriod = 0
        '    .PreserveColumnInfo = True
        '    '.ListObject.DisplayName = "Table_Query_from_MS_Access_Database"
        '    .Refresh(BackgroundQuery:=False)
        'End With

        'ws.Protect()

        'PLLog.Trace1("Exit", Common.PROJECT_NAME)

    End Sub

End Class
