Imports PacificLife.Life

Public Class SurveyQuestionsWorkSheet

    Public Shared Function CreateWorkSheet(ByVal sheetName As String) As Excel.Worksheet
        PLLog.Trace1("Enter", "Scorecard")

        Dim ws As Excel.Worksheet
        Dim startRow As Integer = Globals.cHeaderID_RowShort
        Dim startCol As Integer = Globals.cHeaderID_Column
        Dim headerFontSize As Integer = Globals.cHeaderFontSize
        Dim headerFontSizeSmall As Integer = Globals.cHeaderFontSizeSmall
        Dim headerBold As Boolean = True
        Dim headerUnderline As Boolean = True
        Dim headerWrapText As Boolean = True
        Dim headerHorizontalAlignment As Excel.XlHAlign = Excel.XlHAlign.xlHAlignCenter
        Dim headerVerticalText As Integer = 90

        ws = Util.NewWorksheet(sheetName)

        With ws
            .Range(Globals.cSQ_QuestionsLocationCell).Value = "Questions"
            .Range(Globals.cSQ_ColumnWidthsLocationCell).Value = "Row Width"
            .Range(Globals.cSQ_StatisticsLocationCell).Value = "Statistics"
            .Range(Globals.cSQ_PrimaryQuestionsLocationCell).Value = "Primary Questions"
            .Range(Globals.cSQ_FollowUpQuestionsLocationCell).Value = "Follow-Up Questions"

            ' TODO: Add some columns with Section Headers along with comments so people know what to do.
            ' Also format some columns center aligned.

            Util.AddColumnToSheet(ws, startCol + 0, 15, Globals.WrapText.No, startRow, "Question #", _
                headerFontSize, headerBold, headerUnderline, headerWrapText, headerHorizontalAlignment)
            Util.AddColumnToSheet(ws, startCol + 1, 50, Globals.WrapText.Yes, startRow, "Question Text", _
                headerFontSize, headerBold, headerUnderline, headerWrapText, headerHorizontalAlignment)

            ' These columns control the associated Data sheet.

            Util.AddColumnToSheet(ws, startCol + 2, 5, Globals.WrapText.No, startRow, "Question #", _
                headerFontSizeSmall, headerBold, headerUnderline, headerWrapText, headerHorizontalAlignment, headerVerticalText)
            Util.AddColumnToSheet(ws, startCol + 3, 5, Globals.WrapText.No, startRow, "Width", _
                headerFontSizeSmall, headerBold, headerUnderline, headerWrapText, headerHorizontalAlignment, headerVerticalText)
            Util.AddCommentToCell(ws, startCol + 3, startRow, _
                "Enter width of question column on data sheet.  This helps make the charts fit on the data sheet. 5 is good for numeric answers and 110 for text answers")

            Util.AddColumnToSheet(ws, startCol + 4, 5, Globals.WrapText.No, startRow, "Question #", _
                headerFontSizeSmall, headerBold, headerUnderline, headerWrapText, headerHorizontalAlignment, headerVerticalText)
            Util.AddColumnToSheet(ws, startCol + 5, 5, Globals.WrapText.No, startRow, "Statistics", _
                headerFontSizeSmall, headerBold, headerUnderline, headerWrapText, headerHorizontalAlignment, headerVerticalText)
            Util.AddCommentToCell(ws, startCol + 5, startRow, _
                "Indicate if question generates statistical results. Use 1 for yes and 0 or blank for no.")

            Util.AddColumnToSheet(ws, startCol + 6, 5, Globals.WrapText.No, startRow, "Question #", _
                headerFontSizeSmall, headerBold, headerUnderline, headerWrapText, headerHorizontalAlignment, headerVerticalText)
            Util.AddColumnToSheet(ws, startCol + 7, 5, Globals.WrapText.No, startRow, "Primary Question", _
                headerFontSizeSmall, headerBold, headerUnderline, headerWrapText, headerHorizontalAlignment, headerVerticalText)
            Util.AddCommentToCell(ws, startCol + 7, startRow, _
                "Indicate if primary question. Use 1 for yes and 0 or blank for no.")

            Util.AddColumnToSheet(ws, startCol + 8, 5, Globals.WrapText.No, startRow, "Question #", _
                headerFontSizeSmall, headerBold, headerUnderline, headerWrapText, headerHorizontalAlignment, headerVerticalText)
            Util.AddColumnToSheet(ws, startCol + 9, 5, Globals.WrapText.No, startRow, "Follow-Up Question", _
                headerFontSizeSmall, headerBold, headerUnderline, headerWrapText, headerHorizontalAlignment, headerVerticalText)
            Util.AddCommentToCell(ws, startCol + 9, startRow, _
                "Indicate if follow-up question. Use 1 for yes and 0 or blank for no.  Should be automatically filled in and inverse of Primary Question")

            ' Indicate the parts of the worksheet that need data entered by the user.  Protect the rest.

            'With .Range(.Cells(startRow + 1, startCol), .Cells(startRow + Globals.cNumberTeams, startCol + 6))
            '    .Locked = False

            '    With .Interior
            '        .Pattern = Excel.Constants.xlSolid
            '        .PatternColorIndex = Excel.Constants.xlAutomatic
            '        .ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3
            '        .TintAndShade = 0.599993896298105
            '        .PatternTintAndShade = 0
            '    End With
            'End With

            .PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape

            .Protect(DrawingObjects:=False, Contents:=True, Scenarios:=False)
        End With

        PLLog.Trace1("Exit", "Scorecard")

        Return ws
    End Function

    Public Shared Sub LoadQuestions()
        Dim tempWs As Excel.Worksheet = Globals.ThisAddIn.Application.Sheets.Add()
        Dim questionsWs As Excel.Worksheet

        Util.ScreenUpdatesOff()
        Util.CalculationsOff()

        LoadQuestionsFromMDBFile()

        questionsWs = XmlUtil.ProcessQuestionsXML()

        ' No longer need the raw XML

        Util.DeleteSheet(tempWs)

        Util.CalculationsOn()
        Util.ScreenUpdatesOn()
    End Sub

    Private Shared Sub LoadQuestionsFromMDBFile()
        Dim ws As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet

        Dim inputFile As String = Util.GetFile(Globals.cDEFAULT_SURVEY_FOLDER, "Select File Containing Survey Questions", "Survey Questions Files (*.MDB)|*.MDB")
        Dim inputPath As String = System.IO.Path.GetDirectoryName(inputFile)

        With ws.ListObjects.Add( _
            SourceType:=0, _
            Source:= _
                "ODBC;DSN=MS Access Database" _
                & ";DBQ=" & inputFile _
                & ";DefaultDir=" & inputPath _
                & ";DriverId=25;FIL=MS Access;MaxBufferSize=2048;PageTimeout=5;", _
            Destination:=ws.Range("$A$1")).QueryTable
            .CommandText = "SELECT PdcQuestions.Sequence, PdcQuestions.XmlData FROM PdcQuestions PdcQuestions"
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .BackgroundQuery = True
            .RefreshStyle = Excel.XlCellInsertionMode.xlInsertDeleteCells
            .SavePassword = False
            .SaveData = True
            .AdjustColumnWidth = True
            .RefreshPeriod = 0
            .PreserveColumnInfo = True
            '.ListObject.DisplayName = "Table_Query_from_MS_Access_Database"
            .Refresh(BackgroundQuery:=False)
        End With
    End Sub
End Class
