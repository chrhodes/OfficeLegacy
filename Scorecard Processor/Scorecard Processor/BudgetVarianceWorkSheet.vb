Imports PacificLife.Life

Public Class BudgetVarianceWorkSheet

    Public Shared Function CreateWorkSheet(ByVal sheetName As String) As Excel.Worksheet
        PLLog.Trace1("Enter", "Scorecard")
        Dim ws As Excel.Worksheet
        Dim startRow As Integer = Globals.cHeaderID_RowShort
        Dim startCol As Integer = Globals.cHeaderID_Column
        Dim headerFontSize As Integer = Globals.cHeaderFontSize
        Dim headerBold As Boolean = True
        Dim headerUnderline As Boolean = True
        Dim headerWrapText As Boolean = True
        Dim headerHorizontalAlignment As Excel.XlHAlign = Excel.XlHAlign.xlHAlignCenter

        ws = Util.NewWorksheet(sheetName)

        With ws
            Util.AddColumnToSheet(ws, startCol + 0, 15, startRow, "IT Team", _
                headerFontSize, headerBold, headerUnderline, headerWrapText, headerHorizontalAlignment)
            Util.AddColumnToSheet(ws, startCol + 1, 15, startRow, "Programming Services" & vbLf & "(Acct 246)", _
                headerFontSize, headerBold, headerUnderline, headerWrapText, headerHorizontalAlignment)
            Util.AddColumnToSheet(ws, startCol + 2, 15, startRow, "SW" & vbLf & "(Acct 46)", _
                headerFontSize, headerBold, headerUnderline, headerWrapText, headerHorizontalAlignment)
            Util.AddColumnToSheet(ws, startCol + 3, 15, startRow, "SW Maint" & vbLf & "(Acct 243)", _
                headerFontSize, headerBold, headerUnderline, headerWrapText, headerHorizontalAlignment)
            Util.AddColumnToSheet(ws, startCol + 4, 15, startRow, "Consulting" & vbLf & "(Acct 34)", _
                headerFontSize, headerBold, headerUnderline, headerWrapText, headerHorizontalAlignment)
            Util.AddColumnToSheet(ws, startCol + 5, 15, startRow, "Salaries" & vbLf & "(Acct 10)", _
                headerFontSize, headerBold, headerUnderline, headerWrapText, headerHorizontalAlignment)
            Util.AddColumnToSheet(ws, startCol + 6, 15, startRow, "Capital Expenses" & vbLf & "(Acct 3880)", _
                headerFontSize, headerBold, headerUnderline, headerWrapText, headerHorizontalAlignment)

            ' Indicate the parts of the worksheet that need data entered by the user.  Protect the rest.

            With .Range(.Cells(startRow + 1, startCol), .Cells(startRow + Globals.cNumberTeams, startCol + 6))
                .Locked = False

                With .Interior
                    .Pattern = Excel.Constants.xlSolid
                    .PatternColorIndex = Excel.Constants.xlAutomatic
                    .ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3
                    .TintAndShade = 0.599993896298105
                    .PatternTintAndShade = 0
                End With
            End With

            .PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape

            .Protect(DrawingObjects:=False, Contents:=True, Scenarios:=False)
        End With

        PLLog.Trace1("Exit", "Scorecard")
        Return ws
    End Function

End Class
