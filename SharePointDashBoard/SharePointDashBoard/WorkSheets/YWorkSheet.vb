Imports PacificLife.Life

Public Class YWorkSheet

    Public Shared Function CreateWorkSheet(ByVal sheetName As String) As Excel.Worksheet
        PLLog.Trace1("Enter", Globals.cPLLOG_NAME)

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
            Util.AddColumnToSheet(ws, startCol + 1, 15, startRow, "Open ITRs", _
                headerFontSize, headerBold, headerUnderline, headerWrapText, headerHorizontalAlignment)
            Util.AddColumnToSheet(ws, startCol + 2, 15, startRow, "Closed ITRs", _
                headerFontSize, headerBold, headerUnderline, headerWrapText, headerHorizontalAlignment)
            Util.AddColumnToSheet(ws, startCol + 3, 15, startRow, "Active ITRs", _
                headerFontSize, headerBold, headerUnderline, headerWrapText, headerHorizontalAlignment)

            ' Indicate the parts of the worksheet that need data entered by the user.  Protect the rest.

            With .Range(.Cells(startRow + 1, startCol), .Cells(startRow + Globals.cNumberTeams, startCol + 3))
                .Locked = False

                With .Interior
                    .Pattern = Excel.Constants.xlSolid
                    .PatternColorIndex = Excel.Constants.xlAutomatic
                    .ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3
                    .TintAndShade = 0.599993896298105
                    .PatternTintAndShade = 0
                End With
            End With

            ' Add a row for "All IT" data based on the data provided for each team.  Then SUM and AVERAGE the data.

            .Cells(startRow + Globals.cNumberTeams + 1, startCol).Value = Globals.cAllITString

            .Range(.Cells(startRow + Globals.cNumberTeams + 1, startCol + 1), _
                    .Cells(startRow + Globals.cNumberTeams + 1, startCol + 3)).FormulaR1C1 = _
                    "=SUM(R[-" & Globals.cNumberTeams & "]C:R[-1]C)"

            .Range(.Cells(startRow + Globals.cNumberTeams + 2, startCol + 1), _
                    .Cells(startRow + Globals.cNumberTeams + 2, startCol + 3)).FormulaR1C1 = _
                    "=AVERAGE(R[-" & Globals.cNumberTeams + 1 & "]C:R[-2]C)"

            .Protect(DrawingObjects:=False, Contents:=True, Scenarios:=False)
        End With

        PLLog.Trace1("Exit", Globals.cPLLOG_NAME)

        Return ws
    End Function

End Class
