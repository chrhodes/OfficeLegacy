Imports PacificLife.Life
Imports Microsoft.Office.Interop

Public Class YWorkSheet

    Public Shared Function CreateWorkSheet(ByVal sheetName As String) As Excel.Worksheet
        'PLLog.Trace1("Enter", Common.PROJECT_NAME)

        'Dim ws As Excel.Worksheet
        'Dim startRow As Integer = Common.cHeaderID_RowShort
        'Dim startCol As Integer = Common.cHeaderID_Column
        'Dim headerFontSize As Integer = Common.cHeaderFontSize
        'Dim headerBold As Boolean = True
        'Dim headerUnderline As Boolean = True
        'Dim headerWrapText As Boolean = True
        'Dim headerHorizontalAlignment As Excel.XlHAlign = Excel.XlHAlign.xlHAlignCenter

        'ws = Common.ExcelUtil.NewWorksheet(sheetName)

        'With ws
        '    Common.ExcelUtil.AddColumnToSheet(ws, startCol + 0, 15, startRow, "IT Team", _
        '        headerFontSize, headerBold, headerUnderline, headerWrapText, headerHorizontalAlignment)
        '    Common.ExcelUtil.AddColumnToSheet(ws, startCol + 1, 15, startRow, "Open ITRs", _
        '        headerFontSize, headerBold, headerUnderline, headerWrapText, headerHorizontalAlignment)
        '    Common.ExcelUtil.AddColumnToSheet(ws, startCol + 2, 15, startRow, "Closed ITRs", _
        '        headerFontSize, headerBold, headerUnderline, headerWrapText, headerHorizontalAlignment)
        '    Common.ExcelUtil.AddColumnToSheet(ws, startCol + 3, 15, startRow, "Active ITRs", _
        '        headerFontSize, headerBold, headerUnderline, headerWrapText, headerHorizontalAlignment)

        '    ' Indicate the parts of the worksheet that need data entered by the user.  Protect the rest.

        '    With .Range(.Cells(startRow + 1, startCol), .Cells(startRow + Common.cNumberTeams, startCol + 3))
        '        .Locked = False

        '        With .Interior
        '            .Pattern = Excel.Constants.xlSolid
        '            .PatternColorIndex = Excel.Constants.xlAutomatic
        '            .ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3
        '            .TintAndShade = 0.599993896298105
        '            .PatternTintAndShade = 0
        '        End With
        '    End With

        '    ' Add a row for "All IT" data based on the data provided for each team.  Then SUM and AVERAGE the data.

        '    .Cells(startRow + Common.cNumberTeams + 1, startCol).Value = Common.cAllITString

        '    .Range(.Cells(startRow + Common.cNumberTeams + 1, startCol + 1), _
        '            .Cells(startRow + Common.cNumberTeams + 1, startCol + 3)).FormulaR1C1 = _
        '            "=SUM(R[-" & Common.cNumberTeams & "]C:R[-1]C)"

        '    .Range(.Cells(startRow + Common.cNumberTeams + 2, startCol + 1), _
        '            .Cells(startRow + Common.cNumberTeams + 2, startCol + 3)).FormulaR1C1 = _
        '            "=AVERAGE(R[-" & Common.cNumberTeams + 1 & "]C:R[-2]C)"

        '    .Protect(DrawingObjects:=False, Contents:=True, Scenarios:=False)
        'End With

        'PLLog.Trace1("Exit", Common.PROJECT_NAME)

        'Return ws
        Return Nothing
    End Function

End Class
