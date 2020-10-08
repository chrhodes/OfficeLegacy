Imports PacificLife.Life

Public Class XWorkSheet

    Public Shared Function CreateWorkSheet(ByVal sheetName As String) As Excel.Worksheet
        PLLog.Trace1("Enter", "NEWAPPNAME")

        Dim ws As Excel.Worksheet
        Dim startRow As Integer = Globals.cSD_HeaderIDRow
        Dim startCol As Integer = Globals.cHeaderID_Column
        Dim headerFontSize As Integer = Globals.cHeaderFontSize
        Dim headerBold As Boolean = True
        Dim headerUnderline As Boolean = True

        ws = Util.NewWorksheet(sheetName)

        ' TODO: Get rid of the magic constants and booleans.
        With ws
            Util.AddColumnToSheet(ws, startCol + 0, 15, False, startRow, "Team", headerFontSize, headerBold, headerUnderline)
            Util.AddColumnToSheet(ws, startCol + 1, 15, False, startRow, "Score", headerFontSize, headerBold, headerUnderline)
            Util.AddColumnToSheet(ws, startCol + 2, 15, True, startRow, "On-Time % (Scheduled Weighted)", headerFontSize, headerBold, headerUnderline)
            Util.AddColumnToSheet(ws, startCol + 3, 15, True, startRow, "On-Time % (Actual Weighted)", headerFontSize, headerBold, headerUnderline)
            Util.AddColumnToSheet(ws, startCol + 4, 12, True, startRow, "On-Time % (Average)", headerFontSize, headerBold, headerUnderline)
            Util.AddColumnToSheet(ws, startCol + 5, 12, False, startRow, "# Releases", headerFontSize, headerBold, headerUnderline)
            Util.AddColumnToSheet(ws, startCol + 6, 17, False, startRow, "Manager", headerFontSize, headerBold, headerUnderline)
            Util.AddColumnToSheet(ws, startCol + 7, 11, False, startRow, "Extension", headerFontSize, headerBold, headerUnderline)
            Util.AddColumnToSheet(ws, startCol + 8, 85, False, startRow, "Source File (Full Path)", headerFontSize, headerBold, headerUnderline)
            Util.AddColumnToSheet(ws, startCol + 9, 15, False, startRow, "SheetName", headerFontSize, headerBold, headerUnderline)
            Util.AddColumnToSheet(ws, startCol + 10, 15, False, startRow, "File Errors", headerFontSize, headerBold, headerUnderline)
            Util.AddColumnToSheet(ws, startCol + 11, 25, False, startRow, "Data Sheet Name", headerFontSize, headerBold, headerUnderline)

            ' Indicate the parts of the worksheet that need data entered by the user.  Protect the rest.

            With .Range(.Cells(startRow + 1, startCol), .Cells(startRow + Globals.cNumberTeams, startCol))
                .Locked = False

                With .Interior
                    .Pattern = Excel.Constants.xlSolid
                    .PatternColorIndex = Excel.Constants.xlAutomatic
                    .ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3
                    .TintAndShade = 0.599993896298105
                    .PatternTintAndShade = 0
                End With

                ' Add a List of Team names to choose from.

                With .Validation
                    .Delete()
                    .Add( _
                        Type:=Excel.XlDVType.xlValidateList, _
                        AlertStyle:=Excel.XlDVAlertStyle.xlValidAlertStop, _
                        Operator:=Excel.XlFormatConditionOperator.xlBetween, _
                        Formula1:="=Team_Names")
                    .IgnoreBlank = True
                    .InputTitle = ""
                    .ErrorTitle = ""
                    .InputMessage = ""
                    .ErrorMessage = ""
                    .ShowInput = True
                    .ShowError = True
                End With
            End With

            ' Add lookups for the associated manager and extension.

            With .Range(.Cells(startRow + 1, startCol + 2), _
                        .Cells(startRow + Globals.cNumberTeams, startCol + 2))
                .FormulaR1C1 = "=LOOKUP(RC[-5],Team_Names,Team_Managers)"
            End With

            With .Range(.Cells(startRow + 1, startCol + 3), _
            .Cells(startRow + Globals.cNumberTeams, startCol + 3))
                .FormulaR1C1 = "=LOOKUP(RC[-6],Team_Names,Team_Managers_Extensions)"
            End With

            ' TODO: Get rid of the yucky magic numbers

            With .Range(.Cells(startRow + 1, startCol + 7), .Cells(startRow + Globals.cNumberTeams, startCol + 8))
                .Locked = False

                With .Interior
                    .Pattern = Excel.Constants.xlSolid
                    .PatternColorIndex = Excel.Constants.xlAutomatic
                    .ThemeColor = Excel.XlThemeColor.xlThemeColorAccent3
                    .TintAndShade = 0.599993896298105
                    .PatternTintAndShade = 0
                End With
            End With

            .Cells(startRow + Globals.cNumberTeams + 1, startCol).Value = Globals.cAllITString
            .Cells(startRow + Globals.cNumberTeams + 1, startCol + 1).FormulaR1C1 = "=AVERAGE(R[-" & Globals.cNumberTeams & "]C:R[-1]C)"
            '.Range(.Cells(startRow + 1, startCol + 8), .Cells(startRow + 1 + Globals.cNumberTeams, startCol + 8)).FormulaR1C1 = "=RC[-8]"
            .Range(.Cells(startRow + 1, startCol + 1), .Cells(startRow + 1 + Globals.cNumberTeams, startCol + 1)).Style = "Percent"

            .Protect(DrawingObjects:=False, Contents:=True, Scenarios:=False)
        End With

        PLLog.Trace1("Enter", "NEWAPPNAME")

        Return ws
    End Function

End Class
