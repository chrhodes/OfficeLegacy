Option Explicit On

Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel
Imports PacificLife.Life

''' This class contains all the methods to process On-Time Data
''' except for charting which is in Charts.vb
Public Class OnTimeDataWorkSheet

#Region "AddOnTimeDataToPowerPoint"

#End Region

#Region "AddFormulas"

    Public Shared Sub AddFormulas( _
        ByVal fileNameAndPath As String, _
        ByVal sheetName As String, _
        ByVal teamName As String, _
        ByVal managerName As String, _
        ByVal extension As String _
    )
        PLLog.Trace1("Enter", "Scorecard")

        Dim columnHeadingsRow As Integer
        Dim dataSheet As Worksheet
        'Dim lastRow2 As Integer
        'Dim lastCol2 As Integer
        Dim lastRow As Integer
        Dim lastCol As Integer
        Dim row As Integer
        Dim weightedScheduledFormula As String
        Dim weightedActualFormula As String
        Dim averageFormula As String
        Dim sumActualReleaseDaysFormula As String
        Dim sumScheduledReleaseDaysFormula As String
        Dim numberReleasesFormula As String

        ' ToDo: Make this a constant and get rid of the Magic Cell references.

        With Globals.ThisAddIn.Application
            .Range("A2").Hyperlinks.Add( _
                Anchor:=.Range("A2"), _
                Address:="", _
                SubAddress:="'" & "On-Time Delivery Data" & "'!A1", _
                TextToDisplay:="On-Time Delivery Data")

            .Range("$B$4").Value = "Survey Name"
            .Range(Globals.cSD_SurveyNameCell).Value = "On-Time Delivery"

            .Range("$B$5").Value = "Survey Period"
            .Range(Globals.cSD_SurveyPeriodCell).Formula = "=Survey_Period"

            .Range("$B$12").Value = "Team"
            .Range(Globals.cOTD_TeamName_Cell).Value = teamName

            .Range("$B$11").Value = "Manager"
            .Range(Globals.cOTD_ManagerName_Cell).Value = managerName
            '.Range("D5").Value = extension

            .Range("$B$6").Value = "Input File"
            .Range(Globals.cOTD_InputFile_Cell).Value = fileNameAndPath

            .Range("$B$7").Value = "Input Sheet"
            .Range(Globals.cOTD_InputSheet_Cell).Value = sheetName

            '.Range("B9").Value = "Data Range"
            '.Range("B9").AddComment("Adjust to cover valid data range for charting")

            .Range("$B$13").Value = "Start Data Row"
            .Range("$B$14").Value = "End Data Row"
            .Range("$B$15").Value = "Start Data Column"
            .Range("$B$16").Value = "End Data Column"

            .Columns("G:G").NumberFormat = "0%"

            dataSheet = .ActiveSheet

            lastRow = dataSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row
            lastCol = dataSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column
            'lastRow2 = Util.FindLastRow(dataSheet.Range("A1"))
            'lastCol2 = Util.FindLastColumn(dataSheet.Range("A1"))

            ' TODO: Start this further down as data doesn't get pasted until later

            For row = 1 To lastRow
                'Debug.Print(row & " >" & dataSheet.Cells(row, 1).value & "<")

                Select Case TypeName(dataSheet.Cells(row, 1).value)
                    Case "String"
                        If "ID" = dataSheet.Cells(row, 1).value Then
                            columnHeadingsRow = row
                            ' We found the columnHeadings, now look for the columns we need.

                            GoTo Finished
                        End If

                End Select
            Next row

            MessageBox.Show("Cannot Find ""ID"" in header")
            Exit Sub
Finished:

            .Range(Globals.cOTD_StartDataRow_Cell).Value = columnHeadingsRow + 1
            .Range(Globals.cOTD_StartDataColumn_Cell).Value = Globals.cOTD_DataColumn
            .Range(Globals.cOTD_EndDataRow_Cell).Value = lastRow
            .Range(Globals.cOTD_EndDataColumn_Cell).Value = Globals.cOTD_DataColumn

            ' Now build the formulas used to process the data.

            weightedScheduledFormula = _
            "=SUM(" & _
                "INDIRECT(ADDRESS(" & Globals.cOTD_StartDataRow_Cell & ", " & 13 & ", TRUE)):" & _
                "INDIRECT(ADDRESS(" & Globals.cOTD_EndDataRow_Cell & ", " & 13 & ", TRUE))" & _
                ")"

            .Range("$B$24").Value = "Weighted Scheduled On Time Delivery"
            .Range(Globals.cOTD_WeightedScheduledOnTimePercentage_Cell).Formula = weightedScheduledFormula
            .Range(Globals.cOTD_WeightedScheduledOnTimePercentage_Cell).NumberFormat = "0%"

            weightedActualFormula = _
            "=SUM(" & _
                "INDIRECT(ADDRESS(" & Globals.cOTD_StartDataRow_Cell & ", " & 10 & ", TRUE)):" & _
                "INDIRECT(ADDRESS(" & Globals.cOTD_EndDataRow_Cell & ", " & 10 & ", TRUE))" & _
                ")"

            .Range("$B$25").Value = "Weighted Actual On Time Delivery"
            .Range(Globals.cOTD_WeightedActualOnTimePercentage_Cell).Formula = weightedActualFormula
            .Range(Globals.cOTD_WeightedActualOnTimePercentage_Cell).NumberFormat = "0%"

            averageFormula = _
            "=AVERAGE(" & _
                "INDIRECT(ADDRESS(" & Globals.cOTD_StartDataRow_Cell & ", " & Globals.cOTD_StartDataColumn_Cell & ", TRUE)):" & _
                "INDIRECT(ADDRESS(" & Globals.cOTD_EndDataRow_Cell & ", " & Globals.cOTD_EndDataColumn_Cell & ", TRUE))" & _
                ")"

            .Range("$B$26").Value = "Average On Time Delivery"
            .Range(Globals.cOTD_OnTimePercentage_Cell).Formula = averageFormula
            .Range(Globals.cOTD_OnTimePercentage_Cell).NumberFormat = "0%"

            numberReleasesFormula = "=R14C3 - R13C3 + 1"

            .Range("$B$27").Value = "Number of Releases"
            .Range(Globals.cOTD_NumberReleases_Cell).Formula = numberReleasesFormula
            .Range(Globals.cOTD_NumberReleases_Cell).NumberFormat = "0"

            sumActualReleaseDaysFormula = _
            "=SUM(" & _
                "INDIRECT(ADDRESS(" & Globals.cOTD_StartDataRow_Cell & ", " & 8 & ", TRUE)):" & _
                "INDIRECT(ADDRESS(" & Globals.cOTD_EndDataRow_Cell & ", " & 8 & ", TRUE))" & _
                ")"

            .Range("$B$28").Value = "Total Actual Release Days"
            .Range(Globals.cOTD_TotalActualReleaseDays_Cell).Formula = sumActualReleaseDaysFormula
            .Range(Globals.cOTD_TotalActualReleaseDays_Cell).NumberFormat = "0"

            sumScheduledReleaseDaysFormula = _
            "=SUM(" & _
                "INDIRECT(ADDRESS(" & Globals.cOTD_StartDataRow_Cell & ", " & 11 & ", TRUE)):" & _
                "INDIRECT(ADDRESS(" & Globals.cOTD_EndDataRow_Cell & ", " & 11 & ", TRUE))" & _
                ")"

            .Range("$B$29").Value = "Total Scheduled Release Days"
            .Range(Globals.cOTD_TotalScheduledReleaseDays_Cell).Formula = sumScheduledReleaseDaysFormula
            .Range(Globals.cOTD_TotalScheduledReleaseDays_Cell).NumberFormat = "0"

            ' Now add the headings and the formulas to the data section of the sheet.
            ' TODO: Get rid of magic column numbers

            Util.AddContentToCell(.Cells(columnHeadingsRow, 7), "Average On Time Delivery", _
                8, Globals.MakeBold.Yes, Globals.UnderLine.No, Globals.WrapText.Yes, Excel.XlHAlign.xlHAlignCenter)
            Util.AddContentToCell(.Cells(columnHeadingsRow, 8), "Total Actual Release Days", _
                8, Globals.MakeBold.Yes, Globals.UnderLine.No, Globals.WrapText.Yes, Excel.XlHAlign.xlHAlignCenter)
            Util.AddContentToCell(.Cells(columnHeadingsRow, 9), "% of Total Actual Release Days", _
                8, Globals.MakeBold.Yes, Globals.UnderLine.No, Globals.WrapText.Yes, Excel.XlHAlign.xlHAlignCenter)
            Util.AddContentToCell(.Cells(columnHeadingsRow, 10), "Weighted Actual On Time Delivery", _
                8, Globals.MakeBold.Yes, Globals.UnderLine.No, Globals.WrapText.Yes, Excel.XlHAlign.xlHAlignCenter)
            Util.AddContentToCell(.Cells(columnHeadingsRow, 11), "Total Scheduled Release Days", _
                8, Globals.MakeBold.Yes, Globals.UnderLine.No, Globals.WrapText.Yes, Excel.XlHAlign.xlHAlignCenter)
            Util.AddContentToCell(.Cells(columnHeadingsRow, 12), "% of Total Scheduled Release Days", _
                8, Globals.MakeBold.Yes, Globals.UnderLine.No, Globals.WrapText.Yes, Excel.XlHAlign.xlHAlignCenter)
            Util.AddContentToCell(.Cells(columnHeadingsRow, 13), "Weighted Scheduled On Time Delivery", _
                8, Globals.MakeBold.Yes, Globals.UnderLine.No, Globals.WrapText.Yes, Excel.XlHAlign.xlHAlignCenter)

            Dim formula As String

            ' Average OnTime Delivery

            formula = "=IF("
            formula = formula & "AND(NOT(ISBLANK(RC[-3])),ISNUMBER(RC[-3]), "
            formula = formula & "    NOT(ISBLANK(RC[-2])),ISNUMBER(RC[-2]), "
            formula = formula & "    NOT(ISBLANK(RC[-1])),ISNUMBER(RC[-1])"
            formula = formula & "),NETWORKDAYS(RC[-3],RC[-1])/NETWORKDAYS(RC[-3],RC[-2]),""Not Used"")"

            dataSheet.Range(.Cells(columnHeadingsRow + 1, Globals.cOTD_DataColumn), .Cells(lastRow, Globals.cOTD_DataColumn)).Value = formula
            dataSheet.Range(.Cells(columnHeadingsRow + 1, Globals.cOTD_DataColumn), .Cells(lastRow, Globals.cOTD_DataColumn)).NumberFormat = "0%"

            ' Total Actual Release Days

            formula = "=IF(AND(ISNUMBER(RC[-2]), ISNUMBER(RC[-4])), RC[-2] - RC[-4], ""Not Used"")"
            dataSheet.Range(.Cells(columnHeadingsRow + 1, Globals.cOTD_DataColumn + 1), .Cells(lastRow, Globals.cOTD_DataColumn + 1)).Value = formula
            dataSheet.Range(.Cells(columnHeadingsRow + 1, Globals.cOTD_DataColumn + 1), .Cells(lastRow, Globals.cOTD_DataColumn + 1)).NumberFormat = "0"

            ' % Total Actual Release Days

            formula = "=IF(ISNUMBER(RC[-1]), RC[-1] / R28C3, ""Not Used"")"
            dataSheet.Range(.Cells(columnHeadingsRow + 1, Globals.cOTD_DataColumn + 2), .Cells(lastRow, Globals.cOTD_DataColumn + 2)).Value = formula
            dataSheet.Range(.Cells(columnHeadingsRow + 1, Globals.cOTD_DataColumn + 2), .Cells(lastRow, Globals.cOTD_DataColumn + 2)).NumberFormat = "0.00%"

            ' Weighted Actual OnTime Delivery

            formula = "=IF(ISNUMBER(RC[-3]), RC[-3] * RC[-1], ""Not Used"")"
            dataSheet.Range(.Cells(columnHeadingsRow + 1, Globals.cOTD_DataColumn + 3), .Cells(lastRow, Globals.cOTD_DataColumn + 3)).Value = formula
            dataSheet.Range(.Cells(columnHeadingsRow + 1, Globals.cOTD_DataColumn + 3), .Cells(lastRow, Globals.cOTD_DataColumn + 3)).NumberFormat = "0.00%"

            ' Total Scheduled Release Days

            formula = "=IF(AND(ISNUMBER(RC[-6]), ISNUMBER(RC[-7])), RC[-6] - RC[-7], ""Not Used"")"
            dataSheet.Range(.Cells(columnHeadingsRow + 1, Globals.cOTD_DataColumn + 4), .Cells(lastRow, Globals.cOTD_DataColumn + 4)).Value = formula
            dataSheet.Range(.Cells(columnHeadingsRow + 1, Globals.cOTD_DataColumn + 4), .Cells(lastRow, Globals.cOTD_DataColumn + 4)).NumberFormat = "0"

            ' % Total Scheduled Release Days

            formula = "=IF(ISNUMBER(RC[-1]), RC[-1] / R29C3, ""Not Used"")"
            dataSheet.Range(.Cells(columnHeadingsRow + 1, Globals.cOTD_DataColumn + 5), .Cells(lastRow, Globals.cOTD_DataColumn + 5)).Value = formula
            dataSheet.Range(.Cells(columnHeadingsRow + 1, Globals.cOTD_DataColumn + 5), .Cells(lastRow, Globals.cOTD_DataColumn + 5)).NumberFormat = "0.00%"

            ' Weighted Scheduled OnTime Delivery

            formula = "=IF(ISNUMBER(RC[-6]), RC[-6] * RC[-1], ""Not Used"")"
            dataSheet.Range(.Cells(columnHeadingsRow + 1, Globals.cOTD_DataColumn + 6), .Cells(lastRow, Globals.cOTD_DataColumn + 6)).Value = formula
            dataSheet.Range(.Cells(columnHeadingsRow + 1, Globals.cOTD_DataColumn + 6), .Cells(lastRow, Globals.cOTD_DataColumn + 6)).NumberFormat = "0.00%"

        End With

        PLLog.Trace1("Exit", "Scorecard")

    End Sub

#End Region

#Region "CleanUpDataSheet"

    Public Shared Sub CleanUpDataSheet()
        PLLog.Trace1("Enter", "Scorecard")

        Dim columnHeadingsRow As Integer
        Dim foundStart As Boolean
        Dim foundEnd As Boolean
        Dim endDataRow As Integer
        Dim dataSheet As Worksheet
        Dim lastRow As Integer
        Dim lastCol As Integer
        'Dim lastRow2 As Integer
        'Dim lastCol2 As Integer
        Dim row As Integer
        Dim col As Integer

        With Globals.ThisAddIn.Application

            Debug.Print(.ActiveSheet.Name)
            dataSheet = .ActiveSheet

            lastRow = dataSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row
            lastCol = dataSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column
            'lastRow2 = Util.FindLastRow(.Range("A1"))
            'lastCol2 = Util.FindLastColumn(.Range("A1"))

            For row = 1 To lastRow
                'Debug.Print(row & " >" & dataSheet.Cells(row, 1).value & "<")

                Select Case TypeName(dataSheet.Cells(row, 1).value)
                    Case "String"
                        If foundStart Then
                            ' Just walked off end of data
                            foundEnd = True
                            endDataRow = row - 1
                        Else
                            If "ID" = dataSheet.Cells(row, 1).value Then
                                columnHeadingsRow = row
                                ' We found the columnHeadings, now look for the columns we need.

                                For col = lastCol To 2 Step -1
                                    'Debug.Print(row & ":" & col & " >" & dataSheet.Cells(row, col).value & "<")

                                    Select Case TypeName(dataSheet.Cells(row, col).value)
                                        Case "String"
                                            Select Case dataSheet.Cells(row, col).value.ToString()
                                                Case "Release Name"

                                                Case "Description"

                                                Case "Release Actual Start Date"

                                                Case "Release Scheduled Impl Date"

                                                Case "Release Actual Impl Date"

                                                Case Else
                                                    ' Don't need this column
                                                    'dataSheet.Range(.Cells(row, col), .Cells(.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row, col)).Delete(Shift:=Excel.XlDirection.xlToLeft)
                                                    dataSheet.Range(.Cells(row, col), .Cells(lastRow, col)).Delete(Shift:=Excel.XlDirection.xlToLeft)

                                            End Select

                                        Case Else
                                            ' Don't need this column
                                            'dataSheet.Range(.Cells(row, col), .Cells(.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row, col)).Delete(Shift:=Excel.XlDirection.xlToLeft)
                                            dataSheet.Range(.Cells(row, col), .Cells(lastRow, col)).Delete(Shift:=Excel.XlDirection.xlToLeft)
                                    End Select
                                Next col

                                GoTo Finished
                            End If
                        End If

                    Case "Empty"
                        If foundStart Then
                            ' Just walked off end of data
                            foundEnd = True
                            endDataRow = row - 1
                            Exit For
                        End If

                    Case Else

                End Select
            Next row

Finished:
            .Columns("A:A").ColumnWidth = Globals.cIDColumnWidth
            .Columns("B:B").ColumnWidth = Globals.cReleaseNameColumnWidth
            .Columns("C:C").ColumnWidth = Globals.cDescriptionColumnWidth
            .Columns("D:D").ColumnWidth = Globals.cDateColumnWidth
            .Columns("E:E").ColumnWidth = Globals.cDateColumnWidth
            .Columns("F:F").ColumnWidth = Globals.cDateColumnWidth
            .Columns("G:G").ColumnWidth = Globals.cPercentColumnWidth
            .Cells.Rows.AutoFit()

        End With

        dataSheet = Nothing

        PLLog.Trace1("Exit", "Scorecard")

    End Sub

#End Region

#Region "DeleteOnTimeDataSheets"

    '------------------------------------------------------------------------------------------
    '
    ' DeleteOnTimeDataSheets()
    '
    ' Delete selected sheets.
    '
    '------------------------------------------------------------------------------------------



#End Region

#Region "OpenInputFileAndVerifyDataLayout"
    '------------------------------------------------------------------------------------------
    '
    ' VerifyInputFileDataLayout
    '
    ' Verify InputFile layout.
    '
    '------------------------------------------------------------------------------------------

    Public Shared Function OpenInputFileAndVerifyDataLayout(ByVal filePathAndName As String, ByVal sheetName As String, ByRef errorMessage As String) As Boolean
        PLLog.Trace1("Enter", "Scorecard")

        Dim lastRow As Integer
        Dim lastCol As Integer
        'Dim lastRow2 As Integer
        'Dim lastCol2 As Integer
        Dim dataSheet As Worksheet
        Dim row As Integer
        Dim col As Integer
        Dim sheetValid As Boolean
        Dim idColumnFound As Integer
        Dim releaseNameFound As Integer
        Dim descriptionFound As Integer
        Dim releaseActualStartDateFound As Integer
        Dim releaseScheduledImplDateFound As Integer
        Dim releaseActualImplDateFound As Integer

        OpenInputFileAndVerifyDataLayout = True
        errorMessage = "File Invalid:" & vbLf
        Globals.ThisAddIn.Application.Workbooks.Open(Filename:=filePathAndName, ReadOnly:=True)
        sheetValid = True

        If Not ValidateInputFileWorksheetExists(Globals.ThisAddIn.Application.ActiveWorkbook, sheetName) Then
            sheetValid = False
            GoTo Finished
        End If

        dataSheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets(sheetName)

        ' Make no assumption about where the data might be other than the "ID" field is in
        ' Column 1.  Let Excel tell us the maximum boundaries of the file.

        lastRow = dataSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row
        lastCol = dataSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column

        'lastRow = Util.FindLastRow(dataSheet.Range("A1"))
        'lastCol = Util.FindLastColumn(dataSheet.Range("A1"))

        ' Validate file
        ' Now look for the rows and columns we need.
        For row = 1 To lastRow
            'Debug.Print(row & " >" & dataSheet.Cells(row, 1).value & "<")

            Select Case TypeName(dataSheet.Cells(row, 1).value)
                Case "String"
                    If "ID" = dataSheet.Cells(row, 1).value Then
                        idColumnFound = idColumnFound + 1
                        ' We found the columnHeadings, now look for the columns we need.

                        For col = lastCol To 2 Step -1
                            'Debug.Print(row & ":" & col & " >" & dataSheet.Cells(row, col).value & "<")

                            Select Case TypeName(dataSheet.Cells(row, col).value)
                                Case "String"
                                    Select Case dataSheet.Cells(row, col).value.ToString()
                                        Case "Release Name"
                                            releaseNameFound = releaseNameFound + 1

                                        Case "Description"
                                            descriptionFound = descriptionFound + 1

                                        Case "Release Actual Start Date"
                                            releaseActualStartDateFound = releaseActualStartDateFound + 1

                                        Case "Release Scheduled Impl Date"
                                            releaseScheduledImplDateFound = releaseScheduledImplDateFound + 1

                                        Case "Release Actual Impl Date"
                                            releaseActualImplDateFound = releaseActualImplDateFound + 1

                                    End Select
                            End Select
                        Next col

                        GoTo Finished
                    End If

            End Select
        Next row

Finished:
        If sheetValid = False Then
            errorMessage = errorMessage & "'" & sheetName & "' tab does not exist"
            OpenInputFileAndVerifyDataLayout = False
            Exit Function
        End If

        If idColumnFound = 0 Then
            errorMessage = errorMessage & "Could Not find 'ID' column" & vbLf
            OpenInputFileAndVerifyDataLayout = False
            Exit Function
        ElseIf idColumnFound > 1 Then
            errorMessage = errorMessage & "Found multiple 'ID' columns" & vbLf
            OpenInputFileAndVerifyDataLayout = False
            Exit Function
        End If

        If releaseNameFound = 0 Then
            errorMessage = errorMessage & "Could Not find 'Release Name' column" & vbLf
            OpenInputFileAndVerifyDataLayout = False
        ElseIf releaseNameFound > 1 Then
            errorMessage = errorMessage & "Found multiple 'Release Name' columns" & vbLf
            OpenInputFileAndVerifyDataLayout = False
        End If

        If descriptionFound = 0 Then
            errorMessage = errorMessage & "Could Not find 'Description' column" & vbLf
            OpenInputFileAndVerifyDataLayout = False
        ElseIf releaseNameFound > 1 Then
            errorMessage = errorMessage & "Found multiple 'Description' columns" & vbLf
            OpenInputFileAndVerifyDataLayout = False
        End If

        If releaseActualStartDateFound = 0 Then
            errorMessage = errorMessage & "Could Not find 'Release Actual Start Date' column" & vbLf
            OpenInputFileAndVerifyDataLayout = False
        ElseIf releaseActualStartDateFound > 1 Then
            errorMessage = errorMessage & "Found multiple 'Release Actual Start Date' columns" & vbLf
            OpenInputFileAndVerifyDataLayout = False
        End If

        If releaseScheduledImplDateFound = 0 Then
            errorMessage = errorMessage & "Could Not find 'Release Scheduled Impl Date' column" & vbLf
            OpenInputFileAndVerifyDataLayout = False
        ElseIf releaseScheduledImplDateFound > 1 Then
            errorMessage = errorMessage & "Found multiple 'Release Scheduled Impl Date' columns" & vbLf
            OpenInputFileAndVerifyDataLayout = False
        End If

        If releaseActualImplDateFound = 0 Then
            errorMessage = errorMessage & "Could Not find 'Release Actual Impl Date' column" & vbLf
            OpenInputFileAndVerifyDataLayout = False
        ElseIf releaseActualImplDateFound > 1 Then
            errorMessage = errorMessage & "Found multiple 'Release Actual Impl Date' columns" & vbLf
            OpenInputFileAndVerifyDataLayout = False
        End If

        If OpenInputFileAndVerifyDataLayout Then
            errorMessage = "File Valid"
        End If

        PLLog.Trace1("Exit", "Scorecard")
    End Function
#End Region

#Region "ValidateInputSheet"
    Public Shared Function ValidateInputSheet() As Boolean
        MsgBox("ToDo: Validate current sheet is well formed and something is selected.")
        ValidateInputSheet = True
    End Function
#End Region

#Region "ValidateOnTimeDataFiles"
    '------------------------------------------------------------------------------------------
    '
    ' ValidateOnTimeDataFiles()
    '
    ' Verify each selected file is properly formatted and contains the information needed.
    '
    '------------------------------------------------------------------------------------------


#End Region

#Region "ValidateInputFileWorksheetExists"
    '------------------------------------------------------------------------------------------
    '
    ' ValidateInputFileWorksheetExists
    '
    ' Verify File Exists and can find worksheet containing data.
    '
    '------------------------------------------------------------------------------------------

    Public Shared Function ValidateInputFileWorksheetExists(ByVal Workbook As Workbook, ByVal sheetName As String) As Boolean
        PLLog.Trace1("Enter", "Scorecard")

        Dim ws As Worksheet

        ValidateInputFileWorksheetExists = False

        For Each ws In Workbook.Sheets
            If ws.Name = sheetName Then
                ValidateInputFileWorksheetExists = True
                Exit For
            End If
        Next ws

        PLLog.Trace1("Exit", "Scorecard")
    End Function
#End Region

End Class
