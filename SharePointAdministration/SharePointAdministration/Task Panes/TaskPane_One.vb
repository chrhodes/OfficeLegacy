Imports System.ComponentModel
Imports System.Reflection
Imports System.Runtime

Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel

Public Class TaskPane_One
    'Private _teams As Data.DataSet

    Private _inputFilePath As String = Globals.cDEFAULT_ONTIMEDATA_FOLDER

    Private Sub btnAddOnTimeDataSheets_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddOnTimeDataSheets.Click
        Dim rng As Excel.Range
        Dim inputPathAndFileName As String
        Dim inputSheetName As String
        Dim inputWorkbook As Excel.Workbook
        Dim currentWorkbook As Excel.Workbook
        Dim errorMessage As String = ""
        Dim inputWorkSheet As Excel.Worksheet
        Dim currentSheet As Excel.Worksheet
        Dim dataSheet As Excel.Worksheet
        Dim dataSheetName As String
        'Dim lastRow2 As Integer
        'Dim lastCol2 As Integer
        Dim lastRow As Integer
        Dim lastCol As Integer
        'Dim startRow As Integer
        'Dim startColumn As Integer
        'Dim row As Integer

        ' TODO: This should probably mostly move to OnTimeDataWorkSheet

        If Not OnTimeDataWorkSheet.ValidateInputSheet() Then
            MsgBox("Invalid Worksheet.")
        Else
            With Globals.ThisAddIn.Application
                Util.ScreenUpdatesOff()
                currentSheet = .ActiveSheet
                currentWorkbook = .ActiveWorkbook

                For Each rng In CType(.Selection, Microsoft.Office.Interop.Excel.Range)
                    inputPathAndFileName = rng.Value
                    inputSheetName = rng.Offset(0, Globals.cOTD_SheetName_Offset).Value

                    If inputPathAndFileName <> "" Then
                        If Not OnTimeDataWorkSheet.OpenInputFileAndVerifyDataLayout(inputPathAndFileName, inputSheetName, errorMessage) Then
                            MsgBox("Invalid input file: " & errorMessage)
                        Else
                            ' We have opened an file containing valid On-Time Data.  Extract what we need.
                            ' Make as few assumptions as possible.  The sheet may be poorly formatted or contain
                            ' merged cells.  Assume the worst and just grab everything that might contain data.
                            ' We will clean up the file later in CleanUpDataSheet()

                            inputWorkbook = .ActiveWorkbook
                            inputWorkSheet = .ActiveWorkbook.Sheets(inputSheetName)

                            lastRow = inputWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row
                            lastCol = inputWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column
                            ' We cannot safely use these routines as they don't behave well with merged cells
                            'lastRow2 = Util.FindLastRow(.ActiveSheet.Range("A1"))
                            'lastCol2 = Util.FindLastColumn(.ActiveSheet.Range("A1"))

                            inputWorkSheet.Range(inputWorkSheet.Cells(1, 1), inputWorkSheet.Cells(lastRow, lastCol)).Copy()

                            ' Now return to the Scorecard Workbook, add a sheet for the data, and give it a suitable name.

                            currentWorkbook.Activate()
                            dataSheet = .Sheets.Add()
                            dataSheetName = rng.Offset(0, Globals.cOTD_TeamName_Offset).Value & Globals.cOTD_MetricName
                            .ActiveSheet.Name = dataSheetName

                            '' Some people have added fancy formulas to their reports.  To ensure that we
                            '' get just the values so we can safely delete rows and not have information
                            '' changing just paste in the values and any number formats.

                            '.Range(Globals.cRawDataCell).PasteSpecial( _
                            '    Paste:=XlPasteType.xlPasteValuesAndNumberFormats, _
                            '    Operation:=XlPasteSpecialOperation.xlPasteSpecialOperationNone, _
                            '    SkipBlanks:=False, Transpose:=False)

                            '' Unfortunately pasting the values does not bring the formatting along.
                            '' This makes is hard to show the data to someone and have them recognize it
                            '' so, add the formatting back on top of the data

                            '.Range(Globals.cRawDataCell).PasteSpecial( _
                            '        Paste:=XlPasteType.xlPasteFormats, _
                            '        Operation:=XlPasteSpecialOperation.xlPasteSpecialOperationNone, _
                            '        SkipBlanks:=False, Transpose:=False)

                            ' Ah, man.  This just sucks.  If a cell has formatting applied to just part of 
                            ' it's contents, when pasting the format on top of the values it only applies
                            ' the first formatting it sees.  So, if for example the cell has two dates
                            ' and the first is stiken but the second is not, pasting the format makes
                            ' the whole cell have striken values.  Groan.

                            ' So, looks like we are back to just pasting everything.  Will have to manually
                            ' fix the ID column if needed.

                            .Range(Globals.cRawDataCell).PasteSpecial( _
                                    Paste:=XlPasteType.xlPasteAll, _
                                    Operation:=XlPasteSpecialOperation.xlPasteSpecialOperationNone, _
                                    SkipBlanks:=False, Transpose:=False)

                            OnTimeDataWorkSheet.CleanUpDataSheet()

                            ' Might want to expose this from TaskPane

                            OnTimeDataWorkSheet.AddFormulas( _
                                inputPathAndFileName, inputSheetName, _
                                rng.Offset(0, Globals.cOTD_TeamName_Offset).Value, _
                                rng.Offset(0, Globals.cOTD_Manager_Offset).Value, _
                                rng.Offset(0, Globals.cOTD_Extension_Offset).Value)

                            ' Quietly close the input workbook.
                            ' TODO: Make this a Util Function

                            .DisplayAlerts = False
                            inputWorkbook.Close(False)
                            .DisplayAlerts = True

                            ' Now record what sheet just got added to the Input Worksheet
                            ' along with a link to the sheet.

                            currentSheet.Unprotect()
                            rng.Offset(0, Globals.cOTD_DataSheet_Offset).Value = dataSheetName
                            rng.Offset(0, Globals.cOTD_DataSheet_Offset).Hyperlinks.Add( _
                                        Anchor:=rng.Offset(0, Globals.cOTD_DataSheet_Offset), _
                                        Address:="", _
                                        SubAddress:="'" & dataSheetName & "'!A1", _
                                        TextToDisplay:=dataSheetName)

                            ' Finally add the On-Time Percentage data.

                            rng.Offset(0, Globals.cOTD_WeightedScheduledOnTimePercentage_Offset).Value = _
                                "='" & dataSheetName & "'!" & Globals.cOTD_WeightedScheduledOnTimePercentage_Cell

                            rng.Offset(0, Globals.cOTD_WeightedActualOnTimePercentage_Offset).Value = _
                                "='" & dataSheetName & "'!" & Globals.cOTD_WeightedActualOnTimePercentage_Cell

                            rng.Offset(0, Globals.cOTD_OnTimePercentage_Offset).Value = _
                                "='" & dataSheetName & "'!" & Globals.cOTD_OnTimePercentage_Cell

                            rng.Offset(0, Globals.cOTD_NumberReleases_Offset).Value = _
                                "='" & dataSheetName & "'!" & Globals.cOTD_NumberReleases_Cell

                            currentSheet.Protect()
                        End If
                    End If
                Next rng
            End With
        End If

        Util.ScreenUpdatesOn()
        inputWorkbook = Nothing
    End Sub

    Private Sub btnAddOnTimeDataToPowerPoint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddOnTimeDataToPowerPoint.Click
        Dim rng As Excel.Range
        Dim ws As Excel.Worksheet

        Try
            With Globals.ThisAddIn.Application
                For Each rng In CType(.Selection, Microsoft.Office.Interop.Excel.Range)
                    ' Verify we have a file name and a sheet name

                    If "" <> rng.Value And "" <> rng.Offset(0, Globals.cOTD_DataSheet_Offset).Value Then
                        ' If yes, select the sheet and add the, hopefully existing, chart
                        ' to PowerPoint.

                        ws = .Sheets(rng.Offset(0, Globals.cOTD_DataSheet_Offset).Value)
                        'PowerPointIntegration.AddOnTimeDataToPowerPoint(ws)
                        ws = Nothing
                    End If
                Next rng
            End With
        Catch ex As Exception
            MessageBox.Show("Exception: btnAddOnTimeDataToPowerPoint_Click()")
        End Try
    End Sub

    Private Sub btnBrowseForOnTimeDataFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowseForOnTimeDataFile.Click
        Dim inputFile As String = Util.GetFile(_inputFilePath, "Select File Containing On-Time Data")

        If inputFile.Length > 0 Then
            ' Save the folder in case the user browsed to a new location.
            _inputFilePath = System.IO.Path.GetDirectoryName(inputFile)

            ' Save the file on the worksheet.
            Globals.ThisAddIn.Application.ActiveCell.Value = inputFile
        End If

    End Sub

    Private Sub TaskPane_OnTimeDelivery_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' Ensure that any config data we need is available.  Ok to call multiple times.
        'Config.IntializeApplication()

        For Each dataTable As Data.DataTable In Config.ConfigInfo.Tables
            'Debug.Print(dataTable.TableName)

            'For Each dataColumn As Data.DataColumn In dataTable.Columns
            '    Debug.Print(dataColumn.ColumnName)
            'Next

            Select Case dataTable.TableName
                Case "team"
                    For Each dataRow As Data.DataRow In dataTable.Rows
                        Me.clbOnTimeTeams.Items.Add(dataRow.Item("name")).ToString()
                        Me.cbOnTimeTeams.Items.Add(dataRow.Item("name")).ToString()
                        'Debug.Print(dataRow.Item("name").ToString())
                        'Debug.Print(dataRow.Item("id").ToString())
                        'Debug.Print(dataRow.Item("team_Id").ToString())
                    Next

                    'Case "manager"
                    '    For Each dataRow As Data.DataRow In dataTable.Rows
                    '        Debug.Print(dataRow.Item("manager_Text").ToString())
                    '        Debug.Print(dataRow.Item("ext").ToString())
                    '        Debug.Print(dataRow.Item("team_Id").ToString())
                    '    Next

            End Select
        Next
    End Sub

    Private Sub cbOnTimeTeams_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbOnTimeTeams.SelectedIndexChanged
        Dim ws As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet

        ''' TODO: Verify we have the right worksheet open before splatting things down.

        'ws.Range(Globals.cOTD_TeamNameCell).Value = cbOnTimeTeams.SelectedItem.ToString()
    End Sub

    Private Sub btnDeleteOnTimeDataSheets_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteOnTimeDataSheets.Click
        Dim rng As Excel.Range

        With Globals.ThisAddIn.Application
            For Each rng In CType(.Selection, Microsoft.Office.Interop.Excel.Range)
                .Sheets(rng.Offset(0, 3).Value).Delete()
                Util.ProtectSheet(.ActiveSheet, False)
                ' TODO: Get rid of magic numbers.
                rng.Offset(0, 2).Value = ""
                rng.Offset(0, 3).Clear()
                rng.Offset(0, -3).Clear()
                Util.ProtectSheet(.ActiveSheet, True)
            Next rng
        End With
    End Sub

    Private Sub btnValidateOnTimeDataFiles_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnValidateOnTimeDataFiles.Click
        Dim rng As Microsoft.Office.Interop.Excel.Range
        Dim inputPathAndFileName As String
        Dim inputSheetName As String
        Dim inputWorkbook As Excel.Workbook
        Dim errorMessage As String = ""
        Dim ws As Excel.Worksheet

        If Not OnTimeDataWorkSheet.ValidateInputSheet() Then
            MsgBox("Invalid Input Worksheet.")
        Else
            Util.ScreenUpdatesOff()
            ws = Globals.ThisAddIn.Application.ActiveSheet

            For Each rng In CType(Globals.ThisAddIn.Application.Selection, Excel.Range)
                inputPathAndFileName = rng.Value
                inputSheetName = rng.Offset(0, 1).Value

                If inputPathAndFileName <> "" Then
                    OnTimeDataWorkSheet.OpenInputFileAndVerifyDataLayout(inputPathAndFileName, inputSheetName, errorMessage)
                    ws.Unprotect()
                    rng.Offset(0, 2).Value = errorMessage
                    ws.Protect()
                    Globals.ThisAddIn.Application.ActiveWorkbook.Close(False)
                End If
            Next rng

            Util.ScreenUpdatesOn()
        End If

        inputWorkbook = Nothing
    End Sub

    Private Sub btnOpenOnTimeDataFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOpenOnTimeDataFile.Click
        With Globals.ThisAddIn.Application
            Try
                .Workbooks.Open(Filename:=.ActiveCell.Value, ReadOnly:=True)
            Catch ex As Exception
                MessageBox.Show("File not found.  Must select a cell containing a valid FilePath")
            End Try
        End With
    End Sub

    Private Sub btnCreatePage_OnTimeDeliveryData_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'CreateSheet.NewSheet(Globals.cSN_OnTimeDeliveryData)
    End Sub

    Private Sub btnAddOnTimeCharts_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddOnTimeCharts.Click
        Dim rng As Excel.Range
        Dim ws As Excel.Worksheet
        Dim currentSheet As Excel.Worksheet

        Try
            With Globals.ThisAddIn.Application
                currentSheet = .ActiveSheet

                For Each rng In CType(.Selection, Microsoft.Office.Interop.Excel.Range)
                    ' Verify we have a file name and a sheet name

                    If "" <> rng.Value And "" <> rng.Offset(0, Globals.cOTD_DataSheet_Offset).Value Then
                        ' If yes, select the sheet and add the chart

                        ws = .Sheets(rng.Offset(0, Globals.cOTD_DataSheet_Offset).Value)
                        'Charts.AddOnTimeChartToWorksheet(ws)
                        ws = Nothing
                    End If
                Next rng
            End With

            currentSheet.Activate()
        Catch ex As Exception
            MessageBox.Show("Exception: btnAddOnTimeDataToPowerPoint_Click()")
        End Try
    End Sub

    Private Sub btnAddOnTimeChart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddOnTimeChart.Click
        'Charts.AddOnTimeChartToWorksheet(Globals.ThisAddIn.Application.ActiveSheet)
    End Sub
End Class
