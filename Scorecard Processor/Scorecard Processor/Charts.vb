Option Explicit On

Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Core
Imports PacificLife.Life

Public Class Charts

    ' TODO: This routine doesn't do much other than say something is not
    ' implemented.

    Public Shared Sub AddCharts()
        PLLog.Trace1("Enter", "Scorecard")

        Dim sheetName As String

        sheetName = Globals.ThisAddIn.Application.ActiveSheet.Name

        DeleteChartsFromActiveWorkSheet()

        Select Case sheetName
            Case "Partner Survey Data"
                AddSurveyChartsToWorkSheet(sheetName)

            Case "Business Survey Data"
                AddSurveyChartsToWorkSheet(sheetName)

            Case "IT Survey Data"
                AddSurveyChartsToWorkSheet(sheetName)

            Case "Help Desk Survey Data"
                AddSurveyChartsToWorkSheet(sheetName)

            Case Else
                MsgBox("AddCharts to <" & sheetName & ">", vbOKOnly, "Not Implemented")

        End Select

        PLLog.Trace1("Exit", "Scorecard")
    End Sub

    Private Shared Sub AddSurveyChartsToWorkSheet(ByVal sheetName As String)
        PLLog.Trace1("Enter", "Scorecard")

        Dim chartCount As Integer
        Dim i As Double
        Dim labelRange As Excel.Range
        Dim valueRange As Excel.Range
        Dim chartDataRange As Excel.Range
        Dim chartTitle As String
        Dim question As String
        Dim questionText As String
        Dim average As String
        Dim stdevp As String
        Dim responseCount As Integer
        Dim ws As Excel.Worksheet
        Dim currentWorkbookName As String
        'Dim newChart As Excel.Shape
        Dim newShape As Excel.Shape
        Dim newChart As Excel.Chart

        currentWorkbookName = Globals.ThisAddIn.Application.ActiveWorkbook.Name
        ws = Globals.ThisAddIn.Application.Sheets.Item(sheetName)
        Util.ScreenUpdatesOff()

        With ws
            chartCount = .Range(Globals.cSD_ChartCountCell).Value

            ' The labels appear in a fixed location on the spreadsheet.  Create a range object that will be joined (Union) with the data values below

            labelRange = .Range(.Cells(Globals.cSD_ResponseLabelRowStart, Globals.cSD_ResponseLabelColumn), .Cells(Globals.cSD_ResponseLabelRowEnd, Globals.cSD_ResponseLabelColumn))

            For i = 0 To chartCount - 1
                ' Each question associated with a chart has a follow-up question.  So, move over by two (i * 2)
                ' The questions (and associated chart) start at cValueColumn.  So, ( + cValueColumn)

                ' Create a range object that holds the data values for the current questions then create a chart range object that
                ' holds the labels (above) and the data values.
                valueRange = .Range(.Cells(Globals.cSD_ResponseValueRowStart, (i * 2) + Globals.cSD_ResponseValueColumn), .Cells(Globals.cSD_ResponseValueRowEnd, (i * 2) + Globals.cSD_ResponseValueColumn))
                chartDataRange = .Application.Union(labelRange, valueRange)
                'Util.DisplayExcelRange(chartDataRange)

                ' Gather information that will be placed onto the produced chart.

                responseCount = .Cells(Globals.cSD_ResponseCountRow, (i * 2) + Globals.cSD_ResponseValueColumn).value
                question = .Cells(Globals.cSD_QuestionIDRow, (i * 2) + Globals.cSD_ResponseValueColumn).value
                questionText = .Cells(Globals.cSD_QuestionTextRow, (i * 2) + Globals.cSD_ResponseValueColumn).value

                ' ToDo: Fix hack.  Need to handle bad Average and Stdevp
                On Error Resume Next
                ' ToDo: Remove ChangeTo5To1() code.  For now just don't call.

                average = Format(.Cells(Globals.cSD_OverallAverageRow, (i * 2) + Globals.cSD_ResponseValueColumn).value, "0.0")
                'average = Format(ChangeTo5To1(.Cells(Globals.cSD_OverallAverageRow, (i * 2) + Globals.cSD_ResponseValueColumn).value), "0.0")

                stdevp = Format(.Cells(Globals.cSD_AverageDeviationRow, (i * 2) + Globals.cSD_ResponseValueColumn).value, "0.0")

                chartTitle = question & " - " & questionText & vbCr _
                    & " (" & average & " +/- " & stdevp & ")"

                .Cells(1, (i * 2) + Globals.cSD_ResponseValueColumn).Activate()

                ' Hum, this code worked fine before.  Now Excel 2007 does not seem to like adding charts and then moving
                ' them to the sheet.

                'Globals.ThisAddIn.Application.Charts.Add()
                'Globals.ThisAddIn.Application.ActiveChart.ChartType = Excel.XlChartType.xlColumnClustered
                'Globals.ThisAddIn.Application.ActiveChart.SetSourceData(Source:=chartDataRange, PlotBy:=Excel.XlRowCol.xlColumns)
                'Globals.ThisAddIn.Application.ActiveChart.Location(Where:=Excel.XlChartLocation.xlLocationAsObject, Name:=sheetName)

                newShape = ws.Shapes.AddChart(Excel.XlChartType.xlColumnClustered)
                newChart = newShape.Chart

                With newChart
                    .SetSourceData(Source:=chartDataRange, PlotBy:=Excel.XlRowCol.xlColumns)
                    .Location(Where:=Excel.XlChartLocation.xlLocationAsObject, Name:=sheetName)
                End With

                ' TODO: Clean up code below to use newChart variable above

                'With Globals.ThisAddIn.Application.ActiveChart
                With newChart
                    .HasTitle = True
                    .HasLegend = False
                    .ChartTitle.Characters.Text = chartTitle
                    .Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary).HasTitle = False

                    With .Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary)
                        .HasTitle = True
                        .AxisTitle.Text = Globals.cCH_SurveyValueAxisLabel
                    End With

                    .SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementPrimaryValueAxisTitleRotated)
                End With

                'With Globals.ThisAddIn.Application.ActiveChart.Axes(Excel.XlAxisType.xlValue)
                'With newChart.Axes(Excel.XlAxisType.xlValue)
                '    .MinimumScale = 0
                '    .MaximumScale = responseCount
                '    .MinorUnitIsAuto = True
                '    .MinorUnit = 1.0
                '    .MajorUnitIsAuto = True
                '    .Crosses = Excel.Constants.xlAutomatic
                '    .ReversePlotOrder = False
                '    .ScaleType = Excel.XlTrendlineType.xlLinear
                '    .DisplayUnit = Excel.Constants.xlNone
                'End With
                With newChart.Axes(Excel.XlAxisType.xlValue)
                    .MinimumScale = 0
                    .MaximumScaleIsAuto = True
                    .MinorUnitIsAuto = True
                    .MinorUnit = 1.0
                    .MajorUnitIsAuto = True
                    .Crosses = Excel.Constants.xlAutomatic
                    .ReversePlotOrder = False
                    .ScaleType = Excel.XlTrendlineType.xlLinear
                    .DisplayUnit = Excel.Constants.xlNone
                End With

                newChart.ApplyDataLabels( _
                    Excel.XlDataLabelsType.xlDataLabelsShowValue, _
                    LegendKey:=False, _
                    AutoText:=True, _
                    HasLeaderLines:=False, _
                    ShowSeriesName:=False, _
                    ShowCategoryName:=False, _
                    ShowValue:=True, _
                    ShowPercentage:=False, _
                    ShowBubbleSize:=False)

                ' For some reason cannot move the chart itself.  Need to move the shapes.  See below.
                '        ActiveChart.ChartArea.Left = (i * 90) + 90
                '        ActiveChart.ChartArea.Top = 5

                ' Currenty the chart is active.
                ' ReActivate the underlying worksheet so we can select cells again

                Globals.ThisAddIn.Application.Windows(currentWorkbookName).Activate()

                With Globals.ThisAddIn.Application.ActiveSheet.Shapes(i + 1)
                    .Top = 1
                    .Left = (i * Globals.cSD_SurveyChartSpacing) + Globals.cSD_SurveyChartStartingOffset
                    .Height = Globals.cSD_SurveyChartHeight
                    .Width = Globals.cSD_SurveyChartWidth
                    .ScaleWidth(1.0#, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoScaleFrom.msoScaleFromTopLeft)
                End With

                Globals.ThisAddIn.Application.ActiveSheet.ChartObjects(i + 1).Activate()

                Globals.ThisAddIn.Application.ActiveChart.ChartTitle.Select()
                Globals.ThisAddIn.Application.Selection.AutoScaleFont = True

                With Globals.ThisAddIn.Application.Selection.Font
                    .Name = "Arial"
                    .FontStyle = "Bold"
                    .Size = 12
                    .Strikethrough = False
                    .Superscript = False
                    .Subscript = False
                    .OutlineFont = False
                    .Shadow = False
                    .Underline = Excel.XlUnderlineStyle.xlUnderlineStyleNone
                    .ColorIndex = Excel.Constants.xlAutomatic
                    .Background = Excel.Constants.xlAutomatic
                End With

                Globals.ThisAddIn.Application.ActiveChart.Axes(Excel.XlAxisType.xlValue).Select()
                Globals.ThisAddIn.Application.Selection.TickLabels.AutoScaleFont = True

                With Globals.ThisAddIn.Application.Selection.TickLabels.Font
                    .Name = "Arial"
                    .FontStyle = "Regular"
                    .Size = Globals.cCH_TickLabelFontSize
                    .Strikethrough = False
                    .Superscript = False
                    .Subscript = False
                    .OutlineFont = False
                    .Shadow = False
                    .Underline = Excel.XlUnderlineStyle.xlUnderlineStyleNone
                    .ColorIndex = Excel.Constants.xlAutomatic
                    .Background = Excel.Constants.xlAutomatic
                End With

                With Globals.ThisAddIn.Application.ActiveChart.SeriesCollection(1)
                    .ApplyDataLabels()
                    .DataLabels.Position = Excel.XlDataLabelPosition.xlLabelPositionOutsideEnd
                End With

                ' TODO: Clean up this code.   Looks like a Macro Recorder wrote it :)
                Globals.ThisAddIn.Application.ActiveChart.SeriesCollection(1).DataLabels.Select()
                Globals.ThisAddIn.Application.Selection.AutoScaleFont = True

                With Globals.ThisAddIn.Application.Selection.Font
                    .Name = "Arial"
                    .FontStyle = "Regular"
                    .Size = Globals.cCH_DataLabelFontSize
                    .Strikethrough = False
                    .Superscript = False
                    .Subscript = False
                    .OutlineFont = False
                    .Shadow = False
                    .Underline = Excel.XlUnderlineStyle.xlUnderlineStyleNone
                    .ColorIndex = Excel.Constants.xlAutomatic
                    .Background = Excel.Constants.xlAutomatic
                End With

                Globals.ThisAddIn.Application.ActiveChart.Axes(Excel.XlAxisType.xlCategory).Select()
                Globals.ThisAddIn.Application.Selection.TickLabels.AutoScaleFont = True

                With Globals.ThisAddIn.Application.Selection.TickLabels.Font
                    .Name = "Arial"
                    .FontStyle = "Regular"
                    .Size = Globals.cCH_TickLabelFontSize
                    .Strikethrough = False
                    .Superscript = False
                    .Subscript = False
                    .OutlineFont = False
                    .Shadow = False
                    .Underline = Excel.XlUnderlineStyle.xlUnderlineStyleNone
                    .ColorIndex = Excel.Constants.xlAutomatic
                    .Background = Excel.Constants.xlAutomatic
                End With

                'done:
                valueRange = Nothing
                chartDataRange = Nothing

                ' Currenty the chart is active.
                ' ReActivate the underlying worksheet so we can select cells again

                Globals.ThisAddIn.Application.Windows(currentWorkbookName).Activate()
            Next i

            .Range("A1").Activate()
        End With

        Util.ScreenUpdatesOn()

        PLLog.Trace1("Exit", "Scorecard")
    End Sub

    Public Shared Sub AddOnTimeChartToWorksheet(ByRef ws As Excel.Worksheet)
        PLLog.Trace1("Enter", "Scorecard")

        'Dim ws As Excel.Worksheet
        Dim chartDataRange As Excel.Range
        Dim sheetName As String
        Dim chartTitle As String

        Try
            Charts.DeleteChartsFromActiveWorkSheet()

            With Globals.ThisAddIn.Application
                'With ws
                ws.Activate()
                sheetName = ws.Name
                'ws = .ActiveSheet

                chartDataRange = ws.Range( _
                    ws.Cells( _
                        ws.Range(Globals.cOTD_StartDataRow_Cell).Value, _
                        ws.Range(Globals.cOTD_StartDataColumn_Cell).Value), _
                    ws.Cells( _
                        ws.Range(Globals.cOTD_EndDataRow_Cell).Value, _
                        ws.Range(Globals.cOTD_EndDataColumn_Cell).Value))

                chartTitle = ws.Name

                chartDataRange.Select()
                .Charts.Add()
                .ActiveChart.ChartType = Excel.XlChartType.xlColumnClustered
                .ActiveChart.SetSourceData(Source:=chartDataRange, PlotBy:=Excel.XlRowCol.xlColumns)
                .ActiveChart.SeriesCollection(1).Name = "Series1"
                .ActiveChart.Location(Where:=Excel.XlChartLocation.xlLocationAsObject, Name:=sheetName)

                With .ActiveChart
                    .HasTitle = True
                    .HasLegend = False
                    .ChartTitle.Characters.Text = chartTitle
                    .Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary).HasTitle = True
                    .Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary).AxisTitle.Characters.Text = _
                        "Releases (" & ws.Range(Globals.cOTD_NumberReleases_Cell).Value & ")"

                    With .Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary)
                        .HasTitle = True
                        .AxisTitle.Text = Globals.cCH_OnTimeDeliveryValueAxisLabel
                        .MinimumScale = 0   ' 0%
                        .MaximumScaleIsAuto = True
                        .MajorUnit = 0.2    ' 20%
                    End With

                    .SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementPrimaryValueAxisTitleRotated)
                End With

                With ws.Shapes.Item(1)
                    .Top = Globals.cOTD_ChartTop
                    .Left = Globals.cOTD_ChartLeft
                End With
            End With
        Catch ex As Exception
            MessageBox.Show("Exception: AddOnTimeChart()")
        End Try

        PLLog.Trace1("Exit", "Scorecard")
    End Sub

    Public Shared Sub DeleteChartsFromActiveWorkSheet()
        PLLog.Trace1("Enter", "Scorecard")

        Dim chartCount As Integer

        With Globals.ThisAddIn.Application
            For chartCount = .ActiveSheet.Shapes.Count To 1 Step -1
                .ActiveSheet.Shapes(chartCount).Delete()
            Next chartCount
        End With

        PLLog.Trace1("Exit", "Scorecard")
    End Sub

    Private Shared Function ChangeTo5To1(ByVal value As Single) As Single
        ChangeTo5To1 = 6.0# - value
    End Function

    ' Update the title of the chart to reflect the current values
    ' for Average and Standard Deviation.  They change whenever a new team
    ' is selected.  This saves us from having to recreate the charts.

    Public Shared Sub UpdateChartTitles(ByVal sheetName As String)
        PLLog.Trace1("Enter", "Scorecard")

        Dim chartCount As Integer
        Dim i As Double
        Dim chartTitle As String
        Dim question As String
        Dim questionText As String
        Dim average As String
        Dim stdevp As String
        Dim ws As Excel.Worksheet
        Dim currentWorkbookName As String
        'Dim newChart As Excel.Shape
        Dim charts As Excel.ChartObjects
        Dim chartObj As Excel.ChartObject

        currentWorkbookName = Globals.ThisAddIn.Application.ActiveWorkbook.Name
        ws = Globals.ThisAddIn.Application.Sheets.Item(sheetName)
        Util.ScreenUpdatesOff()

        With ws
            charts = CType(ws.ChartObjects, Excel.ChartObjects)
            chartCount = .Range(Globals.cSD_ChartCountCell).Value

            '' The labels appear in a fixed location on the spreadsheet.  Create a range object that will be joined (Union) with the data values below

            'labelRange = .Range(.Cells(Globals.cSD_ResponseLabelRowStart, Globals.cSD_ResponseLabelColumn), .Cells(Globals.cSD_ResponseLabelRowEnd, Globals.cSD_ResponseLabelColumn))

            For i = 0 To chartCount - 1
                'For Each chartObj As Excel.ChartObject In charts

                question = .Cells(Globals.cSD_QuestionIDRow, (i * 2) + Globals.cSD_ResponseValueColumn).value
                questionText = .Cells(Globals.cSD_QuestionTextRow, (i * 2) + Globals.cSD_ResponseValueColumn).value

                average = Format(.Cells(Globals.cSD_OverallAverageRow, (i * 2) + Globals.cSD_ResponseValueColumn).value, "0.0")
                'average = Format(ChangeTo5To1(.Cells(Globals.cSD_OverallAverageRow, (i * 2) + Globals.cSD_ResponseValueColumn).value), "0.0")

                stdevp = Format(.Cells(Globals.cSD_AverageDeviationRow, (i * 2) + Globals.cSD_ResponseValueColumn).value, "0.0")

                chartTitle = question & " - " & questionText & vbCr _
                    & " (" & average & " +/- " & stdevp & ")"

                ' Item indexes start at 1.
                chartObj = charts.Item(i + 1)
                chartObj.Chart.ChartTitle.Characters.Text = chartTitle
            Next


        End With

        Util.ScreenUpdatesOn()

        PLLog.Trace1("Exit", "Scorecard")
    End Sub
End Class
