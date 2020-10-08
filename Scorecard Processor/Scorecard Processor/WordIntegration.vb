Option Explicit On

Imports Excel = Microsoft.Office.Interop.Excel
Imports PowerPoint = Microsoft.Office.Interop.PowerPoint
Imports Word = Microsoft.Office.Interop.Word


Imports Microsoft.Office.Core
Imports PacificLife.Life

Imports System.Text

Public Class WordIntegration
    Private wdApp As Word.Application

    Public Sub AddSurveyResults(ByVal includeCharts As Boolean, ByVal includeFeedback As Boolean)
        PLLog.Trace1("Enter", "Scorecard")

        If Not includeCharts And Not includeFeedback Then
            Return
        End If

        If wdApp Is Nothing Then
            wdApp = New Word.Application
        End If

        wdApp.Visible = True

        wdApp.Documents.Add(Template:="Normal", NewTemplate:=False, DocumentType:=Word.WdDocumentType.wdTypeDocument, Visible:=True)

        Dim presentationWd As Word.Document
        Dim selectionWd As Word.Selection

        selectionWd = wdApp.Selection
        Dim textRangeWd As Word.Range

        Dim titleText As String
        Dim fileName As String

        Dim ws As Excel.Worksheet

        Dim responsesSb As StringBuilder = New StringBuilder(2000)

        Dim response As String
        Dim question As String = ""
        Dim chartCount As Integer
        Dim questionCount As Integer

        Dim questionCounter As Integer
        Dim teamName As String
        Dim surveyName As String
        Dim surveyPeriod As String
        Dim currentColumn As Integer
        Dim currentRow As Integer

        Dim startDataRow As Integer
        Dim endDataRow As Integer
        Dim responseOffset As Integer
        Dim respondentName As String

        Dim responseCount As Integer
        Dim hasContinuationSlide As Boolean

        ws = Globals.ThisAddIn.Application.ActiveSheet
        presentationWd = wdApp.ActiveDocument
        textRangeWd = wdApp.Selection.Range

        Styles.CreateStyle(wdApp, presentationWd, "Feedback")

        ' Get the data we need off the survey data sheet

        surveyName = ws.Range(Globals.cSD_SurveyNameCell).Value
        surveyPeriod = ws.Range(Globals.cSD_SurveyPeriodCell).Value
        teamName = ws.Range(Globals.cSD_TeamNameCell).Value
        startDataRow = ws.Range(Globals.cSD_StartDataRowCell).Value
        endDataRow = ws.Range(Globals.cSD_EndDataRowCell).Value
        chartCount = ws.Range(Globals.cSD_ChartCountCell).Value
        questionCount = ws.Range(Globals.cSD_QuestionCountCell).Value

        fileName = surveyName & "-" & surveyPeriod & " - " & teamName
        titleText = surveyName & " : " & surveyPeriod & " - " & teamName

        selectionWd.InsertAfter(titleText)
        selectionWd.Style = presentationWd.Styles("Heading 1")
        selectionWd.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
        selectionWd.InsertParagraph()
        selectionWd.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

        ' Initialize the column counter to the first column with data (questions)

        currentColumn = Globals.cSD_ResponseValueColumn

        ' Responses are the column after the question if there is a data response
        ' which would be represented with a chart

        responseOffset = 1

        ' Loop through the columns (questions) and output the data (charts & responses)
        ' to PowerPoint

        For questionCounter = 1 To questionCount
            'responses = ""
            responsesSb.Length = 0

            If includeFeedback Then
                question = _
                    ws.Cells(Globals.cSD_QuestionIDRow, currentColumn).value _
                    & " - " _
                    & ws.Cells(Globals.cSD_QuestionTextRow, currentColumn).value

                selectionWd.InsertAfter(question)
                selectionWd.Style = presentationWd.Styles("Heading 2")
                selectionWd.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                selectionWd.InsertParagraph()
                selectionWd.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
            End If

            If questionCounter > chartCount Then
                ' We are past the questions with data responses
                responseOffset = 0
            Else
                ' Copy the related chart

                If includeCharts Then
                    ws.Shapes.Item(questionCounter).CopyPicture()

                    ' Add the chart to the Word document

                    selectionWd.PasteAndFormat(Word.WdRecoveryType.wdChartPicture)
                    selectionWd.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                    selectionWd.InsertParagraph()
                    selectionWd.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                End If
            End If

            If includeFeedback Then
                ' Now add the question feedback

                responseCount = 0
                hasContinuationSlide = False

                For currentRow = startDataRow To endDataRow
                    'Debug.Print(ws.Cells(currentRow, currentColumn + responseOffset).value)
                    response = ws.Cells(currentRow, currentColumn + responseOffset).value

                    ' Attribute the feedback to either a person (Partner and Business Survey) 
                    ' or  a team (IT Survey).

                    respondentName = ws.Cells(currentRow, Globals.cSD_RespondentName_Column).Value

                    ' TODO: May need to catch some error if don't have a valid team name.

                    If respondentName = "" Then
                        respondentName = ws.Cells(currentRow, Globals.cSD_TeamName_Column).Value
                    End If

                    Debug.Print(questionCounter & " " & currentRow & " >" & respondentName & "< >" & response & "<")

                    ' Select all non-blank responses if All IT is selected as team, else, filter by team.

                    If response <> "" Then
                        If teamName = Globals.cAllITString Then
                            responsesSb.Append(vbCrLf & respondentName & ":" & vbTab & response)
                            responseCount += 1
                        ElseIf teamName = ws.Cells(currentRow, Globals.cSD_TeamName_Column).value Then
                            responsesSb.Append(vbCrLf & respondentName & ":" & vbTab & response)
                            responseCount += 1
                        End If
                    End If

                Next currentRow

                ' Add the responses

                AddSurveyFeedback(selectionWd, question, responsesSb.ToString)

            End If

            ' Now increment current column to go to the next question.

            currentColumn = currentColumn + 1 + responseOffset
        Next questionCounter

        'wdApp.Dialogs(Word.WdWordDialog.wdDialogFileSaveAs).Display()
        Dim workbookPath As String = Globals.ThisAddIn.Application.ActiveWorkbook.Path

        wdApp.ActiveDocument.SaveAs(FileName:=workbookPath & "\Team Scorecards\Word Output\" & fileName.ToString, FileFormat:=Word.WdSaveFormat.wdFormatDocument)
        wdApp.ActiveDocument.Close(SaveChanges:=True)
        'wdApp.Quit(SaveChanges:=False)

        PLLog.Trace1("Exit", "Scorecard")
    End Sub

    Private Sub AddSurveyFeedback( _
            ByRef selectionWd As Word.Selection, _
            ByVal question As String, _
            ByVal responses As String _
        )
        PLLog.Trace1("Enter", "Scorecard")

        selectionWd.InsertAfter(responses)
        selectionWd.Style = selectionWd.Document.Styles("Feedback")
        selectionWd.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
        selectionWd.InsertParagraph()
        selectionWd.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

        selectionWd.InsertBreak(Word.WdBreakType.wdPageBreak)

        PLLog.Trace1("Exit", "Scorecard")
    End Sub

    'Private Shared Sub AddOnTimeDataToPowerPoint(ByVal ws As Excel.Worksheet)
    '    PLLog.Trace1("Enter", "Scorecard")

    '    Dim ppApp As New PowerPoint.Application
    '    '    Set PowerPoint = GetObject("PowerPoint.Application")
    '    Dim presentation As PowerPoint.Presentation
    '    Dim slide As PowerPoint.Slide
    '    Dim shape As PowerPoint.Shape
    '    Dim textRange As PowerPoint.TextRange
    '    Dim titleText As String

    '    Dim teamName As String
    '    Dim metricName As String
    '    Dim surveyPeriod As String
    '    Dim percentage As String
    '    ' ToDo: Need to verify PowerPoint is running.  Lame that it won't start itself.

    '    '    Set presentation = PowerPoint.Presentations.Open(Filename:="\\ind-svr-02\projdata2\LDM\Templates\8-Other\Mtg Presentation template.pot", Untitled:=msoTrue)

    '    ' Get the data we need off the sheet

    '    teamName = ws.Range(Globals.cOTD_TeamName_Cell).Value
    '    metricName = ws.Range(Globals.cOTD_MetricName_Cell).Value
    '    surveyPeriod = ws.Range(Globals.cOTD_SurveyPeriod_Cell).Value
    '    percentage = Format(ws.Range(Globals.cOTD_OnTimePercentage_Cell).Value, "0%")

    '    Try
    '        presentation = ppApp.ActivePresentation

    '        slide = presentation.Slides.Add(Index:=presentation.Slides.Count + 1, Layout:=PowerPoint.PpSlideLayout.ppLayoutTitleOnly)
    '        shape = slide.Shapes(1)
    '        textRange = shape.TextFrame.TextRange
    '        titleText = metricName & " : " & surveyPeriod & " - " & teamName & " Average: " & percentage

    '        FormatSlideTitle(textRange, titleText)

    '        Dim nbrShapes As Integer
    '        nbrShapes = slide.Shapes.Count
    '        'Debug.Print(slide.Shapes.Count)

    '        'ws.Shapes.Item(1).Select()
    '        'ws.Shapes.Item(1).Copy()
    '        ws.Shapes.Item(1).CopyPicture()

    '        slide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPasteEnhancedMetafile)

    '        With slide.Shapes(nbrShapes + 1)
    '            .Top = Globals.cPP_OnTimeChartTop
    '            .Left = Globals.cPP_OnTimeChartLeft
    '            .Width = Globals.cPP_OnTimeChartWidth
    '            .Height = Globals.cPP_OnTimeChartHeight
    '        End With

    '        'Debug.Print(slide.Shapes.Count)
    '    Catch ex As Exception
    '        MessageBox.Show("Exception: AddOnTimeDataToPowerPointIntegration()" & ex.ToString())
    '    End Try

    '    PLLog.Trace1("Exit", "Scorecard")
    'End Sub

    'Private Shared Sub FormatSlideResponses(ByRef textRange As PowerPoint.TextRange, ByVal text As String)
    '    PLLog.Trace2("Enter", "Scorecard")

    '    With textRange
    '        .ParagraphFormat.Bullet.Type = PowerPoint.PpBulletType.ppBulletNone

    '        With .Font
    '            .Name = "Arial"
    '            .Size = Globals.cPP_ResponseFontSize
    '            .Bold = MsoTriState.msoFalse
    '            .Italic = MsoTriState.msoFalse
    '            .Underline = MsoTriState.msoFalse
    '            .Shadow = MsoTriState.msoFalse
    '            .Emboss = MsoTriState.msoFalse
    '            .BaselineOffset = 0
    '            .AutoRotateNumbers = MsoTriState.msoFalse
    '            .Color.SchemeColor = PowerPoint.PpColorSchemeIndex.ppTitle
    '        End With

    '        .Text = text
    '        ' TODO: Is the next line needed?
    '        .Characters.Font.Size = Globals.cPP_ResponseFontSize
    '    End With

    '    PLLog.Trace2("Exit", "Scorecard")
    'End Sub
End Class
