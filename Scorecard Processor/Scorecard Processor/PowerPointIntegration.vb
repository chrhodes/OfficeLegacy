Option Explicit On

Imports Excel = Microsoft.Office.Interop.Excel
Imports PowerPoint = Microsoft.Office.Interop.PowerPoint
Imports Microsoft.Office.Core
Imports PacificLife.Life

'''
Public Class PowerPointIntegration
    Public Sub ShowShapeInfo()
        PLLog.Trace1("Enter", "Scorecard")

        Dim s As Excel.Shape

        For Each s In Globals.ThisAddIn.Application.ActiveSheet.Shapes
            Debug.Print(s.Name)
            '        Debut.Print s.OLEFormat
        Next s

        PLLog.Trace1("Exit", "Scorecard")
    End Sub

    Public Shared Sub AddSurveyResults(ByVal includeCharts As Boolean, ByVal includeFeedback As Boolean)
        PLLog.Trace1("Enter", "Scorecard")

        If Not includeCharts And Not includeFeedback Then
            Return
        End If

        Dim ppApp As New PowerPoint.Application
        ppApp.Activate()

        'If ppApp Is Nothing Then
        '    MessageBox.Show("Must start PowerPoint before running this command")
        '    Return
        'End If

        '    Set PowerPoint = GetObject("PowerPoint.Application")
        Dim presentation As PowerPoint.Presentation
        Dim slide As PowerPoint.Slide
        Dim shape As PowerPoint.Shape
        Dim textRange As PowerPoint.TextRange
        Dim titleText As String

        Dim wsResults As Excel.Worksheet

        Dim responses As String
        Dim response As String
        Dim question As String
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

        Dim fileName As String

        Dim responseCount As Integer
        Dim hasContinuationSlide As Boolean

        wsResults = Globals.ThisAddIn.Application.ActiveSheet

        Select Case wsResults.Name
            Case Globals.cSN_PartnerSurveyData
            Case Globals.cSN_BusinessSurveyData
            Case Globals.cSN_ITSurveyData
                ' Must be one of these for the following code to work.

            Case Else
                MessageBox.Show("Must be on Partner|Business|IT Survey Data Sheet before running this command")
                Return
        End Select

        ' ToDo: Need to verify PowerPoint is running.  Lame that it won't start itself.

        'ppApp.Presentations.Add(MsoTriState.msoTrue)
        ppApp.Presentations.Open(FileName:="\\ind-svr-02\projdata2\LDM\Templates\8-Other\Scorecard Survey Results template.pot", ReadOnly:=MsoTriState.msoTrue)
        presentation = ppApp.ActivePresentation

        ' Get the data we need off the survey data sheet

        surveyName = wsResults.Range(Globals.cSD_SurveyNameCell).Value
        surveyPeriod = wsResults.Range(Globals.cSD_SurveyPeriodCell).Value
        teamName = wsResults.Range(Globals.cSD_TeamNameCell).Value
        startDataRow = wsResults.Range(Globals.cSD_StartDataRowCell).Value
        endDataRow = wsResults.Range(Globals.cSD_EndDataRowCell).Value
        chartCount = wsResults.Range(Globals.cSD_ChartCountCell).Value
        questionCount = wsResults.Range(Globals.cSD_QuestionCountCell).Value

        fileName = surveyName & "-" & surveyPeriod & " - " & teamName

        ' Update the first slide of the presentation

        ModifyTitle(presentation, fileName, Today.ToShortDateString)

        ' Initialize the column counter to the first column with data (questions)

        currentColumn = Globals.cSD_ResponseValueColumn

        ' Responses are the column after the question if there is a data response
        ' which would be represented with a chart

        responseOffset = 1

        ' Loop through the columns (questions) and output the data (charts & responses)
        ' to PowerPoint

        For questionCounter = 1 To questionCount
            responses = ""

            If questionCounter > chartCount Then
                ' We are past the questions with data responses
                responseOffset = 0
            Else
                If includeCharts Then
                    ' Add a slide to the deck containing a Title only

                    slide = presentation.Slides.Add( _
                        Index:=presentation.Slides.Count + 1, _
                        Layout:=PowerPoint.PpSlideLayout.ppLayoutTitleOnly)

                    ' Update the title to reflect the selected survey

                    shape = slide.Shapes(1)
                    textRange = shape.TextFrame.TextRange
                    titleText = surveyName & " : " & surveyPeriod & " - " & teamName

                    FormatSlideTitle(textRange, titleText)

                    ' Copy the related chart

                    ' .Copy is really slow compared to .CopyPicture
                    ' ws.Shapes.Item(questionCounter).Select()
                    ' ws.Shapes(questionCouter).Copy

                    wsResults.Shapes.Item(questionCounter).CopyPicture()

                    ' And add the chart to the PowerPoint deck

                    'nbrShapes = slide.Shapes.Count
                    slide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPasteEnhancedMetafile)

                    'With slide.Shapes(nbrShapes + 1)

                    ' And position it where it needs to go

                    With slide.Shapes(slide.Shapes.Count)
                        .Top = Globals.cPP_SurveyChartTop
                        .Left = Globals.cPP_SurveyChartLeft
                        .Width = Globals.cPP_SurveyChartWidth
                        .Height = Globals.cPP_SurveyChartHeight
                    End With
                End If
            End If

            ' Now add the questions and question feedback

            If includeFeedback Then

                question = _
                    wsResults.Cells(Globals.cSD_QuestionIDRow, currentColumn).value _
                    & " - " _
                    & wsResults.Cells(Globals.cSD_QuestionTextRow, currentColumn).value
                'Debug.Print(question)

                responseCount = 0
                hasContinuationSlide = False

                For currentRow = startDataRow To endDataRow
                    'Debug.Print(ws.Cells(currentRow, currentColumn + responseOffset).value)
                    response = wsResults.Cells(currentRow, currentColumn + responseOffset).value

                    ' Attribute the feedback to either a person (Partner and Business Survey) 
                    ' or  a team (IT Survey).

                    respondentName = wsResults.Cells(currentRow, Globals.cSD_RespondentName_Column).Value

                    ' TODO: May need to catch some error if don't have a valid team name.

                    If respondentName = "" Then
                        respondentName = wsResults.Cells(currentRow, Globals.cSD_TeamName_Column).Value
                    End If

                    'Debug.Print(questionCounter & " " & currentRow & " >" & respondentName & "< >" & response & "<")

                    ' Select all non-blank responses if All IT is selected as team, else, filter by team.

                    If response <> "" Then
                        If teamName = Globals.cAllITString Then
                            responses = responses & vbCr & respondentName & ":" & vbTab & response
                            responseCount += 1
                        ElseIf teamName = wsResults.Cells(currentRow, Globals.cSD_TeamName_Column).value Then
                            responses = responses & vbCr & respondentName & ":" & vbTab & response
                            responseCount += 1
                        End If
                    End If

                    ' TODO: Pay attention to how many characters we have added to the responses.
                    ' If greater than a certain number add the slide and continue on subsequent slides.  

                    If responses.Length > Globals.cPP_MaxResponseLengthPerPage Then
                        AddSurveyFeedbackToSlide(presentation, question, responses)
                        responses = ""

                        If Not hasContinuationSlide Then
                            question = question & " (cont.)"
                        End If

                        hasContinuationSlide = True
                    End If

                Next currentRow

                ' Add the responses (or last set if have previously added parts)

                AddSurveyFeedbackToSlide(presentation, question, responses)
            End If

            ' Now increment current column to go to the next question.

            currentColumn = currentColumn + 1 + responseOffset
        Next questionCounter

        Dim workbookPath As String = Globals.ThisAddIn.Application.ActiveWorkbook.Path

        ppApp.ActivePresentation.SaveAs(FileName:=workbookPath & "\Team Scorecards\PowerPoint Output\" & fileName.ToString, FileFormat:=PowerPoint.PpSaveAsFileType.ppSaveAsPresentation)
        ppApp.ActivePresentation.Close()

        PLLog.Trace1("Exit", "Scorecard")

    End Sub

    Private Shared Sub AddSurveyFeedbackToSlide( _
        ByRef presentation As PowerPoint.Presentation, _
        ByVal question As String, _
        ByVal responses As String _
    )
        PLLog.Trace1("Enter", "Scorecard")

        Dim slideLayoutStyle As Integer
        Dim slide As PowerPoint.Slide
        Dim shape As PowerPoint.Shape
        Dim shape2 As PowerPoint.Shape
        Dim textRange As PowerPoint.TextRange
        Dim textRange2 As PowerPoint.TextRange

        'slideLayoutStyle = PowerPoint.PpSlideLayout.ppLayoutTitle  ' Two shapes, Title and Sub Title
        'slideLayoutStyle = PowerPoint.PpSlideLayout.ppLayoutTitleOnly  ' One shape, Title
        'slideLayoutStyle = PowerPoint.PpSlideLayout.ppLayoutBlank  ' No shapes

        slideLayoutStyle = PowerPoint.PpSlideLayout.ppLayoutText    ' Two shapes, Title and Text

        ' Create a new slide at end of deck (.Count + 1)

        slide = presentation.Slides.Add(Index:=presentation.Slides.Count + 1, Layout:=slideLayoutStyle)

        shape = slide.Shapes(1)
        textRange = shape.TextFrame.TextRange

        FormatSlideTitle(textRange, question)

        shape2 = slide.Shapes(2)
        shape2.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText
        '        shape2.textFrame.AutoSize = ppAutoSizeNone
        '        shape2.textFrame.AutoSize = ppAutoSizeMixed

        With shape2
            .Top = Globals.cPP_SurveyFeedbackTop
            .Left = Globals.cPP_SurveyFeedbackLeft
            .Height = Globals.cPP_SurveyFeedbackHeight
            .Width = Globals.cPP_SurveyFeedbackWidth
        End With

        textRange2 = shape2.TextFrame.TextRange

        With textRange2
            '            .ParagraphFormat.WordWrap = msoTrue

            .ParagraphFormat.Bullet.Type = PowerPoint.PpBulletType.ppBulletNone

            With .Font
                .Name = "Arial"
                .Size = Globals.cPP_ResponseFontSize
                .Bold = MsoTriState.msoFalse
                .Italic = MsoTriState.msoFalse
                .Underline = MsoTriState.msoFalse
                .Shadow = MsoTriState.msoFalse
                .Emboss = MsoTriState.msoFalse
                .BaselineOffset = 0
                .AutoRotateNumbers = MsoTriState.msoFalse
                .Color.SchemeColor = PowerPoint.PpColorSchemeIndex.ppTitle
            End With

            .Text = responses
            .Characters.Font.Size = Globals.cPP_ResponseFontSize
        End With

        ' TODO: Remove all tab stops

        With shape2.TextFrame.Ruler
            '.TabStops.Add(PowerPoint.PpTabStopType.ppTabStopLeft, 72)
            .Levels(1).FirstMargin = 0
            .Levels(1).LeftMargin = Globals.cPP_ResponseLeftMargin
        End With

        PLLog.Trace1("Exit", "Scorecard")
    End Sub

    Public Shared Sub AddOnTimeDataToPowerPoint(ByVal ws As Excel.Worksheet)
        PLLog.Trace1("Enter", "Scorecard")

        Dim ppApp As New PowerPoint.Application
        '    Set PowerPoint = GetObject("PowerPoint.Application")
        Dim presentation As PowerPoint.Presentation
        Dim slide As PowerPoint.Slide
        Dim shape As PowerPoint.Shape
        Dim textRange As PowerPoint.TextRange
        Dim titleText As String

        Dim teamName As String
        Dim metricName As String
        Dim surveyPeriod As String
        Dim percentage As String
        ' ToDo: Need to verify PowerPoint is running.  Lame that it won't start itself.

        '    Set presentation = PowerPoint.Presentations.Open(Filename:="\\ind-svr-02\projdata2\LDM\Templates\8-Other\Mtg Presentation template.pot", Untitled:=msoTrue)

        ' Get the data we need off the sheet

        teamName = ws.Range(Globals.cOTD_TeamName_Cell).Value
        metricName = ws.Range(Globals.cOTD_MetricName_Cell).Value
        surveyPeriod = ws.Range(Globals.cOTD_SurveyPeriod_Cell).Value
        percentage = Format(ws.Range(Globals.cOTD_OnTimePercentage_Cell).Value, "0%")

        Try
            presentation = ppApp.ActivePresentation

            slide = presentation.Slides.Add(Index:=presentation.Slides.Count + 1, Layout:=PowerPoint.PpSlideLayout.ppLayoutTitleOnly)
            shape = slide.Shapes(1)
            textRange = shape.TextFrame.TextRange
            titleText = metricName & " : " & surveyPeriod & " - " & teamName & " Average: " & percentage

            FormatSlideTitle(textRange, titleText)

            Dim nbrShapes As Integer
            nbrShapes = slide.Shapes.Count
            'Debug.Print(slide.Shapes.Count)

            'ws.Shapes.Item(1).Select()
            'ws.Shapes.Item(1).Copy()
            ws.Shapes.Item(1).CopyPicture()

            slide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPasteEnhancedMetafile)

            With slide.Shapes(nbrShapes + 1)
                .Top = Globals.cPP_OnTimeChartTop
                .Left = Globals.cPP_OnTimeChartLeft
                .Width = Globals.cPP_OnTimeChartWidth
                .Height = Globals.cPP_OnTimeChartHeight
            End With

            'Debug.Print(slide.Shapes.Count)
        Catch ex As Exception
            MessageBox.Show("Exception: AddOnTimeDataToPowerPointIntegration()" & ex.ToString())
        End Try

        PLLog.Trace1("Exit", "Scorecard")
    End Sub

    Private Shared Sub FormatSlideTitle(ByRef textRange As PowerPoint.TextRange, ByVal text As String)
        PLLog.Trace2("Enter", "Scorecard")

        With textRange
            .Text = text

            With .Font
                .Name = "Arial"
                .Size = Globals.cPP_TitleFontSize
                .Bold = MsoTriState.msoFalse
                .Italic = MsoTriState.msoFalse
                .Underline = MsoTriState.msoFalse
                .Shadow = MsoTriState.msoFalse
                .Emboss = MsoTriState.msoFalse
                .BaselineOffset = 0
                .AutoRotateNumbers = MsoTriState.msoFalse
                .Color.SchemeColor = PowerPoint.PpColorSchemeIndex.ppTitle
            End With
        End With

        PLLog.Trace2("Exit", "Scorecard")
    End Sub

    Private Shared Sub FormatSlideResponses(ByRef textRange As PowerPoint.TextRange, ByVal text As String)
        PLLog.Trace2("Enter", "Scorecard")

        With textRange
            .ParagraphFormat.Bullet.Type = PowerPoint.PpBulletType.ppBulletNone

            With .Font
                .Name = "Arial"
                .Size = Globals.cPP_ResponseFontSize
                .Bold = MsoTriState.msoFalse
                .Italic = MsoTriState.msoFalse
                .Underline = MsoTriState.msoFalse
                .Shadow = MsoTriState.msoFalse
                .Emboss = MsoTriState.msoFalse
                .BaselineOffset = 0
                .AutoRotateNumbers = MsoTriState.msoFalse
                .Color.SchemeColor = PowerPoint.PpColorSchemeIndex.ppTitle
            End With

            .Text = text
            ' TODO: Is the next line needed?
            .Characters.Font.Size = Globals.cPP_ResponseFontSize
        End With

        PLLog.Trace2("Exit", "Scorecard")
    End Sub

    Private Shared Sub ModifyTitle(ByVal presentation As PowerPoint.Presentation, ByVal titleText As String, ByVal dateText As String)
        Dim sld As PowerPoint.Slide
        'Dim shp As PowerPoint.Shape

        '    For i = 1 To Application.ActivePresentation.SlideMaster.Shapes.Count
        '        ActivePresentation.SlideMaster
        '          ActivePresentation.SlideMaster.Shapes(i).Select
        '    Next i
        '    For i = 0 To ActiveWindow.Selection.SlideRange.Shapes.Count
        '        ActiveWindow.Selection.SlideRange.Shapes(i).Select
        '    Next i

        'Debug.Print(presentation.Slides(1).Shapes.Count)

        ' Note: The template has been specially prepared so the two shapes on the
        ' first slide have names that can be used to select the shape.  This code
        ' will not work if template changes.

        sld = presentation.Slides(1)
        Try
            sld.Shapes("TITLE").TextFrame.TextRange.Text = titleText
            sld.Shapes("DATE").TextFrame.TextRange.Text = dateText
        Catch ex As Exception
            MessageBox.Show("Could not locate TITLE and DATE shapes on first slide.  Check template")
        End Try

        'For i = 1 To sld.Shapes.Count
        '    shp = sld.Shapes(i)

        '    Debug.Print(">" & sld.Name & "<  >" & shp.Name & "<  " & i & "  >" & shp.TextFrame.TextRange.Text & "<")
        '    ActivePresentation.Slides(1).Shapes(i).Select()
        '    '        shp.Name = "Shape" & i
        'Next i

        'sld.Shapes(1).Name = "DATE"
        'sld.Shapes(2).Name = "TITLE"
    End Sub
End Class
