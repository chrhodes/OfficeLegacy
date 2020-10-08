Option Explicit On

Imports Excel = Microsoft.Office.Interop.Excel
Imports PowerPoint = Microsoft.Office.Interop.PowerPoint
Imports Microsoft.Office.Core
Imports PacificLife.Life

Public Class PowerPointOutput
    Private ppApp As PowerPoint.Application

    Public Sub CreateOutput( _
        ByRef workSheet As Excel.Worksheet, _
        ByVal startingRow As Integer, _
        ByVal endingRow As Integer, _
        ByVal fileName As String _
    )
        'If ppApp Is Nothing Then
        '    ppApp = New PowerPoint.Application
        '    ppApp.Activate()
        'End If

        'Dim presentation As PowerPoint.Presentation
        'Dim slide As PowerPoint.Slide
        'Dim shape As PowerPoint.Shape
        'Dim textRange As PowerPoint.TextRange

        'ppApp.Presentations.Open(FileName:="\\ind-svr-02\projdata2\LDM\Templates\8-Other\Scorecard Survey Results template.pot", ReadOnly:=MsoTriState.msoTrue)
        'presentation = ppApp.ActivePresentation

        'slide = presentation.Slides.Add( _
        '    Index:=presentation.Slides.Count + 1, _
        '    Layout:=PowerPoint.PpSlideLayout.ppLayoutTitleOnly)


        'Dim workbookPath As String = Globals.ThisAddIn.Application.ActiveWorkbook.Path

        'ppApp.ActivePresentation.SaveAs(FileName:=workbookPath & "\Team Scorecards\PowerPoint Output\" & fileName.ToString, FileFormat:=PowerPoint.PpSaveAsFileType.ppSaveAsPresentation)
        'ppApp.ActivePresentation.Close()

    End Sub

    'Private Shared Sub AddSurveyFeedbackToSlide( _
    '    ByRef presentation As PowerPoint.Presentation, _
    '    ByVal question As String, _
    '    ByVal responses As String _
    ')
    '    PLLog.Trace1("Enter", "Scorecard")

    '    Dim slideLayoutStyle As Integer
    '    Dim slide As PowerPoint.Slide
    '    Dim shape As PowerPoint.Shape
    '    Dim shape2 As PowerPoint.Shape
    '    Dim textRange As PowerPoint.TextRange
    '    Dim textRange2 As PowerPoint.TextRange

    '    'slideLayoutStyle = PowerPoint.PpSlideLayout.ppLayoutTitle  ' Two shapes, Title and Sub Title
    '    'slideLayoutStyle = PowerPoint.PpSlideLayout.ppLayoutTitleOnly  ' One shape, Title
    '    'slideLayoutStyle = PowerPoint.PpSlideLayout.ppLayoutBlank  ' No shapes

    '    slideLayoutStyle = PowerPoint.PpSlideLayout.ppLayoutText    ' Two shapes, Title and Text

    '    ' Create a new slide at end of deck (.Count + 1)

    '    slide = presentation.Slides.Add(Index:=presentation.Slides.Count + 1, Layout:=slideLayoutStyle)

    '    shape = slide.Shapes(1)
    '    textRange = shape.TextFrame.TextRange

    '    FormatSlideTitle(textRange, question)

    '    shape2 = slide.Shapes(2)
    '    shape2.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText
    '    '        shape2.textFrame.AutoSize = ppAutoSizeNone
    '    '        shape2.textFrame.AutoSize = ppAutoSizeMixed

    '    With shape2
    '        .Top = Globals.cPP_SurveyFeedbackTop
    '        .Left = Globals.cPP_SurveyFeedbackLeft
    '        .Height = Globals.cPP_SurveyFeedbackHeight
    '        .Width = Globals.cPP_SurveyFeedbackWidth
    '    End With

    '    textRange2 = shape2.TextFrame.TextRange

    '    With textRange2
    '        '            .ParagraphFormat.WordWrap = msoTrue

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

    '        .Text = responses
    '        .Characters.Font.Size = Globals.cPP_ResponseFontSize
    '    End With

    '    ' TODO: Remove all tab stops

    '    With shape2.TextFrame.Ruler
    '        '.TabStops.Add(PowerPoint.PpTabStopType.ppTabStopLeft, 72)
    '        .Levels(1).FirstMargin = 0
    '        .Levels(1).LeftMargin = Globals.cPP_ResponseLeftMargin
    '    End With

    '    PLLog.Trace1("Exit", "Scorecard")
    'End Sub

    'Private Shared Sub FormatSlideTitle(ByRef textRange As PowerPoint.TextRange, ByVal text As String)
    '    PLLog.Trace2("Enter", "Scorecard")

    '    With textRange
    '        .Text = text

    '        With .Font
    '            .Name = "Arial"
    '            .Size = Globals.cPP_TitleFontSize
    '            .Bold = MsoTriState.msoFalse
    '            .Italic = MsoTriState.msoFalse
    '            .Underline = MsoTriState.msoFalse
    '            .Shadow = MsoTriState.msoFalse
    '            .Emboss = MsoTriState.msoFalse
    '            .BaselineOffset = 0
    '            .AutoRotateNumbers = MsoTriState.msoFalse
    '            .Color.SchemeColor = PowerPoint.PpColorSchemeIndex.ppTitle
    '        End With
    '    End With

    '    PLLog.Trace2("Exit", "Scorecard")
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

    'Private Shared Sub ModifyTitle(ByVal presentation As PowerPoint.Presentation, ByVal titleText As String, ByVal dateText As String)
    '    Dim sld As PowerPoint.Slide
    '    'Dim shp As PowerPoint.Shape

    '    '    For i = 1 To Application.ActivePresentation.SlideMaster.Shapes.Count
    '    '        ActivePresentation.SlideMaster
    '    '          ActivePresentation.SlideMaster.Shapes(i).Select
    '    '    Next i
    '    '    For i = 0 To ActiveWindow.Selection.SlideRange.Shapes.Count
    '    '        ActiveWindow.Selection.SlideRange.Shapes(i).Select
    '    '    Next i

    '    'Debug.Print(presentation.Slides(1).Shapes.Count)

    '    ' Note: The template has been specially prepared so the two shapes on the
    '    ' first slide have names that can be used to select the shape.  This code
    '    ' will not work if template changes.

    '    sld = presentation.Slides(1)
    '    Try
    '        sld.Shapes("TITLE").TextFrame.TextRange.Text = titleText
    '        sld.Shapes("DATE").TextFrame.TextRange.Text = dateText
    '    Catch ex As Exception
    '        MessageBox.Show("Could not locate TITLE and DATE shapes on first slide.  Check template")
    '    End Try

    '    'For i = 1 To sld.Shapes.Count
    '    '    shp = sld.Shapes(i)

    '    '    Debug.Print(">" & sld.Name & "<  >" & shp.Name & "<  " & i & "  >" & shp.TextFrame.TextRange.Text & "<")
    '    '    ActivePresentation.Slides(1).Shapes(i).Select()
    '    '    '        shp.Name = "Shape" & i
    '    'Next i

    '    'sld.Shapes(1).Name = "DATE"
    '    'sld.Shapes(2).Name = "TITLE"
    'End Sub
End Class
