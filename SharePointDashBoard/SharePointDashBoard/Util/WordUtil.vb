Imports Excel = Microsoft.Office.Interop.Excel
Imports PowerPoint = Microsoft.Office.Interop.PowerPoint
Imports Word = Microsoft.Office.Interop.Word


Imports Microsoft.Office.Core
Imports PacificLife.Life

Public Class WordUtil

    Public Shared Sub CreateStyle(ByRef wdApp As Word.Application, ByRef document As Word.Document, ByVal styleName As String)
        Select Case styleName.ToUpper
            Case "FEEDBACK"
                CreateStyle_Feedback(wdApp, document)

            Case Else
                MessageBox.Show("Unsupported Style: " & styleName)
        End Select
    End Sub

    Private Shared Sub CreateStyle_Feedback(ByRef wdApp As Word.Application, ByRef document As Word.Document)
        'wdApp.WordBasic.FormatStyle(Name:="Feedback", NewName:="", BasedOn:="", _
        '    NextStyle:="", Type:=0, FileName:="", Link:="")

        Try
            document.Styles("Feedback").Delete()
        Catch ex As Exception
            ' May not exist
        End Try

        document.Styles.Add("Feedback", Word.WdStyleType.wdStyleTypeParagraph)

        With document.Styles("Feedback").Font
            .Name = "Calibri"
            .Size = 8
            .Bold = False
            .Italic = False
            .Underline = Word.WdUnderline.wdUnderlineNone
            .UnderlineColor = Word.WdColor.wdColorAutomatic
            .StrikeThrough = False
            .DoubleStrikeThrough = False
            .Outline = False
            .Emboss = False
            .Shadow = False
            .Hidden = False
            .SmallCaps = False
            .AllCaps = False
            .Color = Word.WdColor.wdColorAutomatic
            .Engrave = False
            .Superscript = False
            .Subscript = False
            .Scaling = 100
            .Kerning = 0
            .Animation = Word.WdAnimation.wdAnimationNone
        End With

        With document.Styles("Feedback").ParagraphFormat
            .LeftIndent = wdApp.InchesToPoints(1)
            .RightIndent = wdApp.InchesToPoints(0)
            .SpaceBefore = 0
            .SpaceBeforeAuto = False
            .SpaceAfter = 10
            .SpaceAfterAuto = False
            .LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple
            .LineSpacing = wdApp.LinesToPoints(1.15)
            .Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
            .WidowControl = True
            .KeepWithNext = False
            .KeepTogether = False
            .PageBreakBefore = False
            .NoLineNumber = False
            .Hyphenation = True
            .FirstLineIndent = wdApp.InchesToPoints(-1)
            .OutlineLevel = Word.WdOutlineLevel.wdOutlineLevelBodyText
            .CharacterUnitLeftIndent = 0
            .CharacterUnitRightIndent = 0
            .CharacterUnitFirstLineIndent = 0
            .LineUnitBefore = 0
            .LineUnitAfter = 0
            .MirrorIndents = False
            .TextboxTightWrap = Word.WdTextboxTightWrap.wdTightNone

            With .Shading
                .Texture = Word.WdTextureIndex.wdTextureNone
                .ForegroundPatternColor = Word.WdColor.wdColorAutomatic
                .BackgroundPatternColor = Word.WdColor.wdColorAutomatic
            End With

            .Borders(Word.WdBorderType.wdBorderLeft).LineStyle = Word.WdLineStyle.wdLineStyleNone
            .Borders(Word.WdBorderType.wdBorderRight).LineStyle = Word.WdLineStyle.wdLineStyleNone
            .Borders(Word.WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleNone
            .Borders(Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleNone

            With .Borders
                .DistanceFromTop = 1
                .DistanceFromLeft = 4
                .DistanceFromBottom = 1
                .DistanceFromRight = 4
                .Shadow = False
            End With
        End With

        With document.Styles("Feedback")
            .NoSpaceBetweenParagraphsOfSameStyle = False
            .ParagraphFormat.TabStops.ClearAll()
            .LanguageID = Word.WdLanguageID.wdEnglishUS
            .NoProofing = False
            .Frame.Delete()
        End With

        'With document.Styles("Feedback").ParagraphFormat
        '    With .Shading
        '        .Texture = Word.WdTextureIndex.wdTextureNone
        '        .ForegroundPatternColor = Word.WdColor.wdColorAutomatic
        '        .BackgroundPatternColor = Word.WdColor.wdColorAutomatic
        '    End With

        '    .Borders(Word.WdBorderType.wdBorderLeft).LineStyle = Word.WdLineStyle.wdLineStyleNone
        '    .Borders(Word.WdBorderType.wdBorderRight).LineStyle = Word.WdLineStyle.wdLineStyleNone
        '    .Borders(Word.WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleNone
        '    .Borders(Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleNone

        '    With .Borders
        '        .DistanceFromTop = 1
        '        .DistanceFromLeft = 4
        '        .DistanceFromBottom = 1
        '        .DistanceFromRight = 4
        '        .Shadow = False
        '    End With
        'End With


    End Sub

End Class
