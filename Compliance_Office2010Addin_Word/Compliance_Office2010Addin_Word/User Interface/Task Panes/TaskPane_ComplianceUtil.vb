Imports System.Collections
Imports System.IO
Imports System.Xml
Imports System.Linq
Imports System.Windows.Forms

Imports Microsoft.Office.Interop.Word

Public Class TaskPane_ComplianceUtil
    Private Const cINDEX_WORD_STYLE As String = "IndexWord"
    Private Const cINDEX_HEADING_STYLE As String = "IndexHeading"

#Region "Event Handlers"
    Private Sub btnBuildReplacementWords_Click(sender As System.Object, e As System.EventArgs) Handles btnSaveReplacementWords.Click
        SaveReplacementWordsToXMLFile()
    End Sub

    Private Sub btnCreateIndex_Click(sender As System.Object, e As System.EventArgs) Handles btnCreateIndex.Click
        CreateIndex()
    End Sub

    Private Sub btnCreateIndexStyle_Click(sender As System.Object, e As System.EventArgs) Handles btnCreateIndexStyles.Click
        CreateIndexStyles()
    End Sub

    Private Sub btnFindIndexWords_Click(sender As System.Object, e As System.EventArgs) Handles btnFindIndexWords.Click
        FindIndexWords()
    End Sub

    Private Sub btnLoadReplacementWords_Click(sender As System.Object, e As System.EventArgs) Handles btnLoadReplacementWords.Click
        LoadWordsFromXMLFile("ReplacementWords.xml")
    End Sub

    Private Sub btnMarkIndexWords_Click(sender As System.Object, e As System.EventArgs) Handles btnMarkIndexWords.Click
        MarkIndexWords()
    End Sub

    Private Sub btnTagIndexWords_Click(sender As System.Object, e As System.EventArgs) Handles btnTagIndexWords.Click
        TagIndexWords()
    End Sub

    Private Sub btnZapReplacementWords_Click(sender As System.Object, e As System.EventArgs) Handles btnZapReplacementWords.Click
        ZapReplacementWords()
    End Sub
#End Region

#Region "Main Function Routines"

    Private Sub ApplyStyleToWords(ByVal indexWord As String, ByVal style As String)
        Common.WriteToDebugWindow(String.Format("ApplyStyleToWords:{0} style:{1}", indexWord, style))


        With Globals.ThisAddIn.Application
            .Selection.HomeKey(Unit:=WdUnits.wdStory) 
            .Selection.Find.ClearFormatting

            '.Selection.Find.Replacement.ClearFormatting

            With .Selection.Find
                .Text = indexWord
                .Forward = True
                .Wrap = WdFindWrap.wdFindContinue
                .Format = True
                .MatchCase = False
                .MatchWholeWord = True
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With

            .Selection.Find.Execute ()

            While (.Selection.Find.Found)
                Dim match As Range = .Selection.Range

                Common.WriteToDebugWindow(String.Format("start:{0} end:{1} text:{2} style:{3}",
                                                        match.Start, match.End, match.Text, match.Style.ToString()))

                .Selection.Range.Style = style

                .Selection.Find.Execute()
            End While
        End With
    End Sub

    Private Sub CreateIndex()
        With Globals.ThisAddIn.Application.ActiveDocument.Indexes(1)
            .HeadingSeparator = WdHeadingSeparator.wdHeadingSeparatorBlankLine
            .Type = WdIndexType.wdIndexIndent
            .RightAlignPageNumbers = False
            .NumberOfColumns = 2
            .IndexLanguage = WdLanguageID.wdEnglishUS
            .TabLeader = WdTabLeader.wdTabLeaderDots
        End With
    End Sub

    Private Sub CreateStyle_IndexHeading()
         With Globals.ThisAddIn.Application
            .ActiveDocument.Styles.Add(Name:=cINDEX_HEADING_STYLE, Type:=WdStyleType.wdStyleTypeParagraph)
            .ActiveDocument.Styles(cINDEX_HEADING_STYLE).AutomaticallyUpdate = False

            .ActiveDocument.Styles(cINDEX_HEADING_STYLE).QuickStyle = True

            With .ActiveDocument.Styles(cINDEX_HEADING_STYLE).Font
                .Name = "+Body"
                .Size = 12
                .Bold = True
                .Italic = False
                .Underline = WdUnderline.wdUnderlineNone
                .UnderlineColor = WdColor.wdColorAutomatic
                .StrikeThrough = False
                .DoubleStrikeThrough = False
                .Outline = False
                .Emboss = False
                .Shadow = False
                .Hidden = False
                .SmallCaps = False
                .AllCaps = False
                .Color = 12611584
                .Engrave = False
                .Superscript = False
                .Subscript = False
                .Scaling = 100
                .Kerning = 0
                .Animation = WdAnimation.wdAnimationNone
                .Ligatures = WdLigatures.wdLigaturesNone
                .NumberSpacing = WdNumberSpacing.wdNumberSpacingDefault
                .NumberForm = WdNumberForm.wdNumberFormDefault
                .StylisticSet = WdStylisticSet.wdStylisticSetDefault
                .ContextualAlternates = 0
            End With

            With .ActiveDocument.Styles(cINDEX_HEADING_STYLE).ParagraphFormat
                .LeftIndent =  Globals.ThisAddIn.Application.InchesToPoints(0)
                .RightIndent = Globals.ThisAddIn.Application.InchesToPoints(0)
                .SpaceBefore = 0
                .SpaceBeforeAuto = False
                .SpaceAfter = 6
                .SpaceAfterAuto = False
                .LineSpacingRule = WdLineSpacing.wdLineSpaceMultiple
                .LineSpacing = Globals.ThisAddIn.Application.LinesToPoints(1.15)
                .Alignment = WdParagraphAlignment.wdAlignParagraphCenter
                .WidowControl = True
                .KeepWithNext = False
                .KeepTogether = False
                .PageBreakBefore = False
                .NoLineNumber = False
                .Hyphenation = True
                .FirstLineIndent = Globals.ThisAddIn.Application.InchesToPoints(0)
                .OutlineLevel = WdOutlineLevel.wdOutlineLevelBodyText
                .CharacterUnitLeftIndent = 0
                .CharacterUnitRightIndent = 0
                .CharacterUnitFirstLineIndent = 0
                .LineUnitBefore = 0
                .LineUnitAfter = 0
                .MirrorIndents = False
                .TextboxTightWrap = WdTextboxTightWrap.wdTightNone
            End With

            .ActiveDocument.Styles(cINDEX_HEADING_STYLE).NoSpaceBetweenParagraphsOfSameStyle _
                = False
            .ActiveDocument.Styles(cINDEX_HEADING_STYLE).ParagraphFormat.TabStops.ClearAll()

            With .ActiveDocument.Styles(cINDEX_HEADING_STYLE).ParagraphFormat
                With .Shading
                    .Texture = WdTextureIndex.wdTextureNone
                    .ForegroundPatternColor = WdColor.wdColorAutomatic
                    .BackgroundPatternColor = WdColor.wdColorAutomatic
                End With

                .Borders(WdBorderType.wdBorderLeft).LineStyle = WdLineStyle.wdLineStyleNone
                .Borders(WdBorderType.wdBorderRight).LineStyle =  WdLineStyle.wdLineStyleNone
                .Borders(WdBorderType.wdBorderTop).LineStyle =  WdLineStyle.wdLineStyleNone
                .Borders(WdBorderType.wdBorderBottom).LineStyle =  WdLineStyle.wdLineStyleNone

                With .Borders
                    .DistanceFromTop = 1
                    .DistanceFromLeft = 4
                    .DistanceFromBottom = 1
                    .DistanceFromRight = 4
                    .Shadow = False
                End With
            End With

            .ActiveDocument.Styles(cINDEX_HEADING_STYLE).LanguageID = WdLanguageID.wdEnglishUS
            .ActiveDocument.Styles(cINDEX_HEADING_STYLE).NoProofing = False
            .ActiveDocument.Styles(cINDEX_HEADING_STYLE).Frame.Delete()
        End With
    End Sub

    Private Sub CreateStyle_IndexWord()
        With Globals.ThisAddIn.Application
            .ActiveDocument.Styles.Add(Name:=cINDEX_WORD_STYLE, Type:=WdStyleType.wdStyleTypeCharacter)

            With .ActiveDocument.Styles(cINDEX_WORD_STYLE)
                .QuickStyle = True

                With .Font
                    .Name = "+Body"
                    .Size = 10
                    .Bold = True
                    .Color = 12611584
                    .Borders(1).LineStyle = WdLineStyle.wdLineStyleNone
                    .Borders.Shadow = False

                    '    .Italic = False
                    '    .Underline = wdUnderlineNone
                    '    .UnderlineColor = wdColorAutomatic
                    '    .StrikeThrough = False
                    '    .DoubleStrikeThrough = False
                    '    .Outline = False
                    '    .Emboss = False
                    '    .Shadow = False
                    '    .Hidden = False
                    '    .SmallCaps = False
                    '    .AllCaps = False
                    '    .Color = 12611584
                    '    .Engrave = False
                    '    .Superscript = False
                    '    .Subscript = False
                    '    .Scaling = 100
                    '    .Kerning = 0
                    '    .Animation = wdAnimationNone
                    '    .Ligatures = wdLigaturesNone
                    '    .NumberSpacing = wdNumberSpacingDefault
                    '    .NumberForm = wdNumberFormDefault
                    '    .StylisticSet = wdStylisticSetDefault
                    '    .ContextualAlternates = 0

                    With .Shading
                        .Texture = WdTextureIndex.wdTextureNone
                        .ForegroundPatternColor = WdColor.wdColorAutomatic
                        .BackgroundPatternColor = WdColor.wdColorAutomatic
                    End With
                End With
            End With
        End With
    End Sub

    Private Sub CreateIndexStyles()
        CreateStyle_IndexWord()
        CreateStyle_IndexHeading()

        'ActiveDocument.Styles.Add(Name:="IndexWord", Type:=wdStyleTypeCharacter)
        'With ActiveDocument.Styles("IndexWord").Font
        '    .Name = "+Body"
        '    .Size = 10
        '    .Bold = True
        '    .Color = 12611584
        'End With
        'With ActiveDocument.Styles("IndexWord").Font
        '    With .Shading
        '        .Texture = wdTextureNone
        '        .ForegroundPatternColor = wdColorAutomatic
        '        .BackgroundPatternColor = wdColorAutomatic
        '    End With
        '    .Borders(1).LineStyle = wdLineStyleNone
        '    .Borders.Shadow = False
        'End With
        'Selection.MoveDown(Unit:=wdLine, Count:=1)
        'Selection.MoveRight(Unit:=wdCharacter, Count:=3)
        'Selection.MoveLeft(Unit:=wdCharacter, Count:=13, Extend:=wdExtend)
        'ActiveDocument.Styles.Add(Name:="IndexHeading", Type:=wdStyleTypeParagraph)
        'ActiveDocument.Styles("IndexHeading").AutomaticallyUpdate = False
        'With ActiveDocument.Styles("IndexHeading").Font
        '    .Name = "+Body"
        '    .Size = 12
        '    .Bold = True
        '    .Italic = False
        '    .Underline = wdUnderlineNone
        '    .UnderlineColor = wdColorAutomatic
        '    .StrikeThrough = False
        '    .DoubleStrikeThrough = False
        '    .Outline = False
        '    .Emboss = False
        '    .Shadow = False
        '    .Hidden = False
        '    .SmallCaps = False
        '    .AllCaps = False
        '    .Color = 12611584
        '    .Engrave = False
        '    .Superscript = False
        '    .Subscript = False
        '    .Scaling = 100
        '    .Kerning = 0
        '    .Animation = wdAnimationNone
        '    .Ligatures = wdLigaturesNone
        '    .NumberSpacing = wdNumberSpacingDefault
        '    .NumberForm = wdNumberFormDefault
        '    .StylisticSet = wdStylisticSetDefault
        '    .ContextualAlternates = 0
        'End With
        'With ActiveDocument.Styles("IndexHeading").ParagraphFormat
        '    .LeftIndent = InchesToPoints(0)
        '    .RightIndent = InchesToPoints(0)
        '    .SpaceBefore = 0
        '    .SpaceBeforeAuto = False
        '    .SpaceAfter = 6
        '    .SpaceAfterAuto = False
        '    .LineSpacingRule = wdLineSpaceMultiple
        '    .LineSpacing = LinesToPoints(1.15)
        '    .Alignment = wdAlignParagraphCenter
        '    .WidowControl = True
        '    .KeepWithNext = False
        '    .KeepTogether = False
        '    .PageBreakBefore = False
        '    .NoLineNumber = False
        '    .Hyphenation = True
        '    .FirstLineIndent = InchesToPoints(0)
        '    .OutlineLevel = wdOutlineLevelBodyText
        '    .CharacterUnitLeftIndent = 0
        '    .CharacterUnitRightIndent = 0
        '    .CharacterUnitFirstLineIndent = 0
        '    .LineUnitBefore = 0
        '    .LineUnitAfter = 0
        '    .MirrorIndents = False
        '    .TextboxTightWrap = wdTightNone
        'End With
        'ActiveDocument.Styles("IndexHeading").NoSpaceBetweenParagraphsOfSameStyle _
        '    = False
        'ActiveDocument.Styles("IndexHeading").ParagraphFormat.TabStops.ClearAll()
        'With ActiveDocument.Styles("IndexHeading").ParagraphFormat
        '    With .Shading
        '        .Texture = wdTextureNone
        '        .ForegroundPatternColor = wdColorAutomatic
        '        .BackgroundPatternColor = wdColorAutomatic
        '    End With
        '    .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        '    .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        '    .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        '    .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        '    With .Borders
        '        .DistanceFromTop = 1
        '        .DistanceFromLeft = 4
        '        .DistanceFromBottom = 1
        '        .DistanceFromRight = 4
        '        .Shadow = False
        '    End With
        'End With
        'ActiveDocument.Styles("IndexHeading").LanguageID = wdEnglishUS
        'ActiveDocument.Styles("IndexHeading").NoProofing = False
        'ActiveDocument.Styles("IndexHeading").Frame.Delete()
    End Sub

    Private Sub FindIndexWords()
        With Globals.ThisAddIn.Application
            .Selection.HomeKey(Unit:=WdUnits.wdStory)
            .Selection.ClearFormatting()

            Try
                With .Selection.Find
                    .Style = Globals.ThisAddIn.Application.ActiveDocument.Styles(cINDEX_WORD_STYLE)
                    .Text = ""
                    .Replacement.Text = ""
                    .Forward = True
                    .Wrap = WdFindWrap.wdFindStop
                    .Format = True
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With

                .Selection.Find.Execute()

                While (.Selection.Find.Found)
                    Dim match As Range = .Selection.Range

                    Common.WriteToDebugWindow(String.Format("start:{0} end:{1} text:{2} style:{3}",
                                                            match.Start, match.End, match.Text, match.Style.ToString()))

                    '.ActiveDocument.Indexes.MarkEntry(Range:=.Selection.Range, Entry:=.Selection.Text, EntryAutoText:=.Selection.Text)

                    '' The Index Marker unfortunately takes on the style of the selection, so,
                    '' Search again for the style which finds the marker

                    '.Selection.Find.Execute()

                    '' Remove the formatting from the Selection

                    '.Selection.ClearCharacterAllFormatting()

                    '' Collapse the selection so we continue to search the document

                    '.Selection.Collapse()

                    ' And search for the next word to tag (if any)

                    .Selection.Find.Execute()
                End While
            Catch ex As Exception
                MessageBox.Show(ex.ToString())
            End Try
        End With
    End Sub

    Private Function LoadWordsFromXMLFile(ByVal fileName As String) As XElement
        Dim replacementWords As XElement = Nothing

        openFileDialog.FileName = fileName
        openFileDialog.Filter = "XML Files (*.xml)|*.xml|All files (*.*)|(*.*)"

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            Using streamReader As New StreamReader(openFileDialog.FileName)
                replacementWords = XElement.Load(streamReader)
            End Using
        End If

        Return replacementWords
    End Function

    Private Sub MarkIndexWords()
        ' TODO(crhodes) get this from Config class.  Think about initial directory
        Dim indexWords as XElement = LoadWordsFromXMLFile("IndexWords.xml")

        For Each word As XElement In indexWords.Elements()
            Common.WriteToDebugWindow(word.Value)

            If word.Value.Length > 0 Then
                ApplyStyleToWords(word.Value, cINDEX_WORD_STYLE)
            End If
        Next
    End Sub

    Private Sub ReplaceWord(ByVal phrase As String, ByVal replacementWord As String, ByVal indexWordsOnly As Boolean)
        Common.WriteToDebugWindow(String.Format("ReplaceWord:{0} indexWordsOnly:{1}", phrase, indexWordsOnly))

        With Globals.ThisAddIn.Application
            .Selection.Find.ClearFormatting

            If indexWordsOnly Then
                .Selection.Find.Style = .ActiveDocument.Styles(cINDEX_WORD_STYLE)                
            End If

            .Selection.Find.Replacement.ClearFormatting

            With .Selection.Find
                .Text = phrase
                .Replacement.Text = replacementWord
                .Forward = True
                .Wrap = WdFindWrap.wdFindContinue
                .Format = True
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With

            .Selection.Find.Execute (Replace:=WdReplace.wdReplaceAll)
        End With
    End Sub

    Private Sub SaveReplacementWordsToXMLFile()
        Dim replacementWords As XElement = New XElement("ReplacementWords")
                             
        For Each field As Field In Globals.ThisAddIn.Application.ActiveDocument.Fields
            'Common.WriteToDebugWindow(String.Format("field:{0} text:>{1}<", lField.Type, lField.Code.Text))

            If field.Type = WdFieldType.wdFieldIndexEntry Then
                Dim fieldText As String = field.Code.Text
                Dim indexWord As String = fieldText.Substring(5, fieldText.Length - 7)

                Common.WriteToDebugWindow(String.Format("  fieldText:>{0}< indexWord:>{1}<", fieldText, indexWord))
                replacementWords.Add(New XElement("ReplacementWord", indexWord))
            End If
        Next

        Common.WriteToDebugWindow(replacementWords.ToString())

        saveFileDialog.Filter = "XML Files (*.xml)|*.xml|All files (*.*)|(*.*)"
        ' TODO(crhodes) get this from Config class.  Think about initial directory
        saveFileDialog.FileName = "ReplacementWords.xml"

        If saveFileDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Using streamWriter As StreamWriter = New StreamWriter(saveFileDialog.FileName)
                streamWriter.Write(replacementWords.ToString())
                streamWriter.Flush()
            End Using
        End If

    End Sub

    Private Sub TagIndexWords()
        With Globals.ThisAddIn.Application
            .Selection.HomeKey(Unit:=WdUnits.wdStory)
            .Selection.ClearFormatting()

            With .Selection.Find
                .Style = Globals.ThisAddIn.Application.ActiveDocument.Styles(cINDEX_WORD_STYLE)
                .Text = ""
                .Replacement.Text = ""
                .Forward = True
                .Wrap = WdFindWrap.wdFindStop
                .Format = True
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With

            .Selection.Find.Execute()

            While (.Selection.Find.Found)
                'Dim match As Range = .Selection.Range

                'Common.WriteToDebugWindow(String.Format("start:{0} end:{1} text:{2} style:{3}",
                '                                        match.Start, match.End, match.Text, match.Style.ToString()))

                .ActiveDocument.Indexes.MarkEntry(Range:=.Selection.Range, Entry:=.Selection.Text, EntryAutoText:=.Selection.Text)

                ' The Index Marker unfortunately takes on the style of the selection, so,
                ' Search again for the style which finds the marker

                .Selection.Find.Execute()

                ' Remove the formatting from the Selection

                .Selection.ClearCharacterAllFormatting()

                ' Collapse the selection so we continue to search the document

                .Selection.Collapse()

                ' And search for the next word to tag (if any)

                .Selection.Find.Execute()
            End While
        End With
    End Sub

    Private Sub ZapReplacementWords()
        ' TODO(crhodes) get this from Config class
        Dim replacementWordsXML as XElement = LoadWordsFromXMLFile("ReplacementWords.xml")

        For Each word As XElement In replacementWordsXML.Elements()
            Common.WriteToDebugWindow(word.Value)

            If word.Value.Length > 0 Then
                ReplaceWord(word.Value, txtReplacementWord.Text, ckIndexWordsOnly.Checked)
            End If
        Next
    End Sub
#End Region
End Class
