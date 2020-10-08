using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using Microsoft.Office.Interop.Word;
using WordHlp = AddinHelper.Word;

namespace SupportTools_Word.Actions
{
    class Word_PageFormatting
    {

        public static void AddFooter()
        {
            try
            {
                var wordApp = Globals.ThisAddIn.Application;
                //Debug.Print(.ActiveDocument.Sections.Count())

                foreach (Section documentSection in wordApp.ActiveDocument.Sections)
                {
                    //Debug.Print(documentSection.Index)
                    if (documentSection.Index > 1)
                    {
                        //documentSection.PageSetup.
                        // This needs to be smarter about sections.  Linking to previous pulls the formatting, too.
                        // That makes landscape pages have the same margins as the portrait pages.
                        documentSection.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = true;
                        documentSection.Headers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].LinkToPrevious = true;
                        //Continue For
                    }

                    if (wordApp.ActiveWindow.View.SplitSpecial != WdSpecialPane.wdPaneNone)
                    {
                        wordApp.ActiveWindow.Panes[2].Close();
                    }

                    if (wordApp.ActiveWindow.ActivePane.View.Type == WdViewType.wdNormalView | wordApp.ActiveWindow.ActivePane.View.Type == WdViewType.wdOutlineView)
                    {
                        wordApp.ActiveWindow.ActivePane.View.Type = WdViewType.wdPrintView;
                    }

                    wordApp.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekCurrentPageFooter;
                    //If Selection.HeaderFooter.IsHeader = True Then
                    //    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
                    //Else
                    //    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
                    //End If

                    // Delete any existing stuff and make the font small.
                    // Word 2007: Adjust paragraph spacing as "Normal" paragraph style has space
                    // between lines.  May decide to add our own footer style.

                    var paragraphFormat = wordApp.Selection.ParagraphFormat;



                    wordApp.Selection.WholeStory();
                    wordApp.Selection.Range.Delete();
                    wordApp.Selection.WholeStory();
                    // Decided not to do this incase the style is not available.
                    //.Selection.Style = .ActiveDocument.Styles("No Spacing")
                    wordApp.Selection.Font.Name = "Arial";
                    wordApp.Selection.Font.Size = 5;


                    paragraphFormat.TabStops.ClearAll();
                    paragraphFormat.LeftIndent = Globals.ThisAddIn.Application.InchesToPoints(0);
                    paragraphFormat.RightIndent = Globals.ThisAddIn.Application.InchesToPoints(0);
                    paragraphFormat.SpaceBefore = 0;
                    paragraphFormat.SpaceBeforeAuto = 0;
                    paragraphFormat.SpaceAfter = 0;
                    paragraphFormat.SpaceAfterAuto = 0;
                    paragraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                    //.Alignment = wdAlignParagraphLeft
                    //.WidowControl = True
                    //.KeepWithNext = False
                    //.KeepTogether = False
                    //.PageBreakBefore = False
                    //.NoLineNumber = False
                    //.Hyphenation = True
                    //.FirstLineIndent = InchesToPoints(0)
                    //.OutlineLevel = wdOutlineLevelBodyText
                    //.CharacterUnitLeftIndent = 0
                    //.CharacterUnitRightIndent = 0
                    //.CharacterUnitFirstLineIndent = 0
                    //.LineUnitBefore = 0
                    //.LineUnitAfter = 0
                    //.MirrorIndents = False
                    //.TextboxTightWrap = wdTightNone


                    //.Selection.TypeParagraph()

                    //' Add some tabs to help format the fields.

                    //.Selection.ParagraphFormat.TabStops.Add( _
                    //    Position:=.InchesToPoints(0.5), _
                    //    Alignment:=WdTabAlignment.wdAlignTabLeft, _
                    //    Leader:=WdTabLeader.wdTabLeaderSpaces)
                    //.Selection.ParagraphFormat.TabStops.Add( _
                    //    Position:=.InchesToPoints(1.5), _
                    //    Alignment:=WdTabAlignment.wdAlignTabLeft, _
                    //    Leader:=WdTabLeader.wdTabLeaderSpaces)

                    // Adjust right margin point depending on page orientation.

                    //For Each documentSection As Section In .ActiveDocument.Sections
                    float rightMargin = 0;
                    float leftMargin = 0;
                    float pageWidth = 0;

                    //rightMargin = .ActiveDocument.PageSetup.RightMargin
                    //leftMargin = .ActiveDocument.PageSetup.LeftMargin
                    //pageWidth = .ActiveDocument.PageSetup.PageWidth

                    rightMargin = documentSection.PageSetup.RightMargin;
                    leftMargin = documentSection.PageSetup.LeftMargin;
                    pageWidth = documentSection.PageSetup.PageWidth;

                    //Debug.Print(String.Format("orientation: {0}   pageWidth: {1}   leftMargin: {2}   rightMargin: {3}", documentSection.PageSetup.Orientation, pageWidth, leftMargin, rightMargin))

                    //Next


                    float rightMarginPoint = 0;

                    //rightMarginPoint = .ActiveDocument.PageSetup.PageWidth _
                    //    - .ActiveDocument.PageSetup.RightMargin _
                    //    - .ActiveDocument.PageSetup.LeftMargin

                    rightMarginPoint = pageWidth - leftMargin - rightMargin;

                    //If .ActiveDocument.PageSetup.Orientation = WdOrientation.wdOrientPortrait Then
                    //    .Selection.ParagraphFormat.TabStops.Add( _
                    //        Position:=CInt(rightMarginPoint), _
                    //        Alignment:=WdTabAlignment.wdAlignTabRight, _
                    //        Leader:=WdTabLeader.wdTabLeaderSpaces)
                    //Else
                    //    .Selection.ParagraphFormat.TabStops.Add( _
                    //        Position:=CInt(rightMarginPoint), _
                    //        Alignment:=WdTabAlignment.wdAlignTabRight, _
                    //        Leader:=WdTabLeader.wdTabLeaderSpaces)
                    //End If

                    if (documentSection.PageSetup.Orientation == WdOrientation.wdOrientPortrait)
                    {
                        wordApp.Selection.ParagraphFormat.TabStops.Add(Position: Convert.ToInt32(rightMarginPoint), Alignment: WdTabAlignment.wdAlignTabRight, Leader: WdTabLeader.wdTabLeaderSpaces);
                    }
                    else
                    {
                        wordApp.Selection.ParagraphFormat.TabStops.Add(Position: Convert.ToInt32(rightMarginPoint), Alignment: WdTabAlignment.wdAlignTabRight, Leader: WdTabLeader.wdTabLeaderSpaces);
                    }

                    wordApp.Selection.TypeParagraph();

                    wordApp.Selection.Fields.Add(Range: wordApp.Selection.Range, Type: WdFieldType.wdFieldEmpty, Text: "FILENAME \\p ", PreserveFormatting: true);

                    wordApp.Selection.TypeText(Text: "\tPage ");

                    wordApp.Selection.Fields.Add(Range: wordApp.Selection.Range, Type: WdFieldType.wdFieldEmpty, Text: "PAGE ", PreserveFormatting: true);
                    wordApp.Selection.TypeText(Text: " of ");
                    wordApp.Selection.Fields.Add(Range: wordApp.Selection.Range, Type: WdFieldType.wdFieldEmpty, Text: "NUMPAGES ", PreserveFormatting: true);

                    wordApp.Selection.TypeParagraph();

                    // Add some more tabs to help format the fields.

                    wordApp.Selection.ParagraphFormat.TabStops.Add(Position: wordApp.InchesToPoints(0.5f), Alignment: WdTabAlignment.wdAlignTabLeft, Leader: WdTabLeader.wdTabLeaderSpaces);
                    wordApp.Selection.ParagraphFormat.TabStops.Add(Position: wordApp.InchesToPoints(1.5f), Alignment: WdTabAlignment.wdAlignTabLeft, Leader: WdTabLeader.wdTabLeaderSpaces);

                    wordApp.Selection.TypeText(Text: "Created:\t");

                    wordApp.Selection.Fields.Add(Range: wordApp.Selection.Range, Type: WdFieldType.wdFieldEmpty, Text: "CREATEDATE ", PreserveFormatting: true);

                    wordApp.Selection.TypeText(Text: "\tAuthor: ");
                    wordApp.Selection.Fields.Add(Range: wordApp.Selection.Range, Type: WdFieldType.wdFieldEmpty, Text: "AUTHOR ", PreserveFormatting: true);

                    wordApp.Selection.TypeText(Text: "\tTitle: ");
                    wordApp.Selection.Fields.Add(Range: wordApp.Selection.Range, Type: WdFieldType.wdFieldEmpty, Text: "TITLE ", PreserveFormatting: true);

                    wordApp.Selection.TypeParagraph();

                    wordApp.Selection.TypeText(Text: "Last Saved:\t");
                    wordApp.Selection.Fields.Add(Range: wordApp.Selection.Range, Type: WdFieldType.wdFieldEmpty, Text: "SAVEDATE ", PreserveFormatting: true);

                    wordApp.Selection.TypeText(Text: "\tBy: ");

                    wordApp.Selection.Fields.Add(Range: wordApp.Selection.Range, Type: WdFieldType.wdFieldEmpty, Text: "LASTSAVEDBY ", PreserveFormatting: true);

                    wordApp.Selection.TypeText(Text: "\tSubject: ");
                    wordApp.Selection.Fields.Add(Range: wordApp.Selection.Range, Type: WdFieldType.wdFieldEmpty, Text: "SUBJECT ", PreserveFormatting: true);

                    wordApp.Selection.TypeParagraph();
                    wordApp.Selection.TypeText(Text: "Printed:\t");
                    wordApp.Selection.Fields.Add(Range: wordApp.Selection.Range, Type: WdFieldType.wdFieldEmpty, Text: "PRINTDATE ", PreserveFormatting: true);
                    wordApp.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekMainDocument;
                    //.ActiveWindow.ActivePane.Close()
                }

                // g_Word
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("AddFooter: {0}", ex));
            }
        }
        // Footer_Add


    }
}