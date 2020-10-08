using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

using Microsoft.Office.Interop.Word;

namespace Compliance_Office2010Addin_Word.User_Interface.Task_Panes
{
    public partial class TaskPane_ComplianceUtil : UserControl
    {

        private const string cINDEX_WORD_STYLE = "IndexWord";
        private const string cINDEX_HEADING_STYLE = "IndexHeading";

        public TaskPane_ComplianceUtil()
        {
            InitializeComponent();
        }

        #region Event Handlers
        private void btnClearStylesFromWords_Click(object sender, EventArgs e)
        {
            ClearStylesFromWords();
        }

        private void btnCreateIndexStyles_Click(object sender, EventArgs e)
        {
            CreateIndexStyles();
        }

        private void btnDeleteIndexFields_Click(object sender, EventArgs e)
        {
            DeleteIndexFields();
        }
        private void btnDisplayReadability_Click(object sender, EventArgs e)
        {
            DisplayReadabilityStatistics();
        }

        private void btnFindIndexFields_Click(object sender, EventArgs e)
        {
            FindIndexFields();
        }
        private void btnFindIndexHeadingStyle_Click(object sender, EventArgs e)
        {
            FindIndexHeadingStyle();
        }

        private void btnFindIndexWordStyle_Click(object sender, EventArgs e)
        {
            FindIndexWordStyle();
        }

        private void btnMarkIndexWords_Click(object sender, EventArgs e)
        {
            MarkIndexWords();
        }

        private void btnResetCheck_Click(object sender, EventArgs e)
        {

        }

        private void btnSaveReplacementWords_Click(object sender, EventArgs e)
        {
            SaveReplacementWordsToXMLFile();
        }

        private void btnTagIndexHeadingStyleWords_Click(object sender, EventArgs e)
        {
            TagIndexHeadingStyleWords();
        }

        private void btnTagIndexWordStyleWords_Click(object sender, EventArgs e)
        {
            TagIndexWordStyleWords();
        }

        private void btnUpdateIndex_Click(object sender, EventArgs e)
        {
            UpdateIndex();
        }

        private void btnZapReplacementWords_Click(object sender, EventArgs e)
        {
            ZapReplacementWords();
        }
        #endregion

        #region Main Function Routines

        private void ApplyStyleToWords(string indexWord, string style)
        {
            AddinHelper.Common.WriteToDebugWindow(String.Format("ApplyStyleToWords:{0} style:{1}", indexWord, style));

            Document activeDoc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection selection = Globals.ThisAddIn.Application.Selection;

            selection.HomeKey(Unit: WdUnits.wdStory);
            selection.Find.ClearFormatting();

            selection.Find.Text = indexWord;
            selection.Find.Forward = true;
            selection.Find.Wrap = WdFindWrap.wdFindContinue;
            selection.Find.Format = true;
            selection.Find.MatchCase = false;
            selection.Find.MatchWholeWord = false;
            selection.Find.MatchWildcards = false;
            selection.Find.MatchSoundsLike = false;
            selection.Find.MatchAllWordForms = false;

            selection.Find.Execute();

            while(selection.Find.Found)
            {
                Range match = selection.Range;

                DialogResult choice = MessageBox.Show("Mark as IndexWord?", indexWord, MessageBoxButtons.YesNoCancel);

                switch(choice)
                {
                    case DialogResult.Yes:
                        AddinHelper.Common.WriteToDebugWindow(
                            String.Format("start:{0} end:{1} text:{2} style:{3}",
                                   match.Start, match.End, match.Text, match.get_Style()));

                        // There may be junk on the selection that will show through the style.
                        // Remove it.

                        selection.ClearFormatting();

                        // Then apply the style.

                        selection.Range.set_Style(style);                        
                        break;

                    case DialogResult.No:
                        
                        break;

                    case DialogResult.Cancel:
                        return;
                        break;
                }

                selection.Find.Execute();
            }
        }

        private void ClearStyle(string styleName)
        {
            Document activeDoc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection selection = Globals.ThisAddIn.Application.Selection;

            selection.HomeKey(Unit: WdUnits.wdStory);
            selection.ClearFormatting();

            try
            {
                selection.Find.set_Style(activeDoc.Styles[styleName]);
                selection.Find.Text = "";
                selection.Find.Replacement.Text = "";
                selection.Find.Forward = true;
                selection.Find.Wrap = WdFindWrap.wdFindStop;
                selection.Find.Format = true;
                selection.Find.MatchCase = false;
                selection.Find.MatchWildcards = false;
                selection.Find.MatchSoundsLike = false;
                selection.Find.MatchAllWordForms = false;

                selection.Find.Execute();

                while(selection.Find.Found)
                {
                    Range match = selection.Range;
                    Style matchStyle = match.get_Style();

                    AddinHelper.Common.WriteToDebugWindow(
                        String.Format("start:{0} end:{1} text:{2} style:{3}",
                        match.Start, match.End, match.Text, matchStyle.NameLocal));

                    selection.ClearFormatting();

                    // Collapse the selection so we continue to search the document

                    selection.Collapse(WdCollapseDirection.wdCollapseEnd);

                    // And search for the next word to tag (if any)

                    selection.Find.Execute();
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void ClearStylesFromWords()
        {
            ClearStyle(cINDEX_HEADING_STYLE);
            ClearStyle(cINDEX_WORD_STYLE);
        }

        private void CreateStyle_IndexHeading()
        {
            Common.WordHelper.DeleteStyle(cINDEX_HEADING_STYLE);

            Document activeDoc = Globals.ThisAddIn.Application.ActiveDocument;

            Style newStyle = activeDoc.Styles.Add(Name: cINDEX_HEADING_STYLE, Type: WdStyleType.wdStyleTypeParagraph);
            newStyle.AutomaticallyUpdate = false;
            newStyle.NoSpaceBetweenParagraphsOfSameStyle = false;
            newStyle.ParagraphFormat.TabStops.ClearAll();
            newStyle.LanguageID = WdLanguageID.wdEnglishUS;
            newStyle.NoProofing = 0;
            newStyle.Frame.Delete();

            //activeDoc.Styles[cINDEX_HEADING_SYTLE].AutomaticallyUpdate = false;

            Microsoft.Office.Interop.Word.Font font = activeDoc.Styles[cINDEX_HEADING_STYLE].Font;

            font.Name = "+Body";
            font.Size = 12;
            font.Italic = 0;
            font.Underline = WdUnderline.wdUnderlineNone;
            font.UnderlineColor = WdColor.wdColorAutomatic;
            font.StrikeThrough = 0;
            font.Outline = 0;
            font.Emboss = 0;
            font.Shadow = 0;
            font.Hidden = 0;
            font.SmallCaps = 0;
            font.AllCaps = 0;
            font.Color = WdColor.wdColorBlue; 
            font.Engrave = 0;
            font.Subscript = 0;
            font.Scaling = 100;
            font.Kerning = 0;
            font.Animation = WdAnimation.wdAnimationNone;
            font.Ligatures = WdLigatures.wdLigaturesNone;
            font.NumberSpacing = WdNumberSpacing.wdNumberSpacingDefault;
            font.NumberForm = WdNumberForm.wdNumberFormDefault;
            font.StylisticSet = WdStylisticSet.wdStylisticSetDefault;
            font.ContextualAlternates = 0;

            Microsoft.Office.Interop.Word.ParagraphFormat paragraphFormat = activeDoc.Styles[cINDEX_HEADING_STYLE].ParagraphFormat;

            paragraphFormat.LeftIndent = Globals.ThisAddIn.Application.InchesToPoints(0);
            paragraphFormat.RightIndent = Globals.ThisAddIn.Application.InchesToPoints(0);
            paragraphFormat.SpaceBefore = 0;
            paragraphFormat.SpaceBeforeAuto = 0;
            paragraphFormat.SpaceAfter = 6;
            paragraphFormat.SpaceAfterAuto = 0;
            paragraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceMultiple;
            paragraphFormat.LineSpacing = Globals.ThisAddIn.Application.LinesToPoints((float)1.15);
            paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            // This throws exception.
            //paragraphFormat.WidowControl = 1;
            paragraphFormat.KeepTogether = 0;
            paragraphFormat.KeepWithNext = 0;
            paragraphFormat.PageBreakBefore = 0;
            paragraphFormat.NoLineNumber = 0;
            paragraphFormat.Hyphenation = 0;
            paragraphFormat.FirstLineIndent = Globals.ThisAddIn.Application.InchesToPoints(0);
            paragraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevelBodyText;
            paragraphFormat.CharacterUnitLeftIndent = 0;
            paragraphFormat.CharacterUnitRightIndent = 0;
            paragraphFormat.CharacterUnitFirstLineIndent = 0;
            paragraphFormat.LineUnitBefore = 0;
            paragraphFormat.LineUnitAfter = 0;
            paragraphFormat.MirrorIndents = 0;
            paragraphFormat.TextboxTightWrap = WdTextboxTightWrap.wdTightNone;

            paragraphFormat.Shading.Texture = WdTextureIndex.wdTextureNone;
            paragraphFormat.Shading.ForegroundPatternColor = WdColor.wdColorAutomatic;
            paragraphFormat.Shading.BackgroundPatternColor = WdColor.wdColorAutomatic;

            paragraphFormat.Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleNone;
            paragraphFormat.Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleNone;
            paragraphFormat.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleNone;
            paragraphFormat.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;

            paragraphFormat.Borders.DistanceFromTop = 1;
            paragraphFormat.Borders.DistanceFromLeft = 4;
            paragraphFormat.Borders.DistanceFromBottom = 1;
            paragraphFormat.Borders.DistanceFromRight = 4;
        }

        private void CreateStyle_IndexWord()
        {
            Common.WordHelper.DeleteStyle(cINDEX_WORD_STYLE);

            Document activeDoc = Globals.ThisAddIn.Application.ActiveDocument;

            Style newStyle = activeDoc.Styles.Add(Name: cINDEX_WORD_STYLE, Type: WdStyleType.wdStyleTypeCharacter);
            newStyle.QuickStyle = true;

            //newStyle.AutomaticallyUpdate = false;
            //newStyle.NoSpaceBetweenParagraphsOfSameStyle = false;
            //newStyle.ParagraphFormat.TabStops.ClearAll();
            //newStyle.LanguageID = WdLanguageID.wdEnglishUS;
            //newStyle.NoProofing = 0;
            //newStyle.Frame.Delete();

            //activeDoc.Styles[cINDEX_HEADING_SYTLE].AutomaticallyUpdate = false;

            Microsoft.Office.Interop.Word.Font font = activeDoc.Styles[cINDEX_WORD_STYLE].Font;

            font.Name = "+Body";
            font.Size = 12;
            font.Bold = 1;
            font.Italic = 0;
            font.Underline = WdUnderline.wdUnderlineNone;
            font.UnderlineColor = WdColor.wdColorAutomatic;
            font.StrikeThrough = 0;
            font.Outline = 0;
            font.Emboss = 0;
            font.Shadow = 0;
            font.Hidden = 0;
            font.SmallCaps = 0;
            font.AllCaps = 0;
            font.Color = WdColor.wdColorBlue; 
            font.Engrave = 0;
            font.Subscript = 0;
            font.Scaling = 100;
            font.Kerning = 0;
            font.Animation = WdAnimation.wdAnimationNone;
            font.Ligatures = WdLigatures.wdLigaturesNone;
            font.NumberSpacing = WdNumberSpacing.wdNumberSpacingDefault;
            font.NumberForm = WdNumberForm.wdNumberFormDefault;
            font.StylisticSet = WdStylisticSet.wdStylisticSetDefault;
            font.ContextualAlternates = 0;
            //font.Borders[1].LineStyle = WdLineStyle.wdLineStyleNone;
            font.Borders.Shadow = false;
            font.Shading.Texture = WdTextureIndex.wdTextureNone;
            font.Shading.ForegroundPatternColor = WdColor.wdColorAutomatic;
            font.Shading.BackgroundPatternColor = WdColor.wdColorAutomatic;
        }

        private void CreateIndexStyles()
        {
            CreateStyle_IndexWord();
            CreateStyle_IndexHeading();
        }

        private void DeleteIndexFields()
        {
            // Search the file for IndexEntry fields and delete them.

            foreach (Field field in Globals.ThisAddIn.Application.ActiveDocument.Fields)
            {
            	if (field.Type == WdFieldType.wdFieldIndexEntry)
                {
                    field.Delete();
                }
            }
        }

        private void DisplayReadabilityStatistics()
        {
            ReadabilityStatistics stats = Globals.ThisAddIn.Application.ActiveDocument.ReadabilityStatistics;

            txtReadabilityStatistics.Clear();

            foreach(ReadabilityStatistic stat in stats)
            {
                txtReadabilityStatistics.AppendText(String.Format("{0}: {1}{2}", stat.Name, stat.Value, Environment.NewLine));
            }
        }

        private void FindIndexFields()
        {
            foreach (Field field in Globals.ThisAddIn.Application.ActiveDocument.Fields)
            {
            	if (field.Type == WdFieldType.wdFieldIndexEntry)
                {
                    field.Select();

                    // Pause before searching for next word

                    System.Threading.Thread.Sleep(750);
                }
            }
        }

        private void FindIndexHeadingStyle()
        {
            MessageBox.Show(String.Format("Searching for words marked with {0} style", cINDEX_HEADING_STYLE));
            FindIndexWords(cINDEX_HEADING_STYLE);
        }

        private void FindIndexWordStyle()
        {
            MessageBox.Show(String.Format("Searching for words marked with {0} style", cINDEX_WORD_STYLE));
            FindIndexWords(cINDEX_WORD_STYLE);
        }

        private void FindIndexWords(string styleName)
        {
            Document activeDoc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection selection = Globals.ThisAddIn.Application.Selection;

            selection.HomeKey(Unit: WdUnits.wdStory);
            selection.ClearFormatting();

            try
            {
                selection.Find.set_Style(activeDoc.Styles[styleName]);
                selection.Find.Text = "";
                selection.Find.Replacement.Text = "";
                selection.Find.Forward = true;
                selection.Find.Wrap = WdFindWrap.wdFindStop;
                selection.Find.Format = true;
                selection.Find.MatchCase = false;
                selection.Find.MatchWildcards = false;
                selection.Find.MatchSoundsLike = false;
                selection.Find.MatchAllWordForms = false;

                selection.Find.Execute();

                while(selection.Find.Found)
                {
                    Range match = selection.Range;

                    AddinHelper.Common.WriteToDebugWindow(
                        String.Format("start:{0} end:{1} text:{2} style:{3}",
                        match.Start, match.End, match.Text, match.get_Style()));

                    // Pause before searching for next word

                    System.Threading.Thread.Sleep(750);

                    // Collapse the selection so we continue to search the document

                    selection.Collapse(WdCollapseDirection.wdCollapseEnd);

                    // And search for the next word to tag (if any)

                    selection.Find.Execute();
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private XElement LoadWordsFromXMLFile(string fileName)
        {
            XElement replacementWords = null;

            openFileDialog1.FileName = fileName;
            openFileDialog1.Filter = "XML Files (*.xml)|*.xml|All files (*.*)|(*.*)";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                using(StreamReader streamReader = new StreamReader(openFileDialog1.FileName))
                {
                    replacementWords = XElement.Load(streamReader);
                }
            }

            return replacementWords;
        }

        private void MarkIndexWords()
        {
            // TODO(crhodes) Get this from Config class.  Think about initial directory.

            XElement indexWords = LoadWordsFromXMLFile("IndexWords.xml");

            foreach(XElement word in indexWords.Elements())
            {
                AddinHelper.Common.WriteToDebugWindow(word.Value);

                if (word.Value.Length > 0)
                {
                	ApplyStyleToWords(word.Value, cINDEX_WORD_STYLE);
                }
            }
        }

        private void ReplaceWord(string phrase, string replacementWord, bool indexWordsOnly)
        {
            AddinHelper.Common.WriteToDebugWindow(String.Format("phrase:>{0}<  indexWordsOnly:{1}", phrase, indexWordsOnly));

            Globals.ThisAddIn.Application.Selection.Find.ClearFormatting();

            if (indexWordsOnly)
            {
            	Globals.ThisAddIn.Application.Selection.Find.set_Style(Globals.ThisAddIn.Application.ActiveDocument.Styles[cINDEX_WORD_STYLE]);
            }

            Globals.ThisAddIn.Application.Selection.Find.Replacement.ClearFormatting();

            Microsoft.Office.Interop.Word.Find findObj = Globals.ThisAddIn.Application.Selection.Find;

            findObj.Text = phrase;
            findObj.Replacement.Text = replacementWord;
            findObj.Forward = true;
            findObj.Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue;
            findObj.Format = true;
            findObj.MatchCase = false;
            findObj.MatchWholeWord = false;
            findObj.MatchWildcards = false;
            findObj.MatchSoundsLike = false;
            findObj.MatchAllWordForms = false;

            findObj.Execute(Replace: Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll);
        }

        private void SaveReplacementWordsToXMLFile()
        {
            XElement replacementWords = new XElement("ReplacementWords");
            Dictionary<string, int> possibleWords = new Dictionary<string,int>();

            // Search the file for IndexEntry fields and add them to the dictionary to suppress duplicates

            foreach (Field field in Globals.ThisAddIn.Application.ActiveDocument.Fields)
            {
            	if (field.Type == WdFieldType.wdFieldIndexEntry)
                {
                	string fieldText = field.Code.Text;
                    string indexWord = fieldText.Substring(5,fieldText.Length - 7);
                    int indexLength = indexWord.Length;

                    AddinHelper.Common.WriteToDebugWindow(String.Format("  fieldText:>{0}< indexWord:>{1}<", fieldText, indexWord));

                    try
                    {
                    	possibleWords.Add(indexWord, indexLength);
                    }
                    catch (Exception ex)
                    {
                        AddinHelper.Common.WriteToDebugWindow(String.Format("    Skipping duplicate indexWord:>{0}<", indexWord));
                    }
                }
            }

            // We want to replace the longest phrases first,
            // so get the words out in descending length order

            var orderedWords = from w in possibleWords
                               orderby w.Value descending
                               select w.Key;

            // Add them to the XML

            foreach (string word in orderedWords)
            {
            	replacementWords.Add(new XElement("Word", word));
            }

            // And save the XML to a file

            saveFileDialog1.Filter = "XML Files (*.xml)|*.xml|All files (*.*)|(*.*)";
            saveFileDialog1.FileName = "ReplacementWords.xml";

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
            	using (StreamWriter streamWriter = new StreamWriter(saveFileDialog1.FileName))
                {
                	streamWriter.Write(replacementWords.ToString());
                    streamWriter.Flush();
                }
            }
        }

        private void TagIndexHeadingStyleWords()
        {
            TagIndexWords(cINDEX_HEADING_STYLE);
        }

        private void TagIndexWords(string style)
        {
            Document activeDoc = Globals.ThisAddIn.Application.ActiveDocument;
            Selection selection = Globals.ThisAddIn.Application.Selection;

            selection.HomeKey(Unit: WdUnits.wdStory);
            selection.Find.ClearFormatting();

            // TODO(CRHODES) Verify style exists

            selection.Find.set_Style(Globals.ThisAddIn.Application.ActiveDocument.Styles[style]);
            selection.Find.Text = "";
            selection.Find.Replacement.Text = "";
            selection.Find.Forward = true;
            selection.Find.Wrap = WdFindWrap.wdFindStop;
            selection.Find.Format = true;
            selection.Find.MatchCase = false;
            selection.Find.MatchWholeWord = false;
            selection.Find.MatchWildcards = false;
            selection.Find.MatchSoundsLike = false;
            selection.Find.MatchAllWordForms = false;

            selection.Find.Execute();

            while(selection.Find.Found)
            {
                Range match = selection.Range;

                AddinHelper.Common.WriteToDebugWindow(
                    String.Format("start:{0} end:{1} text:{2} style:{3}",
                                   match.Start, match.End, match.Text, match.get_Style()));

                // Have to do some special magic to handle the Character Style and the Paragraph Style.
                // The Index Marker unfortunately takes on the style of the selection,
                // so, have to find a way of removing the style (character) or moving past it (paragraph).

                if (activeDoc.Styles[style].Type == WdStyleType.wdStyleTypeCharacter)
                {
                	// Mark the found selection as an Index entry
                    
                    activeDoc.Indexes.MarkEntry(Range: selection.Range, Entry: selection.Text, EntryAutoText: selection.Text);

                    // The Index Marker unfortunately takes on the style of the selection, 
                    // so, Search again for the style which finds the marker

                    selection.Find.Execute();

                    // Remove the formatting from the selection so it is not found

                    selection.ClearCharacterAllFormatting();
                }

                if (activeDoc.Styles[style].Type == WdStyleType.wdStyleTypeParagraph)
                {
                	selection.Shrink();
                    activeDoc.Indexes.MarkEntry(Range: selection.Range, Entry: selection.Text, EntryAutoText: selection.Text);
                }

                // Collapse the selection so we continue to search the document

                selection.Collapse(WdCollapseDirection.wdCollapseEnd);

                // And search for the next word to tag (if any)

                selection.Find.Execute();
            }
        }

        private void TagIndexWordStyleWords()
        {
            TagIndexWords(cINDEX_WORD_STYLE);
        }

        private void UpdateIndex()
        {
            try
            {
                Document activeDoc = Globals.ThisAddIn.Application.ActiveDocument;

                if(Globals.ThisAddIn.Application.ActiveDocument.Indexes.Count < 1)
                {
                    MessageBox.Show("Missing Index.  Navigate to where the Index should be located and create using UI");
                }
                else if(Globals.ThisAddIn.Application.ActiveDocument.Indexes.Count > 1)
                {
                    MessageBox.Show("Multiple Indexes detected.  Not supported.  Update manually using UI");
                }
                else
                {
                    activeDoc.Indexes[1].HeadingSeparator = WdHeadingSeparator.wdHeadingSeparatorBlankLine;
                    activeDoc.Indexes[1].Type = WdIndexType.wdIndexIndent;
                    activeDoc.Indexes[1].RightAlignPageNumbers = false;
                    activeDoc.Indexes[1].NumberOfColumns = 2;
                    activeDoc.Indexes[1].IndexLanguage = WdLanguageID.wdEnglishUS;
                    activeDoc.Indexes[1].TabLeader = WdTabLeader.wdTabLeaderDots;
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void ZapReplacementWords()
        {
            XElement replacementWordsXML = LoadWordsFromXMLFile("ReplacementWords.xml");

            foreach(XElement word in replacementWordsXML.Elements())
            {
                if (word.Value.Length > 0)
                {
                	ReplaceWord(word.Value, txtReplacementWord.Text, ckIndexWordsOnly.Checked);
                }
            }
        }

        #endregion
        
    }
}
