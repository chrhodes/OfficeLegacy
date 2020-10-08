Imports System.Reflection
Imports System.Windows.Forms

Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Word

''' <summary>
''' Word_AddFooter
''' </summary>
''' <remarks>
''' This class can be used in two ways.  If calling this from a commandBar, modify
''' the Private Const as needed and then create an instance of this class in the code
''' that loads the command bars.
''' 
''' If calling this from a Ribbon Event handler, call the ActionNameGoesHere method directly.
''' 
''' Rename the ActionNameGoesHere Method and provide code that does something useful.
''' </remarks>
Public Class Word_AddFooter
    Inherits AddinHelper.AppMethod

#Region "Private Constants and Variables"

    Private Const _MODULE_NAME As String = Common.PROJECT_NAME & "AddFooter"
    Private Const _NAME As String = "AddFooter"
    Private Const _BITMAP_NAME As String = "add footer.bmp"
    Private Const _CAPTION As String = "AddFooter"
    Private Const _TOOL_TIP_TEXT As String = "Click to Add Footer"
    Private Const _DESCRIPTION As String = "AddFooter does ..."

#End Region

#Region "Public Methods"

    Public Sub New(ByRef commandBar As CommandBar, ByRef buttonStyle As MsoButtonStyle)
        MyBase.Name = _NAME
        MyBase.CommandBar = commandBar
        MyBase.EventHandler = AddressOf Action
        MyBase.ButtonStyle = buttonStyle
        MyBase.BitMapName = _BITMAP_NAME
        MyBase.Asmbly = [Assembly].GetExecutingAssembly
        MyBase.Caption = _CAPTION
        MyBase.ToolTipText = _TOOL_TIP_TEXT
        MyBase.Description = _DESCRIPTION

        MyBase.Initialize()
    End Sub

    '-----------------------------------------------------------
    '
    ' Sub Footer_Add()
    '
    ' ToDo:
    '   Display dialog box that allow format choices.
    '   Use FSO and get UNC path to file (if possible)
    '   Save and Restore current view type.
    '------------------------------------------------------------

    Public Shared Sub AddFooter()
        Try
            With Globals.ThisAddIn.Application
                'Debug.Print(.ActiveDocument.Sections.Count())

                For Each documentSection As Section In .ActiveDocument.Sections
                    'Debug.Print(documentSection.Index)
                    If documentSection.Index > 1 Then
                        'documentSection.PageSetup.
                        ' This needs to be smarter about sections.  Linking to previous pulls the formatting, too.
                        ' That makes landscape pages have the same margins as the portrait pages.
                        documentSection.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary).LinkToPrevious = True
                        documentSection.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage).LinkToPrevious = True
                        'Continue For
                    End If

                    If .ActiveWindow.View.SplitSpecial <> WdSpecialPane.wdPaneNone Then
                        .ActiveWindow.Panes(2).Close()
                    End If

                    If .ActiveWindow.ActivePane.View.Type = WdViewType.wdNormalView Or _
                        .ActiveWindow.ActivePane.View.Type = WdViewType.wdOutlineView Then
                        .ActiveWindow.ActivePane.View.Type = WdViewType.wdPrintView
                    End If

                    .ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekCurrentPageFooter
                    'If Selection.HeaderFooter.IsHeader = True Then
                    '    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
                    'Else
                    '    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
                    'End If

                    ' Delete any existing stuff and make the font small.
                    ' Word 2007: Adjust paragraph spacing as "Normal" paragraph style has space
                    ' between lines.  May decide to add our own footer style.

                    With .Selection.ParagraphFormat


                    End With

                    .Selection.WholeStory()
                    .Selection.Range.Delete()
                    .Selection.WholeStory()
                    ' Decided not to do this incase the style is not available.
                    '.Selection.Style = .ActiveDocument.Styles("No Spacing")
                    .Selection.Font.Name = "Arial"
                    .Selection.Font.Size = 5

                    With .Selection.ParagraphFormat
                        .TabStops.ClearAll()
                        .LeftIndent = Globals.ThisAddIn.Application.InchesToPoints(0)
                        .RightIndent = Globals.ThisAddIn.Application.InchesToPoints(0)
                        .SpaceBefore = 0
                        .SpaceBeforeAuto = False
                        .SpaceAfter = 0
                        .SpaceAfterAuto = False
                        .LineSpacingRule = WdLineSpacing.wdLineSpaceSingle
                        '.Alignment = wdAlignParagraphLeft
                        '.WidowControl = True
                        '.KeepWithNext = False
                        '.KeepTogether = False
                        '.PageBreakBefore = False
                        '.NoLineNumber = False
                        '.Hyphenation = True
                        '.FirstLineIndent = InchesToPoints(0)
                        '.OutlineLevel = wdOutlineLevelBodyText
                        '.CharacterUnitLeftIndent = 0
                        '.CharacterUnitRightIndent = 0
                        '.CharacterUnitFirstLineIndent = 0
                        '.LineUnitBefore = 0
                        '.LineUnitAfter = 0
                        '.MirrorIndents = False
                        '.TextboxTightWrap = wdTightNone
                    End With


                    '.Selection.TypeParagraph()

                    '' Add some tabs to help format the fields.

                    '.Selection.ParagraphFormat.TabStops.Add( _
                    '    Position:=.InchesToPoints(0.5), _
                    '    Alignment:=WdTabAlignment.wdAlignTabLeft, _
                    '    Leader:=WdTabLeader.wdTabLeaderSpaces)
                    '.Selection.ParagraphFormat.TabStops.Add( _
                    '    Position:=.InchesToPoints(1.5), _
                    '    Alignment:=WdTabAlignment.wdAlignTabLeft, _
                    '    Leader:=WdTabLeader.wdTabLeaderSpaces)

                    ' Adjust right margin point depending on page orientation.

                    'For Each documentSection As Section In .ActiveDocument.Sections
                    Dim rightMargin As Single
                    Dim leftMargin As Single
                    Dim pageWidth As Single

                    'rightMargin = .ActiveDocument.PageSetup.RightMargin
                    'leftMargin = .ActiveDocument.PageSetup.LeftMargin
                    'pageWidth = .ActiveDocument.PageSetup.PageWidth

                    rightMargin = documentSection.PageSetup.RightMargin
                    leftMargin = documentSection.PageSetup.LeftMargin
                    pageWidth = documentSection.PageSetup.PageWidth

                    'Debug.Print(String.Format("orientation: {0}   pageWidth: {1}   leftMargin: {2}   rightMargin: {3}", documentSection.PageSetup.Orientation, pageWidth, leftMargin, rightMargin))

                    'Next


                    Dim rightMarginPoint As Single

                    'rightMarginPoint = .ActiveDocument.PageSetup.PageWidth _
                    '    - .ActiveDocument.PageSetup.RightMargin _
                    '    - .ActiveDocument.PageSetup.LeftMargin

                    rightMarginPoint = pageWidth - leftMargin - rightMargin

                    'If .ActiveDocument.PageSetup.Orientation = WdOrientation.wdOrientPortrait Then
                    '    .Selection.ParagraphFormat.TabStops.Add( _
                    '        Position:=CInt(rightMarginPoint), _
                    '        Alignment:=WdTabAlignment.wdAlignTabRight, _
                    '        Leader:=WdTabLeader.wdTabLeaderSpaces)
                    'Else
                    '    .Selection.ParagraphFormat.TabStops.Add( _
                    '        Position:=CInt(rightMarginPoint), _
                    '        Alignment:=WdTabAlignment.wdAlignTabRight, _
                    '        Leader:=WdTabLeader.wdTabLeaderSpaces)
                    'End If

                    If documentSection.PageSetup.Orientation = WdOrientation.wdOrientPortrait Then
                        .Selection.ParagraphFormat.TabStops.Add( _
                            Position:=CInt(rightMarginPoint), _
                            Alignment:=WdTabAlignment.wdAlignTabRight, _
                            Leader:=WdTabLeader.wdTabLeaderSpaces)
                    Else
                        .Selection.ParagraphFormat.TabStops.Add( _
                            Position:=CInt(rightMarginPoint), _
                            Alignment:=WdTabAlignment.wdAlignTabRight, _
                            Leader:=WdTabLeader.wdTabLeaderSpaces)
                    End If

                    .Selection.TypeParagraph()

                    .Selection.Fields.Add( _
                        Range:=.Selection.Range, _
                        Type:=WdFieldType.wdFieldEmpty, _
                        Text:="FILENAME \p ", _
                        PreserveFormatting:=True)

                    .Selection.TypeText(Text:=vbTab & "Page ")

                    .Selection.Fields.Add( _
                        Range:=.Selection.Range, _
                        Type:=WdFieldType.wdFieldEmpty, Text:= _
                        "PAGE ", _
                        PreserveFormatting:=True)
                    .Selection.TypeText(Text:=" of ")
                    .Selection.Fields.Add( _
                        Range:=.Selection.Range, _
                        Type:=WdFieldType.wdFieldEmpty, _
                        Text:="NUMPAGES ", _
                        PreserveFormatting:=True)

                    .Selection.TypeParagraph()

                    ' Add some more tabs to help format the fields.

                    .Selection.ParagraphFormat.TabStops.Add( _
                        Position:=.InchesToPoints(0.5), _
                        Alignment:=WdTabAlignment.wdAlignTabLeft, _
                        Leader:=WdTabLeader.wdTabLeaderSpaces)
                    .Selection.ParagraphFormat.TabStops.Add( _
                        Position:=.InchesToPoints(1.5), _
                        Alignment:=WdTabAlignment.wdAlignTabLeft, _
                        Leader:=WdTabLeader.wdTabLeaderSpaces)

                    .Selection.TypeText(Text:="Created:" & vbTab)

                    .Selection.Fields.Add( _
                        Range:=.Selection.Range, _
                        Type:=WdFieldType.wdFieldEmpty, _
                        Text:="CREATEDATE ", _
                        PreserveFormatting:=True)

                    .Selection.TypeText(Text:=vbTab & "Author: ")
                    .Selection.Fields.Add( _
                        Range:=.Selection.Range, _
                        Type:=WdFieldType.wdFieldEmpty, _
                        Text:="AUTHOR ", _
                        PreserveFormatting:=True)

                    .Selection.TypeText(Text:=vbTab & "Title: ")
                    .Selection.Fields.Add( _
                        Range:=.Selection.Range, _
                        Type:=WdFieldType.wdFieldEmpty, _
                        Text:="TITLE ", _
                        PreserveFormatting:=True)

                    .Selection.TypeParagraph()

                    .Selection.TypeText(Text:="Last Saved:" & vbTab)
                    .Selection.Fields.Add( _
                        Range:=.Selection.Range, _
                        Type:=WdFieldType.wdFieldEmpty, _
                        Text:="SAVEDATE ", _
                        PreserveFormatting:=True)

                    .Selection.TypeText(Text:=vbTab & "By: ")

                    .Selection.Fields.Add(Range:=.Selection.Range, Type:=WdFieldType.wdFieldEmpty, _
                        Text:="LASTSAVEDBY ", PreserveFormatting:=True)

                    .Selection.TypeText(Text:=vbTab & "Subject: ")
                    .Selection.Fields.Add( _
                        Range:=.Selection.Range, _
                        Type:=WdFieldType.wdFieldEmpty, _
                        Text:="SUBJECT ", _
                        PreserveFormatting:=True)

                    .Selection.TypeParagraph()
                    .Selection.TypeText(Text:="Printed:" & vbTab)
                    .Selection.Fields.Add(Range:=.Selection.Range, Type:=WdFieldType.wdFieldEmpty, _
                        Text:="PRINTDATE ", PreserveFormatting:=True)
                    .ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekMainDocument
                    '.ActiveWindow.ActivePane.Close()
                Next

            End With    ' g_Word
        Catch ex As Exception
            MsgBox(String.Format("AddFooter: {0}", ex))
        End Try
    End Sub ' Footer_Add

#End Region

#Region "Private Methods"

    Private Sub Action(ByVal Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean)
        Try
            Word_AddFooter.AddFooter()
        Catch ex As Exception
            MessageBox.Show(String.Format("Exception: {0}.{1}() - {2}",
                         System.Reflection.Assembly.GetExecutingAssembly().FullName,
                         System.Reflection.MethodInfo.GetCurrentMethod().Name,
                         ex.ToString()
                         ))
        End Try
    End Sub

#End Region

End Class
