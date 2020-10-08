Option Explicit On

Public Class XmlUtil
    Public Shared Function ProcessQuestionsXML() As Excel.Worksheet
        Dim ws As Excel.Worksheet
        Dim targetWorksheet As Excel.Worksheet = Nothing
        'Dim r As Xml.XmlReader
        Dim cellXml As Xml.XmlDocument = New Xml.XmlDocument
        Dim xn As Xml.XmlNode
        Dim ac As Xml.XmlAttributeCollection
        Dim xe As Xml.XmlElement
        Dim xeq As Xml.XmlText

        Dim inputRow As Integer = 2
        Dim outputStartRow As Integer = Globals.cHeaderID_RowShort + 1
        Dim outputEndRow As Integer = outputStartRow

        Dim questionName As String
        Dim questionText As String

        ws = Globals.ThisAddIn.Application.ActiveSheet

        While (ws.Cells(inputRow, 2).Value <> "")
            cellXml.LoadXml(ws.Cells(inputRow, 2).Value)
            Debug.WriteLine(cellXml.DocumentElement.Name, "DocumentElement.Name")

            Select Case cellXml.DocumentElement.Name
                Case "WebProperties"
                    ac = cellXml.DocumentElement.Attributes
                    xn = ac.GetNamedItem("title")
                    Select Case xn.Value
                        Case "Partner Survey"
                            targetWorksheet = SurveyQuestionsWorkSheet.CreateWorkSheet(Globals.cSN_PartnerSurveyQuestions)
                            targetWorksheet.Unprotect()

                        Case "Business Survey"
                            targetWorksheet = SurveyQuestionsWorkSheet.CreateWorkSheet(Globals.cSN_BusinessSurveyQuestions)
                            targetWorksheet.Unprotect()

                        Case "IT Survey"
                            targetWorksheet = SurveyQuestionsWorkSheet.CreateWorkSheet(Globals.cSN_ITSurveyQuestions)
                            targetWorksheet.Unprotect()

                        Case "Help Desk Survey"
                            targetWorksheet = SurveyQuestionsWorkSheet.CreateWorkSheet(Globals.cSN_HelpDeskSurveyQuestions)
                            targetWorksheet.Unprotect()

                        Case Else
                            MessageBox.Show("Unknown tile attribute value in <WebProperties> element (" & xn.Value.ToString, "ProcessQuestionsXML()")
                    End Select

                Case "Question"
                    ac = cellXml.DocumentElement.Attributes
                    DisplayAttributes(ac)
                    xn = ac.GetNamedItem("name")
                    xe = cellXml.FirstChild
                    xeq = xe.FirstChild
                    questionName = xn.Value
                    questionText = xeq.Value
                    System.Diagnostics.Debug.WriteLine("QuestionName:>" & questionName & "< questionText:>" & questionText & "<")
                    targetWorksheet.Cells(outputEndRow, 1).Value = questionName
                    targetWorksheet.Cells(outputEndRow, 2).Value = questionText
                    outputEndRow += 1
            End Select

            inputRow += 1
        End While

        ' TODO: Get rid of 1, 3, 5, ... magic numbers in ADDRESS(...) formula.

        With targetWorksheet
            ' Put the appropriate cell ranges for later use.  Add the supporting formulas
            ' to reference the Questions and any data we can pre-populate.

            ' Questions

            .Range(Globals.cSQ_QuestionsLocationCell).Offset(0, 1).Formula = _
                "=ADDRESS(" & outputStartRow & ",1,,TRUE) & "":"" & ADDRESS(" & outputEndRow - 1 & ",2,TRUE)"

            ' Row Widths

            .Range(Globals.cSQ_ColumnWidthsLocationCell).Offset(0, 1).Formula = _
                "=ADDRESS(" & outputStartRow & ",3,,TRUE) & "":"" & ADDRESS(" & outputEndRow - 1 & ",4,TRUE)"

            .Range(.Cells(outputStartRow, 3), .Cells(outputEndRow - 1, 3)).FormulaR1C1 = _
                "=RC[-2]"

            ' TODO: Make this smarter about being on a question or a follow-up question.
            ' Questions are 5, follow-ups are 100 wide.

            '.Range(.Cells(outputStartRow, 4), .Cells(outputEndRow - 1, 4)).Value = "5"
            .Range(.Cells(outputStartRow, 4), .Cells(outputEndRow - 1, 4)).FormulaR1C1 = _
                "=IF(ISERR(FIND(""b"",RC[-3])),5,110)"

            ' Statistics

            .Range(Globals.cSQ_StatisticsLocationCell).Offset(0, 1).Formula = _
                "=ADDRESS(" & outputStartRow & ",5,,TRUE) & "":"" & ADDRESS(" & outputEndRow - 1 & ",6,TRUE)"

            .Range(.Cells(outputStartRow, 5), .Cells(outputEndRow - 1, 5)).FormulaR1C1 = _
                "=RC[-4]"

            .Range(.Cells(outputStartRow, 6), .Cells(outputEndRow - 1, 6)).FormulaR1C1 = _
                "=IF(ISERR(FIND(""b"",RC[-5])),1,0)"

            ' Primary Questions

            .Range(Globals.cSQ_PrimaryQuestionsLocationCell).Offset(0, 1).Formula = _
                "=ADDRESS(" & outputStartRow & ",7,,TRUE) & "":"" & ADDRESS(" & outputEndRow - 1 & ",8,TRUE)"

            .Range(.Cells(outputStartRow, 7), .Cells(outputEndRow - 1, 7)).FormulaR1C1 = _
                "=RC[-6]"

            ' TODO: Make this smarter by looking at Questions.  If "b" does not appear
            ' in question name, default to Primary.

            .Range(.Cells(outputStartRow, 8), .Cells(outputEndRow - 1, 8)).FormulaR1C1 = _
                "=IF(ISERR(FIND(""b"",RC[-7])),1,0)"

            ' Follow-Up Questions

            .Range(Globals.cSQ_FollowUpQuestionsLocationCell).Offset(0, 1).Formula = _
                "=ADDRESS(" & outputStartRow & ",9,,TRUE) & "":"" & ADDRESS(" & outputEndRow - 1 & ",10,TRUE)"

            .Range(.Cells(outputStartRow, 9), .Cells(outputEndRow - 1, 9)).FormulaR1C1 = _
                "=RC[-8]"

            ' Follow-Up Questions and Primary Questions are mutually exclusive

            .Range(.Cells(outputStartRow, 10), .Cells(outputEndRow - 1, 10)).FormulaR1C1 = _
                "=IF(RC[-2],0,1)"

            ' Now sort the questions so LOOKUP will work correctly

            Dim sortRange As Excel.Range

            sortRange = .Range(.Cells(outputStartRow, 1), .Cells(outputEndRow, 2))

            ' TODO: Fix the "A11" magic number.

            With .Sort
                .SortFields.Clear()
                .SortFields.Add( _
                    Key:=targetWorksheet.Range("A11"), SortOn:=Excel.XlSortOn.xlSortOnValues, _
                    Order:=Excel.XlSortOrder.xlAscending, DataOption:=Excel.XlSortDataOption.xlSortNormal)
                .SetRange(sortRange)
                .Header = Excel.XlYesNoGuess.xlNo
                .MatchCase = False
                .Orientation = Excel.Constants.xlTopToBottom
                .SortMethod = Excel.XlSortMethod.xlPinYin
                .Apply()
            End With

            .Range("C11").Select()
            Globals.ThisAddIn.Application.ActiveWindow.FreezePanes = True
        End With

        Return targetWorksheet
    End Function

    Private Sub btnProcessXML_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim ws As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet
        'Dim r As Xml.XmlReader
        Dim xd As Xml.XmlDocument = New Xml.XmlDocument
        'Dim xn As Xml.XmlNode
        'Dim ac As Xml.XmlAttributeCollection
        Dim xe As Xml.XmlElement
        Dim xeq As Xml.XmlText

        'Dim rng As Excel.Range
        Dim xmlCell As String
        'Dim question As String

        xmlCell = Globals.ThisAddIn.Application.ActiveCell.Value
        xd.LoadXml(xmlCell)
        'Debug.WriteLine(xd.DocumentElement.Name, "DocumentElement.Name")
        'Debug.WriteLine(xd.HasChildNodes & ":" & xd.ChildNodes.Count, "Has Child Nodes")
        'Debug.WriteLine(xd.Value, "xd.Value")

        If "Question" = xd.DocumentElement.Name Then
            xe = xd.FirstChild
            xeq = xe.FirstChild
            Debug.WriteLine(xeq.Value, "xeq.Value")
        End If

        For Each cn As Xml.XmlNode In xd.ChildNodes
            DisplayChildNode(cn)
        Next

        Return

    End Sub

    Public Shared Sub DisplayAttributes(ByVal ac As Xml.XmlAttributeCollection)
        For Each at As Xml.XmlAttribute In ac
            System.Diagnostics.Debug.WriteLine(at.Name & " : " & at.Value, "at.Value")
        Next
    End Sub

    Public Sub DisplayChildNode(ByVal cn As Xml.XmlNode)
        Debug.WriteLine(cn.Name, "cn.Name")
        Debug.WriteLine(cn.Value, "cn.Value")

        Debug.WriteLine(cn.NodeType, "cn.NodeType")

        Select Case cn.NodeType
            Case Xml.XmlNodeType.Attribute

            Case Xml.XmlNodeType.CDATA

            Case Xml.XmlNodeType.Comment

            Case Xml.XmlNodeType.Document

            Case Xml.XmlNodeType.DocumentFragment

            Case Xml.XmlNodeType.DocumentType

            Case Xml.XmlNodeType.Element
                Debug.WriteLine(cn.Attributes.Count, "cn.Attributes.Count")

                For Each at As Xml.XmlAttribute In cn.Attributes
                    System.Diagnostics.Debug.WriteLine(at.Name & " : " & at.Value, "at.Value")
                Next

            Case Xml.XmlNodeType.EndElement

            Case Xml.XmlNodeType.EndEntity

            Case Xml.XmlNodeType.Entity

            Case Xml.XmlNodeType.EntityReference

            Case Xml.XmlNodeType.None

            Case Xml.XmlNodeType.Notation

            Case Xml.XmlNodeType.ProcessingInstruction

            Case Xml.XmlNodeType.SignificantWhitespace

            Case Xml.XmlNodeType.Text

            Case Xml.XmlNodeType.Whitespace

            Case Xml.XmlNodeType.XmlDeclaration

        End Select

        Debug.WriteLine(cn.HasChildNodes & ":" & cn.ChildNodes.Count, "Has Child Nodes")

        For Each cn2 As Xml.XmlNode In cn
            DisplayChildNode(cn2)
        Next
    End Sub
End Class
