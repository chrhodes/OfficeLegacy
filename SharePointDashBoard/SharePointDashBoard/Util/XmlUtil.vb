Option Explicit On

Public Class XmlUtil

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
