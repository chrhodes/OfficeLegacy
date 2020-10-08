Option Explicit On

Imports Excel = Microsoft.Office.Interop.Excel
Imports Word = Microsoft.Office.Interop.Word

Imports Microsoft.Office.DocumentFormat.OpenXml
Imports System.IO.Packaging
Imports System.Xml
Imports System.IO

Imports Microsoft.Office.Core
Imports PacificLife.Life

Public Class WordOutput

    'Private wdApp As Word.Application

    Public Sub CreateOutput( _
        ByRef scRequestList As List(Of STSADMRequestData), _
        ByVal outputFolder As String, _
        ByVal fileNameBase As String _
    )
        PLLog.Trace("Enter", Globals.cPLLog_Category_Name)

        For Each scRequest As STSADMRequestData In scRequestList
            GenerateWordOutputFile(fileNameBase, scRequest)
        Next

        PLLog.Trace("Exit", Globals.cPLLog_Category_Name)
    End Sub


    Private Shared Sub GenerateWordOutputFile(ByVal fileNameBase As String, ByVal scRequest As STSADMRequestData)
        Dim templateFileName As String = Globals.cSharePointSiteCollectionRequestTemplate
        Dim outputFileName = String.Format("{0} {1} - {2}.docx", fileNameBase, scRequest.SharePointFarm, scRequest.Title)

        ' Generate a new output file using a template which already contains bound Content Controls

        System.IO.File.Copy(templateFileName, outputFileName, True)

        ' Open new output file
        ' locate the customXml data package part.
        ' and get a stream writer to access the contents

        Dim pkg As Package = Package.Open(outputFileName, IO.FileMode.Open, IO.FileAccess.ReadWrite)
        Dim uri As Uri = New Uri("/customXml/item1.xml", UriKind.Relative)
        Dim part As PackagePart = pkg.GetPart(uri)
        Dim partWrt As StreamWriter = New StreamWriter(part.GetStream(FileMode.Open, FileAccess.Write))

        ' Create an XmlDocument containing the new data

        Dim doc As Xml.XmlDocument = New Xml.XmlDocument
        doc.LoadXml(scRequest.ToXml)

        ' Add write it back into the package

        doc.Save(partWrt)
        partWrt.Flush()
        partWrt.Close()

        ' Finally, close the new output file containing the updated data

        pkg.Flush()
        pkg.Close()
    End Sub


End Class
