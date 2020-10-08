Imports Microsoft.Office.DocumentFormat.OpenXml
Imports System.IO.Packaging
Imports System.Xml
Imports System.IO
Imports System.Collections

Imports System.Diagnostics

Public Class STSADMRequest
    Private _listOfRequests As List(Of STSADMRequestData)

    Public Property ListOfRequests() As List(Of STSADMRequestData)
        Get
            If _listOfRequests Is Nothing Then
                _listOfRequests = New List(Of STSADMRequestData)
            End If

            Return _listOfRequests
        End Get
        Set(ByVal Value As List(Of STSADMRequestData))
            _listOfRequests = Value
        End Set
    End Property

    Private Sub Sheet1_Startup(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Startup

    End Sub

    Private Sub Sheet1_Shutdown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shutdown

    End Sub

    Private Sub PopulateListOfRequests(ByVal ws As Excel.Worksheet, ByVal startingRow As Integer, ByVal endingRow As Integer)
        For i As Integer = startingRow To endingRow
            Dim scRequest As New STSADMRequestData
            scRequest.PopulateFromExcelRange(ws.Cells(i, 2))

            ListOfRequests.Add(scRequest)
        Next
    End Sub

    Private Sub btnCreateSTSADMCommands_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateSTSADMCommands.Click
        Dim flOutput As FileOutput = New FileOutput
        'Dim ppOutput As PowerPointOutput = New PowerPointOutput
        'Dim wdOutput As WordOutput = New WordOutput

        Dim ws As Excel.Worksheet = Application.ActiveSheet
        Dim startingRow As Integer = Range(Globals.cSC_StartingRow_Cell).Value
        Dim endingRow As Integer = Range(Globals.cSC_EndingRow_Cell).Value
        Dim outputFolder As String = Range(Globals.cSC_STSADMOutput_Folder_Cell).Value
        Dim outputFileName As String = Range(Globals.cSC_STSADMOutput_FileName_Cell).Value

        'Dim listOfRequests As New List(Of SiteCollectionRequestData)

        If ListOfRequests.Count <> 0 Then
            ListOfRequests = Nothing
        End If

        PopulateListOfRequests(ws, startingRow, endingRow)
        flOutput.CreateOutput(ListOfRequests, outputFolder, outputFileName)

        'ppOutput.CreateOutput(ws, startingRow, endingRow, fileName)
        'wdOutput.CreateOutput(ws, startingRow, endingRow, fileName)

    End Sub

    Private Sub btnCreateSiteCollectionRequests_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateSiteCollectionRequests.Click
        'Dim flOutput As FileOutput = New FileOutput
        'Dim ppOutput As PowerPointOutput = New PowerPointOutput
        Dim wdOutput As WordOutput = New WordOutput

        Dim ws As Excel.Worksheet = Application.ActiveSheet
        Dim startingRow As Integer = Range(Globals.cSC_StartingRow_Cell).Value
        Dim endingRow As Integer = Range(Globals.cSC_EndingRow_Cell).Value
        Dim outputFolder As String = Range(Globals.cSC_WordOutput_Folder_Cell).Value
        Dim fileNameBase As String = Range(Globals.cSC_WordOutput_FileNameBase_Cell).Value

        'flOutput.CreateOutput(ws, startingRow, endingRow, fileName)
        'ppOutput.CreateOutput(ws, startingRow, endingRow, fileName)

        If ListOfRequests.Count = 0 Then
            PopulateListOfRequests(ws, startingRow, endingRow)
        End If

        wdOutput.CreateOutput(ListOfRequests, outputFolder, fileNameBase)
    End Sub

End Class

' Class to encapsulate the Excel data row

'Public Class SiteCollectionRequests : Implements IEnumerable
'    Dim requests() As SharePointContainerRequestData
'    Dim count As Integer

'    Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
'        Return requests.GetEnumerator
'    End Function
'End Class
