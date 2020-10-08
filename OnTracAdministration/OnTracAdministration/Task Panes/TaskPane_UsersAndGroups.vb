Imports System.ComponentModel
Imports System.Reflection
Imports System.Runtime

Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel

Public Class TaskPane_UsersAndGroups
    'Private _teams As Data.DataSet

    Private _inputFilePath As String = Globals.cDEFAULT_ONTIMEDATA_FOLDER

    Private _paneWidth As Integer = 300

    Public Property PaneWidth() As Integer
        Get
            Return _paneWidth
        End Get
        Set(ByVal Value As Integer)
            _paneWidth = Value
        End Set
    End Property

    Private Sub btnAddOnTimeDataSheets_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Dim rng As Excel.Range
        'Dim inputPathAndFileName As String
        'Dim inputSheetName As String
        'Dim inputWorkbook As Excel.Workbook
        'Dim currentWorkbook As Excel.Workbook
        'Dim errorMessage As String = ""
        'Dim inputWorkSheet As Excel.Worksheet
        'Dim currentSheet As Excel.Worksheet
        'Dim dataSheet As Excel.Worksheet
        'Dim dataSheetName As String
        ''Dim lastRow2 As Integer
        ''Dim lastCol2 As Integer
        'Dim lastRow As Integer
        'Dim lastCol As Integer
        ''Dim startRow As Integer
        ''Dim startColumn As Integer
        ''Dim row As Integer

        '' TODO: This should probably mostly move to OnTimeDataWorkSheet

        'If Not OnTimeDataWorkSheet.ValidateInputSheet() Then
        '    MsgBox("Invalid Worksheet.")
        'Else
        '    With Globals.ThisAddIn.Application
        '        Util.ScreenUpdatesOff()
        '        currentSheet = .ActiveSheet
        '        currentWorkbook = .ActiveWorkbook

        '        For Each rng In CType(.Selection, Microsoft.Office.Interop.Excel.Range)
        '            inputPathAndFileName = rng.Value
        '            inputSheetName = rng.Offset(0, Globals.cOTD_SheetName_Offset).Value

        '            If inputPathAndFileName <> "" Then
        '                If Not OnTimeDataWorkSheet.OpenInputFileAndVerifyDataLayout(inputPathAndFileName, inputSheetName, errorMessage) Then
        '                    MsgBox("Invalid input file: " & errorMessage)
        '                Else
        '                    ' We have opened an file containing valid On-Time Data.  Extract what we need.
        '                    ' Make as few assumptions as possible.  The sheet may be poorly formatted or contain
        '                    ' merged cells.  Assume the worst and just grab everything that might contain data.
        '                    ' We will clean up the file later in CleanUpDataSheet()

        '                    inputWorkbook = .ActiveWorkbook
        '                    inputWorkSheet = .ActiveWorkbook.Sheets(inputSheetName)

        '                    lastRow = inputWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row
        '                    lastCol = inputWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column
        '                    ' We cannot safely use these routines as they don't behave well with merged cells
        '                    'lastRow2 = Util.FindLastRow(.ActiveSheet.Range("A1"))
        '                    'lastCol2 = Util.FindLastColumn(.ActiveSheet.Range("A1"))

        '                    inputWorkSheet.Range(inputWorkSheet.Cells(1, 1), inputWorkSheet.Cells(lastRow, lastCol)).Copy()

        '                    ' Now return to the Scorecard Workbook, add a sheet for the data, and give it a suitable name.

        '                    currentWorkbook.Activate()
        '                    dataSheet = .Sheets.Add()
        '                    dataSheetName = rng.Offset(0, Globals.cOTD_TeamName_Offset).Value & Globals.cOTD_MetricName
        '                    .ActiveSheet.Name = dataSheetName

        '                    '' Some people have added fancy formulas to their reports.  To ensure that we
        '                    '' get just the values so we can safely delete rows and not have information
        '                    '' changing just paste in the values and any number formats.

        '                    '.Range(Globals.cRawDataCell).PasteSpecial( _
        '                    '    Paste:=XlPasteType.xlPasteValuesAndNumberFormats, _
        '                    '    Operation:=XlPasteSpecialOperation.xlPasteSpecialOperationNone, _
        '                    '    SkipBlanks:=False, Transpose:=False)

        '                    '' Unfortunately pasting the values does not bring the formatting along.
        '                    '' This makes is hard to show the data to someone and have them recognize it
        '                    '' so, add the formatting back on top of the data

        '                    '.Range(Globals.cRawDataCell).PasteSpecial( _
        '                    '        Paste:=XlPasteType.xlPasteFormats, _
        '                    '        Operation:=XlPasteSpecialOperation.xlPasteSpecialOperationNone, _
        '                    '        SkipBlanks:=False, Transpose:=False)

        '                    ' Ah, man.  This just sucks.  If a cell has formatting applied to just part of 
        '                    ' it's contents, when pasting the format on top of the values it only applies
        '                    ' the first formatting it sees.  So, if for example the cell has two dates
        '                    ' and the first is stiken but the second is not, pasting the format makes
        '                    ' the whole cell have striken values.  Groan.

        '                    ' So, looks like we are back to just pasting everything.  Will have to manually
        '                    ' fix the ID column if needed.

        '                    .Range(Globals.cRawDataCell).PasteSpecial( _
        '                            Paste:=XlPasteType.xlPasteAll, _
        '                            Operation:=XlPasteSpecialOperation.xlPasteSpecialOperationNone, _
        '                            SkipBlanks:=False, Transpose:=False)

        '                    OnTimeDataWorkSheet.CleanUpDataSheet()

        '                    ' Might want to expose this from TaskPane

        '                    OnTimeDataWorkSheet.AddFormulas( _
        '                        inputPathAndFileName, inputSheetName, _
        '                        rng.Offset(0, Globals.cOTD_TeamName_Offset).Value, _
        '                        rng.Offset(0, Globals.cOTD_Manager_Offset).Value, _
        '                        rng.Offset(0, Globals.cOTD_Extension_Offset).Value)

        '                    ' Quietly close the input workbook.
        '                    ' TODO: Make this a Util Function

        '                    .DisplayAlerts = False
        '                    inputWorkbook.Close(False)
        '                    .DisplayAlerts = True

        '                    ' Now record what sheet just got added to the Input Worksheet
        '                    ' along with a link to the sheet.

        '                    currentSheet.Unprotect()
        '                    rng.Offset(0, Globals.cOTD_DataSheet_Offset).Value = dataSheetName
        '                    rng.Offset(0, Globals.cOTD_DataSheet_Offset).Hyperlinks.Add( _
        '                                Anchor:=rng.Offset(0, Globals.cOTD_DataSheet_Offset), _
        '                                Address:="", _
        '                                SubAddress:="'" & dataSheetName & "'!A1", _
        '                                TextToDisplay:=dataSheetName)

        '                    ' Finally add the On-Time Percentage data.

        '                    rng.Offset(0, Globals.cOTD_WeightedScheduledOnTimePercentage_Offset).Value = _
        '                        "='" & dataSheetName & "'!" & Globals.cOTD_WeightedScheduledOnTimePercentage_Cell

        '                    rng.Offset(0, Globals.cOTD_WeightedActualOnTimePercentage_Offset).Value = _
        '                        "='" & dataSheetName & "'!" & Globals.cOTD_WeightedActualOnTimePercentage_Cell

        '                    rng.Offset(0, Globals.cOTD_OnTimePercentage_Offset).Value = _
        '                        "='" & dataSheetName & "'!" & Globals.cOTD_OnTimePercentage_Cell

        '                    rng.Offset(0, Globals.cOTD_NumberReleases_Offset).Value = _
        '                        "='" & dataSheetName & "'!" & Globals.cOTD_NumberReleases_Cell

        '                    currentSheet.Protect()
        '                End If
        '            End If
        '        Next rng
        '    End With
        'End If

        'Util.ScreenUpdatesOn()
        'inputWorkbook = Nothing
    End Sub

    Private Sub btnAddOnTimeDataToPowerPoint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Dim rng As Excel.Range
        'Dim ws As Excel.Worksheet

        'Try
        '    With Globals.ThisAddIn.Application
        '        For Each rng In CType(.Selection, Microsoft.Office.Interop.Excel.Range)
        '            ' Verify we have a file name and a sheet name

        '            If "" <> rng.Value And "" <> rng.Offset(0, Globals.cOTD_DataSheet_Offset).Value Then
        '                ' If yes, select the sheet and add the, hopefully existing, chart
        '                ' to PowerPoint.

        '                ws = .Sheets(rng.Offset(0, Globals.cOTD_DataSheet_Offset).Value)
        '                'PowerPointIntegration.AddOnTimeDataToPowerPoint(ws)
        '                ws = Nothing
        '            End If
        '        Next rng
        '    End With
        'Catch ex As Exception
        '    MessageBox.Show("Exception: btnAddOnTimeDataToPowerPoint_Click()")
        'End Try
    End Sub

    Private Sub btnBrowseForOnTimeDataFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Dim inputFile As String = Util.GetFile(_inputFilePath, "Select File Containing On-Time Data")

        'If inputFile.Length > 0 Then
        '    ' Save the folder in case the user browsed to a new location.
        '    _inputFilePath = System.IO.Path.GetDirectoryName(inputFile)

        '    ' Save the file on the worksheet.
        '    Globals.ThisAddIn.Application.ActiveCell.Value = inputFile
        'End If

    End Sub

    Private Sub TaskPane_UsersAndGroups_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ''' Ensure that any config data we need is available.  Ok to call multiple times.
        '''Config.IntializeApplication()

        ''For Each dataTable As Data.DataTable In Config.ConfigInfo.Tables
        ''    'Debug.Print(dataTable.TableName)

        ''    'For Each dataColumn As Data.DataColumn In dataTable.Columns
        ''    '    Debug.Print(dataColumn.ColumnName)
        ''    'Next

        ''    Select Case dataTable.TableName
        ''        Case "team"
        ''            For Each dataRow As Data.DataRow In dataTable.Rows
        ''                Me.clbOnTimeTeams.Items.Add(dataRow.Item("name")).ToString()
        ''                Me.cbOnTimeTeams.Items.Add(dataRow.Item("name")).ToString()
        ''                'Debug.Print(dataRow.Item("name").ToString())
        ''                'Debug.Print(dataRow.Item("id").ToString())
        ''                'Debug.Print(dataRow.Item("team_Id").ToString())
        ''            Next

        ''            'Case "manager"
        ''            '    For Each dataRow As Data.DataRow In dataTable.Rows
        ''            '        Debug.Print(dataRow.Item("manager_Text").ToString())
        ''            '        Debug.Print(dataRow.Item("ext").ToString())
        ''            '        Debug.Print(dataRow.Item("team_Id").ToString())
        ''            '    Next

        ''    End Select
        ''Next
        'Dim rng As Microsoft.Office.Interop.Excel.Range
        'Dim inputPathAndFileName As String
        'Dim inputSheetName As String
        'Dim inputWorkbook As Excel.Workbook
        'Dim errorMessage As String = ""


        'Dim OnTracService As New ontrac1.Webs()

        'Dim webService As String = "http://ontrac/_vti_bin/webs.asmx"

        '' Use credentials of logged on user

        'OnTracService.Credentials = System.Net.CredentialCache.DefaultCredentials

        '' or use specific credentials

        ''Dim cache As System.Net.CredentialCache = New System.Net.CredentialCache()
        ''cache.Add(New Uri(webService), "NTLM", New System.Net.NetworkCredential("pspappca", "Production2007", "PACIFICMUTUAL"))

        ''OnTracService.Credentials = cache

        'OnTracService.Url = webService

        'Util.ScreenUpdatesOff()

        'Dim node As System.Xml.XmlNode = OnTracService.GetAllSubWebCollection()

        ''Dim innerNode As System.Xml.XmlNode = node.FirstChild() ' 

        'Dim i As Integer = 6

        'For Each webNode As System.Xml.XmlNode In node
        '    'ws.Cells(i, 1).Value = webNode.Attributes("Title").Value
        '    'ws.Cells(i, 2).Value = webNode.Attributes("Url").Value
        '    i += 1
        'Next

        'Util.ScreenUpdatesOn()
    End Sub

    Private Sub cbOnTimeTeams_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbOnTimeTeams.SelectedIndexChanged
        'Dim ws As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet

        '''' TODO: Verify we have the right worksheet open before splatting things down.

        ''ws.Range(Globals.cOTD_TeamNameCell).Value = cbOnTimeTeams.SelectedItem.ToString()
    End Sub

    Private Sub btnDeleteOnTimeDataSheets_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Dim rng As Excel.Range

        'With Globals.ThisAddIn.Application
        '    For Each rng In CType(.Selection, Microsoft.Office.Interop.Excel.Range)
        '        .Sheets(rng.Offset(0, 3).Value).Delete()
        '        Util.ProtectSheet(.ActiveSheet, False)
        '        ' TODO: Get rid of magic numbers.
        '        rng.Offset(0, 2).Value = ""
        '        rng.Offset(0, 3).Clear()
        '        rng.Offset(0, -3).Clear()
        '        Util.ProtectSheet(.ActiveSheet, True)
        '    Next rng
        'End With
    End Sub

    Private Sub btnGetGroupCollectionFromSite_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetGroupCollectionFromSite.Click
        Dim rng As Microsoft.Office.Interop.Excel.Range
        Dim inputPathAndFileName As String
        Dim inputSheetName As String
        Dim inputWorkbook As Excel.Workbook
        Dim errorMessage As String = ""
        Dim ws As Excel.Worksheet = Util.NewWorksheet("Groups from Site")
        ws.Cells.Clear()

        Dim OnTracService As New ontrac.UserGroup()

        Dim webServiceURL As String

        If txtURL.TextLength > 0 Then
            webServiceURL = txtURL.Text & "/_vti_bin/usergroup.asmx"
        Else
            MsgBox("Must provide URL")
            Return
        End If

        ' Use credentials of logged on user

        OnTracService.Credentials = System.Net.CredentialCache.DefaultCredentials

        ' or use specific credentials

        'Dim cache As System.Net.CredentialCache = New System.Net.CredentialCache()
        'cache.Add(New Uri(webService), "NTLM", New System.Net.NetworkCredential("pspappca", "Production2007", "PACIFICMUTUAL"))

        'OnTracService.Credentials = cache

        OnTracService.Url = webServiceURL

        Util.ScreenUpdatesOff()

        ws.Name = "Users and Groups from Site"

        Util.AddColumnToSheet(ws, 1, 10, True, 5, "ID")
        Util.AddColumnToSheet(ws, 2, 40, True, 5, "Name")
        Util.AddColumnToSheet(ws, 3, 50, True, 5, "Description")
        Util.AddColumnToSheet(ws, 4, 10, True, 5, "OwnerID")
        Util.AddColumnToSheet(ws, 5, 15, True, 5, "OwnerIsUser")

        Try
            Dim node As System.Xml.XmlNode = OnTracService.GetGroupCollectionFromSite()
            Dim innerNode As System.Xml.XmlNode = node.FirstChild()

            Dim i As Integer = 6

            For Each groupNode As System.Xml.XmlNode In innerNode
                ws.Cells(i, 1).Value = groupNode.Attributes("ID").Value
                ws.Cells(i, 2).Value = groupNode.Attributes("Name").Value
                ws.Cells(i, 3).Value = groupNode.Attributes("Description").Value
                ws.Cells(i, 4).Value = groupNode.Attributes("OwnerID").Value
                ws.Cells(i, 5).Value = groupNode.Attributes("OwnerIsUser").Value
                i += 1
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        Util.ScreenUpdatesOn()

    End Sub

    Private Sub btnGetGroupCollectionFromWeb_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetGroupCollectionFromWeb.Click
        Dim rng As Microsoft.Office.Interop.Excel.Range
        Dim inputPathAndFileName As String
        Dim inputSheetName As String
        Dim inputWorkbook As Excel.Workbook
        Dim errorMessage As String = ""
        Dim ws As Excel.Worksheet = Util.NewWorksheet("Groups from Web")
        ws.Cells.Clear()

        Dim OnTracService As New ontrac.UserGroup()

        Dim webServiceURL As String

        If txtURL.TextLength > 0 Then
            webServiceURL = txtURL.Text & "/_vti_bin/usergroup.asmx"
        Else
            MsgBox("Must provide URL")
            Return
        End If

        ' Use credentials of logged on user

        OnTracService.Credentials = System.Net.CredentialCache.DefaultCredentials

        ' or use specific credentials

        'Dim cache As System.Net.CredentialCache = New System.Net.CredentialCache()
        'cache.Add(New Uri(webService), "NTLM", New System.Net.NetworkCredential("pspappca", "Production2007", "PACIFICMUTUAL"))

        'OnTracService.Credentials = cache

        OnTracService.Url = webServiceURL

        Util.ScreenUpdatesOff()

        Util.AddColumnToSheet(ws, 1, 10, True, 5, "ID")
        Util.AddColumnToSheet(ws, 2, 40, True, 5, "Name")
        Util.AddColumnToSheet(ws, 3, 50, True, 5, "Description")
        Util.AddColumnToSheet(ws, 4, 10, True, 5, "OwnerID")
        Util.AddColumnToSheet(ws, 5, 15, True, 5, "OwnerIsUser")

        Try
            Dim node As System.Xml.XmlNode = OnTracService.GetGroupCollectionFromWeb()
            Dim innerNode As System.Xml.XmlNode = node.FirstChild()

            Dim i As Integer = 6

            For Each groupNode As System.Xml.XmlNode In innerNode
                ws.Cells(i, 1).Value = groupNode.Attributes("ID").Value
                ws.Cells(i, 2).Value = groupNode.Attributes("Name").Value
                ws.Cells(i, 3).Value = groupNode.Attributes("Description").Value
                ws.Cells(i, 4).Value = groupNode.Attributes("OwnerID").Value
                ws.Cells(i, 5).Value = groupNode.Attributes("OwnerIsUser").Value
                i += 1
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        Util.ScreenUpdatesOn()

    End Sub

    Private Sub btnGetUserCollectionFromSite_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetUserCollectionFromSite.Click
        Dim rng As Microsoft.Office.Interop.Excel.Range
        Dim inputPathAndFileName As String
        Dim inputSheetName As String
        Dim inputWorkbook As Excel.Workbook
        Dim errorMessage As String = ""
        Dim ws As Excel.Worksheet = Util.NewWorksheet("Users from Site")
        ws.Cells.Clear()

        Dim OnTracService As New ontrac.UserGroup()

        Dim webServiceURL As String

        If txtURL.TextLength > 0 Then
            webServiceURL = txtURL.Text & "/_vti_bin/usergroup.asmx"
        Else
            MsgBox("Must provide URL")
            Return
        End If

        ' Use credentials of logged on user

        OnTracService.Credentials = System.Net.CredentialCache.DefaultCredentials

        ' or use specific credentials

        'Dim cache As System.Net.CredentialCache = New System.Net.CredentialCache()
        'cache.Add(New Uri(webService), "NTLM", New System.Net.NetworkCredential("pspappca", "Production2007", "PACIFICMUTUAL"))

        'OnTracService.Credentials = cache

        OnTracService.Url = webServiceURL

        Util.ScreenUpdatesOff()

        Util.AddColumnToSheet(ws, 1, 10, True, 5, "ID")
        Util.AddColumnToSheet(ws, 2, 45, True, 5, "Sid")
        Util.AddColumnToSheet(ws, 3, 40, True, 5, "Name")
        Util.AddColumnToSheet(ws, 4, 40, True, 5, "Email")
        Util.AddColumnToSheet(ws, 5, 15, True, 5, "Notes")
        Util.AddColumnToSheet(ws, 5, 15, True, 5, "IsSiteAdmin")
        Util.AddColumnToSheet(ws, 5, 15, True, 5, "IsDomainGroup")

        Try
            Dim node As System.Xml.XmlNode = OnTracService.GetUserCollectionFromSite()
            Dim innerNode As System.Xml.XmlNode = node.FirstChild()

            Dim i As Integer = 6

            For Each groupNode As System.Xml.XmlNode In innerNode
                ws.Cells(i, 1).Value = groupNode.Attributes("ID").Value
                ws.Cells(i, 2).Value = groupNode.Attributes("Sid").Value
                ws.Cells(i, 3).Value = groupNode.Attributes("Name").Value
                ws.Cells(i, 4).Value = groupNode.Attributes("Email").Value
                ws.Cells(i, 5).Value = groupNode.Attributes("Notes").Value
                ws.Cells(i, 5).Value = groupNode.Attributes("IsSiteAdmin").Value
                ws.Cells(i, 5).Value = groupNode.Attributes("IsDomainGroup").Value

                i += 1
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        Util.ScreenUpdatesOn()
    End Sub

    Private Sub btnGetUserCollectionFromWeb_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetUserCollectionFromWeb.Click
        Dim rng As Microsoft.Office.Interop.Excel.Range
        Dim inputPathAndFileName As String
        Dim inputSheetName As String
        Dim inputWorkbook As Excel.Workbook
        Dim errorMessage As String = ""
        Dim ws As Excel.Worksheet = Util.NewWorksheet("Users from Web")
        ws.Cells.Clear()

        Dim OnTracService As New ontrac.UserGroup()

        Dim webServiceURL As String

        If txtURL.TextLength > 0 Then
            webServiceURL = txtURL.Text & "/_vti_bin/usergroup.asmx"
        Else
            MsgBox("Must provide URL")
            Return
        End If

        ' Use credentials of logged on user

        OnTracService.Credentials = System.Net.CredentialCache.DefaultCredentials

        ' or use specific credentials

        'Dim cache As System.Net.CredentialCache = New System.Net.CredentialCache()
        'cache.Add(New Uri(webService), "NTLM", New System.Net.NetworkCredential("pspappca", "Production2007", "PACIFICMUTUAL"))

        'OnTracService.Credentials = cache

        OnTracService.Url = webServiceURL

        Util.ScreenUpdatesOff()

        Util.AddColumnToSheet(ws, 1, 10, True, 5, "ID")
        Util.AddColumnToSheet(ws, 2, 45, True, 5, "Sid")
        Util.AddColumnToSheet(ws, 3, 40, True, 5, "Name")
        Util.AddColumnToSheet(ws, 4, 40, True, 5, "Email")
        Util.AddColumnToSheet(ws, 5, 15, True, 5, "Notes")
        Util.AddColumnToSheet(ws, 5, 15, True, 5, "IsSiteAdmin")
        Util.AddColumnToSheet(ws, 5, 15, True, 5, "IsDomainGroup")

        Try
            Dim node As System.Xml.XmlNode = OnTracService.GetUserCollectionFromWeb()
            Dim innerNode As System.Xml.XmlNode = node.FirstChild()

            Dim i As Integer = 6

            For Each userNode As System.Xml.XmlNode In innerNode
                ws.Cells(i, 1).Value = userNode.Attributes("ID").Value
                ws.Cells(i, 2).Value = userNode.Attributes("Sid").Value
                ws.Cells(i, 3).Value = userNode.Attributes("Name").Value
                ws.Cells(i, 4).Value = userNode.Attributes("Email").Value
                ws.Cells(i, 5).Value = userNode.Attributes("Notes").Value
                ws.Cells(i, 5).Value = userNode.Attributes("IsSiteAdmin").Value
                ws.Cells(i, 5).Value = userNode.Attributes("IsDomainGroup").Value

                i += 1
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        Util.ScreenUpdatesOn()
    End Sub
End Class
