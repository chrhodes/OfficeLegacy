Imports Microsoft.Office.Interop
Imports System.Data

Public Class TaskPane_Config
    Private Sub TaskPane_Config_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' TODO: Ensure that any config data we need is available.  Ok to call multiple times.
        ' This method should populate any controls on this task pane with data from Globals.vb or a loaded config file.

        Me.chkScreenUpdatesOff.Checked = Common.ScreenUpdatesOff
    End Sub

    'Private Sub chkScreenUpdatesOff_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkScreenUpdatesOff.CheckedChanged
    '    'Common.ScreenUpdatesOff = Me.chkScreenUpdatesOff.Checked
    '    'Common.ExcelHelper.EnableScreenUpdatesToggle = Me.chkScreenUpdatesOff.Checked
    'End Sub

    'Private Sub btnReLoadConfigData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReLoadConfigData.Click
    '    Config.ReIntializeApplication()

    '    'For Each dataTable As Data.DataTable In Config.ConfigInfo.Tables
    '    '    Select Case dataTable.TableName
    '    '        Case "team"
    '    '            For Each dataRow As Data.DataRow In dataTable.Rows
    '    '                Me.cbTeamName.Items.Add(dataRow.Item("name")).ToString()
    '    '            Next
    '    '    End Select
    '    'Next

    '    ' TODO: Perhaps could have a property in Config for each table.

    '    For Each dataRow As Data.DataRow In Config.ConfigInfo.Tables("team").Rows
    '        Me.cbTeamName.Items.Add(dataRow.Item("name")).ToString()
    '    Next
    'End Sub

    'Private Sub btnLoadLookups_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoadLookups.Click
    '    Dim ws As Excel.Worksheet
    '    Dim rng As Excel.Range
    '    Dim i As Integer
    '    Dim expression As String
    '    Dim foundManager() As DataRow
    '    Dim foundManagersManager() As DataRow
    '    Dim currentProtectionMode As Boolean
    '    Dim currentSheet As Excel.Worksheet

    '    ' Save where we are, then activate the sheet containing team data.
    '    ' Turn off screen updates so things run a bit faster.

    '    currentSheet = Globals.ThisAddIn.Application.ActiveSheet
    '    ws = Globals.ThisAddIn.Application.Sheets(Common.cSN_Teams)
    '    ws.Activate()
    '    currentProtectionMode = Common.ExcelHelper.ProtectSheet(ws, False)
    '    Common.ExcelHelper.ScreenUpdatesOff()
    '    Common.ExcelHelper.CalculationsOff()

    '    ' First load the TeamID to TeamName lookup data

    '    rng = ws.Range(Common.cLU_TeamsInfoCell)
    '    i = 0   ' Team offset

    '    'Dim managers As Data.DataTable = Config.Teams.Tables("manager")

    '    For Each dataRow As Data.DataRow In Config.ConfigInfo.Tables("team").Rows
    '        ' Fill in team information

    '        rng.Offset(i, 0).Value = dataRow.Item("team_Id").ToString()
    '        rng.Offset(i, 1).Value = dataRow.Item("name").ToString()

    '        i += 1
    '    Next

    '    ' Next, load the TeamName to TeamManager lookup data

    '    rng = ws.Range(Common.cLU_ManagerInfoCell)
    '    i = 0   ' Team offset

    '    Dim managers As Data.DataTable = Config.ConfigInfo.Tables("manager")

    '    For Each dataRow As Data.DataRow In Config.ConfigInfo.Tables("team").Rows
    '        ' Fill in team information

    '        rng.Offset(i, 0).Value = dataRow.Item("team_Id").ToString()
    '        rng.Offset(i, 1).Value = dataRow.Item("name").ToString()

    '        ' Fill in team manager information.  We presume you only have
    '        ' one manager.  Hence we only bother with array(0)

    '        expression = String.Format("team_Id = {0}", dataRow.Item("team_Id"))

    '        foundManager = managers.Select(expression)

    '        rng.Offset(i, 2).Value = foundManager(0).Item("name").ToString()
    '        rng.Offset(i, 3).Value = foundManager(0).Item("ext").ToString()

    '        ' Now see if manager has a manager listed.  This is a bit painful as
    '        ' we have to see if someone is someone elses manager.  There is no
    '        ' direct link from a person to their manager.

    '        expression = String.Format("manager_Id_0 = {0}", foundManager(0).Item("manager_Id"))

    '        foundManagersManager = managers.Select(expression)

    '        If foundManagersManager.GetLength(0) Then
    '            rng.Offset(i, 4).Value = foundManagersManager(0).Item("name").ToString()
    '            rng.Offset(i, 5).Value = foundManagersManager(0).Item("ext").ToString()
    '        End If

    '        i += 1
    '    Next

    '    ' Now create a Defined names that will be used to surface the list of
    '    ' Teams in various places it is needed, e.g. Scorecard worksheet

    '    Dim workbook As Excel.Workbook = Globals.ThisAddIn.Application.ActiveWorkbook
    '    ' TODO: Get rid of hard coding
    '    Dim teamNameRange As String = "=Teams!R5C2:R20C2"
    '    workbook.Names.Item("Team_Names").Delete()
    '    workbook.Names.Add(Name:="Team_Names", RefersToR1C1:=teamNameRange)

    '    Common.ExcelHelper.ProtectSheet(ws, True)
    '    Common.ExcelHelper.CalculationsOn()
    '    Common.ExcelHelper.ScreenUpdatesOn()
    '    currentSheet.Activate()

    'End Sub

    'Private Sub btnCreateNames_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateDefinedNames.Click
    '    Dim workbook As Excel.Workbook = Globals.ThisAddIn.Application.ActiveWorkbook

    '    For Each dataRow As Data.DataRow In Config.DefinedNames.Rows
    '        ' Create the defined names we need.

    '        'Common.ExcelUtil.CreateName(workbook, dataRow.Item("name").ToString, dataRow.Item("targetRange").ToString)
    '    Next
    'End Sub


    'Private Sub chkDisplayDebugMessages_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDisplayDebugMessages.CheckedChanged

    'End Sub
End Class
