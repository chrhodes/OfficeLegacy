Public Class TaskPane_Config
    Private Sub TaskPane_Config_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' Ensure that any config data we need is available.  Ok to call multiple times.
        'Config.IntializeApplication()

        Me.txtSurveyChartWidthExcel.Text = Globals.cSD_SurveyChartWidth.ToString
        Me.txtSurveyChartHeightExcel.Text = Globals.cSD_SurveyChartHeight.ToString
        Me.txtSurveyChartStartingOffsetExcel.Text = Globals.cSD_SurveyChartStartingOffset.ToString
        Me.txtSurveyChartSpacingExcel.Text = Globals.cSD_SurveyChartSpacing.ToString

        Me.txtDataLabelFontSize.Text = Globals.cCH_DataLabelFontSize.ToString
        Me.txtTickLabelFontSize.Text = Globals.cCH_TickLabelFontSize.ToString

        Me.txtSurveyChartTopPowerPoint.Text = Globals.cPP_SurveyChartTop.ToString
        Me.txtSurveyChartLeftPowerPoint.Text = Globals.cPP_SurveyChartLeft.ToString
        Me.txtSurveyChartWidthPowerPoint.Text = Globals.cPP_SurveyChartWidth.ToString
        Me.txtSurveyChartHeightPowerPoint.Text = Globals.cPP_SurveyChartHeight.ToString

        Me.txtOnTimeChartTopPowerPoint.Text = Globals.cPP_OnTimeChartTop.ToString
        Me.txtOnTimeChartLeftPowerPoint.Text = Globals.cPP_OnTimeChartLeft.ToString
        Me.txtOnTimeChartWidthPowerPoint.Text = Globals.cPP_OnTimeChartWidth.ToString
        Me.txtOnTimeChartHeightPowerPoint.Text = Globals.cPP_OnTimeChartHeight.ToString

        Me.txtPowerPointTitleFontSize.Text = Globals.cPP_TitleFontSize.ToString
        Me.txtPowerPointSubQuestionFontSize.Text = Globals.cPP_SubQuestionFontSize.ToString
        Me.txtPowerPointResponseFontSize.Text = Globals.cPP_ResponseFontSize.ToString
        Me.txtPowerPointResponseLeftMargin.Text = Globals.cPP_ResponseLeftMargin.ToString
        Me.txtPowerPointResponseLengthPerPage.Text = Globals.cPP_MaxResponseLengthPerPage.ToString

        Me.chkScreenUpdatesOff.Checked = Globals.cScreenUpdatesOff
    End Sub

    Private Sub txtExcelChartHeight_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSurveyChartHeightExcel.TextChanged
        Globals.cSD_SurveyChartHeight = Me.txtSurveyChartHeightExcel.Text.ToString
    End Sub

    Private Sub txtExcelChartWidth_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSurveyChartWidthExcel.TextChanged
        Globals.cSD_SurveyChartWidth = Me.txtSurveyChartWidthExcel.Text.ToString
    End Sub

    Private Sub txtSurveyChartSpacingExcel_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSurveyChartSpacingExcel.TextChanged
        Globals.cSD_SurveyChartSpacing = Me.txtSurveyChartSpacingExcel.Text.ToString
    End Sub

    Private Sub txtSurveyChartStartingOffsetExcel_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSurveyChartStartingOffsetExcel.TextChanged
        Globals.cSD_SurveyChartStartingOffset = Me.txtSurveyChartStartingOffsetExcel.Text.ToString
    End Sub

    Private Sub txtSurveyChartTopPowerPoint_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSurveyChartTopPowerPoint.TextChanged
        Globals.cPP_SurveyChartTop = Me.txtSurveyChartTopPowerPoint.Text.ToString
    End Sub

    Private Sub txtSurveyChartLeftPowerPoint_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSurveyChartLeftPowerPoint.TextChanged
        Globals.cPP_SurveyChartLeft = Me.txtSurveyChartLeftPowerPoint.Text.ToString
    End Sub

    Private Sub txtSurveyChartHeightPowerPoint_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSurveyChartHeightPowerPoint.TextChanged
        Globals.cPP_SurveyChartHeight = Me.txtSurveyChartHeightPowerPoint.Text.ToString
    End Sub

    Private Sub txtSurveyChartWidthPowerPoint_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSurveyChartWidthPowerPoint.TextChanged
        Globals.cPP_SurveyChartWidth = Me.txtSurveyChartWidthPowerPoint.Text.ToString
    End Sub

    Private Sub chkScreenUpdatesOff_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkScreenUpdatesOff.CheckedChanged
        Globals.cScreenUpdatesOff = Me.chkScreenUpdatesOff.Checked
    End Sub

    Private Sub txtPowerPointTitleFontSize_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPowerPointTitleFontSize.TextChanged
        Globals.cPP_TitleFontSize = Me.txtPowerPointTitleFontSize.Text.ToString
    End Sub

    Private Sub txtPowerPointSubQuestionFontSize_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPowerPointSubQuestionFontSize.TextChanged
        Globals.cPP_SubQuestionFontSize = Me.txtPowerPointSubQuestionFontSize.Text.ToString
    End Sub

    Private Sub txtPowerPointResponseFontSize_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPowerPointResponseFontSize.TextChanged
        Globals.cPP_ResponseFontSize = Me.txtPowerPointResponseFontSize.Text.ToString
    End Sub

    Private Sub txtResponseLeftMargin_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPowerPointResponseLeftMargin.TextChanged
        Globals.cPP_ResponseLeftMargin = Me.txtPowerPointResponseLeftMargin.Text.ToString
    End Sub

    Private Sub txtDataLabelFontSize_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDataLabelFontSize.TextChanged
        Globals.cCH_DataLabelFontSize = Me.txtDataLabelFontSize.Text.ToString
    End Sub

    Private Sub txtTickLabelFontSize_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTickLabelFontSize.TextChanged
        Globals.cCH_TickLabelFontSize = Me.txtTickLabelFontSize.Text.ToString
    End Sub

    Private Sub txtPowerPointResponseLengthPerPage_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPowerPointResponseLengthPerPage.TextChanged
        Globals.cPP_MaxResponseLengthPerPage = Me.txtPowerPointResponseLengthPerPage.Text.ToString
    End Sub

    Private Sub btnFindLast_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindLast.Click
        Util.FindLast()
    End Sub

    Private Sub btnReLoadConfigData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReLoadConfigData.Click
        'TODO: Implement a refresh.
        'MessageBox.Show("Not Implemented")
        ''Config.LoadConfigDataFromXMLFile()
        Config.ReIntializeApplication()
    End Sub

    Private Sub txtOnTimeChartTopPowerPoint_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOnTimeChartTopPowerPoint.TextChanged
        Globals.cPP_OnTimeChartTop = Me.txtOnTimeChartTopPowerPoint.Text.ToString
    End Sub

    Private Sub txtOnTimeChartLeftPowerPoint_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOnTimeChartLeftPowerPoint.TextChanged
        Globals.cPP_OnTimeChartLeft = Me.txtOnTimeChartLeftPowerPoint.Text.ToString
    End Sub

    Private Sub txtOnTimeChartHeightPowerPoint_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOnTimeChartHeightPowerPoint.TextChanged
        Globals.cPP_OnTimeChartHeight = Me.txtOnTimeChartHeightPowerPoint.Text.ToString
    End Sub

    Private Sub txtOnTimeChartWidthPowerPoint_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOnTimeChartWidthPowerPoint.TextChanged
        Globals.cPP_OnTimeChartWidth = Me.txtOnTimeChartWidthPowerPoint.Text.ToString
    End Sub

    Private Sub btnLoadLookups_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoadLookups.Click
        Dim ws As Excel.Worksheet
        Dim rng As Excel.Range
        Dim i As Integer
        Dim expression As String
        Dim foundManager() As DataRow
        Dim foundManagersManager() As DataRow
        Dim currentProtectionMode As Boolean
        Dim currentSheet As Excel.Worksheet

        ' Save where we are, then activate the sheet containing team data.
        ' Turn off screen updates so things run a bit faster.

        currentSheet = Globals.ThisAddIn.Application.ActiveSheet
        ws = Globals.ThisAddIn.Application.Sheets(Globals.cSN_Teams)
        ws.Activate()
        currentProtectionMode = Util.ProtectSheet(ws, False)
        Util.ScreenUpdatesOff()
        Util.CalculationsOff()

        ' First load the TeamID to TeamName lookup data

        rng = ws.Range(Globals.cLU_TeamsInfoCell)
        i = 0   ' Team offset

        'Dim managers As Data.DataTable = Config.Teams.Tables("manager")

        For Each dataRow As Data.DataRow In Config.ConfigInfo.Tables("team").Rows
            ' Fill in team information

            rng.Offset(i, 0).Value = dataRow.Item("team_Id").ToString()
            rng.Offset(i, 1).Value = dataRow.Item("name").ToString()

            '' Fill in team manager information.  We presume you only have
            '' one manager.  Hence we only bother with array(0)

            'expression = "team_Id = " & dataRow.Item("team_Id")

            'foundManager = managers.Select(expression)

            'rng.Offset(i, 2).Value = foundManager(0).Item("name").ToString()
            'rng.Offset(i, 3).Value = foundManager(0).Item("ext").ToString()

            '' Now see if manager has a manager listed.  This is a bit painful as
            '' we have to see if someone is someone elses manager.  There is no
            '' direct link from a person to their manager.

            'expression = "manager_Id_0 = " & foundManager(0).Item("manager_Id")

            'foundManagersManager = managers.Select(expression)

            'If foundManagersManager.GetLength(0) Then
            '    rng.Offset(i, 4).Value = foundManagersManager(0).Item("name").ToString()
            '    rng.Offset(i, 5).Value = foundManagersManager(0).Item("ext").ToString()
            'End If

            i += 1
        Next

        ' Next, load the TeamName to TeamManager lookup data

        rng = ws.Range(Globals.cLU_ManagerInfoCell)
        i = 0   ' Team offset

        Dim managers As Data.DataTable = Config.ConfigInfo.Tables("manager")

        For Each dataRow As Data.DataRow In Config.ConfigInfo.Tables("team").Rows
            ' Fill in team information

            rng.Offset(i, 0).Value = dataRow.Item("team_Id").ToString()
            rng.Offset(i, 1).Value = dataRow.Item("name").ToString()

            ' Fill in team manager information.  We presume you only have
            ' one manager.  Hence we only bother with array(0)

            expression = "team_Id = " & dataRow.Item("team_Id")

            foundManager = managers.Select(expression)

            rng.Offset(i, 2).Value = foundManager(0).Item("name").ToString()
            rng.Offset(i, 3).Value = foundManager(0).Item("ext").ToString()

            ' Now see if manager has a manager listed.  This is a bit painful as
            ' we have to see if someone is someone elses manager.  There is no
            ' direct link from a person to their manager.

            expression = "manager_Id_0 = " & foundManager(0).Item("manager_Id")

            foundManagersManager = managers.Select(expression)

            If foundManagersManager.GetLength(0) Then
                rng.Offset(i, 4).Value = foundManagersManager(0).Item("name").ToString()
                rng.Offset(i, 5).Value = foundManagersManager(0).Item("ext").ToString()
            End If

            i += 1
        Next

        ' Now create a Defined names that will be used to surface the list of
        ' Teams in various places it is needed, e.g. Scorecard worksheet

        Dim workbook As Excel.Workbook = Globals.ThisAddIn.Application.ActiveWorkbook
        ' TODO: Get rid of hard coding
        Dim teamNameRange As String = "=Teams!R5C2:R20C2"
        workbook.Names.Item("Team_Names").Delete()
        workbook.Names.Add(Name:="Team_Names", RefersToR1C1:=teamNameRange)

        Util.ProtectSheet(ws, True)
        Util.CalculationsOn()
        Util.ScreenUpdatesOn()
        currentSheet.Activate()

    End Sub

    Private Sub btnCreateNames_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateNames.Click
        Dim workbook As Excel.Workbook = Globals.ThisAddIn.Application.ActiveWorkbook

        For Each dataRow As Data.DataRow In Config.DefinedNames.Rows
            ' Create the defined names we need.

            Util.CreateName(workbook, dataRow.Item("name").ToString, dataRow.Item("targetRange").ToString)
        Next
    End Sub


    Private Sub chkDisplayDebugMessages_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDisplayDebugMessages.CheckedChanged

    End Sub
End Class
