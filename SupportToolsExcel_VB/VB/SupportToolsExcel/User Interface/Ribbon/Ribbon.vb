Imports Microsoft.Office.Tools.Ribbon
Imports PacificLife.Life

Public Class Ribbon

#Region "Initialization"

    Private Sub Ribbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

#End Region

#Region "Event Handlers"

    Private Sub btnAddFooter_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnAddFooter.Click
        Excel_AddFooter.AddFooter()
    End Sub

    Private Sub btnAddInInfo_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnAddInInfo.Click
        DisplayAddInInfo()
    End Sub

    Private Sub btnDebugWindow_Click( sender As System.Object,  e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnDebugWindow.Click
        DisplayDebugWindow()
    End Sub

    Private Sub btnDeveloperMode_Click( sender As System.Object,  e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnDeveloperMode.Click
        ToggleDeveloperMode()
    End Sub

    Private Sub btnExcelUtilities_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnExcelUtilities.Click
        If Common.TaskPaneExcelUtil Is Nothing Then
            Common.TaskPaneExcelUtil = AddinHelper.TaskPaneUtil.AddTaskPane(New TaskPane_ExcelUtil, "ExcelUtil", Globals.ThisAddIn.CustomTaskPanes)
        Else
            Common.TaskPaneExcelUtil.Visible = Not Common.TaskPaneExcelUtil.Visible
        End If
    End Sub

    Private Sub btnFolderMap_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) 
        Excel_FolderMaps.CreateFolderMap()
    End Sub

    Private Sub btnGroupDown_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) 
        Excel_GroupDown.GroupColumnRangeDown()
    End Sub

    Private Sub btnITRs_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnITRs.Click
        If Common.TaskPaneITRs Is Nothing Then
            Common.TaskPaneITRs = AddinHelper.TaskPaneUtil.AddTaskPane(New TaskPane_ITRs, "ITR Tasks", Globals.ThisAddIn.CustomTaskPanes)
        Else
            Common.TaskPaneITRs.Visible = Not Common.TaskPaneITRs.Visible
        End If
    End Sub

    Private Sub btnNetworkTraces_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnNetworkTraces.Click
        If Common.TaskPaneNetworkTrace Is Nothing Then
            Common.TaskPaneNetworkTrace = AddinHelper.TaskPaneUtil.AddTaskPane(New TaskPane_NetworkTrace, "Network Trace", Globals.ThisAddIn.CustomTaskPanes)
        Else
            Common.TaskPaneNetworkTrace.Visible = Not Common.TaskPaneNetworkTrace.Visible
        End If
    End Sub

    Private Sub btnProtectAllWorksheets_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnProtectAllWorksheets.Click
        Excel_ProtectSheets.ProtectSheets()
    End Sub

    Private Sub btnSearchDown_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) 
        Excel_SearchDown.FindEndOfRangeDown()
    End Sub

    Private Sub btnSearchUp_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) 

    End Sub

    Private Sub btnTableOfContents_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnTableOfContents.Click
        Excel_TableOfContents.CreateTableOfContents()
    End Sub

    Private Sub btnUnGroupSelection_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) 
        Excel_UngroupSelection.UnGroupSelection()
    End Sub

    Private Sub btnUnProtectAllWorksheets_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnUnProtectAllWorksheets.Click
        Excel_UnProtectSheets.UnProtectSheets()
    End Sub

    Private Sub btnWatchWindow_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnWatchWindow.Click
        DisplayWatchWindow()
    End Sub

    Private Sub chkDisplayEvents_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles chkDisplayEvents.Click
        Common.DisplayEvents = Not Common.DisplayEvents
        chkDisplayEvents.Checked = Common.DisplayEvents
    End Sub

    Private Sub chkEnableAppEvents_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles chkEnableAppEvents.Click, cbEnableAppEvents.Click
        Common.HAS_APP_EVENTS = Not Common.HAS_APP_EVENTS
        chkEnableAppEvents.Checked = Common.HAS_APP_EVENTS

        If Common.HAS_APP_EVENTS Then
            If Common.AppEvents Is Nothing Then
                Common.AppEvents = New AppEvents
                Common.AppEvents.Initialize()
            End If
        Else
            Common.AppEvents = Nothing
        End If
    End Sub

#End Region

#Region "Main Function Routines"

    Private Sub DisplayAddInInfo()
        AddinHelper.AddInInfo.DisplayInfo()
    End Sub

    Private Sub DisplayDebugWindow()
        Common.DebugWindow.Show()
    End Sub

    Private Sub DisplayWatchWindow()
        AddinHelper.WatchWindow.DisplayWatchWindow()
    End Sub

    Private Sub ToggleDeveloperMode()
        Common.DeveloperMode = Not Common.DeveloperMode
        Globals.Ribbons.Ribbon.grpDebug.Visible = Common.DeveloperMode
    End Sub

#End Region

End Class
