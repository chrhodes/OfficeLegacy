Imports Microsoft.Office.Tools.Ribbon
Imports PacificLife.Life

Public Class Ribbon

#Region "Initialization"

    Private Sub Ribbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

#End Region

#Region "Event Hanlders"

    Private Sub btnAddFooter_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnAddFooter.Click
        Word_AddFooter.AddFooter()
    End Sub

    Private Sub btnAddInInfo_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnAddInInfo.Click
        DisplayAddInInfo()
    End Sub

    Private Sub btnComplianceUtilities_Click( sender As System.Object,  e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnComplianceUtilities.Click
        If Common.TaskPaneComplianceUtil Is Nothing Then
            Common.TaskPaneComplianceUtil = AddinHelper.TaskPaneUtil.AddTaskPane(New TaskPane_ComplianceUtil, "Compliance Utilities", Globals.ThisAddIn.CustomTaskPanes)
        Else
            Common.TaskPaneComplianceUtil.Visible = Not Common.TaskPaneComplianceUtil.Visible
        End If
    End Sub

    Private Sub btnDebugWindow_Click( sender As System.Object,  e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnDebugWindow.Click
        DisplayDebugWindow()
    End Sub

    Private Sub btnDeveloperMode_Click( sender As System.Object,  e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnDeveloperMode.Click
        ToggleDeveloperMode()
    End Sub

    Private Sub btnWordUtilities_Click( sender As System.Object,  e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnWordUtilities.Click
        If Common.TaskPaneWordUtil Is Nothing Then
            Common.TaskPaneWordUtil = AddinHelper.TaskPaneUtil.AddTaskPane(New TaskPane_WordUtil, "Word Utilities", Globals.ThisAddIn.CustomTaskPanes)
        Else
            Common.TaskPaneWordUtil.Visible = Not Common.TaskPaneWordUtil.Visible
        End If
    End Sub

    Private Sub btnWatchWindow_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnWatchWindow.Click
        DisplayWatchWindow()
    End Sub

    Private Sub cbDisplayEvents_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles cbDisplayEvents.Click
        Common.DisplayEvents = Not Common.DisplayEvents
        cbDisplayEvents.Checked = Common.DisplayEvents
    End Sub

    Private Sub cbEnableAppEvents_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles cbEnableAppEvents.Click
        Common.HAS_APP_EVENTS = Not Common.HAS_APP_EVENTS
        cbEnableAppEvents.Checked = Common.HAS_APP_EVENTS

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
