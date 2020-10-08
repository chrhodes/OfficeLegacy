Imports Microsoft.Office.Tools.Ribbon
Imports PacificLife.Life

Public Class Ribbon

#Region "Initialization"

    Private Sub Ribbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

#End Region

#Region "Event Handlers"

    Private Sub btnAddInInfo_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnAddInInfo.Click
        DisplayAddInInfo()
    End Sub

    Private Sub btnDebugWindow_Click( sender As System.Object,  e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnDebugWindow.Click
        DisplayDebugWindow()
    End Sub

    Private Sub btnDeveloperMode_Click( sender As System.Object,  e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnDeveloperMode.Click
        ToggleDeveloperMode()
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
