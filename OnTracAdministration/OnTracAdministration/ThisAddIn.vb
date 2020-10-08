Imports Microsoft.Practices.EnterpriseLibrary.Logging.ExtraInformation
Imports Microsoft.Practices.EnterpriseLibrary.Logging.Filters
Imports PacificLife.Life

public class ThisAddIn
    Public m_vntPriorCalculationState As Object
    Public priorScreenUpdatingState As Boolean = True

    Private Sub ThisAddIn_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        PLLog.Trace("Enter", "NEWAPPNAME")

        ' Excel 2007 does not like to have these called.

        'RemoveConfigTaskPane()
        'RemoveSurveysTaskPane()
        'RemoveWorksheetsTaskPane()
        'RemoveHelpTaskPane()
        'RemoveOnTimeDeliveryTaskPane()
        'RemoveResultsTaskPane()

        PLLog.Trace("Exit", "NEWAPPNAME")
    End Sub

    Private Sub ThisAddIn_Startup(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Startup
        PLLog.Trace("Enter", "NEWAPPNAME")


        PLLog.Trace("Exit", "NEWAPPNAME")
    End Sub

    Private Sub ThisAddIn_Shutdown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shutdown
        PLLog.Trace("Enter", "NEWAPPNAME")

        ' Excel 2007 does not like to have these called.

        'RemoveConfigTaskPane()
        'RemoveSurveysTaskPane()
        'RemoveWorksheetsTaskPane()
        'RemoveHelpTaskPane()
        'RemoveOnTimeDeliveryTaskPane()
        'RemoveResultsTaskPane()

        PLLog.Trace("Exit", "NEWAPPNAME")
    End Sub

#Region "Task Panes"
    ' Routines to add and remove custom task panes and manage their visibility

#Region "Config"

    Private ctpConfig As Microsoft.Office.Tools.CustomTaskPane

    Public Sub AddTaskPane_Config()
        PLLog.Trace3("Enter", "NEWAPPNAME")
        ctpConfig = Me.CustomTaskPanes.Add(New TaskPane_Config(), "Config Tasks")
        ctpConfig.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
        ctpConfig.Visible = True
        PLLog.Trace3("Exit", "NEWAPPNAME")
    End Sub

    Public Sub RemoveTaskPane_Config()
        PLLog.Trace3("Enter", "NEWAPPNAME")
        Me.CustomTaskPanes.Remove(ctpConfig)
        PLLog.Trace3("Exit", "NEWAPPNAME")
    End Sub

#End Region

#Region "Help"

    Private ctpHelp As Microsoft.Office.Tools.CustomTaskPane

    Public Sub AddTaskPane_Help()
        PLLog.Trace3("Enter", "NEWAPPNAME")
        ctpHelp = Me.CustomTaskPanes.Add(New TaskPane_Help(), "Help Tasks")
        ctpHelp.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
        ctpHelp.Visible = True
        PLLog.Trace3("Exit", "NEWAPPNAME")
    End Sub

    Public Sub RemoveTaskPane_Help()
        PLLog.Trace3("Enter", "NEWAPPNAME")
        Me.CustomTaskPanes.Remove(ctpHelp)
        PLLog.Trace3("Exit", "NEWAPPNAME")
    End Sub

#End Region

#Region "TaskPane UserAdmin"

    Private ctpUserAdmin As Microsoft.Office.Tools.CustomTaskPane

    Public Sub AddTaskPane_UserAdmin()
        PLLog.Trace3("Enter", "NEWAPPNAME")
        Dim taskPane As TaskPane_UserAdmin = New TaskPane_UserAdmin
        ctpUserAdmin = Me.CustomTaskPanes.Add(taskPane, "TaskPane User Administration")
        ctpUserAdmin.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
        ctpUserAdmin.Width = taskPane.PaneWidth
        ctpUserAdmin.Visible = True
        PLLog.Trace3("Exit", "NEWAPPNAME")
    End Sub

    Public Sub RemoveTaskPane_UserAdmin()
        PLLog.Trace3("Enter", "NEWAPPNAME")
        Me.CustomTaskPanes.Remove(ctpUserAdmin)
        PLLog.Trace3("Exit", "NEWAPPNAME")
    End Sub

#End Region

#Region "TaskPane UsersAndGroups"

    Private ctpUsersAndGroups As Microsoft.Office.Tools.CustomTaskPane

    Public Sub AddTaskPane_UsersAndGroups()
        PLLog.Trace3("Enter", "NEWAPPNAME")
        Dim taskPane As TaskPane_UsersAndGroups = New TaskPane_UsersAndGroups
        ctpUsersAndGroups = Me.CustomTaskPanes.Add(taskPane, "TaskPane Users and Groups")
        ctpUsersAndGroups.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
        ctpUsersAndGroups.Width = taskPane.PaneWidth
        ctpUsersAndGroups.Visible = True
        PLLog.Trace3("Exit", "NEWAPPNAME")
    End Sub

    Public Sub RemoveTaskPane_UsersAndGroups()
        PLLog.Trace3("Enter", "NEWAPPNAME")
        Me.CustomTaskPanes.Remove(ctpUsersAndGroups)
        PLLog.Trace3("Exit", "NEWAPPNAME")
    End Sub

#End Region

#Region "TaskPane Webs"

    Private ctpWebs As Microsoft.Office.Tools.CustomTaskPane

    Public Sub AddTaskPane_Webs()
        PLLog.Trace3("Enter", "NEWAPPNAME")
        Dim taskPane As TaskPane_Webs = New TaskPane_Webs

        'ctpWebs = Me.CustomTaskPanes.Add(New TaskPane_Webs(), "TaskPane Webs")
        ctpWebs = Me.CustomTaskPanes.Add(taskPane, "TaskPane Webs")
        ctpWebs.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
        Debug.Print(taskPane.Width)
        Debug.Print(taskPane.PaneWidth)
        ctpWebs.Width = taskPane.PaneWidth
        'ctpWebs.Width = 300
        ctpWebs.Visible = True
        PLLog.Trace3("Exit", "NEWAPPNAME")
    End Sub

    Public Sub RemoveTaskPane_Webs()
        PLLog.Trace3("Enter", "NEWAPPNAME")
        Me.CustomTaskPanes.Remove(ctpWebs)
        PLLog.Trace3("Exit", "NEWAPPNAME")
    End Sub

#End Region

#Region "TaskPane Two"

    Private ctpTwo As Microsoft.Office.Tools.CustomTaskPane

    Public Sub AddTaskPane_Two()
        PLLog.Trace3("Enter", "NEWAPPNAME")
        ctpTwo = Me.CustomTaskPanes.Add(New TaskPane_UserAdmin(), "TaskPane Two")
        ctpTwo.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
        ctpTwo.Visible = True
        PLLog.Trace3("Exit", "NEWAPPNAME")
    End Sub

    Public Sub RemoveTaskPane_Two()
        PLLog.Trace3("Enter", "NEWAPPNAME")
        Me.CustomTaskPanes.Remove(ctpTwo)
        PLLog.Trace3("Exit", "NEWAPPNAME")
    End Sub

#End Region

#Region "Worksheets"

    Private ctpCreateSheets As Microsoft.Office.Tools.CustomTaskPane

    Public Sub AddTaskPane_CreateSheets()
        PLLog.Trace3("Enter", "NEWAPPNAME")
        ctpCreateSheets = Me.CustomTaskPanes.Add(New TaskPane_CreateSheets(), "TaskPane CreateSheets")
        ctpCreateSheets.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
        ctpCreateSheets.Visible = True
        PLLog.Trace3("Exit", "NEWAPPNAME")
    End Sub

    Public Sub RemoveTaskPane_CreateSheets()
        PLLog.Trace3("Enter", "NEWAPPNAME")
        Me.CustomTaskPanes.Remove(ctpCreateSheets)
        PLLog.Trace3("Exit", "NEWAPPNAME")
    End Sub

#End Region

#End Region

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
