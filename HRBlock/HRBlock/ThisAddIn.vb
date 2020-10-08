Imports PacificLife.Life

public class ThisAddIn
    Private _CmdBars As CmdBars
    Private _AppEvents As AppEvents

    Public AppEventsWatchWindow As AddinHelper.WatchWindow

    Public ExcelUtil As New ExcelHelper.ExcelHelper

    Private Sub ThisAddIn_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        PLLog.Trace("Enter", Globals.PROJECT_NAME)

        PLLog.Trace("Exit", Globals.PROJECT_NAME)
    End Sub

    Private Sub ThisAddIn_Startup(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Startup
        PLLog.Trace("Enter", Globals.PROJECT_NAME)

        Try
            _CmdBars = New CmdBars
            _CmdBars.CreateCommandBars()

            _AppEvents = New AppEvents
            _AppEvents.Initialize()

            ' Set the context for the ExcelUtil code to this application
            ExcelUtil.Application = Globals.ThisAddIn.Application
        Catch ex As Exception
            PLLog.Error(ex, Globals.PROJECT_NAME)
        End Try

        PLLog.Trace("Exit", Globals.PROJECT_NAME)
    End Sub

    Private Sub ThisAddIn_Shutdown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shutdown
        PLLog.Trace("Enter", Globals.PROJECT_NAME)

        Try
            _CmdBars.RemoveCommandBars()
            _CmdBars = Nothing
            _AppEvents = Nothing
        Catch ex As Exception
            PLLog.Error(ex, Globals.PROJECT_NAME)
        End Try

        PLLog.Trace("Exit", Globals.PROJECT_NAME)
    End Sub

#Region "Task Panes"
    ' Routines to add and remove custom task panes and manage their visibility

    ' The Ribbon class does not stay loaded during the lifetime of our addin so keep the task pane
    ' variables here so they stay loaded.

#Region "Config"

    Private _taskPaneConfig As Microsoft.Office.Tools.CustomTaskPane

    'Public Sub AddTaskPane_Config()
    '    PLLog.Trace3("Enter", Globals.PROJECT_NAME)
    '    ctpConfig = Me.CustomTaskPanes.Add(New TaskPane_Config(), "Config Tasks")
    '    ctpConfig.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
    '    ctpConfig.Visible = True
    '    PLLog.Trace3("Exit", Globals.PROJECT_NAME)
    'End Sub

    'Public Sub RemoveTaskPane_Config()
    '    PLLog.Trace3("Enter", Globals.PROJECT_NAME)
    '    Me.CustomTaskPanes.Remove(ctpConfig)
    '    PLLog.Trace3("Exit", Globals.PROJECT_NAME)
    'End Sub

    Public Property TaskPaneConfig() As Microsoft.Office.Tools.CustomTaskPane
        Get
            Return _taskPaneConfig
        End Get
        Set(ByVal value As Microsoft.Office.Tools.CustomTaskPane)
            _taskPaneConfig = value
        End Set
    End Property
#End Region

#Region "Help"

    Private _taskPaneHelp As Microsoft.Office.Tools.CustomTaskPane

    'Public Sub AddTaskPane_Help()
    '    PLLog.Trace3("Enter", Globals.PROJECT_NAME)
    '    ctpHelp = Me.CustomTaskPanes.Add(New TaskPane_Help(), "Help Tasks")
    '    ctpHelp.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
    '    ctpHelp.Visible = True
    '    PLLog.Trace3("Exit", Globals.PROJECT_NAME)
    'End Sub

    'Public Sub RemoveTaskPane_Help()
    '    PLLog.Trace3("Enter", Globals.PROJECT_NAME)
    '    Me.CustomTaskPanes.Remove(ctpHelp)
    '    PLLog.Trace3("Exit", Globals.PROJECT_NAME)
    'End Sub

    Public Property TaskPaneHelp() As Microsoft.Office.Tools.CustomTaskPane
        Get
            Return _taskPaneHelp
        End Get
        Set(ByVal value As Microsoft.Office.Tools.CustomTaskPane)
            _taskPaneHelp = value
        End Set
    End Property
#End Region

#Region "TaskPane HRB"

    Private _taskPaneHRB As Microsoft.Office.Tools.CustomTaskPane

    Public Property TaskPaneHRB() As Microsoft.Office.Tools.CustomTaskPane
        Get
            Return _taskPaneHRB
        End Get
        Set(ByVal value As Microsoft.Office.Tools.CustomTaskPane)
            _taskPaneHRB = value
        End Set
    End Property

#End Region

#Region "TaskPane Two"

    Private _taskPaneTwo As Microsoft.Office.Tools.CustomTaskPane

    Public Property TaskPaneTwo() As Microsoft.Office.Tools.CustomTaskPane
        Get
            Return _taskPaneTwo
        End Get
        Set(ByVal value As Microsoft.Office.Tools.CustomTaskPane)
            _taskPaneTwo = value
        End Set
    End Property
#End Region

#Region "Worksheets"

    Private _taskPaneCreateSheets As Microsoft.Office.Tools.CustomTaskPane

    Public Property TaskPaneCreateSheets() As Microsoft.Office.Tools.CustomTaskPane
        Get
            Return _taskPaneCreateSheets
        End Get
        Set(ByVal value As Microsoft.Office.Tools.CustomTaskPane)
            _taskPaneCreateSheets = value
        End Set
    End Property

#End Region

#End Region

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
