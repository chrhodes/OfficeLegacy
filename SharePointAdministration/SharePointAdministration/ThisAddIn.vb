Imports Microsoft.Practices.EnterpriseLibrary.Logging.ExtraInformation
Imports Microsoft.Practices.EnterpriseLibrary.Logging.Filters
Imports PacificLife.Life

public class ThisAddIn
    Public m_vntPriorCalculationState As Object
    Public priorScreenUpdatingState As Boolean = True

    Private Sub ThisAddIn_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        PLLog.Trace("Enter", Globals.cPLLOG_NAME)

        ' Excel 2007 does not like to have these called.

        'RemoveConfigTaskPane()
        'RemoveSurveysTaskPane()
        'RemoveWorksheetsTaskPane()
        'RemoveHelpTaskPane()
        'RemoveOnTimeDeliveryTaskPane()
        'RemoveResultsTaskPane()

        PLLog.Trace("Exit", Globals.cPLLOG_NAME)
    End Sub

    Private Sub ThisAddIn_Startup(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Startup
        PLLog.Trace("Enter", Globals.cPLLOG_NAME)


        PLLog.Trace("Exit", Globals.cPLLOG_NAME)
    End Sub

    Private Sub ThisAddIn_Shutdown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shutdown
        PLLog.Trace("Enter", Globals.cPLLOG_NAME)

        ' Excel 2007 does not like to have these called.

        'RemoveConfigTaskPane()
        'RemoveSurveysTaskPane()
        'RemoveWorksheetsTaskPane()
        'RemoveHelpTaskPane()
        'RemoveOnTimeDeliveryTaskPane()
        'RemoveResultsTaskPane()

        PLLog.Trace("Exit", Globals.cPLLOG_NAME)
    End Sub

#Region "Task Panes"
    ' Routines to add and remove custom task panes and manage their visibility

#Region "Config"

    Private ctpConfig As Microsoft.Office.Tools.CustomTaskPane

    Public Sub AddTaskPane_Config()
        PLLog.Trace3("Enter", Globals.cPLLOG_NAME)
        ctpConfig = Me.CustomTaskPanes.Add(New TaskPane_Config(), "Config Tasks")
        ctpConfig.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
        ctpConfig.Visible = True
        PLLog.Trace3("Exit", Globals.cPLLOG_NAME)
    End Sub

    Public Sub RemoveTaskPane_Config()
        PLLog.Trace3("Enter", Globals.cPLLOG_NAME)
        Me.CustomTaskPanes.Remove(ctpConfig)
        PLLog.Trace3("Exit", Globals.cPLLOG_NAME)
    End Sub

#End Region

#Region "Help"

    Private ctpHelp As Microsoft.Office.Tools.CustomTaskPane

    Public Sub AddTaskPane_Help()
        PLLog.Trace3("Enter", Globals.cPLLOG_NAME)
        ctpHelp = Me.CustomTaskPanes.Add(New TaskPane_Help(), "Help Tasks")
        ctpHelp.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
        ctpHelp.Visible = True
        PLLog.Trace3("Exit", Globals.cPLLOG_NAME)
    End Sub

    Public Sub RemoveTaskPane_Help()
        PLLog.Trace3("Enter", Globals.cPLLOG_NAME)
        Me.CustomTaskPanes.Remove(ctpHelp)
        PLLog.Trace3("Exit", Globals.cPLLOG_NAME)
    End Sub

#End Region

#Region "TaskPane One"

    Private ctpOne As Microsoft.Office.Tools.CustomTaskPane

    Public Sub AddTaskPane_One()
        PLLog.Trace3("Enter", Globals.cPLLOG_NAME)
        ctpOne = Me.CustomTaskPanes.Add(New TaskPane_One(), "TaskPane One")
        ctpOne.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
        ctpOne.Visible = True
        PLLog.Trace3("Exit", Globals.cPLLOG_NAME)
    End Sub

    Public Sub RemoveTaskPane_One()
        PLLog.Trace3("Enter", Globals.cPLLOG_NAME)
        Me.CustomTaskPanes.Remove(ctpOne)
        PLLog.Trace3("Exit", Globals.cPLLOG_NAME)
    End Sub

#End Region

#Region "TaskPane Two"

    Private ctpTwo As Microsoft.Office.Tools.CustomTaskPane

    Public Sub AddTaskPane_Two()
        PLLog.Trace3("Enter", Globals.cPLLOG_NAME)
        ctpTwo = Me.CustomTaskPanes.Add(New TaskPane_Two(), "TaskPane Two")
        ctpTwo.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
        ctpTwo.Visible = True
        PLLog.Trace3("Exit", Globals.cPLLOG_NAME)
    End Sub

    Public Sub RemoveTaskPane_Two()
        PLLog.Trace3("Enter", Globals.cPLLOG_NAME)
        Me.CustomTaskPanes.Remove(ctpTwo)
        PLLog.Trace3("Exit", Globals.cPLLOG_NAME)
    End Sub

#End Region

#Region "Worksheets"

    Private ctpCreateSheets As Microsoft.Office.Tools.CustomTaskPane

    Public Sub AddTaskPane_CreateSheets()
        PLLog.Trace3("Enter", Globals.cPLLOG_NAME)
        ctpCreateSheets = Me.CustomTaskPanes.Add(New TaskPane_CreateSheets(), "TaskPane CreateSheets")
        ctpCreateSheets.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
        ctpCreateSheets.Visible = True
        PLLog.Trace3("Exit", Globals.cPLLOG_NAME)
    End Sub

    Public Sub RemoveTaskPane_CreateSheets()
        PLLog.Trace3("Enter", Globals.cPLLOG_NAME)
        Me.CustomTaskPanes.Remove(ctpCreateSheets)
        PLLog.Trace3("Exit", Globals.cPLLOG_NAME)
    End Sub

#End Region

#End Region

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
