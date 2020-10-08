Imports Microsoft.Practices.EnterpriseLibrary.Logging.ExtraInformation
Imports Microsoft.Practices.EnterpriseLibrary.Logging.Filters
Imports PacificLife.Life

public class ThisAddIn
    Public m_vntPriorCalculationState As Object
    Public priorScreenUpdatingState As Boolean = True

    Private Sub ThisAddIn_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        PLLog.Trace("Enter", "Scorecard")

        ' Excel 2007 does not like to have these called.

        'RemoveConfigTaskPane()
        'RemoveSurveysTaskPane()
        'RemoveWorksheetsTaskPane()
        'RemoveHelpTaskPane()
        'RemoveOnTimeDeliveryTaskPane()
        'RemoveResultsTaskPane()

        PLLog.Trace("Exit", "Scorecard")
    End Sub

    Private Sub ThisAddIn_Startup(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Startup
        PLLog.Trace("Enter", "Scorecard")


        PLLog.Trace("Exit", "Scorecard")
    End Sub

    Private Sub ThisAddIn_Shutdown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shutdown
        PLLog.Trace("Enter", "Scorecard")

        ' Excel 2007 does not like to have these called.

        'RemoveConfigTaskPane()
        'RemoveSurveysTaskPane()
        'RemoveWorksheetsTaskPane()
        'RemoveHelpTaskPane()
        'RemoveOnTimeDeliveryTaskPane()
        'RemoveResultsTaskPane()

        PLLog.Trace("Exit", "Scorecard")
    End Sub

#Region "Task Panes"
    ' Routines to add and remove custom task panes and manage their visibility

#Region "Config"

    Private ctpConfig As Microsoft.Office.Tools.CustomTaskPane

    Public Sub AddConfigTaskPane()
        PLLog.Trace3("Enter", "Scorecard")
        ctpConfig = Me.CustomTaskPanes.Add(New TaskPane_Config(), "Config Tasks")
        ctpConfig.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
        ctpConfig.Visible = True
        PLLog.Trace3("Exit", "Scorecard")
    End Sub

    Public Sub RemoveConfigTaskPane()
        PLLog.Trace3("Enter", "Scorecard")
        Me.CustomTaskPanes.Remove(ctpConfig)
        PLLog.Trace3("Exit", "Scorecard")
    End Sub

#End Region

#Region "Help"

    Private ctpHelp As Microsoft.Office.Tools.CustomTaskPane

    Public Sub AddHelpTaskPane()
        PLLog.Trace3("Enter", "Scorecard")
        ctpHelp = Me.CustomTaskPanes.Add(New TaskPane_Help(), "Help Tasks")
        ctpHelp.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
        ctpHelp.Visible = True
        PLLog.Trace3("Exit", "Scorecard")
    End Sub

    Public Sub RemoveHelpTaskPane()
        PLLog.Trace3("Enter", "Scorecard")
        Me.CustomTaskPanes.Remove(ctpHelp)
        PLLog.Trace3("Exit", "Scorecard")
    End Sub

#End Region

#Region "On-Time Delivery"

    Private ctpOnTimeDelivery As Microsoft.Office.Tools.CustomTaskPane

    Public Sub AddOnTimeDeliveryTaskPane()
        PLLog.Trace3("Enter", "Scorecard")
        ctpOnTimeDelivery = Me.CustomTaskPanes.Add(New TaskPane_OnTimeDelivery(), "On-Time Delivery Tasks")
        ctpOnTimeDelivery.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
        ctpOnTimeDelivery.Visible = True
        PLLog.Trace3("Exit", "Scorecard")
    End Sub

    Public Sub RemoveOnTimeDeliveryTaskPane()
        PLLog.Trace3("Enter", "Scorecard")
        Me.CustomTaskPanes.Remove(ctpOnTimeDelivery)
        PLLog.Trace3("Exit", "Scorecard")
    End Sub

#End Region

#Region "Results"

    Private ctpResults As Microsoft.Office.Tools.CustomTaskPane

    Public Sub AddResultsTaskPane()
        PLLog.Trace3("Enter", "Scorecard")
        ctpResults = Me.CustomTaskPanes.Add(New TaskPane_Results(), "Results Tasks")
        ctpResults.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
        ctpResults.Visible = True
        PLLog.Trace3("Exit", "Scorecard")
    End Sub

    Public Sub RemoveResultsTaskPane()
        PLLog.Trace3("Enter", "Scorecard")
        Me.CustomTaskPanes.Remove(ctpResults)
        PLLog.Trace3("Exit", "Scorecard")
    End Sub

#End Region

#Region "Surveys"

    Private ctpSurveys As Microsoft.Office.Tools.CustomTaskPane

    Public Sub AddSurveysTaskPane()
        PLLog.Trace3("Enter", "Scorecard")
        ctpSurveys = Me.CustomTaskPanes.Add(New TaskPane_Surveys(), "Surveys Tasks")
        ctpSurveys.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
        ctpSurveys.Visible = True
        PLLog.Trace3("Exit", "Scorecard")
    End Sub

    Public Sub RemoveSurveysTaskPane()
        PLLog.Trace3("Enter", "Scorecard")
        Me.CustomTaskPanes.Remove(ctpSurveys)
        PLLog.Trace3("Exit", "Scorecard")
    End Sub

#End Region

#Region "Worksheets"

    Private ctpWorksheets As Microsoft.Office.Tools.CustomTaskPane

    Public Sub AddWorksheetsTaskPane()
        PLLog.Trace3("Enter", "Scorecard")
        ctpWorksheets = Me.CustomTaskPanes.Add(New TaskPane_Worksheets(), "Worksheets Tasks")
        ctpWorksheets.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
        ctpWorksheets.Visible = True
        PLLog.Trace3("Exit", "Scorecard")
    End Sub

    Public Sub RemoveWorksheetsTaskPane()
        PLLog.Trace3("Enter", "Scorecard")
        Me.CustomTaskPanes.Remove(ctpWorksheets)
        PLLog.Trace3("Exit", "Scorecard")
    End Sub

#End Region

#End Region

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
