Imports Microsoft.Practices.EnterpriseLibrary.Logging.ExtraInformation
Imports Microsoft.Practices.EnterpriseLibrary.Logging.Filters
Imports PacificLife.Life

public class ThisAddIn
    Public m_vntPriorCalculationState As Object
    Public priorScreenUpdatingState As Boolean = True
    Private Dim _servers As Servers
    Private _webService As SystemManagementWS.WMIInfoWS
    Dim _cache As System.Net.CredentialCache


    Public ReadOnly Property Servers() As Servers
        Get
            Return _servers
        End Get
    End Property

    Public ReadOnly Property WebService() As SystemManagementWS.WMIInfoWS
        Get
            Return _webService
        End Get
    End Property

    Public ReadOnly Property Cache() As System.Net.CredentialCache
        Get
            Return _cache
        End Get
    End Property

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

        _webService = New SystemManagementWS.WMIInfoWS()
        _cache = New System.Net.CredentialCache
        ' Populate WebService Credential Cache with servers
        _servers = New Servers(Cache, WebService)
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

#Region "TaskPane FarmHealth"

    Private ctpFarmHealth As Microsoft.Office.Tools.CustomTaskPane

    Public Sub AddTaskPane_FarmHealth()
        PLLog.Trace3("Enter", Globals.cPLLOG_NAME)
        ctpFarmHealth = Me.CustomTaskPanes.Add(New TaskPane_FarmHealth(), "TaskPane FarmHealth")
        ctpFarmHealth.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
        ctpFarmHealth.Visible = True
        PLLog.Trace3("Exit", Globals.cPLLOG_NAME)
    End Sub

    Public Sub RemoveTaskPane_FarmHealth()
        PLLog.Trace3("Enter", Globals.cPLLOG_NAME)
        Me.CustomTaskPanes.Remove(ctpFarmHealth)
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

#Region "TaskPane LogicalDisk"

    Private ctpLogicalDiskInfo As Microsoft.Office.Tools.CustomTaskPane

    Public Sub AddTaskPane_LogicalDiskInfo()
        PLLog.Trace3("Enter", Globals.cPLLOG_NAME)
        ctpLogicalDiskInfo = Me.CustomTaskPanes.Add(New TaskPane_LogicalDiskInfo(), "TaskPane LogicalDiskInfo")
        ctpLogicalDiskInfo.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
        ctpLogicalDiskInfo.Visible = True
        PLLog.Trace3("Exit", Globals.cPLLOG_NAME)
    End Sub

    Public Sub RemoveTaskPane_LogicalDiskInfo()
        PLLog.Trace3("Enter", Globals.cPLLOG_NAME)
        Me.CustomTaskPanes.Remove(ctpLogicalDiskInfo)
        PLLog.Trace3("Exit", Globals.cPLLOG_NAME)
    End Sub

#End Region

#Region "TaskPane MemoryDevice"

    Private ctpMemoryDeviceInfo As Microsoft.Office.Tools.CustomTaskPane

    Public Sub AddTaskPane_MemoryDeviceInfo()
        PLLog.Trace3("Enter", Globals.cPLLOG_NAME)
        ctpMemoryDeviceInfo = Me.CustomTaskPanes.Add(New TaskPane_MemoryDeviceInfo(), "TaskPane MemoryDeviceInfo")
        ctpMemoryDeviceInfo.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
        ctpMemoryDeviceInfo.Visible = True
        PLLog.Trace3("Exit", Globals.cPLLOG_NAME)
    End Sub

    Public Sub RemoveTaskPane_MemoryDeviceInfo()
        PLLog.Trace3("Enter", Globals.cPLLOG_NAME)
        Me.CustomTaskPanes.Remove(ctpMemoryDeviceInfo)
        PLLog.Trace3("Exit", Globals.cPLLOG_NAME)
    End Sub

#End Region

#Region "TaskPane PhysicalMemory"

    Private ctpPhysicalMemoryInfo As Microsoft.Office.Tools.CustomTaskPane

    Public Sub AddTaskPane_PhysicalMemoryInfo()
        PLLog.Trace3("Enter", Globals.cPLLOG_NAME)
        ctpPhysicalMemoryInfo = Me.CustomTaskPanes.Add(New TaskPane_PhysicalMemoryInfo(), "TaskPane PhysicalMemoryInfo")
        ctpPhysicalMemoryInfo.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
        ctpPhysicalMemoryInfo.Visible = True
        PLLog.Trace3("Exit", Globals.cPLLOG_NAME)
    End Sub

    Public Sub RemoveTaskPane_PhysicalMemoryInfo()
        PLLog.Trace3("Enter", Globals.cPLLOG_NAME)
        Me.CustomTaskPanes.Remove(ctpPhysicalMemoryInfo)
        PLLog.Trace3("Exit", Globals.cPLLOG_NAME)
    End Sub

#End Region

#Region "TaskPane Processor"

    Private ctpProcessorInfo As Microsoft.Office.Tools.CustomTaskPane

    Public Sub AddTaskPane_ProcessorInfo()
        PLLog.Trace3("Enter", Globals.cPLLOG_NAME)
        ctpProcessorInfo = Me.CustomTaskPanes.Add(New TaskPane_ProcessorInfo(), "TaskPane ProcessorInfo")
        ctpProcessorInfo.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
        ctpProcessorInfo.Visible = True
        PLLog.Trace3("Exit", Globals.cPLLOG_NAME)
    End Sub

    Public Sub RemoveTaskPane_ProcessorInfo()
        PLLog.Trace3("Enter", Globals.cPLLOG_NAME)
        Me.CustomTaskPanes.Remove(ctpProcessorInfo)
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
