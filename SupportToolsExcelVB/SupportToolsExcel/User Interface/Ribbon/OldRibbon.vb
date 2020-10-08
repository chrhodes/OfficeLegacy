Imports System.IO
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports Office = Microsoft.Office.Core
Imports PacificLife.Life

Partial Public Class ThisAddIn

    Private OldRibbon As OldRibbon

    Protected Overrides Function RequestService(ByVal serviceGuid As Guid) As Object
        If serviceGuid = GetType(Office.IRibbonExtensibility).GUID Then
            If OldRibbon Is Nothing Then
                OldRibbon = New OldRibbon()
            End If
            Return OldRibbon
        End If

        Return MyBase.RequestService(serviceGuid)
    End Function

End Class

<ComVisible(True)> _
Public Class OldRibbon
    Implements Office.IRibbonExtensibility

    Private ribbon As Office.IRibbonUI

    Public Sub New()
    End Sub

    Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI
        Return GetResourceText("SupportToolsExcel.Ribbon.xml")
    End Function

#Region "Ribbon Callbacks"

    Public Sub OnLoad(ByVal ribbonUI As Office.IRibbonUI)
        Me.ribbon = ribbonUI
    End Sub


#Region "Task Panes"
    ' Routines to add and remove custom task panes and manage their visibility

    Private Function AddTaskPane(ByRef taskPane As System.Windows.Forms.UserControl, ByVal name As String) As Microsoft.Office.Tools.CustomTaskPane
        PLLog.Trace3("Enter", Common.PROJECT_NAME)

        Dim ctp As Microsoft.Office.Tools.CustomTaskPane
        ctp = Globals.ThisAddIn.CustomTaskPanes.Add(taskPane, name)
        ctp.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
        ctp.Visible = True

        PLLog.Trace3("Exit", Common.PROJECT_NAME)
        Return ctp

    End Function

    Private Sub RemoveTaskPane(ByRef taskPane As Microsoft.Office.Tools.CustomTaskPane)
        PLLog.Trace3("Enter", Common.PROJECT_NAME)

        Globals.ThisAddIn.CustomTaskPanes.Remove(taskPane)

        PLLog.Trace3("Exit", Common.PROJECT_NAME)
    End Sub

#Region "TaskPane Config"

    Public Sub TaskPane_Config_Click(ByVal control As Office.IRibbonControl)
        If Globals.ThisAddIn.TaskPaneConfig Is Nothing Then
            Globals.ThisAddIn.TaskPaneConfig = AddTaskPane(New TaskPane_Config, "Config Tasks")
        Else
            Globals.ThisAddIn.TaskPaneConfig.Visible = Not Globals.ThisAddIn.TaskPaneConfig.Visible
        End If
    End Sub

    Public Sub RemoveTaskPane_Config()
        RemoveTaskPane(Globals.ThisAddIn.TaskPaneConfig)
    End Sub

#End Region

#Region "TaskPane ExcelUtil"

    Public Sub TaskPane_ExcelUtil_Click(ByVal control As Office.IRibbonControl)
        If Globals.ThisAddIn.TaskPaneExcelUtil Is Nothing Then
            Globals.ThisAddIn.TaskPaneExcelUtil = AddTaskPane(New TaskPane_ExcelUtil, "ExcelUtil")
        Else
            Globals.ThisAddIn.TaskPaneExcelUtil.Visible = Not Globals.ThisAddIn.TaskPaneITRs.Visible
        End If
    End Sub

    Public Sub RemoveTaskPane_ExcelUtil()
        RemoveTaskPane(Globals.ThisAddIn.TaskPaneExcelUtil)
    End Sub
#End Region

#Region "TaskPane Help"

    Public Sub TaskPane_Help_Click(ByVal control As Office.IRibbonControl)
        If Globals.ThisAddIn.TaskPaneHelp Is Nothing Then
            Globals.ThisAddIn.TaskPaneHelp = AddTaskPane(New TaskPane_Help, "Help Tasks")
        Else
            Globals.ThisAddIn.TaskPaneHelp.Visible = Not Globals.ThisAddIn.TaskPaneHelp.Visible
        End If
    End Sub

    ' TODO: May want to remove based on Control key or something
    Private Sub RemoveTaskPane_Help()
        RemoveTaskPane(Globals.ThisAddIn.TaskPaneHelp)
    End Sub

#End Region

#Region "TaskPane ITRs"

    Public Sub TaskPane_ITRs_Click(ByVal control As Office.IRibbonControl)
        If Globals.ThisAddIn.TaskPaneITRs Is Nothing Then
            Globals.ThisAddIn.TaskPaneITRs = AddTaskPane(New TaskPane_ITRs, "ITR Tasks")
        Else
            Globals.ThisAddIn.TaskPaneITRs.Visible = Not Globals.ThisAddIn.TaskPaneITRs.Visible
        End If
    End Sub

    Public Sub RemoveTaskPane_ITRs()
        RemoveTaskPane(Globals.ThisAddIn.TaskPaneITRs)
    End Sub
#End Region

#Region "TaskPane NetworkTrace"

    Public Sub TaskPane_NetworkTrace_Click(ByVal control As Office.IRibbonControl)
        If Globals.ThisAddIn.TaskPaneNetworkTrace Is Nothing Then
            Globals.ThisAddIn.TaskPaneNetworkTrace = AddTaskPane(New TaskPane_NetworkTrace, "Network Trace")
        Else
            Globals.ThisAddIn.TaskPaneNetworkTrace.Visible = Not Globals.ThisAddIn.TaskPaneNetworkTrace.Visible
        End If
    End Sub

    Public Sub RemoveTaskPane_NetworkTrace()
        RemoveTaskPane(Globals.ThisAddIn.TaskPaneNetworkTrace)
    End Sub
#End Region

#End Region

#End Region

#Region "Helpers"

    Private Shared Function GetResourceText(ByVal resourceName As String) As String
        Dim asm As Assembly = Assembly.GetExecutingAssembly()
        Dim resourceNames() As String = asm.GetManifestResourceNames()
        For i As Integer = 0 To resourceNames.Length - 1
            If String.Compare(resourceName, resourceNames(i), StringComparison.OrdinalIgnoreCase) = 0 Then
                Using resourceReader As StreamReader = New StreamReader(asm.GetManifestResourceStream(resourceNames(i)))
                    If resourceReader IsNot Nothing Then
                        Return resourceReader.ReadToEnd()
                    End If
                End Using
            End If
        Next
        Return Nothing
    End Function

#End Region

End Class
