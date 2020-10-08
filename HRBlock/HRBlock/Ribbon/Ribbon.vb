Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.Text
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Office = Microsoft.Office.Core
Imports PacificLife.Life

Partial Public Class ThisAddIn

    Private ribbon As Ribbon

    Protected Overrides Function RequestService(ByVal serviceGuid As Guid) As Object
        If serviceGuid = GetType(Office.IRibbonExtensibility).GUID Then
            If ribbon Is Nothing Then
                ribbon = New Ribbon()
            End If
            Return ribbon
        End If

        Return MyBase.RequestService(serviceGuid)
    End Function

End Class

<ComVisible(True)> _
    Public Class Ribbon
    Implements Office.IRibbonExtensibility

    Private ribbon As Office.IRibbonUI

    Public Sub New()
    End Sub

    Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI
        Return GetResourceText("HRBlock.Ribbon.xml")
    End Function

#Region "Ribbon Callbacks"

    Public Sub OnLoad(ByVal ribbonUI As Office.IRibbonUI)
        Me.ribbon = ribbonUI
    End Sub

#Region "Task Panes"

    ' Routines to add and remove custom task panes and manage their visibility

    Private Function AddTaskPane(ByRef taskPane As System.Windows.Forms.UserControl, ByVal name As String) As Microsoft.Office.Tools.CustomTaskPane
        PLLog.Trace3("Enter", Globals.PROJECT_NAME)

        Dim ctp As Microsoft.Office.Tools.CustomTaskPane
        ctp = Globals.ThisAddIn.CustomTaskPanes.Add(taskPane, name)
        ctp.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
        ctp.Visible = True

        PLLog.Trace3("Exit", Globals.PROJECT_NAME)
        Return ctp

    End Function

    Private Sub RemoveTaskPane(ByRef taskPane As Microsoft.Office.Tools.CustomTaskPane)
        PLLog.Trace3("Enter", Globals.PROJECT_NAME)

        Globals.ThisAddIn.CustomTaskPanes.Remove(taskPane)

        PLLog.Trace3("Exit", Globals.PROJECT_NAME)
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

#Region "TaskPane CreateSheets"

    Public Sub TaskPane_CreateSheets_Click(ByVal control As Office.IRibbonControl)
        If Globals.ThisAddIn.TaskPaneCreateSheets Is Nothing Then
            Globals.ThisAddIn.TaskPaneCreateSheets = AddTaskPane(New TaskPane_CreateSheets, "Create Sheets")
        Else
            Globals.ThisAddIn.TaskPaneCreateSheets.Visible = Not Globals.ThisAddIn.TaskPaneCreateSheets.Visible
        End If
    End Sub

    ' TODO: May want to remove based on Control key or something
    Private Sub RemoveTaskPane_CreateSheets()
        RemoveTaskPane(Globals.ThisAddIn.TaskPaneHelp)
    End Sub

#End Region

#Region "TaskPane HRB"

    Public Sub TaskPane_HRB_Click(ByVal control As Office.IRibbonControl)
        If Globals.ThisAddIn.TaskPaneHRB Is Nothing Then
            Globals.ThisAddIn.TaskPaneHRB = AddTaskPane(New TaskPane_HRB, "HRB Tasks")
        Else
            Globals.ThisAddIn.TaskPaneHRB.Visible = Not Globals.ThisAddIn.TaskPaneHRB.Visible
        End If
    End Sub

    ' TODO: May want to remove based on Control key or something
    Private Sub RemoveTaskPane_HRB()
        RemoveTaskPane(Globals.ThisAddIn.TaskPaneHelp)
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
