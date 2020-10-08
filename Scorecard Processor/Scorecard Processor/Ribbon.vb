Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.Text
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Office = Microsoft.Office.Core

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
        Return GetResourceText("Scorecard_Processor.Ribbon.xml")
    End Function

#Region "Ribbon Callbacks"

    Public Sub OnLoad(ByVal ribbonUI As Office.IRibbonUI)
        Me.ribbon = ribbonUI
    End Sub

    Private ConfigTaskPaneExists As Boolean = False

    Public Sub ConfigTaskPaneClick(ByVal control As Office.IRibbonControl)
        If Not ConfigTaskPaneExists Then
            Globals.ThisAddIn.AddConfigTaskPane()
        Else
            Globals.ThisAddIn.RemoveConfigTaskPane()
        End If

        ConfigTaskPaneExists = Not ConfigTaskPaneExists
    End Sub

    Private HelpTaskPaneExists As Boolean = False

    Public Sub HelpTaskPaneClick(ByVal control As Office.IRibbonControl)
        If Not HelpTaskPaneExists Then
            Globals.ThisAddIn.AddHelpTaskPane()
        Else
            Globals.ThisAddIn.RemoveHelpTaskPane()
        End If

        HelpTaskPaneExists = Not HelpTaskPaneExists
    End Sub

    Private ResultsTaskPaneExists As Boolean = False

    Public Sub ResultsTaskPaneClick(ByVal control As Office.IRibbonControl)
        If Not ResultsTaskPaneExists Then
            Globals.ThisAddIn.AddResultsTaskPane()
        Else
            Globals.ThisAddIn.RemoveResultsTaskPane()
        End If

        ResultsTaskPaneExists = Not ResultsTaskPaneExists
    End Sub

    Private SurveysTaskPaneExists As Boolean = False

    Public Sub SurveysTaskPaneClick(ByVal control As Office.IRibbonControl)
        If Not SurveysTaskPaneExists Then
            Globals.ThisAddIn.AddSurveysTaskPane()
        Else
            Globals.ThisAddIn.RemoveSurveysTaskPane()
        End If

        SurveysTaskPaneExists = Not SurveysTaskPaneExists
    End Sub

    Private WorksheetsTaskPaneExists As Boolean = False

    Public Sub WorksheetsTaskPaneClick(ByVal control As Office.IRibbonControl)
        If Not WorksheetsTaskPaneExists Then
            Globals.ThisAddIn.AddWorksheetsTaskPane()
        Else
            Globals.ThisAddIn.RemoveWorksheetsTaskPane()
        End If

        WorksheetsTaskPaneExists = Not WorksheetsTaskPaneExists
    End Sub

    Private OnTimeDeliveryTaskPaneExists As Boolean = False

    Public Sub OnTimeDeliveryTaskPaneClick(ByVal control As Office.IRibbonControl)
        If Not OnTimeDeliveryTaskPaneExists Then
            Globals.ThisAddIn.AddOnTimeDeliveryTaskPane()
        Else
            Globals.ThisAddIn.RemoveOnTimeDeliveryTaskPane()
        End If

        OnTimeDeliveryTaskPaneExists = Not OnTimeDeliveryTaskPaneExists
    End Sub
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
