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
        Return GetResourceText("OnTracAdministration.Ribbon.xml")
    End Function

#Region "Ribbon Callbacks"

    Public Sub OnLoad(ByVal ribbonUI As Office.IRibbonUI)
        Me.ribbon = ribbonUI
    End Sub

    Private _TaskPane_Config_Exists As Boolean = False

    Public Sub TaskPane_Config_Click(ByVal control As Office.IRibbonControl)
        If Not _TaskPane_Config_Exists Then
            Globals.ThisAddIn.AddTaskPane_Config()
        Else
            Globals.ThisAddIn.RemoveTaskPane_Config()
        End If

        _TaskPane_Config_Exists = Not _TaskPane_Config_Exists
    End Sub

    Private _TaskPane_Help_Exists As Boolean = False

    Public Sub TaskPane_Help_Click(ByVal control As Office.IRibbonControl)
        If Not _TaskPane_Help_Exists Then
            Globals.ThisAddIn.AddTaskPane_Help()
        Else
            Globals.ThisAddIn.RemoveTaskPane_Help()
        End If

        _TaskPane_Help_Exists = Not _TaskPane_Help_Exists
    End Sub

    Private _TaskPane_Two_Exists As Boolean = False

    Public Sub TaskPane_Two_Click(ByVal control As Office.IRibbonControl)
        If Not _TaskPane_Two_Exists Then
            Globals.ThisAddIn.AddTaskPane_Two()
        Else
            Globals.ThisAddIn.RemoveTaskPane_Two()
        End If

        _TaskPane_Two_Exists = Not _TaskPane_Two_Exists
    End Sub

    Private _TaskPane_CreateSheets_Exists As Boolean = False

    Public Sub TaskPane_CreateSheets_Click(ByVal control As Office.IRibbonControl)
        If Not _TaskPane_CreateSheets_Exists Then
            Globals.ThisAddIn.AddTaskPane_CreateSheets()
        Else
            Globals.ThisAddIn.RemoveTaskPane_CreateSheets()
        End If

        _TaskPane_CreateSheets_Exists = Not _TaskPane_CreateSheets_Exists
    End Sub

    Private _TaskPane_UserAdmin_Exists As Boolean = False

    Public Sub TaskPane_USerAdmin_Click(ByVal control As Office.IRibbonControl)
        If Not _TaskPane_UserAdmin_Exists Then
            Globals.ThisAddIn.AddTaskPane_UserAdmin()
        Else
            Globals.ThisAddIn.RemoveTaskPane_UserAdmin()
        End If

        _TaskPane_UserAdmin_Exists = Not _TaskPane_UserAdmin_Exists
    End Sub

    Private _TaskPane_UsersAndGroups_Exists As Boolean = False

    Public Sub TaskPane_UsersAndGroups_Click(ByVal control As Office.IRibbonControl)
        If Not _TaskPane_UsersAndGroups_Exists Then
            Globals.ThisAddIn.AddTaskPane_UsersAndGroups()
        Else
            Globals.ThisAddIn.RemoveTaskPane_UsersAndGroups()
        End If

        _TaskPane_UsersAndGroups_Exists = Not _TaskPane_UsersAndGroups_Exists
    End Sub
#End Region

    Private _TaskPane_Webs_Exists As Boolean = False

    Public Sub TaskPane_Webs_Click(ByVal control As Office.IRibbonControl)
        If Not _TaskPane_Webs_Exists Then
            Globals.ThisAddIn.AddTaskPane_Webs()
        Else
            Globals.ThisAddIn.RemoveTaskPane_Webs()
        End If

        _TaskPane_Webs_Exists = Not _TaskPane_Webs_Exists
    End Sub

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
