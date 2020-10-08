Imports System.Reflection
Imports System.Windows.Forms

Imports Microsoft.Office.Core

''' <summary>
''' Word_GetSiteInfo
''' </summary>
''' <remarks>
''' This class can be used in two ways.  If calling this from a commandBar, modify
''' the Private Const as needed and then create an instance of this class in the code
''' that loads the command bars.
''' 
''' If calling this from a Ribbon Event handler, call the ActionNameGoesHere method directly.
''' 
''' Rename the ActionNameGoesHere Method and provide code that does something useful.
''' </remarks>
Public Class Word_GetSiteInfo
    Inherits AddinHelper.AppMethod

#Region "Private Constants and Variables"

    Private Const _MODULE_NAME As String = Common.PROJECT_NAME & "GetSiteInfo"
    Private Const _NAME As String = "GetSiteInfo"
    Private Const _BITMAP_NAME As String = "GetSiteInfo.bmp"
    Private Const _CAPTION As String = "SiteInfo"
    Private Const _TOOL_TIP_TEXT As String = "Get information from SharePoint site"
    Private Const _DESCRIPTION As String = "GetSiteInfo does ..."

#End Region

#Region "Public Methods"

    Public Sub New(ByRef commandBar As CommandBar, ByRef buttonStyle As MsoButtonStyle)
        MyBase.Name = _NAME
        MyBase.CommandBar = commandBar
        MyBase.EventHandler = AddressOf Action
        MyBase.ButtonStyle = buttonStyle
        MyBase.BitMapName = _BITMAP_NAME
        MyBase.Asmbly = [Assembly].GetExecutingAssembly
        MyBase.Caption = _CAPTION
        MyBase.ToolTipText = _TOOL_TIP_TEXT
        MyBase.Description = _DESCRIPTION

        MyBase.Initialize()
    End Sub

    Public Shared Sub GetSiteInfo()
        MessageBox.Show("GetSiteInfo coming soon")
    End Sub

#End Region

#Region "Private Methods"

    Private Sub Action(ByVal Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean)
        Try
            Me.GetSiteInfo()
        Catch ex As Exception

        End Try
    End Sub

#End Region

End Class
