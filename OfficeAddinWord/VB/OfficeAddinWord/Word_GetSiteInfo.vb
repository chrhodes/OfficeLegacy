Imports System.Reflection

Imports AddinHelper
Imports Microsoft.Office.Core

Public Class Word_GetSiteInfo
    Inherits AddinHelper.AppMethod

    Private Const _MODULE_NAME As String = Globals.PROJECT_NAME & "GetSiteInfo"
    Private Const _NAME As String = "GetSiteInfo"
    Private Const _BITMAP_NAME As String = "GetSiteInfo.bmp"
    Private Const _CAPTION As String = "SiteInfo"
    Private Const _TOOL_TIP_TEXT As String = "Get information from SharePoint site"
    Private Const _DESCRIPTION As String = "GetSiteInfo does ..."

    '**********************************************************************
    '   P u b l i c    M e t h o d s
    '**********************************************************************

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

    Private Sub Action(ByVal Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean)
        Try
            Me.GetSiteInfo()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub GetSiteInfo()
        MessageBox.Show("GetSiteInfo coming soon")
    End Sub
End Class
