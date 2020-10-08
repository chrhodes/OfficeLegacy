Option Strict Off

Imports System.Environment
Imports System.Reflection
Imports System.Text

Imports AddinHelper
Imports Microsoft.Office.Core

Public Class EnvironmentInfo
    Inherits AddinHelper.AppMethod

    Private Const _MODULE_NAME As String = Globals.PROJECT_NAME & "EnvironmentInfo"
    Private Const _NAME As String = "EnvironmentInfo"
    Private Const _BITMAP_NAME As String = "Action1.bmp"
    Private Const _CAPTION As String = "EnvInfo"
    Private Const _TOOL_TIP_TEXT As String = "Environment Info"
    Private Const _DESCRIPTION As String = "EnvironmentInfo does ..."

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
            Dim sb As StringBuilder = New StringBuilder

            sb.AppendLine("Current Directory: " & System.Environment.CurrentDirectory())
            'sb.AppendLine("Environment Variables: " & System.Environment.GetEnvironmentVariables())
            sb.AppendLine("UserDomainName: " & System.Environment.UserDomainName())
            sb.AppendLine("UserName: " & System.Environment.UserName())
            sb.AppendLine("CLR Version: " & System.Environment.Version.ToString())

            MsgBox(sb.ToString())
        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try
    End Sub
End Class
