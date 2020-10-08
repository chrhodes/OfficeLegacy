Imports System.Reflection

Imports AddinHelper
Imports Microsoft.Office.Core
Imports System.Text

Public Class AddinInfo
    Inherits AddinHelper.AppMethod

    Private Const _MODULE_NAME As String = Globals.PROJECT_NAME & "AddinInfo"
    Private Const _NAME As String = "AddinInfo"
    Private Const _BITMAP_NAME As String = "AddinInfo.bmp"
    Private Const _CAPTION As String = "AddinInfo"
    Private Const _TOOL_TIP_TEXT As String = "Addin Information"
    Private Const _DESCRIPTION As String = "AddinInfo does ..."

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
        Dim sb As New StringBuilder(100)

        sb.AppendLine(String.Format("ProductName:   {0}", My.Application.Info.ProductName))
        sb.AppendLine(String.Format("Title:         {0}", My.Application.Info.Title))
        sb.AppendLine(String.Format("AssemblyName:  {0}", My.Application.Info.AssemblyName))
        sb.AppendLine(String.Format("DirectoryPath: {0}", My.Application.Info.DirectoryPath))
        sb.AppendLine(String.Format("Version:       {0}", My.Application.Info.Version))
        sb.AppendLine(String.Format("Description:   {0}", My.Application.Info.Description))

        MsgBox(sb.ToString())
    End Sub
End Class
