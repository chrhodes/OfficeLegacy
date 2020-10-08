Imports System.Reflection

Imports AddinHelper
Imports Microsoft.Office.Core

Public Class Action2
    Inherits AddinHelper.AppMethod

    Private Const _MODULE_NAME As String = Globals.PROJECT_NAME & "Action2"
    Private Const _NAME As String = "Action2"
    Private Const _BITMAP_NAME As String = "Action2.bmp"
    Private Const _CAPTION As String = "Action2"
    Private Const _TOOL_TIP_TEXT As String = "Click for Action2"
    Private Const _DESCRIPTION As String = "Action2 does ..."

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
        MsgBox("Totaly Cool Action 2")
    End Sub
End Class
