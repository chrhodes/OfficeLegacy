Imports System.Reflection

Imports AddinHelper
Imports Microsoft.Office.Core

''' <summary>
''' Word_WatchWindow
''' </summary>
''' <remarks>To Use this class modify the class name and change the constants.
''' Update the Action Method with code that does something useful.</remarks>
Public Class Word_WatchWindow
    Inherits AddinHelper.AppMethod

#Region "Private Variables"

    Private Const _MODULE_NAME As String = Globals.PROJECT_NAME & "Word_WatchWindow"
    Private Const _NAME As String = "Word_WatchWindow"
    Private Const _BITMAP_NAME As String = "WatchWindow.bmp"
    Private Const _CAPTION As String = "Word_WatchWindow"
    Private Const _TOOL_TIP_TEXT As String = "Click for Word_WatchWindow"
    Private Const _DESCRIPTION As String = "Word_WatchWindow does ..."

#End Region

#Region "Public Methods"

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="commandBar">Which bar to add method to</param>
    ''' <param name="buttonStyle">The type of button to put on the bar</param>
    ''' <remarks></remarks>
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
#End Region

#Region "Private Methods"

    Private Sub Action(ByVal Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean)
        If Globals.ThisAddIn.AppEventsWatchWindow Is Nothing Then
            Globals.ThisAddIn.AppEventsWatchWindow = New AddinHelper.WatchWindow
            Globals.ThisAddIn.AppEventsWatchWindow.Show()
        Else
            Globals.ThisAddIn.AppEventsWatchWindow.Close()
            Globals.ThisAddIn.AppEventsWatchWindow = Nothing
        End If
    End Sub
#End Region
End Class
