Imports System.Reflection
Imports System.Windows.Forms
Imports AddinHelper
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel

''' <summary>
''' Excel_AllPortrait
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
Public Class Excel_AllPortrait
    Inherits AddinHelper.AppMethod

#Region "Private Constants and Variables"

    Private Const _MODULE_NAME As String = Common.PROJECT_NAME & "Excel_AllPortrait"
    Private Const _NAME As String = "Excel_AllPortrait"
    Private Const _BITMAP_NAME As String = "format all portrait.bmp"
    Private Const _CAPTION As String = "All Portrait"
    Private Const _TOOL_TIP_TEXT As String = "All worksheets Portrait"
    Private Const _DESCRIPTION As String = "Excel_AllPortrait does ..."

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

    Public Shared Sub AllPortrait()
        Dim workSheet As Worksheet

        With Globals.ThisAddIn.Application
            For Each workSheet In .ActiveWorkbook.Sheets
                workSheet.PageSetup.Orientation = XlPageOrientation.xlPortrait
            Next
        End With
    End Sub
#End Region

#Region "Private Methods"

    Private Sub Action(ByVal Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean)
        Try
            AllPortrait()
        Catch ex As Exception
            MessageBox.Show(String.Format("Exception: {0}.{1}() - {2}",
                         System.Reflection.Assembly.GetExecutingAssembly().FullName,
                         System.Reflection.MethodInfo.GetCurrentMethod().Name,
                         ex.ToString()
                         ))
        End Try
    End Sub

#End Region
End Class
