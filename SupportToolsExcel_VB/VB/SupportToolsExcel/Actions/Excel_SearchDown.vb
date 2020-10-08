Option Strict Off

Imports System.Reflection
Imports System.Windows.Forms

Imports AddinHelper
Imports Microsoft.Office.Core
Imports PacificLife.Life

''' <summary>
''' Excel_SearchDown
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
Public Class Excel_SearchDown
    Inherits AddinHelper.AppMethod

#Region "General Addin Constants"

    Private Const _MODULE_NAME As String = Common.PROJECT_NAME & "Excel_SearchDown"
    Private Const _NAME As String = "Excel_SearchDown"
    Private Const _BITMAP_NAME As String = "search down.bmp"
    Private Const _CAPTION As String = "Search Down"
    Private Const _TOOL_TIP_TEXT As String = "Search Down"
    Private Const _DESCRIPTION As String = "Excel_SearchDown does ..."
#End Region

#Region "Private Constants and Ennumerations"

    Private Const _ERROR_EMPTY_CELL As String = "Cell is empty.  Must select a popluated starting cell first."
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

    '--------------------------------------------------------------------------------
    '
    ' FindEndOfRangeDown()
    '
    ' Search down from current cell looking for end of current range.
    ' End of current range is the next occupied cell in the current section
    ' with no intervening outdented sections.  Current range ends just before
    ' subsequent outdented sections or on last occupied row.
    '
    '--------------------------------------------------------------------------------

    Public Shared Sub FindEndOfRangeDown()
        Dim intStartRow As Integer
        Dim intStartCol As Integer
        Dim intLastRow As Integer
        Dim intLastCol As Integer
        Dim intEndRowOfSection As Integer

        With Globals.ThisAddIn.Application
            If IsNothing(.ActiveCell.Value) Then
                MsgBox(_ERROR_EMPTY_CELL)
            Else
                ' Get the last populated cell on the worksheet.
                intLastRow = .ActiveSheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell).Row()
                intLastCol = .ActiveSheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell).Column()

                ' Save where we currently are located.
                intStartRow = .ActiveCell.Row
                intStartCol = .ActiveCell.Column

                intEndRowOfSection = Excel_FolderMaps.GetEndOfSectionDown(intStartRow, intStartCol, intLastRow)
                .Cells(intEndRowOfSection, intStartCol).Select()
            End If
        End With
    End Sub

#End Region

#Region "Private Methods"
    
    Private Sub Action(ByVal Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean)
        Try
            FindEndOfRangeDown()
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
