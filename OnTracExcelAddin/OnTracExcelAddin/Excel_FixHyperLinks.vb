Option Strict Off

Imports System.Reflection

Imports AddinHelper
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop
Imports PacificLife.Life

'Imports ENLog = Microsoft.Practices.EnterpriseLibrary.Logging

''' <summary>
''' Action1
''' </summary>
''' <remarks>To Use this class modify the class name and change the constants.
''' Update the Action Method with code that does something useful.</remarks>
Public Class Excel_FixHyperLinks
    Inherits AddinHelper.AppMethod

#Region "Private Variables"

    Private Const _MODULE_NAME As String = Globals.PROJECT_NAME & "Excel_FixHyperLinks"
    Private Const _NAME As String = "Excel_FixHyperLinks"
    Private Const _BITMAP_NAME As String = "FixHyperLinks.bmp"
    Private Const _CAPTION As String = "Fix HyperLinks"
    Private Const _TOOL_TIP_TEXT As String = "Fix HyperLinks"
    Private Const _DESCRIPTION As String = "Excel_FixHyperLinks does ..."

    Private Const _ERROR_EMPTY_CELL As String = "Cell is empty.  Must select a popluated starting cell first."
    Private Const _ERROR_NOHYPERLINK_CELL As String = "Cell contains no HyperLinks.  Must select one or more cells with HyperLinks first.  Continue with other cells?"
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
        FixHyperLinks()
    End Sub

    Shared Sub FixHyperLinks()
        Dim selectedCells As Excel.Range = Globals.ThisAddIn.Application.Selection

        With Globals.ThisAddIn.Application
            If IsNothing(.ActiveCell.Value) Then
                MsgBox(_ERROR_EMPTY_CELL)
            Else
                Try
                    For Each cell As Excel.Range In selectedCells
                        'DisplayCellInfo(cell)     

                        Select Case (cell.Hyperlinks.Count)
                            Case 0
                                Dim result As DialogResult = MessageBox.Show(_ERROR_NOHYPERLINK_CELL, "Invalid Cell", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                                If result = DialogResult.Yes Then
                                    Continue For
                                Else
                                    Return
                                End If

                            Case 1
                                RepairLink(cell.Hyperlinks.Item(1))

                            Case Else
                                MessageBox.Show(String.Format("Cell R{0}C{1} contains more than one HyperLink!", cell.Row, cell.Column), "Unexpected Cell Contents", MessageBoxButtons.OK, MessageBoxIcon.Error)

                        End Select

                    Next

                Catch ex As Exception
                    'PLLog.Error(ex, "FixHyperLinks")
                    Throw (ex)
                End Try
            End If
        End With
    End Sub

    Public Shared Sub RepairLink(ByRef link As Excel.Hyperlink)
        Dim currentAddress As String = link.Address
        Dim startOfLink As Integer = currentAddress.IndexOf("'", 0)
        Dim endOfLink As Integer = currentAddress.IndexOf("'", startOfLink + 1)

        ' Skip past the "'"
        startOfLink += 1

        Dim updatedAddress As String = currentAddress.Substring(startOfLink, endOfLink - startOfLink)
        link.Address = updatedAddress
    End Sub
#End Region
End Class
