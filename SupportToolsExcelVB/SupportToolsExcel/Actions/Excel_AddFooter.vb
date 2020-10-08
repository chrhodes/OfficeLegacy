Option Strict Off
Imports System.Reflection

Imports AddinHelper
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Text

''' <summary>
''' Excel_AddFooter
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
Public Class Excel_AddFooter
    Inherits AddinHelper.AppMethod

#Region "Private Constants and Variables"

    Private Const _MODULE_NAME As String = Common.PROJECT_NAME & "AddFooter"
    Private Const _NAME As String = "AddFooter"
    Private Const _BITMAP_NAME As String = "add footer.bmp"
    Private Const _CAPTION As String = "AddFooter"
    Private Const _TOOL_TIP_TEXT As String = "Click to Add Footer"
    Private Const _DESCRIPTION As String = "AddFooter does ..."

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

    '------------------------------------------------------------
    '
    ' Sub Footer_Add()
    '
    ' ToDo:
    '   Display dialog box that allow format choices.
    '   See Format page for ideas.
    '------------------------------------------------------------

    Public Shared Sub AddFooter()
        Dim workBook As Excel.Workbook
        Dim workSheet As Excel.Worksheet
        Dim sb As StringBuilder = New StringBuilder

        Try
            With Globals.ThisAddIn.Application
                workBook = .ActiveWorkbook

                Common.ExcelHelper.CalculationsOff()

                For Each workSheet In workBook.Sheets
                    With workSheet.PageSetup
                        sb.Length = 0
                        ' Five point font, path, and filename
                        sb.Append("&5&Z&F")
                        sb.Append(vbLf & "Created: ")
                        sb.Append(Common.ExcelHelper.GetBuiltInPropertyValue(workBook, "Creation Date"))
                        sb.Append(vbLf & "Last Saved: ")
                        sb.Append(Common.ExcelHelper.GetBuiltInPropertyValue(workBook, "Last Save Time"))
                        sb.Append(vbLf & "Last Printed: ")
                        sb.Append(Common.ExcelHelper.GetBuiltInPropertyValue(workBook, "Last Print Date"))
                        'Debug.Print("LeftFooter:>" & sb.ToString & "<")
                        .LeftFooter = sb.ToString()

                        .CenterFooter = ""

                        sb.Length = 0
                        ' Five point font, page of pages
                        sb.Append("&5&P - &N")
                        sb.Append(vbLf & "Title :")
                        sb.Append(Common.ExcelHelper.GetBuiltInPropertyValue(workBook, "Title"))
                        sb.Append(vbLf & "Subject: ")
                        sb.Append(Common.ExcelHelper.GetBuiltInPropertyValue(workBook, "Subject"))
                        'Debug.Print("RightFooter:>" & sb.ToString & "<")
                        .RightFooter = sb.ToString()
                    End With

                    ' TODO: Indicate we have added a custom footer.  This will be looked for
                    ' in the before close event.

                    If Not Common.ExcelHelper.HasCustomFooter() Then
                        Common.ExcelHelper.CustomFooterExists(True)
                    End If

                Next

                Common.ExcelHelper.CalculationsOn()
            End With
        Catch ex As Exception
            MsgBox("AddFooter:" & ex.ToString)
        End Try
    End Sub ' Footer_Add

    'Friend Shared Function HasCustomFooter() As Boolean
    '    Dim prp As Office.DocumentProperty
    '    Dim prps As Office.DocumentProperties

    '    Try
    '        Try
    '            prps = Globals.ThisAddIn.Application.ActiveWorkbook.CustomDocumentProperties
    '            prp = prps.Item("HasCustomFooter")
    '            ' If the property exists we don't really care about the value
    '            Return True

    '        Catch ex As Exception
    '            ' Exception is thrown if property does not exist
    '            Return False
    '        End Try
    '    Finally

    '    End Try
    'End Function

    'Friend Shared Sub CustomFooterExists(ByVal hasCustomFooter As Boolean)
    '    Dim prp As Office.DocumentProperty
    '    Dim prps As Office.DocumentProperties

    '    Try
    '        Try
    '            prps = Globals.ThisAddIn.Application.ActiveWorkbook.CustomDocumentProperties
    '            ' Add a new property.
    '            prp = prps.Add("HasCustomFooter", False, _
    '             Office.MsoDocProperties.msoPropertyTypeBoolean, True)
    '        Catch ex As Exception
    '            'PLLog.Error(ex, Globals.PROJECT_NAME)
    '            MessageBox.Show("CustomFooterExists() Unable to add HasCustomFooter property" & ex.Message)
    '        End Try
    '    Finally

    '    End Try
    'End Sub

#End Region

#Region "Private Methods"

    Private Sub Action(ByVal Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean)
        Try
            AddFooter()
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
