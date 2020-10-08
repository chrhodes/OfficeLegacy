Option Strict Off

Imports System.IO
Imports System.Reflection
Imports System.Windows.Forms

Imports AddinHelper
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel
'Imports PacificLife.Life.Enterprise

''' <summary>
''' Excel_TableOfContents
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
Public Class Excel_TableOfContents
    Inherits AddinHelper.AppMethod

#Region "Private Constants and Variables"

    Private Const _MODULE_NAME As String = Common.PROJECT_NAME & "Excel_TableOfContents"
    Private Const _NAME As String = "Excel_TableOfContents"
    Private Const _BITMAP_NAME As String = "table of contents.bmp"
    Private Const _CAPTION As String = "Table Of Contents"
    Private Const _TOOL_TIP_TEXT As String = "Create/Update Table of Contents"
    Private Const _DESCRIPTION As String = "Create Table of Contents does ..."

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

    Public Shared Sub CreateTableOfContents()
        With Globals.ThisAddIn.Application
            Dim sht As Worksheet
            Dim shtTableOfContents As Worksheet
            Dim intRow As Integer
            Dim intCol As Integer
            Dim currentSheetProtectionMode As Boolean

            intRow = 3  ' Starting Row
            intCol = 1  ' Starting Col

            Try
                .ActiveWorkbook.Sheets("Table of Contents").Activate()
            Catch ex As Exception
                .ActiveWorkbook.Sheets.Add()
                .ActiveSheet.Name = "Table of Contents"
            End Try

            .ActiveSheet.Columns("A:A").Select()
            .Selection.ClearContents()

            shtTableOfContents = .ActiveSheet

            For Each sht In .ActiveWorkbook.Sheets
                Select Case sht.Name
                    Case "Table of Contents"

                    Case Else
                        With sht
                            ' Unprotect the sheet before adding the hyperlink

                            currentSheetProtectionMode = Common.ExcelHelper.ProtectSheet(sht, False)
                            .Cells(1, 1).Value = "Table of Contents"

                            .Hyperlinks.Add( _
                                Anchor:=sht.Cells(1, 1), _
                                Address:="", _
                                SubAddress:="'" & shtTableOfContents.Name & "'!A1", _
                                TextToDisplay:=shtTableOfContents.Name)

                            ' Then restore the prior setting
                            Common.ExcelHelper.ProtectSheet(sht, currentSheetProtectionMode)
                        End With

                        With shtTableOfContents
                            .Cells(intRow, intCol).Value = sht.Name

                            .Hyperlinks.Add( _
                                Anchor:=.Cells(intRow, intCol), _
                                Address:="", _
                                SubAddress:="'" & sht.Name & "'!A1", _
                                TextToDisplay:=sht.Name)
                        End With

                        intRow = intRow + 1
                End Select
            Next sht

            shtTableOfContents.Columns("A:A").EntireColumn.AutoFit()
        End With
    End Sub
#End Region

#Region "Private Methods"

    Private Sub Action(ByVal Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean)
        Try
            CreateTableOfContents()
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
