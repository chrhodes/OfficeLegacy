Option Explicit On

Imports System.Collections
Imports System.Collections.Generic
Imports System.Diagnostics

Namespace Util

    ''' <summary>
    ''' Contains general purpose Excel Helper routines.
    ''' Depends on Globals.ExcelApp being intialized to current Excel Application.
    ''' This is typically done in Startup code.
    ''' </summary>
    ''' <remarks></remarks>
    Public Class ExcelHelper
        '**********************************************************************
        '   C o n s t a n t s
        '**********************************************************************

        Const cPointsToInch As Integer = 72    ' Hard coded for Excel? 1/72 Inches.
        Const cStartHour As Integer = 8        ' Times before this
        Const cEndHour As Integer = 20         ' and after this are hilighted.

        '**********************************************************************
        '   P u b l i c    T y p e s    a n d    V a r i a b l e s
        '**********************************************************************


        '**********************************************************************
        '   P r i v a t e    T y p e s    a n d    V a r i a b l e s
        '**********************************************************************

        'Private m_vntPriorCalculationState As Object
        'Private m_vntPriorScreenUpdatingState As Object

        '**********************************************************************
        '   P u b l i c    M e t h o d s
        '**********************************************************************

        Public Shared Sub AddColumnToSheet( _
        ByRef ws As Excel.Worksheet, _
        ByVal columnNumber As Integer, _
        ByVal columnWidth As Integer, _
        ByVal columnWrapText As Boolean, _
        ByVal headerRow As Integer, _
        Optional ByVal headerTitle As String = "", _
        Optional ByVal headerFontSize As Integer = Globals.cHeaderFontSize, _
        Optional ByVal headerBold As Boolean = True, _
        Optional ByVal headerUnderline As Boolean = True, _
        Optional ByVal headerWrapText As Boolean = True, _
        Optional ByVal headerHorizontalAlignment As Excel.XlHAlign = Excel.XlHAlign.xlHAlignGeneral, _
        Optional ByVal orientation As Integer = 0 _
    )
            With ws
                .Columns(columnNumber).ColumnWidth = columnWidth
                .Columns(columnNumber).WrapText = columnWrapText

                If headerTitle <> "" Then
                    With .Cells(headerRow, columnNumber)
                        .Value = headerTitle
                        .Font.Size = headerFontSize
                        .Font.Bold = headerBold
                        .Font.Underline = headerUnderline
                        .WrapText = headerWrapText
                        .HorizontalAlignment = headerHorizontalAlignment
                        .Orientation = orientation
                    End With
                End If
            End With

        End Sub

        Public Shared Sub AddCommentToCell( _
            ByRef ws As Excel.Worksheet, _
            ByVal column As Integer, _
            ByVal row As Integer, _
            ByVal text As String, _
            Optional ByVal headerFontSize As Integer = Globals.cHeaderFontSize, _
            Optional ByVal headerBold As Boolean = True, _
            Optional ByVal headerUnderline As Boolean = True, _
            Optional ByVal headerWrapText As Boolean = True, _
            Optional ByVal headerHorizontalAlignment As Excel.XlHAlign = Excel.XlHAlign.xlHAlignGeneral _
        )
            ws.Cells(row, column).AddComment(text)
            ' TODO: Determine how to format the text differently.
            'With ws
            '    With .Cells(row, column)
            '        .Value = headerTitle
            '        .Font.Size = headerFontSize
            '        .Font.Bold = headerBold
            '        .Font.Underline = headerUnderline
            '        .WrapText = headerWrapText
            '        .HorizontalAlignment = headerHorizontalAlignment
            '    End With
            'End With

        End Sub

        Public Shared Sub AddContentToCell( _
            ByVal rng As Excel.Range, _
            ByVal text As String, _
            Optional ByVal fontSize As Integer = 10, _
            Optional ByVal bold As Boolean = False, _
            Optional ByVal underline As Boolean = False, _
            Optional ByVal wrapText As Boolean = False, _
            Optional ByVal horizontalAlignment As Excel.XlHAlign = Excel.XlHAlign.xlHAlignGeneral _
        )
            With rng
                .Value = text
                .Font.Size = fontSize
                .Font.Bold = bold
                .Font.Underline = underline
                .WrapText = wrapText
                .HorizontalAlignment = horizontalAlignment
            End With

        End Sub

        Public Shared Sub ApplicationInfo()
            Try
                Debug.Print("Application.CommonAppDataPath:" & Application.CommonAppDataPath.ToString)
                Debug.Print("Application.CommonAppDataRegistry:" & Application.CommonAppDataRegistry.ToString)
                Debug.Print("Application.CompanyName:" & Application.CompanyName.ToString)
                Debug.Print("Application.CurrentCulture:" & Application.CurrentCulture.ToString)
                Debug.Print("Application.CurrentInputLanguage:" & Application.CurrentInputLanguage.ToString)
                Debug.Print("Application.ExecutablePath:" & Application.ExecutablePath.ToString)
                Debug.Print("Application.LocalUserAppDataPath:" & Application.LocalUserAppDataPath.ToString)
                Debug.Print("Application.ProductName:" & Application.ProductName.ToString)
                Debug.Print("Application.ProductVersion:" & Application.ProductVersion.ToString)
                Debug.Print("Application.SafeTopLevelCaptionFormat:" & Application.SafeTopLevelCaptionFormat.ToString)
                Debug.Print("Application.StartupPath:" & Application.StartupPath.ToString)
                Debug.Print("Application.UserAppDataPath:" & Application.UserAppDataPath.ToString)
                Debug.Print("Application.UserAppDataRegistry:" & Application.UserAppDataRegistry.ToString)

                Debug.Print("ThisAddin.Application.StartupPath:" & Globals.ExcelApp.Application.StartupPath.ToString)
                Debug.Print("ThisAddin.Application.ActiveWorkbook.Name:" & Globals.ExcelApp.Application.ActiveWorkbook.Name.ToString)
                Debug.Print("ThisAddin.Application.ActiveWorkbook.Path:" & Globals.ExcelApp.Application.ActiveWorkbook.Path.ToString)
                Debug.Print("ThisAddin.Application.ActiveWorkbook.FullName:" & Globals.ExcelApp.Application.ActiveWorkbook.FullName.ToString)
                Debug.Print("ThisAddin.Application.ActiveWorkbook.FullNameURLEncoded:" & Globals.ExcelApp.Application.ActiveWorkbook.FullNameURLEncoded.ToString)

                Debug.Print("ThisAddin.Application.DefaultFilePath:" & Globals.ExcelApp.Application.DefaultFilePath.ToString)
                Debug.Print("ThisAddin.Application.Name:" & Globals.ExcelApp.Application.Name.ToString)
                Debug.Print("ThisAddin.Application.NetworkTemplatesPath:" & Globals.ExcelApp.Application.NetworkTemplatesPath.ToString)
                Debug.Print("ThisAddin.Application.Path:" & Globals.ExcelApp.Application.Path.ToString)
            Catch ex As Exception
                MessageBox.Show("ApplicationInfo():" & ex.ToString)
            End Try
        End Sub

        Public Shared Function BaseName(ByVal strName As String) As String
            BaseName = Left(strName, InStr(1, strName, ".", vbTextCompare) - 1)
        End Function

        Public Shared Sub CalculationsOff()
            ' Don't bother trying to save current if no open workbooks.

            With Globals.ExcelApp.Application
                If .Workbooks.Count > 0 Then
                    Globals.ExcelApp.m_vntPriorCalculationState = .Calculation
                    .Calculation = Excel.XlCalculation.xlCalculationManual
                Else
                    ' Assume the intent is to run with calculation and screen updates on.
                    ' Hopefully we never get called with no workbooks open.
                    Globals.ExcelApp.m_vntPriorCalculationState = Excel.XlCalculation.xlCalculationAutomatic
                End If
            End With
        End Sub ' CalculationsOff

        Public Shared Sub CalculationsOn()
            With Globals.ExcelApp.Application
                .Calculation = Globals.ExcelApp.m_vntPriorCalculationState
            End With
        End Sub ' CalculationsOn

        Public Shared Sub DisplayNames(ByVal ws As Excel.Worksheet)
            For Each name As Excel.Name In ws.Names
                Debug.Print(name.Category)
                Debug.Print(name.Comment)
                Debug.Print(name.Index)
                Debug.Print(name.Name)
                Debug.Print(name.RefersTo.ToString)
                Debug.Print(name.RefersToRange.ToString)
                Debug.Print(name.ToString)
                Debug.Print(name.Value)
                Debug.Print(name.WorkbookParameter)
            Next
        End Sub

        Public Shared Sub DisplayNames(ByVal wb As Excel.Workbook)
            For Each name As Excel.Name In wb.Names
                Debug.Print("Comment: " & name.Comment)
                Debug.Print("Index: " & name.Index)

                Select Case name.MacroType
                    Case Excel.XlXLMMacroType.xlCommand
                        Debug.Print("MacroType: xlCommand")

                    Case Excel.XlXLMMacroType.xlFunction
                        Debug.Print("MacroType: xlFunction")

                    Case Excel.XlXLMMacroType.xlNotXLM
                        Debug.Print("MacroType: xlNotXLM")

                End Select

                Debug.Print("Name: " & name.Name)
                Debug.Print("RefersTo: " & name.RefersTo.ToString)
                'Debug.Print(name.RefersToRange.ToString)
                'Debug.Print(name.ToString)
                Debug.Print("Value: " & name.Value)
                Debug.Print("WorkbookParamter: " & name.WorkbookParameter)
            Next
        End Sub


        Public Shared Sub DeleteSheet(ByVal ws As Excel.Worksheet, Optional ByVal prompt As Boolean = False)
            Dim priorState As Boolean

            priorState = Globals.ExcelApp.Application.DisplayAlerts()

            If prompt Then
                Globals.ExcelApp.Application.DisplayAlerts = True
                ws.Delete()
            Else
                Globals.ExcelApp.Application.DisplayAlerts = False
                ws.Delete()

            End If

            Globals.ExcelApp.Application.DisplayAlerts = priorState
        End Sub

        Public Shared Sub DisplayExcelRange(ByVal rng As Excel.Range)
            Debug.Print(rng.Address)
            Debug.Print("Address: " & rng.Address & " Rows: " & rng.Rows.Count & " Columns: " & rng.Columns.Count)

            For Each c As Excel.Range In rng.Cells
                Debug.Print("Value: >" & c.Value & "< Row: " & c.Row & " Col: " & c.Column)
            Next
        End Sub

        Public Shared Sub DisplayListObjects(ByVal ws As Excel.Worksheet)
            For Each lo As Excel.ListObject In ws.ListObjects
                Debug.Print(lo.Comment)
                Debug.Print(lo.DataBodyRange.ToString)
                Debug.Print(lo.DisplayName)
                Debug.Print(lo.HeaderRowRange.ToString)
                Debug.Print(lo.InsertRowRange.ToString)
                Debug.Print(lo.Name)
                Debug.Print(lo.Range.ToString)
                Debug.Print(lo.SharePointURL)
                Debug.Print(lo.ToString)
                Debug.Print(lo.TotalsRowRange.ToString)
            Next
        End Sub

        Public Shared Sub FindLast()
            Dim currentCellRange As Excel.Range
            Dim currentRowRange As Excel.Range
            Dim currentColumnRange As Excel.Range
            Dim lastRow As Long
            Dim lastColumn As Long

            currentCellRange = Globals.ExcelApp.Application.ActiveCell
            currentRowRange = Globals.ExcelApp.Application.ActiveSheet.Rows.Item(currentCellRange.Row)
            currentColumnRange = Globals.ExcelApp.Application.ActiveSheet.Columns.Item(currentCellRange.Column)

            lastRow = currentCellRange.Find("*", , , , Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious).Row
            lastColumn = currentCellRange.Find("*", , , , Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious).Column

            MsgBox("Row=" & lastRow & " Column=" & lastColumn, vbOKOnly, "Cell Find")

            lastRow = currentRowRange.Find("*", , , , Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious).Row
            lastColumn = currentRowRange.Find("*", , , , Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious).Column

            MsgBox("Row=" & lastRow & " Column=" & lastColumn, vbOKOnly, "Row Find")

            lastRow = currentColumnRange.Find("*", , , , Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious).Row
            lastColumn = currentColumnRange.Find("*", , , , Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious).Column

            MsgBox("Row=" & lastRow & " Column=" & lastColumn, vbOKOnly, "Column Find")
        End Sub

        Public Shared Function FindLastColumn(ByVal searchFromCell As Excel.Range) As Long
            Dim currentRowRange As Excel.Range
            'Dim currentColumnRange As Excel.Range
            'Dim lastRow As Long
            Dim lastColumn As Long

            If searchFromCell Is Nothing Then
                MessageBox.Show("FindLastColumn(): searchFromCell is Nothing")
            Else
                Try
                    currentRowRange = Globals.ExcelApp.Application.ActiveSheet.Rows.Item(searchFromCell.Row)
                    'currentColumnRange = Globals.ThisAddIn.Application.ActiveSheet.Columns.Item(searchFromCell.Column)

                    'lastRow = currentRowRange.Find("*", , , , Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious).Row
                    lastColumn = currentRowRange.Find("*", , , , Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious).Column

                    'MsgBox("Row=" & lastRow & " Column=" & lastColumn, vbOKOnly, "Row Find")

                    'lastRow = currentColumnRange.Find("*", , , , Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious).Row
                    'lastColumn = currentColumnRange.Find("*", , , , Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious).Column

                    'MsgBox("Row=" & lastRow & " Column=" & lastColumn, vbOKOnly, "Column Find")
                Catch ex As Exception
                    MessageBox.Show(ex.ToString())
                End Try
            End If

            Return lastColumn
        End Function

        Public Shared Function FindLastRow(ByVal searchFromCell As Excel.Range) As Long
            'Dim currentCellRange As Excel.Range
            'Dim currentRowRange As Excel.Range
            Dim currentColumnRange As Excel.Range
            Dim lastRow As Long
            'Dim lastColumn As Long

            If searchFromCell Is Nothing Then
                MessageBox.Show("FindLastRow(): searchFromCell is Nothing")
            Else
                Try
                    'currentRowRange = Globals.ThisAddIn.Application.ActiveSheet.Rows.Item(searchFromCell.Row)
                    currentColumnRange = Globals.ExcelApp.Application.ActiveSheet.Columns.Item(searchFromCell.Column)

                    'lastRow = currentRowRange.Find("*", , , , Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious).Row
                    'lastColumn = currentRowRange.Find("*", , , , Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious).Column

                    'MsgBox("Row=" & lastRow & " Column=" & lastColumn, vbOKOnly, "Row Find")

                    lastRow = currentColumnRange.Find("*", , , , Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious).Row
                    'lastColumn = currentColumnRange.Find("*", , , , Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious).Column

                    'MsgBox("Row=" & lastRow & " Column=" & lastColumn, vbOKOnly, "Column Find")
                Catch ex As Exception
                    MessageBox.Show(ex.ToString())
                End Try
            End If

            Return lastRow
        End Function

        Public Shared Function GetFile( _
            Optional ByVal initialFolder As String = "", _
            Optional ByVal dialogTitle As String = "Open", _
            Optional ByVal fileFilter As String = "All Files (*.*)|*.*") As String
            Dim ofd As New OpenFileDialog
            Dim result As System.Windows.Forms.DialogResult

            ofd.Multiselect = False
            ofd.InitialDirectory = initialFolder
            ofd.Title = dialogTitle
            ofd.Filter = fileFilter

            result = ofd.ShowDialog()

            Debug.WriteLine(ofd.FileName)

            Return ofd.FileName

        End Function

        Public Shared Function ProtectSheet( _
            ByRef sht As Microsoft.Office.Interop.Excel.Worksheet, _
            ByVal protectMode As Boolean _
        ) As Boolean
            If protectMode = True Then
                sht.Protect()
            Else
                sht.Unprotect()
            End If

            Return sht.ProtectionMode
        End Function

        Function SafeName(ByVal strS As String) As String
            Dim strSafe As String

            strSafe = Replace(strS, "/", " ")
            SafeName = strSafe
        End Function

        Public Shared Function NewWorksheet( _
            ByVal sheetName As String, _
            Optional ByVal beforeSheetName As String = "", _
            Optional ByVal afterSheetName As String = "" _
        ) As Excel.Worksheet

            With Globals.ThisWorkbook.Application.ActiveWorkbook

                For Each ws As Excel.Worksheet In .Worksheets
                    If ws.Name = sheetName Then
                        ' Sheet exists.  Ask user what to do.
                        MessageBox.Show("Sheet: >" & sheetName & "< already exists.")
                        Return ws
                    End If
                Next

                If beforeSheetName <> "" Then
                    .Sheets.Add(.Sheets(beforeSheetName))
                ElseIf afterSheetName <> "" Then
                    .Sheets.Add(, .Sheets(afterSheetName))
                Else
                    .Sheets.Add()
                End If

                .ActiveSheet.Name = sheetName

                Return .ActiveSheet
            End With

        End Function

        '' Do this with regular expressions.

        'Function SafeSheetName(ByVal strName As String) As String
        '    Dim strSafe As String

        '    strSafe = Replace(strName, "/", "")
        '    strSafe = Replace(strSafe, " ", "")
        '    SafeSheetName = Left(strSafe, Globals.cMaxSheetNameLen)
        'End Function

        Public Shared Sub ScreenUpdatesOff()
            If True = Globals.cScreenUpdatesOff Then
                With Globals.ExcelApp.Application
                    If .Workbooks.Count > 0 Then
                        Globals.ExcelApp.priorScreenUpdatingState = .ScreenUpdating
                        .ScreenUpdating = False
                    Else
                        ' Assume the intent is to run with screen updates on.
                        Globals.ExcelApp.priorScreenUpdatingState = True
                        .ScreenUpdating = False
                    End If
                End With
            End If
        End Sub

        Public Shared Sub ScreenUpdatesOn()
            With Globals.ExcelApp.Application
                .ScreenUpdating = Globals.ExcelApp.priorScreenUpdatingState
            End With
        End Sub

        '**********************************************************************
        '   P r i v a t e    M e t h o d s
        '**********************************************************************

    End Class

End Namespace
