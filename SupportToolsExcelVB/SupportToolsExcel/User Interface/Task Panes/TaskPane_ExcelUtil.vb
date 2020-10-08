Imports Microsoft.Office.Interop
Imports PacificLife.Life
Imports System.Collections.Generic
Imports System.Collections
Imports System.Diagnostics

Public Class TaskPane_ExcelUtil

#Region "Private Constants and Ennumerations"

    Private Const _ERROR_EMPTY_CELL As String = "Cell is empty.  Must select a popluated starting cell first."

    Private Const _INDENT_LEVEL As Short = 1
    Private Const _COL_WIDTH As Short = 3
    Private Const _NOTE_WIDTH As Short = 20
    Private Const _FILE_FONT_SIZE As Short = 6
    Private Const _FOLDER_FONT_SIZE As Short = 8

    Private Const _HEADING_ROW As Integer = 2
    Private Const _INITIAL_ROW As Integer = _HEADING_ROW + 1

    Private Const _FOLDER_INFO_COL As Short = 1 ' Folder level info starts here
    Private Const _FOLDER_INFO_LEN As Short = 10
    Private Const _FILE_INFO_COL As Short = 5   ' File Info starts here
    Private Const _FILE_INFO_LEN As Short = 4
    Private Const _NOTE_COL As Short = 10
    Private Const _INITIAL_COL As Short = 11    ' Map Info starts here

    Private Const _MAKE_BOLD As Boolean = True

    Public Enum _DateType As Integer
        LastCreate = 1
        LastWrite = 2
        LastAccess = 3
    End Enum

#End Region

    Private Sub btnGetLastRowColInfo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetLastRowColInfo.Click
        GetLastRowColInfo()
    End Sub

    Private Sub btnDeleteDuplicateRows_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteDuplicateRows.Click
        DeleteDuplicateRows()
    End Sub

    Private Sub btnCreateFolderMap_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateFolderMap.Click
        Excel_FolderMaps.CreateFolderMap()
    End Sub

    Private Sub btnGroupDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGroupDown.Click
        Excel_GroupDown.GroupColumnRangeDown()
    End Sub

    Private Sub btnSearchDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearchDown.Click
        Excel_SearchDown.FindEndOfRangeDown()
    End Sub

    Private Sub btnSearchUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearchUp.Click
        'SearchUp()
    End Sub

    Private Sub btnUnGroupSelection_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUnGroupSelection.Click
        Excel_UngroupSelection.UnGroupSelection()
    End Sub

    Private Sub DeleteDuplicateRows()
        Dim currentCell As Excel.Range = Globals.ThisAddIn.Application.ActiveCell
        Dim firstRow As Integer = currentCell.Row
        Dim lastRow As Integer = currentCell.SpecialCells(Excel.Constants.xlLastCell).Row
        Dim foundRows As Hashtable = New Hashtable()
        Dim activeSheet As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet
        lastRow = lastRow
        Dim rowsToDelete As List(Of String) = New List(Of String)
        Dim count As Integer = 0
        Dim rng As Excel.Range

        Debug.Print(String.Format("{0} - {1}", firstRow, lastRow))

        For i As Integer = lastRow To firstRow Step -1
            rng = activeSheet.Cells(i, currentCell.Column)

            Debug.Print(String.Format("{0} - {1} - {2}", i, rng.Value, lastRow))
            If foundRows.Contains(rng.Value) Then
                rowsToDelete.Add(String.Format("{0}:{1}", i, i))

                'currentCell.Rows(String.Format("{0}:{1}", i + currentCell.Row - 1, i + currentCell.Row - 1)).Delete()
            Else
                foundRows.Add(rng.Value, "X")
            End If
        Next i

        For Each row As String In rowsToDelete
            Debug.Print(row)
            activeSheet.Rows(row).Delete()
        Next

    End Sub

    Private Sub GetLastRowColInfo()
        Dim rng As Excel.Range = Globals.ThisAddIn.Application.ActiveCell

        txtLastRowSearch.Text = Common.ExcelHelper.FindLastRow(rng).ToString
        txtLastColSearch.Text = Common.ExcelHelper.FindLastColumn(rng).ToString
        txtLastRowSpecial.Text = rng.SpecialCells(Excel.Constants.xlLastCell).Row.ToString
        txtLastColSpecial.Text = rng.SpecialCells(Excel.Constants.xlLastCell).Column.ToString
    End Sub

    'Shared Function GetEndOfSectionDown( _
    '    ByVal intStartRow As Integer, _
    '    ByVal intStartCol As Integer, _
    '    ByVal intLastRow As Integer _
    ') As Integer
    '    Dim intMatchingRow As Integer

    '    With Globals.ThisAddIn.Application
    '        ' Search down for a matching cell
    '        intMatchingRow = .Cells(intStartRow, intStartCol).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).Row

    '        If intStartCol = _INITIAL_COL Then
    '            ' We have back'd all the way back to the first column.
    '            ' Return either then next matching cell down or the last
    '            ' populated row on the sheet.

    '            If intMatchingRow < intLastRow Then
    '                ' Section ends on the row prior to the match.
    '                GetEndOfSectionDown = intMatchingRow - 1
    '            Else
    '                ' Return end of populated section
    '                GetEndOfSectionDown = intLastRow
    '            End If
    '        Else
    '            If intMatchingRow <= intLastRow Then
    '                ' Back up one column and search down for a populated cell.
    '                ' Treat row prior to matching row as new end.
    '                GetEndOfSectionDown = GetEndOfSectionDown(intStartRow, intStartCol - 1, intMatchingRow - 1)
    '            Else
    '                ' Back up one column and search down for a populated cell.
    '                ' Treat end of worksheet as end.
    '                GetEndOfSectionDown = GetEndOfSectionDown(intStartRow, intStartCol - 1, intLastRow)
    '            End If
    '        End If
    '    End With
    'End Function

    'Shared Function GetStartOfSectionUp( _
    '    ByVal intStartRow As Integer, _
    '    ByVal intStartCol As Integer _
    ') As Integer

    '    With Globals.ThisAddIn.Application
    '        ' Search Up for matching cell
    '        GetStartOfSectionUp = .Cells(intStartRow, intStartCol).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
    '    End With
    'End Function

    'Private Sub GroupDown()
    '    Dim intStartRow As Integer
    '    Dim intStartCol As Integer
    '    Dim intLastRow As Integer
    '    Dim intLastCol As Integer
    '    Dim intEndRowOfSection As Integer

    '    With Globals.ThisAddIn.Application
    '        If IsNothing(.ActiveCell.Value) Then
    '            MsgBox(_ERROR_EMPTY_CELL)
    '        Else
    '            Try
    '                ' Get the last populated cell on the worksheet.
    '                intLastRow = .ActiveSheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell).Row()
    '                intLastCol = .ActiveSheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell).Column()

    '                ' Save where we currently are located.
    '                intStartRow = .ActiveCell.Row
    '                intStartCol = .ActiveCell.Column

    '                intEndRowOfSection = GetEndOfSectionDown(intStartRow, intStartCol, intLastRow)
    '                .Rows(intStartRow + 1 & ":" & intEndRowOfSection).Select()

    '                ' Group the hilighted rows so can collapse if desired.
    '                .Selection.Rows.Group()

    '                ' TODO:
    '                ' Select a cell at the bottom of the range so can easily collapse
    '                '.Cells(intEndRowOfSection, intStartCol).Select()
    '                .Selection.Rows.Hidden = True
    '            Catch ex As Exception
    '                PLLog.Error(ex, Common.PROJECT_NAME)
    '                Throw (ex)
    '            End Try
    '        End If
    '    End With
    'End Sub

    'Private Sub SearchDown()
    '    Dim intStartRow As Integer
    '    Dim intStartCol As Integer
    '    Dim intLastRow As Integer
    '    Dim intLastCol As Integer
    '    Dim intEndRowOfSection As Integer

    '    With Globals.ThisAddIn.Application
    '        If IsNothing(.ActiveCell.Value) Then
    '            MsgBox(_ERROR_EMPTY_CELL)
    '        Else
    '            Try
    '                ' Get the last populated cell on the worksheet.
    '                intLastRow = .ActiveSheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell).Row()
    '                intLastCol = .ActiveSheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell).Column()

    '                ' Save where we currently are located.
    '                intStartRow = .ActiveCell.Row
    '                intStartCol = .ActiveCell.Column

    '                intEndRowOfSection = GetEndOfSectionDown(intStartRow, intStartCol, intLastRow)
    '                .Cells(intEndRowOfSection, intStartCol).Select()
    '            Catch ex As Exception
    '                PLLog.Error(ex, Common.PROJECT_NAME)
    '                Throw (ex)
    '            End Try
    '        End If
    '    End With
    'End Sub

    ''--------------------------------------------------------------------------------
    ''
    '' SearchUp()
    ''
    '' Search up from current cell looking for start of current range.
    '' Start of current range is the next occupied cell in the current section
    '' with no intervening outdented sections.  Current range ends just before
    '' subsequent outdented sections or on last occupied row.
    ''
    ''--------------------------------------------------------------------------------

    'Private Sub SearchUp()
    '    Dim intStartRow As Integer
    '    Dim intStartCol As Integer
    '    Dim intStartRowOfSection As Integer

    '    With Globals.ThisAddIn.Application
    '        Try
    '            ' Save where we currently are located.
    '            intStartRow = .ActiveCell.Row
    '            intStartCol = .ActiveCell.Column

    '            intStartRowOfSection = GetStartOfSectionUp(intStartRow, intStartCol)
    '            .Cells(intStartRowOfSection, intStartCol).Select()
    '        Catch ex As Exception
    '            PLLog.Error(ex, Common.PROJECT_NAME)
    '            Throw (ex)
    '        End Try
    '    End With
    'End Sub

    'Private Sub UnGroupSelection()
    '    Globals.ThisAddIn.Application.Selection.Rows.UnGroup()
    'End Sub
End Class
