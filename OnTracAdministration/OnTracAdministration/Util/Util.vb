Option Explicit On

Imports System.Collections
Imports System.Collections.Generic

Public Class Util
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

            Debug.Print("ThisAddin.Application.StartupPath:" & Globals.ThisAddIn.Application.StartupPath.ToString)
            Debug.Print("ThisAddin.Application.ActiveWorkbook.Name:" & Globals.ThisAddIn.Application.ActiveWorkbook.Name.ToString)
            Debug.Print("ThisAddin.Application.ActiveWorkbook.Path:" & Globals.ThisAddIn.Application.ActiveWorkbook.Path.ToString)
            Debug.Print("ThisAddin.Application.ActiveWorkbook.FullName:" & Globals.ThisAddIn.Application.ActiveWorkbook.FullName.ToString)
            Debug.Print("ThisAddin.Application.ActiveWorkbook.FullNameURLEncoded:" & Globals.ThisAddIn.Application.ActiveWorkbook.FullNameURLEncoded.ToString)

            Debug.Print("ThisAddin.Application.DefaultFilePath:" & Globals.ThisAddIn.Application.DefaultFilePath.ToString)
            Debug.Print("ThisAddin.Application.Name:" & Globals.ThisAddIn.Application.Name.ToString)
            Debug.Print("ThisAddin.Application.NetworkTemplatesPath:" & Globals.ThisAddIn.Application.NetworkTemplatesPath.ToString)
            Debug.Print("ThisAddin.Application.Path:" & Globals.ThisAddIn.Application.Path.ToString)
        Catch ex As Exception
            MessageBox.Show("ApplicationInfo():" & ex.ToString)
        End Try
    End Sub

    Public Shared Function BaseName(ByVal strName As String) As String
        BaseName = Left(strName, InStr(1, strName, ".", vbTextCompare) - 1)
    End Function

    Public Shared Sub CalculationsOff()
        ' Don't bother trying to save current if no open workbooks.

        With Globals.ThisAddIn.Application
            If .Workbooks.Count > 0 Then
                Globals.ThisAddIn.m_vntPriorCalculationState = .Calculation
                .Calculation = Excel.XlCalculation.xlCalculationManual
            Else
                ' Assume the intent is to run with calculation and screen updates on.
                ' Hopefully we never get called with no workbooks open.
                Globals.ThisAddIn.m_vntPriorCalculationState = Excel.XlCalculation.xlCalculationAutomatic
            End If
        End With
    End Sub ' CalculationsOff

    Public Shared Sub CalculationsOn()
        With Globals.ThisAddIn.Application
            .Calculation = Globals.ThisAddIn.m_vntPriorCalculationState
        End With
    End Sub ' CalculationsOn

    Public Shared Sub CopySurveyValuesToAllTeamsSurveyWorksheets()
        Dim destinationWs As Excel.Worksheet
        Dim sourceSheet As Excel.Worksheet
        Dim targetRange As Excel.Range
        Dim teamName As String

        ' Get the currently selected team

        teamName = Globals.ThisAddIn.Application.Sheets(Globals.cSN_ScoreCards_IndividualTeam).Range(Globals.cSCIT_TeamNameCell).Value

        ' Get a reference to the Survey Mapping worksheet and

        sourceSheet = Globals.ThisAddIn.Application.Sheets(Globals.cSN_ScoreCards_SurveyMapping)

        ' Copy the IT Survey Values for the currently selected team

        sourceSheet.Range(Globals.cSM_ITSurveyCells).Copy()

        destinationWs = Globals.ThisAddIn.Application.Sheets(Globals.cSN_ITSurvey_AllTeams)

        targetRange = destinationWs.Range(Config.TeamNameToCells(teamName))

        targetRange.PasteSpecial(Paste:=Excel.XlPasteType.xlPasteValuesAndNumberFormats, _
            Operation:=Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, _
            SkipBlanks:=False, Transpose:=False)

        ' and the number of Survey responses

        targetRange.Offset(Globals.cSM_ResponseCountOffset, 0).Value = sourceSheet.Range(Globals.cSM_ITSurveyResponsesCell).Value

        ' Copy the Business Survey Values for the currently selected team

        sourceSheet.Range(Globals.cSM_BusinessSurveyCells).Copy()

        destinationWs = Globals.ThisAddIn.Application.Sheets(Globals.cSN_BusinessSurvey_AllTeams)

        targetRange = destinationWs.Range(Config.TeamNameToCells(teamName))

        targetRange.PasteSpecial(Paste:=Excel.XlPasteType.xlPasteValuesAndNumberFormats, _
          Operation:=Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, _
          SkipBlanks:=False, Transpose:=False)

        ' and the number of Survey responses

        targetRange.Offset(Globals.cSM_ResponseCountOffset, 0).Value = sourceSheet.Range(Globals.cSM_BusinessSurveyResponsesCell).Value

        ' Copy the Partner Survey Values

        sourceSheet.Range(Globals.cSM_PartnerSurveyCells).Copy()

        destinationWs = Globals.ThisAddIn.Application.Sheets(Globals.cSN_PartnerSurvey_AllTeams)

        targetRange = destinationWs.Range(Config.TeamNameToCells(teamName))

        targetRange.PasteSpecial(Paste:=Excel.XlPasteType.xlPasteValuesAndNumberFormats, _
            Operation:=Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, _
            SkipBlanks:=False, Transpose:=False)

        ' and the number of Survey responses

        targetRange.Offset(Globals.cSM_ResponseCountOffset, 0).Value = sourceSheet.Range(Globals.cSM_PartnerSurveyResponsesCell).Value

    End Sub

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



    Public Shared Sub CopyScorecardValuesToAllTeamsScorecardWorksheet()
        '
        ' CopyValues Macro
        '
        Dim teamName As String
        Dim targetRange As Excel.Range
        'Dim targetAddress As String
        Dim sourceWs As Excel.Worksheet
        Dim destinationWs As Excel.Worksheet

        Util.ScreenUpdatesOff()

        ' Get references to the Individual Team Scorecard worksheet (source)
        ' and to the All Teams Scorecard worksheet (destination)

        sourceWs = Globals.ThisAddIn.Application.Sheets(Globals.cSN_ScoreCards_IndividualTeam)
        destinationWs = Globals.ThisAddIn.Application.Sheets(Globals.cSN_ScoreCards_AllTeams)

        ' Get the team that is currently selected

        teamName = sourceWs.Range(Globals.cSCIT_TeamNameCell).Value

        ' and copy the current values for the selected team

        sourceWs.Range(Globals.cSCIT_TeamScoreCells).Copy()

        ' Determine where the values should be pasted on the destination worksheet

        targetRange = destinationWs.Range(Config.TeamNameToCells(teamName))

        ' and paste them there.

        targetRange.PasteSpecial(Paste:=Excel.XlPasteType.xlPasteValuesAndNumberFormats, _
            Operation:=Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, _
            SkipBlanks:=False, Transpose:=False)

        ' Finally, copy ITR processing data.

        targetRange.Offset(31, 0).Value = sourceWs.Range(Globals.cSCIT_OpenedITRsCell).Value
        targetRange.Offset(32, 0).Value = sourceWs.Range(Globals.cSCIT_ClosedITRsCell).Value
        targetRange.Offset(33, 0).Value = sourceWs.Range(Globals.cSCIT_ActiveITRsCell).Value

        Util.ScreenUpdatesOn()

    End Sub

    Public Shared Sub CreateName(ByVal wb As Excel.Workbook, ByVal name As String, ByVal targetRange As String)
        Try
            wb.Names.Item(name).Delete()
        Catch ex As Exception

        End Try

        wb.Names.Add(Name:=name, RefersToR1C1:=targetRange)

    End Sub

    Public Shared Sub DeleteSheet(ByVal ws As Excel.Worksheet, Optional ByVal prompt As Boolean = False)
        Dim priorState As Boolean

        priorState = Globals.ThisAddIn.Application.DisplayAlerts

        If prompt Then
            Globals.ThisAddIn.Application.DisplayAlerts = True
            ws.Delete()
        Else
            Globals.ThisAddIn.Application.DisplayAlerts = False
            ws.Delete()

        End If

        Globals.ThisAddIn.Application.DisplayAlerts = priorState
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

        currentCellRange = Globals.ThisAddIn.Application.ActiveCell
        currentRowRange = Globals.ThisAddIn.Application.ActiveSheet.Rows.Item(currentCellRange.Row)
        currentColumnRange = Globals.ThisAddIn.Application.ActiveSheet.Columns.Item(currentCellRange.Column)

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
                currentRowRange = Globals.ThisAddIn.Application.ActiveSheet.Rows.Item(searchFromCell.Row)
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
                currentColumnRange = Globals.ThisAddIn.Application.ActiveSheet.Columns.Item(searchFromCell.Column)

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
        Dim wb As Excel.Workbook = Globals.ThisAddIn.Application.ActiveWorkbook

        ' There might not be any open Workbooks.
        If wb Is Nothing Then
            wb = Globals.ThisAddIn.Application.Workbooks.Add()
        End If

        With wb
            Try
                For Each ws As Excel.Worksheet In .Worksheets
                    If ws.Name = sheetName Then
                        ' Sheet exists.  Ask user what to do.
                        Dim result As System.Windows.Forms.DialogResult = MessageBox.Show("Sheet: >" & sheetName & "< already exists.  Overwrite?", "WorkSheet Exists", MessageBoxButtons.YesNo)
                        If result = DialogResult.Yes Then
                            ws.Cells.Clear()
                            Return ws
                        Else
                            ' TODO: Decide how best to handle this.  For now just tweak the Sheetname
                            sheetName = sheetName & "1"
                        End If
                    End If
                Next
            Catch ex As Exception
                ' Likely there are no worksheets open.
            End Try

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

    ' Do this with regular expressions.

    Function SafeSheetName(ByVal strName As String) As String
        Dim strSafe As String

        strSafe = Replace(strName, "/", "")
        strSafe = Replace(strSafe, " ", "")
        SafeSheetName = Left(strSafe, Globals.cMaxSheetNameLen)
    End Function

    Public Shared Sub ScreenUpdatesOff()
        If True = Globals.cScreenUpdatesOff Then
            With Globals.ThisAddIn.Application
                If .Workbooks.Count > 0 Then
                    Globals.ThisAddIn.priorScreenUpdatingState = .ScreenUpdating
                    .ScreenUpdating = False
                Else
                    ' Assume the intent is to run with screen updates on.
                    Globals.ThisAddIn.priorScreenUpdatingState = True
                    .ScreenUpdating = False
                End If
            End With
        End If
    End Sub

    Public Shared Sub ScreenUpdatesOn()
        With Globals.ThisAddIn.Application
            .ScreenUpdating = Globals.ThisAddIn.priorScreenUpdatingState
        End With
    End Sub

    '**********************************************************************
    '   P r i v a t e    M e t h o d s
    '**********************************************************************

End Class
