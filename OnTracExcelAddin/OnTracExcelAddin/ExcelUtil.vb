Option Strict Off
Option Explicit On

Imports System.Text

Imports Microsoft.Office.Core
'Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop
Imports PacificLife.Life

Class ExcelUtil
    '********************************************************************************
    '
    ' $Workfile: ExcelUtil.vb $
    ' $Revision: 1 $
    '
    ' Description:
    '   This module contains ....
    '
    ' Public Methods:
    '   Method(arg1, arg2) As Type
    '
    ' Public Types and Variables:
    '   Note: Put these in modGlobals.bas unless need here.
    '
    ' ToDo:
    '   List of ideas for improvement.
    '
    ' $History: ExcelUtil.vb $
'
'*****************  Version 1  *****************
'User: Crhodes      Date: 2/02/11    Time: 2:20p
'Created in $/Office/OnTracExcelAddin/OnTracExcelAddin
'
'*****************  Version 1  *****************
'User: Crhodes      Date: 7/20/07    Time: 4:00p
'Created in $/VSTO/OfficeAddin/OfficeAddin/OfficeAddin
    '
    '********************************************************************************


    '**********************************************************************
    '   E x t e r n a l    F u n c t i o n    D e c l a r a t i o n s
    '**********************************************************************
    ' Put these in modGlobals.bas


    '**********************************************************************
    '   P u b l i c    C o n s t a n t s
    '**********************************************************************
    ' Put these in modGlobals.bas


    '**********************************************************************
    '   P u b l i c    T y p e s    a n d    V a r i a b l e s
    '**********************************************************************
    ' Put these in modGlobals.bas


    '**********************************************************************
    '   P r i v a t e    C o n s t a n t s
    '**********************************************************************

    Private Const cMODULE_NAME As String = Globals.PROJECT_NAME & ".modExcel"

    '**********************************************************************
    '   P r i v a t e    T y p e s    a n d    V a r i a b l e s
    '**********************************************************************

    '**********************************************************************
    '   P u b l i c    M e t h o d s
    '**********************************************************************
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
                    .Orientation = Orientation
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
                Globals.PriorCalculationState = .Calculation
                .Calculation = Excel.XlCalculation.xlCalculationManual
            Else
                ' Assume the intent is to run with calculation and screen updates on.
                ' Hopefully we never get called with no workbooks open.
                Globals.PriorCalculationState = Excel.XlCalculation.xlCalculationAutomatic
            End If
        End With
    End Sub ' CalculationsOff

    Public Shared Sub CalculationsOn()
        With Globals.ThisAddIn.Application
            .Calculation = Globals.PriorCalculationState
        End With
    End Sub ' CalculationsOn

    '    '----------------------------------------------------------------------
    '    '
    '    ' CountDelimiters
    '    '
    '    ' Return how many times SearchChar appears in Files.
    '    '
    '    ' ToDo:
    '    '
    '    '----------------------------------------------------------------------

    '    Public Function CountDelimiters( _
    '        ByVal strFiles As String, _
    '        ByVal vntSearchChar As Object _
    '    ) As Integer
    '        On Error GoTo PROC_ERROR
    '        Const cRoutineName = "CountDelimiters"

    '        Dim i As Integer
    '        Dim intResult As Integer

    '        For i = 1 To Len(strFiles)
    '            If Mid(strFiles, i, 1) = vntSearchChar Then
    '                intResult = intResult + 1
    '            End If
    '        Next i

    '        CountDelimiters = intResult
    'PROC_EXIT:
    '        Exit Function

    'PROC_ERROR:
    '        Err.Raise(Err.Number, Err.Source, _
    '            cRoutineName & "():" & Err.Description, _
    '            Err.HelpFile, Err.HelpContext)
    '        GoTo PROC_EXIT
    '        Resume Next
    '    End Function    ' CountDelimiters

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

    Public Shared Sub DisplayCellInfo(ByVal cell As Excel.Range)
        With cell
            Debug.Print("Value:>" & .Value & "<")
            Debug.Print("ID:>" & .ID & "<")
            Debug.Print("" & .Hyperlinks.Count)
            DisplayHyperLinkInfo(.Hyperlinks.Item(1))
        End With
    End Sub

    Public Shared Sub DisplayExcelRange(ByVal rng As Excel.Range)
        Debug.Print(rng.Address)
    End Sub

    Public Shared Sub DisplayHyperLinkInfo(ByVal link As Excel.Hyperlink)
        With link
            Debug.Print("Address:>" & .Address.ToString & "<")
            Debug.Print("Name:>" & .Name.ToString & "<")
            'Debug.Print("ScreenTip:>" & .ScreenTip.ToString & "<")
            'Debug.Print("SubAddress:>" & .SubAddress.ToString & "<")
            Debug.Print("TextToDisplay:>" & .TextToDisplay.ToString & "<")
        End With
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

    Public Shared Function GetBuiltInPropertyValue( _
ByVal objDoc As Object, _
ByVal PropName As String _
) As String

        ' This procedure returns the value of the built-in document
        ' property specified in the strPropName argument for the Office
        ' document object specified in the objDoc argument.

        Dim prpDocProp As DocumentProperty
        Dim varValue As Object

        Const ERR_BADPROPERTY As Long = 5
        Const ERR_BADDOCOBJ As Long = 438
        Const ERR_BADCONTEXT As Long = -2147467259

        Try
            prpDocProp = objDoc.BuiltinDocumentProperties(PropName)

            With prpDocProp
                varValue = .Value
                If Len(varValue) <> 0 Then
                    GetBuiltInPropertyValue = varValue
                Else
                    GetBuiltInPropertyValue = "<Not Set>"
                End If
            End With
        Catch ex As Exception
            Select Case Err.Number
                Case ERR_BADDOCOBJ
                    GetBuiltInPropertyValue = "<No Object.BuiltInProperties>"
                Case ERR_BADPROPERTY
                    GetBuiltInPropertyValue = "<Property not in collection>"
                Case ERR_BADCONTEXT
                    GetBuiltInPropertyValue = "<Value not available in this context>"
                Case Else
                    GetBuiltInPropertyValue = "<BuiltInProperty_Get:Unknown error>"
            End Select
        End Try
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

    ''----------------------------------------------------------------------
    ''
    '' CreateChartName
    ''
    ''----------------------------------------------------------------------


    'Function CreateChartName(ByVal intDataColumn As Integer) As String
    '    Dim strName As String
    '    Dim strCategory As String
    '    Dim strCounter As String

    '    ' HACK
    '    ' Need to figure out better way of naming charts.  Only get
    '    ' 30 characters on sheet name.  Need to either shorten or
    '    ' group in separate workbook.

    '    ' Keep the full device name.

    '    strName = Cells(cDeviceRow, intDataColumn).value
    '    strCategory = Cells(cCategoryRow, intDataColumn).value
    '    strCounter = Cells(cCounterRow, intDataColumn).value

    '    ' Clean up the known categories, truncate the rest.

    '    Select Case strCategory
    '        Case "Network Interface"    ' Learn RE and do Network Interface*
    '            strCategory = "Network"

    '        Case "PhysicalDisk(_Total)"
    '            strCategory = "PhyDsk"

    '        Case "Processor(_Total)"
    '            strCategory = "Processor"

    '        Case Else
    '            strCategory = Left(strCategory, 9)
    '    End Select

    '    ' Clean up the known counters, truncate the rest.

    '    Select Case strCounter
    '        Case "Bytes Total/sec"
    '            strCounter = "BytSec"

    '        Case "Context Switches/sec"
    '            strCounter = "CSwtSec"

    '        Case "% Disk Time"
    '            strCounter = "%Time"

    '        Case "Pages/sec"
    '            strCounter = "PgSec"

    '        Case "% Processor Time"
    '            strCounter = "%Time"

    '        Case "% Privileged Time"
    '            strCounter = "%Priv"

    '        Case Else
    '            strCounter = Left(strCounter, 9)
    '    End Select

    '    ' Be sure there are no problematic characters.

    '    CreateChartName = SafeSheetName(strName & "_" & strCategory & "_" & strCounter)
    'End Function

    '    '----------------------------------------------------------------------
    '    '
    '    ' CreateHyperLinkName
    '    '
    '    ' Returns a name suitable for using as a hyperlink.
    '    '
    '    ' ToDo:
    '    '
    '    '----------------------------------------------------------------------

    '    Public Function CreateHyperLinkName( _
    '        ByVal strDataSheet As String, _
    '        ByVal strReportType As String _
    '    ) As String
    '        On Error GoTo PROC_ERROR
    '        Const cRoutineName = "CreateHyperLinkName"

    '        Dim strMachineType As String
    '        Dim strDate As String
    '        Dim intStart As Integer
    '        Dim intEnd As Integer

    '        intEnd = InStr(1, strDataSheet, ".", vbTextCompare)
    '        strMachineType = Left(strDataSheet, intEnd - 1)

    '        intStart = InStr(intEnd + 1, strDataSheet, ".", vbTextCompare) + 1
    '        intEnd = InStr(intStart + 1, strDataSheet, ".", vbTextCompare)

    '        ' This should handle "cmb1.sv.Aug1.txt" and "cmb1.sv.Aug"

    '        If intEnd = 0 Then
    '            strDate = Mid(strDataSheet, intStart, Len(strDataSheet))
    '        Else
    '            strDate = Mid(strDataSheet, intStart, intEnd - intStart)
    '        End If

    '        CreateHyperLinkName = _
    '            strMachineType & "." _
    '            & strReportType & "." _
    '            & strDate

    'PROC_EXIT:
    '        Exit Function

    'PROC_ERROR:
    '        Err.Raise(Err.Number, Err.Source, _
    '            cRoutineName & "():" & Err.Description, _
    '            Err.HelpFile, Err.HelpContext)
    '        GoTo PROC_EXIT
    '        Resume Next
    '    End Function

    ''----------------------------------------------------------------------
    ''
    '' CreateMonthlySummarySheet
    ''
    '' Returns a name suitable for using as a hyperlink.
    ''
    '' ToDo:
    ''
    ''----------------------------------------------------------------------

    'Sub CreateMonthlySummarySheet()
    '    Dim strMonth As String
    '    Dim strYear As String
    '    Dim strS As String
    '    Dim dblDate As Double

    '    Sheets.Add()
    '    Cells.Select()

    '    With Selection.Font
    '        .Name = "Arial"
    '        .Size = 8
    '    End With

    '    strMonth = InputBox("Enter Month", , MonthName(Month(Now)))
    '    strYear = InputBox("Enter Year", , Year(Now))

    '    ' Add the statistics control lines.

    '    Range("F2").value = "Rows"
    '    Range("E3").value = "From"
    '    Range("E4").value = "To"
    '    Range("F3").value = 11          ' Start Row for AVERAGE
    '    ' Hard code for 31 day months.  Likely a reasonable hack.
    '    Range("F4").value = 11 + 30     ' End Row for AVERAGE

    '    ' Add the statistics rows.

    '    Range("A6").value = "High"
    '    Range("A7").value = "Average"
    '    Range("A8").value = "Low"


    '    AddStatisticsFormulas(Range("B7"))
    '    AddStatisticsFormulas(Range("D7"))
    '    AddStatisticsFormulas(Range("F7"))
    '    AddStatisticsFormulas(Range("H7"))
    '    AddStatisticsFormulas(Range("J7"))
    '    AddStatisticsFormulas(Range("L7"))
    '    AddStatisticsFormulas(Range("N7"))
    '    AddStatisticsFormulas(Range("P7"))
    '    AddStatisticsFormulas(Range("R7"))
    '    AddStatisticsFormulas(Range("T7"))

    '    ' Add some lines to separate the sections

    '    Range("A6:U8").Select()

    '    With Selection.Borders(xlEdgeTop)
    '        .LineStyle = xlContinuous
    '        .Weight = xlMedium
    '        .ColorIndex = xlAutomatic
    '    End With

    '    With Selection.Borders(xlEdgeBottom)
    '        .LineStyle = xlContinuous
    '        .Weight = xlMedium
    '        .ColorIndex = xlAutomatic
    '    End With

    '    ' Add the same headings that we used on individual days, but
    '    ' lower to accomodate the statistics rows we just added..

    '    InitializeSummarySheet( _
    '        strMonth & " " & strYear, _
    '        "Monthly Summary for " & strMonth & " " & strYear, _
    '        Range("A9"))

    '    ' Make a few tweaks.

    '    Range("A1").Select()

    '    With Selection.Font
    '        .Bold = True
    '        .Size = 12
    '    End With

    '    ' Make this fill a range with dates so can see where data goes.
    '    Range("A12").value = "<Date>"
    'End Sub

    ''----------------------------------------------------------------------
    ''
    '' InitializeSummarySheet
    ''
    '' Take argument for starting location.  Read time range stuff off form
    '' or pass in.
    ''
    '' ToDo:
    ''
    ''----------------------------------------------------------------------

    'Sub InitializeSummarySheet( _
    'ByVal strSheetName As String, _
    'ByVal strTitle As String, _
    'ByVal rngStartingRow As Range _
    ')
    '    ActiveSheet.Name = strSheetName
    '    Range("A1").value = strTitle

    '    ' Statistics can be generated based on two time ranges.  No effort
    '    ' has been made to make these smart, e.g. if overlapped.

    '    Range("A3").value = "Time Range 1"
    '    Range("B3").value = "7:00"
    '    Range("C3").value = "12:00"
    '    Range("A4").value = "Time Range 2"
    '    Range("B4").value = "13:00"
    '    Range("C4").value = "17:00"

    '    ' The following values  will get overwritten with values from the data file.
    '    ' These are the row values that control the beginning and ending
    '    ' off the two time ranges shown above.

    '    Range("D3").value = cNbrAddedRows + 1
    '    Range("E3").value = cNbrAddedRows + 2
    '    Range("D4").value = cNbrAddedRows + 3
    '    Range("E4").value = cNbrAddedRows + 4

    '    AddSummaryColumn(rngStartingRow.Offset(0, 1), "CPU Idle")
    '    AddSummaryColumn(rngStartingRow.Offset(0, 3), "Disk Wait")
    '    AddSummaryColumn(rngStartingRow.Offset(0, 5), "Disk Ops")
    '    AddSummaryColumn(rngStartingRow.Offset(0, 7), "BES")
    '    AddSummaryColumn(rngStartingRow.Offset(0, 9), "CSM")
    '    AddSummaryColumn(rngStartingRow.Offset(0, 11), "DOC")
    '    AddSummaryColumn(rngStartingRow.Offset(0, 13), "INX")
    '    AddSummaryColumn(rngStartingRow.Offset(0, 15), "PRI")
    '    AddSummaryColumn(rngStartingRow.Offset(0, 17), "SEC")
    '    AddSummaryColumn(rngStartingRow.Offset(0, 19), "WFL")
    'End Sub

    '' TODO: Replease Add and link with this routine.

    'Sub LinkSheetIndirect( _
    '    ByVal strSheetName As String, _
    '    ByVal blnNewWorksheet As Boolean, _
    '    ByVal strIndexSheetName As String, _
    '    ByVal strLinkName As String, _
    '    ByVal xlOrientation As Excel.XlPageOrientation, _
    '    ByVal strLinkRowCell As String, _
    '    ByVal strLinkColCell As String, _
    '    ByVal strBeforeSheetName As String, _
    '    ByVal strAfterSheetName As String, _
    '    ByVal blnIncrementRow As Boolean, _
    '    ByVal blnIncrementCol As Boolean, _
    '    Optional ByVal sngColWidth As Single = 8.43 _
    ')
    '    Dim intLinkRow As Integer
    '    Dim intLinkCol As Integer

    '    If blnNewWorksheet Then
    '        NewWorksheet(strSheetName, strBeforeSheetName, strAfterSheetName)
    '    End If

    '    Worksheet_Format(ActiveSheet.Name, xlOrientation)
    '    Columns().ColumnWidth = sngColWidth

    '    Worksheets(strIndexSheetName).Activate()
    '    intLinkRow = Range(strLinkRowCell).value
    '    intLinkCol = Range(strLinkColCell).value
    '    Cells(intLinkRow, intLinkCol).value = strSheetName

    '    ' TODO: Fix the code here.
    '    Cells(intLinkRow, intLinkCol).Select()
    '    ActiveSheet.Hyperlinks.Add(Anchor:=Selection, Address:="", SubAddress:= _
    '        "'" & strSheetName & "'!A1", TextToDisplay:=strLinkName)

    '    If True = blnIncrementRow Then
    '        Range(strLinkRowCell).value = intLinkRow + 1
    '    End If

    '    If True = blnIncrementCol Then
    '        Range(strLinkColCell).value = intLinkCol + 1
    '    End If

    '    Worksheets(strSheetName).Activate()
    'End Sub ' LinkSheetIndirect

    'Sub NewWorksheet(strWorksheetName As String, _
    '    Optional strBeforeSheetName As String, _
    '    Optional strAfterSheetName As String _
    ')
    '        If strBeforeSheetName <> "" Then
    '            ActiveWorkbook.Sheets.Add(Sheets(strBeforeSheetName))
    '        ElseIf strAfterSheetName <> "" Then
    '            ActiveWorkbook.Sheets.Add, Sheets(strAfterSheetName)
    '        Else
    '            ActiveWorkbook.Sheets.Add()
    '        End If

    '        ActiveSheet.Name = strWorksheetName
    '    End Sub

    'Sub PopulateSummarySheet(ByVal strSheetName As String)
    '    Dim strSUMSheetName As String
    '    Dim strSVSheetName As String

    '    Exit Sub

    '    Sheets(strSheetName).Select()

    '    strSUMSheetName = Range("C1").value
    '    strSVSheetName = Range("C2").value

    '    ' CPU Idle
    '    Range("B8").formula = "=" & strSUMSheetName & "!B1"
    '    Range("C8").formula = "=" & strSUMSheetName & "!B2"

    '    ' Disk Wait
    '    Range("D8").formula = "=" & strSUMSheetName & "!E1"
    '    Range("E8").formula = "=" & strSUMSheetName & "!E2"

    '    ' Disk Wait
    '    Range("F8").formula = "=" & strSUMSheetName & "!F1"
    '    Range("G8").formula = "=" & strSUMSheetName & "!F2"

    '    ' BES
    '    Range("H8").formula = "=" & strSVSheetName & "!F1"
    '    Range("I8").formula = "=" & strSVSheetName & "!F2"

    '    ' CSM
    '    Range("J8").formula = "=" & strSVSheetName & "!H1"
    '    Range("K8").formula = "=" & strSVSheetName & "!H2"

    '    ' DOC
    '    Range("L8").formula = "=" & strSVSheetName & "!J1"
    '    Range("M8").formula = "=" & strSVSheetName & "!J2"

    '    ' INX
    '    Range("N8").formula = "=" & strSVSheetName & "!L1"
    '    Range("O8").formula = "=" & strSVSheetName & "!L2"

    '    ' PRI
    '    Range("P8").formula = "=" & strSVSheetName & "!P1"
    '    Range("Q8").formula = "=" & strSVSheetName & "!P2"

    '    ' SEC
    '    Range("R8").formula = "=" & strSVSheetName & "!R1"
    '    Range("S8").formula = "=" & strSVSheetName & "!R2"

    '    ' WFL
    '    Range("T8").formula = "=" & strSVSheetName & "!Z1"
    '    Range("U8").formula = "=" & strSVSheetName & "!Z2"

    'End Sub

    'Function CreateWorkbookName(ByVal dtDate As Date) As String
    '    Dim strS As String

    '    ' TODO: Pass in Application Name
    '    strS = "Viewpoint_" & Format(dtDate, "yyyy_mm_dd") & "_Data"
    '    CreateWorkbookName = strS
    'End Function

    ''----------------------------------------------------------------------
    ''
    '' EmptyWorkbook
    ''
    '' Returns the name of a workbook containing one sheet.
    '' Creates new workbook if blnCreateNew
    '' All existing sheets except for strWorksheet are removed.
    ''
    '' ToDo:
    ''
    ''----------------------------------------------------------------------
    'Function EmptyWorkbook( _
    'ByVal strWorksheetName As String, _
    'ByVal blnCreateNew _
    ') As String
    '    Dim shtWS As Worksheet
    '    Dim strKeepName As String

    '    If True = blnCreateNew Then
    '        Workbooks.Add()
    '        ' Keep the name so we don't have to worry about what name is given.
    '        strKeepName = ActiveSheet.Name
    '    Else
    '        strKeepName = strWorksheetName
    '    End If

    '    Application.DisplayAlerts = False   ' Yes, delete the damn things!

    '    ' Remove all other worksheets

    '    For Each shtWS In ActiveWorkbook.Sheets
    '        If strKeepName <> shtWS.Name Then
    '            shtWS.Delete()
    '        End If
    '    Next shtWS

    '    Application.DisplayAlerts = True

    '    ActiveSheet.Name = strWorksheetName
    'End Function

    ''----------------------------------------------------------------------
    ''
    '' ExtractBaseCategory
    ''
    '' ToDo:
    ''   Consider making this smarter.
    ''   For now just hack off everything past the first "("
    ''
    ''----------------------------------------------------------------------

    'Function ExtractBaseCategory(ByVal strRawCategory As String) As String
    '    Dim intOffset As Integer

    '    intOffset = InStr(1, strRawCategory, "(", vbTextCompare)

    '    If intOffset > 0 Then
    '        ExtractBaseCategory = Left(strRawCategory, intOffset - 1)
    '    Else
    '        ExtractBaseCategory = strRawCategory
    '    End If
    'End Function

    ''----------------------------------------------------------------------
    ''
    '' ExtractReportType
    ''
    '' FileNet data files are presumed to look like
    ''   <srv type>.<report type>.<MonDD>.<ext>
    '' This routine returns the <report type> which drives downstream
    '' behavior.
    ''
    ''----------------------------------------------------------------------

    'Function ExtractReportType( _
    'ByVal strFilename As String _
    ') As String
    '    Dim intStart As Integer
    '    Dim intEnd As Integer

    '    If 2 > CountDelimiters(strFilename, ".") Then
    '        MsgBox("Invalid File.  Missing "".<report type>.""" & strFilename)
    '        ExtractReportType = ""
    '        Exit Function
    '    Else
    '        intStart = InStr(1, strFilename, ".", vbTextCompare)
    '        intEnd = InStr(intStart + 1, strFilename, ".", vbTextCompare)
    '        ExtractReportType = Mid(strFilename, intStart + 1, intEnd - intStart - 1)
    '    End If
    'End Function    ' ExtractReportType

    '    '----------------------------------------------------------------------
    '    '
    '    ' FindTimeRange
    '    '
    '    ' Return Data Rows covered by time range.
    '    '
    '    ' ToDo:
    '    '   Support more paper sizes.
    '    '----------------------------------------------------------------------

    '    Function FindTimeRange( _
    '        ByVal strDataSheet As String, _
    '        ByRef intDataStartRow As Integer, _
    '        ByRef intDataEndRow As Integer, _
    '        ByVal dtStartDateTime As Date, _
    '        ByVal dtEndDateTime As Date _
    '    ) As Boolean
    '        On Error GoTo PROC_ERROR
    '        Const cRoutineName = "FindTimeRange"

    '        Dim i As Integer
    '        Dim dtCloseEnough As Date

    '        Sheets(strDataSheet).Select()
    '        dtCloseEnough = Cells(cSamplePeriodRow, cTimeColumn).value

    '        ' The logic depends on how the time is ordered in the file.  FileNet puts most recent first.
    '        ' Perfmon puts most recent last.

    '        For i = intDataStartRow To intDataEndRow
    '            '        Debug.Print dtStartTime, CDate(Cells(i, cTimeColumn).Value), CDbl(Abs(Cells(i, cTimeColumn).Value - dtStartTime))
    '            '       Debug.Print Cells(i, cTimeColumn + 1).Value >= dtStartTime
    '            '        If Abs(Cells(i, cTimeColumn).Value - dblStartTime) < cCloseEnough Then
    '            If Abs(Cells(i, cDateTimeColumn).value - dtStartDateTime) < dtCloseEnough Then
    '                intDataStartRow = i
    '                Exit For
    '            End If
    '        Next i

    '        For i = intDataEndRow To intDataStartRow Step -1
    '            '        Debug.Print Cells(i, cTimeColumn).Value, Abs(Cells(i, cTimeColumn).Value - dblEndTime)
    '            '        If Cells(i, cTimeColumn).Value <= dblEndTime Then
    '            If Abs(Cells(i, cDateTimeColumn).value - dtEndDateTime) < dtCloseEnough Then
    '                intDataEndRow = i
    '                Exit For
    '            End If
    '        Next i

    'PROC_EXIT:
    '        Exit Function

    'PROC_ERROR:
    '        Err.Raise(Err.Number, Err.Source, _
    '            cRoutineName & "():" & Err.Description, _
    '            Err.HelpFile, Err.HelpContext)
    '        GoTo PROC_EXIT
    '        Resume Next
    '    End Function

    'Sub TestGetSheetSize()
    '    Dim intHeight As Integer
    '    Dim intWidth As Integer

    '    GetSheetSize(intHeight, intWidth)
    '    MsgBox("Height = " & intHeight & "   Width = " & intWidth)
    'End Sub

    'Function FoundDataWorksheet() As Boolean
    '    Dim shtWS As Worksheet

    '    FoundDataWorksheet = False

    '    For Each shtWS In ActiveWorkbook.Worksheets
    '        shtWS.Activate()

    '        If "Data File Name" = Range("A1").value Then
    '            FoundDataWorksheet = True
    '            Exit For
    '        End If
    '    Next shtWS
    'End Function

    '    '----------------------------------------------------------------------
    '    '
    '    ' GetSheetSize
    '    '
    '    ' Returns printout sheet size in Points.
    '    '
    '    ' ToDo:
    '    '   Support more paper sizes.
    '    '----------------------------------------------------------------------

    '    Sub GetSheetSize(ByRef intHeight As Integer, ByRef intWidth As Integer)
    '        On Error GoTo PROC_ERROR
    '        Const cRoutineName = "GetSheetSize"

    '        Dim shtS As Excel.Worksheet
    '        Dim dblHeight As Double
    '        Dim dblWidth As Double
    '        Dim dblX As Double
    '        Dim intHeightMargin As Integer  ' In Points
    '        Dim intWidthMargin As Integer   ' In Points

    '        shtS = ActiveSheet

    '        ' Assume xlPortrait orientation.  Will swap below if needed.

    '        Select Case shtS.PageSetup.PaperSize
    '            Case xlPaperLetter
    '                dblHeight = 11
    '                dblWidth = 8.5

    '            Case xlPaperLegal
    '                dblHeight = 14
    '                dblWidth = 8.5

    '            Case Else
    '                MsgBox("Unsupported Paper Size")
    '                dblHeight = 11
    '                dblWidth = 8.5

    '        End Select

    '        If shtS.PageSetup.Orientation <> xlPortrait Then
    '            ' Swap dimensions if xlLandscape
    '            dblX = dblHeight
    '            dblHeight = dblWidth
    '            dblWidth = dblX
    '        End If

    '        ' Subtract Margins from total paper size.

    '        With shtS.PageSetup
    '            intHeightMargin = .TopMargin + .BottomMargin
    '            intWidthMargin = .LeftMargin + .RightMargin
    '        End With

    '        ' Return size in points.

    '        intHeight = Application.InchesToPoints(dblHeight) - intHeightMargin
    '        intWidth = Application.InchesToPoints(dblWidth) - intWidthMargin
    '        '
    '        '    If frmCharts.chkEnforcePrinterPageSize.Value = True Then
    '        '        intHeight = intHeight * 3
    '        '    End If

    'PROC_EXIT:
    '        Exit Sub

    'PROC_ERROR:
    '        Err.Raise(Err.Number, Err.Source, _
    '            cRoutineName & "():" & Err.Description, _
    '            Err.HelpFile, Err.HelpContext)
    '        GoTo PROC_EXIT
    '        Resume Next
    '    End Sub

    'Public Function GetSummaryWorkbookName(ByVal intYear, ByVal strApplication) As String
    '    GetSummaryWorkbookName = intYear & strApplication & " Data Summary" & ".xls"
    'End Function

    'Public Function GetSummaryWorksheetName(ByVal strDevice, ByVal strCategory, ByVal strCounter) As String
    '    ' Return the appropriate worksheet name from input values.
    '    ' For now we only use the Device.

    '    GetSummaryWorksheetName = strDevice
    'End Function

    '''
    ''' Protect or unprotect the sheet.  Return the current setting before
    ''' and changes made.
    '''
    Public Shared Function ProtectSheet( _
        ByRef sht As Microsoft.Office.Interop.Excel.Worksheet, _
        ByVal protectMode As Boolean _
    ) As Boolean
        Dim currentMode As Boolean = sht.ProtectContents

        If protectMode = True Then
            sht.Protect()
        Else
            sht.Unprotect()
        End If

        Return currentMode
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

        With Globals.ThisAddIn.Application.ActiveWorkbook

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
                    Globals.PriorScreenUpdatingState = .ScreenUpdating
                    .ScreenUpdating = False
                Else
                    ' Assume the intent is to run with screen updates on.
                    Globals.PriorScreenUpdatingState = True
                    .ScreenUpdating = False
                End If
            End With
        End If
    End Sub

    Public Shared Sub ScreenUpdatesOn()
        With Globals.ThisAddIn.Application
            .ScreenUpdating = Globals.PriorScreenUpdatingState
        End With
    End Sub

    Public Sub SetCellValue( _
    ByVal rngR As Excel.Range, _
    ByVal vntValue As Object, _
    Optional ByVal lngHorizontalAlignment As Long = XlHAlign.xlHAlignLeft, _
    Optional ByVal strComment As String = "")

        With rngR
            .Value = vntValue
            .HorizontalAlignment = lngHorizontalAlignment
            If "" <> strComment Then
                .AddComment()
                .Comment.Visible = False
                .Comment.Text(Text:=strComment)
            End If
        End With
    End Sub


    'Public Sub TestScreenOff()
    '    Application.ScreenUpdating = False

    '    Application.Workbooks.Add()

    '    Application.ScreenUpdating = True
    'End Sub

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

                ExcelUtil.CalculationsOff()

                For Each workSheet In workBook.Sheets
                    With workSheet.PageSetup
                        sb.Length = 0
                        ' Five point font, path, and filename
                        sb.Append("&5&Z&F")
                        sb.Append(vbLf & "Created: ")
                        sb.Append(ExcelUtil.GetBuiltInPropertyValue(workBook, "Creation Date"))
                        sb.Append(vbLf & "Last Saved: ")
                        sb.Append(ExcelUtil.GetBuiltInPropertyValue(workBook, "Last Save Time"))
                        sb.Append(vbLf & "Last Printed: ")
                        sb.Append(ExcelUtil.GetBuiltInPropertyValue(workBook, "Last Print Date"))
                        'Debug.Print("LeftFooter:>" & sb.ToString & "<")
                        .LeftFooter = sb.ToString()

                        .CenterFooter = ""

                        sb.Length = 0
                        ' Five point font, page of pages
                        sb.Append("&5&P - &N")
                        sb.Append(vbLf & "Title :")
                        sb.Append(ExcelUtil.GetBuiltInPropertyValue(workBook, "Title"))
                        sb.Append(vbLf & "Subject: ")
                        sb.Append(ExcelUtil.GetBuiltInPropertyValue(workBook, "Subject"))
                        'Debug.Print("RightFooter:>" & sb.ToString & "<")
                        .RightFooter = sb.ToString()
                    End With

                    ' TODO: Indicate we have added a custom footer.  This will be looked for
                    ' in the before close event.

                    If Not HasCustomFooter() Then
                        CustomFooterExists(True)
                    End If

                Next

                ExcelUtil.CalculationsOn()
            End With
        Catch ex As Exception
            MsgBox("AddFooter:" & ex.ToString)
        End Try
    End Sub ' Footer_Add

    Friend Shared Function HasCustomFooter() As Boolean
        Dim prp As Office.DocumentProperty
        Dim prps As Office.DocumentProperties

        Try
            Try
                prps = Globals.ThisAddIn.Application.ActiveWorkbook.CustomDocumentProperties
                prp = prps.Item("HasCustomFooter")
                ' If the property exists we don't really care about the value
                Return True

            Catch ex As Exception
                ' Exception is thrown if property does not exist
                Return False
            End Try
        Finally

        End Try
    End Function

    Friend Shared Sub CustomFooterExists(ByVal hasCustomFooter As Boolean)
        Dim prp As Office.DocumentProperty
        Dim prps As Office.DocumentProperties

        Try
            Try
                prps = Globals.ThisAddIn.Application.ActiveWorkbook.CustomDocumentProperties
                ' Add a new property.
                prp = prps.Add("HasCustomFooter", False, _
                 Office.MsoDocProperties.msoPropertyTypeBoolean, True)
            Catch ex As Exception
                'PLLog.Error(ex, "OnTracExcelAddin")
                MessageBox.Show("CustomFooterExists() Unable to add HasCustomFooter property" & ex.Message)
            End Try
        Finally

        End Try
    End Sub

    Private Shared Sub DumpPropertyCollection( _
     ByVal prps As Office.DocumentProperties, _
     ByVal rng As Excel.Range, ByRef i As Integer)
        Dim prp As Office.DocumentProperty

        For Each prp In prps
            rng.Offset(i, 0).Value = prp.Name
            Try
                If Not prp.Value Is Nothing Then
                    rng.Offset(i, 1).Value = _
                     prp.Value.ToString
                End If
            Catch
                ' Do nothing at all.
            End Try
            i += 1
        Next
    End Sub

    Public Shared Sub ZapPageBreaks()
        Dim i As Integer
        Dim sht As Excel.Worksheet

        Dim vPB As Excel.VPageBreak
        Dim hPB As Excel.HPageBreak

        With Globals.ThisAddIn.Application

            For Each sht In .ActiveWorkbook.Sheets
                .ActiveSheet.PageSetup.PrintArea = ""

                Debug.Print(sht.Name)
                '        Debug.Print sht.HPageBreaks.Count
                '        Debug.Print sht.VPageBreaks.Count
                ' For some reason the page break handling is not clean.
                ' There are different types of page breaks, that is clear.
                ' Unfortunately the For Each hPB errors out if only Automatic
                ' Page breaks.  Wrap in try catch for AddIn
                On Error Resume Next
                With sht
                    If .VPageBreaks.Count > 0 Then

                        For Each vPB In .VPageBreaks
                            If vPB.Type = Excel.XlPageBreak.xlPageBreakManual Then
                                vPB.Delete()
                            End If
                        Next vPB
                    End If

                    If .HPageBreaks.Count > 0 Then
                        For Each hPB In .HPageBreaks
                            If hPB.Type = Excel.XlPageBreak.xlPageBreakManual Then
                                hPB.Delete()
                            End If
                        Next hPB
                    End If

                    '            For i = .HPageBreaks.Count To 1 Step -1
                    ''                Debug.Print .HPageBreaks.Item(i).Type
                    ''                Debug.Print .HPageBreaks.Item(i).Location
                    ''                Debug.Print .HPageBreaks.Item(i).Extent
                    '
                    '                If .HPageBreaks.Item(i).Type = xlPageBreakManual Then
                    '                    .HPageBreaks.Item(i).Delete
                    '                End If
                    '            Next i
                    '
                    '            For i = .VPageBreaks.Count To 1 Step -1
                    ''                Debug.Print .VPageBreaks.Item(i).Type
                    ''                Debug.Print .VPageBreaks.Item(i).Location
                    ''                Debug.Print .VPageBreaks.Item(i).Extent
                    '
                    '                If .VPageBreaks.Item(i).Type = xlPageBreakManual Then
                    '                    .VPageBreaks.Item(i).Delete
                    '                End If
                    '            Next i

                End With
            Next sht
        End With
    End Sub


    '**********************************************************************
    '   P r i v a t e    M e t h o d s
    '**********************************************************************


    '********************************************************************************
    '   End $Workfile: ExcelUtil.vb $
    '       $Revision: 1 $
    '********************************************************************************
End Class