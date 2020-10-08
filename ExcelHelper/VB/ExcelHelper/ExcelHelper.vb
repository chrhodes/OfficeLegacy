Imports System.Windows.Forms
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop

Public Class ExcelHelper
    'Private Const cMODULE_NAME As String = Globals.PROJECT_NAME & ".modExcel"

    Public Const cPointsToInch As Integer = 72    ' Hard coded for Excel? 1/72 Inches.
    Public Const cStartHour As Integer = 8        ' Times before this
    Public Const cEndHour As Integer = 20         ' and after this are hilighted.

    'Private m_vntPriorCalculationState As Object
    'Private m_vntPriorScreenUpdatingState As Object

    Public Const cHeaderFontSize As Integer = 12
    Public Const cHeaderFontSizeMedium As Integer = 10
    Public Const cHeaderFontSizeSmall As Integer = 8

    Public PriorCalculationState As Excel.XlCalculation
    Public PriorScreenUpdatingState As Boolean

    Private _application As Excel.Application

    Public Property Application() As Excel.Application
        Get
            Return _application
        End Get
        Set(ByVal Value As Excel.Application)
            _application = Value
        End Set
    End Property

    Private _enableScreenUpdatesToggle As Boolean = True

    Public Property EnableScreenUpdatesToggle() As Boolean
        Get
            Return _enableScreenUpdatesToggle
        End Get
        Set(ByVal Value As Boolean)
            _enableScreenUpdatesToggle = Value
        End Set
    End Property


    '**********************************************************************
    '   P u b l i c    M e t h o d s
    '**********************************************************************

    Public Sub AddColumnToSheet( _
        ByRef ws As Excel.Worksheet, _
        ByVal columnNumber As Integer, _
        ByVal columnWidth As Integer, _
        ByVal columnWrapText As Boolean, _
        ByVal headerRow As Integer, _
        Optional ByVal headerTitle As String = "", _
        Optional ByVal headerFontSize As Integer = cHeaderFontSize, _
        Optional ByVal headerBold As Boolean = True, _
        Optional ByVal headerUnderline As Boolean = True, _
        Optional ByVal headerWrapText As Boolean = True, _
        Optional ByVal headerHorizontalAlignment As Excel.XlHAlign = Excel.XlHAlign.xlHAlignGeneral, _
        Optional ByVal orientation As Integer = 0, _
        Optional ByVal columnNumberFormat As String = "General" _
    )

        With ws
            .Columns(columnNumber).ColumnWidth = columnWidth
            .Columns(columnNumber).WrapText = columnWrapText
            .Columns(columnNumber).NumberFormat = columnNumberFormat

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

    Public Sub AddNewColumnToSheet( _
        ByRef ws As Excel.Worksheet, _
        ByVal columnNumber As Integer, _
        ByVal columnWidth As Integer, _
        ByVal columnWrapText As Boolean, _
        ByVal headerRow As Integer, _
        ByVal shiftDirection As Excel.XlDirection, _
        ByVal insertFormatOrigin As Excel.XlInsertFormatOrigin, _
        Optional ByVal headerTitle As String = "", _
        Optional ByVal headerFontSize As Integer = cHeaderFontSize, _
        Optional ByVal headerBold As Boolean = True, _
        Optional ByVal headerUnderline As Boolean = True, _
        Optional ByVal headerWrapText As Boolean = True, _
        Optional ByVal headerHorizontalAlignment As Excel.XlHAlign = Excel.XlHAlign.xlHAlignGeneral, _
        Optional ByVal orientation As Integer = 0, _
        Optional ByVal columnNumberFormat As String = "General" _
)
        With ws
            '.Columns(columnNumber).Select()
            '.Selection.Insert(Shift:=shiftDirection, CopyOrigin:=insertFormatOrigin)
            .Columns(columnNumber).Insert(Shift:=shiftDirection, CopyOrigin:=insertFormatOrigin)

            .Columns(columnNumber).ColumnWidth = columnWidth
            .Columns(columnNumber).WrapText = columnWrapText
            .Columns(columnNumber).NumberFormat = columnNumberFormat

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

    Public Sub AddCommentToCell( _
        ByRef ws As Excel.Worksheet, _
        ByVal column As Integer, _
        ByVal row As Integer, _
        ByVal text As String, _
        Optional ByVal headerFontSize As Integer = Common.cHeaderFontSize, _
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

    Public Sub AddContentToCell( _
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

    'Public Sub ApplicationInfo()
    '    Try
    '        Debug.Print("Application.CommonAppDataPath:" & Application.CommonAppDataPath.ToString)
    '        Debug.Print("Application.CommonAppDataRegistry:" & Application.CommonAppDataRegistry.ToString)
    '        Debug.Print("Application.CompanyName:" & Application.CompanyName.ToString)
    '        Debug.Print("Application.CurrentCulture:" & Application.CurrentCulture.ToString)
    '        Debug.Print("Application.CurrentInputLanguage:" & Application.CurrentInputLanguage.ToString)
    '        Debug.Print("Application.ExecutablePath:" & Application.ExecutablePath.ToString)
    '        Debug.Print("Application.LocalUserAppDataPath:" & Application.LocalUserAppDataPath.ToString)
    '        Debug.Print("Application.ProductName:" & Application.ProductName.ToString)
    '        Debug.Print("Application.ProductVersion:" & Application.ProductVersion.ToString)
    '        Debug.Print("Application.SafeTopLevelCaptionFormat:" & Application.SafeTopLevelCaptionFormat.ToString)
    '        Debug.Print("Application.StartupPath:" & Application.StartupPath.ToString)
    '        Debug.Print("Application.UserAppDataPath:" & Application.UserAppDataPath.ToString)
    '        Debug.Print("Application.UserAppDataRegistry:" & Application.UserAppDataRegistry.ToString)

    '        Debug.Print("ThisAddin.Application.StartupPath:" & Globals.ThisAddIn.Application.StartupPath.ToString)
    '        Debug.Print("ThisAddin.Application.ActiveWorkbook.Name:" & Globals.ThisAddIn.Application.ActiveWorkbook.Name.ToString)
    '        Debug.Print("ThisAddin.Application.ActiveWorkbook.Path:" & Globals.ThisAddIn.Application.ActiveWorkbook.Path.ToString)
    '        Debug.Print("ThisAddin.Application.ActiveWorkbook.FullName:" & Globals.ThisAddIn.Application.ActiveWorkbook.FullName.ToString)
    '        Debug.Print("ThisAddin.Application.ActiveWorkbook.FullNameURLEncoded:" & Globals.ThisAddIn.Application.ActiveWorkbook.FullNameURLEncoded.ToString)

    '        Debug.Print("ThisAddin.Application.DefaultFilePath:" & Globals.ThisAddIn.Application.DefaultFilePath.ToString)
    '        Debug.Print("ThisAddin.Application.Name:" & Globals.ThisAddIn.Application.Name.ToString)
    '        Debug.Print("ThisAddin.Application.NetworkTemplatesPath:" & Globals.ThisAddIn.Application.NetworkTemplatesPath.ToString)
    '        Debug.Print("ThisAddin.Application.Path:" & Globals.ThisAddIn.Application.Path.ToString)
    '    Catch ex As Exception
    '        MessageBox.Show("ApplicationInfo():" & ex.ToString)
    '    End Try
    'End Sub

    Public Sub CalculationsOff()
        ' Don't bother trying to save current if no open workbooks.

        With Application
            If .Workbooks.Count > 0 Then
                Common.PriorCalculationState = .Calculation
                .Calculation = Excel.XlCalculation.xlCalculationManual
            Else
                ' Assume the intent is to run with calculation and screen updates on.
                ' Hopefully we never get called with no workbooks open.
                Common.PriorCalculationState = Excel.XlCalculation.xlCalculationAutomatic
            End If
        End With
    End Sub ' CalculationsOff

    Public Sub CalculationsOn()
        With Application
            .Calculation = Common.PriorCalculationState
        End With
    End Sub ' CalculationsOn

    Public Sub DeleteSheet(ByVal ws As Excel.Worksheet, Optional ByVal prompt As Boolean = False)
        Dim priorState As Boolean

        priorState = Application.DisplayAlerts

        If prompt Then
            Application.DisplayAlerts = True
            ws.Delete()
        Else
            Application.DisplayAlerts = False
            ws.Delete()

        End If

        Application.DisplayAlerts = priorState
    End Sub

    Public Sub DisplayExcelRange(ByVal rng As Excel.Range)
        Debug.Print(rng.Address)
    End Sub

    Public Sub FindLast()
        Dim currentCellRange As Excel.Range
        Dim currentRowRange As Excel.Range
        Dim currentColumnRange As Excel.Range
        Dim lastRow As Long
        Dim lastColumn As Long

        currentCellRange = Application.ActiveCell
        currentRowRange = Application.ActiveSheet.Rows.Item(currentCellRange.Row)
        currentColumnRange = Application.ActiveSheet.Columns.Item(currentCellRange.Column)

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

    Public Function FindLastColumn(ByVal searchFromCell As Excel.Range) As Long
        Dim currentRowRange As Excel.Range
        'Dim currentColumnRange As Excel.Range
        'Dim lastRow As Long
        Dim lastColumn As Long

        If searchFromCell Is Nothing Then
            MessageBox.Show("FindLastColumn(): searchFromCell is Nothing")
        Else
            Try
                currentRowRange = Application.ActiveSheet.Rows.Item(searchFromCell.Row)
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

    Public Function FindLastRow(ByVal searchFromCell As Excel.Range) As Long
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
                currentColumnRange = Application.ActiveSheet.Columns.Item(searchFromCell.Column)

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

    Public Function GetBuiltInPropertyValue( _
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

    Public Function GetFile( _
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

    Public Function GetOpenFileName( _
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

    Public Function GetSaveFileName( _
    Optional ByVal initialFolder As String = "", _
    Optional ByVal proposedFileName As String = "", _
    Optional ByVal dialogTitle As String = "Open", _
    Optional ByVal fileFilter As String = "All Files (*.*)|*.*") As String
        Dim sfd As New SaveFileDialog
        Dim result As System.Windows.Forms.DialogResult

        sfd.InitialDirectory = initialFolder
        sfd.Title = dialogTitle
        sfd.Filter = fileFilter
        sfd.FileName = proposedFileName

        result = sfd.ShowDialog()

        Debug.WriteLine(sfd.FileName)

        Return sfd.FileName

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
    '    If strBeforeSheetName <> "" Then
    '        Application.ActiveWorkbook.Sheets.Add(Sheets(strBeforeSheetName))
    '    ElseIf strAfterSheetName <> "" Then
    '        Application.ActiveWorkbook.Sheets.Add, Sheets(strAfterSheetName)
    '    Else
    '        ActiveWorkbook.Sheets.Add()
    '    End If

    '    ActiveSheet.Name = strWorksheetName
    'End Sub


    '----------------------------------------------------------------------
    '
    ' EmptyWorkbook
    '
    ' Returns the name of a workbook containing one sheet.
    ' Creates new workbook if blnCreateNew
    ' All existing sheets except for strWorksheet are removed.
    '
    ' ToDo:
    '
    '----------------------------------------------------------------------
    Function EmptyWorkbook( _
        ByVal strWorksheetName As String, _
        ByVal blnCreateNew As Boolean _
    ) As String
        Dim shtWS As Excel.Worksheet
        Dim strKeepName As String

        If True = blnCreateNew Then
            Application.Workbooks.Add()
            ' Keep the name so we don't have to worry about what name is given.
            strKeepName = Application.ActiveSheet.Name
        Else
            strKeepName = strWorksheetName
        End If

        Application.DisplayAlerts = False   ' Yes, delete the damn things!

        ' Remove all other worksheets

        For Each shtWS In Application.ActiveWorkbook.Sheets
            If strKeepName <> shtWS.Name Then
                shtWS.Delete()
            End If
        Next shtWS

        Application.DisplayAlerts = True

        Application.ActiveSheet.Name = strWorksheetName

        Return strWorksheetName
    End Function


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

    '''
    ''' Protect or unprotect the sheet.  Return the current setting before
    ''' and changes made.
    '''
    Public Function ProtectSheet( _
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

    Public Function HasCustomFooter() As Boolean
        'Dim prp As Office.DocumentProperty
        'Dim prps As Office.DocumentProperties
        Dim prp As DocumentProperty
        Dim prps As DocumentProperties

        Try
            Try
                prps = Application.ActiveWorkbook.CustomDocumentProperties
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

    Public Sub CustomFooterExists(ByVal hasCustomFooter As Boolean)
        Dim prp As DocumentProperty
        Dim prps As DocumentProperties

        Try
            Try
                prps = Application.ActiveWorkbook.CustomDocumentProperties
                ' Add a new property.
                prp = prps.Add("HasCustomFooter", False, _
                 MsoDocProperties.msoPropertyTypeBoolean, True)
            Catch ex As Exception
                'PLLog.Error(ex, Globals.PROJECT_NAME)
                MessageBox.Show("CustomFooterExists() Unable to add HasCustomFooter property" & ex.Message)
            End Try
        Finally

        End Try
    End Sub

    Public Function DuplicateWorksheet( _
        ByVal sourceSheetName As String, _
        ByVal destinationSheetName As String, _
        Optional ByVal beforeSheetName As String = "", _
        Optional ByVal afterSheetName As String = "" _
    ) As Excel.Worksheet

        With Application.ActiveWorkbook
            For Each ws As Excel.Worksheet In .Worksheets
                If ws.Name = destinationSheetName Then
                    ' TODO: Sheet exists.  Ask user what to do.
                    MessageBox.Show(String.Format("Destination Sheet: >{0}< already exists.", sourceSheetName))
                    Return ws
                End If
            Next

            If beforeSheetName <> "" Then
                .Sheets(sourceSheetName).Copy(Before:=.Sheets(beforeSheetName))
            ElseIf afterSheetName <> "" Then
                .Sheets(sourceSheetName).Copy(After:=.Sheets(afterSheetName))
            Else
                .Sheets(sourceSheetName).Copy()
            End If

            .ActiveSheet.Name = destinationSheetName

            Return .ActiveSheet
        End With
    End Function

    Public Function NewWorksheet( _
        ByVal sheetName As String, _
        Optional ByVal beforeSheetName As String = "", _
        Optional ByVal afterSheetName As String = "" _
    ) As Excel.Worksheet

        With Application.ActiveWorkbook

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
        SafeSheetName = Left(strSafe, Common.cMaxSheetNameLen)
    End Function

    Public Sub ScreenUpdatesOff()
        If True = EnableScreenUpdatesToggle Then
            With Application
                If .Workbooks.Count > 0 Then
                    PriorScreenUpdatingState = .ScreenUpdating
                    .ScreenUpdating = False
                Else
                    ' Assume the intent is to run with screen updates on.
                    PriorScreenUpdatingState = True
                    .ScreenUpdating = False
                End If
            End With
        End If
    End Sub

    Public Sub ScreenUpdatesOn()
        Application.ScreenUpdating = PriorScreenUpdatingState
    End Sub

    Public Sub SetCellValue( _
        ByVal rngR As Excel.Range, _
        ByVal vntValue As Object, _
        Optional ByVal lngHorizontalAlignment As Long = Excel.XlHAlign.xlHAlignLeft, _
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


    'Private Sub DumpPropertyCollection( _
    ' ByVal prps As Office.DocumentProperties, _
    ' ByVal rng As Excel.Range, ByRef i As Integer)
    '    Dim prp As Office.DocumentProperty

    '    For Each prp In prps
    '        rng.Offset(i, 0).Value = prp.Name
    '        Try
    '            If Not prp.Value Is Nothing Then
    '                rng.Offset(i, 1).Value = _
    '                 prp.Value.ToString
    '            End If
    '        Catch
    '            ' Do nothing at all.
    '        End Try
    '        i += 1
    '    Next
    'End Sub

    Public Sub ZapPageBreaks()
        Dim i As Integer
        Dim sht As Excel.Worksheet

        Dim vPB As Excel.VPageBreak
        Dim hPB As Excel.HPageBreak

        With Application

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
End Class
