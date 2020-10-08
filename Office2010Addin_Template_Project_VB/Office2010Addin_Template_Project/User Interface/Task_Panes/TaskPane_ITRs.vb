Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Reflection
Imports System.Runtime
Imports System.Windows.Forms

Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop

Public Class TaskPane_ITRs
    'Private _teams As Data.DataSet

    Private _inputFilePath As String = Common.cDEFAULT_FOLDER

    'Private WithEvents _frmITRDetail As frmITRDetail

    'Public Property frmITRDetail As frmITRDetail
    '    Get
    '        If _frmITRDetail Is Nothing Then
    '            _frmITRDetail = New frmITRDetail()
    '        End If

    '        Return _frmITRDetail
    '    End Get
    '    Set(ByVal Value As frmITRDetail)
    '        _frmITRDetail = Value
    '    End Set
    'End Property

    'Private Sub MethodName() Handles _frmITRDetail.FormClosed
    '    _frmITRDetail = Nothing
    'End Sub

#Region "Initialization"

    Private Sub TaskPane_One_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '' Ensure that any config data we need is available.  Ok to call multiple times.
        ''Config.IntializeApplication()

        'For Each dataTable As Data.DataTable In Config.ConfigInfo.Tables
        '    'Debug.Print(dataTable.TableName)

        '    'For Each dataColumn As Data.DataColumn In dataTable.Columns
        '    '    Debug.Print(dataColumn.ColumnName)
        '    'Next

        '    Select Case dataTable.TableName
        '        Case "team"
        '            For Each dataRow As Data.DataRow In dataTable.Rows
        '                Me.clbOnTimeTeams.Items.Add(dataRow.Item("name")).ToString()
        '                Me.cbOnTimeTeams.Items.Add(dataRow.Item("name")).ToString()
        '                'Debug.Print(dataRow.Item("name").ToString())
        '                'Debug.Print(dataRow.Item("id").ToString())
        '                'Debug.Print(dataRow.Item("team_Id").ToString())
        '            Next

        '            'Case "manager"
        '            '    For Each dataRow As Data.DataRow In dataTable.Rows
        '            '        Debug.Print(dataRow.Item("manager_Text").ToString())
        '            '        Debug.Print(dataRow.Item("ext").ToString())
        '            '        Debug.Print(dataRow.Item("team_Id").ToString())
        '            '    Next

        '    End Select
        'Next
    End Sub

#End Region

#Region "Event Handlers"

    Private Sub btnDisplayITRDetail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDisplayITRDetail.Click
        DisplayITRDetail()
    End Sub

    Private Sub btnGetITRInformation_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetITRInformation.Click
        'GetITRInformation()
    End Sub

#End Region

#Region "Main Function Routines"

    'Sub AddAgeColumn()
    '    With Globals.ThisAddIn.Application
    '        .Columns("D:D").Select()
    '        .Selection.Insert(Shift:=Excel.XlDirection.xlToRight, CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove)
    '        .Range("D5").FormulaR1C1 = "Age"

    '        Dim startRow As Integer
    '        Dim endRow As Integer
    '        Dim startColumn As Integer
    '        Dim endColumn As Integer

    '        startColumn = 4
    '        endColumn = startColumn
    '        startRow = 6
    '        endRow = .ActiveCell.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row

    '        .Range(.Cells(startRow, startColumn), .Cells(endRow, endColumn)).Select()

    '        With .Selection
    '            .FormulaR1C1 = "=TODAY()-RC[-1]"
    '            .NumberFormat = "0"
    '        End With
    '    End With
    'End Sub

    'Private Sub AddPageBreaks()
    '    Dim currentITR As Excel.Range
    '    Dim currentRow As Integer
    '    Dim ws As Excel.Worksheet

    '    ' TODO: Add a constant for this 
    '    currentRow = Common.cFI_SecondITRRow

    '    With Globals.ThisAddIn.Application
    '        ws = .ActiveSheet

    '        currentITR = .Cells(currentRow, 2)

    '        While currentITR.Value > 0
    '            Debug.Print(currentITR.Value)

    '            ws.HPageBreaks.Add(.Cells(currentRow, 15))
    '            ws.VPageBreaks.Add(.Cells(currentRow, 15))

    '            ' Skip past four new rows
    '            currentRow = currentRow + 5
    '            currentITR = .Cells(currentRow, 2)
    '        End While
    '    End With
    'End Sub

    'Private Shared Sub DeleteRow(ByVal ws As Excel.Worksheet, ByVal i As Integer)
    '    'With ws.Rows(i).Interior
    '    '    .Pattern = Excel.Constants.xlSolid
    '    '    .PatternColorIndex = Excel.Constants.xlAutomatic
    '    '    .Color = 5296274
    '    '    .TintAndShade = 0
    '    '    .PatternTintAndShade = 0
    '    'End With

    '    ws.Rows(i).Delete()
    'End Sub

    'Delegate Sub DuplicateRowAction(ByVal workSheet As Excel.Worksheet, ByVal rowNumber As Integer)

    'Private Shared Sub AddListObjects()
    '    With Globals.ThisAddIn.Application
    '        ' Since we have deleted rows, save the workbook to ensure the .UsedRange gets reset.
    '        ' If we don't do this now, when we add the ListObjects using .SpecialCells
    '        ' we will get more rows that we need.  Alternatively could use the ExcelUtil.{LastRow,LastColumn) routines
    '        ' TODO: Make save silent

    '        .ActiveWorkbook.Save()

    '        .Sheets("ITRInfo").ListObjects.Add( _
    '                                    Excel.XlListObjectSourceType.xlSrcRange, _
    '                                    .Range(.Sheets("ITRInfo").Range(Common.cITRHeader_Cell), _
    '                                    .Sheets("ITRInfo").Range(Common.cITRHeader_Cell).SpecialCells(Excel.Constants.xlLastCell)), , _
    '                                    Excel.XlYesNoGuess.xlYes).Name = "SourceITRInfo"

    '        .Sheets("ITRInfoWithResources").ListObjects.Add( _
    '                                    Excel.XlListObjectSourceType.xlSrcRange, _
    '                                    .Range(.Sheets("ITRInfoWithResources").Range(Common.cITRHeader_Cell), _
    '                                    .Sheets("ITRInfoWithResources").Range(Common.cITRHeader_Cell).SpecialCells(Excel.Constants.xlLastCell)), , _
    '                                    Excel.XlYesNoGuess.xlYes).Name = "SourceITRResourceInfo"
    '    End With
    'End Sub

    'Private Sub GetITR(ByVal itrRow As Excel.Range)
    '    itrRow.Cells(1, 2).Value = "Subject"
    '    itrRow.Cells(1, 3).Value = "Application"
    'End Sub

    Private Sub DisplayITRDetail()
        'Dim itrDetail() As ApplicationDS.sp_ITRDetailRow
        'Dim itrID As Integer = Globals.ThisAddIn.Application.ActiveCell.Value

        Dim itrID As Integer
        Dim selectedText As String = Globals.ThisAddIn.Application.Selection.Text

        If Not IsValidITR(selectedText) Then
            MessageBox.Show(String.Format("Selection >{0}< is not a valid ITR #", selectedText))
            Return
        Else
            itrID = CInt(selectedText)

            Try
                ' If GetITRInformation has not populated the Common.ApplicationDS.sp_ITRDetail table 
                ' this won't return anything.

                'itrDetail = Common.ApplicationDS.sp_ITRDetail.Select(String.Format("ID = {0}", itrID))

                ' So, to be safe just use the table adapter to call the stored procedure.

                Dim ta As ApplicationDSTableAdapters.sp_ITRDetailTableAdapter = New ApplicationDSTableAdapters.sp_ITRDetailTableAdapter()
                Dim itrDetail As ApplicationDS.sp_ITRDetailRow
                Dim itrDataTable As ApplicationDS.sp_ITRDetailDataTable
                itrDataTable = ta.GetData(itrID)
                itrDetail = itrDataTable.Rows(0)

                If Not itrDetail Is Nothing Then
                    Dim frmITRDetail As frmITRDetail = New frmITRDetail()

                    For Each ctrl As Control In frmITRDetail.Controls
                        Try
                            If Not ctrl Is Nothing Then
                                If TypeOf ctrl Is TextBox Then
                                    Dim fieldName = ctrl.Name.Replace("txt", "")
                                    ctrl.Text = itrDetail.Item(fieldName).ToString()
                                End If
                            Else
                                MessageBox.Show("ctrl is nothing??")
                            End If
                        Catch ex As System.ArgumentException
                            ' Column likely does not exist on row/table
                        Catch ex As Exception
                            MessageBox.Show(String.Format("Exception: {0}.{1}() - {2}",
                                System.Reflection.Assembly.GetExecutingAssembly().FullName,
                                System.Reflection.MethodInfo.GetCurrentMethod().Name,
                                ex.ToString()
                                ))
                        End Try

                    Next

                    'Dim date1 As Date = Today()
                    'Dim date2 As Date = Convert.ToDateTime(itrDetail.RequestID).Date
                    'Dim age As TimeSpan = date1.Subtract(date2)

                    'frmITRDetail.txtAge.Text = age.Days.ToString()

                    frmITRDetail.txtAge.Text = (Today().Subtract(Date.Parse(itrDetail.RequestID))).Days.ToString()

                    'frmDetail.txtID.Text = itrDetail.ID

                    'frmDetail.txtApplication.Text = itrDetail.Application
                    'frmDetail.txtDescription.Text = itrDetail.Description
                    'frmDetail.txtComments.Text = itrDetail.Comments
                    'frmDetail.txtJustification.Text = itrDetail.Justification
                    frmITRDetail.Show()
                Else
                    MessageBox.Show("Could not find ITR")
                End If


            Catch ex As Exception

                MessageBox.Show(String.Format("Exception: {0}.{1}() - {2}",
                             System.Reflection.Assembly.GetExecutingAssembly().FullName,
                             System.Reflection.MethodInfo.GetCurrentMethod().Name,
                             ex.ToString()
                             ))
            End Try
        End If
    End Sub

    Private Function IsValidITR(ByVal selectedText As String) As Boolean
        If selectedText.Trim.Length <> 5 Then
            Return False
        End If

        Return RegularExpressions.Regex.Match(selectedText, "[0-9]{5}").Success
    End Function

    'Private Shared Sub DuplicateInputSheet()
    '    With Globals.ThisAddIn
    '        Common.ExcelHelper.DuplicateWorksheet("SourceITRs", "ITRInfo", "SourceITRs")
    '        .Application.ActiveWorkbook.Sheets("ITRInfo").Columns(Common.cITRInfo_CommentColumns).Delete()

    '        Common.ExcelHelper.DuplicateWorksheet("SourceITRs", "ITRInfoWithResources", "ITRInfo")
    '        .Application.ActiveWorkbook.Sheets("ITRInfoWithResources").Columns(Common.cITRITRInfoWithResources_CommentColumns).Delete()

    '        Common.ExcelHelper.DuplicateWorksheet("SourceITRs", "FormatedITRs", "SourceITRs")
    '    End With
    'End Sub
    Private Sub btnDuplicateInputSheet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'DuplicateInputSheet()
    End Sub

    Private Sub btnFormatSourceITRs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'FormatSourceITRs()
    End Sub

    Private Sub FormatSourceITRs()
        'Common.ExcelUtil.ScreenUpdatesOff()
        'FormatITRs()
        'Common.ExcelUtil.ScreenUpdatesOn()
    End Sub

    'Private Sub GetITRInformation()
    '    Dim currentCell As Excel.Range = Globals.ThisAddIn.Application.ActiveCell
    '    Dim firstRow As Integer = currentCell.Row
    '    'Dim lastRow As Integer = currentCell.SpecialCells(Excel.Constants.xlLastCell).Row
    '    Dim lastRow As Integer = Common.ExcelHelper.FindLastRow(currentCell)
    '    'Dim foundRows As Hashtable = New Hashtable()
    '    Dim activeSheet As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet
    '    lastRow = lastRow
    '    Dim itrRow As Excel.Range

    '    'Debug.Print(String.Format("{0} - {1}", firstRow, lastRow))

    '    Dim ta As ApplicationDSTableAdapters.sp_ITRDetailTableAdapter = New ApplicationDSTableAdapters.sp_ITRDetailTableAdapter()
    '    Dim itrDataRow As ApplicationDS.sp_ITRDetailRow
    '    Dim itrDataTable As ApplicationDS.sp_ITRDetailDataTable

    '    Dim tr As ApplicationDS.sp_ITRDetailRow

    '    For i As Integer = firstRow To lastRow
    '        itrRow = activeSheet.Cells(i, currentCell.Column)

    '        Common.ApplicationDS.sp_ITRDetail.ImportRow(itrDataRow)

    '        itrRow.Cells(1, 2).Value = itrDataRow.Subject
    '        itrRow.Cells(1, 3).Value = itrDataRow.Application.Replace("/Life IT/Technical Services/", "")

    '        Try
    '            itrRow.Cells(1, 4).Value = itrDataRow.Description.Length    ' Current Condition
    '        Catch ex As Exception
    '            itrRow.Cells(1, 4).Value = "<none>"
    '        End Try
    '        Try
    '            itrRow.Cells(1, 5).Value = itrDataRow.Justification.Length  ' Desired Outcome
    '        Catch ex As Exception
    '            itrRow.Cells(1, 5).Value = "<none>"
    '        End Try
    '        Try
    '            itrRow.Cells(1, 6).Value = itrDataRow.Comments.Length       ' Comments
    '        Catch ex As Exception
    '            itrRow.Cells(1, 6).Value = "<none>"
    '        End Try
    '    Next i
    'End Sub


    'Sub HilightDuplicateRows(ByVal worksheetName As String)
    '    Dim ws As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets(worksheetName)

    '    ws.Columns(Common.cRESOURCEID_COLUMN).Delete()

    '    Dim previousITR As Long
    '    Dim currentITR As Long
    '    Dim lastITRrow As Integer
    '    Dim foundMatching As Boolean = False

    '    lastITRrow = ws.Range("B5").SpecialCells(Excel.Constants.xlLastCell).Row
    '    ' and then walk the list of ITRs and delete duplicates
    '    Debug.Print(lastITRrow)

    '    For i As Integer = lastITRrow To Common.cFirstITRRow Step -1
    '        Debug.Print(String.Format("previousITR {0}  currentITR {1}   i {2}  lastITRRow {3}", _
    '            previousITR, ws.Cells(i, Common.cITRID_COLUMN).Value, i, lastITRrow))
    '        currentITR = CLng(ws.Cells(i, Common.cITRID_COLUMN).Value)

    '        If currentITR = previousITR Then
    '            Debug.Print("Matching ITR")
    '            HilightRow(ws, i)
    '            lastITRrow = lastITRrow - 1

    '            If Not foundMatching Then
    '                ' Hilight the previous row which matched the current row.
    '                HilightRow(ws, i - 1)
    '            End If

    '            foundMatching = True
    '        Else
    '            Debug.Print("New ITR")
    '            foundMatching = False
    '        End If

    '        previousITR = currentITR
    '    Next i
    'End Sub

    'Private Shared Sub HilightRow2(ByVal ws As Excel.Worksheet, ByVal i As Integer)
    '    With ws.Rows(i).Interior
    '        .Pattern = Excel.Constants.xlGray50
    '        .PatternColorIndex = Excel.Constants.xlAutomatic
    '        .Color = 5296274
    '        .TintAndShade = 0
    '        .PatternTintAndShade = 0
    '    End With

    '    'ws.Rows(i).Delete()
    'End Sub

    'Private Shared Sub HilightRow(ByVal ws As Excel.Worksheet, ByVal i As Integer)
    '    With ws.Rows(i).Interior
    '        .Pattern = Excel.Constants.xlSolid
    '        .PatternColorIndex = Excel.Constants.xlAutomatic
    '        .Color = 5296274
    '        .TintAndShade = 0
    '        .PatternTintAndShade = 0
    '    End With

    '    'ws.Rows(i).Delete()
    'End Sub

    'Private Sub ProcessDuplicateRows(ByVal worksheetName As String, ByVal deleteAction As DuplicateRowAction, ByVal hilightAction As DuplicateRowAction)
    '    Dim ws As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets(worksheetName)

    '    Dim previousITR As Long
    '    Dim currentITR As Long
    '    Dim lastITRrow As Integer = ws.Range("B5").SpecialCells(Excel.Constants.xlLastCell).Row

    '    ' and then walk the list of ITRs and delete duplicates
    '    Debug.Print(lastITRrow)

    '    For i As Integer = lastITRrow To Common.cFirstITRRow Step -1
    '        'Debug.Print(String.Format("previousITR {0}  currentITR {1}   i {2}  lastITRRow {3}", _
    '        '    previousITR, ws.Cells(i, Common.cITRID_COLUMN).Value, i, lastITRrow))
    '        currentITR = CLng(ws.Cells(i, Common.cITRID_COLUMN).Value)

    '        If currentITR = previousITR Then
    '            'Debug.Print("Matching ITR")
    '            ' First hilight the previous matching ITR
    '            hilightAction(ws, i + 1)
    '            ' Then take the deleteAction on the current row.  This might be a hilight operation.
    '            deleteAction(ws, i)
    '            lastITRrow = lastITRrow - 1
    '        Else
    '            'Debug.Print("New ITR")
    '        End If

    '        previousITR = currentITR
    '    Next i
    'End Sub

    'Private Sub ProcessDuplicates()
    '    'Common.ExcelUtil.ScreenUpdatesOff()
    '    ' Delete the Resource ID column from the ITRInfo sheet
    '    Globals.ThisAddIn.Application.Sheets("ITRInfo").Columns(Common.cRESOURCEID_COLUMN).Delete()
    '    ProcessDuplicateRows("ITRInfo", AddressOf DeleteRow, AddressOf HilightRow)

    '    ' Just Hilight the ITRs with multiple Resources.  We keep the Resource ID column here.
    '    ProcessDuplicateRows("ITRInfoWithResources", AddressOf HilightRow2, AddressOf HilightRow)
    '    'Common.ExcelUtil.ScreenUpdatesOn()
    'End Sub

    'Private Sub ProcessDynamicOutput()
    '    'TODO: Drive this off Config file and/or TaskPane_Config
    '    'Common.ExcelUtil.ScreenUpdatesOff()

    '    Globals.ThisAddIn.Application.Sheets("dynamicRepo").Name = "SourceITRs"
    '    ' Move down a row so the TOC doesn't step on things
    '    Globals.ThisAddIn.Application.ActiveSheet.Rows("1:1").Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftDown, CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove)
    '    AddAgeColumn()
    '    UpdateApplicationNames()
    '    Globals.ThisAddIn.Application.ActiveSheet.Range("A1").Select()
    '    SaveOutputFile()

    '    'Common.ExcelUtil.ScreenUpdatesOn()
    'End Sub

    'Sub RemoveDuplicateRows(ByVal worksheetName As String, ByVal action As DuplicateRowAction)
    '    Dim ws As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets(worksheetName)

    '    ws.Columns(Common.cRESOURCEID_COLUMN).Delete()

    '    Dim previousITR As Long
    '    Dim currentITR As Long
    '    Dim lastITRrow As Integer
    '    Dim foundMatching As Boolean = False

    '    lastITRrow = ws.Range("B5").SpecialCells(Excel.Constants.xlLastCell).Row
    '    ' and then walk the list of ITRs and delete duplicates
    '    Debug.Print(lastITRrow)

    '    For i As Integer = lastITRrow To Common.cFirstITRRow Step -1
    '        Debug.Print(String.Format("previousITR {0}  currentITR {1}   i {2}  lastITRRow {3}", _
    '            previousITR, ws.Cells(i, Common.cITRID_COLUMN).Value, i, lastITRrow))
    '        currentITR = CLng(ws.Cells(i, Common.cITRID_COLUMN).Value)

    '        If currentITR = previousITR Then
    '            Debug.Print("Matching ITR")
    '            DeleteRow(ws, i)
    '            lastITRrow = lastITRrow - 1

    '            If Not foundMatching Then
    '                ' Hilight the previous row which matched the current row.
    '                HilightRow(ws, i - 1)
    '            End If

    '            foundMatching = True
    '        Else
    '            Debug.Print("New ITR")
    '            foundMatching = False
    '        End If

    '        previousITR = currentITR
    '    Next i
    'End Sub

    'Private Sub SaveOutputFile()
    '    Dim outputFileName As String = GetDefaultOutputFileName()

    '    With Globals.ThisAddIn.Application
    '        outputFileName = Common.ExcelHelper.GetSaveFileName(Common.cDEFAULT_FOLDER, outputFileName, "Enter Save As Name")

    '        If outputFileName.Length > 0 Then
    '            .ActiveWorkbook.SaveAs(Filename:=outputFileName, FileFormat:=Excel.XlFileFormat.xlOpenXMLWorkbook, CreateBackup:=False)
    '        End If
    '    End With
    'End Sub

    'Sub FormatITRs()
    '    Globals.ThisAddIn.Application.ActiveWorkbook.Sheets("FormatedITRs").Activate()

    '    With Globals.ThisAddIn.Application
    '        '.Rows("2:3").Select()
    '        '.Selection.Delete(Shift:=Excel.XlDirection.xlUp)
    '        .Rows("5:5").Select()
    '    End With

    '    FormatPrinterOutput()

    '    MergeAssignedResources("FormatedITRs")

    '    MoveITRCommentsBelow()

    '    FormatColumns_FormatedITRs()

    '    AddPageBreaks()

    'End Sub

    'Sub FormatColumns_FormatedITRs()

    '    With Globals.ThisAddIn.Application
    '        .Columns(Common.cFI_Application_Column_Range).ColumnWidth = 25
    '        .Columns(Common.cFI_ITRID_Column_Range).ColumnWidth = 5
    '        .Columns(Common.cFI_EnteredOn_Column_Range).ColumnWidth = 9
    '        .Columns(Common.cFI_Age_Column_Range).ColumnWidth = 5
    '        .Columns(Common.cFI_EnteredBy_Column_Range).ColumnWidth = 9
    '        .Columns(Common.cFI_RequestedBy_Column_Range).ColumnWidth = 8
    '        .Columns(Common.cFI_ReleaseNbr_Column_Range).ColumnWidth = 7
    '        .Columns(Common.cFI_PatRank_Column_Range).ColumnWidth = 4
    '        .Columns(Common.cFI_Category_Column_Range).ColumnWidth = 12
    '        .Columns(Common.cFI_Status_Column_Range).ColumnWidth = 10
    '        .Columns(Common.cFI_Severity_Column_Range).ColumnWidth = 10
    '        .Columns(Common.cFI_LOE_Column_Range).ColumnWidth = 10
    '        .Columns(Common.cFI_Subject_Column_Range).ColumnWidth = 35
    '        .Columns(Common.cFI_Resource_Column_Range).Columnwidth = 10
    '    End With
    'End Sub

    'Sub FormatColumns_PivotTable(ByVal workSheetName As String)
    '    ' TODO: Vary this by sheet.  Some may want to be landscape, some portrait.  Perhaps pass parameter.
    '    With Globals.ThisAddIn.Application
    '        .Worksheets(workSheetName).Activate()
    '        .Columns(Common.cPT_ITR_Column_Range).ColumnWidth = 75
    '        .Columns(Common.cPT_Count_Column_Range).ColumnWidth = 6
    '    End With
    'End Sub

    'Sub FormatPrinterOutput()
    '    With Globals.ThisAddIn.Application
    '        With .ActiveSheet.PageSetup
    '            .PrintTitleRows = "$5:$5"
    '            .PrintTitleColumns = ""
    '        End With

    '        .ActiveSheet.PageSetup.PrintArea = ""
    '        ' TODO: Take Center Header as parameter
    '        With .ActiveSheet.PageSetup
    '            .LeftHeader = ""
    '            .CenterHeader = "Open Delivery Services ITRs"
    '            .RightHeader = ""
    '            .LeftFooter = ""
    '            .CenterFooter = ""
    '            .RightFooter = ""
    '            .LeftMargin = .Application.InchesToPoints(0.25)
    '            .RightMargin = .Application.InchesToPoints(0.25)
    '            .TopMargin = .Application.InchesToPoints(0.75)
    '            .BottomMargin = .Application.InchesToPoints(0.75)
    '            .HeaderMargin = .Application.InchesToPoints(0.3)
    '            .FooterMargin = .Application.InchesToPoints(0.3)
    '            .PrintHeadings = False
    '            .PrintGridlines = False
    '            .PrintComments = Excel.XlPrintLocation.xlPrintNoComments
    '            .PrintQuality = 600
    '            .CenterHorizontally = False
    '            .CenterVertically = False
    '            .Orientation = Excel.XlPageOrientation.xlLandscape
    '            .Draft = False
    '            .PaperSize = Excel.XlPaperSize.xlPaperLegal
    '            .FirstPageNumber = Excel.Constants.xlAutomatic
    '            .Order = Excel.XlOrder.xlDownThenOver
    '            .BlackAndWhite = False
    '            .Zoom = 100
    '            .PrintErrors = Excel.XlPrintErrors.xlPrintErrorsDisplayed
    '            .OddAndEvenPagesHeaderFooter = False
    '            .DifferentFirstPageHeaderFooter = False
    '            .ScaleWithDocHeaderFooter = True
    '            .AlignMarginsHeaderFooter = True
    '            .EvenPage.LeftHeader.Text = ""
    '            .EvenPage.CenterHeader.Text = ""
    '            .EvenPage.RightHeader.Text = ""
    '            .EvenPage.LeftFooter.Text = ""
    '            .EvenPage.CenterFooter.Text = ""
    '            .EvenPage.RightFooter.Text = ""
    '            .FirstPage.LeftHeader.Text = ""
    '            .FirstPage.CenterHeader.Text = ""
    '            .FirstPage.RightHeader.Text = ""
    '            .FirstPage.LeftFooter.Text = ""
    '            .FirstPage.CenterFooter.Text = ""
    '            .FirstPage.RightFooter.Text = ""
    '        End With
    '    End With
    'End Sub

    'Private Sub MergeDuplicateRows()
    '    Dim currentCell As Excel.Range = Globals.ThisAddIn.Application.ActiveCell
    '    Dim firstRow As Integer = currentCell.Row
    '    'Dim lastRow As Integer = currentCell.SpecialCells(Excel.Constants.xlLastCell).Row
    '    Dim lastRow As Integer = Common.ExcelHelper.FindLastRow(currentCell)
    '    Dim foundITRs As Dictionary(Of String, Excel.Range) = New Dictionary(Of String, Excel.Range)
    '    Dim duplicateITRs As Dictionary(Of String, String) = New Dictionary(Of String, String)
    '    Dim activeSheet As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet
    '    Dim rowsToDelete As List(Of String) = New List(Of String)
    '    Dim rng As Excel.Range

    '    Debug.Print(String.Format("{0} - {1}", firstRow, lastRow))

    '    For i As Integer = lastRow To firstRow Step -1
    '        rng = activeSheet.Cells(i, currentCell.Column)

    '        Debug.Print(String.Format("{0} - {1} - {2}", i, rng.Value, rng.Cells(1, 7).Value))

    '        If foundITRs.ContainsKey(rng.Value) Then
    '            If duplicateITRs.ContainsKey(rng.Value) Then
    '                ' This is not the first duplicate.  Add resource if not there already.
    '                If Not duplicateITRs(rng.Value).Contains(rng.Cells(1, 7).Value) Then
    '                    duplicateITRs(rng.Value) = duplicateITRs(rng.Value) & ", " & rng.Cells(1, 7).Value
    '                End If

    '            Else
    '                ' Add the Resource from the duplicate row
    '                duplicateITRs(rng.Value) = rng.Cells(1, 7).Value
    '            End If

    '            rowsToDelete.Add(String.Format("{0}:{1}", i, i))
    '        Else
    '            foundITRs(rng.Value) = rng
    '        End If
    '    Next i

    '    For Each row As KeyValuePair(Of String, Excel.Range) In foundITRs
    '        Debug.Print(String.Format("{0} - {1} - {2}", row.Value.Row, row.Value.Cells(1, 1).Value, row.Value.Cells(1, 7).Value))
    '    Next

    '    For Each row As KeyValuePair(Of String, String) In duplicateITRs
    '        Debug.Print(String.Format("{0} - {1}", row.Key, row.Value))

    '        Dim resources As String = activeSheet.Cells(foundITRs(row.Key).Row, 7).Value

    '        If row.Value.Contains(resources) Then
    '            activeSheet.Cells(foundITRs(row.Key).Row, 7).Value = row.Value
    '        Else
    '            activeSheet.Cells(foundITRs(row.Key).Row, 7).Value = resources & ", " & row.Value
    '        End If

    '    Next

    '    For Each row As String In rowsToDelete
    '        Debug.Print(row)
    '        activeSheet.Rows(row).Delete()
    '    Next
    'End Sub

    'Private Sub MergeResourceAndDeleteRow(ByVal ws As Object, ByVal i As Integer)
    '    Dim allResources As String = ws.Cells(i + 1, 14).Value
    '    Dim resource As String = ws.Cells(i, 14).Value

    '    allResources = String.Format("{0}, {1}", allResources, resource)
    '    ws.Cells(i + 1, 14).Value = allResources
    '    ws.Rows(i).Delete()
    'End Sub

    'Private Sub MergeAssignedResources(ByVal worksheetName As String)
    '    Dim ws As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets(worksheetName)
    '    ' Walk the list of ITRs looking for duplicates and merge the assigned resources if any.
    '    Dim previousITR As Long
    '    Dim currentITR As Long
    '    Dim lastITRrow As Integer = ws.Range("B5").SpecialCells(Excel.Constants.xlLastCell).Row

    '    ' and then walk the list of ITRs and delete duplicates
    '    Debug.Print(lastITRrow)

    '    For i As Integer = lastITRrow To Common.cFirstITRRow Step -1
    '        'Debug.Print(String.Format("previousITR {0}  currentITR {1}   i {2}  lastITRRow {3}", _
    '        '    previousITR, ws.Cells(i, Common.cITRID_COLUMN).Value, i, lastITRrow))
    '        currentITR = CLng(ws.Cells(i, Common.cITRID_COLUMN).Value)

    '        If currentITR = previousITR Then
    '            'Debug.Print("Matching ITR")
    '            ' First hilight the previous matching ITR
    '            MergeResourceAndDeleteRow(ws, i)
    '            lastITRrow = lastITRrow - 1
    '        Else
    '            'Debug.Print("New ITR")
    '        End If

    '        previousITR = currentITR
    '    Next i

    'End Sub

    'Sub MoveCommentsBelowAndGroup()
    '    Dim currentCell As Excel.Range

    '    With Globals.ThisAddIn.Application
    '        currentCell = .ActiveCell

    '        .Rows(currentCell.Row).Select()

    '        ' Make room for comments

    '        .Selection.Insert(CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove)    ' Description of Current Condition
    '        .Selection.Insert(CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove)    ' Description fo Desired Outcome
    '        .Selection.Insert(CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove)    ' Prioritization Comments
    '        .Selection.Insert(CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove)    ' Comments

    '        ' Add Row Headers

    '        currentCell.Offset(-4, 0).Value = "Description of Current Condition"
    '        currentCell.Offset(-3, 0).Value = "Description of Desired Outcome"
    '        currentCell.Offset(-2, 0).Value = "Prioritization Comments"
    '        currentCell.Offset(-1, 0).Value = "Comments"

    '        ' Merge all cells for each row so have a long field for values.  This makes
    '        ' cell wrapping work better

    '        currentCell.Offset(-4, 1).Range(.Cells(1, 1), .Cells(1, 13)).Merge()
    '        currentCell.Offset(-3, 1).Range(.Cells(1, 1), .Cells(1, 13)).Merge()
    '        currentCell.Offset(-2, 1).Range(.Cells(1, 1), .Cells(1, 13)).Merge()
    '        currentCell.Offset(-1, 1).Range(.Cells(1, 1), .Cells(1, 13)).Merge()

    '        ' Copy the values from the previous row into the new locations.

    '        currentCell.Offset(-4, 1).Value = currentCell.Offset(-5, 14).Value
    '        currentCell.Offset(-3, 1).Value = currentCell.Offset(-5, 15).Value
    '        currentCell.Offset(-2, 1).Value = currentCell.Offset(-5, 16).Value
    '        currentCell.Offset(-1, 1).Value = currentCell.Offset(-5, 17).Value

    '        ' Update the row height  based on the length of the comments

    '        Call SetRowHeight(currentCell.Offset(-4, 1))
    '        Call SetRowHeight(currentCell.Offset(-3, 1))
    '        Call SetRowHeight(currentCell.Offset(-2, 1))
    '        Call SetRowHeight(currentCell.Offset(-1, 1))

    '        ' Group the new comment rows so they can be collapsed.

    '        currentCell.Offset(-4, 0).Range(.Cells(1, 1), .Cells(4, 1)).Rows.Group()

    '        ' TODO: Probably add the Page Breaks here.  Maybe optional config flag to drive
    '    End With

    'End Sub

    'Sub SetRowHeight(ByVal commentCell As Excel.Range)
    '    Dim rowHeight As Single
    '    Dim textLength As Integer

    '    textLength = Len(commentCell.Value)

    '    '    If textLength > 160 Then
    '    rowHeight = ((textLength \ 225.0#) * 15.75) + 15.75
    '    '    Else
    '    '        rowHeight = 15.75
    '    '    End If

    '    ' Max

    '    If rowHeight > 409 Then
    '        rowHeight = 409
    '    End If

    '    Debug.Print(textLength, rowHeight)
    '    commentCell.Rows.RowHeight = rowHeight
    'End Sub

    'Sub UpdateApplicationNames()
    '    With Globals.ThisAddIn.Application
    '        .Columns("A:A").Select()

    '        ' Strip out "/Life IT/Technical Services"

    '        .Selection.Replace(What:="/Life IT/Technical Services/", Replacement:="", _
    '            LookAt:=Excel.XlLookAt.xlPart, SearchOrder:=Excel.XlSearchOrder.xlByRows, _
    '            MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False)

    '        ' Shorten "Application Infrastructure"

    '        .Selection.Replace(What:="Application Infrastructure/", Replacement:="AI/", _
    '            LookAt:=Excel.XlLookAt.xlPart, SearchOrder:=Excel.XlSearchOrder.xlByRows, _
    '            MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False)

    '        ' Shorten "Common Architecture Components"

    '        .Selection.Replace(What:="Common Architecture Components", Replacement:="Common Arch", _
    '            LookAt:=Excel.XlLookAt.xlPart, SearchOrder:=Excel.XlSearchOrder.xlByRows, _
    '            MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False)

    '        ' Shorten "Development Support Services"

    '        .Selection.Replace(What:="Development Support Services/", Replacement:="DSS/", _
    '            LookAt:=Excel.XlLookAt.xlPart, SearchOrder:=Excel.XlSearchOrder.xlByRows, _
    '            MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False)

    '        ' Shorten "EAC Console Apps"

    '        .Selection.Replace(What:="EAC Console Apps/", Replacement:="EAC/", _
    '            LookAt:=Excel.XlLookAt.xlPart, SearchOrder:=Excel.XlSearchOrder.xlByRows, _
    '            MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False)

    '        ' Shorten "Integration Services"

    '        .Selection.Replace(What:="Integration Services/", Replacement:="IS/", _
    '            LookAt:=Excel.XlLookAt.xlPart, SearchOrder:=Excel.XlSearchOrder.xlByRows, _
    '            MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False)

    '        ' Shorten "Biztalk Solution Development"

    '        .Selection.Replace(What:="Biztalk Solution Development", Replacement:="BizTalk", _
    '            LookAt:=Excel.XlLookAt.xlPart, SearchOrder:=Excel.XlSearchOrder.xlByRows, _
    '            MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False)

    '        ' Shorten "TrackSuite"

    '        .Selection.Replace(What:="TrackSuite/", Replacement:="TS/", _
    '            LookAt:=Excel.XlLookAt.xlPart, SearchOrder:=Excel.XlSearchOrder.xlByRows, _
    '            MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False)

    '        ' Shorten "Multi Case Reporting"

    '        .Selection.Replace(What:="Multi Case Reporting", Replacement:="MCR", _
    '            LookAt:=Excel.XlLookAt.xlPart, SearchOrder:=Excel.XlSearchOrder.xlByRows, _
    '            MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False)
    '    End With
    'End Sub

#End Region


End Class
