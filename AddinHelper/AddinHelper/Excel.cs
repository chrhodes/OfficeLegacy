using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace AddinHelper
{
    public class Excel
    {
        public const int cMaxSheetNameLen = 30;
        #region "Debug Constants"
        //Public Shared ScreenUpdatesToggleEnabled As Boolean = True
        #endregion

        public static object HostApp;
        public static object AddInInstance;
        // Some stuff about who started us and how they started us.
        public static string AppName;

        public static string AppVersion;

        public static short StartMode;
        public const int cHeaderFontSize = 12;
        public const int cHeaderFontSizeMedium = 10;

        public const int cHeaderFontSizeSmall = 8;
        public static XlCalculation PriorCalculationState;

        public static bool PriorScreenUpdatingState = false;

        public enum WrapText : byte
        {
	        Yes = 1,
	        No = 0
        }

        public enum MakeBold : byte
        {
	        Yes = 1,
	        No = 0
        }

        public enum UnderLine : byte
        {
	        Yes = 1,
	        No = 0
        }

        //Private Const cMODULE_NAME As String = Globals.PROJECT_NAME & ".modExcel"

	        // Hard coded for Excel? 1/72 Inches.
        public const int cPointsToInch = 72;
	        // Times before this
        public const int cStartHour = 8;
	        // and after this are hilighted.
        public const int cEndHour = 20;

        public static Microsoft.Office.Interop.Excel.Application ExcelApplication
        {
            get;
            set;
        }

        public static bool DisplayScreenUpdates
        {
            get;
            set;
        }

        #region Main Function Routines

        public static void AddColumnHeaderToSheet(
            ref Worksheet ws,
            int columnNumber,
            int columnWidth,
            int headerRow, 
            string headerTitle = "", 
            int headerFontSize = cHeaderFontSize, 
            MakeBold headerBold = MakeBold.Yes,
            XlUnderlineStyle headerUnderline = XlUnderlineStyle.xlUnderlineStyleNone, 
            WrapText headerWrapText = WrapText.Yes,
            XlHAlign headerHorizontalAlignment = XlHAlign.xlHAlignGeneral, 
            int headerOrientation = 0)
        {
            Range column = (Range)ws.Columns[columnNumber];

            column.ColumnWidth = columnWidth;

	        if ( ! string.IsNullOrEmpty(headerTitle))
            {
		        Range header = (Range)ws.Cells[headerRow, columnNumber];
		        header.Value = headerTitle;
		        header.Font.Size = headerFontSize;
		        header.Font.Bold = headerBold;
		        header.Font.Underline = headerUnderline;
		        header.WrapText = headerWrapText;
		        header.HorizontalAlignment = headerHorizontalAlignment;
		        header.Orientation = headerOrientation;
	        }
        }

        public static void AddColumnToSheet(
            ref Worksheet ws, 
            int columnNumber, 
            int columnWidth, 
            bool columnWrapText, 
            string columnNumberFormat,
            XlDirection shiftDirection, 
            XlInsertFormatOrigin insertFormatOrigin, 
            int headerRow, 
            string headerTitle = "", 
            int headerFontSize = cHeaderFontSize,
            MakeBold headerBold = MakeBold.Yes,
            XlUnderlineStyle headerUnderline = XlUnderlineStyle.xlUnderlineStyleSingle,
            WrapText headerWrapText = WrapText.Yes, 
            XlHAlign headerHorizontalAlignment = XlHAlign.xlHAlignCenter, 
            int headerOrientation = 0)
        {
            // Insert the new column and apply things that pertain to all cells in the column
            //((Range)ws.Columns[columnNumber]).Insert(Shift: shiftDirection, CopyOrigin: insertFormatOrigin);

            Range newColumn = (Range)ws.Columns[columnNumber];
            newColumn.Insert(Shift: shiftDirection, CopyOrigin: insertFormatOrigin);
            newColumn.WrapText = columnWrapText;
            newColumn.NumberFormat = columnNumberFormat;

            // Pass all the rest on
            AddColumnHeaderToSheet(ref ws, 
                columnNumber,
                columnWidth,
                headerRow, 
                headerTitle, 
                headerFontSize, 
                headerBold, 
                headerUnderline, 
                headerWrapText, 
                headerHorizontalAlignment, 
                headerOrientation);
        }

        public static void AddCommentToCell(ref Worksheet ws, int column, int row, string text, int headerFontSize = cHeaderFontSize, bool headerBold = true, bool headerUnderline = true, bool headerWrapText = true, XlHAlign headerHorizontalAlignment = XlHAlign.xlHAlignGeneral)
        {
	        ((Range)ws.Cells[row, column]).AddComment(text);

	        // TODO: Determine how to format the text differently.
	        //With ws
	        //    With .Cells(row, column)
	        //        .Value = headerTitle
	        //        .Font.Size = headerFontSize
	        //        .Font.Bold = headerBold
	        //        .Font.Underline = headerUnderline
	        //        .WrapText = headerWrapText
	        //        .HorizontalAlignment = headerHorizontalAlignment
	        //    End With
	        //End With
        }

        public static void AddContentToCell(Range rng, string text, 
            int fontSize = 10,
            MakeBold bold = MakeBold.No, 
            XlUnderlineStyle underline = XlUnderlineStyle.xlUnderlineStyleNone, 
            WrapText wrapText = WrapText.No, 
            XlHAlign horizontalAlignment = XlHAlign.xlHAlignGeneral)
        {
            rng.Value = text;
	        rng.Font.Size = fontSize;
	        rng.Font.Bold = (int)bold;
	        rng.Font.Underline = underline;
	        rng.WrapText = (int)wrapText;
	        rng.HorizontalAlignment = horizontalAlignment;
            //rng.NumberFormat = "MM/DD/YYYY";
        }

        //Public Sub ApplicationInfo()
        //    Try
        //        Debug.Print("Application.CommonAppDataPath:" & Application.CommonAppDataPath.ToString)
        //        Debug.Print("Application.CommonAppDataRegistry:" & Application.CommonAppDataRegistry.ToString)
        //        Debug.Print("Application.CompanyName:" & Application.CompanyName.ToString)
        //        Debug.Print("Application.CurrentCulture:" & Application.CurrentCulture.ToString)
        //        Debug.Print("Application.CurrentInputLanguage:" & Application.CurrentInputLanguage.ToString)
        //        Debug.Print("Application.ExecutablePath:" & Application.ExecutablePath.ToString)
        //        Debug.Print("Application.LocalUserAppDataPath:" & Application.LocalUserAppDataPath.ToString)
        //        Debug.Print("Application.ProductName:" & Application.ProductName.ToString)
        //        Debug.Print("Application.ProductVersion:" & Application.ProductVersion.ToString)
        //        Debug.Print("Application.SafeTopLevelCaptionFormat:" & Application.SafeTopLevelCaptionFormat.ToString)
        //        Debug.Print("Application.StartupPath:" & Application.StartupPath.ToString)
        //        Debug.Print("Application.UserAppDataPath:" & Application.UserAppDataPath.ToString)
        //        Debug.Print("Application.UserAppDataRegistry:" & Application.UserAppDataRegistry.ToString)

        //        Debug.Print("ThisAddin.Application.StartupPath:" & Globals.ThisAddIn.Application.StartupPath.ToString)
        //        Debug.Print("ThisAddin.Application.ActiveWorkbook.Name:" & Globals.ThisAddIn.Application.ActiveWorkbook.Name.ToString)
        //        Debug.Print("ThisAddin.Application.ActiveWorkbook.Path:" & Globals.ThisAddIn.Application.ActiveWorkbook.Path.ToString)
        //        Debug.Print("ThisAddin.Application.ActiveWorkbook.FullName:" & Globals.ThisAddIn.Application.ActiveWorkbook.FullName.ToString)
        //        Debug.Print("ThisAddin.Application.ActiveWorkbook.FullNameURLEncoded:" & Globals.ThisAddIn.Application.ActiveWorkbook.FullNameURLEncoded.ToString)

        //        Debug.Print("ThisAddin.Application.DefaultFilePath:" & Globals.ThisAddIn.Application.DefaultFilePath.ToString)
        //        Debug.Print("ThisAddin.Application.Name:" & Globals.ThisAddIn.Application.Name.ToString)
        //        Debug.Print("ThisAddin.Application.NetworkTemplatesPath:" & Globals.ThisAddIn.Application.NetworkTemplatesPath.ToString)
        //        Debug.Print("ThisAddin.Application.Path:" & Globals.ThisAddIn.Application.Path.ToString)
        //    Catch ex As Exception
        //        MessageBox.Show("ApplicationInfo():" & ex.ToString)
        //    End Try
        //End Sub

        public static void CalculationsOff()
        {
	        // Don't bother trying to save current if no open workbooks.

	        if (ExcelApplication.Workbooks.Count > 0)
            {
		        PriorCalculationState = ExcelApplication.Calculation;
		        ExcelApplication.Calculation = XlCalculation.xlCalculationManual;
	        }
            else
            {
		        // Assume the intent is to run with calculation and screen updates on.
		        // Hopefully we never get called with no workbooks open.
		        PriorCalculationState = XlCalculation.xlCalculationAutomatic;
	        }
        }

        public static void CalculationsOn()
        {
	        ExcelApplication.Calculation = PriorCalculationState;
        }

        public static void CustomFooterExists(bool hasCustomFooter)
        {
		    try
            {
			    Office.DocumentProperties prps = (Office.DocumentProperties)ExcelApplication.ActiveWorkbook.CustomDocumentProperties;
			    // Add a new property.
                Office.DocumentProperty prp = prps.Add("HasCustomFooter", false, Office.MsoDocProperties.msoPropertyTypeBoolean, hasCustomFooter);
		    } catch (Exception ex)
            {
			    //PLLog.Error(ex, Globals.PROJECT_NAME)
			    MessageBox.Show("CustomFooterExists() Unable to add HasCustomFooter property" + ex.Message);
		    }
        }

        public static void CustomTableOfContentsExists(bool hasCustomTableOfContents)
        {
            Office.DocumentProperty prp = default(Office.DocumentProperty);
            Office.DocumentProperties prps = default(Office.DocumentProperties);

            try
            {
                prps = (Office.DocumentProperties)ExcelApplication.ActiveWorkbook.CustomDocumentProperties;
                // Add a new property.
                prp = prps.Add("HasCustomTableOfContents", false, Office.MsoDocProperties.msoPropertyTypeBoolean, hasCustomTableOfContents);
            }
            catch(Exception ex)
            {
                //PLLog.Error(ex, Globals.PROJECT_NAME)
                MessageBox.Show("CustomTableOfContentsExists() Unable to add HasCustomTableOfContents property" + ex.Message);
            }
        }

        public static void DeleteSheet(Worksheet ws, bool prompt = false)
        {
	        bool priorState = false;

	        priorState = ExcelApplication.DisplayAlerts;

	        if (prompt)
            {
		        ExcelApplication.DisplayAlerts = true;
		        ws.Delete();
	        }
            else
            {
		        ExcelApplication.DisplayAlerts = false;
		        ws.Delete();
	        }

	        ExcelApplication.DisplayAlerts = priorState;
        }

        public void DisplayExcelRange(Range rng)
        {
            //Debug.Print(rng.Address);
        }

        public static Worksheet DuplicateWorksheet(string sourceSheetName, string destinationSheetName, string beforeSheetName = "", string afterSheetName = "")
        {

	        Workbook wb = ExcelApplication.ActiveWorkbook;

	        foreach (Worksheet ws in wb.Worksheets)
            {
		        if (ws.Name == destinationSheetName)
                {
			        // TODO: Sheet exists.  Ask user what to do.
			        MessageBox.Show(string.Format("Destination Sheet: >{0}< already exists.", sourceSheetName));
			        return ws;
		        }
	        }

	        if (!string.IsNullOrEmpty(beforeSheetName))
            {
		        ((Worksheet)wb.Sheets[sourceSheetName]).Copy(Before: wb.Sheets[beforeSheetName]);
	        }
            else if (!string.IsNullOrEmpty(afterSheetName))
            {
		        ((Worksheet)wb.Sheets[sourceSheetName]).Copy(After: wb.Sheets[afterSheetName]);
	        }
            else
            {
		        ((Worksheet)wb.Sheets[sourceSheetName]).Copy();
	        }

	        ((Worksheet)wb.ActiveSheet).Name = destinationSheetName;

	        return (Worksheet)wb.ActiveSheet;
        }
        
        //    '----------------------------------------------------------------------
        //    '
        //    ' CreateHyperLinkName
        //    '
        //    ' Returns a name suitable for using as a hyperlink.
        //    '
        //    ' ToDo:
        //    '
        //    '----------------------------------------------------------------------

        //    Public Function CreateHyperLinkName( _
        //        ByVal strDataSheet As String, _
        //        ByVal strReportType As String _
        //    ) As String
        //        On Error GoTo PROC_ERROR
        //        Const cRoutineName = "CreateHyperLinkName"

        //        Dim strMachineType As String
        //        Dim strDate As String
        //        Dim intStart As Integer
        //        Dim intEnd As Integer

        //        intEnd = InStr(1, strDataSheet, ".", vbTextCompare)
        //        strMachineType = Left(strDataSheet, intEnd - 1)

        //        intStart = InStr(intEnd + 1, strDataSheet, ".", vbTextCompare) + 1
        //        intEnd = InStr(intStart + 1, strDataSheet, ".", vbTextCompare)

        //        ' This should handle "cmb1.sv.Aug1.txt" and "cmb1.sv.Aug"

        //        If intEnd = 0 Then
        //            strDate = Mid(strDataSheet, intStart, Len(strDataSheet))
        //        Else
        //            strDate = Mid(strDataSheet, intStart, intEnd - intStart)
        //        End If

        //        CreateHyperLinkName = _
        //            strMachineType & "." _
        //            & strReportType & "." _
        //            & strDate

        //PROC_EXIT:
        //        Exit Function

        //PROC_ERROR:
        //        Err.Raise(Err.Number, Err.Source, _
        //            cRoutineName & "():" & Err.Description, _
        //            Err.HelpFile, Err.HelpContext)
        //        GoTo PROC_EXIT
        //        Resume Next
        //    End Function
        

        //' TODO: Replease Add and link with this routine.

        //Sub LinkSheetIndirect( _
        //    ByVal strSheetName As String, _
        //    ByVal blnNewWorksheet As Boolean, _
        //    ByVal strIndexSheetName As String, _
        //    ByVal strLinkName As String, _
        //    ByVal xlOrientation As Excel.XlPageOrientation, _
        //    ByVal strLinkRowCell As String, _
        //    ByVal strLinkColCell As String, _
        //    ByVal strBeforeSheetName As String, _
        //    ByVal strAfterSheetName As String, _
        //    ByVal blnIncrementRow As Boolean, _
        //    ByVal blnIncrementCol As Boolean, _
        //    Optional ByVal sngColWidth As Single = 8.43 _
        //)
        //    Dim intLinkRow As Integer
        //    Dim intLinkCol As Integer

        //    If blnNewWorksheet Then
        //        NewWorksheet(strSheetName, strBeforeSheetName, strAfterSheetName)
        //    End If

        //    Worksheet_Format(ActiveSheet.Name, xlOrientation)
        //    Columns().ColumnWidth = sngColWidth

        //    Worksheets(strIndexSheetName).Activate()
        //    intLinkRow = Range(strLinkRowCell).value
        //    intLinkCol = Range(strLinkColCell).value
        //    Cells(intLinkRow, intLinkCol).value = strSheetName

        //    ' TODO: Fix the code here.
        //    Cells(intLinkRow, intLinkCol).Select()
        //    ActiveSheet.Hyperlinks.Add(Anchor:=Selection, Address:="", SubAddress:= _
        //        "'" & strSheetName & "'!A1", TextToDisplay:=strLinkName)

        //    If True = blnIncrementRow Then
        //        Range(strLinkRowCell).value = intLinkRow + 1
        //    End If

        //    If True = blnIncrementCol Then
        //        Range(strLinkColCell).value = intLinkCol + 1
        //    End If

        //    Worksheets(strSheetName).Activate()
        //End Sub ' LinkSheetIndirect

        //Sub NewWorksheet(strWorksheetName As String, _
        //    Optional strBeforeSheetName As String, _
        //    Optional strAfterSheetName As String _
        //)
        //    If strBeforeSheetName <> "" Then
        //        Application.ActiveWorkbook.Sheets.Add(Sheets(strBeforeSheetName))
        //    ElseIf strAfterSheetName <> "" Then
        //        Application.ActiveWorkbook.Sheets.Add, Sheets(strAfterSheetName)
        //    Else
        //        ActiveWorkbook.Sheets.Add()
        //    End If

        //    ActiveSheet.Name = strWorksheetName
        //End Sub


        //----------------------------------------------------------------------
        //
        // EmptyWorkbook
        //
        // Returns the name of a workbook containing one sheet.
        // Creates new workbook if blnCreateNew
        // All existing sheets except for strWorksheet are removed.
        //
        // ToDo:
        //
        //----------------------------------------------------------------------

        public static string EmptyWorkbook(string strWorksheetName, bool blnCreateNew)
        {
            //Worksheet shtWS = default(Worksheet);
	        string strKeepName = null;

	        if (true == blnCreateNew)
            {
		        ExcelApplication.Workbooks.Add();
		        // Keep the name so we don't have to worry about what name is given.
		        strKeepName = ((Worksheet)ExcelApplication.ActiveSheet).Name;
	        }
            else
            {
		        strKeepName = strWorksheetName;
	        }

	        ExcelApplication.DisplayAlerts = false;
	        // Yes, delete the damn things!

	        // Remove all other worksheets

	        foreach (Worksheet ws in ExcelApplication.ActiveWorkbook.Sheets)
            {
		        if (strKeepName != ws.Name)
                {
			        ws.Delete();
		        }
	        }

	        ExcelApplication.DisplayAlerts = true;

	        ((Worksheet)ExcelApplication.ActiveSheet).Name = strWorksheetName;

	        return strWorksheetName;
        }

        public static int FindFirst_PopulatedColumn_InRow(Range searchFromCell)
        {
            Worksheet searchWorksheet = searchFromCell.Worksheet;
            Range searchRange = default(Range);
            int columnNumber = searchFromCell.Column;
            int currentRow = searchFromCell.Row;
            int lastColSpecial = searchFromCell.SpecialCells(XlCellType.xlCellTypeLastCell).Column;

            try
            {            
                searchRange = searchWorksheet.Range[searchWorksheet.Cells[currentRow, 1], searchWorksheet.Cells[currentRow, lastColSpecial]];

                columnNumber = searchRange.Find("*", 
                After: searchWorksheet.Cells[currentRow, lastColSpecial], 
                SearchOrder: XlSearchOrder.xlByRows, 
                SearchDirection: XlSearchDirection.xlNext).Column;
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            return columnNumber;
        }

        public static int FindFirst_PopulatedRow_InColumn(Range searchFromCell)
        {
            Worksheet searchWorksheet = searchFromCell.Worksheet;
            Range currentColumnRange = default(Range);
            int rowNumber = 0;
            int currentColumn = searchFromCell.Column;
            int lastRowSpecial = searchFromCell.SpecialCells(XlCellType.xlCellTypeLastCell).Row;

            try
            {
                currentColumnRange = (Range)searchWorksheet.Columns.Item[searchFromCell.Column];
                rowNumber = currentColumnRange.Find("*", 
                    After: searchWorksheet.Cells[lastRowSpecial, currentColumn], 
                    SearchOrder: XlSearchOrder.xlByColumns, 
                    SearchDirection: XlSearchDirection.xlNext).Row;
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            return rowNumber;
        }

        public static int FindLast_PopulatedColumn_InRow(Range searchFromCell)
        {
            Worksheet searchWorksheet = searchFromCell.Worksheet;
            Range searchRange = default(Range);
            int columnNumber = searchFromCell.Column;
            int currentRow = searchFromCell.Row;
            int lastColSpecial = searchFromCell.SpecialCells(XlCellType.xlCellTypeLastCell).Column;

            try
            {
                searchRange = searchWorksheet.Range[searchWorksheet.Cells[currentRow, 1], searchWorksheet.Cells[currentRow, lastColSpecial]];

                columnNumber = searchRange.Find("*", 
                    After: searchWorksheet.Cells[currentRow, 1], 
                    SearchOrder: XlSearchOrder.xlByRows, 
                    SearchDirection: XlSearchDirection.xlPrevious).Column;
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            return columnNumber;
        }

        public static int FindLast_PopulatedRow_InColumn(Range searchFromCell)
        {
            Worksheet searchWorksheet = searchFromCell.Worksheet;
	        Range currentColumnRange = default(Range);
	        int lastRow = 0;
            int currentColumn = searchFromCell.Column;
            
		    try
            {
			    currentColumnRange = (Range)searchWorksheet.Columns.Item[searchFromCell.Column];
                lastRow = currentColumnRange.Find("*", 
                    After: searchWorksheet.Cells[1, currentColumn], 
                    SearchOrder: XlSearchOrder.xlByColumns, 
                    SearchDirection: XlSearchDirection.xlPrevious).Row;
		    }
            catch (Exception ex)
            {
			    MessageBox.Show(ex.ToString());
		    }

	        return lastRow;            
        }

        public static int FindNext_PopulatedColumn_InRow(Range searchFromCell)
        {
            int columnNumber = 0;
            Range nextMatch;

            try
            {
                nextMatch = searchFromCell.Find("*", SearchOrder: XlSearchOrder.xlByRows, SearchDirection: XlSearchDirection.xlNext);

                if (nextMatch.Row != searchFromCell.Row)
                {
                	return searchFromCell.Column;
                }

                columnNumber = nextMatch.Column;

                if(columnNumber < searchFromCell.Column)
                {
                    columnNumber = searchFromCell.Column;
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            return columnNumber;
        }

        public static int FindNext_PopulatedRow_InColumn(Range searchFromCell)
        {
            int  rowNumber = 0;
            Range nextMatch;

            try
            {
                nextMatch = searchFromCell.Find("*", SearchOrder: XlSearchOrder.xlByColumns, SearchDirection: XlSearchDirection.xlNext);

                if(nextMatch.Column != searchFromCell.Column)
                {
                    return searchFromCell.Row;
                }

                rowNumber = nextMatch.Row;

                if(rowNumber < searchFromCell.Row)
                {
                    rowNumber = searchFromCell.Row;
                }

            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            return rowNumber;
        }

        public static int FindPrevious_PopulatedColumn_InRow(Range searchFromCell)
        {
            int columnNumber = 0;
            Range nextMatch;

            try
            {
                nextMatch = searchFromCell.Find("*", SearchOrder: XlSearchOrder.xlByRows, SearchDirection: XlSearchDirection.xlPrevious);

                if(nextMatch.Row != searchFromCell.Row)
                {
                    return searchFromCell.Column;
                }

                columnNumber = nextMatch.Column;

                if (columnNumber > searchFromCell.Column)
                {
                	columnNumber = searchFromCell.Column;
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            return columnNumber;
        }

        public static int FindPrevious_PopulatedRow_InColumn(Range searchFromCell)
        {
            int  rowNumber = 0;
            Range nextMatch;

            try
            {
                nextMatch = searchFromCell.Find("*", SearchOrder: XlSearchOrder.xlByColumns, SearchDirection: XlSearchDirection.xlPrevious);

                if(nextMatch.Column != searchFromCell.Column)
                {
                    return searchFromCell.Row;
                }

                rowNumber = nextMatch.Row;

                if(rowNumber > searchFromCell.Row)
                {
                    rowNumber = searchFromCell.Row;
                }

            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            return rowNumber;
        }

        public static string GetBuiltInPropertyValue(Workbook workBook, string propName)
        {
	        string returnValue = "";

            //// This procedure returns the value of the built-in document
            //// property specified in the propName argument for the Office
            //// document object specified in the workBook argument.

            //DocumentProperty prpDocProp = default(DocumentProperty);
            //object varValue = null;

            //const long ERR_BADPROPERTY = 5;
            //const long ERR_BADDOCOBJ = 438;
            //const long ERR_BADCONTEXT = -2147467259;

            Office.DocumentProperties documentProperties = (Office.DocumentProperties)workBook.BuiltinDocumentProperties;

            if (documentProperties != null)
            {
                try
                {
                    Office.DocumentProperty property = documentProperties[propName];
                    returnValue = property.Value.ToString();
                }
                catch(System.Runtime.InteropServices.COMException)
                {
                    //Common.WriteToDebugWindow(string.Format("GetBuiltInPropertyValue({0}) COMException ErrorCode:({1})", propName, ex.ErrorCode));
                    returnValue = "Property Not Available Yet";
                }
                catch(NullReferenceException)
                {
                    //Common.WriteToDebugWindow(string.Format("GetBuiltInPropertyValue({0}) NullReferenceException", propName));
                    returnValue = "Property Not Set";
                }
                catch(Exception)
                {
                    //Common.WriteToDebugWindow(string.Format("GetBuiltInPropertyValue({0}) Exception >{1})<", propName, ex));
                }
            }

            return returnValue;
        }

        public static string GetFile(string initialFolder = "", string dialogTitle = "Open", string fileFilter = "All Files (*.*)|*.*")
        {
            using(OpenFileDialog ofd = new OpenFileDialog())
            {
                DialogResult result = default(System.Windows.Forms.DialogResult);
                ofd.Multiselect = false;
                ofd.InitialDirectory = initialFolder;
                ofd.Title = dialogTitle;
                ofd.Filter = fileFilter;
                result = ofd.ShowDialog();
                //Debug.WriteLine(ofd.FileName);
                return ofd.FileName;
            }
        }

        public static string GetOpenFileName(string initialFolder = "", string dialogTitle = "Open", string fileFilter = "All Files (*.*)|*.*")
        {
            using(OpenFileDialog ofd = new OpenFileDialog())
            {
                DialogResult result = default(System.Windows.Forms.DialogResult);
                ofd.Multiselect = false;
                ofd.InitialDirectory = initialFolder;
                ofd.Title = dialogTitle;
                ofd.Filter = fileFilter;
                result = ofd.ShowDialog();
                //Debug.WriteLine(ofd.FileName);
                return ofd.FileName;
            }
        }

        public static string GetSaveFileName(string initialFolder = "", string proposedFileName = "", string dialogTitle = "Open", string fileFilter = "All Files (*.*)|*.*")
        {
            using(SaveFileDialog sfd = new SaveFileDialog())
            {
                DialogResult result = default(DialogResult);
                sfd.InitialDirectory = initialFolder;
                sfd.Title = dialogTitle;
                sfd.Filter = fileFilter;
                sfd.FileName = proposedFileName;
                result = sfd.ShowDialog();
                //Debug.WriteLine(sfd.FileName);
                return sfd.FileName;
            }
        }

        public static bool HasCustomFooter()
        {
	        Office.DocumentProperty prp = default(Office.DocumentProperty);
	        Office.DocumentProperties prps = default(Office.DocumentProperties);

		    try
            {
			    prps = (Office.DocumentProperties)ExcelApplication.ActiveWorkbook.CustomDocumentProperties;
			    prp = prps["HasCustomFooter"];
			    // If the property exists we don't really care about the value
			    return true;

		    } 
            catch (Exception)
            {
			    // Exception is thrown if property does not exist
			    return false;
		    }
        }

        public static bool HasCustomTableOfContents()
        {
            try
            {
                Office.DocumentProperties prps = (Office.DocumentProperties)ExcelApplication.ActiveWorkbook.CustomDocumentProperties;
                Office.DocumentProperty prp = prps["HasCustomTableOfContents"];
                // If the property exists we don't really care about the value
                return true;
            }
            catch(Exception)
            {
                // Exception is thrown if property does not exist
                return false;
            }
        }

        /// <summary>
        /// Add a new worksheet.
        /// </summary>
        /// <param name="sheetName">Name of Worksheet.  beforeSheetName can use FIRST, afterSheetName can use LAST</param>
        /// <param name="beforeSheetName">Sheet name or FIRST</param>
        /// <param name="afterSheetName">Sheet name or LAST</param>
        /// <remarks>beforeSheetName can use FIRST, afterSheetName can use LAST</remarks>
        /// <returns></returns>
        public static Worksheet NewWorksheet(string sheetName, string beforeSheetName = "", string afterSheetName = "")
        {
	        Workbook wb = ExcelApplication.ActiveWorkbook;
            Worksheet newWs;

	        foreach (Worksheet ws in wb.Worksheets)
            {
		        if (ws.Name == sheetName)
                {
			        // Sheet exists.  Ask user what to do.
                    DialogResult answer = MessageBox.Show("Preserve Existing WorkSheet: >" + sheetName + "< ?", "Requested WorkSheet Already Exists", MessageBoxButtons.YesNo);

			        if (DialogResult.Yes == answer)
                    {
                        return ws;
                    }
                    else
                    {
                        ws.Delete();
                    }
		        }
	        }

	        if (!string.IsNullOrEmpty(beforeSheetName))
            {
                if ("FIRST" == beforeSheetName)
                {
                    newWs = (Worksheet)wb.Sheets.Add(Before: wb.Sheets[1]);
                }
                else
                {
                    newWs = (Worksheet)wb.Sheets.Add(Before: wb.Sheets[beforeSheetName]);           
                }
	        } 
            else if (!string.IsNullOrEmpty(afterSheetName))
            {
                if("LAST" == afterSheetName)
                {
                    newWs = (Worksheet)wb.Sheets.Add(After: wb.Sheets[wb.Sheets.Count]);
                }
                else
                {
                    newWs = (Worksheet)wb.Sheets.Add(After: wb.Sheets[afterSheetName]);
                }
	        } 
            else
            {
                newWs = (Worksheet)wb.Sheets.Add();
	        }

	        newWs.Name = sheetName;
            // Make the outline things show up the way CHR likes :)
            newWs.Outline.SummaryColumn = XlSummaryColumn.xlSummaryOnLeft;
            newWs.Outline.SummaryRow = XlSummaryRow.xlSummaryAbove;

	        return newWs;
        }

        //Sub TestGetSheetSize()
        //    Dim intHeight As Integer
        //    Dim intWidth As Integer

        //    GetSheetSize(intHeight, intWidth)
        //    MsgBox("Height = " & intHeight & "   Width = " & intWidth)
        //End Sub

        //Function FoundDataWorksheet() As Boolean
        //    Dim shtWS As Worksheet

        //    FoundDataWorksheet = False

        //    For Each shtWS In ActiveWorkbook.Worksheets
        //        shtWS.Activate()

        //        If "Data File Name" = Range("A1").value Then
        //            FoundDataWorksheet = True
        //            Exit For
        //        End If
        //    Next shtWS
        //End Function

        //    '----------------------------------------------------------------------
        //    '
        //    ' GetSheetSize
        //    '
        //    ' Returns printout sheet size in Points.
        //    '
        //    ' ToDo:
        //    '   Support more paper sizes.
        //    '----------------------------------------------------------------------

        //    Sub GetSheetSize(ByRef intHeight As Integer, ByRef intWidth As Integer)
        //        On Error GoTo PROC_ERROR
        //        Const cRoutineName = "GetSheetSize"

        //        Dim shtS As Excel.Worksheet
        //        Dim dblHeight As Double
        //        Dim dblWidth As Double
        //        Dim dblX As Double
        //        Dim intHeightMargin As Integer  ' In Points
        //        Dim intWidthMargin As Integer   ' In Points

        //        shtS = ActiveSheet

        //        ' Assume xlPortrait orientation.  Will swap below if needed.

        //        Select Case shtS.PageSetup.PaperSize
        //            Case xlPaperLetter
        //                dblHeight = 11
        //                dblWidth = 8.5

        //            Case xlPaperLegal
        //                dblHeight = 14
        //                dblWidth = 8.5

        //            Case Else
        //                MsgBox("Unsupported Paper Size")
        //                dblHeight = 11
        //                dblWidth = 8.5

        //        End Select

        //        If shtS.PageSetup.Orientation <> xlPortrait Then
        //            ' Swap dimensions if xlLandscape
        //            dblX = dblHeight
        //            dblHeight = dblWidth
        //            dblWidth = dblX
        //        End If

        //        ' Subtract Margins from total paper size.

        //        With shtS.PageSetup
        //            intHeightMargin = .TopMargin + .BottomMargin
        //            intWidthMargin = .LeftMargin + .RightMargin
        //        End With

        //        ' Return size in points.

        //        intHeight = Application.InchesToPoints(dblHeight) - intHeightMargin
        //        intWidth = Application.InchesToPoints(dblWidth) - intWidthMargin
        //        '
        //        '    If frmCharts.chkEnforcePrinterPageSize.Value = True Then
        //        '        intHeight = intHeight * 3
        //        '    End If

        //PROC_EXIT:
        //        Exit Sub

        //PROC_ERROR:
        //        Err.Raise(Err.Number, Err.Source, _
        //            cRoutineName & "():" & Err.Description, _
        //            Err.HelpFile, Err.HelpContext)
        //        GoTo PROC_EXIT
        //        Resume Next
        //    End Sub

        ///
        /// Protect or unprotect the sheet.  Return the current setting before
        /// and changes made.
        ///
        public static bool ProtectSheet(Worksheet workSheet, bool protectMode)
        {
	        bool currentMode = workSheet.ProtectContents;

	        if (protectMode == true) {
		        workSheet.Protect();
	        } else {
		        workSheet.Unprotect();
	        }

	        return currentMode;
        }

        public static string SafeName(string strS)
        {
            //string strSafe = null;
            //strSafe = strS.Replace("/", " ");
            ////strSafe = Strings.Replace(strS, "/", " ");
	        return strS.Replace("/", " ");
        }

        // Do this with regular expressions.

        public static string SafeSheetName(string workSheetName)
        {
            workSheetName = workSheetName.Replace("/", "-");
            workSheetName = workSheetName.Replace(@"\", "-");
            workSheetName = workSheetName.Replace("[", "");
            workSheetName = workSheetName.Replace("]", "");
            workSheetName = workSheetName.Replace(" ", "");

            return workSheetName.Substring(0, (workSheetName.Length > cMaxSheetNameLen ? cMaxSheetNameLen : workSheetName.Length));
        }

        public static void ScreenUpdatesOff()
        {
	        if (false == DisplayScreenUpdates)
            {
		        if (ExcelApplication.Workbooks.Count > 0)
                {
			        PriorScreenUpdatingState = ExcelApplication.ScreenUpdating;
			        ExcelApplication.ScreenUpdating = false;
		        } 
                else
                {
			        // Assume the intent is to run with screen updates on.
			        PriorScreenUpdatingState = true;
			        ExcelApplication.ScreenUpdating = false;
		        }
	        }
        }

        public static void ScreenUpdatesOn(bool force = false)
        {
            if (force)
            {
            	ExcelApplication.ScreenUpdating = true;
            }
            else
            {
	            ExcelApplication.ScreenUpdating = PriorScreenUpdatingState;                
            }

        }

        public static void SetCellValue(Range rngR, object vntValue, XlHAlign horizontalAlignment = XlHAlign.xlHAlignLeft, string strComment = "")
        {
            //var _with12 = rngR;
	        rngR.Value = vntValue;
	        rngR.HorizontalAlignment = horizontalAlignment;
	        if (!string.IsNullOrEmpty(strComment))
            {
		        rngR.AddComment();
		        rngR.Comment.Visible = false;
		        rngR.Comment.Text(Text: strComment);
	        }
        }


        //Public Sub TestScreenOff()
        //    Application.ScreenUpdating = False

        //    Application.Workbooks.Add()

        //    Application.ScreenUpdating = True
        //End Sub


        //Private Sub DumpPropertyCollection( _
        // ByVal prps As Office.DocumentProperties, _
        // ByVal rng As Excel.Range, ByRef i As Integer)
        //    Dim prp As Office.DocumentProperty

        //    For Each prp In prps
        //        rng.Offset(i, 0).Value = prp.Name
        //        Try
        //            If Not prp.Value Is Nothing Then
        //                rng.Offset(i, 1).Value = _
        //                 prp.Value.ToString
        //            End If
        //        Catch
        //            ' Do nothing at all.
        //        End Try
        //        i += 1
        //    Next
        //End Sub

        public static void ZapPageBreaks()
        {
	        foreach ( Worksheet ws in ExcelApplication.ActiveWorkbook.Sheets)
            {
		       ws.PageSetup.PrintArea = "";

                //Debug.Print(ws.Name);
		        //        Debug.Print sht.HPageBreaks.Count
		        //        Debug.Print sht.VPageBreaks.Count
		        // For some reason the page break handling is not clean.
		        // There are different types of page breaks, that is clear.
		        // Unfortunately the For Each hPB errors out if only Automatic
		        // Page breaks.  Wrap in try catch for AddIn
		         // ERROR: Not supported in C#: OnErrorStatement

		        if (ws.VPageBreaks.Count > 0)
                {
			        foreach (VPageBreak vPB in ws.VPageBreaks)
                    {
				        if (vPB.Type == XlPageBreak.xlPageBreakManual)
                        {
					        vPB.Delete();
				        }
			        }
		        }

		        if (ws.HPageBreaks.Count > 0)
                {
			        foreach (HPageBreak hPB in ws.HPageBreaks)
                    {
				        if (hPB.Type == XlPageBreak.xlPageBreakManual)
                        {
					        hPB.Delete();
				        }
			        }
		        }
	        }
        }

        #endregion
    }
}
