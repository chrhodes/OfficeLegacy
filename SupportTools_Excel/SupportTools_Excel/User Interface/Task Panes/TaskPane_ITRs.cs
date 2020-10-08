using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.ComponentModel;
using System.Reflection;
using System.Runtime;
using System.Text;
using System.Windows.Forms;

using Microsoft.Office.Core;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Excel;
using ExcelHlp = AddinHelper.Excel;

namespace SupportTools_Excel.User_Interface.Task_Panes
{

     public partial class TaskPane_ITRs : UserControl
    {
        //1
        //private string _inputFilePath = Common.cDEFAULT_FOLDER;

        //Private WithEvents _frmITRDetail As frmITRDetail

        //Public Property frmITRDetail As frmITRDetail
        //    Get
        //        If _frmITRDetail Is Nothing Then
        //            _frmITRDetail = New frmITRDetail()
        //        End If

        //        Return _frmITRDetail
        //    End Get
        //    Set(ByVal Value As frmITRDetail)
        //        _frmITRDetail = Value
        //    End Set
        //End Property

        //Private Sub MethodName() Handles _frmITRDetail.FormClosed
        //    _frmITRDetail = Nothing
        //End Sub

        #region "Initialization"

        public TaskPane_ITRs()
        {
            InitializeComponent();
        }

        #endregion

        #region "Event Handlers"

        private void btnAddListObjects_Click(System.Object sender, System.EventArgs e)
        {
	        AddListObjects();
        }

        private void btnAddPageBreaks_Click(System.Object sender, System.EventArgs e)
        {
	        AddPageBreaks();
        }

        private void btnAddPivotTables_Click(System.Object sender, System.EventArgs e)
        {
	        //Common.ExcelUtil.ScreenUpdatesOff()
	        AddPivotTables();
	        //Common.ExcelUtil.ScreenUpdatesOn()
        }

        private void btnDisplayITRDetail_Click(System.Object sender, System.EventArgs e)
        {
	        DisplayITRDetail();
        }

        private void btnGetITRInformation_Click(System.Object sender, System.EventArgs e)
        {
	        GetITRInformation();
        }

        private void btnMergeDuplicateRows_Click(System.Object sender, System.EventArgs e)
        {
	        MergeDuplicateRows();
        }

        private void btnProcessDARTOutput_Click(System.Object sender, System.EventArgs e)
        {
	        ProcessDynamicOutput();
	        DuplicateInputSheet();
	        ProcessDuplicates();
	        AddListObjects();
	        AddPivotTables();
	        //FormatSourceITRs()
        }

        private void btnProcessDuplicateRow_Click(System.Object sender, System.EventArgs e)
        {
	        ProcessDuplicates();
        }

        private void btnProcessDynamicOutput_Click(System.Object sender, System.EventArgs e)
        {
	        ProcessDynamicOutput();
        }


        private void cmbTeamName_SelectedIndexChanged(System.Object sender, System.EventArgs e)
        {
	        Common.TeamName = cmbTeamName.Text;
        }

        #endregion

        #region "Main Function Routines"

        public void AddAgeColumn()
        {
            Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;
            Range activeCell = Globals.ThisAddIn.Application.ActiveCell;
            Range formatRange;

	        activeSheet.Columns["D:D"].Insert(Shift: XlDirection.xlToRight, CopyOrigin: XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
	        activeSheet.Range["D5"].FormulaR1C1 = "Age";

	        int startRow = 0;
	        int endRow = 0;
	        int startColumn = 0;
	        int endColumn = 0;

	        startColumn = 4;
	        endColumn = startColumn;
	        startRow = 6;
            
	        endRow = activeCell.SpecialCells(XlCellType.xlCellTypeLastCell).Row;

            formatRange = activeSheet.Range[activeSheet.Cells[startRow, startColumn], activeSheet.Cells[endRow, endColumn]];

	        formatRange.FormulaR1C1 = "=TODAY()-RC[-1]";
	        formatRange.NumberFormat = "0";
        }

        private void AddPageBreaks()
        {
	        Range currentITR;
	        int currentRow = 0;
	        Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;

	        // TODO: Add a constant for this 
	        currentRow = Common.cFI_SecondITRRow;

	        currentITR = activeSheet.Cells[currentRow, 2];

	        while (currentITR.Value > 0)
            {
                //Debug.Print(currentITR.Value);

		        activeSheet.HPageBreaks.Add(activeSheet.Cells[currentRow, 15]);
		        activeSheet.VPageBreaks.Add(activeSheet.Cells[currentRow, 15]);

		        // Skip past four new rows
		        currentRow = currentRow + 5;
		        currentITR = activeSheet.Cells[currentRow, 2];
	        }
        }

        private static void DeleteRow(Worksheet ws, int i)
        {
	        //With ws.Rows(i).Interior
	        //    .Pattern = Constants.xlSolid
	        //    .PatternColorIndex = Constants.xlAutomatic
	        //    .Color = 5296274
	        //    .TintAndShade = 0
	        //    .PatternTintAndShade = 0
	        //End With

	        ws.Rows[i].Delete();
        }

        private string GetDefaultOutputFileName()
        {
	        return string.Format("{0}-ITRs {1}{2}", Common.TeamName, DateTime.Now.ToString("yyyy.MM.dd"), ".xlsx");
        }

        public delegate void DuplicateRowAction(Worksheet workSheet, int rowNumber);

        private static void AddListObjects()
        {
            Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;
            Microsoft.Office.Interop.Excel.Application excelApp = Globals.ThisAddIn.Application;

	        // Since we have deleted rows, save the workbook to ensure the .UsedRange gets reset.
	        // If we don't do this now, when we add the ListObjects using .SpecialCells
	        // we will get more rows that we need.  Alternatively could use the ExcelUtil.{LastRow,LastColumn) routines
	        // TODO: Make save silent

            excelApp.ActiveWorkbook.Save();

            Worksheet wsITRInfo = excelApp.Sheets["ITRInfo"];
            Worksheet wsITRInfoWithResources = excelApp.Sheets["ITRInfoWithResources"];

	        wsITRInfo.ListObjects.Add(
                SourceType: XlListObjectSourceType.xlSrcRange, 
                Source: excelApp.Range[wsITRInfo.Range[Common.cITRHeader_Cell], wsITRInfo.Range[Common.cITRHeader_Cell].SpecialCells(XlCellType.xlCellTypeLastCell)],
                XlListObjectHasHeaders: XlYesNoGuess.xlYes).Name = "SourceITRInfo";

	        wsITRInfoWithResources.ListObjects.Add(
                SourceType: XlListObjectSourceType.xlSrcRange, 
                Source: excelApp.Range[wsITRInfoWithResources.Range[Common.cITRHeader_Cell], wsITRInfoWithResources.Range[Common.cITRHeader_Cell].SpecialCells(XlCellType.xlCellTypeLastCell)],
                XlListObjectHasHeaders: XlYesNoGuess.xlYes).Name = "SourceITRResourceInfo";
        }

        //Private Sub GetITR(ByVal itrRow As Range)
        //    itrRow.Cells[1, 2].Value = "Subject"
        //    itrRow.Cells[1, 3].Value = "Application"
        //End Sub

        private void DisplayITRDetail()
        {
	        try {
		        int currentRow = Globals.ThisAddIn.Application.ActiveCell.Row;
		        int itrID = Globals.ThisAddIn.Application.ActiveSheet.Cells[currentRow, 1].Value;

		        // If GetITRInformation has not populated the Common.ApplicationDS.sp_ITRDetail table 
		        // this won't return anything.

		        //itrDetail = Common.ApplicationDS.sp_ITRDetail.Select(String.Format("ID = {0}", itrID))

		        // So, to be safe just use the table adapter to call the stored procedure.

		        Data.ApplicationDSTableAdapters.sp_ITRDetailTableAdapter ta = new Data.ApplicationDSTableAdapters.sp_ITRDetailTableAdapter();
		        Data.ApplicationDS.sp_ITRDetailRow itrDetail;
		        Data.ApplicationDS.sp_ITRDetailDataTable itrDataTable;
		        itrDataTable = ta.GetData(itrID.ToString());
		        itrDetail = (Data.ApplicationDS.sp_ITRDetailRow)itrDataTable.Rows[0];

		        if ((itrDetail != null))
                {
                    User_Interface.Forms.frmITRDetail frmITRDetail = new User_Interface.Forms.frmITRDetail();

			        foreach (Control ctrl in frmITRDetail.Controls)
                    {
				        try
                        {
					        if ((ctrl != null))
                            {
						        if (ctrl is System.Windows.Forms.TextBox)
                                {
                                    // TODO (crhodes): This doesn't look good.
							        dynamic fieldName = ctrl.Name.Replace("txt", "");
                                    ctrl.Text = itrDetail.ItemArray[fieldName].ToString();
                                    //ctrl.Text = itrDetail.Item(fieldName).ToString();
						        }
					        }
                            else
                            {
						        MessageBox.Show("ctrl is nothing??");
					        }
				        }
                        catch (System.ArgumentException ex)
                        {
                            MessageBox.Show(ex.ToString());
				        // Column likely does not exist on row/table
				        }
                        catch (Exception ex) 
                        {
					        MessageBox.Show(string.Format("Exception: {0}.{1}() - {2}", System.Reflection.Assembly.GetExecutingAssembly().FullName, System.Reflection.MethodInfo.GetCurrentMethod().Name, ex.ToString()));
				        }

			        }

			        //Dim date1 As Date = Today()
			        //Dim date2 As Date = Convert.ToDateTime(itrDetail.RequestID).Date
			        //Dim age As TimeSpan = date1.Subtract(date2)

			        //frmITRDetail.txtAge.Text = age.Days.ToString()

			        frmITRDetail.txtAge.Text = DateTime.Today.Subtract(System.DateTime.Parse(itrDetail.RequestID)).Days.ToString();

			        //frmDetail.txtID.Text = itrDetail.ID

			        //frmDetail.txtApplication.Text = itrDetail.Application
			        //frmDetail.txtDescription.Text = itrDetail.Description
			        //frmDetail.txtComments.Text = itrDetail.Comments
			        //frmDetail.txtJustification.Text = itrDetail.Justification
			        frmITRDetail.Show();
			        frmITRDetail.txtID.Focus();
			        frmITRDetail.txtID.SelectionLength = 0;
		        }
                else
                {
			        MessageBox.Show("Could not find ITR");
		        }
	        }
            catch (Exception ex)
            {
		        MessageBox.Show(string.Format("Exception: {0}.{1}() - {2}", System.Reflection.Assembly.GetExecutingAssembly().FullName, System.Reflection.MethodInfo.GetCurrentMethod().Name, ex.ToString()));
	        }
        }

        private static void DuplicateInputSheet()
        {
            Workbook activeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;

	        var _with5 = Globals.ThisAddIn;
            ExcelHlp.DuplicateWorksheet("SourceITRs", "ITRInfo", "SourceITRs");
	        activeWorkbook.Sheets["ITRInfo"].Columns[Common.cITRInfo_CommentColumns].Delete();

            ExcelHlp.DuplicateWorksheet("SourceITRs", "ITRInfoWithResources", "ITRInfo");
	        activeWorkbook.Sheets["ITRInfoWithResources"].Columns[Common.cITRITRInfoWithResources_CommentColumns].Delete();

            ExcelHlp.DuplicateWorksheet("SourceITRs", "FormatedITRs", "SourceITRs");
        }

        private void btnDuplicateInputSheet_Click(System.Object sender, System.EventArgs e)
        {
	        DuplicateInputSheet();
        }

        private void btnFormatSourceITRs_Click(System.Object sender, System.EventArgs e)
        {
	        FormatSourceITRs();
        }

        private void FormatSourceITRs()
        {
	        //Common.ExcelUtil.ScreenUpdatesOff()
	        FormatITRs();
	        //Common.ExcelUtil.ScreenUpdatesOn()
        }

        private void GetITRInformation()
        {
	        Range currentCell = Globals.ThisAddIn.Application.ActiveCell;
	        int firstRow = currentCell.Row;
	        //Dim lastRow As Integer = currentCell.SpecialCells(Constants.xlLastCell).Row
            int lastRow = ExcelHlp.FindLast_PopulatedRow_InColumn(currentCell);
	        //Dim foundRows As Hashtable = New Hashtable()
	        Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;
	        Range itrRow = default(Range);

	        //Debug.Print(String.Format("{0} - {1}", firstRow, lastRow))

	        //Dim ta As ApplicationDSTableAdapters.sp_ITRDetailTableAdapter = New ApplicationDSTableAdapters.sp_ITRDetailTableAdapter()
	        Data.ApplicationDS.sp_ITRDetailRow itrDataRow = null;
	        //Dim itrDataTable As ApplicationDS.sp_ITRDetailDataTable

	        //Dim tr As ApplicationDS.sp_ITRDetailRow

	        for (int i = firstRow; i <= lastRow; i++)
            {
		        itrRow = activeSheet.Cells[i, currentCell.Column];

		        Common.ApplicationDS.sp_ITRDetail.ImportRow(itrDataRow);

		        itrRow.Cells[1, 2].Value = itrDataRow.Subject;
		        itrRow.Cells[1, 3].Value = itrDataRow.Application.Replace("/Life IT/Technical Services/", "");

		        try
                {
			        itrRow.Cells[1, 4].Value = itrDataRow.Description.Length;
			        // Current Condition
		        }
                catch (Exception)
                {
			        itrRow.Cells[1, 4].Value = "<none>";
		        }
		        try
                {
			        itrRow.Cells[1, 5].Value = itrDataRow.Justification.Length;
			        // Desired Outcome
		        }
                catch (Exception)
                {
			        itrRow.Cells[1, 5].Value = "<none>";
		        }
		        try
                {
			        itrRow.Cells[1, 6].Value = itrDataRow.Comments.Length;
			        // Comments
		        }
                catch (Exception)
                {
			        itrRow.Cells[1, 6].Value = "<none>";
		        }
	        }
        }


        public void HilightDuplicateRows(string worksheetName)
        {
	        Worksheet ws = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[worksheetName];

	        ws.Columns[Common.cRESOURCEID_COLUMN].Delete();

	        long previousITR = 0;
	        long currentITR = 0;
	        int lastITRrow = 0;
	        bool foundMatching = false;

	        lastITRrow = ws.Range["B5"].SpecialCells(XlCellType.xlCellTypeLastCell).Row;
	        // and then walk the list of ITRs and delete duplicates
            //Debug.Print(lastITRrow);

	        for (int i = lastITRrow; i >= Common.cFirstITRRow; i += -1)
            {
                //Debug.Print(string.Format("previousITR {0}  currentITR {1}   i {2}  lastITRRow {3}", previousITR, ws.Cells[i, Common.cITRID_COLUMN].Value, i, lastITRrow));
		        currentITR = Convert.ToInt32(ws.Cells[i, Common.cITRID_COLUMN].Value);

		        if (currentITR == previousITR)
                {
                    //Debug.Print("Matching ITR");
			        HilightRow(ws, i);
			        lastITRrow = lastITRrow - 1;

			        if (!foundMatching)
                    {
				        // Hilight the previous row which matched the current row.
				        HilightRow(ws, i - 1);
			        }

			        foundMatching = true;
		        }
                else 
                {
                    //Debug.Print("New ITR");
			        foundMatching = false;
		        }

		        previousITR = currentITR;
	        }
        }

        private static void HilightRow2(Worksheet ws, int i)
        {
            Range row = ws.Rows[i];
            
	        row.Interior.Pattern = Constants.xlGray50;
	        row.Interior.PatternColorIndex = Constants.xlAutomatic;
	        row.Interior.Color = 5296274;
	        row.Interior.TintAndShade = 0;
	        row.Interior.PatternTintAndShade = 0;

	        //ws.Rows[i].Delete()
        }

        private static void HilightRow(Worksheet ws, int i)
        {
            Range row = ws.Rows[i];

	        row.Interior.Pattern = Constants.xlSolid;
	        row.Interior.PatternColorIndex = Constants.xlAutomatic;
	        row.Interior.Color = 5296274;
	        row.Interior.TintAndShade = 0;
	        row.Interior.PatternTintAndShade = 0;

	        //ws.Rows[i].Delete()
        }

        private bool IsValidITR(string selectedText)
        {
	        if (selectedText.Trim().Length != 5)
            {
		        return false;
	        }

	        return System.Text.RegularExpressions.Regex.Match(selectedText, "[0-9]{5}").Success;
        }

        private void ProcessDuplicateRows(string worksheetName, DuplicateRowAction deleteAction, DuplicateRowAction hilightAction)
        {
	        Worksheet ws = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[worksheetName];

	        long previousITR = 0;
	        long currentITR = 0;
	        int lastITRrow = ws.Range["B5"].SpecialCells(XlCellType.xlCellTypeLastCell).Row;

	        // and then walk the list of ITRs and delete duplicates
            //Debug.Print(lastITRrow);

	        for (int i = lastITRrow; i >= Common.cFirstITRRow; i += -1)
            {
		        //Debug.Print(String.Format("previousITR {0}  currentITR {1}   i {2}  lastITRRow {3}", _
		        //    previousITR, ws.Cells[i, Common.cITRID_COLUMN].Value, i, lastITRrow))
		        currentITR = Convert.ToInt64(ws.Cells[i, Common.cITRID_COLUMN].Value);

		        if (currentITR == previousITR)
                {
			        //Debug.Print("Matching ITR")
			        // First hilight the previous matching ITR
			        hilightAction(ws, i + 1);
			        // Then take the deleteAction on the current row.  This might be a hilight operation.
			        deleteAction(ws, i);
			        lastITRrow = lastITRrow - 1;
		        }
                else
                {
			        //Debug.Print("New ITR")
		        }

		        previousITR = currentITR;
	        }
        }

        private void ProcessDuplicates()
        {
	        //Common.ExcelUtil.ScreenUpdatesOff()
	        // Delete the Resource ID column from the ITRInfo sheet
	        Globals.ThisAddIn.Application.Sheets["ITRInfo"].Columns(Common.cRESOURCEID_COLUMN).Delete();
	        ProcessDuplicateRows("ITRInfo", DeleteRow, HilightRow);

	        // Just Hilight the ITRs with multiple Resources.  We keep the Resource ID column here.
	        ProcessDuplicateRows("ITRInfoWithResources", HilightRow2, HilightRow);
	        //Common.ExcelUtil.ScreenUpdatesOn()
        }

        private void ProcessDynamicOutput()
        {
	        //TODO: Drive this off Config file and/or TaskPane_Config
	        //Common.ExcelUtil.ScreenUpdatesOff()

	        Globals.ThisAddIn.Application.Sheets["dynamicRepo"].Name = "SourceITRs";
	        // Move down a row so the TOC doesn't step on things
	        Globals.ThisAddIn.Application.ActiveSheet.Rows("1:1").Insert(Shift: XlInsertShiftDirection.xlShiftDown, CopyOrigin: XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
	        AddAgeColumn();
	        UpdateApplicationNames();
	        Globals.ThisAddIn.Application.ActiveSheet.Range("A1").Select();
	        SaveOutputFile();

	        //Common.ExcelUtil.ScreenUpdatesOn()
        }

        public void RemoveDuplicateRows(string worksheetName, DuplicateRowAction action)
        {
	        Worksheet ws = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[worksheetName];

	        ws.Columns[Common.cRESOURCEID_COLUMN].Delete();

	        long previousITR = 0;
	        long currentITR = 0;
	        int lastITRrow = 0;
	        bool foundMatching = false;

	        lastITRrow = ws.Range["B5"].SpecialCells(XlCellType.xlCellTypeLastCell).Row;
	        // and then walk the list of ITRs and delete duplicates
            //Debug.Print(lastITRrow);

	        for (int i = lastITRrow; i >= Common.cFirstITRRow; i += -1)
            {
                //Debug.Print(string.Format("previousITR {0}  currentITR {1}   i {2}  lastITRRow {3}", previousITR, ws.Cells[i, Common.cITRID_COLUMN].Value, i, lastITRrow));
		        currentITR = Convert.ToInt64(ws.Cells[i, Common.cITRID_COLUMN].Value);

		        if (currentITR == previousITR)
                {
                    //Debug.Print("Matching ITR");
			        DeleteRow(ws, i);
			        lastITRrow = lastITRrow - 1;

			        if (!foundMatching)
                    {
				        // Hilight the previous row which matched the current row.
				        HilightRow(ws, i - 1);
			        }

			        foundMatching = true;
		        }
                else
                {
                    //Debug.Print("New ITR");
			        foundMatching = false;
		        }

		        previousITR = currentITR;
	        }
        }

        private void SaveOutputFile()
        {
	        string outputFileName = GetDefaultOutputFileName();

	        var _with8 = Globals.ThisAddIn.Application;
	        outputFileName = ExcelHlp.GetSaveFileName(Common.cDEFAULT_FOLDER, outputFileName, "Enter Save As Name");

	        if (outputFileName.Length > 0) {
		        _with8.ActiveWorkbook.SaveAs(Filename: outputFileName, FileFormat: XlFileFormat.xlOpenXMLWorkbook, CreateBackup: false);
	        }
        }

        public void FormatITRs()
        {
	        Globals.ThisAddIn.Application.ActiveWorkbook.Sheets["FormatedITRs"].Activate();

	        var _with9 = Globals.ThisAddIn.Application;
	        //.Rows("2:3").Select()
	        //.Selection.Delete(Shift:=XlDirection.xlUp)
	        _with9.Rows["5:5"].Select();

	        FormatPrinterOutput();

	        MergeAssignedResources("FormatedITRs");

	        MoveITRCommentsBelow();

	        FormatColumns_FormatedITRs();

	        AddPageBreaks();

        }


        public void FormatColumns_FormatedITRs()
        {
	        Microsoft.Office.Interop.Excel.Application excelApp = Globals.ThisAddIn.Application;

	        excelApp.Columns[Common.cFI_Application_Column_Range].ColumnWidth = 25;
	        excelApp.Columns[Common.cFI_ITRID_Column_Range].ColumnWidth = 5;
	        excelApp.Columns[Common.cFI_EnteredOn_Column_Range].ColumnWidth = 9;
	        excelApp.Columns[Common.cFI_Age_Column_Range].ColumnWidth = 5;
	        excelApp.Columns[Common.cFI_EnteredBy_Column_Range].ColumnWidth = 9;
	        excelApp.Columns[Common.cFI_RequestedBy_Column_Range].ColumnWidth = 8;
	        excelApp.Columns[Common.cFI_ReleaseNbr_Column_Range].ColumnWidth = 7;
	        excelApp.Columns[Common.cFI_PatRank_Column_Range].ColumnWidth = 4;
	        excelApp.Columns[Common.cFI_Category_Column_Range].ColumnWidth = 12;
	        excelApp.Columns[Common.cFI_Status_Column_Range].ColumnWidth = 10;
	        excelApp.Columns[Common.cFI_Severity_Column_Range].ColumnWidth = 10;
	        excelApp.Columns[Common.cFI_LOE_Column_Range].ColumnWidth = 10;
	        excelApp.Columns[Common.cFI_Subject_Column_Range].ColumnWidth = 35;
	        excelApp.Columns[Common.cFI_Resource_Column_Range].Columnwidth = 10;
        }

        public void FormatColumns_PivotTable(string workSheetName)
        {
	        // TODO: Vary this by sheet.  Some may want to be landscape, some portrait.  Perhaps pass parameter.
	        Microsoft.Office.Interop.Excel.Application excelApp = Globals.ThisAddIn.Application;

	        excelApp.Sheets[workSheetName].Activate();
	        excelApp.Columns[Common.cPT_ITR_Column_Range].ColumnWidth = 75;
	        excelApp.Columns[Common.cPT_Count_Column_Range].ColumnWidth = 6;
        }

        public void FormatPrinterOutput()
        {
            Globals.ThisAddIn.Application.PrintCommunication = false;

            Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;
            PageSetup pageSetup = activeSheet.PageSetup;
            
	        pageSetup.PrintTitleRows = "$5:$5";
            pageSetup.PrintTitleColumns = "";
	        pageSetup.PrintArea = "";

	        // TODO: Take Center Header as parameter

	        pageSetup.LeftHeader = "";
	        pageSetup.CenterHeader = "Open Delivery Services ITRs";
	        pageSetup.RightHeader = "";
	        pageSetup.LeftFooter = "";
	        pageSetup.CenterFooter = "";
	        pageSetup.RightFooter = "";
	        pageSetup.LeftMargin = Globals.ThisAddIn.Application.InchesToPoints(0.25);
	        pageSetup.RightMargin = Globals.ThisAddIn.Application.InchesToPoints(0.25);
	        pageSetup.TopMargin = Globals.ThisAddIn.Application.InchesToPoints(0.75);
	        pageSetup.BottomMargin = Globals.ThisAddIn.Application.InchesToPoints(0.75);
	        pageSetup.HeaderMargin = Globals.ThisAddIn.Application.InchesToPoints(0.3);
	        pageSetup.FooterMargin = Globals.ThisAddIn.Application.InchesToPoints(0.3);
	        pageSetup.PrintHeadings = false;
	        pageSetup.PrintGridlines = false;
	        pageSetup.PrintComments = XlPrintLocation.xlPrintNoComments;
	        pageSetup.PrintQuality = 600;
	        pageSetup.CenterHorizontally = false;
	        pageSetup.CenterVertically = false;
	        pageSetup.Orientation = XlPageOrientation.xlLandscape;
	        pageSetup.Draft = false;
	        pageSetup.PaperSize = XlPaperSize.xlPaperLegal;
	        pageSetup.FirstPageNumber = (int)Constants.xlAutomatic;
	        pageSetup.Order = XlOrder.xlDownThenOver;
	        pageSetup.BlackAndWhite = false;
	        pageSetup.Zoom = 100;
	        pageSetup.PrintErrors = XlPrintErrors.xlPrintErrorsDisplayed;
	        pageSetup.OddAndEvenPagesHeaderFooter = false;
	        pageSetup.DifferentFirstPageHeaderFooter = false;
	        pageSetup.ScaleWithDocHeaderFooter = true;
	        pageSetup.AlignMarginsHeaderFooter = true;
	        pageSetup.EvenPage.LeftHeader.Text = "";
	        pageSetup.EvenPage.CenterHeader.Text = "";
	        pageSetup.EvenPage.RightHeader.Text = "";
	        pageSetup.EvenPage.LeftFooter.Text = "";
	        pageSetup.EvenPage.CenterFooter.Text = "";
	        pageSetup.EvenPage.RightFooter.Text = "";
	        pageSetup.FirstPage.LeftHeader.Text = "";
	        pageSetup.FirstPage.CenterHeader.Text = "";
	        pageSetup.FirstPage.RightHeader.Text = "";
	        pageSetup.FirstPage.LeftFooter.Text = "";
	        pageSetup.FirstPage.CenterFooter.Text = "";
	        pageSetup.FirstPage.RightFooter.Text = "";

            Globals.ThisAddIn.Application.PrintCommunication = true;
        }

        private void MergeDuplicateRows()
        {
	        Range currentCell = Globals.ThisAddIn.Application.ActiveCell;
	        int firstRow = currentCell.Row;
	        //Dim lastRow As Integer = currentCell.SpecialCells(Constants.xlLastCell).Row
            int lastRow = ExcelHlp.FindLast_PopulatedRow_InColumn(currentCell);
	        Dictionary<string, Range> foundITRs = new Dictionary<string, Range>();
	        Dictionary<string, string> duplicateITRs = new Dictionary<string, string>();
	        Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;
	        List<string> rowsToDelete = new List<string>();
	        Range rng = default(Range);

	        Debug.Print(string.Format("{0} - {1}", firstRow, lastRow));

	        for (int i = lastRow; i >= firstRow; i += -1)
            {
		        rng = activeSheet.Cells[i, currentCell.Column];

                //Debug.Print(string.Format("{0} - {1} - {2}", i, rng.Value, rng.Cells[1, 7].Value));

		        if (foundITRs.ContainsKey(rng.Value))
                {
			        if (duplicateITRs.ContainsKey(rng.Value))
                    {
				        // This is not the first duplicate.  Add resource if not there already.
				        if (!duplicateITRs[rng.Value].Contains(rng.Cells[1, 7].Value))
                        {
					        duplicateITRs[rng.Value] = duplicateITRs[rng.Value] + ", " + rng.Cells[1, 7].Value;
				        }

			        } 
                    else 
                    {
				        // Add the Resource from the duplicate row
				        duplicateITRs[rng.Value] = rng.Cells[1, 7].Value;
			        }

			        rowsToDelete.Add(string.Format("{0}:{1}", i, i));
		        }
                else 
                {
			        foundITRs[rng.Value] = rng;
		        }
	        }

            //foreach (KeyValuePair<string, Range> row in foundITRs)
            //{
            //    //Debug.Print(string.Format("{0} - {1} - {2}", row.Value.Row, row.Value.Cells[1, 1].Value, row.Value.Cells[1, 7].Value));
            //}

	        foreach (KeyValuePair<string, string> row in duplicateITRs) 
            {
                //Debug.Print(string.Format("{0} - {1}", row.Key, row.Value));

		        string resources = activeSheet.Cells[foundITRs[row.Key].Row, 7].Value;

		        if (row.Value.Contains(resources))
                {
			        activeSheet.Cells[foundITRs[row.Key].Row, 7].Value = row.Value;
		        }
                else 
                {
			        activeSheet.Cells[foundITRs[row.Key].Row, 7].Value = resources + ", " + row.Value;
		        }

	        }

	        foreach (string row in rowsToDelete)
            {
                //Debug.Print(row);
		        activeSheet.Rows[row].Delete();
	        }
        }

        private void MergeResourceAndDeleteRow(Worksheet ws, int i)
        {
	        string allResources = ws.Cells[i + 1, 14].Value;
	        string resource = ws.Cells[i, 14].Value;

	        allResources = string.Format("{0}, {1}", allResources, resource);
	        ws.Cells[i + 1, 14].Value = allResources;
	        ws.Rows[i].Delete();
        }

        private void MergeAssignedResources(string worksheetName)
        {
	        Worksheet ws = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[worksheetName];
	        // Walk the list of ITRs looking for duplicates and merge the assigned resources if any.
	        long previousITR = 0;
	        long currentITR = 0;
	        int lastITRrow = ws.Range["B5"].SpecialCells(XlCellType.xlCellTypeLastCell).Row;

	        // and then walk the list of ITRs and delete duplicates
            //Debug.Print(lastITRrow);

	        for (int i = lastITRrow; i >= Common.cFirstITRRow; i += -1)
            {
		        //Debug.Print(String.Format("previousITR {0}  currentITR {1}   i {2}  lastITRRow {3}", _
		        //    previousITR, ws.Cells[i, Common.cITRID_COLUMN].Value, i, lastITRrow))
		        currentITR = Convert.ToInt32(ws.Cells[i, Common.cITRID_COLUMN].Value);

		        if (currentITR == previousITR)
                {
			        //Debug.Print("Matching ITR")
			        // First hilight the previous matching ITR
			        MergeResourceAndDeleteRow(ws, i);
			        lastITRrow = lastITRrow - 1;
		        } 
                else 
                {
			        //Debug.Print("New ITR")
		        }

		        previousITR = currentITR;
	        }

        }

        public void MoveCommentsBelowAndGroup()
        {
            Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;
            Range currentCell = Globals.ThisAddIn.Application.ActiveCell;
            Range currentRow = Globals.ThisAddIn.Application.Rows[currentCell.Row];

	        // Make room for comments

	        currentRow.Insert(CopyOrigin: XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
	        // Description of Current Condition
	        currentRow.Insert(CopyOrigin: XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
	        // Description of Desired Outcome
	        currentRow.Insert(CopyOrigin: XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
	        // Prioritization Comments
	        currentRow.Insert(CopyOrigin: XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
	        // Comments

	        // Add Row Headers

	        currentCell.Offset[-4, 0].Value = "Description of Current Condition";
	        currentCell.Offset[-3, 0].Value = "Description of Desired Outcome";
	        currentCell.Offset[-2, 0].Value = "Prioritization Comments";
	        currentCell.Offset[-1, 0].Value = "Comments";

	        // Merge all cells for each row so we have a long field for values.  This makes cell wrapping work better

	        currentCell.Offset[-4, 1].Range[activeSheet.Cells[1, 1], activeSheet.Cells[1, 13]].Merge();
	        currentCell.Offset[-3, 1].Range[activeSheet.Cells[1, 1], activeSheet.Cells[1, 13]].Merge();
	        currentCell.Offset[-2, 1].Range[activeSheet.Cells[1, 1], activeSheet.Cells[1, 13]].Merge();
	        currentCell.Offset[-1, 1].Range[activeSheet.Cells[1, 1], activeSheet.Cells[1, 13]].Merge();

	        // Copy the values from the previous row into the new locations.

	        currentCell.Offset[-4, 1].Value = currentCell.Offset[-5, 14].Value;
	        currentCell.Offset[-3, 1].Value = currentCell.Offset[-5, 15].Value;
	        currentCell.Offset[-2, 1].Value = currentCell.Offset[-5, 16].Value;
	        currentCell.Offset[-1, 1].Value = currentCell.Offset[-5, 17].Value;

	        // Update the row height  based on the length of the comments

	        SetRowHeight(currentCell.Offset[-4, 1]);
	        SetRowHeight(currentCell.Offset[-3, 1]);
	        SetRowHeight(currentCell.Offset[-2, 1]);
	        SetRowHeight(currentCell.Offset[-1, 1]);

	        // Group the new comment rows so they can be collapsed.

	        currentCell.Offset[-4, 0].Range[activeSheet.Cells[1, 1], activeSheet.Cells[4, 1]].Rows.Group();

	        // TODO: Probably add the Page Breaks here.  Maybe optional config flag to drive
        }

        public void SetRowHeight(Range commentCell)
        {
	        float rowHeight = 0;
	        int textLength = 0;

            textLength = commentCell.Value.Length();

	        //    If textLength > 160 Then
	        rowHeight = (float)(((textLength / 225.0) * 15.75) + 15.75);
	        //    Else
	        //        rowHeight = 15.75
	        //    End If

	        // Max

	        if (rowHeight > 409)
            {
		        rowHeight = 409;
	        }

            //Debug.Print(textLength, rowHeight);
	        commentCell.Rows.RowHeight = rowHeight;
        }

        public void UpdateApplicationNames()
        {
	        var _with16 = Globals.ThisAddIn.Application;
	        _with16.Columns["A:A"].Select();

	        // Strip out "/Life IT/Technical Services"

	        _with16.Selection.Replace(What: "/Life IT/Technical Services/", Replacement: "", LookAt: XlLookAt.xlPart, SearchOrder: XlSearchOrder.xlByRows, MatchCase: false, SearchFormat: false, ReplaceFormat: false);

	        // Shorten "Application Infrastructure"

	        _with16.Selection.Replace(What: "Application Infrastructure/", Replacement: "AI/", LookAt: XlLookAt.xlPart, SearchOrder: XlSearchOrder.xlByRows, MatchCase: false, SearchFormat: false, ReplaceFormat: false);

	        // Shorten "Common Architecture Components"

	        _with16.Selection.Replace(What: "Common Architecture Components", Replacement: "Common Arch", LookAt: XlLookAt.xlPart, SearchOrder: XlSearchOrder.xlByRows, MatchCase: false, SearchFormat: false, ReplaceFormat: false);

	        // Shorten "Development Support Services"

	        _with16.Selection.Replace(What: "Development Support Services/", Replacement: "DSS/", LookAt: XlLookAt.xlPart, SearchOrder: XlSearchOrder.xlByRows, MatchCase: false, SearchFormat: false, ReplaceFormat: false);

	        // Shorten "EAC Console Apps"

	        _with16.Selection.Replace(What: "EAC Console Apps/", Replacement: "EAC/", LookAt: XlLookAt.xlPart, SearchOrder: XlSearchOrder.xlByRows, MatchCase: false, SearchFormat: false, ReplaceFormat: false);

	        // Shorten "Integration Services"

	        _with16.Selection.Replace(What: "Integration Services/", Replacement: "IS/", LookAt: XlLookAt.xlPart, SearchOrder: XlSearchOrder.xlByRows, MatchCase: false, SearchFormat: false, ReplaceFormat: false);

	        // Shorten "Biztalk Solution Development"

	        _with16.Selection.Replace(What: "Biztalk Solution Development", Replacement: "BizTalk", LookAt: XlLookAt.xlPart, SearchOrder: XlSearchOrder.xlByRows, MatchCase: false, SearchFormat: false, ReplaceFormat: false);

	        // Shorten "TrackSuite"

	        _with16.Selection.Replace(What: "TrackSuite/", Replacement: "TS/", LookAt: XlLookAt.xlPart, SearchOrder: XlSearchOrder.xlByRows, MatchCase: false, SearchFormat: false, ReplaceFormat: false);

	        // Shorten "Multi Case Reporting"

	        _with16.Selection.Replace(What: "Multi Case Reporting", Replacement: "MCR", LookAt: XlLookAt.xlPart, SearchOrder: XlSearchOrder.xlByRows, MatchCase: false, SearchFormat: false, ReplaceFormat: false);
        }

        #endregion

        #region "Pivot Table Routines"

        public void AddPivotTables()
        {
	        AddPivotTable_PivotByPatRank();
	        AddPivotTable_PivotByResource();
	        AddPivotTable_PivotByApplicationStatus();
	        AddPivotTable_PivotByStatusApplication();
        }

        public void AddPivotTable_PivotByPatRank()
        {
	        string pivotTableName = "ITRByApplicationPatRank";
	        string pivotTableHeader = "ITR by Application Pat Rank";
	        string sourceData = "SourceITRInfo";

            Worksheet activeSheet = PivotTable_Create(pivotTableName, pivotTableHeader, sourceData);
            PivotField field;

	        field = activeSheet.PivotTables(pivotTableName).PivotFields("Application");
            field.Orientation = Microsoft.Office.Interop.Excel.XlPivotFieldOrientation.xlRowField;
	        field.Position = 1;

	        field = activeSheet.PivotTables(pivotTableName).PivotFields("PAT Rank");
            field.Orientation = Microsoft.Office.Interop.Excel.XlPivotFieldOrientation.xlRowField;
	        field.Position = 2;

	        field = activeSheet.PivotTables(pivotTableName).PivotFields("ITR ID");
            field.Orientation = Microsoft.Office.Interop.Excel.XlPivotFieldOrientation.xlRowField;
	        field.Position = 3;

	        field = activeSheet.PivotTables(pivotTableName).PivotFields("Subject");
            field.Orientation = Microsoft.Office.Interop.Excel.XlPivotFieldOrientation.xlRowField;
	        field.Position = 4;

	        PivotTable_CollapseSections(pivotTableName, "Application");
	        FormatColumns_PivotTable(pivotTableName);
        }

        public void AddPivotTable_PivotByResource()
        {
	        string pivotTableName = "ITRByResource";
	        string pivotTableHeader = "ITR by Resource";
	        string sourceData = "SourceITRResourceInfo";

            Worksheet activeSheet = PivotTable_Create(pivotTableName, pivotTableHeader, sourceData);
            PivotField field;

            // For some reason need to do this first and not after adding  fields (which is how it is recorded)
	        // as the Status field disappears???

	        activeSheet.PivotTables(pivotTableName).AddDataField(activeSheet.PivotTables(pivotTableName).PivotFields("Application"), "Count", XlConsolidationFunction.xlCount);

	        field = activeSheet.PivotTables(pivotTableName).PivotFields("Resource ID");
            field.Orientation = Microsoft.Office.Interop.Excel.XlPivotFieldOrientation.xlRowField;
	        field.Position = 1;

	        field = activeSheet.PivotTables(pivotTableName).PivotFields("ITR ID");
            field.Orientation = Microsoft.Office.Interop.Excel.XlPivotFieldOrientation.xlRowField;
	        field.Position = 2;

	        field = activeSheet.PivotTables(pivotTableName).PivotFields("Application");
            field.Orientation = Microsoft.Office.Interop.Excel.XlPivotFieldOrientation.xlRowField;
	        field.Position = 3;

	        field = activeSheet.PivotTables(pivotTableName).PivotFields("Subject");
            field.Orientation = Microsoft.Office.Interop.Excel.XlPivotFieldOrientation.xlRowField;
	        field.Position = 4;

	        field = activeSheet.PivotTables(pivotTableName).PivotFields("Status");
            field.Orientation = Microsoft.Office.Interop.Excel.XlPivotFieldOrientation.xlRowField;
	        field.Position = 5;

	        field = activeSheet.PivotTables(pivotTableName).PivotFields("LOE");
            field.Orientation = Microsoft.Office.Interop.Excel.XlPivotFieldOrientation.xlRowField;
	        field.Position = 6;

	        field = activeSheet.PivotTables(pivotTableName).PivotFields("Age");
            field.Orientation = Microsoft.Office.Interop.Excel.XlPivotFieldOrientation.xlRowField;
	        field.Position = 7;

	        PivotTable_CollapseSections(pivotTableName, "Status");
	        FormatColumns_PivotTable(pivotTableName);
        }

        public void AddPivotTable_PivotByApplicationStatus()
        {
	        string pivotTableName = "ITRByApplicationStatus";
	        string pivotTableHeader = "ITR by Application Status";
	        string sourceData = "SourceITRInfo";

	        // For some reason need to do this first and not after adding  fields (which is how it is recorded)
	        // as the Status field disappears???

            Worksheet activeSheet = PivotTable_Create(pivotTableName, pivotTableHeader, sourceData);
            PivotField field;

	        activeSheet.PivotTables(pivotTableName).AddDataField(activeSheet.PivotTables(pivotTableName).PivotFields("Status"), "Count", XlConsolidationFunction.xlCount);

	        field = activeSheet.PivotTables(pivotTableName).PivotFields("Application");
            field.Orientation = Microsoft.Office.Interop.Excel.XlPivotFieldOrientation.xlRowField;
	        field.Position = 1;

	        field = activeSheet.PivotTables(pivotTableName).PivotFields("Status");
            field.Orientation = Microsoft.Office.Interop.Excel.XlPivotFieldOrientation.xlRowField;
	        field.Position = 2;

	        field = activeSheet.PivotTables(pivotTableName).PivotFields("ITR ID");
            field.Orientation = Microsoft.Office.Interop.Excel.XlPivotFieldOrientation.xlRowField;
	        field.Position = 3;

	        field = activeSheet.PivotTables(pivotTableName).PivotFields("Subject");
            field.Orientation = Microsoft.Office.Interop.Excel.XlPivotFieldOrientation.xlRowField;
	        field.Position = 4;

	        PivotTable_CollapseSections(pivotTableName, "Status");
	        FormatColumns_PivotTable(pivotTableName);
        }

        private void AddPivotTable_PivotByStatusApplication()
        {
	        string pivotTableName = "ITRByStatusApplication";
	        string pivotTableHeader = "ITR by Status Application";
	        string sourceData = "SourceITRResourceInfo";
           
            Worksheet activeSheet = PivotTable_Create(pivotTableName, pivotTableHeader, sourceData);
            PivotField field;
            
	        activeSheet.PivotTables(pivotTableName).AddDataField(activeSheet.PivotTables(pivotTableName).PivotFields("Status"), "Count", XlConsolidationFunction.xlCount);

	        field = activeSheet.PivotTables(pivotTableName).PivotFields("Status");
	        field.Orientation = Microsoft.Office.Interop.Excel.XlPivotFieldOrientation.xlPageField;
	        field.Position = 1;

	        field = activeSheet.PivotTables(pivotTableName).PivotFields("Application");
            field.Orientation = Microsoft.Office.Interop.Excel.XlPivotFieldOrientation.xlRowField;
	        field.Position = 1;

	        field = activeSheet.PivotTables(pivotTableName).PivotFields("ITR ID");
            field.Orientation = Microsoft.Office.Interop.Excel.XlPivotFieldOrientation.xlRowField;
	        field.Position = 2;

	        field = activeSheet.PivotTables(pivotTableName).PivotFields("Subject");
            field.Orientation = Microsoft.Office.Interop.Excel.XlPivotFieldOrientation.xlRowField;
	        field.Position = 3;

	        field = activeSheet.PivotTables(pivotTableName).PivotFields("Age");
            field.Orientation = Microsoft.Office.Interop.Excel.XlPivotFieldOrientation.xlRowField;
	        field.Position = 4;

	        field = activeSheet.PivotTables(pivotTableName).PivotFields("Resource ID");
            field.Orientation = Microsoft.Office.Interop.Excel.XlPivotFieldOrientation.xlRowField;
	        field.Position = 5;

	        PivotTable_CollapseSections(pivotTableName, "Status");
	        FormatColumns_PivotTable(pivotTableName);
        }

        public void MoveITRCommentsBelow()
        {
            Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;

	        Range currentCell;
	        int currentRow = 0;

	        // TODO: Add a constant for this 
	        currentRow = 7;
	        // Second ITR

	        currentCell = activeSheet.Cells[currentRow, 1];

	        while (!string.IsNullOrEmpty(currentCell.Value))
            {
                //Debug.Print(currentCell.Value);

		        currentCell.Activate();
		        MoveCommentsBelowAndGroup();

		        // Skip past four new rows
		        currentRow = currentRow + 5;
		        currentCell = activeSheet.Cells[currentRow, 1];
	        }

	        // Do one more as we reach backward for the data for the last ITR.

	        currentCell.Activate();
	        MoveCommentsBelowAndGroup();

	        // Now that the comment data has moved, delete the un-needed columns.
	        activeSheet.Columns["O:R"].Delete();

        }

        public void PivotTable_CollapseSections(string pivotTableName, string fieldName)
        {
	        PivotTable pT = default(PivotTable);
	        PivotField pF = default(PivotField);
	        //Dim pI As PivotItem

	        pT = Globals.ThisAddIn.Application.ActiveSheet.PivotTables(pivotTableName);

	        // Not sure what the difference is between these two.  Both seem to work.

	        pF = pT.PivotFields(fieldName);
	        //    Set pF = pT.RowFields(fieldName)

	        // Collapse each of the sections

	        // TODO: Figure out why this throws exception??  Works in VBA??
	        //For Each pI In pF.PivotItems
	        //    pI.ShowDetail = False
	        //Next
        }

        // TODO: Have this take a before/after sheet argument so can put sheets where we want them.
        public Worksheet PivotTable_Create(string pivotTableName, string pivotTableHeader, string sourceData)
        {
	        string pivotTableDestination = null;
	        Worksheet newSheet = Globals.ThisAddIn.Application.Sheets.Add();
	        newSheet.Name = pivotTableName;

	        pivotTableDestination = string.Format("{0}!R3C1", newSheet.Name);

            //PivotCaches caches = Globals.ThisAddIn.Application.ActiveWorkbook.PivotCaches();
            //caches.Create(...);

            (Globals.ThisAddIn.Application.ActiveWorkbook.PivotCaches()).Create(
                SourceType: XlPivotTableSourceType.xlDatabase, 
                SourceData: sourceData, 
                Version: XlPivotTableVersionList.xlPivotTableVersion12).CreatePivotTable(TableDestination: pivotTableDestination, 
                TableName: pivotTableName, 
                DefaultVersion: XlPivotTableVersionList.xlPivotTableVersion12);

	        Globals.ThisAddIn.Application.ActiveSheet.PivotTables(pivotTableName).CompactLayoutRowHeader = pivotTableHeader;

            return newSheet;
        }

    #endregion

    }

}
