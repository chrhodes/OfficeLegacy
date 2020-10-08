using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using Microsoft.Office.Interop.Excel;

using ExcelHlp=AddinHelper.Excel;

namespace SupportTools_Excel.User_Interface.Task_Panes
{
    public partial class TaskPane_MTreaty : UserControl
    {        
        #region Initialization

        public TaskPane_MTreaty()
        {
            InitializeComponent();
        }

        private void TaskPane_MTreaty_Load(object sender, EventArgs e)
        {
            ucEnvironmentList1.PopulateListFromFile(@"C:\Temp\ExcelMtreatyConfig.xml");
            ucFileTypeList1.PopulateListFromFile(@"C:\Temp\ExcelMtreatyConfig.xml");
            CurrentDate = DateTime.Now;

            ResetToInitialState();
        }

        #endregion

        public DateTime CurrentDate
        {
            get;
            set;
        }
        public DateTime CycleDate
        {
            get;
            set;
        }
        public string FileType
        {
            get;
            set;
        }
        public string OutputPath
        {
            get;
            set;
        }
        public string Path
        {
            get;
            set;
        }
        public string RelativePath
        {
            get;
            set;
        }
        public int RecordCount
        {
            get;
            set;
        }
        public double TotalAmount
        {
            get;
            set;
        }
        
        #region Event Handlers

        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            Worksheet ws;
            Workbook sourceWorkbook;

            if (null == (sourceWorkbook = OpenManualFile()))
            {
            	MessageBox.Show("Cannot open source workbook");
                return;
            }

            if (null == (ws = SelectSourceWorkSheet(sourceWorkbook)))
            {
            	MessageBox.Show("Cannot locate source worksheet");
                return;
            }

            ws.Copy();
            sourceWorkbook.Close(SaveChanges: false);
            // Not sure if we need to do this.  Compare ws before and after someday to see.
            ws = Globals.ThisAddIn.Application.ActiveSheet;

            string asOfDate = ws.Range["A1"].Value;
            asOfDate = asOfDate.ToUpper().Replace("AS OF ", "");
            dtpAsOfDate.Value = DateTime.Parse(asOfDate);
            // Make the user pick again.
            pnlTask_ConfirmAsOfDate.BackColor = Color.White;
            dtpAsOfDate.Enabled = true;
            dtpAsOfDate.Focus();
        }

        private void btnPlayAgain_Click(object sender, EventArgs e)
        {
            ResetToInitialState();

            if (Globals.ThisAddIn.Application.ActiveWorkbook != null)
            {
                if (DialogResult.Yes == MessageBox.Show("Close current workbook?", "Close Files?", MessageBoxButtons.YesNo))
                {
            	    Globals.ThisAddIn.Application.ActiveWorkbook.Close(SaveChanges: false);
                }
            }
        }

        private void btnProcessFile_Click(object sender, EventArgs e)
        {
            Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet;

            ProcessManualFile(ws);
            pnlTask_ProcessFile.BackColor = Color.Green;
            btnProcessFile.Enabled = false;
            btnSaveOutputFiles.Enabled = true;
            btnSaveOutputFiles.Focus();
        }

        private void btnReviewOutput_Click(object sender, EventArgs e)
        {
            OpenOutputFolder();
        }

        private void btnSaveOutputFiles_Click(object sender, EventArgs e)
        {
            Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet;
            SaveDetailFile();
            SaveControlFile();
            btnSaveOutputFiles.Enabled = false;
        }

        private void dtpAsOfDate_MouseUp(object sender, MouseEventArgs e)
        {
            pnlTask_ConfirmAsOfDate.BackColor = Color.Green;
            btnProcessFile.Enabled = true;
            //btnProcessFile.Focus();
        }

        private void dtpAsOfDate_ValueChanged(object sender, EventArgs e)
        {
            CycleDate = dtpAsOfDate.Value;
            pnlTask_ConfirmAsOfDate.BackColor = Color.Green;
            btnProcessFile.Enabled = true;
            //btnProcessFile.Focus();
        }

        private void ucEnvironmentList1_EnvironmentselectionChanged_Event()
        {
            Path = ucEnvironmentList1.Path;
            pnlTask_SelectEnvironment.BackColor = Color.Green;
            ucFileTypeList1.Enabled = true;
            ucFileTypeList1.Focus();
        }

        private void ucFileTypeList1_FileTypeSelectionChanged_Event()
        {
            FileType = ucFileTypeList1.FileType;
            RelativePath = ucFileTypeList1.RelativePath;
            pnlTask_SelectFileType.BackColor = Color.Green;
            btnOpenFile.Enabled = true;
            btnOpenFile.Focus();
        }

        #endregion

        #region Main Function Routines

        private Workbook OpenManualFile()
        {
            string initialDirectory = String.Format(@"{0}\{1}\{2}", 
                Path, RelativePath, "Original_Files");

            Common.WriteToDebugWindow(String.Format(@"OpenFile: {0}", initialDirectory));
            Workbook wb = null;

            wb = Globals.ThisAddIn.Application.Workbooks.Open(
                Filename: ExcelHlp.GetOpenFileName(initialDirectory), 
                UpdateLinks: false, 
                ReadOnly: true);

            if (wb != null)
            {
                pnlTask_OpenFile.BackColor = Color.Green;
                pnlTask_ConfirmAsOfDate.BackColor = Color.White;                ;
            }
            else
            {
            	pnlTask_OpenFile.BackColor = Color.Red;
            }

            return wb;
        }

        private void OpenOutputFolder()
        {
            try
            {
                System.Diagnostics.Process.Start("explorer.exe", OutputPath);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void ProcessManualFile(Worksheet ws)
        {
            Common.WriteToDebugWindow(String.Format("ProcessManualFile()"));

            switch(FileType)
            {
                case "FundServiceFees":
                    ProcessFundServiceFees(ws);
                    break;

                case "FundAdvisoryFees":
                    ProcessFundAdvisoryFees(ws);
                    break;

                case "CashManagementFees":
                    ProcessCashManagementFees(ws);
                    break;

                case "VITsFees":
                    ProcessVITsFees(ws);
                    break;

                default:
                    MessageBox.Show("Processing of Manual FileType: {0} not currently supported", FileType);
                    break;
            }
        }
        
        #region Manual File Processing Routines

        private void ProcessCashManagementFees(Worksheet ws)
        {
            MessageBox.Show("Not Implemented Yet");
            Common.WriteToDebugWindow(String.Format("ProcessCashManagementFees()"));
        }

        private void ProcessFundAdvisoryFees(Worksheet ws)
        {
            Common.WriteToDebugWindow(String.Format("ProcessFundAdvisoryFees()"));

            // Delete the top two rows

            ((Range)ws.Rows[1]).Delete();
            ((Range)ws.Rows[1]).Delete();

            // Remove the commas

            ((Range)ws.Columns["E:E"]).NumberFormat = "0.00";

            // Determine the number of rows

            RecordCount = ExcelHlp.FindLast_PopulatedRow_InColumn(ws.Range["A1"]);

            // Replace the Fund Codes

            UpdateFundCodes(ws);

            // Add a new column

            ws.Columns["A:A"].Insert(Shift: XlDirection.xlToRight, CopyOrigin: XlInsertFormatOrigin.xlFormatFromLeftOrAbove);

            // And populate with As of Date

            DateTime asOfDate = DateTime.Parse(dtpAsOfDate.Text);
            ws.Range[ws.Cells[1, 1], ws.Cells[RecordCount, 1]].Value = asOfDate.ToShortDateString();

            // Gather some information for the control file

            Range sumRange = ws.Range[string.Format("F1:F{0}", RecordCount)];
            TotalAmount = (double)Globals.ThisAddIn.Application.WorksheetFunction.Sum(sumRange);
                        
            CleanOutExtraDataOnWorksheet(ws);

            Common.WriteToDebugWindow(string.Format("CycleDate:{0},CurrentDate:{1},RecordCount={2},TotalAmount={3}",
                CycleDate, CurrentDate, RecordCount, TotalAmount));
        }

        /// <summary>
        /// Ensure the worksheet only has rows with actual data.  Remove cells with blanks, total rows, etc.
        /// </summary>
        /// <param name="ws"></param>
        private void CleanOutExtraDataOnWorksheet(Worksheet ws)
        {
            Common.WriteToDebugWindow(String.Format("CleanOutExtraDataOnWorksheet()"));

            // TODO: May need different end columns.  For now presume the 6 from fees spreadsheets.
            // Take the switch(FileType) code from GetOutputFileName();
         
            Range dataRange = ws.Range[ws.Cells[1, 1], ws.Cells[RecordCount, 6]];
            dataRange.Copy();
            Globals.ThisAddIn.Application.ActiveWorkbook.Sheets.Add();
            ((Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Paste();
            // Quietly delete the source worksheet to avoid confusion.
            ExcelHlp.DeleteSheet(ws);
        }

        private void ProcessFundServiceFees(Worksheet ws)
        {
            Common.WriteToDebugWindow(String.Format("ProcessFundServiceFees()"));

            // Delete the top two rows

            ((Range)ws.Rows[1]).Delete();
            ((Range)ws.Rows[1]).Delete();

            // Remove the commas

            ((Range)ws.Columns["E:E"]).NumberFormat = "0.00";

            // Determine the number of rows

            RecordCount = ExcelHlp.FindLast_PopulatedRow_InColumn(ws.Range["A1"]);

            // Replace the Fund Codes

            UpdateFundCodes(ws);

            // Add a new column

            ws.Columns["A:A"].Insert(Shift: XlDirection.xlToRight, CopyOrigin: XlInsertFormatOrigin.xlFormatFromLeftOrAbove);

            // And populate with As of Date

            DateTime asOfDate = DateTime.Parse(dtpAsOfDate.Text);
            ws.Range[ws.Cells[1, 1], ws.Cells[RecordCount, 1]].Value = asOfDate.ToShortDateString();

            // Gather some information for the control file

            Range sumRange = ws.Range[string.Format("F1:F{0}", RecordCount)];
            TotalAmount = (double)Globals.ThisAddIn.Application.WorksheetFunction.Sum(sumRange);
                        
            CleanOutExtraDataOnWorksheet(ws);

            Common.WriteToDebugWindow(string.Format("CycleDate:{0},CurrentDate:{1},RecordCount={2},TotalAmount={3}",
                CycleDate, CurrentDate, RecordCount, TotalAmount));
        }

        private void ProcessVITsFees(Worksheet ws)
        {
            MessageBox.Show("Not Implemented Yet - Need processing steps.");            
            Common.WriteToDebugWindow(String.Format("ProcessVITsFees()"));
        }

        #endregion

        private string GetControlFileName()
        {
            Common.WriteToDebugWindow(String.Format("GetControlFileName()"));
            string fileName = null;

            // NB. All filenames start with an "_".  This prevents the file from being automatically processed.
            // This may change as we gain confidence in process.  A file dialog is presented to user at which
            // point they may change the name.

            switch(FileType)
            {
                case "FundServiceFees":
                    fileName = String.Format("_{0}_{1}.csv", "ServiceFees", "Control");
                    break;

                case "FundAdvisoryFees":
                    fileName = String.Format("_{0}_{1}.csv", "AdvisoryFees", "Control");
                    break;

                case "CashManagementFees":
                    fileName = String.Format("_{0}_{1}.csv", "CashManagementFees", "Control");
                    break;

                case "VITsFees":
                    fileName = String.Format("_{0}_{1}.csv", "VITsFees", "Control");
                    break;

                default:
                    MessageBox.Show("Processing of Manual FileType: {0} not currently supported", FileType);
                    break;
            }

            return fileName;
        }

        private string GetOutputFileName()
        {
            Common.WriteToDebugWindow(String.Format("GetOutputFileName()"));
            string fileName = null;

            // NB. All filenames start with an "_".  This prevents the file from being automatically processed.
            // This may change as we gain confidence in process.  A file dialog is presented to user at which
            // point they may change the name.

            switch(FileType)
            {
                case "FundServiceFees":
                    fileName = String.Format("_{0}_{1}.csv", "ServiceFees", "Detail");
                    break;

                case "FundAdvisoryFees":
                    fileName = String.Format("_{0}_{1}.csv", "AdvisoryFees", "Detail");
                    break;

                case "CashManagementFees":
                    fileName = String.Format("_{0}_{1}.csv", "CashManagementFees", "Detail");
                    break;

                case "VITsFees":
                    fileName = String.Format("_{0}_{1}.csv", "VITsFees", "Detail");
                    break;

                default:
                    MessageBox.Show("Processing of Manual FileType: {0} not currently supported", FileType);
                    break;
            }

            return fileName;
        }

        private void ResetToInitialState()
        {
            Common.WriteToDebugWindow(String.Format("ResetToInitialState()"));

            pnlTask_SelectEnvironment.BackColor = Color.White;
            pnlTask_SelectFileType.BackColor = Color.White;
            pnlTask_OpenFile.BackColor = Color.White;
            pnlTask_ConfirmAsOfDate.BackColor = Color.White;
            pnlTask_ProcessFile.BackColor = Color.White;
            pnlTask_SaveControlFile.BackColor = Color.White;
            pnlTask_SaveDetailFile.BackColor = Color.White;

            ucFileTypeList1.Enabled = false;
            btnOpenFile.Enabled = false;
            dtpAsOfDate.Enabled = false;
            btnProcessFile.Enabled = false;
            btnSaveOutputFiles.Enabled = false;
            ucEnvironmentList1.Focus();
        }

        private void SaveControlFile()
        {
            Common.WriteToDebugWindow(String.Format("SaveControlFile:{0},{1},{2},{3:F2}", 
                        CycleDate.ToShortDateString(), CurrentDate.ToShortDateString(), RecordCount, TotalAmount));

            string fileName = null;
            string initialDirectory = OutputPath;

            if(null != (fileName = GetControlFileName()))
            {
                string outputFile = ExcelHlp.GetSaveFileName(initialDirectory, fileName);

                FileInfo fileInfo = new FileInfo(outputFile);
                OutputPath = fileInfo.DirectoryName;

                using(StreamWriter writer = new StreamWriter(outputFile))
                {
                    // TODO:(crhodes) Remove TotalAmount for CMF output
                    writer.Write(String.Format("{0},{1},{2},{3:F2}", 
                        CycleDate.ToShortDateString(), CurrentDate.ToShortDateString(), RecordCount, TotalAmount));
                    writer.Close();
                }

                pnlTask_SaveControlFile.BackColor = Color.Green;
            }
            else
            {
                pnlTask_SaveControlFile.BackColor = Color.Red;                
            }
        }

        private void SaveDetailFile()
        {
            Common.WriteToDebugWindow(String.Format(@"SaveDetailFile: {0} - {1}\{2}", FileType, Path, RelativePath));
            string fileName = null;
            string initialDirectory = String.Format(@"{0}\{1}", Path, RelativePath);
            Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;

            if(null != (fileName = GetOutputFileName()))
            {
                string outputFile = ExcelHlp.GetSaveFileName(initialDirectory, fileName);

                FileInfo fileInfo = new FileInfo(outputFile);
                OutputPath = fileInfo.DirectoryName;

                wb.SaveAs(Filename: outputFile, FileFormat: XlFileFormat.xlCSV);
                pnlTask_SaveDetailFile.BackColor = Color.Green;
            }
            else
            {
                pnlTask_SaveDetailFile.BackColor = Color.Red;                
            }
        }

        private Worksheet SelectSourceWorkSheet(Workbook sourceWorkbook)
        {
            Common.WriteToDebugWindow(String.Format("SelectSourceWorkSheet()"));
            Worksheet ws = null;

            switch(FileType)
            {
                case "FundServiceFees":
                    try
                    {
                         ws = sourceWorkbook.Worksheets[Common.cMTREATY_FUND_SERVICE_FEES_SHEETNAME];
                    }
                    catch(Exception)
                    {
                       Common.WriteToDebugWindow(String.Format("Cannot find required worksheet: {0}", Common.cMTREATY_FUND_SERVICE_FEES_SHEETNAME));
                    }

                    break;

                case "FundAdvisoryFees":
                    try
                    {
                         ws = sourceWorkbook.Worksheets[Common.cMTREATY_FUND_ADVISORY_FEES_SHEETNAME];
                    }
                    catch(Exception)
                    {
                        Common.WriteToDebugWindow(String.Format("Cannot find required worksheet: {0}", Common.cMTREATY_FUND_ADVISORY_FEES_SHEETNAME));
                    }

                    break;

                case "CashManagementFees":
                    try
                    {
                         ws = sourceWorkbook.Worksheets[Common.cMTREATY_CASH_MANAGEMENT_FEES_SHEETNAME];
                    }
                    catch(Exception)
                    {
                        Common.WriteToDebugWindow(String.Format("Cannot find required worksheet: {0}", Common.cMTREATY_CASH_MANAGEMENT_FEES_SHEETNAME));
                    }

                    break;

                case "VITsFees":
                    try
                    {
                         ws = sourceWorkbook.Worksheets[Common.cMTREATY_VITS_FEES_SHEETNAME];
                    }
                    catch(Exception)
                    {
                        Common.WriteToDebugWindow(String.Format("Cannot find required worksheet: {0}", Common.cMTREATY_VITS_FEES_SHEETNAME));
                    }

                    break;

                default:
                    MessageBox.Show("Processing of Manual FileType: {0} not currently supported", FileType);
                    break;
            }

            return ws;
        }

        /// <summary>
        /// Use PDALL FundCode for funds with certain FundNames.
        /// </summary>
        /// <param name="ws"></param>
        private void UpdateFundCodes(Worksheet ws)
        {
            Common.WriteToDebugWindow(String.Format("UpdateFundCodes()"));
            
            // The fund names are inconsistently mixed case in the spreadsheet.
            // Match against UPPERCASE and convert the value below.


            string[] fundNames = { 
                                    //"PD CONSERVATIVE GROWTH", 
                                    //"PD MODERATE GROWTH", 
                                    //"PD GROWTH", 
                                    "PD AGGREGATE BOND INDEX", 
                                    "PD HIGH YIELD BOND MARKET", 
                                    "PD LARGE-CAP GROWTH INDEX", 
                                    "PD LARGE-CAP VALUE INDEX", 
                                    "PD SMALL-CAP GROWTH INDEX", 
                                    "PD SMALL-CAP VALUE INDEX", 
                                    "PD INTERNATIONAL LARGE-CAP", 
                                    "PD EMERGING MARKETS" };

            Range matchRange = null;

            try
            {
                ((Range)ws.Columns["C:C"]).Replace(What: "INFLMNG", Replacement: "INFMNG", LookAt: XlLookAt.xlPart);
            }
            catch(Exception)
            {
                // May not be any matches.  Quietly ignore.                
            }

            // Traverse the fund names looking for ones in the list that should be converted to PDALL

            for(int i = 1 ; i <= RecordCount; i++)
            {
                string fundName = ws.Cells[i, 4].Value.ToUpper();

                if (fundNames.LongCount(p => p == fundName) > 0)
                {
                	ws.Cells[i,3].Value = "PDALL";
                }

                //Common.WriteToDebugWindow(String.Format("FundCode:{0}  FundName:{1}  Count:{2}", 
                //    ws.Cells[i,3].Value, ws.Cells[i, 4].Value, fundNames.LongCount(p => p == testString)));

            }
        }

        //private void UpdateFundCodes(Worksheet ws)
        //{
        //    Common.WriteToDebugWindow(String.Format("UpdateFundCodes()"));

        //    string[] fundCodes = { "PS4A", "PS4B", "PS4D", "PS4E", "PS4F", "PS4G", "PS4H", "PS4I" };

        //    Globals.ThisAddIn.Application.DisplayAlerts = false;

        //    try
        //    {
        //        ((Range)ws.Columns["C:C"]).Replace(What: "INFLMNG", Replacement: "INFMNG", LookAt: XlLookAt.xlPart);
        //    }
        //    catch(Exception)
        //    {
        //        // May not be any matches.  Quietly ignore.                
        //    }

        //    foreach(string code in fundCodes)
        //    {
        //        try
        //        {
        //             ((Range)ws.Columns["C:C"]).Replace(What: code, Replacement: "PDALL", LookAt: XlLookAt.xlPart);   
        //        }
        //        catch(Exception)
        //        {
        //            // May not be any matches.  Quietly ignore.
        //        }
        //    }

        //    Globals.ThisAddIn.Application.DisplayAlerts = true;
        //}

        #endregion
    }
}
