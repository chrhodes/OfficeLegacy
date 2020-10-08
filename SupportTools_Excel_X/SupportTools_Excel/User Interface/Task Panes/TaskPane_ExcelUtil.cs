using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using ExcelHlp = AddinHelper.Excel;

namespace SupportTools_Excel.User_Interface.Task_Panes
{
    public partial class TaskPane_ExcelUtil : UserControl
    {
        private const string _ERROR_EMPTY_CELL = "Cell is empty.  Must select a populated starting cell first.";
        private const short _INDENT_LEVEL = 1;
        private const short _COL_WIDTH = 3;
        private const short _NOTE_WIDTH = 20;
        private const short _FILE_FONT_SIZE = 6;

        private const short _FOLDER_FONT_SIZE = 8;
        private const int _HEADING_ROW = 2;

        private const int _INITIAL_ROW = _HEADING_ROW + 1;
	        // Folder level info starts here
        private const short _FOLDER_INFO_COL = 1;
        private const short _FOLDER_INFO_LEN = 10;
	        // File Info starts here
        private const short _FILE_INFO_COL = 5;
        private const short _FILE_INFO_LEN = 4;
        private const short _NOTE_COL = 10;
	        // Map Info starts here
        private const short _INITIAL_COL = 11;
        private const bool _MAKE_BOLD = true;

        public enum _DateType : int
        {
	        LastCreate = 1,
	        LastWrite = 2,
	        LastAccess = 3
        }

        #region Initialization

         public TaskPane_ExcelUtil()
        {
            InitializeComponent();
        }

        private void TaskPane_ExcelUtil_Load(object sender, EventArgs e)
        {

        }

        #endregion

        #region Event Handlers

        private void btnAddFooter_Click(object sender, EventArgs e)
        {
            Actions.Excel_PageFormatting.AddFooter();
        }

        private void btnAddHeader_Click(object sender, EventArgs e)
        {
            Actions.Excel_PageFormatting.AddHeader();
        }

        private void btnCreateFolderMap_Click(object sender, EventArgs e)
        {
            Actions.Excel_FolderMaps.CreateFolderMap();
        }

        private void btnCreateTableOfContents_Click(object sender, EventArgs e)
        {
            Actions.Excel_TableOfContents.CreateTableOfContents();
        }

        private void btnDeleteDuplicateRows_Click(object sender, EventArgs e)
        {
            DeleteDuplicateRows();
        }
        private void btnDisplayDocumentProperties_Click(object sender, EventArgs e)
        {
            DisplayDocumentProperties();
        }

        private void btnFormatLandscape_Click(object sender, EventArgs e)
        {
            FormatLandscape();
        }

        private void btnFormatPortrait_Click(object sender, EventArgs e)
        {
            FormatPortrait();
        }

        private void btnGetRowColInfo_Click(object sender, EventArgs e)
        {
            GetRowColInfo();
        }
        private void btnGroupDown_Click(object sender, EventArgs e)
        {
            GroupDown();
        }
        private void btnGroupDownAll_Click(object sender, EventArgs e)
        {
            GroupDownAll();
        }

        private void btnProtectSheets_Click(object sender, EventArgs e)
        {
            ProtectSheets();
        }

        private void btnSearchDown_Click(object sender, EventArgs e)
        {
            SearchDown();
        }

        private void btnSearchRight_Click(object sender, EventArgs e)
        {
            SearchRight();
        }

        private void btnSearchUp_Click(object sender, EventArgs e)
        {
            SearchUp();
        }

        private void btnUnProtectSheets_Click(object sender, EventArgs e)
        {
            UnProtectSheets();
        }
        private void btnSearchLeft_Click(object sender, EventArgs e)
        {
            SearchLeft();
        }

        private void btnUnGroupSelection_Click(object sender, EventArgs e)
        {

        }

        #endregion

        #region Main Function Routines

        private void AddHeader()
        {
            throw new NotImplementedException();
        }

        private void DeleteDuplicateRows()
        {
	        Range currentCell = Globals.ThisAddIn.Application.ActiveCell;
	        int firstRow = currentCell.Row;
	        int lastRow = currentCell.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
	        Hashtable foundRows = new Hashtable();
	        Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;

	        List<string> rowsToDelete = new List<string>();

	        Range rng = default(Range);

            Common.WriteToDebugWindow(string.Format("{0} - {1}", firstRow, lastRow));

	        for (int i = lastRow; i >= firstRow; i += -1)
            {
		        rng = activeSheet.Cells[i, currentCell.Column];

		        Common.WriteToDebugWindow(string.Format("{0} - {1} - {2}", i, rng.Value, lastRow));

		        if (foundRows.Contains(rng.Value))
                {
			        rowsToDelete.Add(string.Format("{0}:{1}", i, i));
		        }
                else
                {
			        foundRows.Add(rng.Value, "X");
		        }
	        }

	        foreach (string row in rowsToDelete)
            {
		        Common.WriteToDebugWindow(row);
		        activeSheet.Rows[row].Delete();
	        }
        }

        private void DisplayDocumentProperties()
        {
            Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;

            Office.DocumentProperties documentProperties = (Office.DocumentProperties)wb.BuiltinDocumentProperties;
            StringBuilder sb = new StringBuilder();
            string value = "";
            string name = "";

            if(documentProperties != null)
            {
                foreach(Office.DocumentProperty property in documentProperties)
                {
                    try
                    {
                        name = property.Name;
                        value = property.Value.ToString();
                    }
                    catch(System.Runtime.InteropServices.COMException ex)
                    {
                        Common.WriteToDebugWindow(string.Format("GetBuiltInPropertyValue({0}) COMException ErrorCode:({1})", name, ex.ErrorCode));
                        value = "Property Not Available Yet";
                    }
                    catch(NullReferenceException )
                    {
                        Common.WriteToDebugWindow(string.Format("GetBuiltInPropertyValue({0}) NullReferenceException", name));
                        value = "Property Not Set";
                    }
                    catch(Exception ex)
                    {
                        Common.WriteToDebugWindow(string.Format("GetBuiltInPropertyValue({0}) Exception >{1})<", name, ex));
                        value = "";
                    }

                    sb.AppendLine(string.Format("P:{0}->{1}<", name, value));
                }
            }
            else
            {
                sb.AppendLine("No Document Properties");
            }

            MessageBox.Show(sb.ToString(), "Document Properties", MessageBoxButtons.OK);
        }

        private void FormatLandscape()
        {
            Globals.ThisAddIn.Application.PrintCommunication = false;

            foreach (Worksheet sheet in Globals.ThisAddIn.Application.ActiveWorkbook.Sheets)
            {
            	sheet.PageSetup.Orientation = XlPageOrientation.xlLandscape;
            }

            Globals.ThisAddIn.Application.PrintCommunication = true;
        }

        private void FormatPortrait()
        {
            Globals.ThisAddIn.Application.PrintCommunication = false;

            foreach (Worksheet sheet in Globals.ThisAddIn.Application.ActiveWorkbook.Sheets)
            {
                sheet.PageSetup.Orientation = XlPageOrientation.xlPortrait;
            }

            Globals.ThisAddIn.Application.PrintCommunication = false;
        }
        
        private void GetRowColInfo()
        {
	        Range rng = Globals.ThisAddIn.Application.ActiveCell;

            txtLastRowSpecial.Text = rng.SpecialCells(XlCellType.xlCellTypeLastCell).Row.ToString();
	        txtLastColSpecial.Text = rng.SpecialCells(XlCellType.xlCellTypeLastCell).Column.ToString();

            // TODO (crhodes): Make FindXXXPopulated{Row,Col} take an enum: First, Next,Previous, Last

	        txtFirstRowSearch.Text = ExcelHlp.FindFirst_PopulatedRow_InColumn(rng).ToString();
	        txtFirstColSearch.Text = ExcelHlp.FindFirst_PopulatedColumn_InRow(rng).ToString();

            txtPreviousRowSearch.Text = ExcelHlp.FindPrevious_PopulatedRow_InColumn(rng).ToString();
            txtPreviousColSearch.Text = ExcelHlp.FindPrevious_PopulatedColumn_InRow(rng).ToString();

            txtNextRowSearch.Text = ExcelHlp.FindNext_PopulatedRow_InColumn(rng).ToString();
            txtNextColSearch.Text = ExcelHlp.FindNext_PopulatedColumn_InRow(rng).ToString();

            txtLastRowSearch.Text = ExcelHlp.FindLast_PopulatedRow_InColumn(rng).ToString();
            txtLastColSearch.Text = ExcelHlp.FindLast_PopulatedColumn_InRow(rng).ToString();
        }

        private void GroupDown()
        {
            Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;
            Range activeCell = Globals.ThisAddIn.Application.ActiveCell;
            
            int currentRow;
            int currentColumn;
            int lastPopulatedRow;
            int lastPopulatedColumn;
            int endRowOfSection;

            if (activeCell.Value == null)
            {
            	MessageBox.Show(_ERROR_EMPTY_CELL);
            }
            else
            {
                // Get the last populated cell on the worksheet

                lastPopulatedRow = activeCell.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
                lastPopulatedColumn = activeCell.SpecialCells(XlCellType.xlCellTypeLastCell).Column;

                // Save where we currently are located
                currentRow = activeCell.Row;
                currentColumn = activeCell.Column;

                endRowOfSection = Actions.Excel_FolderMaps.GetEndOfSectionDown(currentRow, currentColumn, lastPopulatedRow, currentColumn);
                activeSheet.Rows[currentRow + 1 + ":" + endRowOfSection].Group();

                //activeSheet.Cells[endRowOfSection, startColumn].Select();
                activeSheet.Rows[currentRow + 1 + ":" + endRowOfSection].Hidden = true;
            }
        }

        private void GroupDownAll()
        {
            Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;
            Range activeCell = Globals.ThisAddIn.Application.ActiveCell;

            int currentRow = activeCell.Row;
            int currentColumn = activeCell.Column;

            int lastPopulatedRow;
            int lastPopulatedColumn;
            int endRowOfSection;

            if (activeCell.Value == null)
            {
                MessageBox.Show(_ERROR_EMPTY_CELL);
            }
            else
            {
                // Get the last populated cell on the worksheet

                lastPopulatedRow = activeCell.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
                lastPopulatedColumn = activeCell.SpecialCells(XlCellType.xlCellTypeLastCell).Column;

                while (currentRow < lastPopulatedRow)
                {
                    endRowOfSection = Actions.Excel_FolderMaps.GetEndOfSectionDown(currentRow, currentColumn, lastPopulatedRow, currentColumn);
                    activeSheet.Rows[currentRow + 1 + ":" + endRowOfSection].Group();
                    activeSheet.Rows[currentRow + 1 + ":" + endRowOfSection].Hidden = true;

                    // Move to the next possible row to collapse.  
                    currentRow = endRowOfSection + 1;                    

                    // Keep going if next row does not have an empty cell or past end of data.
                    while (activeSheet.Cells[currentRow + 1, currentColumn].Value != null)
                    {
                        if (currentRow >= lastPopulatedRow)
                        {
                        	break;
                        }

                        currentRow++;
                    }
                }
            }
        }

        private void ProtectSheets()
        {
            foreach (Worksheet sheet in Globals.ThisAddIn.Application.ActiveWorkbook.Sheets)
            {
                sheet.Protect();
            }
        }

        private void SearchDown()
        {
            Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;
            Range activeCell = Globals.ThisAddIn.Application.ActiveCell;

            int matchRow = ExcelHlp.FindNext_PopulatedRow_InColumn(activeCell);
            int matchColumn = activeCell.Column;

            activeSheet.Cells[matchRow, matchColumn].Select();
        }

        private void SearchLeft()
        {
            Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;
            Range activeCell = Globals.ThisAddIn.Application.ActiveCell;

            int matchRow = activeCell.Row;
            int matchColumn = ExcelHlp.FindPrevious_PopulatedColumn_InRow(activeCell);

            activeSheet.Cells[matchRow, matchColumn].Select();
        }

        private void SearchRight()
        {
            Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;
            Range activeCell = Globals.ThisAddIn.Application.ActiveCell;

            int matchRow = activeCell.Row;
            int matchColumn = ExcelHlp.FindNext_PopulatedColumn_InRow(activeCell);

            activeSheet.Cells[matchRow, matchColumn].Select();
        }

        private void SearchUp()
        {
            Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;
            Range activeCell = Globals.ThisAddIn.Application.ActiveCell;

            int matchRow = ExcelHlp.FindPrevious_PopulatedRow_InColumn(activeCell);
            int matchColumn = activeCell.Column;

            activeSheet.Cells[matchRow, matchColumn].Select();
        }

        private void UnProtectSheets()
        {
            foreach (Worksheet sheet in Globals.ThisAddIn.Application.ActiveWorkbook.Sheets)
            {
                sheet.Unprotect();
            }
        }
        #endregion

    }
}