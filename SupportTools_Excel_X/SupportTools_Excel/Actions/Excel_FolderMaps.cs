using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

using Microsoft.Office.Interop.Excel;

namespace SupportTools_Excel.Actions
{
    class Excel_FolderMaps
    {
        private const int _INDENT_LEVEL = 1;
        private const int _COL_WIDTH = 3;
        private const int _NOTE_WIDTH = 20;
        private const int _FILE_FONT_SIZE = 6;

        private const int _FOLDER_FONT_SIZE = 8;
        private const int _HEADING_ROW = 2;

        private const int _INITIAL_ROW = _HEADING_ROW + 1;

        private const string _ERROR_EMPTY_CELL = "Cell is empty.  Must select a populated starting cell first.";
        // Folder level info starts here
        private const int _FOLDER_INFO_COL = 1;
        private const int _FOLDER_INFO_LEN = 10;
        // File Info starts here
        private const int _FILE_INFO_COL = 5;
        private const int _FILE_INFO_LEN = 4;
        private const int _NOTE_COL = 10;
        // Map Info starts here
        private const int _INITIAL_COL = 11;


        private const bool _MAKE_BOLD = true;
        public enum _DateType : int
        {
            LastCreate = 1,
            LastWrite = 2,
            LastAccess = 3
        }

        #region "Private Types and Variables"

        private static int _Row;

        private static int _Column;

        private static string _StartingFolder;
        private static bool _LimitLevels;
        private static int _LimitLevel;
        private static bool _GroupResults;
        private static int _GroupLevel;
        private static int _GroupStartRow;
        private static bool _PatternMatchFolderHighlight;
        private static string _FolderMatchPattern;
        private static bool _PatternMatchFileOutput;
        private static string _FileMatchPattern;
        private static bool _ShowFiles;
        private static bool _ShowFolders;
        private static bool _SkipFoldersWithNoFiles;
        private static int _TotalFolders;
        private static int _TotalFiles;
        private static long _TotalSize;
        private static int _MonthsSinceCreated;
        private static int _MonthsSinceWritten;
        private static int _MonthsSinceAccessed;
        private static int _MonthsCreatedColor;
        private static int _MonthsWrittenColor;
        private static int _MonthsAccessedColor;
        private static int _MonthsDefaultColor;
        private static bool _ColorCodeDates;
        private static int _FolderMatchColor;
        private static int _PatternMatchFileColor;
        private static int _NoAccessColor;
        private static int _PathTooLongColor;
        private static bool _CheckIllegalCharacters;
        private static int _IllegalCharactersColor;
        private static string _IllegalFileCharacters;

        private static string _IllegalFolderCharacters;
        private static bool _CheckSharePointFileNameLength;
        private static int _IllegalFileNameLengthColor;

        private static int _MaxFileNameLength;

        #endregion

        public static void CreateFolderMap()
        {
            using (User_Interface.Forms.frmExcel_FolderMaps frmF = new User_Interface.Forms.frmExcel_FolderMaps())
            {
                try
                {
                    if (System.Windows.Forms.DialogResult.Cancel == frmF.ShowDialog())
                    {
                        return;
                    }

                    Globals.ThisAddIn.Application.ScreenUpdating = false;

                    // Grab the information we need from the form.  I guess we could
                    // unload the form and ditch it, but that would lose any state.
                    InitializeLocalsFromFormFields(frmF);

                    if (_StartingFolder.Length > 0)
                    {
                        PopulateFolderMap(ref _StartingFolder);
                        SaveFolderMap(_StartingFolder);
                    }
                    else
                    {
                        MessageBox.Show("Must select starting folder");
                    }

                }
                catch (Exception ex)
                {
                    // TODO: EntLib
                    MessageBox.Show(ex.ToString());
                    throw (ex);
                }
                finally
                {
                    Globals.ThisAddIn.Application.ScreenUpdating = true;
                }
            }
        }

        private static void InitializeLocalsFromFormFields(User_Interface.Forms.frmExcel_FolderMaps frmF)
        {
            _StartingFolder = frmF.txtStartingFolder.Text;
            _LimitLevels = frmF.chkLimitLevels.Checked;
            _LimitLevel = (int)(Convert.ToInt16(frmF.txtLimitLevel.Text) + _INITIAL_COL);
            // Map data starts at _INITIAL_COL
            _GroupResults = frmF.chkGroupResults.Checked;
            _GroupLevel = Convert.ToInt16(frmF.txtGroupLevel.Text);
            _PatternMatchFolderHighlight = frmF.chkPatternMatchFolderHighlight.Checked;
            _FolderMatchPattern = frmF.txtFolderMatchPattern.Text;
            _PatternMatchFileOutput = frmF.chkPatternMatchFileOutput.Checked;
            _FileMatchPattern = frmF.txtFileMatchPattern.Text;
            _ShowFiles = frmF.chkShowFiles.Checked;
            _ShowFolders = frmF.chkShowFolders.Checked;
            _SkipFoldersWithNoFiles = frmF.chkSkipFoldersWithNoFiles.Checked;

            _FolderMatchColor = RGB(frmF.pnlFolderHighlightColor.BackColor.R, frmF.pnlFolderHighlightColor.BackColor.G, frmF.pnlFolderHighlightColor.BackColor.B);
            _PathTooLongColor = RGB(frmF.pnlPathTooLongColor.BackColor.R, frmF.pnlPathTooLongColor.BackColor.G, frmF.pnlPathTooLongColor.BackColor.B);
            _NoAccessColor = RGB(frmF.pnlNoAccessColor.BackColor.R, frmF.pnlNoAccessColor.BackColor.G, frmF.pnlNoAccessColor.BackColor.B);

            _PatternMatchFileColor = RGB(frmF.pnlPatternMatchFileColor.BackColor.R, frmF.pnlPatternMatchFileColor.BackColor.G, frmF.pnlPatternMatchFileColor.BackColor.B);

            _ColorCodeDates = frmF.chkColorCodeDates.Checked;

            _MonthsDefaultColor = RGB(frmF.pnlDefaultColor.BackColor.R, frmF.pnlDefaultColor.BackColor.G, frmF.pnlDefaultColor.BackColor.B);

            _MonthsCreatedColor = RGB(frmF.pnlMonthCreatedColor.BackColor.R, frmF.pnlMonthCreatedColor.BackColor.G, frmF.pnlMonthCreatedColor.BackColor.B);
            _MonthsWrittenColor = RGB(frmF.pnlMonthWrittenColor.BackColor.R, frmF.pnlMonthWrittenColor.BackColor.G, frmF.pnlMonthWrittenColor.BackColor.B);
            _MonthsAccessedColor = RGB(frmF.pnlMonthAccessedColor.BackColor.R, frmF.pnlMonthAccessedColor.BackColor.G, frmF.pnlMonthAccessedColor.BackColor.B);

            _MonthsSinceCreated = Convert.ToInt16(frmF.txtMonthsSinceCreated.Text);
            _MonthsSinceWritten = Convert.ToInt16(frmF.txtMonthsSinceWritten.Text);
            _MonthsSinceAccessed = Convert.ToInt16(frmF.txtMonthsSinceAccessed.Text);

            _PathTooLongColor = RGB(frmF.pnlPathTooLongColor.BackColor.R, frmF.pnlPathTooLongColor.BackColor.G, frmF.pnlPathTooLongColor.BackColor.B);

            _CheckIllegalCharacters = frmF.chkIllegalCharacters.Checked;
            _IllegalCharactersColor = frmF.pnlIllegalCharactersColor.BackColor.ToArgb();
            _IllegalFileCharacters = frmF.txtIllegalFileCharacters.Text;
            _IllegalFolderCharacters = frmF.txtIllegalFolderCharacters.Text;

            _CheckSharePointFileNameLength = frmF.chkFileNameLength.Checked;

            _IllegalFileNameLengthColor = RGB(frmF.pnlIllegalFileNameLengthColor.BackColor.R, frmF.pnlIllegalFileNameLengthColor.BackColor.G, frmF.pnlIllegalFileNameLengthColor.BackColor.B);
            _MaxFileNameLength = Convert.ToInt32(frmF.txtMaxFileNameLength.Text);

        }

        private static int RGB(Byte red, Byte green, Byte blue)
        {
            return (blue << 0) | (green << 8) | (red << 16);
        }

        // $Description :Open the folder to map and descend.  Called from frmExcel_FolderMaps.

        public static void PopulateFolderMap(ref string startingFolder)
        {
            Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet workSheet;

	        try
            {
		        // Set starting point for folder output

		        _Row = _INITIAL_ROW;
		        _Column = _INITIAL_COL;
		        _TotalFolders = 0;
		        _TotalFiles = 0;
		        _TotalSize = 0;

		        // Always add a new sheet so can accumulate results in Workbook.
		        // Need to handle case when no Workbook exists.  Tried a variety of 
		        // ways to determine empty workbook.  PERSONAL.XLS may be open if 
		        // macros have been created so cannot rely on Workbooks.count.
		        // Worksheets.Count throws an exception.
		        //
		        //Dim wb as Microsoft.Office.Interop.Excel.Workbook
		        //
		        //For Each wb In .Workbooks
		        //    Debug.Print(wb.Name)
		        //Next

		        //Dim ws As Microsoft.Office.Interop.Excel.Worksheet

		        //For Each ws In .Worksheets
		        //    Debug.Print(ws.Name)
		        //Next

		        // ActiveWorkbook seems to work reliably.  Maybe Interop issue, who knows.

		        if (workbook == null)
                {
			        workbook = Globals.ThisAddIn.Application.Workbooks.Add();
			        // Get a new WorkSheet (or more :)) for free.
                    workSheet = workbook.ActiveSheet;
		        }
                else 
                {
			        // TODO: Prompt to use existing sheet if found.
			        workSheet = workbook.Worksheets.Add();
		        }

		        workSheet.Cells[_HEADING_ROW, _FOLDER_INFO_COL].Value = "Cummulative Folder Count";
		        workSheet.Cells[_HEADING_ROW, _FOLDER_INFO_COL + 1].Value = "Cummulative File Count";
		        workSheet.Cells[_HEADING_ROW, _FOLDER_INFO_COL + 2].Value = "Cummulative Size";
                               
		        workSheet.Cells[_HEADING_ROW, _FILE_INFO_COL].Value = "Count";
		        workSheet.Cells[_HEADING_ROW, _FILE_INFO_COL + 1].Value = "Size";
		        workSheet.Cells[_HEADING_ROW, _FILE_INFO_COL + 2].Value = "Last Create";
		        workSheet.Cells[_HEADING_ROW, _FILE_INFO_COL + 3].Value = "Last Write";
		        workSheet.Cells[_HEADING_ROW, _FILE_INFO_COL + 4].Value = "Last Access";

		        workSheet.Cells[_HEADING_ROW, _INITIAL_COL].Value = startingFolder;

		        int startingRow = _Row;
		        int startingColumn = _Column;

                //int numberFoldersLocal = 0;
                //int numberFilesLocal = 0;
                //long sizeFilesLocal = 0;

		        FileInfo dirInfo = new FileInfo(startingFolder);

		        System.DateTime maxLastCreate = System.DateTime.MinValue;
		        System.DateTime maxLastWrite = System.DateTime.MinValue;
		        System.DateTime maxLastAccess = System.DateTime.MinValue;

                //int column = 1;
		        int fontSize = 10;

		        // Note: maxLastDate is passed in, even though we don't use the updated value.
		        // it is used during the recursion process when ListFolders calls itself.  Do
		        // not be mislead and rework this logic!

		        ListFolders(startingFolder, _Column, fontSize, ref _TotalFolders, ref _TotalFiles, ref _TotalSize, ref maxLastCreate, ref maxLastWrite, ref maxLastAccess, _ShowFiles, _ShowFolders);

		        XlPageOrientation orientation = XlPageOrientation.xlPortrait;

		        FormatFolderMapSheet(startingFolder, orientation);
	        }
            catch (Exception ex)
            {
		        MessageBox.Show(ex.Message);
		        throw ex;
	        }
        }

        public static void SaveFolderMap(string startingFolder)
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                string strOutputFile = null;
                try
                {
                    saveFileDialog.FileName = "ExcelFolderMaps.xls";
                    saveFileDialog.InitialDirectory = startingFolder;

                    if (System.Windows.Forms.DialogResult.Cancel == saveFileDialog.ShowDialog())
                    {
                        return;
                    }
                    else
                    {
                        strOutputFile = saveFileDialog.FileName;
                    }
                    if (string.IsNullOrEmpty(strOutputFile))
                    {
                        return;
                    }
                    Globals.ThisAddIn.Application.ActiveWorkbook.SaveAs(strOutputFile);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        public static int GetEndOfSectionDown(int startRow, int startCol, int lastPopulatedRow, int initialColumn)
        {
            Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;

            int functionReturnValue = 0;
            int matchingRow = 0;

            // Search down for a matching cell
            matchingRow = activeSheet.Cells[startRow, startCol].End(XlDirection.xlDown).Row;

            if (startCol == initialColumn)
                //if (startCol == _INITIAL_COL)
            {
                // We have back'd all the way back to the first column.
                // Return either the next matching cell down or the last populated row on the sheet.

                if (matchingRow < lastPopulatedRow)
                {
                    // Section ends on the row prior to the match.
                    functionReturnValue = matchingRow - 1;
                }
                else
                {
                    // Return end of populated section
                    functionReturnValue = lastPopulatedRow;
                }
            }
            else
            {
                if (matchingRow <= lastPopulatedRow)
                {
                    // Back up one column and search down for a populated cell.
                    // Treat row prior to matching row as new end.
                    functionReturnValue = GetEndOfSectionDown(startRow, startCol - 1, matchingRow - 1, initialColumn);
                }
                else
                {
                    // Back up one column and search down for a populated cell.
                    // Treat end of worksheet as end.
                    functionReturnValue = GetEndOfSectionDown(startRow, startCol - 1, lastPopulatedRow, initialColumn);
                }
            }
            return functionReturnValue;
        }


        private static void ListFiles(
            ref string folder, int column, ref int numberFiles, ref long sizeFiles, 
            ref System.DateTime maxLastCreateDate, ref System.DateTime maxLastWriteDate, ref System.DateTime maxLastAccessDate, 
            bool showFiles)
        {
            string file = null;
            int defaultFileColor = (int)System.ConsoleColor.Black;

            string[] files = Directory.GetFiles(folder);
            // Full path names

            numberFiles = 0;
            sizeFiles = 0;

            bool fileAdded = false;

            foreach (string file_loopVariable in files)
            {
                file = file_loopVariable;
                // Is there a better way to use FileInfo?  Initialize just one and reuse?
                try
                {
                    fileAdded = false;

                    FileInfo fileInfo = new FileInfo(file);

                    if (_PatternMatchFileOutput)
                    {
                        if ( ! Regex.Match(fileInfo.Name, _FileMatchPattern).Success)
                        {
                            continue;
                        }
                        else
                        {
                            defaultFileColor = _PatternMatchFileColor;
                        }
                    }

                    numberFiles += 1;
                    sizeFiles += fileInfo.Length;

                    // If there is a local file with a more current date, alert our caller.

                    if (maxLastCreateDate < fileInfo.CreationTime)
                    {
                        maxLastCreateDate = fileInfo.CreationTime;
                    }

                    if (maxLastWriteDate < fileInfo.LastWriteTime)
                    {
                        maxLastWriteDate = fileInfo.LastWriteTime;
                    }

                    if (maxLastAccessDate < fileInfo.LastAccessTime)
                    {
                        maxLastAccessDate = fileInfo.LastAccessTime;
                    }

                    // Note, we add file rows if checking for illegal characters or file name length
                    // even if not listing files.

                    if (_CheckIllegalCharacters)
                    {
                        if (HasIllegalFileNameCharacters(fileInfo.Name))
                        {
                            AddFileRow(fileInfo, column, _FILE_FONT_SIZE, false, _IllegalCharactersColor);
                            fileAdded = true;
                        }
                    }

                    // TODO: This overrides an IllegalCharacters color.  Maybe bold or something else.
                    if (_CheckSharePointFileNameLength)
                    {
                        if (fileInfo.Name.Length > _MaxFileNameLength)
                        {
                            AddFileRow(fileInfo, column, _FILE_FONT_SIZE, false, _IllegalFileNameLengthColor);
                            fileAdded = true;
                        }
                    }

                    if (showFiles & !fileAdded)
                    {
                        AddFileRow(fileInfo, column, _FILE_FONT_SIZE, false, defaultFileColor);
                    }
                }
                catch (PathTooLongException)
                {
                    // TODO: We now know that the current file exists at a path+filename that exceeds
                    // the allowable lengths, however, we don't have the size and date info.
                    // Need to do some extra work to get this.  Not sure it is worth it.  For now
                    // just log that the path is too long.

                    string baseFileName = null;

                    if (_ShowFolders)
                    {
                        // There will be folder information to show where the file is located so just
                        // display the filename.

                        baseFileName = file.Substring(file.LastIndexOf("\\") + 1);
                    }
                    else
                    {
                        // Need to display the full path to the file since no folders are being displayed.

                        baseFileName = file;
                    }

                    //System.IO.Directory.SetCurrentDirectory(folder)
                    //' Get the base filename skipping the last "\"

                    //' This still blows up!
                    //Dim fileInfo As FileInfo = New FileInfo(baseFileName)
                    //sizeFiles += fileInfo.Length

                    string errorInfo = string.Format("{0} - path({1}) + filename({2}) too long({3})", baseFileName, folder.Length, baseFileName.Length, file.Length);
                    AddErrorRow(errorInfo, column, _FILE_FONT_SIZE, false, _PathTooLongColor);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                    throw ex;
                }
            }
        }

        //   This routine calls ListFiles() and then calls
        //   itself recursively to descend the folder hierarchy.

        private static void ListFolders(
            string startingFolder, int column, int fontSize, 
            ref int numberFoldersCummulative, ref int numberFilesCummulative, ref long sizeFilesCummulative, 
            ref System.DateTime maxLastCreateDate, ref System.DateTime maxLastWriteDate, ref System.DateTime maxLastAccessDate, 
            bool showFiles = false, bool showFolders = true)
        {
            //Dim objFld As Scripting.Folder
            int intColMax = 0;
            //Dim innerDir As String

            int numberFoldersCummulativeLocal = 0;
            int numberFilesCummulativeLocal = 0;
            long sizeFilesCummulativeLocal = 0;

            int numberFoldersLocal = 0;
            int numberFilesLocal = 0;
            long sizeFilesLocal = 0;

            FileInfo dirInfo = new FileInfo(startingFolder);
            bool folderAdded = false;

            // Start off with Date.MinValue for dates.
            // The ListFiles methods will further update this with file information.

            System.DateTime localDirCreateDate = System.DateTime.MinValue;
            System.DateTime localDirLastWriteDate = System.DateTime.MinValue;
            System.DateTime localDirLastAccessDate = System.DateTime.MinValue;

            if (_CheckIllegalCharacters)
            {
                if (HasIllegalFolderNameCharacters(dirInfo.Name))
                {
                    AddFolderRow(dirInfo, column, fontSize, _MAKE_BOLD, _IllegalCharactersColor);
                    folderAdded = true;
                }
            }

            if (showFolders & !folderAdded)
            {
                AddFolderRow(dirInfo, column, fontSize, _MAKE_BOLD);
            }

            // Save the current location as we need to come back and add the totals to this row
            // We already bumped _Row (in AddFolderRow) when we added the Folder we are on
            int currentRow = _Row - 1;
            int currentColumn = column;

            try
            {
                // First list the files in the current folder.  
                // Note: We call ListFiles even if not showing files (blnShowFiles is False) 
                // so we can get information about the files to include in the totals.

                ListFiles(ref startingFolder, column + _INDENT_LEVEL, ref numberFilesLocal, ref sizeFilesLocal, ref localDirCreateDate, ref localDirLastWriteDate, ref localDirLastAccessDate, showFiles);

                // Update the dates with the information from the files that were found.
                // The dates will not have changed if there were no local files.  If
                // that is the case use the directory dates.

                if (localDirCreateDate > maxLastCreateDate)
                {
                    maxLastCreateDate = localDirCreateDate;
                }
                else
                {
                    maxLastCreateDate = dirInfo.CreationTime;
                }

                if (localDirLastWriteDate > maxLastWriteDate)
                {
                    maxLastWriteDate = localDirLastWriteDate;
                }
                else
                {
                    maxLastWriteDate = dirInfo.LastWriteTime;
                }

                if (localDirLastAccessDate > maxLastAccessDate)
                {
                    maxLastAccessDate = localDirLastAccessDate;
                }
                else
                {
                    maxLastAccessDate = dirInfo.LastAccessTime;
                }
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow("ListFolders (ListFiles): " + ex.ToString());
                //PLLog.Error(ex, System.Data.Common.PROJECT_NAME);
                throw (ex);
                // So we can add code to catch later.
            }

            string[] dirs = Directory.GetDirectories(startingFolder);

            // Then explore each sub folder

            foreach (string innerDir in dirs)
            {
                try
                {
                    FileInfo innerDirInfo = new FileInfo(innerDir);

                    numberFoldersLocal += 1;

                    int numberFoldersI = 0;
                    int numberFilesI = 0;
                    long sizeFilesI = 0;

                    // Stop if limiting the depth of the exploration

                    if (false == _LimitLevels | (true == _LimitLevels & _LimitLevel > column))
                    {
                        // Call ourselves recursively to display sub folders.

                        localDirCreateDate = System.DateTime.MinValue;
                        localDirLastWriteDate = System.DateTime.MinValue;
                        localDirLastAccessDate = System.DateTime.MinValue;

                        ListFolders(innerDir, column + _INDENT_LEVEL, _FOLDER_FONT_SIZE, ref numberFoldersI, ref numberFilesI, ref sizeFilesI, ref localDirCreateDate, ref localDirLastWriteDate, ref localDirLastAccessDate, showFiles, showFolders);

                        if (localDirCreateDate > maxLastCreateDate)
                        {
                            maxLastCreateDate = localDirCreateDate;
                        }

                        if (localDirLastWriteDate > maxLastWriteDate)
                        {
                            maxLastWriteDate = localDirLastWriteDate;
                        }

                        if (localDirLastAccessDate > maxLastAccessDate)
                        {
                            maxLastAccessDate = localDirLastAccessDate;
                        }

                        numberFoldersCummulativeLocal += numberFoldersI;
                        numberFilesCummulativeLocal += numberFilesI;
                        sizeFilesCummulativeLocal += sizeFilesI;

                        intColMax = column + _INDENT_LEVEL;
                    }
                }
                catch (UnauthorizedAccessException)
                {
                    // Add a message to indicate there is no access to the item previously added.
                    // Move over +1 so grouping isn't impacted.
                    AddErrorRow("<No Access>", column + _INDENT_LEVEL + 1, _FOLDER_FONT_SIZE, false, _NoAccessColor);
                }
                catch (System.IO.PathTooLongException)
                {
                    //Dim indexS As String = File.LastIndexOf("\")
                    //Dim length As Integer = File.Length
                    //Dim errorInfo As String = String.Format("{0} - Path too long ({1})", File.Substring(File.LastIndexOf("\")), File.Length)
                    //AddErrorRow(errorInfo, column, _FILE_FONT_SIZE)
                }
                catch (Exception ex)
                {
                    Common.WriteToDebugWindow("ListFolders (fe Dir): " + ex.ToString());
                    throw (ex);
                    // TODO: PLException.PLApplicationException.Publish(ex)
                }
            }

            // Add the numberFolders, numberFiles, and sizeFiles to the folder row 
            // now that we know stuff about the files in the current folder and below

            numberFoldersCummulativeLocal += numberFoldersLocal;
            numberFilesCummulativeLocal += numberFilesLocal;
            sizeFilesCummulativeLocal += sizeFilesLocal;

            try
            {
                UpdateFolderRow(currentRow, currentColumn, numberFoldersCummulativeLocal, numberFilesCummulativeLocal, sizeFilesCummulativeLocal, numberFoldersLocal, numberFilesLocal, sizeFilesLocal, maxLastCreateDate, maxLastWriteDate,
                maxLastAccessDate, fontSize);
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow("ListFolders (UpdateFolderRow): " + ex.ToString());
                throw (ex);
                // So we can add code to catch later
            }

            numberFoldersCummulative += numberFoldersCummulativeLocal;
            numberFilesCummulative += numberFilesCummulativeLocal;
            sizeFilesCummulative += sizeFilesCummulativeLocal;

            if (_SkipFoldersWithNoFiles)
            {
                if (0 == numberFilesLocal & 0 == numberFilesCummulative)
                {
                    Globals.ThisAddIn.Application.ActiveSheet.Rows(currentRow).Delete();
                    _Row -= 1;
                }
            }

            // Save highest column used
            if (intColMax > _Column)
            {
                _Column = intColMax;
            }
        }

        // Formats the Folder Map Sheet and Page.  Can this be changed to
        // call the common worksheet format thing?

        private static void FormatFolderMapSheet(string strSheetName, XlPageOrientation enumOrientation = XlPageOrientation.xlPortrait)
        {
            int i = 0;
            Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;
            Range formatRange;

            //var _with4 = Globals.ThisAddIn.Application;


            for (i = _INITIAL_COL; i <= _Column; i++)
            {
                activeSheet.Columns[i].ColumnWidth = _COL_WIDTH;
            }

            activeSheet.Columns[_Column + 1].ColumnWidth = _NOTE_WIDTH;

            //activeSheet.Range("A2:I2").Select();

            formatRange = activeSheet.Range["A2:I2"];

            formatRange.HorizontalAlignment = Constants.xlGeneral;
            formatRange.VerticalAlignment = Constants.xlBottom;
            formatRange.WrapText = true;
            formatRange.Orientation = 0;
            formatRange.AddIndent = false;
            formatRange.IndentLevel = 0;
            formatRange.ShrinkToFit = false;
            //_with5.ReadingOrder = Constants.xlContext;
            formatRange.MergeCells = false;
            formatRange.Font.Bold = true;

            activeSheet.Columns["A:A"].ColumnWidth = 13.57;
            activeSheet.Columns["B:B"].ColumnWidth = 13.57;
            activeSheet.Columns["C:C"].ColumnWidth = 13.43;

            activeSheet.Columns["E:E"].ColumnWidth = 6.43;
            activeSheet.Columns["F:F"].ColumnWidth = 7.71;
            activeSheet.Columns["G:G"].ColumnWidth = 13.57;
            activeSheet.Columns["H:H"].ColumnWidth = 13.57;
            activeSheet.Columns["I:I"].ColumnWidth = 13.43;

            //_with4.Range["K2"].Select();

            formatRange = activeSheet.Range["K2"];

            formatRange.Font.Name = "Arial";
            formatRange.Font.Size = 12;
            formatRange.Font.Strikethrough = false;
            formatRange.Font.Superscript = false;
            formatRange.Font.Subscript = false;
            formatRange.Font.OutlineFont = false;
            formatRange.Font.Shadow = false;
            formatRange.Font.Underline = XlUnderlineStyle.xlUnderlineStyleNone;
            formatRange.Font.ColorIndex = Constants.xlAutomatic;
            formatRange.Font.Bold = true;

            activeSheet.Columns["A:J"].Group();
            activeSheet.Columns["A:D"].Group();
            activeSheet.Columns["G:I"].Group();

            activeSheet.Range["A1"].Select();

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dirInfo"></param>
        /// <param name="column"></param>
        /// <param name="fontSize"></param>
        /// <param name="makeBold"></param>
        /// <param name="fontColor"></param>
        private static void AddFolderRow(FileInfo dirInfo, int column = 1, int fontSize = 10, bool makeBold = false, int fontColor = (int)System.ConsoleColor.Black)
        {
            string strS = null;

            if (true == _GroupResults)
            {
                if (_GroupLevel == column)
                {
                    // Mark starting point.  We may need it later for grouping.  Don't reset.
                    if (0 == _GroupStartRow)
                    {
                        _GroupStartRow = _Row;
                    }
                }
            }

            if (_PatternMatchFolderHighlight)
            {
                if ((Regex.Match(dirInfo.Name, _FolderMatchPattern).Success))
                {
                    fontColor = _FolderMatchColor;
                }
            }

            Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;

            var folderRow = activeSheet.Cells[_Row, column];
            // If we are not displaying folders but for some reason
            // are calling AddFolderRow(), then use full path.

            if (_ShowFolders)
            {
                folderRow.Value = string.Format("{0}\\", dirInfo.Name);
            }
            else
            {
                folderRow.Value = string.Format("{0}\\", dirInfo.FullName);
            }

            folderRow.Font.Bold = makeBold;
            folderRow.Font.Size = fontSize;
            folderRow.Font.Color = fontColor;

            // Now check if we need to do any grouping.

            if (_GroupStartRow > 0)
            {
                // We have been adding rows to a possible grouping set.
                if (_GroupLevel > column)
                {
                    // We have transitioned up the chain above the grouping set.
                    strS = string.Format("{0}:{1}", _GroupStartRow, _Row - 1);
                    //strS = _GroupStartRow + ":" + _Row - 1;
                    Globals.ThisAddIn.Application.ActiveSheet.Rows(strS).Group();
                    // Reset the grouping counter.
                    _GroupStartRow = 0;
                }
            }

            // Next row to add content
            _Row += 1;
        }

        //   Adds a line to the spreadsheet indented at the appropriate level.

        private static void AddFileRow(FileInfo fileInfo, int column = 1, int fontSize = 10, bool makeBold = false, int fontColor = (int)System.ConsoleColor.Black)
        {
            Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;
            Range formatRange;

            string strS = null;

            if (true == _GroupResults)
            {
                if (_GroupLevel == column)
                {
                    // Mark starting point.  We may need it later for grouping.  Don't reset.
                    if (0 == _GroupStartRow)
                    {
                        _GroupStartRow = _Row;
                    }
                }
            }

            formatRange = activeSheet.Cells[_Row, column];
            // If we are not displaying folders but for some reason 
            // are calling AddFileRow(), then use full path.

            if (_ShowFolders)
            {
                formatRange.Value = string.Format(" - {0}", fileInfo.Name);
            }
            else
            {
                formatRange.Value = string.Format(" - {0}", fileInfo.FullName);
            }

            formatRange.Font.Bold = makeBold;
            formatRange.Font.Size = fontSize;
            //.ColorIndex = fontColor
            formatRange.Font.Color = fontColor;

            // Start at _FILE_INFO_COL + 1 as we don't display file count on file row.

            formatRange = activeSheet.Cells[_Row, _FILE_INFO_COL + 1];
            formatRange.Value = fileInfo.Length;
            formatRange.NumberFormat = "#,##0_);(#,##0)";

            formatRange.Font.Bold = makeBold;
            formatRange.Font.Size = fontSize;

            Range rng = default(Microsoft.Office.Interop.Excel.Range);
            System.DateTime dateD = default(System.DateTime);

            rng = activeSheet.Cells[_Row, _FILE_INFO_COL + 2];
            dateD = fileInfo.CreationTime;
            rng.Value = dateD;

            ColorCodeDate(rng, dateD, false, fontSize, _DateType.LastCreate);

            rng = activeSheet.Cells[_Row, _FILE_INFO_COL + 3];
            dateD = fileInfo.LastWriteTime;
            rng.Value = dateD;

            ColorCodeDate(rng, dateD, false, fontSize, _DateType.LastWrite);

            rng = activeSheet.Cells[_Row, _FILE_INFO_COL + 4];
            dateD = fileInfo.LastAccessTime;
            rng.Value = dateD;

            ColorCodeDate(rng, dateD, false, fontSize, _DateType.LastAccess);

            // Now check if we need to do any grouping.

            if (_GroupStartRow > 0)
            {
                // We have been adding rows to a possible grouping set.
                if (_GroupLevel > column)
                {
                    // We have transitioned up the chain above the grouping set.
                    //strS = _GroupStartRow + ":" + _Row - 1;
                    strS = string.Format("{0}:{1}", _GroupStartRow, _Row - 1);
                    activeSheet.Rows[strS].Group();
                    // Reset the grouping counter.
                    _GroupStartRow = 0;
                }
            }

            // Next row to add content.
            _Row = _Row + 1;
        }

        private static void AddErrorRow(string errorInfo, int column = 1, int fontSize = 10, bool makeBold = false, int fontColor = (int)System.ConsoleColor.Black)
        {
            Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;
            Range formatRange;

            string strS = null;

            //    Debug.Print m_lngRow & " " & intCol & ":" & m_intGroupStartRow & " " & strText

            if (true == _GroupResults)
            {
                if (_GroupLevel == column)
                {
                    // Mark starting point.  We may need it later for grouping.  Don't reset.
                    if (0 == _GroupStartRow)
                    {
                        _GroupStartRow = _Row;
                    }
                }
            }

            var _with15 = Globals.ThisAddIn.Application;
            formatRange = activeSheet.Cells[_Row, column];
            formatRange.Value = string.Format(" - {0}", errorInfo);

            formatRange.Font.Bold = makeBold;
            formatRange.Font.Size = fontSize;
            formatRange.Font.Color = fontColor;

            // Now check if we need to do any grouping.

            if (_GroupStartRow >  0)
            {
                // We have been adding rows to a possible grouping set.
                if (_GroupLevel > column)
                {
                    // We have transitioned up the chain above the grouping set.
                    //strS = _GroupStartRow + ":" + _Row - 1;
                    strS = string.Format("{0}:{1}", _GroupStartRow, _Row - 1);
                    activeSheet.Rows[strS].Group();
                    // Reset the grouping counter.
                    _GroupStartRow = 0;
                }
            }

            // Next row to add content.
            _Row = _Row + 1;
        }

        private static void UpdateFolderRow(int row, int column, int numberFoldersCummulative, int numberFilesCummulative, long sizeFilesCummulative, int numberFolders, int numberFiles, long sizeFiles, System.DateTime maxLastCreateDate, System.DateTime maxLastWriteDate,
        System.DateTime maxLastAccessDate, int fontSize = 10)
        {
            Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;
            Range formatRange;

            var excelApp = Globals.ThisAddIn.Application;

            formatRange = activeSheet.Cells[row, _FOLDER_INFO_COL];
            formatRange.Value = numberFoldersCummulative;
            formatRange.NumberFormat = "#,##0_);(#,##0)";
            formatRange.Font.Size = fontSize;
            // .Bold = blnBold

            formatRange = activeSheet.Cells[row, _FOLDER_INFO_COL + 1];
            formatRange.Value = numberFilesCummulative;
            formatRange.NumberFormat = "#,##0_);(#,##0)";
            formatRange.Font.Size = fontSize;

            formatRange = activeSheet.Cells[row, _FOLDER_INFO_COL + 2];
            formatRange.Value = sizeFilesCummulative;
            formatRange.NumberFormat = "#,##0_);(#,##0)";
            formatRange.Font.Size = fontSize;

            formatRange = activeSheet.Cells[row, _FILE_INFO_COL];
            formatRange.Value = numberFiles;
            formatRange.NumberFormat = "#,##0_);(#,##0)";
            formatRange.Font.Size = fontSize;

            formatRange = activeSheet.Cells[row, _FILE_INFO_COL + 1];
            formatRange.Value = sizeFiles;
            formatRange.NumberFormat = "#,##0_);(#,##0)";
            formatRange.Font.Size = fontSize;

            System.DateTime dateD = default(System.DateTime);

            formatRange = activeSheet.Cells[row, _FILE_INFO_COL + 2];
            dateD = maxLastCreateDate;
            formatRange.Value = dateD;

            ColorCodeDate(formatRange, dateD, false, fontSize, _DateType.LastCreate);

            formatRange = activeSheet.Cells[row, _FILE_INFO_COL + 3];
            dateD = maxLastWriteDate;
            formatRange.Value = dateD;

            ColorCodeDate(formatRange, dateD, false, fontSize, _DateType.LastWrite);

            formatRange = activeSheet.Cells[row, _FILE_INFO_COL + 4];
            dateD = maxLastAccessDate;
            formatRange.Value = dateD;

            ColorCodeDate(formatRange, dateD, false, fontSize, _DateType.LastAccess);
        }

        private static void ColorCodeDate(Range formatRange, DateTime dt, bool makeBold, int fontSize, _DateType dateType)
        {
            formatRange.Font.Bold = makeBold;
            formatRange.Font.Size = fontSize;

            if (_ColorCodeDates)
            {
                switch (dateType)
                {
                    case _DateType.LastCreate:
                        if ((DateTime.Compare(DateTime.Now, dt.AddMonths(_MonthsSinceCreated)) > 0))
                        {
                            formatRange.Font.Color = _MonthsCreatedColor;
                        }
                        else
                        {
                            formatRange.Font.Color = _MonthsDefaultColor;
                        }

                        break;

                    case _DateType.LastWrite:
                        if ((DateTime.Compare(DateTime.Now, dt.AddMonths(_MonthsSinceWritten)) > 0))
                        {
                            formatRange.Font.Color = _MonthsWrittenColor;
                        }
                        else
                        {
                            formatRange.Font.Color = _MonthsDefaultColor;
                        }

                        break;

                    case _DateType.LastAccess:
                        if ((DateTime.Compare(DateTime.Now, dt.AddMonths(_MonthsSinceAccessed)) > 0))
                        {
                            formatRange.Font.Color = _MonthsAccessedColor;
                        }
                        else
                        {
                            formatRange.Font.Color = _MonthsDefaultColor;
                        }

                        break;
                }
            }
        }

        private static bool HasIllegalFileNameCharacters(string name)
        {
            Match illegalCharactersMatch = Regex.Match(name, _IllegalFileCharacters);

            return illegalCharactersMatch.Success;
        }

        private static bool HasIllegalFolderNameCharacters(string name)
        {
            Match illegalCharactersMatch = Regex.Match(name, _IllegalFolderCharacters);

            return illegalCharactersMatch.Success;
        }

    }
}
