using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop;
using ExcelHlp = AddinHelper.Excel;

namespace SupportTools_Excel.Actions
{
    class Excel_TableOfContents
    {
        public static void CreateTableOfContents()
        {
            Worksheet tableOfContents;
            int row = 3;        // Starting Row
            int column = 1;     // Starting Column
            Boolean currentSheetProtectionMode;
            Workbook activeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
           
            try
            {
                tableOfContents = activeWorkbook.Sheets["Table of Contents"];
            }
            catch(Exception)
            {
                tableOfContents = ExcelHlp.NewWorksheet("Table of Contents", beforeSheetName: "FIRST");
                //tableOfContents = activeWorkbook.Sheets.Add();
                //tableOfContents.Name = "Table of Contents";
            }

            tableOfContents.Columns["A:A"].ClearContents();

            foreach(Worksheet sheet in activeWorkbook.Sheets)
            {
                switch(sheet.Name)
                {
                    case "Table of Contents":
                        break;

                    default:
                        // Unprotect the sheet before adding the hyperlink
                        currentSheetProtectionMode = ExcelHlp.ProtectSheet(sheet, false);
                        sheet.Cells[1, 1].Value = "Table of Contents";
                        sheet.Hyperlinks.Add(Anchor: sheet.Cells[1, 1], Address: "", SubAddress: "'" + tableOfContents.Name + "'!A1", TextToDisplay: tableOfContents.Name);
                        // Then restore the setting
                        ExcelHlp.ProtectSheet(sheet, currentSheetProtectionMode);

                        // Now update the Table of Contents Sheet
                        tableOfContents.Cells[row, column].Value = sheet.Name;
                        tableOfContents.Hyperlinks.Add(Anchor: tableOfContents.Cells[row, column], Address: "", SubAddress: "'" + sheet.Name + "'!A1", TextToDisplay: sheet.Name);

                        row += 1;
                        break;
                }
            }

            tableOfContents.Columns["A:A"].EntireColumn.AutoFit();

            if(!ExcelHlp.HasCustomTableOfContents())
            {
                ExcelHlp.CustomTableOfContentsExists(true);
            }

        }
    }
}
