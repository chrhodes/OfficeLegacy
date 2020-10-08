using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using Microsoft.Office.Interop.Excel;
using ExcelHlp = AddinHelper.Excel;

namespace SupportTools_Excel.Actions
{
    class Excel_PageFormatting
    {
        public static void AddHeader()
        {
            Globals.ThisAddIn.Application.PrintCommunication = false;

            Workbook workBook;
            //Worksheet workSheet;
            StringBuilder sb = new StringBuilder();

            try
            {
                workBook = Globals.ThisAddIn.Application.ActiveWorkbook;
                ExcelHlp.CalculationsOff();

                foreach (Worksheet workSheet in workBook.Sheets)
                {
                    sb.Length = 0;

                    sb.AppendFormat("&A");
                    //sb.Append("&5&Z&F");    // Five point font, path, and filename
                    //sb.AppendFormat("\nCreated: {0} By: {1}", 
                    //    ExcelHlp.GetBuiltInPropertyValue(workBook, "Creation Date"),
                    //    ExcelHlp.GetBuiltInPropertyValue(workBook, "Author"));
                    //sb.AppendFormat("\nLast Saved: {0} By: {1}",
                    //    ExcelHlp.GetBuiltInPropertyValue(workBook, "Last Save Time"),
                    //    ExcelHlp.GetBuiltInPropertyValue(workBook, "Last author"));
                    //sb.AppendFormat("\nLast Printed: {0}",
                    //    ExcelHlp.GetBuiltInPropertyValue(workBook, "Last Print Date"));

                    //if (Common.DebugMode)
                    //{
                    //    Common.WriteToDebugWindow("LeftFooter:>" + sb.ToString() + "<");                    	;
                    //}

                    //workSheet.PageSetup.LeftHeader = @"&10 Left &A";
                    workSheet.PageSetup.LeftHeader = "";

                    workSheet.PageSetup.CenterHeader = workSheet.Name;

                    sb.Length = 0;

                    //sb.Append("&5&P - &N"); // Five point font, page of pages
                    //sb.Append("\nTitle :");
                    //sb.Append(ExcelHlp.GetBuiltInPropertyValue(workBook, "Title"));
                    //sb.Append("\nSubject: ");
                    //sb.Append(ExcelHlp.GetBuiltInPropertyValue(workBook, "Subject"));

                    //if (Common.DebugMode)
                    //{
                    //    Common.WriteToDebugWindow("RightFooter:>" + sb.ToString() + "<");                    	;
                    //}

                    workSheet.PageSetup.RightHeader = "";
                    //workSheet.PageSetup.RightHeader = @"&5 Right";

                    // Indicate we have added a custom footer.
                    // This gets checked in the BeforeClose event.

                    //if ( ! ExcelHlp.HasCustomFooter())
                    //{
                    //    ExcelHlp.CustomFooterExists(true);
                    //}
                }

                ExcelHlp.CalculationsOn();
            }
            catch (Exception ex)
            {
                MessageBox.Show("AddHeader:" + ex.ToString());
            }

            Globals.ThisAddIn.Application.PrintCommunication = true;
        }

        public static void AddFooter()
        {
            Globals.ThisAddIn.Application.PrintCommunication = false;

            Workbook workBook;
            //Worksheet workSheet;
            StringBuilder sb = new StringBuilder();

            try
            {
                workBook = Globals.ThisAddIn.Application.ActiveWorkbook;
                ExcelHlp.CalculationsOff();

                foreach(Worksheet workSheet in workBook.Sheets)
                {
                    sb.Length = 0;

                    sb.Append("&5&Z&F");    // Five point font, path, and filename
                    sb.AppendFormat("\nCreated: {0} By: {1}",
                        ExcelHlp.GetBuiltInPropertyValue(workBook, "Creation Date"),
                        ExcelHlp.GetBuiltInPropertyValue(workBook, "Author"));
                    sb.AppendFormat("\nLast Saved: {0} By: {1}",
                        ExcelHlp.GetBuiltInPropertyValue(workBook, "Last Save Time"),
                        ExcelHlp.GetBuiltInPropertyValue(workBook, "Last author"));
                    sb.AppendFormat("\nLast Printed: {0}",
                        ExcelHlp.GetBuiltInPropertyValue(workBook, "Last Print Date"));

                    if(Common.DebugMode)
                    {
                        Common.WriteToDebugWindow("LeftFooter:>" + sb.ToString() + "<"); ;
                    }

                    workSheet.PageSetup.LeftFooter = sb.ToString();

                    workSheet.PageSetup.CenterFooter = "";

                    sb.Length = 0;

                    sb.Append("&5&P - &N"); // Five point font, page of pages
                    sb.Append("\nTitle :");
                    sb.Append(ExcelHlp.GetBuiltInPropertyValue(workBook, "Title"));
                    sb.Append("\nSubject: ");
                    sb.Append(ExcelHlp.GetBuiltInPropertyValue(workBook, "Subject"));

                    if(Common.DebugMode)
                    {
                        Common.WriteToDebugWindow("RightFooter:>" + sb.ToString() + "<"); ;
                    }

                    workSheet.PageSetup.RightFooter = sb.ToString();

                    // Indicate we have added a custom footer.
                    // This gets checked in the BeforeClose event.

                    if(!ExcelHlp.HasCustomFooter())
                    {
                        ExcelHlp.CustomFooterExists(true);
                    }
                }

                ExcelHlp.CalculationsOn();
            }
            catch(Exception ex)
            {
                MessageBox.Show("AddFooter:" + ex.ToString());
            }

            Globals.ThisAddIn.Application.PrintCommunication = true;
        }

    }
}