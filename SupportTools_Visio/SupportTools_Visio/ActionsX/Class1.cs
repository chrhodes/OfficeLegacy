using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Xml.Linq;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using Visio = Microsoft.Office.Interop.Visio;
using System.Data;

namespace SupportTools_Visio.Actions
{
    public static class Class1
    {
        public static void DoSomething()
        {
            //Dim UndoScopeID1 As Long
            //UndoScopeID1 = Application.BeginUndoScope("Insert Page")

            int undoScope = Globals.ThisAddIn.Application.BeginUndoScope("GenerateTOCPage");

            //Dim vsoPage1 As Visio.Page
            //Set vsoPage1 = ActiveDocument.Pages.Add

            Visio.Page tocPage = Globals.ThisAddIn.Application.ActiveDocument.Pages.Add();

            //vsoPage1.Name = "Table of Contents"
            //vsoPage1.Background = False
            //vsoPage1.Index = 1

            tocPage.Name = "Table of Contents";
            tocPage.Background = 0;
            tocPage.Index = 1;

            //vsoPage1.PageSheet.CellsSRC(visSectionObject, visRowPage, visPageWidth).FormulaU = "11 in"
            //vsoPage1.PageSheet.CellsSRC(visSectionObject, visRowPage, visPageHeight).FormulaU = "8.5 in"
            //vsoPage1.PageSheet.CellsSRC(visSectionObject, visRowPage, 38).FormulaU = "0"
            //vsoPage1.PageSheet.CellsSRC(visSectionObject, visRowPageLayout, visPLOSplit).FormulaForceU = "1"
            //vsoPage1.PageSheet.CellsSRC(visSectionObject, visRowPrintProperties, visPrintPropertiesPageOrientation).FormulaU = "2"
            //vsoPage1.PageSheet.CellsSRC(visSectionUser, 0, visUserValue).FormulaForceU = ""


            //Application.EndUndoScope UndoScopeID1, True

            Globals.ThisAddIn.Application.EndUndoScope(undoScope, true);
        }

        private static string URLQueryString(Visio.Shape shape)
        {
            string result;

            switch ((int)shape.Cells["Prop.AbstractionID"].ResultIU)
            {
                case Common.ABSTRACTION_RELATION:
                    // Relation Shape Concatenation
                    Visio.Shape beginShape = shape.Connects[1].ToCell.Shape;
                    result = string.Format("?id={0}&type={1}&relation=",
                        beginShape.Cells["Prop.ShapeID"].ResultIU.ToString(),
                        beginShape.Cells["Prop.ShapeTypeID"].ResultIU.ToString(),
                        shape.Cells["Prop.ShapeID"].ResultIU.ToString());

                    break;

                default :
                    // Item Shape Concatenation
                    result = string.Format("?type={0}&id={1}",
                        shape.Cells["Prop.ShapeTypeID"].ResultIU.ToString(),
                        shape.Cells["Prop.ShapeID"].ResultIU.ToString());
                    break;
            }

            return result;
        }

        private static bool ShapeExists(int key, int type)
        {
            try
            {
                string shapeName = String.Format("{0}.{1}", SetMaster(type), key);
                Visio.Shape shape = Globals.ThisAddIn.Application.ActivePage.Shapes[shapeName];

                if (null != shape)
                {
                    // Just to double-check for different relation types
                    if (shape.Cells["Prop.ShapeTypeID"].ResultIU == type)
                    {
                        // This is a duplicate
                    	return true;
                    }
                    else
                    {
                        // THis is a differnt relationship
                        return false;
                    }
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                // This is not a duplicate
                return false;                
            }
        }

        private static string SetMaster(int shapeType)
        {
            string result = "Unknown";

            switch (shapeType)
            {
                case Common.SHAPETYPE_ACTOR:
                    result = "Actor";
                    break;

                case Common.SHAPETYPE_APPLICATION:
                    result = "Application";
                    break;

                case Common.SHAPETYPE_BATCHFILE:
                    result = "Batch File";
                    break;

                case Common.SHAPETYPE_BROWSER:
                    result = "Browser";
                    break;

                case Common.SHAPETYPE_BUSINESS:
                    result = "Business Process";
                    break;

                case Common.SHAPETYPE_BUSINESSRELATION:
                    result = "Relation";
                    break;

                case Common.SHAPETYPE_CROSSRELATION:
                    result = "Relation";
                    break;

                case Common.SHAPETYPE_DATA:
                    result = "Data";
                    break;

                case Common.SHAPETYPE_DECISION:
                    result = "Decision";
                    break;

                case Common.SHAPETYPE_DESKTOP:
                    result = "Desktop";
                    break;

                case Common.SHAPETYPE_DEVICERELATION:
                    result = "Relation";
                    break;

                case Common.SHAPETYPE_END:
                    result = "End";
                    break;

                case Common.SHAPETYPE_FILE:
                    result = "File";
                    break;

                case Common.SHAPETYPE_FILEREPORT:
                    result = "File/Report";
                    break;

                case Common.SHAPETYPE_FIREWALL:
                    result = "Firewall";
                    break;

                case Common.SHAPETYPE_FOLDER:
                    result = "Folder";
                    break;

                case Common.SHAPETYPE_IP:
                    result = "IP";
                    break;

                case Common.SHAPETYPE_JOB:
                    result = "Job";
                    break;

                case Common.SHAPETYPE_LAPTOP:
                    result = "Laptop";
                    break;

                case Common.SHAPETYPE_LOADBALANCER:
                    result = "LoadBalancer";
                    break;

                case Common.SHAPETYPE_MAINFRAME:
                    result = "Mainframe";
                    break;

                case Common.SHAPETYPE_OBJECTRELATION:
                    result = "Relation";
                    break;

                case Common.SHAPETYPE_PDA:
                    result = "PDA";
                    break;

                case Common.SHAPETYPE_PIECE:
                    result = "Piece";
                    break;

                case Common.SHAPETYPE_PRINTER:
                    result = "Printer";
                    break;

                case Common.SHAPETYPE_QUEUE:
                    result = "Queue";
                    break;

                case Common.SHAPETYPE_ROUTER:
                    result = "Router";
                    break;

                case Common.SHAPETYPE_SAN:
                    result = "SAN";
                    break;

                case Common.SHAPETYPE_SERVER:
                    result = "Server";
                    break;

                case Common.SHAPETYPE_SQLINSTANCE:
                    result = "SQL Instance";
                    break;

                case Common.SHAPETYPE_START:
                    result = "Start";
                    break;

                case Common.SHAPETYPE_STEP:
                    result = "Process Step";
                    break;

                case Common.SHAPETYPE_STEPRELATION:
                    result = "Relation";
                    break;

                case Common.SHAPETYPE_SUBSTEP:
                    result = "SubStep";
                    break;

                case Common.SHAPETYPE_SYSTEM:
                    result = "System";
                    break;

                case Common.SHAPETYPE_SYSTEMRELATION:
                    result = "Relation";
                    break;

                case Common.SHAPETYPE_TABLET:
                    result = "Tablet";
                    break;

                case Common.SHAPETYPE_TIME:
                    result = "Clock";
                    break;

            }

            return result;
        }

        // TODO: FIgure out what ListView is
        private static DataTable CreateTable_Assocs(ListView shapes)
        {
            DataTable table = new DataTable("Data");

            table.Columns.Add(new DataColumn("Assoc_TypeID", Type.GetType("System.Int32")));
            table.Columns.Add(new DataColumn("Assoc_ID", Type.GetType("System.Int32")));
            table.Columns.Add(new DataColumn("Assoc_Name", Type.GetType("System.String")));
            table.Columns.Add(new DataColumn("Assoc_Desc", Type.GetType("System.String")));
            table.Columns.Add(new DataColumn("Assoc_AbstractionID", Type.GetType("System.Int32")));
            table.Columns.Add(new DataColumn("Begin_TableID", Type.GetType("System.Int32")));
            table.Columns.Add(new DataColumn("Begin_TypeID", Type.GetType("System.Int32")));
            table.Columns.Add(new DataColumn("Begin_ID", Type.GetType("System.Int32")));
            table.Columns.Add(new DataColumn("End_TableID", Type.GetType("System.Int32")));
            table.Columns.Add(new DataColumn("End_TypeID", Type.GetType("System.Int32")));
            table.Columns.Add(new DataColumn("End_ID", Type.GetType("System.Int32")));

            table.AcceptChanges();

            for (int i = 0; i < shapes.SelectedItems.Count - 1; i++)
            {
                DataRow row = table.NewRow();
                row["Assoc_TypeID"] = shapes.SelectedItems[i];
            }

            return table;
        }
    }
}
