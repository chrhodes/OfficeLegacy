using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using Visio=Microsoft.Office.Interop.Visio;
//using Smarts3;

namespace SupportTools_Visio.User_Interface.Forms
{
    public partial class frmRetrieveShape : Form
    {
        public frmRetrieveShape()
        {
            InitializeComponent();
        }

        private void frmRetrieveShape_Load(object sender, EventArgs e)
        {
            int abstractionID = 0;
            int shapeTypeID = 0;

            Visio.Application app = Globals.ThisAddIn.Application;
            var visSelection = app.ActiveWindow.Selection;

            if (visSelection.Count == 0)
            {
                string message = string.Format("{0}\n{1}",
                    System.Reflection.MethodInfo.GetCurrentMethod().Name,
                    "No Shape Selected");

            	MessageBox.Show(message);
            }
            else
            {
                var visShape = visSelection[1];
                //var visShape2 = visSelection[0];

                abstractionID = (int)visShape.Cells["Prop.AbstractionID"].ResultIU;
                shapeTypeID = (int)visShape.Cells["Prop.ShapeTypeID"].ResultIU;
            }

            string cs = Config.SmartsDBConnection;

            //Smarts3.Data.DAL dal = new Smarts3.Data.DAL(cs);

            //Smarts3.Data.ApplicationDS.ShapeListDataTable dataTable = dal.GetShapeList_Browser(shapeTypeID);

            User_Controls_WPF.wucRetrieveShape wuc = (User_Controls_WPF.wucRetrieveShape)elementHost1.Child;

            //wuc.lbItems.DataContext = dataTable;
            //wuc.lvItems.DataContext = dataTable;
            //wuc.gcItems.DataContext = dataTable;

            //using (SqlConnection sqlConn = new SqlConnection(Config.SmartsDBConnection))
            //{
            //    sqlConn.Open();

            //    DataTable sqlTable = new DataTable();
            //    //DataTable sqlTable = GetShapeListTable();

            //    SqlDataAdapter dataAdapter = new SqlDataAdapter();
            //    string sql = string.Format("EXEC GetShapeList_Browser {0}", shapeTypeID);
                    
            //    using (SqlCommand sqlCmd = new SqlCommand(sql, sqlConn))
            //    {
            //        sqlTable = sqlCmd.ExecuteDataTable(sqlConn);

            //        User_Controls_WPF.wucRetrieveShape wuc = (User_Controls_WPF.wucRetrieveShape)elementHost1.Child;

            //        wuc.lbItems.DataContext = sqlTable;
            //        wuc.lvItems.DataContext = sqlTable;
            //        wuc.gcItems.DataContext = sqlTable;
            //        //wuc.lbItems.ItemsSource = sqlTable.Rows;
            //        //wuc.lbItems.DisplayMemberPath = "Display_Name";
            //    }

            //}

        }

        private DataTable GetShapeListTable()
        {
            DataTable table = new DataTable();
            DataColumn column;

            column = new DataColumn("Item_Table");
            column.DataType = System.Type.GetType("System.String");
            table.Columns.Add(column);

            column = new DataColumn("Item_ID");
            column.DataType = System.Type.GetType("System.String");
            table.Columns.Add(column);

            column = new DataColumn("Item_Name");
            column.DataType = System.Type.GetType("System.String");
            table.Columns.Add(column);

            column = new DataColumn("Container_ID");
            column.DataType = System.Type.GetType("System.String");
            table.Columns.Add(column);

            column = new DataColumn("Container_TypeID");
            column.DataType = System.Type.GetType("System.String");
            table.Columns.Add(column);

            column = new DataColumn("Container_Name");
            column.DataType = System.Type.GetType("System.String");
            table.Columns.Add(column);

            column = new DataColumn("Item_Desc");
            column.DataType = System.Type.GetType("System.String");
            table.Columns.Add(column);

            column = new DataColumn("Item_TypeID");
            column.DataType = System.Type.GetType("System.String");
            table.Columns.Add(column);

            column = new DataColumn("Item_TypeName");
            column.DataType = System.Type.GetType("System.String");
            table.Columns.Add(column);

            column = new DataColumn("Display_Name");
            column.Caption = "Display_Name";
            column.DataType = System.Type.GetType("System.String");
            table.Columns.Add(column);

            return table;
        }
    }
}
