using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Text;

namespace AddinHelper
{
    public class Common
    {
        public const string TAG_PREFIX = "VNC";
        public const string PROJECT_NAME = "AddInHelper";

        public static Boolean HasAppEvents = true;  // Custom Header and Footer need this enabled.
        public static Boolean DisplayEvents = false;

        public static Boolean DebugSQL
        {
            get;
            set;
        }
        public static Boolean DebugLevel1
        {
            get;
            set;
        }
        public static Boolean DebugLevel2
        {
            get;
            set;
        }
        public static Boolean DebugMode
        {
            get;
            set;
        }
        public static Boolean DeveloperMode
        {
            get;
            set;
        }

        private static AddinHelper.User_Interface.Forms.frmDebugWindow _DebugWindow;
        public static AddinHelper.User_Interface.Forms.frmDebugWindow DebugWindow
        {
            get
            {
                if (_DebugWindow == null)
                {
                	_DebugWindow = new User_Interface.Forms.frmDebugWindow();
                }

                return _DebugWindow;
            }
            set
            {
                _DebugWindow = value;
            }
        }

        private static AddinHelper.User_Interface.Forms.frmWatchWindow _WatchWindow;
        public static AddinHelper.User_Interface.Forms.frmWatchWindow WatchWindow
        {
            get
            {
                if (_WatchWindow == null)
                {
                	_WatchWindow = new User_Interface.Forms.frmWatchWindow();
                }
                return _WatchWindow;
            }
            set
            {
                _WatchWindow = value;
            }
        }

        private static void DisplayDataSet(DataSet dataSet)
        {
            DisplayTables(dataSet.Tables);
        }

        private static void DisplayTables(DataTableCollection tables)
        {
            foreach(DataTable table in tables)
            {
                WriteToDebugWindow(string.Format("Table:   >{0}<", table.TableName));
                WriteToDebugWindow("Columns:");

                foreach(DataColumn column in table.Columns)
                {
                    WriteToDebugWindow(string.Format(" >{0}<", column.ColumnName));
                }

                WriteToDebugWindow("");
                WriteToDebugWindow(string.Format("Rows:{0}", Environment.NewLine));

                foreach(DataRow row in table.Rows)
                {
                    foreach(DataColumn column in table.Columns)
                    {
                        WriteToDebugWindow(string.Format(" >{0}<", row[column.ColumnName]));
                    }
                    WriteToDebugWindow("");
                }
            }
        }

        public static long WriteToWatchWindow(string message)
        {
            if (DeveloperMode)
            {
            	WatchWindow.AddOutputLine(message);
            }

            return Stopwatch.GetTimestamp();
        }

        public static long WriteToWatchWindow(string message, long startTicks)
        {
            if(DeveloperMode)
            {
                WatchWindow.AddOutputLine(message + "-" + (Stopwatch.GetTimestamp() - startTicks) / Stopwatch.Frequency);
            }
                    

            return Stopwatch.GetTimestamp();
        }

        public static long WriteToDebugWindow(string message)
        {
            if (DeveloperMode)
            {
            	DebugWindow.AddOutputLine(message);
            }

            return Stopwatch.GetTimestamp();
        }

        public static long WriteToDebugWindow(string message, long startTicks)
        {
            
            if(DeveloperMode)
            {
                DebugWindow.AddOutputLine(message + "-" + (Stopwatch.GetTimestamp() - startTicks) / Stopwatch.Frequency);
            }

            return Stopwatch.GetTimestamp();
        }
    }
}
