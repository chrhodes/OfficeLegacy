using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SMO = Microsoft.SqlServer.Management.Smo;

namespace SMOHelper
{
    public class Table
    {
        #region Initialization

        /// <summary>
        /// Initializes a new instance of the Table class.
        /// </summary>
        public Table(SMO.Table table)
        {
            AddinHelper.Common.WriteToDebugWindow(string.Format("SMOH.{0}({1})", "Table", table));
            // Save the real table in case we need to get something from it.

            _table = table;

            CreateDate = table.CreateDate.ToString("yyyy-MM-dd hh:mm:ss");

            try
            {
                DateLastModified = table.DateLastModified.ToString("yyyy-MM-dd hh:mm:ss");
            }
            catch(Exception )
            {
                DateLastModified = "Not Available";
            }

            DataSpaceUsed = table.DataSpaceUsed.ToString();
            ID = table.ID.ToString(); ;
            Name = table.Name;
            Owner = table.Owner;
            RowCount = table.RowCount.ToString();

            try
            {
                ExtendedProperties = table.ExtendedProperties;
            }
            catch (Exception )
            {

            }
        }

        #endregion
        
        public SMO.Table _table;

        public string CreateDate;
        public string DataSpaceUsed;
        public string DateLastModified;
        public string ID;
        public string Name;
        public string Owner;
        public string RowCount;

        private Dictionary<string, Column> _Columns;
        public Dictionary<string, Column> Columns
        {
            get
            {
                if (null == _Columns)
                {
                    _Columns = new Dictionary<string, Column>();

                    foreach (SMO.Column column in _table.Columns)
                    {
                        SMOHelper.Column col = new SMOHelper.Column(column);

                        _Columns.Add(column.Name, col);
                    };
                }

                return _Columns; 
            }
            set
            {
                _Columns = value;
            }
        }

        private SMO.ExtendedPropertyCollection _ExtendedProperties;
        public SMO.ExtendedPropertyCollection ExtendedProperties
        {
            get { return _ExtendedProperties; }
            set { _ExtendedProperties = value; }
        }
    }
}
