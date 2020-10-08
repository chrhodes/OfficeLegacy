using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SMO = Microsoft.SqlServer.Management.Smo;

namespace SMOHelper
{
    public class Column
    {
        #region Initialization

        /// <summary>
        /// Initializes a new instance of the Column class.
        /// </summary>
        public Column(SMO.Column column)
        {
            AddinHelper.Common.WriteToDebugWindow(string.Format("SMOH.{0}({1})", "Column", column));
            Name = column.Name;
            DataType = column.DataType.ToString();
            Default = column.Default;
            MaximumLength = column.DataType.MaximumLength.ToString();
            NumericPrecision = column.DataType.NumericPrecision.ToString();
            NumericScale = column.DataType.NumericScale.ToString();
            ID = column.ID.ToString(); ;
            Identity = column.Identity.ToString();
            InPrimaryKey = column.InPrimaryKey.ToString();
            IsForeignKey = column.IsForeignKey.ToString();
            Nullable = column.Nullable.ToString();
        }

        #endregion 
        
        public string Name;
        public string DataType;
        public string Default;
        public string MaximumLength;
        public string NumericPrecision;
        public string NumericScale;
        public string ID;
        public string Identity;
        public string InPrimaryKey;
        public string IsForeignKey;
        public string Nullable;
    }
}
