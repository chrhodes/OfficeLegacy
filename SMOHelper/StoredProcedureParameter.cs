using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SMO = Microsoft.SqlServer.Management.Smo;

namespace SMOHelper
{
    public class StoredProcedureParameter
    {
        #region Initialization

        /// <summary>
        /// Initializes a new instance of the StoredProcedure class.
        /// </summary>
        public StoredProcedureParameter(SMO.StoredProcedureParameter storedProcedureParameter)
        {
            AddinHelper.Common.WriteToDebugWindow(string.Format("SMOH.{0}({1})", "StoredProcedureParameter", storedProcedureParameter));
            
            DataType = storedProcedureParameter.DataType.ToString();
            DefaultValue = storedProcedureParameter.DefaultValue;
            MaximumLength = storedProcedureParameter.DataType.MaximumLength.ToString();
            Name = storedProcedureParameter.Name;
            NumericPrecision = storedProcedureParameter.DataType.NumericPrecision.ToString();
            NumericScale = storedProcedureParameter.DataType.NumericScale.ToString();     
        }

        #endregion

        public string DataType;
        public string Name;
        public string DefaultValue;
        public string MaximumLength;
        public string NumericPrecision;
        public string NumericScale;
    }
}
