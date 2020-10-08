using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SMO = Microsoft.SqlServer.Management.Smo;

namespace SMOHelper
{
    public class StoredProcedure
    {
        #region Initialization

        /// <summary>
        /// Initializes a new instance of the StoredProcedure class.
        /// </summary>
        public StoredProcedure(SMO.StoredProcedure storedProcedure)
        {
            AddinHelper.Common.WriteToDebugWindow(string.Format("SMOH.{0}({1})", "StoredProcedure", storedProcedure));

            _storedProcedure = storedProcedure;

            CreateDate = storedProcedure.CreateDate.ToString();

            try
            {
                DateLastModified = storedProcedure.DateLastModified.ToString("yyyy-MM-dd hh:mm:ss");
            }
            catch(Exception )
            {
                DateLastModified = "Not Available";
            }

            ID = storedProcedure.ID.ToString();
            IsSystemObject = storedProcedure.IsSystemObject;

            try
            {
                MethodName = storedProcedure.MethodName;
            }
            catch(Exception )
            {
                MethodName = "Not Available";
            }
            
            Name = storedProcedure.Name;
            Owner = storedProcedure.Owner;

            try
            {
                TextHeader = storedProcedure.TextHeader;
            }
            catch(Exception)
            {
                TextHeader = "<No Access>";
            }

            try
            {
                TextBody = storedProcedure.TextBody;
            }
            catch(Exception)
            {
                TextBody = "<No Access>";
            }

            try
            {
                ExtendedProperties = storedProcedure.ExtendedProperties;
            }
            catch (Exception )
            {

            }
        }

        #endregion

        private SMO.StoredProcedure _storedProcedure;

        public string CreateDate;
        public string DateLastModified;
        public string ID;
        public bool IsSystemObject;
        public string MethodName;
        public string Name;
        public string Owner;
        public string TextBody;
        public string TextHeader;


        private SMO.ExtendedPropertyCollection _ExtentedProperties;
        public SMO.ExtendedPropertyCollection ExtendedProperties
        {
            get { return _ExtentedProperties; }
            set { _ExtentedProperties = value; }
        }

        private Dictionary<string, StoredProcedureParameter> _Parameters;
        public Dictionary<string, StoredProcedureParameter> Parameters
        {
            get
            {
                if(null == _Parameters)
                {
                    _Parameters = new Dictionary<string, StoredProcedureParameter>();

                    foreach(SMO.StoredProcedureParameter spParameter in _storedProcedure.Parameters)
                    {
                        SMOHelper.StoredProcedureParameter parameter = new SMOHelper.StoredProcedureParameter(spParameter);

                        _Parameters.Add(spParameter.Name, parameter);
                    };
                }

                return _Parameters;
            }
            set
            {
                _Parameters = value;
            }
        }
    }
}
