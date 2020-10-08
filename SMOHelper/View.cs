using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SMO = Microsoft.SqlServer.Management.Smo;

namespace SMOHelper
{
    public class View
    {
        #region Initialization

        /// <summary>
        /// Initializes a new instance of the View class.
        /// </summary>
        public View(SMO.View view)
        {
            AddinHelper.Common.WriteToDebugWindow(string.Format("SMOH.{0}({1})", "View", view));
            // Save the real view in case we need to get something from it.
            _view = view;

            CreateDate = view.CreateDate.ToString("yyyy-MM-dd hh:mm:ss");

            try
            {
                DateLastModified = view.DateLastModified.ToString("yyyy-MM-dd hh:mm:ss");
            }
            catch(Exception)
            {
                DateLastModified = "Not Available";
            }

            ID = view.ID.ToString();
            IsSystemObject = view.IsSystemObject;
            Name = view.Name;
            Owner = view.Owner;

            try
            {
                ExtendedProperties = view.ExtendedProperties;
            }
            catch (Exception )
            {

            }
        }

        #endregion

        public SMO.View _view;

        public string CreateDate;
        public string DataSpaceUsed;
        public string DateLastModified;
        public string ID;
        public bool IsSystemObject;
        public string Name;
        public string Owner;

        private Dictionary<string, Column> _Columns;
        public Dictionary<string, Column> Columns
        {
            get
            {
                if (null == _Columns)
                {
                    _Columns = new Dictionary<string, Column>();

                    foreach (SMO.Column column in _view.Columns)
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
