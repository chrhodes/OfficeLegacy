using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using SMO = Microsoft.SqlServer.Management.Smo;

namespace SMOHelper
{
    class ExtendedProperty
    {
        /// <summary>
        /// Initializes a new instance of the ExtendedProperty class.
        /// </summary>
        public ExtendedProperty(SMO.ExtendedProperty extendedProperty)
        {
            AddinHelper.Common.WriteToDebugWindow(string.Format("SMOH.{0}({1})", "ExtendedProperty", extendedProperty));
            Name = extendedProperty.Name;
            Value = extendedProperty.Value.ToString();
        }

        // Fields...
        private string _Value;
        private string _Name;

        public string Name
        {
            get { return _Name; }
            set
            {
                _Name = value;
            }
        }


        public string Value
        {
            get { return _Value; }
            set
            {
                _Value = value;
            }
        }
    }
}
