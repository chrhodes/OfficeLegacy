using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SMO = Microsoft.SqlServer.Management.Smo;

namespace SMOHelper
{
    public class Login
    {
        #region Initialization

        /// <summary>
        /// Initializes a new instance of the Login class.
        /// </summary>
        public Login(SMO.Login login)
        {
            AddinHelper.Common.WriteToDebugWindow(string.Format("SMOH.{0}({1})", "Login", login));
            CreateDate = login.CreateDate.ToString("yyyy-MM-dd hh:mm:ss");
            DefaultDatabase = login.DefaultDatabase;
            Name = login.Name;
        }

        #endregion

        public string CreateDate;
        public string DefaultDatabase;
        public string Name;

    }
}
