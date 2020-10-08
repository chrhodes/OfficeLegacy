using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;

namespace AddinHelper
{
    public class Word
    {
        private Microsoft.Office.Interop.Word.Application _WordApplication;
        public Microsoft.Office.Interop.Word.Application WordApplication
        {
	        get { return _WordApplication; }
	        set { _WordApplication = value; }
        }


        private bool _enableScreenUpdatesToggle = true;
        public bool EnableScreenUpdatesToggle
        {
	        get { return _enableScreenUpdatesToggle; }
	        set { _enableScreenUpdatesToggle = value; }
        }

        /// <summary>
        /// Deletes the style if present.
        /// </summary>
        /// <param name="styleName"></param>
        public void DeleteStyle(string styleName)
        {
            try
            {
                WordApplication.ActiveDocument.Styles[styleName].Delete();
            }
            catch(Exception)
            {
                // Style did not exist
            }

        }
    }
}
