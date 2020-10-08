using System;
using System.Windows.Forms;
using System.Reflection;
using Microsoft.Office.Core;

namespace AddinHelper
{
    /// <summary>
    /// AddInInfo
    /// </summary>
    /// <remarks>
    /// This class can be used in two ways.  If calling this from a commandBar, modify
    /// the Private Const as needed and then create an instance of this class in the code
    /// that loads the command bars.
    /// 
    /// If calling this from a Ribbon Event handler, call the ActionNameGoesHere method directly.
    /// 
    /// Rename the ActionNameGoesHere Method and provide code that does something useful.
    /// </remarks>
    public class AddInInfo : AddinHelper.AppMethod
    {

        #region "Private Constants and Variables"

        private const string _MODULE_NAME = Common.PROJECT_NAME + "AddInInfo";
        private const string _NAME = "AddInInfo";
        private const string _BITMAP_NAME = "AddinInfo.bmp";
        private const string _CAPTION = "AddInInfo";
        private const string _TOOL_TIP_TEXT = "Click for AddInInfo";
        private const string _DESCRIPTION = "AddInInfo does ...";
        #endregion

        #region "Public Methods"

        /// <summary>
        /// 
        /// </summary>
        /// <param name="commandBar">Which bar to add method to</param>
        /// <param name="buttonStyle">The type of button to put on the bar</param>
        /// <remarks></remarks>
        public AddInInfo(ref CommandBar commandBar, ref MsoButtonStyle buttonStyle)
        {
            base.Name = _NAME;
            base.CommandBar = commandBar;
            base.EventHandler = Action;
            base.ButtonStyle = buttonStyle;
            base.BitMapName = _BITMAP_NAME;
            base.Asmbly = Assembly.GetExecutingAssembly();
            base.Caption = _CAPTION;
            base.ToolTipText = _TOOL_TIP_TEXT;
            base.Description = _DESCRIPTION;

            base.Initialize();
        }


        public static void DisplayInfo()
        {
            //AssemblyHelper.AssemblyInformation info = new AssemblyHelper.AssemblyInformation(System.Reflection.Assembly.GetExecutingAssembly());
            AssemblyHelper.AssemblyInformation info = new AssemblyHelper.AssemblyInformation(System.Reflection.Assembly.GetCallingAssembly());
            MessageBox.Show(info.ToString());
        }

        #endregion

        #region "Private Methods"

        private void Action(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                DisplayInfo();
            }
            catch(Exception ex)
            {
                MessageBox.Show(string.Format("Exception: {0}.{1}() - {2}", System.Reflection.Assembly.GetExecutingAssembly().FullName, System.Reflection.MethodInfo.GetCurrentMethod().Name, ex.ToString()));
            }
        }

        #endregion

    }
}