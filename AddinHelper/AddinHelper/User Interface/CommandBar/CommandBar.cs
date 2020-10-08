//INSTANT C# NOTE: Formerly VB.NET project-level imports:
using System;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

using Office = Microsoft.Office.Core;

using MsoExcel = Microsoft.Office.Interop.Excel;
using MsoOutlook = Microsoft.Office.Interop.Outlook;
using MsoPowerPoint = Microsoft.Office.Interop.PowerPoint;
using MsoProject = Microsoft.Office.Interop.MSProject;
using MsoVisio = Microsoft.Office.Interop.Visio;
using MsoWord = Microsoft.Office.Interop.Word;

using stdole;

namespace AddinHelper
{
    public class CommandBarHelper
    {
        private const string cMODULE_NAME = Common.PROJECT_NAME + ".CommandBarHelper";

        private object missing = System.Reflection.Missing.Value;

        public static void DisplayCommandBars(object objApp, string appName)
        {
            // Show all the command bars
            // Will need to cast objApp if Option Strict
            string s = "";
            Office.CommandBars cbars = null;

            switch (appName)
            {
                case "Microsoft Excel":
                    MsoExcel.Application appExcel = (MsoExcel.Application)objApp;
                    cbars = appExcel.CommandBars;
                    break;

                case "Microsoft PowerPoint":
                    MsoPowerPoint.Application appPowerPoint = (MsoPowerPoint.Application)objApp;
                    cbars = appPowerPoint.CommandBars;
                    break;

                case "Microsoft Project":
                    MsoProject.Application appProject = (MsoProject.Application)objApp;
                    cbars = appProject.CommandBars;
                    break;

                case "Microsoft Visio":
                    MsoVisio.Application appVisio = (MsoVisio.Application)objApp;
                    cbars = (Office.CommandBars)appVisio.CommandBars;
                    break;
                
                case "Microsoft Word":
                    MsoWord.Application appWord = (MsoWord.Application)objApp;
                    cbars = appWord.CommandBars;
                    break;

                case "Outlook":
                    MsoOutlook.Application appOutlook = (MsoOutlook.Application)objApp;
                    MsoOutlook.Explorer explorer = appOutlook.ActiveExplorer();
                    cbars = explorer.CommandBars;
                    break;

                default:
                    MessageBox.Show("DisplayCommandBars:Unknown appName->" + appName + "<");
                    return;

            }

            foreach (Office.CommandBar bar in cbars)
            {
                s = s + String.Format("{0}\n", bar.Name);
            }

            MessageBox.Show(s);
        }

        public static Office.CommandBarPopup AddMenu(Office.CommandBar cbar, string menuName)
        {
            Office.CommandBarPopup menuItem = null;

            // Clear out any things we have previously added with the same name at this level.

            foreach (Office.CommandBarControl cbc in cbar.Controls)
            {
                if (cbc.Caption == menuName)
                {
                    cbc.Delete(Missing.Value);
                }
            }

            // Now that things are tidy, create the new one.

            menuItem = (Office.CommandBarPopup)cbar.Controls.Add(Office.MsoControlType.msoControlPopup, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            menuItem.Caption = menuName;
            return menuItem;
        }

        public static void RemoveMenu(object objApp, string appName, Office.CommandBar cbar, string menuName, string Name)
        {
            Office.CommandBars cbars = null;

            try
            {
                switch (appName)
                {
                    case "Microsoft Excel":
                        MsoExcel.Application appExcel = (MsoExcel.Application)objApp;
                        cbars = appExcel.CommandBars;
                        break;

                    case "Microsoft PowerPoint":
                        MsoPowerPoint.Application appPowerPoint = (MsoPowerPoint.Application)objApp;
                        cbars = appPowerPoint.CommandBars;
                        break;

                    case "Microsoft Project":
                        MsoProject.Application appProject = (MsoProject.Application)objApp;
                        cbars = appProject.CommandBars;
                        break;

                    case "Microsoft Visio":
                        MsoVisio.Application appVisio = (MsoVisio.Application)objApp;
                        cbars = (Office.CommandBars)appVisio.CommandBars;
                        break;

                    case "Microsoft Word":
                        MsoWord.Application appWord = (MsoWord.Application)objApp;
                        cbars = appWord.CommandBars;
                        break;

                        // TODO: There is no ActiveExplorer when we exit app.  How to handle?
                    //case "Outlook":
                    //    Outlook.Application appOutlook = (Outlook.Application)objApp;
                    //    Outlook.Explorer explorer = appOutlook.ActiveExplorer();
                    //    cbars = explorer.CommandBars;
                    //    break;

                    default:
                        MessageBox.Show("RemoveMenu:Unknown appName->" + appName + "<");
                        return;

                }
  
                cbar = cbars[menuName];

                foreach (Office.CommandBarControl cbc in cbar.Controls)
                {
                    if (cbc.Caption == Name)
                    {
                        cbc.Delete(Missing.Value);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                throw;
            }
        }

        public static Office.CommandBarPopup AddSubMenu(Office.CommandBarPopup Menu, string Name)
        {
            Office.CommandBarPopup menu = null;
            //TODO: INSTANT C# TODO TASK: Instant C# could not resolve the named parameters in the following line:
            menu = (Office.CommandBarPopup)Menu.Controls.Add(Office.MsoControlType.msoControlPopup, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

            menu.Caption = Name;
            return menu;
        }

        /// <summary>
        /// AddToolBar
        /// </summary>
        /// <param name="cbars"></param>
        /// <param name="toolBarName"></param>
        /// <returns></returns>
        /// 
        public static Office.CommandBar
        AddToolBar(
            Office.CommandBars cbars,
            string toolBarName)
        {
            Office.CommandBar cbar;

            // Check to see if the toolbar already exists.  Delete it if found.
            // Need to do this in a loop as some apps allow multiple bars with
            // same name.

            foreach (Office.CommandBar cb in cbars)
            {
                if (cb.Name == toolBarName)
                {
                    cb.Delete();
                }
            }

            cbar = cbars.Add(toolBarName, Office.MsoBarPosition.msoBarTop, false, true);
            cbar.Visible = true;
            return cbar;
        }

        //----------------------------------------------------------------------
        //
        // AddButton()
        //
        // Add buttons to toolbars or menus.  The Icon routines demonstrate a
        // variety of approaches to use.
        //
        // Must pass in assembly with resources.
        //
        //----------------------------------------------------------------------

        public static Office.CommandBarButton
        AddButton(
            Office.CommandBarButton cbb,
            Office.CommandBar cbar,
            string name,
            string caption,
            string description,
            string toolTipText,
            Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler eventHandler)
        {
            cbb = (Office.CommandBarButton)cbar.Controls.Add(Office.MsoControlType.msoControlButton, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

            cbb.Style = Office.MsoButtonStyle.msoButtonCaption;
            cbb.Caption = caption; 
            cbb.DescriptionText = description;
            cbb.TooltipText = toolTipText;
            cbb.Tag = Common.TAG_PREFIX + name;
            cbb.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(eventHandler);
            return cbb;
        }

        public static Office.CommandBarButton
        AddButton(
            Office.CommandBarButton cbb,
            Office.CommandBar cbar,
            string name,
            Assembly asmbly,
            string bitMapName,
            string description,
            string toolTipText,
            Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler eventHandler
        )
        {
            object missing = System.Reflection.Missing.Value;

            try
            {
                cbb = (Office.CommandBarButton)cbar.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, missing, missing);

                //imgStream = asmbly.GetManifestResourceStream(bitMapName);
                //pb.BackgroundImage = Image.FromStream(imgStream);
                //imgStream.Close();

                cbb.Style = Office.MsoButtonStyle.msoButtonIcon;
                cbb.DescriptionText = description;
                //cbb.Picture = ConvertImage.ImageToIPicture(pb.BackgroundImage);
                cbb.Picture = ConvertImage.GetPicture(asmbly, bitMapName);
                cbb.TooltipText = toolTipText;
                cbb.Tag = Common.TAG_PREFIX + name;
                cbb.Click += eventHandler;
                return cbb;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                // TODO: Add Logging
                throw ex;
            }
        }

        public static Office.CommandBarButton
        AddButton(
            Office.CommandBarButton cbb,
            Office.CommandBar cbar,
            string name,
            Assembly asmbly,
            string bitMapName,
            string caption,
            string description,
            string toolTipText,
            Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler eventHandler)
        {
            object missing = System.Reflection.Missing.Value;

            try
            {
                cbb = (Office.CommandBarButton)cbar.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, missing, missing);

                cbb.Style = Office.MsoButtonStyle.msoButtonIconAndCaption;
                cbb.Caption = caption;
                cbb.DescriptionText = description;
                cbb.Picture = ConvertImage.GetPicture(asmbly, bitMapName);
                cbb.TooltipText = toolTipText;
                cbb.Tag = Common.TAG_PREFIX + name;
                cbb.Click += eventHandler;
                return cbb;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                // TODO: Add Logging
                throw ex;
            }
        }

        #region Misc Routines - Commented Out


        //        // Added from prior code
        //        public void ListCommandBars(object objApp)
        //        {

        //            // Purpose: Lists all available command bars in
        //            // this Office application.

        ////INSTANT C# NOTE: Commented this declaration since looping variables in 'foreach' loops are declared in the 'foreach' header in C#
        ////			CommandBar objCommandBar = null;
        //            string strResults = null;

        //            System.Diagnostics.Debug.WriteLine("Command bars in this application:");

        //            foreach (Office.CommandBar objCommandBar in objApp.CommandBars)
        //            {
        //                System.Diagnostics.Debug.WriteLine(objCommandBar.Name + " (" + GetCBType(ref objCommandBar) + ")");
        //            }

        //        }

        //        public void ListCommandBarControls(Office.CommandBar objCommandBar)
        //        {

        //            // Purpose: Given a command bar, lists all of the controls
        //            // on the command bar.

        ////INSTANT C# NOTE: Commented this declaration since looping variables in 'foreach' loops are declared in the 'foreach' header in C#
        ////			CommandBarControl objCommandBarControl = null;
        //            string strControlList = null;

        //            strControlList = "Controls for the '" + objCommandBar.Name + "' command bar:" + System.Environment.NewLine;

        //            foreach (Office.CommandBarControl objCommandBarControl in objCommandBar.Controls)
        //            {

        //                strControlList = strControlList + objCommandBarControl.Caption + " (" + GetCBCtlType(ref objCommandBarControl) + ")" + System.Environment.NewLine;

        //            }

        //            MessageBox.Show(strControlList);

        //        }


        //public void ListButtonPicturesAndIDs(object objApp, short intStart, short intEnd)
        //{
        //    try
        //    {
        //            // Purpose: Given a starting and ending number, creates a
        //            // command bar with pictures corresponding to the face IDs
        //            // in the range of numbers provided.

        //        Office.CommandBar objCommandBar = null;
        //        Office.CommandBarButton objCommandBarButton = null;
        //            short intButton = 0;

        ////			On Error GoTo ListButtonPicturesAndIDs_Err

        //            if (intStart > intEnd)
        //            {

        //                MessageBox.Show("Ending number must be smaller than starting number. " + "Please try again.");

        //                return;

        //            }

        //            foreach (Office.CommandBar objCommandBarWithinLoop in objApp.CommandBars)
        //            {
        //            objCommandBar = objCommandBarWithinLoop;

        //                if (objCommandBarWithinLoop.Name == "Button Pictures and IDs")
        //                {

        //                    objCommandBarWithinLoop.Delete();

        //                }

        //            }

        //            objCommandBar = objApp.CommandBars.Add("Button Pictures and IDs", Missing.Value, Missing.Value, true);

        ////INSTANT C# NOTE: The ending condition of VB 'For' loops is tested only on entry to the loop. Instant C# has created a temporary variable in order to use the initial value of intEnd for every iteration:
        //            short tempFor1 = intEnd;
        //            for (intButton = intStart; intButton <= tempFor1; intButton++)
        //            {

        //                objCommandBarButton = objCommandBar.Controls.Add(Office.MsoControlType.msoControlButton, Missing.Value, Missing.Value, Missing.Value, true);

        //                objCommandBarButton.FaceId = intButton;
        //                objCommandBarButton.TooltipText = "FaceID = " + intButton;

        //            }

        //            objCommandBar.Visible = true;

        //    ListButtonPicturesAndIDs_End:

        //            return;

        //            }

        //    catch
        //    {

        //    //TODO: INSTANT C# TODO TASK: Calls to the VB 'Err' object are not converted by Instant C#:
        //                switch (Err.Number)
        //                {

        //                    case -2147467259: // Invalid FaceIDs.
        //                        MessageBox.Show("Invalid range of numbers for face IDs. " + "Please try again.");
        //                        break;
        //                    default:
        //    //TODO: INSTANT C# TODO TASK: Calls to the VB 'Err' object are not converted by Instant C#:
        //                        MessageBox.Show("Error " + Err.Number + ": " + Err.Description);

        //                        break;
        //                }

        //                goto ListButtonPicturesAndIDs_End;

        //    }
        //}


        //        public void ListOutlookExplorerCommandBarNames(object objApp)
        //        {

        //            // Purpose: Lists all command bar names for the current explorer.
        //            // Note: This code only works in Outlook!

        ////INSTANT C# NOTE: Commented this declaration since looping variables in 'foreach' loops are declared in the 'foreach' header in C#
        ////			CommandBar objCommandBar = null;

        //            //UPGRADE_WARNING: Couldn't resolve default property of object Application.ActiveExplorer. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        //            foreach (Office.CommandBar objCommandBar in objApp.ActiveExplorer.CommandBars)
        //            {

        //                System.Diagnostics.Debug.WriteLine(objCommandBar.Name);

        //            }

        //        }


        //        public void ListCommandBarControlIDs(object objApp)
        //        {

        //            // Purpose: Lists all command bar control IDs for the
        //            // current application.

        ////INSTANT C# NOTE: Commented this declaration since looping variables in 'foreach' loops are declared in the 'foreach' header in C#
        ////			CommandBar objCommandBar = null;
        ////INSTANT C# NOTE: Commented this declaration since looping variables in 'foreach' loops are declared in the 'foreach' header in C#
        ////			CommandBarControl objCommandBarControl = null;

        //            // Replace the next line with:
        //            // For Each objCommandBar In Application.ActiveExplorer.CommandBars <- For Outlook
        //            // For Each objCommandBar In Application.VBE.CommandBars <- For Visual Basic Editor
        //            foreach (Office.CommandBar objCommandBar in objApp.CommandBars)
        //            {
        //                foreach (Office.CommandBarControl objCommandBarControl in objCommandBar.Controls)
        //                {
        //                    System.Diagnostics.Debug.WriteLine(objCommandBarControl.Caption + " " + objCommandBarControl.Id);
        //                }
        //            }
        //        }


        //public bool ChangeCommandBarPosition(object objApp, string strCommandBarName, Office.MsoBarPosition msoPosition)
        //{
        //    try
        //    {

        //            // Purpose: Changes the position of a command bar.
        //            // Accepts:
        //            //   strCommandBarName: The name of the command bar to change position.
        //            // Returns: True if the command bar was successfully moved.

        //            //	On Error GoTo ChangeCommandBarPosition_Err

        //            // Replace the next line of code with:
        //            // Application.ActiveExplorer.CommandBars.Item(strCommandBarName).Position = _
        //            //msoPosition <- For Outlook
        //            // Application.VBE.CommandBars.Item(strCommandBarName).Position = _
        //            //msoPosition <- For the Visual Basic Editor
        //            objApp.CommandBars[strCommandBarName].Position = msoPosition;

        //        ChangeCommandBarPosition_End:

        //        return true;

        //    }

        //    catch
        //    {
        //        return false;
        //    }
        //}

        //Public Sub CommandBarDocumenter(ByVal objApp As Object)

        //    ' Purpose: Writes, to a text file, information about all
        //    ' command bars in the current application.
        //    ' You must first set a reference to the Microsoft Scripting Runtime
        //    ' (scrrun.dll) for this code to run correctly.

        //    ' Note: This code only works with Microsoft Office XP.

        //    Dim objCommandBar As CommandBar
        //    Dim strType As String
        //    Dim strPosition As String
        //    Dim strProtection As String
        //    Dim objFileSaveDialog As Microsoft.Office.Core.FileDialog
        //    Dim objFSO As Scripting.FileSystemObject
        //    Dim objTextStream As Scripting.TextStream
        //    Const SAVE_BUTTON As Short = -1

        //    objFileSaveDialog = objApp.FileDialog(MsoFileDialogType.msoFileDialogSaveAs)

        //    objFileSaveDialog.Title = "Save Results As"

        //    ' User clicked "Save" button.
        //    If objFileSaveDialog.Show = SAVE_BUTTON Then

        //        objFSO = New Scripting.FileSystemObject
        //        objTextStream = objFSO.CreateTextFile(objFileSaveDialog.SelectedItems.Item(1))

        //        objTextStream.WriteLine("Name" & vbTab & "Type" & vbTab & "Enabled" & vbTab & "Visible" & vbTab & "Index" & vbTab & "Position" & vbTab & "Protection" & vbTab & "Row Index" & vbTab & "Top" & vbTab & "Height" & vbTab & "Left" & vbTab & "Width")

        //        ' Replace the next line with:
        //        ' For Each objCommandBar In Application.ActiveExplorer.CommandBars _
        //        '<- For Outlook
        //        ' For Each objCommandBar In Application.VBE.CommandBars <- For _
        //        'Visual Basic Editor
        //        For Each objCommandBar In objApp.CommandBars

        //            Select Case objCommandBar.Type
        //                Case MsoBarType.msoBarTypeMenuBar
        //                    strType = "Menu Bar"
        //                Case MsoBarType.msoBarTypeNormal
        //                    strType = "Normal"
        //                Case MsoBarType.msoBarTypePopup
        //                    strType = "Pop-Up"
        //            End Select

        //            Select Case objCommandBar.Position
        //                Case MsoBarPosition.msoBarBottom
        //                    strPosition = "Bottom"
        //                Case MsoBarPosition.msoBarFloating
        //                    strPosition = "Floating"
        //                Case MsoBarPosition.msoBarLeft
        //                    strPosition = "Left"
        //                Case MsoBarPosition.msoBarMenuBar
        //                    strPosition = "Menu Bar"
        //                Case MsoBarPosition.msoBarPopup
        //                    strPosition = "Pop-Up"
        //                Case MsoBarPosition.msoBarRight
        //                    strPosition = "Right"
        //                Case MsoBarPosition.msoBarTop
        //                    strPosition = "Top"
        //            End Select

        //            Select Case objCommandBar.Protection
        //                Case MsoBarProtection.msoBarNoChangeDock
        //                    strProtection = "No Change Dock"
        //                Case MsoBarProtection.msoBarNoChangeVisible
        //                    strProtection = "No Change Visible"
        //                Case MsoBarProtection.msoBarNoCustomize
        //                    strProtection = "No Customize"
        //                Case MsoBarProtection.msoBarNoHorizontalDock
        //                    strProtection = "No Horizontal Dock"
        //                Case MsoBarProtection.msoBarNoMove
        //                    strProtection = "No Move"
        //                Case MsoBarProtection.msoBarNoProtection
        //                    strProtection = "No Protection"
        //                Case MsoBarProtection.msoBarNoResize
        //                    strProtection = "No Resize"
        //                Case MsoBarProtection.msoBarNoVerticalDock
        //                    strProtection = "No Vertical Dock"
        //            End Select

        //            objTextStream.WriteLine(objCommandBar.Name & vbTab & strType & vbTab & objCommandBar.Enabled & vbTab & objCommandBar.Visible & vbTab & objCommandBar.Index & vbTab & strPosition & vbTab & strProtection & vbTab & objCommandBar.RowIndex & vbTab & objCommandBar.Top & vbTab & objCommandBar.Height & vbTab & objCommandBar.Left & vbTab & objCommandBar.Width)

        //        Next objCommandBar

        //        objTextStream.Close()

        //        MsgBox("Results written to " & objFileSaveDialog.SelectedItems.Item(1) & ".")

        //    End If

        //End Sub

        //Sub CommandBarControlDocumenter(ByVal objApp As Object)

        //    ' Purpose: Writes, to a text file, information about all
        //    ' command bar controls in the current application.

        //    ' You must first set a reference to the Microsoft Scripting Runtime
        //    ' (scrrun.dll) for this code to run correctly.

        //    ' Note: This code only works with Microsoft Office XP.

        //    Dim objCommandBar As CommandBar
        //    Dim objCommandBarControl As CommandBarControl
        //    Dim strOLEUsage As String
        //    Dim strType As String
        //    Dim objFileSaveDialog As Microsoft.Office.Core.FileDialog
        //    Dim objFSO As Scripting.FileSystemObject
        //    Dim objTextStream As Scripting.TextStream
        //    Const SAVE_BUTTON As Short = -1

        //    objFileSaveDialog = objApp.FileDialog(MsoFileDialogType.msoFileDialogSaveAs)

        //    objFileSaveDialog.Title = "Save Results As"

        //    ' User clicked "Save" button.
        //    If objFileSaveDialog.Show = SAVE_BUTTON Then

        //        objFSO = New Scripting.FileSystemObject
        //        objTextStream = objFSO.CreateTextFile(objFileSaveDialog.SelectedItems.Item(1))

        //        objTextStream.WriteLine("ID" & vbTab & "Index" & vbTab & "Caption" & vbTab & "Parent" & vbTab & "DescriptionText" & vbTab & "BuiltIn" & vbTab & "Enabled" & vbTab & "IsPriorityDropped" & vbTab & "OLEUsage" & vbTab & "Priority" & vbTab & "Tag" & vbTab & "TooltipText" & vbTab & "Type" & vbTab & "Visible" & vbTab & "Height" & vbTab & "Width")

        //        ' Replace the next line with:
        //        ' For Each objCommandBar In Application.ActiveExplorer.CommandBars <- For Outlook
        //        ' For Each objCommandBar In Application.VBE.CommandBars _
        //        '<- For Visual Basic Editor
        //        For Each objCommandBar In objApp.CommandBars

        //            For Each objCommandBarControl In objCommandBar.Controls

        //                Select Case objCommandBarControl.OLEUsage

        //                    Case MsoControlOLEUsage.msoControlOLEUsageBoth
        //                        strOLEUsage = "Both"
        //                    Case MsoControlOLEUsage.msoControlOLEUsageClient
        //                        strOLEUsage = "Client"
        //                    Case MsoControlOLEUsage.msoControlOLEUsageNeither
        //                        strOLEUsage = "Neither"
        //                    Case MsoControlOLEUsage.msoControlOLEUsageServer
        //                        strOLEUsage = "Server"

        //                End Select

        //                Select Case objCommandBarControl.Type

        //                    Case MsoControlType.msoControlActiveX
        //                        strType = "ActiveX"
        //                    Case MsoControlType.msoControlAutoCompleteCombo
        //                        strType = "Auto-Complete Combo Box"
        //                    Case MsoControlType.msoControlButton
        //                        strType = "Button"
        //                    Case MsoControlType.msoControlButtonDropdown
        //                        strType = "Drop-Down Button"
        //                    Case MsoControlType.msoControlButtonPopup
        //                        strType = "Popup Button"
        //                    Case MsoControlType.msoControlComboBox
        //                        strType = "Combo Box"
        //                    Case MsoControlType.msoControlCustom
        //                        strType = "Custom"
        //                    Case MsoControlType.msoControlDropdown
        //                        strType = "Drop-Down"
        //                    Case MsoControlType.msoControlEdit
        //                        strType = "Edit"
        //                    Case MsoControlType.msoControlExpandingGrid
        //                        strType = "Expanding Grid"
        //                    Case MsoControlType.msoControlGauge
        //                        strType = "Gauge"
        //                    Case MsoControlType.msoControlGenericDropdown
        //                        strType = "Generic Drop-Down"
        //                    Case MsoControlType.msoControlGraphicCombo
        //                        strType = "Graphic Combo Box"
        //                    Case MsoControlType.msoControlGraphicDropdown
        //                        strType = "Graphic Drop-Down"
        //                    Case MsoControlType.msoControlGraphicPopup
        //                        strType = "Popup Graphic"
        //                    Case MsoControlType.msoControlGrid
        //                        strType = "Grid"
        //                    Case MsoControlType.msoControlLabel
        //                        strType = "Label"
        //                    Case MsoControlType.msoControlLabelEx
        //                        strType = "LabelEx"
        //                    Case MsoControlType.msoControlOCXDropdown
        //                        strType = "OCX Drop-Down"
        //                    Case MsoControlType.msoControlPane
        //                        strType = "Pane"
        //                    Case MsoControlType.msoControlPopup
        //                        strType = "Popup"
        //                    Case MsoControlType.msoControlSpinner
        //                        strType = "Spinner"
        //                    Case MsoControlType.msoControlSplitButtonMRUPopup
        //                        strType = "Split Button MRU Popup"
        //                    Case MsoControlType.msoControlSplitButtonPopup
        //                        strType = "Button Popup"
        //                    Case MsoControlType.msoControlSplitDropdown
        //                        strType = "Split Drop-Down"
        //                    Case MsoControlType.msoControlSplitExpandingGrid
        //                        strType = "Split Expanding Grid"
        //                    Case MsoControlType.msoControlWorkPane
        //                        strType = "Work Pane"

        //                End Select

        //                objTextStream.WriteLine(objCommandBarControl.Id & vbTab & objCommandBarControl.Index & vbTab & objCommandBarControl.Caption & vbTab & objCommandBarControl.Parent.Name & vbTab & objCommandBarControl.DescriptionText & vbTab & objCommandBarControl.BuiltIn & vbTab & objCommandBarControl.Enabled & vbTab & objCommandBarControl.IsPriorityDropped & vbTab & strOLEUsage & vbTab & objCommandBarControl.Priority & vbTab & objCommandBarControl.Tag & vbTab & objCommandBarControl.TooltipText & vbTab & strType & vbTab & objCommandBarControl.Visible & vbTab & objCommandBarControl.Height & vbTab & objCommandBarControl.Width)

        //            Next objCommandBarControl

        //        Next objCommandBar

        //        objTextStream.Close()

        //        MsgBox("Results written to " & objFileSaveDialog.SelectedItems.Item(1) & ".")

        //    End If
        //End Sub


        //        public bool CBToolbarShow(object objApp, ref string strCBarName, ref bool blnVisible)
        //        {
        //            int templngPosition1 = Office.MsoBarPosition.msoBarTop;
        //            return CBToolbarShow(objApp, ref strCBarName, ref blnVisible, ref templngPosition1);
        //        }

        ////INSTANT C# NOTE: C# does not support optional parameters. Overloaded method(s) are created above.
        ////ORIGINAL LINE: Function CBToolbarShow(ByVal objApp As Object, ByRef strCBarName As String, ByRef blnVisible As Boolean, Optional ByRef lngPosition As Integer = MsoBarPosition.msoBarTop) As Boolean
        //        public bool CBToolbarShow(object objApp, ref string strCBarName, ref bool blnVisible, ref int lngPosition)
        //        {
        //            bool tempCBToolbarShow = false;
        //            try
        //            {

        //                    // This procedure displays or hides the command bar specified in the
        //                    // strCBarName argument according to the value of the blnVisible
        //                    // argument. The optional lngPosition argument specifies where the
        //                    // command bar will appear on the screen.

        //                Office.CommandBar cbrCmdBar = null;

        //        //			On Error GoTo CBToolbarShow_Err

        //                    cbrCmdBar = objApp.CommandBars(strCBarName);

        //                    // Show only toolbars.
        //                    if (cbrCmdBar.Type > Office.MsoBarType.msoBarTypeNormal)
        //                        return false;
        //                    // If Position argument is invalid, set to the default
        //                    // msoBarTop position.
        //                    if (lngPosition < Office.MsoBarPosition.msoBarLeft | lngPosition > MOffice.soBarPosition.msoBarMenuBar)
        //                        lngPosition = Office.MsoBarPosition.msoBarTop;

        //                    cbrCmdBar.Visible = blnVisible;
        //                    cbrCmdBar.Position = lngPosition;

        //                    tempCBToolbarShow = true;

        //            CBToolbarShow_End:
        //                    return tempCBToolbarShow;
        //            }

        //            catch
        //            {
        //                        tempCBToolbarShow = false;
        //                        goto CBToolbarShow_End;
        //            }
        //            return tempCBToolbarShow;
        //        }

        //public static bool DeleteCommandBar(object objApp, ref string strCBarName)
        //{
        //    // Delete the command bar specified by strCBarName. If the
        //    // command bar does not exist, an error will occur and that
        //    // error is ignored here.

        //    try
        //    {
        //        objApp.CommandBars(strCBarName).Delete();
        //        return true;
        //    }
        //    catch (Exception ex)
        //    {
        //        return false;
        //        //throw;
        //    }
        //}

        #endregion

        public string GetCBType(ref Office.CommandBar cbrBar)
        {
            // Returns a string representing the command bar type.

            string strCBType = null;

            //INSTANT C# NOTE: The following VB 'Select Case' included range-type or non-constant 'Case' expressions and was converted to C# 'if-else' logic:
            //			Select Case cbrBar.Type
            //ORIGINAL LINE: Case MsoBarType.msoBarTypeNormal
            if (cbrBar.Type == Office.MsoBarType.msoBarTypeNormal)
                strCBType = "Toolbar";
            //ORIGINAL LINE: Case MsoBarType.msoBarTypeMenuBar
            else if (cbrBar.Type == Office.MsoBarType.msoBarTypeMenuBar)
                strCBType = "Menu bar";
            //ORIGINAL LINE: Case MsoBarType.msoBarTypePopup
            else if (cbrBar.Type == Office.MsoBarType.msoBarTypePopup)
                strCBType = "Popup menu";

            return strCBType;
        }

        public string GetCBCtlType(ref Office.CommandBarControl ctlCBarControl)
        {
            // Returns a string representing the command bar control's type.

            string strType = null;

            //INSTANT C# NOTE: The following VB 'Select Case' included range-type or non-constant 'Case' expressions and was converted to C# 'if-else' logic:
            //			Select Case ctlCBarControl.Type
            //ORIGINAL LINE: Case MsoControlType.msoControlCustom
            if (ctlCBarControl.Type == Office.MsoControlType.msoControlCustom)
                strType = "Custom";
            //ORIGINAL LINE: Case MsoControlType.msoControlButton
            else if (ctlCBarControl.Type == Office.MsoControlType.msoControlButton)
                strType = "Button";
            //ORIGINAL LINE: Case MsoControlType.msoControlEdit
            else if (ctlCBarControl.Type == Office.MsoControlType.msoControlEdit)
                strType = "Edit";
            //ORIGINAL LINE: Case MsoControlType.msoControlDropdown
            else if (ctlCBarControl.Type == Office.MsoControlType.msoControlDropdown)
                strType = "Dropdown";
            //ORIGINAL LINE: Case MsoControlType.msoControlComboBox
            else if (ctlCBarControl.Type == Office.MsoControlType.msoControlComboBox)
                strType = "Combobox";
            //ORIGINAL LINE: Case MsoControlType.msoControlButtonDropdown
            else if (ctlCBarControl.Type == Office.MsoControlType.msoControlButtonDropdown)
                strType = "ButtonDropdown";
            //ORIGINAL LINE: Case MsoControlType.msoControlSplitDropdown
            else if (ctlCBarControl.Type == Office.MsoControlType.msoControlSplitDropdown)
                strType = "SplitDropdown";
            //ORIGINAL LINE: Case MsoControlType.msoControlOCXDropdown
            else if (ctlCBarControl.Type == Office.MsoControlType.msoControlOCXDropdown)
                strType = "OCXDropdown";
            //ORIGINAL LINE: Case MsoControlType.msoControlGenericDropdown
            else if (ctlCBarControl.Type == Office.MsoControlType.msoControlGenericDropdown)
                strType = "GenericDropdown";
            //ORIGINAL LINE: Case MsoControlType.msoControlGraphicDropdown
            else if (ctlCBarControl.Type == Office.MsoControlType.msoControlGraphicDropdown)
                strType = "GraphicDropdown";
            //ORIGINAL LINE: Case MsoControlType.msoControlPopup
            else if (ctlCBarControl.Type == Office.MsoControlType.msoControlPopup)
                strType = "Popup";
            //ORIGINAL LINE: Case MsoControlType.msoControlGraphicPopup
            else if (ctlCBarControl.Type == Office.MsoControlType.msoControlGraphicPopup)
                strType = "GraphicPopup";
            //ORIGINAL LINE: Case MsoControlType.msoControlButtonPopup
            else if (ctlCBarControl.Type == Office.MsoControlType.msoControlButtonPopup)
                strType = "ButtonPopup";
            //ORIGINAL LINE: Case MsoControlType.msoControlSplitButtonPopup
            else if (ctlCBarControl.Type == Office.MsoControlType.msoControlSplitButtonPopup)
                strType = "SplitButtonPopup";
            //ORIGINAL LINE: Case MsoControlType.msoControlSplitButtonMRUPopup
            else if (ctlCBarControl.Type == Office.MsoControlType.msoControlSplitButtonMRUPopup)
                strType = "SplitButtonMRUPopup";
            //ORIGINAL LINE: Case MsoControlType.msoControlLabel
            else if (ctlCBarControl.Type == Office.MsoControlType.msoControlLabel)
                strType = "Label";
            //ORIGINAL LINE: Case MsoControlType.msoControlExpandingGrid
            else if (ctlCBarControl.Type == Office.MsoControlType.msoControlExpandingGrid)
                strType = "ExpandingGrid";
            //ORIGINAL LINE: Case MsoControlType.msoControlSplitExpandingGrid
            else if (ctlCBarControl.Type == Office.MsoControlType.msoControlSplitExpandingGrid)
                strType = "SplitExpandingGrid";
            //ORIGINAL LINE: Case MsoControlType.msoControlGrid
            else if (ctlCBarControl.Type == Office.MsoControlType.msoControlGrid)
                strType = "Grid";
            //ORIGINAL LINE: Case MsoControlType.msoControlGauge
            else if (ctlCBarControl.Type == Office.MsoControlType.msoControlGauge)
                strType = "Gauge";
            //ORIGINAL LINE: Case MsoControlType.msoControlGraphicCombo
            else if (ctlCBarControl.Type == Office.MsoControlType.msoControlGraphicCombo)
                strType = "GraphicCombo";
            //ORIGINAL LINE: Case MsoControlType.msoControlPane
            else if (ctlCBarControl.Type == Office.MsoControlType.msoControlPane)
                strType = "Pane";

            return strType;
        }

        #region Misc Routines - Commented Out


        //        public bool CBCtlToggleEnabled(object objApp, ref string strCBarName, ref string strCtlCaption)
        //        {
        //            bool tempCBCtlToggleEnabled = false;

        //            // Toggle the Enabled property of the strCtlCaption control
        //            // on the strCBarName command bar.

        //            Office.CommandBarControl ctlCBarControl = null;

        ////TODO: INSTANT C# TODO TASK: The 'On Error Resume Next' statement is not converted by Instant C#:
        //            //On Error Resume Next

        //            ctlCBarControl = objApp.Office.CommandBars(strCBarName).Controls[strCtlCaption];
        //            ctlCBarControl.Enabled = ! ctlCBarControl.Enabled;
        ////TODO: INSTANT C# TODO TASK: Calls to the VB 'Err' object are not converted by Instant C#:
        //            if (Err.Number == 0)
        //                tempCBCtlToggleEnabled = true;
        //            else
        //                tempCBCtlToggleEnabled = false;
        //            return tempCBCtlToggleEnabled;
        //        }

        //        public bool CBCtlToggleVisible(object objApp, ref string strCBarName, ref string strCtlCaption)
        //        {
        //            bool tempCBCtlToggleVisible = false;

        //            // Toggle the Visible property of the strCtlCaption control
        //            // on the strCBarName command bar.

        //            Office.CommandBarControl ctlCBarControl = null;

        ////TODO: INSTANT C# TODO TASK: The 'On Error Resume Next' statement is not converted by Instant C#:
        //            //On Error Resume Next

        //            ctlCBarControl = objApp.CommandBars(strCBarName).Controls[strCtlCaption];
        //            ctlCBarControl.Visible = ! ctlCBarControl.Visible;
        ////TODO: INSTANT C# TODO TASK: Calls to the VB 'Err' object are not converted by Instant C#:
        //            if (Err.Number == 0)
        //                tempCBCtlToggleVisible = true;
        //            else
        //                tempCBCtlToggleVisible = false;
        //            return tempCBCtlToggleVisible;
        //        }

        //        public bool CBCtlToggleState(object objApp, ref string strCBarName, ref string strCtlCaption)
        //        {
        //            bool tempCBCtlToggleState = false;

        //            // Toggle the State property of the strCtlCaption control
        //            // on the strCBarName command bar. The State property is
        //            // read-only for built-in controls, so if strCtlCaption
        //            // is a built-in control, return False and exit the procedure.

        //            Office.CommandBarControl ctlCBarControl = null;

        ////TODO: INSTANT C# TODO TASK: The 'On Error Resume Next' statement is not converted by Instant C#:
        //            //On Error Resume Next

        //            ctlCBarControl = objApp.CommandBars(strCBarName).Controls[strCtlCaption];

        //            if (ctlCBarControl.BuiltIn == true)
        //                return false;

        //            if (ctlCBarControl.Type != Office.MsoControlType.msoControlButton)
        //                return false;

        //            //UPGRADE_WARNING: Couldn't resolve default property of object ctlCBarControl.State. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        //            ctlCBarControl.State = ! ctlCBarControl.State;

        ////TODO: INSTANT C# TODO TASK: Calls to the VB 'Err' object are not converted by Instant C#:
        //            if (Err.Number == 0)
        //                tempCBCtlToggleState = true;
        //            else
        //                tempCBCtlToggleState = false;

        //            return tempCBCtlToggleState;
        //        }

        //        public object DeleteCBControl(object objApp, ref string strCBarName, ref string strCtlName)
        //        {
        ////TODO: INSTANT C# TODO TASK: The 'On Error Resume Next' statement is not converted by Instant C#:
        //            //On Error Resume Next
        //            objApp.CommandBars(strCBarName).Controls[strCtlName].Delete();
        //            //INSTANT C# NOTE: Inserted the following 'return' since all code paths must return a value in C#:
        //            return null;
        //        }

        //        //public CommandBars()
        //        //{
        //        //    //' Point the logging code to our configuration file.
        //        //    //Log.Logging.ConfigFile = m_LogConfigFile
        //        //    PacificLife.Life.Enterprise.Log.Logging.Log(PacificLife.Life.Enterprise.Log.Severity.Verbose, cMODULE_NAME + "-New", "ObjectLifetime");
        //        //}
        //    }
        #endregion
    }

    public class ConvertImage : System.Windows.Forms.AxHost
    {
        public ConvertImage()
            : base("59EE46BA-677D-4d20-BF10-8D8067CB8B33")
        {
        }
        
        public static IPictureDisp
        GetPicture(Assembly asmbly, string bitMapName)
        {
            Stream imgStream;
            PictureBox pb = new PictureBox();
            string fullBitmapName = asmbly.GetName().Name + "." + bitMapName;

            try
            {
                imgStream = asmbly.GetManifestResourceStream(fullBitmapName);
                pb.BackgroundImage = Image.FromStream(imgStream);
                imgStream.Close();
                return ConvertImage.ImageToIPicture(pb.BackgroundImage);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Missing bitmap? " + fullBitmapName + ":" + ex.ToString());
                imgStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("AddinHelper.missing.bmp");
                pb.BackgroundImage = Image.FromStream(imgStream);
                imgStream.Close();
                return ConvertImage.ImageToIPicture(pb.BackgroundImage);
            }

            #region old ideas

            //Stream imgStream = null;
            //Dim imgList As New ImageList   ' Uncomment if using below
            //PictureBox pb = new PictureBox();
            //Dim dataO As New DataObject    ' Uncomment if using below

            //try
            //{
            //imgStream = asmbly.GetManifestResourceStream(BitMap);
            // Use this if multiple images
            //imgList.Images.Add(Image.FromStream(imgStream))
            // Use this if one image
            //pb.BackgroundImage = Image.FromStream(imgStream);
            //imgStream.Close();

            //tempAddButtonWithIcon.Style = Office.MsoButtonStyle.msoButtonIcon;
            //tempAddButtonWithIcon.DescriptionText = DescriptionText;

            // bit map to button options.
            // ToDo: Pick one of these two techniques.  Both are flawed in using the
            // clipboard as something important may be overwritten.

            // 1
            //dataO.SetData(System.Windows.Forms.DataFormats.Bitmap, imgList.Images.Item(0))
            //Clipboard.SetDataObject(dataO)
            //.PasteFace()

            // 2
            //Clipboard.SetDataObject(pb.BackgroundImage)
            //.PasteFace()

            // 3 good - no clipboard usage - useful when need list of images.
            //.Picture = ConvertImage.ImageToIPicture(imgList.Images.Item(0))


            // 4 good - no clipboard usage
            //tempAddButtonWithIcon.Picture = ConvertImage.ImageToIPicture(pb.BackgroundImage);

            // End bit map to button options
            //}
            #endregion
        }

        public static IPictureDisp
        ImageToIPicture(Image image)
        {
            return (IPictureDisp)AxHost.GetIPictureDispFromPicture(image);
        }

        public static System.Drawing.Image
        IPictureToImage(IPictureDisp image)
        {
            return (Image)AxHost.GetPictureFromIPicture(image);
        }
    }
} 