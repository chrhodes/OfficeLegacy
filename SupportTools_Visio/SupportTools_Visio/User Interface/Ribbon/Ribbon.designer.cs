﻿namespace SupportTools_Visio
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.grpSupportTools = this.Factory.CreateRibbonGroup();
            this.tabSupportTools = this.Factory.CreateRibbonTab();
            this.grpTaskPanes = this.Factory.CreateRibbonGroup();
            this.btnAppUtilities = this.Factory.CreateRibbonButton();
            this.grpDebug = this.Factory.CreateRibbonGroup();
            this.chkEnableAppEvents = this.Factory.CreateRibbonCheckBox();
            this.chkDisplayEvents = this.Factory.CreateRibbonCheckBox();
            this.chkDisplayChattyEvents = this.Factory.CreateRibbonCheckBox();
            this.grpHelp = this.Factory.CreateRibbonGroup();
            this.btnAddInInfo = this.Factory.CreateRibbonButton();
            this.btnDeveloperMode = this.Factory.CreateRibbonButton();
            this.rgVisio_Utilities = this.Factory.CreateRibbonGroup();
            this.btnAddTableOfContents = this.Factory.CreateRibbonButton();
            this.btnDebugWindow = this.Factory.CreateRibbonButton();
            this.btnWatchWindow = this.Factory.CreateRibbonButton();
            this.rgSMARTS = this.Factory.CreateRibbonGroup();
            this.btnRetrive = this.Factory.CreateRibbonButton();
            this.btnReleatedProcess = this.Factory.CreateRibbonButton();
            this.btnValidate = this.Factory.CreateRibbonButton();
            this.btnHilight = this.Factory.CreateRibbonButton();
            this.btnWebPage = this.Factory.CreateRibbonButton();
            this.btnRelatedIntfrastructure = this.Factory.CreateRibbonButton();
            this.btnRelatedSystem = this.Factory.CreateRibbonButton();
            this.btnNavigateUp = this.Factory.CreateRibbonButton();
            this.btnNavigateDown = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.tabSupportTools.SuspendLayout();
            this.grpTaskPanes.SuspendLayout();
            this.grpDebug.SuspendLayout();
            this.grpHelp.SuspendLayout();
            this.rgVisio_Utilities.SuspendLayout();
            this.rgSMARTS.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.grpSupportTools);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // grpSupportTools
            // 
            this.grpSupportTools.Label = "Support Tools";
            this.grpSupportTools.Name = "grpSupportTools";
            // 
            // tabSupportTools
            // 
            this.tabSupportTools.Groups.Add(this.grpTaskPanes);
            this.tabSupportTools.Groups.Add(this.grpDebug);
            this.tabSupportTools.Groups.Add(this.grpHelp);
            this.tabSupportTools.Groups.Add(this.rgVisio_Utilities);
            this.tabSupportTools.Groups.Add(this.rgSMARTS);
            this.tabSupportTools.Label = "Support Tools";
            this.tabSupportTools.Name = "tabSupportTools";
            // 
            // grpTaskPanes
            // 
            this.grpTaskPanes.Items.Add(this.btnAppUtilities);
            this.grpTaskPanes.Label = "Task Panes";
            this.grpTaskPanes.Name = "grpTaskPanes";
            // 
            // btnAppUtilities
            // 
            this.btnAppUtilities.Label = "Visio Utilities";
            this.btnAppUtilities.Name = "btnAppUtilities";
            this.btnAppUtilities.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAppUtilities_Click);
            // 
            // grpDebug
            // 
            this.grpDebug.Items.Add(this.btnDebugWindow);
            this.grpDebug.Items.Add(this.btnWatchWindow);
            this.grpDebug.Items.Add(this.chkEnableAppEvents);
            this.grpDebug.Items.Add(this.chkDisplayEvents);
            this.grpDebug.Items.Add(this.chkDisplayChattyEvents);
            this.grpDebug.Label = "Debug";
            this.grpDebug.Name = "grpDebug";
            this.grpDebug.Visible = false;
            // 
            // chkEnableAppEvents
            // 
            this.chkEnableAppEvents.Label = "Enable App Events";
            this.chkEnableAppEvents.Name = "chkEnableAppEvents";
            this.chkEnableAppEvents.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.chkEnableAppEvents_Click);
            // 
            // chkDisplayEvents
            // 
            this.chkDisplayEvents.Label = "Display Events";
            this.chkDisplayEvents.Name = "chkDisplayEvents";
            this.chkDisplayEvents.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.chkDisplayEvents_Click);
            // 
            // chkDisplayChattyEvents
            // 
            this.chkDisplayChattyEvents.Label = "Display Chatty Events";
            this.chkDisplayChattyEvents.Name = "chkDisplayChattyEvents";
            this.chkDisplayChattyEvents.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.chkDisplayChattyEvents_Click);
            // 
            // grpHelp
            // 
            this.grpHelp.Items.Add(this.btnAddInInfo);
            this.grpHelp.Items.Add(this.btnDeveloperMode);
            this.grpHelp.Label = "Help";
            this.grpHelp.Name = "grpHelp";
            // 
            // btnAddInInfo
            // 
            this.btnAddInInfo.Label = "AddIn Info";
            this.btnAddInInfo.Name = "btnAddInInfo";
            this.btnAddInInfo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddInInfo_Click);
            // 
            // btnDeveloperMode
            // 
            this.btnDeveloperMode.Label = "Developer Mode";
            this.btnDeveloperMode.Name = "btnDeveloperMode";
            this.btnDeveloperMode.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDeveloperMode_Click);
            // 
            // rgVisio_Utilities
            // 
            this.rgVisio_Utilities.Items.Add(this.btnAddTableOfContents);
            this.rgVisio_Utilities.Label = "Visio Utilities";
            this.rgVisio_Utilities.Name = "rgVisio_Utilities";
            // 
            // btnAddTableOfContents
            // 
            this.btnAddTableOfContents.Label = "Add Table of Contents";
            this.btnAddTableOfContents.Name = "btnAddTableOfContents";
            this.btnAddTableOfContents.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddTableOfContents_Click);
            // 
            // btnDebugWindow
            // 
            this.btnDebugWindow.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnDebugWindow.Image = global::SupportTools_Visio.Properties.Resources.Auto_Debug_System_icon;
            this.btnDebugWindow.Label = "Debug Window";
            this.btnDebugWindow.Name = "btnDebugWindow";
            this.btnDebugWindow.ShowImage = true;
            this.btnDebugWindow.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDebugWindow_Click);
            // 
            // btnWatchWindow
            // 
            this.btnWatchWindow.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnWatchWindow.Image = global::SupportTools_Visio.Properties.Resources.WatchWindow;
            this.btnWatchWindow.Label = "Watch Window";
            this.btnWatchWindow.Name = "btnWatchWindow";
            this.btnWatchWindow.ShowImage = true;
            this.btnWatchWindow.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnWatchWindow_Click);
            // 
            // rgSMARTS
            // 
            this.rgSMARTS.Items.Add(this.btnRetrive);
            this.rgSMARTS.Items.Add(this.btnWebPage);
            this.rgSMARTS.Items.Add(this.btnValidate);
            this.rgSMARTS.Items.Add(this.btnReleatedProcess);
            this.rgSMARTS.Items.Add(this.btnRelatedSystem);
            this.rgSMARTS.Items.Add(this.btnRelatedIntfrastructure);
            this.rgSMARTS.Items.Add(this.btnNavigateUp);
            this.rgSMARTS.Items.Add(this.btnNavigateDown);
            this.rgSMARTS.Items.Add(this.btnHilight);
            this.rgSMARTS.Label = "SMARTS";
            this.rgSMARTS.Name = "rgSMARTS";
            // 
            // btnRetrive
            // 
            this.btnRetrive.Label = "Retrive";
            this.btnRetrive.Name = "btnRetrive";
            this.btnRetrive.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRetrive_Click);
            // 
            // btnReleatedProcess
            // 
            this.btnReleatedProcess.Label = "Related Process";
            this.btnReleatedProcess.Name = "btnReleatedProcess";
            this.btnReleatedProcess.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRelatedProcess_Click);
            // 
            // btnValidate
            // 
            this.btnValidate.Label = "Validate";
            this.btnValidate.Name = "btnValidate";
            this.btnValidate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnValidate_Click);
            // 
            // btnHilight
            // 
            this.btnHilight.Label = "Hilight";
            this.btnHilight.Name = "btnHilight";
            this.btnHilight.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnHilight_Click);
            // 
            // btnWebPage
            // 
            this.btnWebPage.Label = "WebPage";
            this.btnWebPage.Name = "btnWebPage";
            this.btnWebPage.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnWebPage_Click);
            // 
            // btnRelatedIntfrastructure
            // 
            this.btnRelatedIntfrastructure.Label = "Related Infrastructure";
            this.btnRelatedIntfrastructure.Name = "btnRelatedIntfrastructure";
            this.btnRelatedIntfrastructure.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRelatedIntfrastructure_Click);
            // 
            // btnRelatedSystem
            // 
            this.btnRelatedSystem.Label = "Related System";
            this.btnRelatedSystem.Name = "btnRelatedSystem";
            this.btnRelatedSystem.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRelatedSystem_Click);
            // 
            // btnNavigateUp
            // 
            this.btnNavigateUp.Label = "Navigate Up";
            this.btnNavigateUp.Name = "btnNavigateUp";
            this.btnNavigateUp.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnNavigateUp_Click);
            // 
            // btnNavigateDown
            // 
            this.btnNavigateDown.Label = "Navigate Down";
            this.btnNavigateDown.Name = "btnNavigateDown";
            this.btnNavigateDown.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnNavigateDown_Click);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Visio.Drawing";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.tabSupportTools);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.tabSupportTools.ResumeLayout(false);
            this.tabSupportTools.PerformLayout();
            this.grpTaskPanes.ResumeLayout(false);
            this.grpTaskPanes.PerformLayout();
            this.grpDebug.ResumeLayout(false);
            this.grpDebug.PerformLayout();
            this.grpHelp.ResumeLayout(false);
            this.grpHelp.PerformLayout();
            this.rgVisio_Utilities.ResumeLayout(false);
            this.rgVisio_Utilities.PerformLayout();
            this.rgSMARTS.ResumeLayout(false);
            this.rgSMARTS.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpSupportTools;
        private Microsoft.Office.Tools.Ribbon.RibbonTab tabSupportTools;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpTaskPanes;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpDebug;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDebugWindow;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnWatchWindow;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkEnableAppEvents;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkDisplayEvents;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpHelp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddInInfo;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDeveloperMode;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAppUtilities;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkDisplayChattyEvents;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup rgVisio_Utilities;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddTableOfContents;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup rgSMARTS;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRetrive;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnWebPage;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnValidate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReleatedProcess;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRelatedSystem;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRelatedIntfrastructure;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnNavigateUp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnNavigateDown;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnHilight;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get
            {
                return this.GetRibbon<Ribbon>();
            }
        }
    }
}
