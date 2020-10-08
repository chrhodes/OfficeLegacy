Partial Class Ribbon
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
   Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.grpSupportTools = Me.Factory.CreateRibbonGroup
        Me.tabSupportTools = Me.Factory.CreateRibbonTab
        Me.grpDocument = Me.Factory.CreateRibbonGroup
        Me.grpTaskPanes = Me.Factory.CreateRibbonGroup
        Me.grpDebug = Me.Factory.CreateRibbonGroup
        Me.chkEnableAppEvents = Me.Factory.CreateRibbonCheckBox
        Me.chkDisplayEvents = Me.Factory.CreateRibbonCheckBox
        Me.grpHelp = Me.Factory.CreateRibbonGroup
        Me.cbEnableAppEvents = Me.Factory.CreateRibbonCheckBox
        Me.Button1 = Me.Factory.CreateRibbonButton
        Me.btnAddFooter = Me.Factory.CreateRibbonButton
        Me.btnTableOfContents = Me.Factory.CreateRibbonButton
        Me.btnProtectAllWorksheets = Me.Factory.CreateRibbonButton
        Me.btnUnProtectAllWorksheets = Me.Factory.CreateRibbonButton
        Me.btnITRs = Me.Factory.CreateRibbonButton
        Me.btnNetworkTraces = Me.Factory.CreateRibbonButton
        Me.btnExcelUtilities = Me.Factory.CreateRibbonButton
        Me.btnDebugWindow = Me.Factory.CreateRibbonButton
        Me.btnWatchWindow = Me.Factory.CreateRibbonButton
        Me.btnAddInInfo = Me.Factory.CreateRibbonButton
        Me.btnDeveloperMode = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout
        Me.grpSupportTools.SuspendLayout
        Me.tabSupportTools.SuspendLayout
        Me.grpDocument.SuspendLayout
        Me.grpTaskPanes.SuspendLayout
        Me.grpDebug.SuspendLayout
        Me.grpHelp.SuspendLayout
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.grpSupportTools)
        Me.Tab1.Label = "TabAddIns"
        Me.Tab1.Name = "Tab1"
        '
        'grpSupportTools
        '
        Me.grpSupportTools.Items.Add(Me.Button1)
        Me.grpSupportTools.Label = "Support Tools"
        Me.grpSupportTools.Name = "grpSupportTools"
        '
        'tabSupportTools
        '
        Me.tabSupportTools.Groups.Add(Me.grpDocument)
        Me.tabSupportTools.Groups.Add(Me.grpTaskPanes)
        Me.tabSupportTools.Groups.Add(Me.grpDebug)
        Me.tabSupportTools.Groups.Add(Me.grpHelp)
        Me.tabSupportTools.Label = "Support Tools"
        Me.tabSupportTools.Name = "tabSupportTools"
        '
        'grpDocument
        '
        Me.grpDocument.Items.Add(Me.btnAddFooter)
        Me.grpDocument.Items.Add(Me.btnTableOfContents)
        Me.grpDocument.Items.Add(Me.btnProtectAllWorksheets)
        Me.grpDocument.Items.Add(Me.btnUnProtectAllWorksheets)
        Me.grpDocument.Label = "Document"
        Me.grpDocument.Name = "grpDocument"
        '
        'grpTaskPanes
        '
        Me.grpTaskPanes.Items.Add(Me.btnITRs)
        Me.grpTaskPanes.Items.Add(Me.btnNetworkTraces)
        Me.grpTaskPanes.Items.Add(Me.btnExcelUtilities)
        Me.grpTaskPanes.Label = "Task Panes"
        Me.grpTaskPanes.Name = "grpTaskPanes"
        '
        'grpDebug
        '
        Me.grpDebug.Items.Add(Me.btnDebugWindow)
        Me.grpDebug.Items.Add(Me.btnWatchWindow)
        Me.grpDebug.Items.Add(Me.chkEnableAppEvents)
        Me.grpDebug.Items.Add(Me.chkDisplayEvents)
        Me.grpDebug.Label = "Debug"
        Me.grpDebug.Name = "grpDebug"
        Me.grpDebug.Visible = false
        '
        'chkEnableAppEvents
        '
        Me.chkEnableAppEvents.Label = "Enable App Events"
        Me.chkEnableAppEvents.Name = "chkEnableAppEvents"
        '
        'chkDisplayEvents
        '
        Me.chkDisplayEvents.Label = "Display Events"
        Me.chkDisplayEvents.Name = "chkDisplayEvents"
        '
        'grpHelp
        '
        Me.grpHelp.Items.Add(Me.btnAddInInfo)
        Me.grpHelp.Items.Add(Me.btnDeveloperMode)
        Me.grpHelp.Label = "Help"
        Me.grpHelp.Name = "grpHelp"
        '
        'cbEnableAppEvents
        '
        Me.cbEnableAppEvents.Label = "Enable App Events"
        Me.cbEnableAppEvents.Name = "cbEnableAppEvents"
        '
        'Button1
        '
        Me.Button1.Label = "Button1"
        Me.Button1.Name = "Button1"
        '
        'btnAddFooter
        '
        Me.btnAddFooter.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnAddFooter.Image = Global.SupportToolsExcel.My.Resources.Resources.add_footer
        Me.btnAddFooter.Label = "Add Footer"
        Me.btnAddFooter.Name = "btnAddFooter"
        Me.btnAddFooter.ShowImage = true
        '
        'btnTableOfContents
        '
        Me.btnTableOfContents.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnTableOfContents.Image = Global.SupportToolsExcel.My.Resources.Resources.table_of_contents
        Me.btnTableOfContents.Label = "Add Table of Contents"
        Me.btnTableOfContents.Name = "btnTableOfContents"
        Me.btnTableOfContents.ShowImage = true
        '
        'btnProtectAllWorksheets
        '
        Me.btnProtectAllWorksheets.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnProtectAllWorksheets.Image = Global.SupportToolsExcel.My.Resources.Resources.protect_sheets
        Me.btnProtectAllWorksheets.Label = "Protect All Worksheets"
        Me.btnProtectAllWorksheets.Name = "btnProtectAllWorksheets"
        Me.btnProtectAllWorksheets.ShowImage = true
        '
        'btnUnProtectAllWorksheets
        '
        Me.btnUnProtectAllWorksheets.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnUnProtectAllWorksheets.Image = Global.SupportToolsExcel.My.Resources.Resources.unprotect_sheets
        Me.btnUnProtectAllWorksheets.Label = "UnProtect All Worksheets"
        Me.btnUnProtectAllWorksheets.Name = "btnUnProtectAllWorksheets"
        Me.btnUnProtectAllWorksheets.ShowImage = true
        '
        'btnITRs
        '
        Me.btnITRs.Label = "ITRs"
        Me.btnITRs.Name = "btnITRs"
        '
        'btnNetworkTraces
        '
        Me.btnNetworkTraces.Label = "Network Traces"
        Me.btnNetworkTraces.Name = "btnNetworkTraces"
        '
        'btnExcelUtilities
        '
        Me.btnExcelUtilities.Label = "Excel Utilities"
        Me.btnExcelUtilities.Name = "btnExcelUtilities"
        '
        'btnDebugWindow
        '
        Me.btnDebugWindow.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnDebugWindow.Image = Global.SupportToolsExcel.My.Resources.Resources.Auto_Debug_System_icon
        Me.btnDebugWindow.Label = "Debug Window"
        Me.btnDebugWindow.Name = "btnDebugWindow"
        Me.btnDebugWindow.ShowImage = true
        '
        'btnWatchWindow
        '
        Me.btnWatchWindow.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnWatchWindow.Image = Global.SupportToolsExcel.My.Resources.Resources.WatchWindow
        Me.btnWatchWindow.Label = "Watch Window"
        Me.btnWatchWindow.Name = "btnWatchWindow"
        Me.btnWatchWindow.ShowImage = true
        '
        'btnAddInInfo
        '
        Me.btnAddInInfo.Label = "AddIn Info"
        Me.btnAddInInfo.Name = "btnAddInInfo"
        '
        'btnDeveloperMode
        '
        Me.btnDeveloperMode.Label = "Developer Mode"
        Me.btnDeveloperMode.Name = "btnDeveloperMode"
        '
        'Ribbon
        '
        Me.Name = "Ribbon"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.Tab1)
        Me.Tabs.Add(Me.tabSupportTools)
        Me.Tab1.ResumeLayout(false)
        Me.Tab1.PerformLayout
        Me.grpSupportTools.ResumeLayout(false)
        Me.grpSupportTools.PerformLayout
        Me.tabSupportTools.ResumeLayout(false)
        Me.tabSupportTools.PerformLayout
        Me.grpDocument.ResumeLayout(false)
        Me.grpDocument.PerformLayout
        Me.grpTaskPanes.ResumeLayout(false)
        Me.grpTaskPanes.PerformLayout
        Me.grpDebug.ResumeLayout(false)
        Me.grpDebug.PerformLayout
        Me.grpHelp.ResumeLayout(false)
        Me.grpHelp.PerformLayout

End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents grpSupportTools As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents tabSupportTools As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents grpDocument As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnAddFooter As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnProtectAllWorksheets As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnUnProtectAllWorksheets As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpTaskPanes As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnITRs As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnNetworkTraces As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnExcelUtilities As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpDebug As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnWatchWindow As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grpHelp As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnAddInInfo As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents chkEnableAppEvents As Microsoft.Office.Tools.Ribbon.RibbonCheckBox
    Friend WithEvents chkDisplayEvents As Microsoft.Office.Tools.Ribbon.RibbonCheckBox
    Friend WithEvents cbEnableAppEvents As Microsoft.Office.Tools.Ribbon.RibbonCheckBox
    Friend WithEvents btnTableOfContents As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnDebugWindow As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnDeveloperMode As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon() As Ribbon
        Get
            Return Me.GetRibbon(Of Ribbon)()
        End Get
    End Property
End Class
