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
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.SupportToolsProject = Me.Factory.CreateRibbonTab
        Me.grpDocument = Me.Factory.CreateRibbonGroup
        Me.grpTaskPanes = Me.Factory.CreateRibbonGroup
        Me.btnITRs = Me.Factory.CreateRibbonButton
        Me.btnProjectUtilities = Me.Factory.CreateRibbonButton
        Me.grpDebug = Me.Factory.CreateRibbonGroup
        Me.cbEnableAppEvents = Me.Factory.CreateRibbonCheckBox
        Me.cbDisplayEvents = Me.Factory.CreateRibbonCheckBox
        Me.grpHelp = Me.Factory.CreateRibbonGroup
        Me.btnAddInInfo = Me.Factory.CreateRibbonButton
        Me.btnDeveloperMode = Me.Factory.CreateRibbonButton
        Me.btnDebugWindow = Me.Factory.CreateRibbonButton
        Me.btnWatchWindow = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout
        Me.SupportToolsProject.SuspendLayout
        Me.grpTaskPanes.SuspendLayout
        Me.grpDebug.SuspendLayout
        Me.grpHelp.SuspendLayout
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Label = "TabAddIns"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Label = "SupportToolsProject"
        Me.Group1.Name = "Group1"
        '
        'SupportToolsProject
        '
        Me.SupportToolsProject.Groups.Add(Me.grpDocument)
        Me.SupportToolsProject.Groups.Add(Me.grpTaskPanes)
        Me.SupportToolsProject.Groups.Add(Me.grpDebug)
        Me.SupportToolsProject.Groups.Add(Me.grpHelp)
        Me.SupportToolsProject.Label = "Support Tools"
        Me.SupportToolsProject.Name = "SupportToolsProject"
        '
        'grpDocument
        '
        Me.grpDocument.Label = "Document"
        Me.grpDocument.Name = "grpDocument"
        '
        'grpTaskPanes
        '
        Me.grpTaskPanes.Items.Add(Me.btnITRs)
        Me.grpTaskPanes.Items.Add(Me.btnProjectUtilities)
        Me.grpTaskPanes.Label = "Task Panes"
        Me.grpTaskPanes.Name = "grpTaskPanes"
        '
        'btnITRs
        '
        Me.btnITRs.Label = "ITRs"
        Me.btnITRs.Name = "btnITRs"
        '
        'btnProjectUtilities
        '
        Me.btnProjectUtilities.Label = "Project Utilities"
        Me.btnProjectUtilities.Name = "btnProjectUtilities"
        '
        'grpDebug
        '
        Me.grpDebug.Items.Add(Me.btnDebugWindow)
        Me.grpDebug.Items.Add(Me.btnWatchWindow)
        Me.grpDebug.Items.Add(Me.cbEnableAppEvents)
        Me.grpDebug.Items.Add(Me.cbDisplayEvents)
        Me.grpDebug.Label = "Debug"
        Me.grpDebug.Name = "grpDebug"
        '
        'cbEnableAppEvents
        '
        Me.cbEnableAppEvents.Label = "Enable App Events"
        Me.cbEnableAppEvents.Name = "cbEnableAppEvents"
        '
        'cbDisplayEvents
        '
        Me.cbDisplayEvents.Label = "Display Events"
        Me.cbDisplayEvents.Name = "cbDisplayEvents"
        '
        'grpHelp
        '
        Me.grpHelp.Items.Add(Me.btnAddInInfo)
        Me.grpHelp.Items.Add(Me.btnDeveloperMode)
        Me.grpHelp.Label = "Help"
        Me.grpHelp.Name = "grpHelp"
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
        'btnDebugWindow
        '
        Me.btnDebugWindow.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnDebugWindow.Image = Global.Office2010Addin_Template_Project.My.Resources.Resources.Auto_Debug_System_icon
        Me.btnDebugWindow.Label = "Debug Window"
        Me.btnDebugWindow.Name = "btnDebugWindow"
        Me.btnDebugWindow.ShowImage = true
        '
        'btnWatchWindow
        '
        Me.btnWatchWindow.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnWatchWindow.Image = Global.Office2010Addin_Template_Project.My.Resources.Resources.WatchWindow
        Me.btnWatchWindow.Label = "Watch Window"
        Me.btnWatchWindow.Name = "btnWatchWindow"
        Me.btnWatchWindow.ShowImage = true
        '
        'Ribbon
        '
        Me.Name = "Ribbon"
        Me.RibbonType = "Microsoft.Project.Project"
        Me.Tabs.Add(Me.Tab1)
        Me.Tabs.Add(Me.SupportToolsProject)
        Me.Tab1.ResumeLayout(false)
        Me.Tab1.PerformLayout
        Me.SupportToolsProject.ResumeLayout(false)
        Me.SupportToolsProject.PerformLayout
        Me.grpTaskPanes.ResumeLayout(false)
        Me.grpTaskPanes.PerformLayout
        Me.grpDebug.ResumeLayout(false)
        Me.grpDebug.PerformLayout
        Me.grpHelp.ResumeLayout(false)
        Me.grpHelp.PerformLayout

End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents SupportToolsProject As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents grpDocument As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents grpTaskPanes As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents grpDebug As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents grpHelp As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnAddInInfo As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnWatchWindow As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents cbEnableAppEvents As Microsoft.Office.Tools.Ribbon.RibbonCheckBox
    Friend WithEvents cbDisplayEvents As Microsoft.Office.Tools.Ribbon.RibbonCheckBox
    Friend WithEvents btnITRs As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnProjectUtilities As Microsoft.Office.Tools.Ribbon.RibbonButton
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
