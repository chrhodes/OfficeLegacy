''' <summary>
''' Common items declared at the Class level.
''' </summary>
''' <remarks>Use this class for any thing you want globally available.  
''' Place only Shared items in this class.  This Class cannot be instantiated.</remarks>
Public NotInheritable Class Common
    Public Const PROJECT_NAME As String = "SupportToolsWord"
    'Public Const PROJECT_VERSION As String = "1.0.0"
    'Public Const DATA_VERSION As String = "1.0.0"
    'Public Const CHART_VERSION As String = "1.0.0"

    'Public Shared CmdBars As CmdBars
    Public Shared AppEvents As AppEvents

    ' These control if Application Events are traced.  
    ' Useful when learning where to hook code in.

    Public Shared AppEventsWatchWindow As AddinHelper.WatchWindow
    Public Shared HAS_APP_EVENTS As Boolean = False
    Public Shared DisplayEvents As Boolean = False

    Public Shared WordHelper As New WordHelper.WordHelper

    Public Const cDEFAULT_FOLDER As String = "G:\Integration Team"
    Public Shared TeamName As String = "Integration Services"


#Region "Task Panes"
    ' Routines to add and remove custom task panes and manage their visibility
    ' The Ribbon class does not stay loaded during the lifetime of our addin so keep the task pane
    ' variables here so they stay loaded.

    Public Shared Property TaskPaneConfig() As Microsoft.Office.Tools.CustomTaskPane

    Public Shared Property TaskPaneWordUtil() As Microsoft.Office.Tools.CustomTaskPane

    Public Shared Property TaskPaneHelp() As Microsoft.Office.Tools.CustomTaskPane

    Public Shared Property TaskPaneITRs() As Microsoft.Office.Tools.CustomTaskPane

#End Region


    Public Shared Sub Initialize()
        'SwapScreenBaseControl = Nothing
        'InitializedUserControls = Nothing

        ' TODO: Add code to (re)Initialize anything that needs to start clear
        ' when ucBase.Reload() is called.

    End Sub

#Region "Core"

    'Public Shared Property SwapScreenBaseControl() As ucScreenBase

    '''' <summary>
    ''''  This property holds a keyed reference to each instantiated swappable user control.
    ''''  The ucSwapScreenBase loads the control into it's Controls collection to display
    ''''  the screen.  Maintaining the controls in a static Dictionary allows state to be maintained.
    '''' </summary>
    'Private Shared _InitializedUserControls As Dictionary(Of String, ucScreenBase)
    'Public Shared Property InitializedUserControls() As Dictionary(Of String, ucScreenBase)
    '    Get
    '        If _InitializedUserControls Is Nothing Then
    '            _InitializedUserControls = New Dictionary(Of String, ucScreenBase)()
    '        End If

    '        Return _InitializedUserControls
    '    End Get
    '    Set(ByVal value As Dictionary(Of String, ucScreenBase))
    '        _InitializedUserControls = value
    '    End Set
    'End Property

    ' TODO: Add as many DebugLevels as needed.
    ' Add accompanying checkboxes on frmDebugWindow

    Public Shared Property DebugSQL() As Boolean
    Public Shared Property DebugLevel1() As Boolean
    Public Shared Property DebugLevel2() As Boolean
    ''' <summary>
    ''' Indicates whether the UI is running in DeveloperMode
    ''' </summary>
    Public Shared Property DeveloperMode() As Boolean
    ''' <summary>
    ''' Indicates whether the UI is running in DebugMode
    ''' </summary>
    Public Shared Property DebugMode() As Boolean

    Private Shared _debugWindow As System.Windows.Forms.Form

    Public Shared Property DebugWindow() As frmDebugWindow
        Get
            If _debugWindow Is Nothing Then
                _debugWindow = New frmDebugWindow()
            End If
            Return _debugWindow
        End Get
        Set(ByVal value As frmDebugWindow)
            _debugWindow = value
        End Set
    End Property

    Public Shared Sub WriteToDebugWindow(ByVal message As String)
        Dim frm As frmDebugWindow = DebugWindow
        frm.txtOutput.AppendText(Environment.NewLine)
        frm.txtOutput.AppendText(message)
    End Sub

#End Region

    Private Shared _applicationDS As ApplicationDS
    Public Shared Property ApplicationDS As ApplicationDS
        Get
            If _applicationDS Is Nothing Then
                _applicationDS = New ApplicationDS

                ' TODO: Add any other initialization of things related to the ApplicationDS

            End If

            Return _applicationDS
        End Get
        Set(ByVal Value As ApplicationDS)
            _applicationDS = Value
        End Set
    End Property

End Class
