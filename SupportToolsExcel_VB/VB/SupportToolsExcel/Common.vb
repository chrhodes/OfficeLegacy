''' <summary>
''' Common items declared at the Class level.
''' </summary>
''' <remarks>Use this class for any thing you want globally available.  
''' Place only Shared items in this class.  This Class cannot be instantiated.</remarks>
Public NotInheritable Class Common
    Public Const PROJECT_NAME As String = "SupportToolsExcel"

    Public Shared AppEvents As AppEvents

    'Public Shared AppEventsWatchWindow As AddinHelper.WatchWindow

    'Public Shared ExcelHelper As New ExcelHelper.ExcelHelper
    Public Shared ExcelHelper As New AddinHelper.Excel

    ' Routines to add and remove custom task panes and manage their visibility
    ' The Ribbon class does not stay loaded during the lifetime of our addin so keep the task pane
    ' variables here so they stay loaded.

    Public Shared Property TaskPaneConfig() As Microsoft.Office.Tools.CustomTaskPane

    Public Shared Property TaskPaneExcelUtil() As Microsoft.Office.Tools.CustomTaskPane

    Public Shared Property TaskPaneHelp() As Microsoft.Office.Tools.CustomTaskPane

    Public Shared Property TaskPaneITRs() As Microsoft.Office.Tools.CustomTaskPane

    Public Shared Property TaskPaneNetworkTrace() As Microsoft.Office.Tools.CustomTaskPane


    Public Const PROJECT_VERSION As String = "1.0.0"
    Public Const DATA_VERSION As String = "1.0.0"
    Public Const CHART_VERSION As String = "1.0.0"

    ' TODO: Should read this from config file.

    Public Const cDEFAULT_FOLDER As String = "G:\Integration Team"


    Public Shared HAS_APP_EVENTS As Boolean = False
    Public Shared DisplayEvents As Boolean = False

    ' These control if Application presents additional Command Bars.
    ' Addins that are shard by applications will have multiple lines set true.

    Public Shared HAS_EXCEL_CBAR_EVENTS As Boolean = False

    Public Const cMaxFileNameLength As Integer = 128

    'Public Const cDataEntryCell As Integer = cGreen

    ' Charting Colors

    'Public Const cBlack As Integer = 1
    'Public Const cColor2 As Integer = 2
    'Public Const cColor3 As Integer = 6
    'Public Const cRed As Integer = 3
    'Public Const cGreen As Integer = 4
    'Public Const cBlue As Integer = 5
    'Public Const cPink As Integer = 7

    'Public Const cOrange As Integer = 46

    'Public Const cLT_TURQUOISE As Integer = 34
    'Public Const cLT_GREEN As Integer = 35
    'Public Const cROSE As Integer = 38
    'Public Const cLT_YELLOW As Integer = 36
    'Public Const cTAN As Integer = 40
    'Public Const cGOLD As Integer = 44

#Region "Debug constants"

    Public Shared ScreenUpdatesOff As Boolean = True

#End Region

    '------------------------------------------------------------
    '   Constants to control the Worksheets
    '------------------------------------------------------------

    ' Naming convention notes:
    ' 
    ' Constants that refer to a cell should end in _Cell
    ' Constants that represent offsets should end in _Offset
    '
    ' Examples
    '
    ' Cell addresses

    'Public Const cOTD_MetricName_Cell As String = "$C$4"
    'Public Const cOTD_SurveyPeriod_Cell As String = "$C$5"
    'Public Const cOTD_InputFile_Cell As String = "$C$6"
    'Public Const cOTD_InputSheet_Cell As String = "$C$7"

    ' Offsets from file name on On-Time Data worksheet

    'Public Const cOTD_TeamName_Offset As Integer = -8
    'Public Const cOTD_Score_Offset As Integer = -7

    ' Constansts that are relative to a sheet should be prefixed with 
    ' a few letters that represent the sheet name

    ' Examples

    '#Region "Scorecard - Individual Team worksheet (cSCIT_) constants "

    '    'Public Const cSCIT_TeamName_Cell As String = "$B$2"
    '    'Public Const cSCIT_SurveyPeriod_Cell As String = "$E$2"

    '    'Public Const cSCIT_OpenedITRs_Cell As String = "$A$42"
    '    'Public Const cSCIT_ClosedITRs_Cell As String = "$B$42"
    '    'Public Const cSCIT_ActiveITRs_Cell As String = "$C$42"

    '    'Public Const cSCIT_TeamScore_Cells As String = "$F$3:$F$33"

    '#End Region

    '#Region "Scorecard - All Teams worksheet (cSCAT_) constants"

    '    Public Const cSCAT_Results_Cells As String = "$F$3:$T$36"
    '    Public Const cSCAT_Scorecard_Cells As String = "$A$2:$S$36"

    '#End Region


#Region "Sheet Names (cSN_) constants"

    '------------------------------------------------------------
    '   Constants to access the various sheets with names we
    '   want to keep constant because they are referenced in
    '   the code.
    '------------------------------------------------------------

    Public Const cSN_Lookups As String = "Lookups"
    Public Const cSN_Teams As String = "Teams"

    Public Const cSN_X As String = "Sheet X"
    Public Const cSN_Y As String = "Sheet Y"
    Public Const cSN_Z As String = "Sheet Z"

#End Region
    Public Const cITRHeader_Cell As String = "A5"
    Public Const cITRInfo_CommentColumns As String = "$O:$R"
    Public Const cITRITRInfoWithResources_CommentColumns As String = "$O:$R"

    Public Const cFirstITRRow As Integer = 6
    Public Const cFI_SecondITRRow As Integer = 11

    Public Const cAPPLICATION_COLUMN As Integer = 1
    Public Const cITRID_COLUMN As Integer = 2
    Public Const cENTEREDON_COLUMN As Integer = 3
    Private Const cAGE_COLUMN As Integer = 4    ' Added by code
    Public Const cENTEREDBY_COLUMN As Integer = 5
    Public Const cREQUESTEDBY_COLUMN As Integer = 6
    Public Const cRELEASENBR_COLUMN As Integer = 7
    Public Const cPATRANK_COLUMN As Integer = 8
    Public Const cCATEGORY_COLUMN As Integer = 9
    Public Const cSTATUS_COLUMN As Integer = 10
    Public Const cSEVERITY_COLUMN As Integer = 11
    Public Const cLOE_COLUMN As Integer = 12
    Public Const cSUBJECT_COLUMN As Integer = 13
    Public Const cRESOURCEID_COLUMN As Integer = 14
    Public Const cCURRENTCONDITION_COLUMN As Integer = 15
    Public Const cDESIREDOUTCOME_COLUMN As Integer = 16
    Public Const cPRIORITIZATIONCOMMENTS_COLUMN As Integer = 17
    Public Const cCOMMENTS_COLUMN As Integer = 18

#Region "Errors worksheet (cER_) constants"

    Public Const cER_HeaderRow As Integer = 5
    Public Const cER_FirstDataRow As Integer = 6

    Public Const cER_FrameNumber_Column As Integer = 1
    Public Const cER_TimeOfDay_Column As Integer = 2
    Public Const cER_TimeOffset_Column As Integer = 3
    Public Const cER_ConvId_Column As Integer = 4
    Public Const cER_TCPState_Column As Integer = 5
    Public Const cER_Source_Column As Integer = 6
    Public Const cER_Destination_Column As Integer = 7
    Public Const cER_TCPFlags_Column As Integer = 8
    Public Const cER_TCPLength_Column As Integer = 9
    Public Const cER_TCPSeqNumber_Column As Integer = 10
    Public Const cER_TCPAckNumber_Column As Integer = 11
    Public Const cER_TCPNextSeqNumber_Column As Integer = 12
    Public Const cER_WindowSize_Column As Integer = 13
    Public Const cER_Description_Column As Integer = 14

    Public Const cER_FrameNumber_Column_Range As String = "A:A"
    Public Const cER_TimeOfDay_Column_Range As String = "B:B"
    Public Const cER_TimeOffset_Column_Range As String = "C:C"
    Public Const cER_ConvId_Column_Range As String = "D:D"
    Public Const cER_TCPState_Column_Range As String = "E:E"
    Public Const cER_Source_Column_Range As String = "F:F"
    Public Const cER_Destination_Column_Range As String = "G:G"
    Public Const cER_TCPFlags_Column_Range As String = "H:H"
    Public Const cER_TCPLength_Column_Range As String = "I:I"
    Public Const cER_TCPSeqNumber_Column_Range As String = "J:J"
    Public Const cER_TCPAckNumber_Column_Range As String = "K:K"
    Public Const cER_TCPNextSeqNumber_Column_Range As String = "L:L"
    Public Const cER_WindowSize_Column_Range As String = "M:M"
    Public Const cER_Description_Column_Range As String = "N:N"

#End Region

#Region "FormatedITRs worksheet (cFI_) constants"

    Public Const cFI_Application_Column_Range As String = "A:A"
    Public Const cFI_ITRID_Column_Range As String = "B:B"
    Public Const cFI_EnteredOn_Column_Range As String = "C:C"
    Public Const cFI_Age_Column_Range As String = "D:D"
    Public Const cFI_EnteredBy_Column_Range As String = "E:E"
    Public Const cFI_RequestedBy_Column_Range As String = "F:F"
    Public Const cFI_ReleaseNbr_Column_Range As String = "G:G"
    Public Const cFI_PatRank_Column_Range As String = "H:H"
    Public Const cFI_Category_Column_Range As String = "I:I"
    Public Const cFI_Status_Column_Range As String = "J:J"
    Public Const cFI_Severity_Column_Range As String = "K:K"
    Public Const cFI_LOE_Column_Range As String = "L:L"
    Public Const cFI_Subject_Column_Range As String = "M:M"
    Public Const cFI_Resource_Column_Range As String = "N:N"

#End Region

#Region "PivotTable worksheet(s) (cPT_) constants"

    Public Const cPT_ITR_Column_Range As String = "A:A"
    Public Const cPT_Count_Column_Range As String = "B:B"

#End Region


#Region "Lookups worksheet (cLU_) constants "

    ' Used on Lookups worksheet and Teams worksheet.

    Public Const cLU_TeamsInfoCell As String = "$A$5"
    Public Const cLU_ManagerInfoCell As String = "$A$29"

#End Region

#Region "Enumerations"

    Public Enum WrapText As Byte
        Yes = 1
        No = 0
    End Enum

    Public Enum MakeBold As Byte
        Yes = 1
        No = 0
    End Enum

    Public Enum UnderLine As Byte
        Yes = 1
        No = 0
    End Enum

#End Region

    ' TODO: Read this from Config File / TaskPane_Config
    Public Shared TeamName As String = "Integration Services"

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
