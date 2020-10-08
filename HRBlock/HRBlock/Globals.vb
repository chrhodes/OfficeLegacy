Partial Friend NotInheritable Class Globals
    '**********************************************************************
    '   P u b l i c    C o n s t a n t s
    '**********************************************************************

    Public Const PROJECT_NAME As String = "HBRBlock"
    Public Const PROJECT_VERSION As String = "1.0.0"
    Public Const DATA_VERSION As String = "1.0.0"
    Public Const CHART_VERSION As String = "1.0.0"

    Public Const cDEFAULT_FOLDER As String = "C:\Temp"

    ' Cross Application Menus and ToolBars.

    Public Const COMMON_MENU_NAME As String = "xNewOfficeAddinMenu"
    Public Const COMMON_MENU_CAPTION As String = "&xNewOfficeAddin"
    Public Const COMMON_MENU_TAG As String = COMMON_MENU_NAME + "_TAG"
    ' The "TB" at the end distinguishes the name from the Application ToolBars, infra.
    Public Const COMMON_TOOLBAR_NAME As String = "xNewOfficeAddinTB"
    Public Const COMMON_TOOLBAR_TAG As String = COMMON_TOOLBAR_NAME + "_TAG"

    ' Application Specific ToolBars.

    Public Const APPLICATION_TOOLBAR_NAME As String = "xNewOfficeAddinToolBar"
    Public Const APPLICATION_TOOLBAR_TAG As String = APPLICATION_TOOLBAR_NAME + "_TAG"

    ' This controls what command bars get created during startup in Connect.OnStartupComplete

    Public Shared HAS_COMMON_MENU As Boolean = False
    Public Shared HAS_COMMON_TOOLBAR As Boolean = False
    Public Shared HAS_APPLICATION_MENU As Boolean = False
    Public Shared HAS_APPLICATION_TOOLBAR As Boolean = False

    ' These control if Application Events are traced.  
    ' Useful when learning where to hook code in.

    Public Shared HAS_EXCEL_APP_EVENTS As Boolean = False

    ' These control if Application presents additional Command Bars.
    ' Addins that are shard by applications will have multiple lines set true.

    Public Shared HAS_EXCEL_CBAR_EVENTS As Boolean = False

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

    Public Shared cScreenUpdatesOff As Boolean = True

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

End Class
