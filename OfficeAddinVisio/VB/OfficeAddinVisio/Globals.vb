Option Strict Off
Option Explicit On

Imports Microsoft.Office.Interop

Partial Friend NotInheritable Class Globals
    Public Const LOG_SECTION As String = "xNewOfficeAddin"
    ' This should probably come from AssemblyInfo.vb so don't have to carry a
    ' modGlobals file just for this.
    Public Const PROJECT_NAME As String = "OfficeAddinVisio"

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

    Public Shared HAS_COMMON_MENU As Boolean = True
    Public Shared HAS_COMMON_TOOLBAR As Boolean = True
    Public Shared HAS_APPLICATION_MENU As Boolean = False
    Public Shared HAS_APPLICATION_TOOLBAR As Boolean = True

    ' These control if Application Events are traced.  
    ' Useful when learning where to hook code in.

    Public Shared HAS_VISIO_APP_EVENTS As Boolean = True

    ' These control if Application presents additional Command Bars.
    ' Addins that are shard by applications will have multiple lines set true.

    Public Shared HAS_VISIO_CBAR_EVENTS As Boolean = False
    Public Const cMaxSheetNameLen As Integer = 40


#Region "Debug Constants"
    Public Shared cScreenUpdatesOff As Boolean = True
#End Region

    '**********************************************************************
    '   P u b l i c    T y p e s    a n d    V a r i a b l e s
    '**********************************************************************
    ' See Connect for cleanup routines: OnDisconnection,
    ' Menu_Remove, ToolBar_Remove.

    Public Shared HostApp As Object
    Public Shared AddInInstance As Object
    'Public Shared ConnectMode As Extensibility.ext_ConnectMode
    ' Some stuff about who started us and how they started us.
    Public Shared AppName As String
    Public Shared AppVersion As String

    Public Shared StartMode As Short

    ' This are used by application specific code to scope operations that may
    ' be named the same across applications.  e.g. 
    ' With Excel
    '   .Application.operation
    ' End With

    'Public Shared Excel As Microsoft.Office.Interop.Excel.Application

    Public Const cHeaderFontSize As Integer = 12
    Public Const cHeaderFontSizeMedium As Integer = 10
    Public Const cHeaderFontSizeSmall As Integer = 8

    'Public Shared PriorCalculationState As Excel.XlCalculation
    Public Shared PriorScreenUpdatingState As Boolean

    Public Class Colors
        Public Const cBlack As Integer = 1
        Public Const cWhite As Integer = 2
        Public Const cRed As Integer = 3
        Public Const cGreen As Integer = 4
        Public Const cBlue As Integer = 5
        Public Const cYellow As Integer = 6
        Public Const cMagenta As Integer = 7
        Public Const cCyan As Integer = 8

        Public Const cOrange As Integer = 46

        Public Const cLT_TURQUOISE As Integer = 34
        Public Const cLT_GREEN As Integer = 35
        Public Const cROSE As Integer = 38
        Public Const cLT_YELLOW As Integer = 36
        Public Const cTAN As Integer = 40
        Public Const cGOLD As Integer = 44
    End Class


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

    Public Class Regex
        ' SharePoint Folder/File/Document Libraries may not contain any of the following characters
        '   / \ : * ? " < > | <TAB> { } % ~ &
        ' nor may they end in periods or contain embedded double periods.
        ' The following regular expressions capture these rules.
        Public Const cIllegalFileCharacters As String = "[/\\:\*\?""<>\|#\{}%~&]|\.\."  ' SharePoint disallowed
        Public Const cIllegalFolderCharacters As String = "[:\*\?""<>\|#\{}%~&]"   ' SharePoint disallowed
    End Class


    Public Const cMaxFileNameLength As Integer = 128

End Class
