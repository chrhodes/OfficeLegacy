Option Strict Off
Option Explicit On

Imports Microsoft.Office.Interop

Partial Friend NotInheritable Class Globals
    Public Const LOG_SECTION As String = "xNewOfficeAddin"
    ' This should probably come from AssemblyInfo.vb so don't have to carry a
    ' modGlobals file just for this.
    Public Const PROJECT_NAME As String = "OfficeAddinWord"

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

    Public Shared HAS_WORD_APP_EVENTS As Boolean = True

    ' These control if Application presents additional Command Bars.
    ' Addins that are shard by applications will have multiple lines set true.

    Public Shared HAS_WORD_CBAR_EVENTS As Boolean = False

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
End Class
