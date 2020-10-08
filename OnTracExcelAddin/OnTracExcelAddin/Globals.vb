Option Strict Off
Option Explicit On

Imports Microsoft.Office.Interop

Partial Friend NotInheritable Class Globals
    '********************************************************************************
    '
    ' $Workfile: Globals.vb $
    ' $Revision: 1 $
    '
    ' Description:
    '   This module contains all globals (Public scope).
    '   There should be no publicly scoped variables in other modules.
    '   There may be publicly scoped variables in class modules.??  Research this.
    '
    ' Public Methods:
    '   Method(arg1, arg2) As Type
    '
    ' Public Types and Variables:
    '   Name as Type
    '
    ' ToDo:
    '   List of ideas for improvement.
    '   Make sure all variables here have g_ in front of them.
    '
    ' $History: Globals.vb $
'
'*****************  Version 1  *****************
'User: Crhodes      Date: 2/02/11    Time: 2:20p
'Created in $/Office/OnTracExcelAddin/OnTracExcelAddin
    '
    '*****************  Version 1  *****************
    'User: Crhodes      Date: 7/20/07    Time: 4:00p
    'Created in $/VSTO/OfficeAddin/OfficeAddin/OfficeAddin
    '
    '********************************************************************************


    '**********************************************************************
    '   E x t e r n a l    F u n c t i o n    D e c l a r a t i o n s
    '**********************************************************************

    '**********************************************************************
    '   P u b l i c    C o n s t a n t s
    '**********************************************************************
    ' ToDo: Change this to something more meaningful for the application.

    Public Const LOG_SECTION As String = "xNewOfficeAddin"
    ' This should probably come from AssemblyInfo.vb so don't have to carry a
    ' modGlobals file just for this.
    Public Const PROJECT_NAME As String = "xNewOfficeAddin"

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

    Public Shared HAS_EXCEL_APP_EVENTS As Boolean = False

    ' These control if Application presents additional Command Bars.
    ' Addins that are shard by applications will have multiple lines set true.

    Public Shared HAS_EXCEL_CBAR_EVENTS As Boolean = False
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

    Public Shared PriorCalculationState As Excel.XlCalculation
    Public Shared PriorScreenUpdatingState As Boolean

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

    '********************************************************************************
    '   End $Workfile: Globals.vb $
    '       $Revision: 1 $
    '********************************************************************************
End Class
