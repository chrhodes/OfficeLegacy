Partial Friend NotInheritable Class Globals
    Public Shared ExcelApp As Excel.Application

    '**********************************************************************
    '   P u b l i c    C o n s t a n t s
    '**********************************************************************

    Public Const cProject_Name As String = "Project X"
    Public Const cProject_Version As String = "1.0.0"
    Public Const cData_Version As String = "1.0.0"
    Public Const cChart_Version As String = "1.0.0"
    Public Const cPLLog_Category_Name As String = "IVRVoiceFilesCleanup"

    ' TODO: Perhaps load these from Config File.

    Public Const cDefault_Output_Folder As String = "C:\temp\"

#Region "Debug constants"

    Public Shared cScreenUpdatesOff As Boolean = True

#End Region

    '------------------------------------------------------------
    '   Constants to control the Worksheets
    '------------------------------------------------------------

#Region "Configuration worksheet constants (cCFG_)"

    Public Const cCFG_TT10_FilePrefix_Cell As String = "$A$5"
    Public Const cCFG_TT10_MessagePrefix_Cell As String = "$B$5"
    Public Const cCFG_TT10_MessageSuffix_Cell As String = "$C$5"

    Public Const cCFG_TT11_FilePrefix_Cell As String = "$A$6"
    Public Const cCFG_TT11_MessagePrefix_Cell As String = "$B$6"
    Public Const cCFG_TT11_MessageSuffix_Cell As String = "$C$6"

    Public Const cCFG_TT12_FilePrefix_Cell As String = "$A$7"
    Public Const cCFG_TT12_MessagePrefix_Cell As String = "$B$7"
    Public Const cCFG_TT12_MessageSuffix_Cell As String = "$C$7"

    Public Const cCFG_TT13_FilePrefix_Cell As String = "$A$8"
    Public Const cCFG_TT13_MessagePrefix_Cell As String = "$B$8"
    Public Const cCFG_TT13_MessageSuffix_Cell As String = "$C$8"

    Public Const cCFG_TT14_FilePrefix_Cell As String = "$A$9"
    Public Const cCFG_TT14_MessagePrefix_Cell As String = "$B$9"
    Public Const cCFG_TT14_MessageSuffix_Cell As String = "$C$9"

    Public Const cCFG_InputWorksheetName_Cell As String = "$B$11"
    Public Const cCFG_InputStartingRow_Cell As String = "$B$12"
    Public Const cCFG_InputEndingRow_Cell As String = "$B$13"
    Public Const cCFG_Input_OutputWorksheetName_Cell As String = "$B$14"
    Public Const cCFG_Input_OutputStartingRow_Cell As String = "$B$15"

    Public Const cCFG_Output_OutputWorksheetName_Cell As String = "$B$17"
    Public Const cCFG_OutputFolderPath_Cell As String = "$B$18"
    Public Const cCFG_OutputStartingRow_Cell As String = "$B$19"
    Public Const cCFG_OutputEndingRow_Cell As String = "$B$20"
    Public Const cCFG_CommandFileName_Cell As String = "$B$21"

    Public Const cSC_Quota_Offset As Integer = 22

    ' Row and Column addresses

    'Public Const cSD_RawDataRow As Integer = 31
    'Public Const cSD_RawDataColumn As Integer = 1

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

    Public Const cHeaderFontSize = 16

End Class

