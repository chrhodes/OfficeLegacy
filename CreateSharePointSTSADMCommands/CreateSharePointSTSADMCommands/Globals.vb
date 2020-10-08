Partial Friend NotInheritable Class Globals
    '**********************************************************************
    '   P u b l i c    C o n s t a n t s
    '**********************************************************************

    Public Const cProject_Name As String = "Project X"
    Public Const cProject_Version As String = "1.0.0"
    Public Const cData_Version As String = "1.0.0"
    Public Const cChart_Version As String = "1.0.0"
    Public Const cPLLog_Category_Name As String = "CreateSharePointSTSADMCommands"

    ' TODO: Perhaps load these from Config File.

    Public Const cDefault_Output_Folder As String = "C:\temp\"

#Region "Debug constants"

    Public Shared cScreenUpdatesOff As Boolean = True

#End Region

    '------------------------------------------------------------
    '   Constants to control the Worksheets
    '------------------------------------------------------------

#Region "Project Constants"

    Public Const cSC_TeamNameCell As String = "$B$2"
    Public Const cSharePointSiteCollectionRequestTemplate As String = "C:\temp\New SharePoint Site Collection Request.docx"

#End Region

#Region "CreateSiteCollections worksheet constants (cSC_)"

    Public Const cSC_StartingRow_Cell As String = "$E$2"
    Public Const cSC_EndingRow_Cell As String = "$E$3"
    Public Const cSC_STSADMOutput_Folder_Cell As String = "$E$4"
    Public Const cSC_STSADMOutput_FileName_Cell As String = "$E$5"
    Public Const cSC_WordOutput_Folder_Cell As String = "$E$6"
    Public Const cSC_WordOutput_FileNameBase_Cell As String = "$E$7"

    Public Const cSC_RequestType_Offset As Integer = 0

    Public Const cSC_RequestedBy_Offset As Integer = 1
    Public Const cSC_RequestDate_Offset As Integer = 2
    Public Const cSC_DateNeeded_Offset As Integer = 3
    Public Const cSC_Purpose_Offset As Integer = 4

    Public Const cSC_SharePointFarm_Offset As Integer = 5

    Public Const cSC_WebApplication_Offset As Integer = 6

    Public Const cSC_ManagedPathName_Offset As Integer = 7
    Public Const cSC_ManagedPathType_Offset As Integer = 8

    Public Const cSC_UrlName_Offset As Integer = 9

    Public Const cSC_UniquePermissions_Offset As Integer = 10

    Public Const cSC_Title_Offset As Integer = 11
    Public Const cSC_Description_Offset As Integer = 12

    Public Const cSC_SiteURL_Offset As Integer = 13

    Public Const cSC_WebApplicationURL_Offset As Integer = 14

    Public Const cSC_SiteTemplate_Offset As Integer = 15

    Public Const cSC_PrimaryOwnerLogin_Offset As Integer = 16
    Public Const cSC_PrimaryOwnerName_Offset As Integer = 17
    Public Const cSC_PrimaryOwnerEmail_Offset As Integer = 18

    Public Const cSC_SecondaryOwnerLogin_Offset As Integer = 19
    Public Const cSC_SecondaryOwnerName_Offset As Integer = 20
    Public Const cSC_SecondaryOwnerEmail_Offset As Integer = 21

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

End Class

