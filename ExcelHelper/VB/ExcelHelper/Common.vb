Option Strict Off
Option Explicit On

Imports Microsoft.Office.Interop

Friend NotInheritable Class Common

    Public Const cMaxSheetNameLen As Integer = 40

#Region "Debug Constants"
    'Public Shared ScreenUpdatesToggleEnabled As Boolean = True
#End Region

    Public Shared HostApp As Object
    Public Shared AddInInstance As Object
    ' Some stuff about who started us and how they started us.
    Public Shared AppName As String
    Public Shared AppVersion As String

    Public Shared StartMode As Short

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
End Class

