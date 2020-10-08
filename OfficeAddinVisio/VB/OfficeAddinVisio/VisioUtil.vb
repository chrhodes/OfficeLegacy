Option Strict Off
Option Explicit On

Imports System.Text

Imports Microsoft.Office.Core
'Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop
Imports PacificLife.Life

Class VisioUtil
    '**********************************************************************
    '   P r i v a t e    C o n s t a n t s
    '**********************************************************************

    Private Const cMODULE_NAME As String = Globals.PROJECT_NAME & ".VisioUtil"

    '**********************************************************************
    '   C o n s t a n t s
    '**********************************************************************

    Const cPointsToInch As Integer = 72    ' Hard coded for Excel? 1/72 Inches.
    Const cStartHour As Integer = 8        ' Times before this
    Const cEndHour As Integer = 20         ' and after this are hilighted.

    '**********************************************************************
    '   P u b l i c    T y p e s    a n d    V a r i a b l e s
    '**********************************************************************


    '**********************************************************************
    '   P r i v a t e    T y p e s    a n d    V a r i a b l e s
    '**********************************************************************

    'Private m_vntPriorCalculationState As Object
    'Private m_vntPriorScreenUpdatingState As Object

    '**********************************************************************
    '   P u b l i c    M e t h o d s
    '**********************************************************************


    Public Shared Sub ApplicationInfo()
        'Try
        '    Debug.Print("Application.CommonAppDataPath:" & Application.CommonAppDataPath.ToString)
        '    Debug.Print("Application.CommonAppDataRegistry:" & Application.CommonAppDataRegistry.ToString)
        '    Debug.Print("Application.CompanyName:" & Application.CompanyName.ToString)
        '    Debug.Print("Application.CurrentCulture:" & Application.CurrentCulture.ToString)
        '    Debug.Print("Application.CurrentInputLanguage:" & Application.CurrentInputLanguage.ToString)
        '    Debug.Print("Application.ExecutablePath:" & Application.ExecutablePath.ToString)
        '    Debug.Print("Application.LocalUserAppDataPath:" & Application.LocalUserAppDataPath.ToString)
        '    Debug.Print("Application.ProductName:" & Application.ProductName.ToString)
        '    Debug.Print("Application.ProductVersion:" & Application.ProductVersion.ToString)
        '    Debug.Print("Application.SafeTopLevelCaptionFormat:" & Application.SafeTopLevelCaptionFormat.ToString)
        '    Debug.Print("Application.StartupPath:" & Application.StartupPath.ToString)
        '    Debug.Print("Application.UserAppDataPath:" & Application.UserAppDataPath.ToString)
        '    Debug.Print("Application.UserAppDataRegistry:" & Application.UserAppDataRegistry.ToString)

        '    Debug.Print("ThisAddin.Application.StartupPath:" & Globals.ThisAddIn.Application.StartupPath.ToString)
        '    Debug.Print("ThisAddin.Application.ActiveWorkbook.Name:" & Globals.ThisAddIn.Application.ActiveWorkbook.Name.ToString)
        '    Debug.Print("ThisAddin.Application.ActiveWorkbook.Path:" & Globals.ThisAddIn.Application.ActiveWorkbook.Path.ToString)
        '    Debug.Print("ThisAddin.Application.ActiveWorkbook.FullName:" & Globals.ThisAddIn.Application.ActiveWorkbook.FullName.ToString)
        '    Debug.Print("ThisAddin.Application.ActiveWorkbook.FullNameURLEncoded:" & Globals.ThisAddIn.Application.ActiveWorkbook.FullNameURLEncoded.ToString)

        '    Debug.Print("ThisAddin.Application.DefaultFilePath:" & Globals.ThisAddIn.Application.DefaultFilePath.ToString)
        '    Debug.Print("ThisAddin.Application.Name:" & Globals.ThisAddIn.Application.Name.ToString)
        '    Debug.Print("ThisAddin.Application.NetworkTemplatesPath:" & Globals.ThisAddIn.Application.NetworkTemplatesPath.ToString)
        '    Debug.Print("ThisAddin.Application.Path:" & Globals.ThisAddIn.Application.Path.ToString)
        'Catch ex As Exception
        '    MessageBox.Show("ApplicationInfo():" & ex.ToString)
        'End Try
    End Sub

    '------------------------------------------------------------
    '
    ' Sub Footer_Add()
    '
    ' ToDo:
    '   Display dialog box that allow format choices.
    '   See Format page for ideas.
    '------------------------------------------------------------

    Public Shared Sub AddFooter()
        'Dim workBook As Excel.Workbook
        'Dim workSheet As Excel.Worksheet
        'Dim sb As StringBuilder = New StringBuilder

        'Try
        '    With Globals.ThisAddIn.Application
        '        workBook = .ActiveWorkbook

        '        VisioUtil.CalculationsOff()

        '        For Each workSheet In workBook.Sheets
        '            With workSheet.PageSetup
        '                sb.Length = 0
        '                ' Five point font, path, and filename
        '                sb.Append("&5&Z&F")
        '                sb.Append(vbLf & "Created: ")
        '                sb.Append(ExcelUtil.GetBuiltInPropertyValue(workBook, "Creation Date"))
        '                sb.Append(vbLf & "Last Saved: ")
        '                sb.Append(ExcelUtil.GetBuiltInPropertyValue(workBook, "Last Save Time"))
        '                sb.Append(vbLf & "Last Printed: ")
        '                sb.Append(ExcelUtil.GetBuiltInPropertyValue(workBook, "Last Print Date"))
        '                'Debug.Print("LeftFooter:>" & sb.ToString & "<")
        '                .LeftFooter = sb.ToString()

        '                .CenterFooter = ""

        '                sb.Length = 0
        '                ' Five point font, page of pages
        '                sb.Append("&5&P - &N")
        '                sb.Append(vbLf & "Title :")
        '                sb.Append(ExcelUtil.GetBuiltInPropertyValue(workBook, "Title"))
        '                sb.Append(vbLf & "Subject: ")
        '                sb.Append(ExcelUtil.GetBuiltInPropertyValue(workBook, "Subject"))
        '                'Debug.Print("RightFooter:>" & sb.ToString & "<")
        '                .RightFooter = sb.ToString()
        '            End With

        '            ' TODO: Indicate we have added a custom footer.  This will be looked for
        '            ' in the before close event.

        '            If Not HasCustomFooter() Then
        '                CustomFooterExists(True)
        '            End If

        '        Next

        '        VisioUtil.CalculationsOn()
        '    End With
        'Catch ex As Exception
        '    MsgBox("AddFooter:" & ex.ToString)
        'End Try
    End Sub ' Footer_Add

    Friend Shared Function HasCustomFooter() As Boolean
        'Dim prp As Office.DocumentProperty
        'Dim prps As Office.DocumentProperties

        'Try
        '    Try
        '        prps = Globals.ThisAddIn.Application.ActiveWorkbook.CustomDocumentProperties
        '        prp = prps.Item("HasCustomFooter")
        '        ' If the property exists we don't really care about the value
        '        Return True

        '    Catch ex As Exception
        '        ' Exception is thrown if property does not exist
        '        Return False
        '    End Try
        'Finally

        'End Try
    End Function

    Friend Shared Sub CustomFooterExists(ByVal hasCustomFooter As Boolean)
        'Dim prp As Office.DocumentProperty
        'Dim prps As Office.DocumentProperties

        'Try
        '    Try
        '        prps = Globals.ThisAddIn.Application.ActiveWorkbook.CustomDocumentProperties
        '        ' Add a new property.
        '        prp = prps.Add("HasCustomFooter", False, _
        '         Office.MsoDocProperties.msoPropertyTypeBoolean, True)
        '    Catch ex As Exception
        '        PLLog.Error(ex, "OfficeAddinExcel")
        '        MessageBox.Show("CustomFooterExists() Unable to add HasCustomFooter property" & ex.Message)
        '    End Try
        'Finally

        'End Try
    End Sub
End Class