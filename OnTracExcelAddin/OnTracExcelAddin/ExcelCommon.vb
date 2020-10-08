Option Strict Off
Option Explicit On

Imports System.Text

Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel
Imports PacificLife.Life

Class ExcelCommon
    '********************************************************************************
    '
    ' $Workfile: ExcelCommon.vb $
    ' $Revision: 1 $
    '
    ' Description:
    '   This module contains ....
    '
    ' Public Methods:
    '   Method(arg1, arg2) As Type
    '
    ' Public Types and Variables:
    '   Note: Put these in modGlobals.bas unless need here.
    '
    ' ToDo:
    '   List of ideas for improvement.
    '
    ' $History: ExcelCommon.vb $
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
    ' Put these in modGlobals.bas


    '**********************************************************************
    '   P u b l i c    C o n s t a n t s
    '**********************************************************************
    ' Put these in modGlobals.bas


    '**********************************************************************
    '   P u b l i c    T y p e s    a n d    V a r i a b l e s
    '**********************************************************************
    ' Put these in modGlobals.bas


    '**********************************************************************
    '   P r i v a t e    C o n s t a n t s
    '**********************************************************************

    Private Const cMODULE_NAME As String = Globals.PROJECT_NAME & ".modExcel"

    '**********************************************************************
    '   P r i v a t e    T y p e s    a n d    V a r i a b l e s
    '**********************************************************************

    '**********************************************************************
    '   P u b l i c    M e t h o d s
    '**********************************************************************

    '------------------------------------------------------------
    '
    ' Sub Footer_Add()
    '
    ' ToDo:
    '   Display dialog box that allow format choices.
    '   See Format page for ideas.
    '------------------------------------------------------------

    Public Shared Sub AddFooter()
        Dim workBook As Workbook
        Dim workSheet As Worksheet
        Dim sb As StringBuilder = New StringBuilder

        Try
            With Globals.ThisAddIn.Application
                workBook = .ActiveWorkbook

                For Each workSheet In workBook.Sheets
                    With workSheet.PageSetup
                        sb.Length = 0
                        ' Five point font, path, and filename
                        sb.Append("&5&Z&F")
                        sb.Append(vbLf & "Created: ")
                        sb.Append(Util.GetBuiltInPropertyValue(workBook, "Creation Date"))
                        sb.Append(vbLf & "Last Saved: ")
                        sb.Append(Util.GetBuiltInPropertyValue(workBook, "Last Save Time"))
                        sb.Append(vbLf & "Last Printed: ")
                        sb.Append(Util.GetBuiltInPropertyValue(workBook, "Last Print Date"))
                        .LeftFooter = sb.ToString()

                        .CenterFooter = ""

                        sb.Length = 0
                        ' Five point font, page of pages
                        sb.Append("&5&P - &N")
                        sb.Append(vbLf & "Title :")
                        sb.Append(Util.GetBuiltInPropertyValue(workBook, "Title"))
                        sb.Append(vbLf & "Subject: ")
                        sb.Append(Util.GetBuiltInPropertyValue(workBook, "Subject"))
                        .RightFooter = sb.ToString()
                    End With
                Next

                ' TODO: Indicate we have added a custom footer.  This will be looked for
                ' in the before close event.

                If Not HasCustomFooter() Then
                    CustomFooterExists(True)
                End If

            End With
        Catch ex As Exception
            'PLLog.Error(ex, "OfficeAddinExcel")
        End Try
    End Sub ' Footer_Add

    Friend Shared Function HasCustomFooter() As Boolean
        Dim prp As Office.DocumentProperty
        Dim prps As Office.DocumentProperties

        Try
            Try
                prps = Globals.ThisAddIn.Application.ActiveWorkbook.CustomDocumentProperties
                prp = prps.Item("HasCustomFooter")
                ' If the property exists we don't really care about the value
                Return True

            Catch ex As Exception
                ' Exception is thrown if property does not exist
                Return False
            End Try
        Finally

        End Try
    End Function

    Friend Shared Sub CustomFooterExists(ByVal hasCustomFooter As Boolean)
        Dim prp As Office.DocumentProperty
        Dim prps As Office.DocumentProperties

        Try
            Try
                prps = Globals.ThisAddIn.Application.ActiveWorkbook.CustomDocumentProperties
                ' Add a new property.
                prp = prps.Add("HasCustomFooter", False, _
                 Office.MsoDocProperties.msoPropertyTypeBoolean, True)
            Catch ex As Exception
                'PLLog.Error(ex, "OfficeAddinExcel")
                MessageBox.Show("CustomFooterExists() Unable to add HasCustomFooter property" & ex.Message)
            End Try
        Finally

        End Try
    End Sub

    Private Shared Sub DumpPropertyCollection( _
     ByVal prps As Office.DocumentProperties, _
     ByVal rng As Range, ByRef i As Integer)
        Dim prp As Office.DocumentProperty

        For Each prp In prps
            rng.Offset(i, 0).Value = prp.Name
            Try
                If Not prp.Value Is Nothing Then
                    rng.Offset(i, 1).Value = _
                     prp.Value.ToString
                End If
            Catch
                ' Do nothing at all.
            End Try
            i += 1
        Next
    End Sub


    Public Shared Sub ZapPageBreaks()
        Dim i As Integer
        Dim sht As Worksheet

        Dim vPB As VPageBreak
        Dim hPB As HPageBreak

        With Globals.ThisAddIn.Application

            For Each sht In .ActiveWorkbook.Sheets
                .ActiveSheet.PageSetup.PrintArea = ""

                Debug.Print(sht.Name)
                '        Debug.Print sht.HPageBreaks.Count
                '        Debug.Print sht.VPageBreaks.Count
                ' For some reason the page break handling is not clean.
                ' There are different types of page breaks, that is clear.
                ' Unfortunately the For Each hPB errors out if only Automatic
                ' Page breaks.  Wrap in try catch for AddIn
                On Error Resume Next
                With sht
                    If .VPageBreaks.Count > 0 Then

                        For Each vPB In .VPageBreaks
                            If vPB.Type = XlPageBreak.xlPageBreakManual Then
                                vPB.Delete()
                            End If
                        Next vPB
                    End If

                    If .HPageBreaks.Count > 0 Then
                        For Each hPB In .HPageBreaks
                            If hPB.Type = XlPageBreak.xlPageBreakManual Then
                                hPB.Delete()
                            End If
                        Next hPB
                    End If

                    '            For i = .HPageBreaks.Count To 1 Step -1
                    ''                Debug.Print .HPageBreaks.Item(i).Type
                    ''                Debug.Print .HPageBreaks.Item(i).Location
                    ''                Debug.Print .HPageBreaks.Item(i).Extent
                    '
                    '                If .HPageBreaks.Item(i).Type = xlPageBreakManual Then
                    '                    .HPageBreaks.Item(i).Delete
                    '                End If
                    '            Next i
                    '
                    '            For i = .VPageBreaks.Count To 1 Step -1
                    ''                Debug.Print .VPageBreaks.Item(i).Type
                    ''                Debug.Print .VPageBreaks.Item(i).Location
                    ''                Debug.Print .VPageBreaks.Item(i).Extent
                    '
                    '                If .VPageBreaks.Item(i).Type = xlPageBreakManual Then
                    '                    .VPageBreaks.Item(i).Delete
                    '                End If
                    '            Next i

                End With
            Next sht
        End With
    End Sub


    '**********************************************************************
    '   P r i v a t e    M e t h o d s
    '**********************************************************************


    '********************************************************************************
    '   End $Workfile: ExcelCommon.vb $
    '       $Revision: 1 $
    '********************************************************************************
End Class