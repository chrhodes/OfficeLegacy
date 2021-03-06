Option Strict Off
Option Explicit On 

Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop
Imports Microsoft.Practices.EnterpriseLibrary.Logging
Imports Microsoft.Practices.EnterpriseLibrary.Logging.ExtraInformation
Imports Microsoft.Practices.EnterpriseLibrary.Logging.Filters
Imports System.Text
Imports PacificLife.Life

Public Class ExcelAppEvents
    '*******************************************************************************
    '
    ' ExcelAppEvents.vb
    '
    ' This class provides handlers for events generated by Excel.
    '
    ' For the DCMAddin we are interested in the WorkbookOpen event.
    '
    ' This file is not likely to change.  The action occurs in ProcesFile.vb
    '
    '*******************************************************************************

    '**********************************************************************
    '   P u b l i c    E v e n t s
    '**********************************************************************

    ' This catches the events from the application
    Public WithEvents ExcelAppEvent As Microsoft.Office.Interop.Excel.Application

    '**********************************************************************
    '   P r i v a t e    C o n s t a n t s
    '**********************************************************************

    Private Const cMODULE_NAME As String = Globals.cPROJECT_NAME & ".ExcelAppEvents"

    '**********************************************************************
    '
    '   M  E  T  H  O  D  S
    '
    '**********************************************************************

    '************************************************************
    '   P r i v a t e      M e t h o d s
    '***********************************************************

    Private Sub ExcelAppEvent_WorkbookOpen(ByVal Wb As Microsoft.Office.Interop.Excel.Workbook) Handles ExcelAppEvent.WorkbookOpen
        PLLog.Trace4("Enter", Globals.cPLLOG_NAME)

        Dim logMessage As String = ""

        ' This is mostly for debugging.  Be sure to test with SharePoint stored files.
#If DEBUG Then
        Dim sb As StringBuilder = New StringBuilder(200)

        sb.Append(String.Format("{0,30}:{1}{2}", "WorkBook.Name", Wb.Name, vbCrLf))
        sb.Append(String.Format("{0,30}:{1}{2}", "WorkBook.FullName", Wb.FullName, vbCrLf))
        sb.Append(String.Format("{0,30}:{1}{2}", "WorkBook.FullNameURLEncoded", Wb.FullNameURLEncoded, vbCrLf))
        sb.Append(String.Format("{0,30}:{1}{2}", "WorkBook.FileFormat", Wb.FileFormat.ToString(), vbCrLf))
        sb.Append(String.Format("{0,30}:{1}{2}", "WorkBook.Path", Wb.Path, vbCrLf))
        Debug.Print(sb.ToString())
#End If

        ' Check to see if we can find a configurtion file in the folder 
        ' containing the file that was just opened.

        If (System.IO.File.Exists(Wb.Path & "\" & Globals.cCONFIG_FILE_NAME)) Then
            logMessage = String.Format("Config File {0} found in {1}.  Loading contents", Globals.cCONFIG_FILE_NAME, Wb.Path)
#If DEBUG Then
            Debug.Print(logMessage)
#End If
            PLLog.Debug(logMessage, Globals.cPLLOG_NAME)

            Try
                Config.LoadFileConfigDataFromXMLFile(Wb.Path & "\" & Globals.cCONFIG_FILE_NAME)
            Catch ex As Exception
                PLLog.Error(ex, Globals.cPLLOG_NAME)
            End Try

            Dim pf As ProcessFile = New ProcessFile

            ' Decide if this is a file we want to process by looking at information
            ' in the configuration file and comparing it to the filename.

            Dim fileNumber As Integer = 0

            If pf.ShouldProcessFile(Config.FileConfigInfo, Wb.Name, fileNumber) Then
                logMessage = "Processing File"
#If DEBUG Then
                Debug.Print(logMessage)
#End If
                PLLog.Debug(logMessage, Globals.cPLLOG_NAME)

                ' Add a new worksheet and load the file using the configuration settings
                Dim ws As Excel.Worksheet = Wb.Worksheets.Add()

                Try
                    pf.ProcessFile(Config.FileConfigInfo, Wb.FullName, ws, fileNumber)
                    PLLog.Info(String.Format("Processed {0}", Wb.FullName), Globals.cPLLOG_NAME)
                Catch ex As Exception
                    PLLog.Error(ex, Globals.cPLLOG_NAME)
                End Try
            Else
                logMessage = "Not Processing File"
#If DEBUG Then
                Debug.Print(logMessage)
#End If
                PLLog.Debug(logMessage, Globals.cPLLOG_NAME)
            End If
        Else
            logMessage = String.Format("No Config File {0} found in {1}.", Globals.cCONFIG_FILE_NAME, Wb.Path)
#If DEBUG Then
            Debug.Print(logMessage)
#End If
            PLLog.Debug(logMessage, Globals.cPLLOG_NAME)
        End If

        PLLog.Trace4("Exit", Globals.cPLLOG_NAME)
    End Sub

End Class