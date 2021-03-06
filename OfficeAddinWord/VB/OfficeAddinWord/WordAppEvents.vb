Option Strict Off
Option Explicit On

Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Word
Imports PacificLife.Life

Public Class WordAppEvents
    ' This catches the events from the application
    Public WithEvents WordApplication As Microsoft.Office.Interop.Word.Application

    Private Const cMODULE_NAME As String = Globals.PROJECT_NAME & ".WordAppEvents"

    Private Sub WordAppEvent_DocumentBeforeClose(ByVal doc As Microsoft.Office.Interop.Word.Document, ByRef Cancel As Boolean) Handles WordApplication.DocumentBeforeClose
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "DocumentBeforeClose")
    End Sub

    Private Sub WordAppEvent_DocumentBeforePrint(ByVal doc As Microsoft.Office.Interop.Word.Document, ByRef Cancel As Boolean) Handles WordApplication.DocumentBeforePrint
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "DocumentBeforePrint")
    End Sub

    Private Sub WordAppEvent_DocumentBeforeSave(ByVal doc As Microsoft.Office.Interop.Word.Document, ByRef SaveAsUI As Boolean, ByRef Cancel As Boolean) Handles WordApplication.DocumentBeforeSave
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "DocumentBeforeSave")
    End Sub

    Private Sub WordAppEvent_DocumentChange() Handles WordApplication.DocumentChange
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "DocumentChange")
    End Sub

    Private Sub WordAppEvent_DocumentOpen(ByVal doc As Microsoft.Office.Interop.Word.Document) Handles WordApplication.DocumentOpen
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "DocumentOpen")
    End Sub

    Private Sub WordApplication_DocumentSync(ByVal Doc As Microsoft.Office.Interop.Word.Document, ByVal SyncEventType As Microsoft.Office.Core.MsoSyncEventType) Handles WordApplication.DocumentSync
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "DocumentSync")
    End Sub

    Private Sub WordAppEvent_EPostageInsert(ByVal doc As Microsoft.Office.Interop.Word.Document) Handles WordApplication.EPostageInsert
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "EPostageInsert")
    End Sub

    Private Sub WordAppEvent_EPostageInsertEx(ByVal Doc As Microsoft.Office.Interop.Word.Document, ByVal cpDeliveryAddrStart As Integer, ByVal cpDeliveryAddrEnd As Integer, ByVal cpReturnAddrStart As Integer, ByVal cpReturnAddrEnd As Integer, ByVal xaWidth As Integer, ByVal yaHeight As Integer, ByVal bstrPrinterName As String, ByVal bstrPaperFeed As String, ByVal fPrint As Boolean, ByRef fCancel As Boolean) Handles WordApplication.EPostageInsertEx
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "EPostageInsertEx")
    End Sub

    Private Sub WordAppEvent_EPostagePropertyDialog(ByVal doc As Microsoft.Office.Interop.Word.Document) Handles WordApplication.EPostagePropertyDialog
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "EPostagePropertyDialog")
    End Sub

    Private Sub WordAppEvent_MailMergeAfterMerge(ByVal doc As Microsoft.Office.Interop.Word.Document, ByVal DocResult As Microsoft.Office.Interop.Word.Document) Handles WordApplication.MailMergeAfterMerge
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "MailMergeAfterMerge")
    End Sub

    Private Sub WordAppEvent_MailMergeAfterRecordMerge(ByVal doc As Microsoft.Office.Interop.Word.Document) Handles WordApplication.MailMergeAfterRecordMerge
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "MailMergeAfterRecordMerge")
    End Sub

    Private Sub WordAppEvent_MailMergeBeforeMerge(ByVal doc As Microsoft.Office.Interop.Word.Document, ByVal StartRecord As Integer, ByVal EndRecord As Integer, ByRef Cancel As Boolean) Handles WordApplication.MailMergeBeforeMerge
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "MailMergeBeforeMerge")
    End Sub

    Private Sub WordAppEvent_MailMergeBeforeRecordMerge(ByVal doc As Microsoft.Office.Interop.Word.Document, ByRef Cancel As Boolean) Handles WordApplication.MailMergeBeforeRecordMerge
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "MailMergeBeforeRecordMerge")
    End Sub

    Private Sub WordAppEvent_MailMergeDataSourceLoad(ByVal doc As Microsoft.Office.Interop.Word.Document) Handles WordApplication.MailMergeDataSourceLoad
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "MailMergeDataSourceLoad")
    End Sub

    Private Sub WordAppEvent_MailMergeDataSourceValidate(ByVal doc As Microsoft.Office.Interop.Word.Document, ByRef Handled As Boolean) Handles WordApplication.MailMergeDataSourceValidate
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "MailMergeDataSourceValidate")
    End Sub

    Private Sub WordApplication_MailMergeDataSourceValidate2(ByVal Doc As Microsoft.Office.Interop.Word.Document, ByRef Handled As Boolean) Handles WordApplication.MailMergeDataSourceValidate2
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "MailMergeDataSourceValidate2")
    End Sub

    Private Sub WordAppEvent_MailMergeWizardSendToCustom(ByVal doc As Microsoft.Office.Interop.Word.Document) Handles WordApplication.MailMergeWizardSendToCustom
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "MailMergeWizardSendToCustom")
    End Sub

    Private Sub WordAppEvent_MailMergeWizardStateChange(ByVal doc As Microsoft.Office.Interop.Word.Document, ByRef FromState As Integer, ByRef ToState As Integer, ByRef Handled As Boolean) Handles WordApplication.MailMergeWizardStateChange
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "MailMergeWizardStateChange")
    End Sub

    Private Sub WordAppEvent_NewDocument(ByVal doc As Microsoft.Office.Interop.Word.Document) Handles WordApplication.NewDocument
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "NewDocument")
    End Sub

    Private Sub WordAppEvent_Quit() Handles WordApplication.Quit
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "Quit")
    End Sub

    Private Sub WordAppEvent_Startup() Handles WordApplication.Startup
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "Startup")
    End Sub

    Private Sub WordAppEvent_WindowActivate(ByVal doc As Microsoft.Office.Interop.Word.Document, ByVal Wn As Microsoft.Office.Interop.Word.Window) Handles WordApplication.WindowActivate
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "WindowActivate")
    End Sub

    Private Sub WordAppEvent_WindowBeforeDoubleClick(ByVal Sel As Microsoft.Office.Interop.Word.Selection, ByRef Cancel As Boolean) Handles WordApplication.WindowBeforeDoubleClick
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "WindowBeforeDoubleClick")
    End Sub

    Private Sub WordAppEvent_WindowBeforeRightClick(ByVal Sel As Microsoft.Office.Interop.Word.Selection, ByRef Cancel As Boolean) Handles WordApplication.WindowBeforeRightClick
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "WindowBeforeRightClick")
    End Sub

    Private Sub WordAppEvent_WindowDeactivate(ByVal doc As Microsoft.Office.Interop.Word.Document, ByVal Wn As Microsoft.Office.Interop.Word.Window) Handles WordApplication.WindowDeactivate
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "WindowDeactivate")
    End Sub

    Private Sub WordAppEvent_WindowSelectionChange(ByVal Sel As Microsoft.Office.Interop.Word.Selection) Handles WordApplication.WindowSelectionChange
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "WindowSelectionChange")
    End Sub

    Private Sub WordAppEvent_WindowSize(ByVal doc As Microsoft.Office.Interop.Word.Document, ByVal Wn As Microsoft.Office.Interop.Word.Window) Handles WordApplication.WindowSize
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "WindowSize")
    End Sub

    Private Sub WordAppEvent_XMLSelectionChange(ByVal Sel As Microsoft.Office.Interop.Word.Selection, ByVal OldXMLNode As Microsoft.Office.Interop.Word.XMLNode, ByVal NewXMLNode As Microsoft.Office.Interop.Word.XMLNode, ByRef Reason As Integer) Handles WordApplication.XMLSelectionChange
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "XMLSelectionChange")
    End Sub

    Private Sub WordAppEvent_XMLValidationError(ByVal XMLNode As Microsoft.Office.Interop.Word.XMLNode) Handles WordApplication.XMLValidationError
        Static i As Short
        i = i + 1
        PLLog.Trace1(i.ToString(), cMODULE_NAME)
        DisplayInWatchWindow(i, "XMLValidationError")
    End Sub

    Private Shared Sub DisplayInWatchWindow(ByVal i As Short, ByVal outputLine As String)
        Dim watchWindow As AddinHelper.WatchWindow = Globals.ThisAddIn.AppEventsWatchWindow

        If Not watchWindow Is Nothing Then
            watchWindow.AddOutputLine(String.Format("{0}:{1}", outputLine, i))
        End If
    End Sub
End Class