Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Reflection
Imports System.Runtime

Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel

Public Class TaskPane_FarmHealth
    Private Sub TaskPane_FarmHealth_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        For Each host As String In Globals.ThisAddIn.Servers.Hosts
            clbHosts.Items.Add(host)
        Next
    End Sub

    Private Sub btnSayHello_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSayHello.Click
        Dim rng As Range
        Dim ws As Worksheet

        ws = Util.NewWorksheet("Hello Info")
        rng = ws.Range("A5")

        Dim i As Integer = 1

        Util.AddColumnToSheet(ws, 1, 32, "General", False, 5, "Host")
        Util.AddColumnToSheet(ws, 2, 65, "General", False, 5, "Message", )
        Util.AddColumnToSheet(ws, 3, 15, "0", False, 5, "Duration (ms)")

        For Each host In clbHosts.CheckedItems
            Dim startTicks As Long = Stopwatch.GetTimestamp()

            Globals.ThisAddIn.WebService.Url = String.Format("http://{0}/SystemManagement/WMIInfoWS.asmx", host)

            rng.Offset(i, 0).Value = host
            rng.Offset(i, 1).Value = Globals.ThisAddIn.WebService.HelloWorld()
            rng.Offset(i, 2).Value = ((Stopwatch.GetTimestamp - startTicks) / (Stopwatch.Frequency / 1000)).ToString()

            i += 1
        Next
    End Sub

    Private Sub clbHosts_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles clbHosts.DoubleClick

        For i As Integer = 0 To clbHosts.Items.Count - 1
            clbHosts.SetItemChecked(i, True)
        Next

    End Sub

End Class
