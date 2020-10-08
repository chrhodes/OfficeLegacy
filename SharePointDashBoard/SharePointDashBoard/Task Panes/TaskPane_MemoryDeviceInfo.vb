Imports System.ComponentModel
Imports System.Reflection
Imports System.Runtime

Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel

Public Class TaskPane_MemoryDeviceInfo
    
    Private Sub TaskPane_MemoryDeviceInfo_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        For Each host As String In Globals.ThisAddIn.Servers.Hosts
            clbHosts.Items.Add(host)
        Next
    End Sub

    Private Sub btnGetAllInfo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetAllInfo.Click
        Dim rng As Range
        Dim ws As Worksheet
        Dim GB As Long = 1024 * 1024

        ws = Util.NewWorksheet("AllMemoryDeviceInfo")
        rng = ws.Range("A5")

        Dim i As Integer = 1

        Util.AddColumnToSheet(ws, 1, 32, "General", False, 5, "Host")
        Util.AddColumnToSheet(ws, 2, 16, "General", False, 5, "DeviceID")
        Util.AddColumnToSheet(ws, 3, 16, "General", False, 5, "Starting Address", )
        Util.AddColumnToSheet(ws, 4, 16, "General", False, 5, "Ending Address")

        For Each host In clbHosts.CheckedItems
            Globals.ThisAddIn.WebService.Url = String.Format("http://{0}/SystemManagement/WMIInfoWS.asmx", host)

            For Each memoryDevice As SystemManagementWS.Win32MemoryDeviceS In Globals.ThisAddIn.WebService.GetWin32MemoryDevice
                rng.Offset(i, 0).Value = host
                rng.Offset(i, 1).Value = memoryDevice.DeviceID
                rng.Offset(i, 2).Value = memoryDevice.StartingAddress
                rng.Offset(i, 3).Value = memoryDevice.EndingAddress

                i += 1
            Next

            i += 1  ' Skip a row between hosts
        Next
    End Sub

    Private Sub btnGetInfo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetInfo.Click
        Dim rng As Range
        Dim ws As Worksheet
        Dim GB As Long = 1024 * 1024

        ws = Util.NewWorksheet("MemoryDeviceInfo")
        rng = ws.Range("A5")

        Dim i As Integer = 1

        Util.AddColumnToSheet(ws, 1, 32, "General", False, 5, "Host")
        Util.AddColumnToSheet(ws, 2, 16, "General", False, 5, "DeviceID")
        Util.AddColumnToSheet(ws, 3, 16, "General", False, 5, "Starting Address", )
        Util.AddColumnToSheet(ws, 4, 16, "General", False, 5, "Ending Address")

        For Each host In clbHosts.CheckedItems
            Globals.ThisAddIn.WebService.Url = String.Format("http://{0}/SystemManagement/WMIInfoWS.asmx", host)

            For Each memoryDevice As SystemManagementWS.Win32MemoryDeviceS In Globals.ThisAddIn.WebService.GetWin32MemoryDevice
                rng.Offset(i, 0).Value = host
                rng.Offset(i, 1).Value = memoryDevice.DeviceID
                rng.Offset(i, 2).Value = memoryDevice.StartingAddress
                rng.Offset(i, 3).Value = memoryDevice.EndingAddress

                i += 1
            Next

            i += 1  ' Skip a row between hosts
        Next

    End Sub

    Private Sub clbHosts_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles clbHosts.DoubleClick
        For i As Integer = 0 To clbHosts.Items.Count - 1
            clbHosts.SetItemChecked(i, True)
        Next
    End Sub
End Class
