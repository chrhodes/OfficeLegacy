Imports System.ComponentModel
Imports System.Reflection
Imports System.Runtime

Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel

Public Class TaskPane_ProcessorInfo
    Private Sub TaskPane_ProcessorInfo_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        For Each host As String In Globals.ThisAddIn.Servers.Hosts
            clbHosts.Items.Add(host)
        Next
    End Sub

    Private Sub btnGetAllInfo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetAllInfo.Click
        Dim rng As Range
        Dim ws As Worksheet
        Dim GB As Long = 1024 * 1024

        ws = Util.NewWorksheet("AllProcessorInfo")
        rng = ws.Range("A5")

        Dim i As Integer = 1

        Util.AddColumnToSheet(ws, 1, 32, "General", False, 5, "Host")
        Util.AddColumnToSheet(ws, 2, 15, "General", False, 5, "AddressWidth", )
        Util.AddColumnToSheet(ws, 3, 32, "General", False, 5, "Caption")
        Util.AddColumnToSheet(ws, 4, 15, "General", False, 5, "DataWidth")
        Util.AddColumnToSheet(ws, 5, 15, "General", False, 5, "L2CacheSize")
        Util.AddColumnToSheet(ws, 6, 17, "General", False, 5, "MaxClockSpeed")
        Util.AddColumnToSheet(ws, 7, 45, "0.00", False, 5, "Name")
        Util.AddColumnToSheet(ws, 8, 15, "0.00", False, 5, "NumberOfCores")
        Util.AddColumnToSheet(ws, 9, 15, "0.00", False, 5, "NumberOfLogicalProcessors")

        For Each host In clbHosts.CheckedItems
            Globals.ThisAddIn.WebService.Url = String.Format("http://{0}/SystemManagement/WMIInfoWS.asmx", host)

            For Each processor As SystemManagementWS.Win32ProcessorS In Globals.ThisAddIn.WebService.GetWin32Processor
                rng.Offset(i, 0).Value = host
                rng.Offset(i, 1).Value = processor.AddressWidth
                rng.Offset(i, 2).Value = processor.Caption
                rng.Offset(i, 3).Value = processor.DataWidth
                rng.Offset(i, 4).Value = processor.L2CacheSize
                rng.Offset(i, 5).Value = processor.MaxClockSpeed
                rng.Offset(i, 6).Value = processor.Name
                rng.Offset(i, 7).Value = processor.NumberOfCores
                rng.Offset(i, 8).Value = processor.NumberOfLogicalProcessors

                i += 1
            Next

            i += 1  ' Skip a row between hosts
        Next
    End Sub

    Private Sub btnGetInfo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetInfo.Click
        Dim rng As Range
        Dim ws As Worksheet
        Dim GB As Long = 1024 * 1024

        ws = Util.NewWorksheet("ProcessorInfo")
        rng = ws.Range("A5")

        Dim i As Integer = 1

        Util.AddColumnToSheet(ws, 1, 32, "General", False, 5, "Host")
        Util.AddColumnToSheet(ws, 2, 32, "General", False, 5, "Caption")
        Util.AddColumnToSheet(ws, 3, 17, "General", False, 5, "MaxClockSpeed")
        Util.AddColumnToSheet(ws, 4, 45, "0.00", False, 5, "Name")

        For Each host In clbHosts.CheckedItems
            Globals.ThisAddIn.WebService.Url = String.Format("http://{0}/SystemManagement/WMIInfoWS.asmx", host)

            For Each processor As SystemManagementWS.Win32ProcessorS In Globals.ThisAddIn.WebService.GetWin32Processor
                rng.Offset(i, 0).Value = host
                rng.Offset(i, 1).Value = processor.Caption
                rng.Offset(i, 2).Value = processor.MaxClockSpeed
                rng.Offset(i, 3).Value = processor.Name

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
