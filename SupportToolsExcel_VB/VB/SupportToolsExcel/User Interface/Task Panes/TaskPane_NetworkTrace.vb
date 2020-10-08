Imports System.Collections
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Reflection
Imports System.Runtime
Imports System.Text
Imports System.Windows.Forms

Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop

Public Class TaskPane_NetworkTrace

#Region "Button Handlers"

    Private Sub btnClearData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearData.Click
        ClearData()
    End Sub

    Private Sub btnCreateAnalysisSheet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateAnalysisSheet.Click
        If txtSheetName.Text.Length > 0 Then
            CreateAnalysisSheet(txtSheetName.Text)
            FormatSheet()
        Else
            MessageBox.Show("SheetName not entered.")
            txtSheetName.Focus()
        End If

    End Sub

    Private Sub btnDetectHosts_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDetectHosts.Click
        txtHostCount.Text = DetectHosts()
    End Sub

    Private Sub btnDuplicateColumns_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDuplicateColumns.Click
        DuplicateColumns()
    End Sub

    Private Sub btnFormatColumns_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFormatColumns.Click
        FormatColumns()
    End Sub

    Private Sub btnFormatSheet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFormatSheet.Click
        FormatSheet()
    End Sub

    Private Sub btnFormatTrace_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFormatTrace.Click
        RemoveHex()
        FormatColumns()
        ClearData()

        HilightTraceSheet()
        HilightTime()
    End Sub

    Private Sub btnHilightErrorSheet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHilightErrorSheet.Click
        HilightErrorSheet()
    End Sub


    Private Sub btnHilightLostFrames_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHilightLostFrames.Click
        HilightLostFrames()
    End Sub

    Private Sub btnHilightTraceSheet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHilightTraceSheet.Click
        HilightTraceSheet()
    End Sub

    Private Sub btnHilightTime_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHilightTime.Click
        HilightTime()
    End Sub

    Private Sub btnRemoveColumns_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemoveColumns.Click
        RemoveColumns()
    End Sub

    Private Sub btnRemoveHex_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemoveHex.Click
        RemoveHex()
    End Sub

#End Region

#Region "Main Methods"

    'Private Sub AddTimeOffset()
    '    Dim ws As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet
    '    ' Find the end of the trace
    '    Dim lastRow As Integer = ws.Range("A5").SpecialCells(Excel.Constants.xlLastCell).Row

    '    With Globals.ThisAddIn.Application
    '        .Columns("C:C").Insert(Shift:=Excel.XlDirection.xlDown, CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove)
    '        .Range("C4").FormulaR1C1 = "Hour Offset"
    '        .Range(.Cells(5, 3), .Cells(lastRow, 3)).FormulaR1C1 = "=RC[-1]/3600"
    '        .Columns("C:C").NumberFormat = "0.00"
    '        .Columns("C:C").ColumnWidth = 5.29
    '    End With
    'End Sub

    Sub ClearData()
        Dim ws As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet
        ' Find the end of the trace
        Dim lastIProw As Integer = ws.Range("A5").SpecialCells(Excel.Constants.xlLastCell).Row
        Dim currentValue As String = ""
        Dim currentIP As String

        ' Walk the list of IPs and clear out the appropriate data.  Assume first row is the
        ' source address.
        Debug.Print(lastIProw)

        Dim sourceAddress As String = ws.Range("E5").Value

        ' TODO: Get rid of magic numbers
        For i As Integer = 5 To lastIProw
            'Debug.Print(String.Format("previousIP >{0}<  currentIP >{1}<   i {2}  lastITRRow {3}", _
            '    previousIP, ws.Cells(i, 5).Value, i, lastIProw))

            currentIP = ws.Cells(i, 5).Value
            Dim r As Excel.Range = ws.Cells(i, 5)

            If currentIP = sourceAddress Then
                ' Clear out Destination
                r.Offset(0, 6).Value = ""
                r.Offset(0, 7).Value = ""
                r.Offset(0, 8).Value = ""
                r.Offset(0, 9).Value = ""
                r.Offset(0, 10).Value = ""
                r.Offset(0, 11).Value = ""
            Else
                ' Clear out Source
                r.Offset(0, 0).Value = ""
                r.Offset(0, 1).Value = ""
                r.Offset(0, 2).Value = ""
                r.Offset(0, 3).Value = ""
                r.Offset(0, 4).Value = ""
                r.Offset(0, 5).Value = ""
            End If
        Next i
    End Sub

    Private Sub CreateAnalysisSheet(ByVal sheetName As String)
        Dim ws As Excel.Worksheet = Common.ExcelHelper.NewWorksheet(sheetName, , Globals.ThisAddIn.Application.ActiveSheet.Name)
        ws.PageSetup.PrintTitleRows = "$5:$5"
        ws.PageSetup.PrintTitleColumns = ""

        With ws.Cells.Font
            .Name = "Calibri"
            .Size = 10
            '.Strikethrough = False
            '.Superscript = False
            '.Subscript = False
            '.OutlineFont = False
            '.Shadow = False
            '.TintAndShade = 0
            '.ThemeFont = xlThemeFontMinor
        End With

        'AddColumnToSheet(ws, 
        '                   columnNumber:=,columnWidth:=,columnWrapText:=,headerRow:=,headerTitle:=,
        '                   headerFontSize:=,headerBold:=,headerUnderline:=,headerWrapText:=,headerHorizontalAlignment,orientation:=)

        ' Don't worry too much about the column widths.  They are adjusted in FormatColumns()
        Common.ExcelHelper.AddColumnToSheet(ws, Common.cER_FrameNumber_Column, 7, False, Common.cER_HeaderRow, "Frame #", 10, True, True, True)
        Common.ExcelHelper.AddColumnToSheet(ws, Common.cER_TimeOfDay_Column, 12, False, Common.cER_HeaderRow, "Time of Day", 10, True, True, , , , "hh:mm:ss;@")
        Common.ExcelHelper.AddColumnToSheet(ws, Common.cER_TimeOffset_Column, 13, False, Common.cER_HeaderRow, "Time Offset", 10, True, True, , , , "0.000000")
        Common.ExcelHelper.AddColumnToSheet(ws, Common.cER_ConvId_Column, 25, False, Common.cER_HeaderRow, "Conv ID", 10, True, True)
        Common.ExcelHelper.AddColumnToSheet(ws, Common.cER_TCPState_Column, 14, False, Common.cER_HeaderRow, "TCP State", 10, True, True)
        Common.ExcelHelper.AddColumnToSheet(ws, Common.cER_Source_Column, 11, False, Common.cER_HeaderRow, "Source", 10, True, True)
        Common.ExcelHelper.AddColumnToSheet(ws, Common.cER_Destination_Column, 11, False, Common.cER_HeaderRow, "Destination", 10, True, True)
        Common.ExcelHelper.AddColumnToSheet(ws, Common.cER_TCPFlags_Column, 7, False, Common.cER_HeaderRow, "TCP Flags", 10, True, True)
        Common.ExcelHelper.AddColumnToSheet(ws, Common.cER_TCPLength_Column, 6, False, Common.cER_HeaderRow, "Length", 10, True, True)
        Common.ExcelHelper.AddColumnToSheet(ws, Common.cER_TCPSeqNumber_Column, 11, False, Common.cER_HeaderRow, "TCP Seq #", 10, True, True)
        Common.ExcelHelper.AddColumnToSheet(ws, Common.cER_TCPAckNumber_Column, 11, False, Common.cER_HeaderRow, "TCP Acq #", 10, True, True)
        Common.ExcelHelper.AddColumnToSheet(ws, Common.cER_TCPNextSeqNumber_Column, 11, False, Common.cER_HeaderRow, "TCP Next Seq #", 10, True, True)
        Common.ExcelHelper.AddColumnToSheet(ws, Common.cER_WindowSize_Column, 11, False, Common.cER_HeaderRow, "Window Size", 10, True, True)
        Common.ExcelHelper.AddColumnToSheet(ws, Common.cER_Description_Column, 50, False, Common.cER_HeaderRow, "Description", 10, True, True)

        ' This is for the analysis data

        ws.Range("A3").Value = "Packets"
        ws.Range("B3").Value = 1    ' Just to make the formula below happy
        ws.Range("B3").NumberFormat = "0"

        ws.Range("C3").FormulaR1C1 = "=RC[2]/R3C2"
        ws.Range("C4").FormulaR1C1 = "=RC[2]/R3C2"
        ws.Range("C3:C4").Style = "Percent"
        ws.Range("C3:C4").NumberFormat = "0.000%"

        ws.Range("D3").Value = "Resets"
        ws.Range("D4").Value = "Retransmissions"

        ws.Range("E2").Value = "Total"

        ws.Range("F1").Value = "Source"

        ws.Range("F2").Value = "MainFrame"
        ws.Range("G2").Value = "LSBizA01CLv"
        ws.Range("H2").Value = "LSDCM02v"
        ws.Range("I2").Value = "Other"

        ws.Range("J1").Value = "Destination"

        ws.Range("J2").Value = "MainFrame"
        ws.Range("K2").Value = "LSBizA01CLv"
        ws.Range("L2").Value = "LSDCM02v"
        ws.Range("M2").Value = "Other"

        With ws.Range("D3:M3").Interior
            .Pattern = Excel.Constants.xlSolid
            .PatternColorIndex = Excel.Constants.xlAutomatic
            .Color = 255        ' Red
        End With

        With ws.Range("D4:M4").Interior
            .Pattern = Excel.Constants.xlSolid
            .PatternColorIndex = Excel.Constants.xlAutomatic
            .Color = 49407          ' Orange
        End With
    End Sub

    ' Walk the trace looking for unique host(IP) names

    Private Function DetectHosts() As Integer
        Dim ws As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet
        ' Find the end of the trace
        Dim lastTraceRow As Integer = ws.Range("A5").SpecialCells(Excel.Constants.xlLastCell).Row
        Dim hosts As Dictionary(Of String, Integer) = New Dictionary(Of String, Integer)

        Dim sourceHost As String = ""
        Dim destinationHost As String = ""

        ' TODO: Get rid of magic numbers
        For i As Integer = 5 To lastTraceRow
            sourceHost = ws.Cells(i, 5).Value
            destinationHost = ws.Cells(i, 6).Value

            If Not hosts.ContainsKey(sourceHost) Then
                hosts.Add(sourceHost, hosts.Count + 1)
            End If

            If Not hosts.ContainsKey(destinationHost) Then
                hosts.Add(destinationHost, hosts.Count + 1)
            End If
        Next i

        'For Each host In hosts
        '    Debug.Print(host.Key, host.Value)
        'Next

        Return hosts.Count

    End Function

    Sub DuplicateColumns()
        Dim hostCount As Integer = CInt(txtHostCount.Text)

        With Globals.ThisAddIn.Application
            For i As Integer = 1 To hostCount - 1
                .Columns("E:K").Select()
                .Selection.Copy()
                .Columns((i * 7) + 5).Select()
                .Selection.Insert(Shift:=Excel.XlDirection.xlToRight)
            Next i
        End With
    End Sub

    Sub FixDestinationAddresses()
        Dim ws As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet
        ' Find the end of the trace
        Dim lastIProw As Integer = ws.Range("A5").SpecialCells(Excel.Constants.xlLastCell).Row
        Dim currentValue As String = ""
        Dim currentIP As String

        ' Walk the list of addresses and delete the destination rows.  

        Dim sourceAddress As String = ws.Range("E5").Value

        ' TODO: Get rid of magic numbers
        For i As Integer = 5 To lastIProw
            'Debug.Print(String.Format("previousIP >{0}<  currentIP >{1}<   i {2}  lastITRRow {3}", _
            '    previousIP, ws.Cells(i, 5).Value, i, lastIProw))

            currentIP = ws.Cells(i, 5).Value
            Dim r As Excel.Range = ws.Cells(i, 5)

            If currentIP = sourceAddress Then
                ' Clear out Source
                r.Offset(0, 0).Value = ""
                r.Offset(0, 1).Value = ""
                r.Offset(0, 2).Value = ""
                r.Offset(0, 3).Value = ""
                r.Offset(0, 4).Value = ""
                r.Offset(0, 5).Value = ""
            Else
                ' Clear out Destination
                r.Offset(0, 6).Value = ""
                r.Offset(0, 7).Value = ""
                r.Offset(0, 8).Value = ""
                r.Offset(0, 9).Value = ""
                r.Offset(0, 10).Value = ""
                r.Offset(0, 11).Value = ""
            End If

            'previousIP = currentIP
        Next i
    End Sub

    Private Sub FormatColumns()
        With Globals.ThisAddIn.Application
            .Columns(Common.cER_FrameNumber_Column_Range).ColumnWidth = 7

            .Columns(Common.cER_TimeOfDay_Column_Range).ColumnWidth = 12
            '.Columns(Common.cER_TimeOfDay_Column_Range).NumberFormat = "hh:mm:ss;@"

            .Columns(Common.cER_TimeOffset_Column_Range).ColumnWidth = 13
            '.Columns(Common.cER_TimeOffset_Column_Range).NumberFormat = "0.000000"

            .Columns(Common.cER_ConvId_Column_Range).ColumnWidth = 25
            '.Columns("C:C").Format.Font = 10

            .Columns(Common.cER_TCPState_Column_Range).ColumnWidth = 14

            .Columns(Common.cER_Source_Column_Range).ColumnWidth = 11

            .Columns(Common.cER_Destination_Column_Range).ColumnWidth = 11

            .Columns(Common.cER_TCPFlags_Column_Range).ColumnWidth = 7

            .Columns(Common.cER_TCPLength_Column_Range).ColumnWidth = 6

            .Columns(Common.cER_TCPSeqNumber_Column_Range).ColumnWidth = 11

            .Columns(Common.cER_TCPAckNumber_Column_Range).ColumnWidth = 11

            .Columns(Common.cER_TCPNextSeqNumber_Column_Range).ColumnWidth = 11

            .Columns(Common.cER_WindowSize_Column_Range).ColumnWidth = 11

            .Columns(Common.cER_Description_Column_Range).ColumnWidth = 50

            .Range("A2").Select()
        End With

    End Sub

    Private Sub FormatSheet()
        With Globals.ThisAddIn.Application
            With .ActiveSheet.PageSetup
                .PrintTitleRows = "$5:$5"
                .PrintTitleColumns = ""
                .PrintArea = ""
                .LeftHeader = "&F"
                .CenterHeader = ""
                .RightHeader = "&A"
                .LeftFooter = ""
                .CenterFooter = ""
                .RightFooter = ""
                .LeftMargin = .Application.InchesToPoints(0.25)
                .RightMargin = .Application.InchesToPoints(0.25)
                .TopMargin = .Application.InchesToPoints(0.75)
                .BottomMargin = .Application.InchesToPoints(0.75)
                .HeaderMargin = .Application.InchesToPoints(0.3)
                .FooterMargin = .Application.InchesToPoints(0.3)
                .PrintHeadings = False
                .PrintGridlines = True
                .PrintComments = Excel.XlPrintLocation.xlPrintNoComments
                .PrintQuality = 600
                .CenterHorizontally = False
                .CenterVertically = False
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .Draft = False
                .PaperSize = Excel.XlPaperSize.xlPaperLegal
                .FirstPageNumber = Excel.Constants.xlAutomatic
                .Order = Excel.XlOrder.xlDownThenOver
                .BlackAndWhite = False
                .Zoom = 100
                .PrintErrors = Excel.XlPrintErrors.xlPrintErrorsDisplayed
                .OddAndEvenPagesHeaderFooter = False
                .DifferentFirstPageHeaderFooter = False
                .ScaleWithDocHeaderFooter = True
                .AlignMarginsHeaderFooter = True
                .EvenPage.LeftHeader.Text = ""
                .EvenPage.CenterHeader.Text = ""
                .EvenPage.RightHeader.Text = ""
                .EvenPage.LeftFooter.Text = ""
                .EvenPage.CenterFooter.Text = ""
                .EvenPage.RightFooter.Text = ""
                .FirstPage.LeftHeader.Text = ""
                .FirstPage.CenterHeader.Text = ""
                .FirstPage.RightHeader.Text = ""
                .FirstPage.LeftFooter.Text = ""
                .FirstPage.CenterFooter.Text = ""
                .FirstPage.RightFooter.Text = ""
            End With

            '.Application.PrintCommunication = True
        End With
    End Sub

    Private Sub HilightErrorSheet()
        Dim ws As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet
        ' Find the end of the trace
        Dim lastIProw As Integer = ws.Range("A5").SpecialCells(Excel.Constants.xlLastCell).Row
        Dim source As String
        Dim destination As String
        Dim flags As String = ""
        Dim description As String = ""
        Dim countResets As Integer = 0
        Dim countRetransmissions As Integer = 0
        Dim countResetsDestination As Dictionary(Of String, Integer) = New Dictionary(Of String, Integer)
        Dim countRetransmissionsDestination As Dictionary(Of String, Integer) = New Dictionary(Of String, Integer)
        Dim countResetsSource As Dictionary(Of String, Integer) = New Dictionary(Of String, Integer)
        Dim countRetransmissionsSource As Dictionary(Of String, Integer) = New Dictionary(Of String, Integer)

        ' May want to put this on the CreateAnalysisSheet()
        'ws.Range("D2").Value = "Resets"
        'ws.Range("D3").Value = "Retransmissions"

        'ws.Range("E1").Value = "Total"
        'ws.Range("F1").Value = "MainFrame"
        'ws.Range("G1").Value = "LSBizA01CLv"
        'ws.Range("H1").Value = "LSDCM02v"
        'ws.Range("I1").Value = "Other"

        'ws.Range("J1").Value = "MainFrame"
        'ws.Range("K1").Value = "LSBizA01CLv"
        'ws.Range("L1").Value = "LSDCM02v"
        'ws.Range("M1").Value = "Other"

        'With ws.Range("D2:M2").Interior
        '    .Pattern = Excel.Constants.xlSolid
        '    .PatternColorIndex = Excel.Constants.xlAutomatic
        '    .Color = 255        ' Red
        'End With

        'With ws.Range("D3:M3").Interior
        '    .Pattern = Excel.Constants.xlSolid
        '    .PatternColorIndex = Excel.Constants.xlAutomatic
        '    .Color = 49407          ' Orange
        'End With

        ' Initialize the Dictionaries
        countResetsSource.Add("MainFrame", 0)
        countResetsSource.Add("LSBizA01CLv", 0)
        countResetsSource.Add("LSDCM02v", 0)
        countResetsSource.Add("Other", 0)

        countRetransmissionsSource.Add("MainFrame", 0)
        countRetransmissionsSource.Add("LSBizA01CLv", 0)
        countRetransmissionsSource.Add("LSDCM02v", 0)
        countRetransmissionsSource.Add("Other", 0)

        countResetsDestination.Add("MainFrame", 0)
        countResetsDestination.Add("LSBizA01CLv", 0)
        countResetsDestination.Add("LSDCM02v", 0)
        countResetsDestination.Add("Other", 0)

        countRetransmissionsDestination.Add("MainFrame", 0)
        countRetransmissionsDestination.Add("LSBizA01CLv", 0)
        countRetransmissionsDestination.Add("LSDCM02v", 0)
        countRetransmissionsDestination.Add("Other", 0)

        ' TODO: Get rid of magic numbers
        For i As Integer = Common.cER_FirstDataRow To lastIProw
            source = ws.Cells(i, Common.cER_Source_Column).Value
            destination = ws.Cells(i, Common.cER_Destination_Column).Value
            flags = ws.Cells(i, Common.cER_TCPFlags_Column).Value
            description = ws.Cells(i, Common.cER_Description_Column).Value

            If Not flags Is Nothing Then
                If flags.Contains("R") Then
                    countResets += 1

                    Select Case source
                        Case "MainFrame", "LSBizA01CLv", "LSDCM02v"
                            countResetsSource(source) += 1

                            With ws.Range(ws.Cells(i, Common.cER_Source_Column), ws.Cells(i, Common.cER_TCPFlags_Column)).Interior
                                .Pattern = Excel.Constants.xlSolid
                                .PatternColorIndex = Excel.Constants.xlAutomatic
                                .Color = 255        ' Red
                            End With

                        Case Else
                            countResetsSource("Other") += 1

                            With ws.Range(ws.Cells(i, Common.cER_TCPFlags_Column), ws.Cells(i, Common.cER_TCPFlags_Column)).Interior
                                .Pattern = Excel.Constants.xlSolid
                                .PatternColorIndex = Excel.Constants.xlAutomatic
                                .Color = 255        ' Red
                            End With

                            ' Clear the color in case we copied from a formatted file
                            With ws.Range(ws.Cells(i, Common.cER_Source_Column), ws.Cells(i, Common.cER_Destination_Column)).Interior
                                .Pattern = Excel.Constants.xlNone
                                .TintAndShade = 0
                                .PatternTintAndShade = 0
                            End With

                    End Select

                    Select Case destination
                        Case "MainFrame", "LSBizA01CLv", "LSDCM02v"
                            countResetsDestination(destination) += 1

                        Case Else
                            countResetsDestination("Other") += 1

                            ' Clear the color in case we copied from a formatted file
                            With ws.Range(ws.Cells(i, Common.cER_Source_Column), ws.Cells(i, Common.cER_Destination_Column)).Interior
                                .Pattern = Excel.Constants.xlNone
                                .TintAndShade = 0
                                .PatternTintAndShade = 0
                            End With

                    End Select
                End If
            End If

            If description.Contains("ReTransmit") Then
                countRetransmissions += 1

                Select Case source
                    Case "MainFrame", "LSBizA01CLv", "LSDCM02v"
                        countRetransmissionsSource(source) += 1

                        With ws.Range(ws.Cells(i, Common.cER_Source_Column), ws.Cells(i, Common.cER_TCPFlags_Column)).Interior
                            .Pattern = Excel.Constants.xlSolid
                            .PatternColorIndex = Excel.Constants.xlAutomatic
                            .Color = 49407          ' Orange
                        End With

                    Case Else
                        countRetransmissionsSource("Other") += 1

                        With ws.Range(ws.Cells(i, Common.cER_TCPFlags_Column), ws.Cells(i, Common.cER_TCPFlags_Column)).Interior
                            .Pattern = Excel.Constants.xlSolid
                            .PatternColorIndex = Excel.Constants.xlAutomatic
                            .Color = 49407          ' Orange
                        End With

                        ' Clear the color in case we copied from a formatted file
                        With ws.Range(ws.Cells(i, Common.cER_Source_Column), ws.Cells(i, Common.cER_Destination_Column)).Interior
                            .Pattern = Excel.Constants.xlNone
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With

                End Select

                Select Case destination
                    Case "MainFrame", "LSBizA01CLv", "LSDCM02v"
                        countRetransmissionsDestination(destination) += 1

                    Case Else
                        countRetransmissionsDestination("Other") += 1

                        ' Clear the color in case we copied from a formatted file
                        With ws.Range(ws.Cells(i, Common.cER_Source_Column), ws.Cells(i, Common.cER_Destination_Column)).Interior
                            .Pattern = Excel.Constants.xlNone
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With

                End Select
            End If
        Next i

        ws.Range("E3").Value = countResets
        ws.Range("E4").Value = countRetransmissions

        'For Each counter In countResetsSource
        '    Debug.Print(String.Format(">{0}< {1}", counter.Key, counter.Value))
        'Next

        'For Each counter In countRetransmissionsSource
        '    Debug.Print(String.Format(">{0}< {1}", counter.Key, counter.Value))
        'Next

        ws.Range("F3").Value = countResetsSource("MainFrame")
        ws.Range("G3").Value = countResetsSource("LSBizA01CLv")
        ws.Range("H3").Value = countResetsSource("LSDCM02v")
        ws.Range("I3").Value = countResetsSource("Other")

        ws.Range("F4").Value = countRetransmissionsSource("MainFrame")
        ws.Range("G4").Value = countRetransmissionsSource("LSBizA01CLv")
        ws.Range("H4").Value = countRetransmissionsSource("LSDCM02v")
        ws.Range("I4").Value = countRetransmissionsSource("Other")

        ws.Range("J3").Value = countResetsDestination("MainFrame")
        ws.Range("K3").Value = countResetsDestination("LSBizA01CLv")
        ws.Range("L3").Value = countResetsDestination("LSDCM02v")
        ws.Range("M3").Value = countResetsDestination("Other")

        ws.Range("J4").Value = countRetransmissionsDestination("MainFrame")
        ws.Range("K4").Value = countRetransmissionsDestination("LSBizA01CLv")
        ws.Range("L4").Value = countRetransmissionsDestination("LSDCM02v")
        ws.Range("M4").Value = countRetransmissionsDestination("Other")
    End Sub

    Private Sub HilightLostFrames()
        Dim ws As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet
        ' Find the end of the trace
        Dim lastIProw As Integer = ws.Range("A5").SpecialCells(Excel.Constants.xlLastCell).Row
        Dim source As String
        Dim destination As String
        Dim flags As String = ""
        Dim description As String = ""
        Dim length As Integer
        Dim seqNumber As Long
        Dim ackNumber As Long
        Dim nextSeqNumber As Long

        Dim countResets As Integer = 0
        Dim countRetransmissions As Integer = 0

        Dim nextPacketSource As Dictionary(Of String, Long) = New Dictionary(Of String, Long)
        Dim highestPacketSource As Dictionary(Of String, Long) = New Dictionary(Of String, Long)
        Dim lastPacketRow As Dictionary(Of String, Integer) = New Dictionary(Of String, Integer)

        'Dim countResetsDestination As Dictionary(Of String, Integer) = New Dictionary(Of String, Integer)
        'Dim countRetransmissionsDestination As Dictionary(Of String, Integer) = New Dictionary(Of String, Integer)
        'Dim countResetsSource As Dictionary(Of String, Integer) = New Dictionary(Of String, Integer)
        'Dim countRetransmissionsSource As Dictionary(Of String, Integer) = New Dictionary(Of String, Integer)

        ' May want to put this on the CreateAnalysisSheet()
        'ws.Range("D2").Value = "Resets"
        'ws.Range("D3").Value = "Retransmissions"

        'ws.Range("E1").Value = "Total"
        'ws.Range("F1").Value = "MainFrame"
        'ws.Range("G1").Value = "LSBizA01CLv"
        'ws.Range("H1").Value = "LSDCM02v"
        'ws.Range("I1").Value = "Other"

        'ws.Range("J1").Value = "MainFrame"
        'ws.Range("K1").Value = "LSBizA01CLv"
        'ws.Range("L1").Value = "LSDCM02v"
        'ws.Range("M1").Value = "Other"

        'With ws.Range("D2:M2").Interior
        '    .Pattern = Excel.Constants.xlSolid
        '    .PatternColorIndex = Excel.Constants.xlAutomatic
        '    .Color = 255        ' Red
        'End With

        'With ws.Range("D3:M3").Interior
        '    .Pattern = Excel.Constants.xlSolid
        '    .PatternColorIndex = Excel.Constants.xlAutomatic
        '    .Color = 49407          ' Orange
        'End With

        ' Initialize the Dictionaries

        nextPacketSource.Add("MainFrame", 0)
        nextPacketSource.Add("LSBizA01CLv", 0)
        nextPacketSource.Add("LSDCM02v", 0)

        lastPacketRow.Add("MainFrame", 0)
        lastPacketRow.Add("LSBizA01CLv", 0)
        lastPacketRow.Add("LSDCM02v", 0)

        highestPacketSource.Add("MainFrame", 0)
        highestPacketSource.Add("LSBizA01CLv", 0)
        highestPacketSource.Add("LSDCM02v", 0)

        ' TODO: Get rid of magic numbers
        For i As Integer = Common.cER_FirstDataRow To lastIProw
            source = ws.Cells(i, Common.cER_Source_Column).Value
            destination = ws.Cells(i, Common.cER_Destination_Column).Value
            flags = ws.Cells(i, Common.cER_TCPFlags_Column).Value
            description = ws.Cells(i, Common.cER_Description_Column).Value
            length = ws.Cells(i, Common.cER_TCPLength_Column).Value
            seqNumber = ws.Cells(i, Common.cER_TCPSeqNumber_Column).Value
            ackNumber = ws.Cells(i, Common.cER_TCPAckNumber_Column).Value
            nextSeqNumber = ws.Cells(i, Common.cER_TCPNextSeqNumber_Column).Value

            If Not flags Is Nothing Then


                If flags.Contains("S") Then
                    nextPacketSource(source) = nextSeqNumber
                    lastPacketRow(source) = i
                    highestPacketSource(source) = nextSeqNumber
                Else

                    If seqNumber < highestPacketSource(source) Then
                        With ws.Range(ws.Cells(i, Common.cER_Source_Column), ws.Cells(i, Common.cER_TCPSeqNumber_Column)).Interior
                            .Pattern = Excel.Constants.xlSolid
                            .PatternColorIndex = Excel.Constants.xlAutomatic
                            .ThemeColor = Excel.XlThemeColor.xlThemeColorAccent4    ' Purples
                            .TintAndShade = 0.399975585192419
                        End With
                    ElseIf nextPacketSource(source) <> seqNumber Then

                        With ws.Range(ws.Cells(lastPacketRow(source), Common.cER_TCPNextSeqNumber_Column), ws.Cells(lastPacketRow(source), Common.cER_TCPNextSeqNumber_Column)).Interior
                            .Pattern = Excel.Constants.xlSolid
                            .PatternColorIndex = Excel.Constants.xlAutomatic
                            .ThemeColor = Excel.XlThemeColor.xlThemeColorAccent2    ' Reds
                            .TintAndShade = 0.399975585192419
                        End With

                        With ws.Range(ws.Cells(i, Common.cER_Source_Column), ws.Cells(i, Common.cER_TCPSeqNumber_Column)).Interior
                            .Pattern = Excel.Constants.xlSolid
                            .PatternColorIndex = Excel.Constants.xlAutomatic
                            '.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent6    ' Oranges
                            .ThemeColor = Excel.XlThemeColor.xlThemeColorAccent2    ' Reds
                            .TintAndShade = 0.399975585192419
                        End With
                    End If

                    nextPacketSource(source) = nextSeqNumber
                    lastPacketRow(source) = i


                    If nextSeqNumber > highestPacketSource(source) Then
                        highestPacketSource(source) = nextSeqNumber
                    End If

                End If
            End If

        Next i

    End Sub

    Private Sub HilightTraceSheet()
        Dim ws As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet
        ' Find the end of the trace
        Dim lastIProw As Integer = ws.Range("A5").SpecialCells(Excel.Constants.xlLastCell).Row
        Dim flags As String = ""
        Dim description As String = ""

        ' TODO: Get rid of magic numbers
        For i As Integer = 5 To lastIProw
            flags = ws.Cells(i, 8).Value
            description = ws.Cells(i, 13).Value

            If Not flags Is Nothing Then
                If flags.Contains("S") Then
                    With ws.Range(ws.Cells(i, 6), ws.Cells(i, 7)).Interior
                        'With ws.Cells(i, 6).Interior
                        .Pattern = Excel.Constants.xlSolid
                        .PatternColorIndex = Excel.Constants.xlAutomatic
                        .Color = 5296274    ' Green
                    End With
                End If

                If flags.Contains("F") Then
                    With ws.Range(ws.Cells(i, 6), ws.Cells(i, 7)).Interior
                        .Pattern = Excel.Constants.xlSolid
                        .PatternColorIndex = Excel.Constants.xlAutomatic
                        .Color = 15773696   ' Blue
                    End With
                End If

                If flags.Contains("R") Then
                    With ws.Range(ws.Cells(i, 6), ws.Cells(i, 7)).Interior
                        .Pattern = Excel.Constants.xlSolid
                        .PatternColorIndex = Excel.Constants.xlAutomatic
                        .Color = 255        ' Red
                    End With
                End If
            End If

            If description.Contains("ReTransmit") Then
                With ws.Range(ws.Cells(i, 6), ws.Cells(i, 7)).Interior
                    .Pattern = Excel.Constants.xlSolid
                    .PatternColorIndex = Excel.Constants.xlAutomatic
                    .Color = 49407          ' Orange
                End With
            End If
        Next i
    End Sub

    Private Sub HilightTime()
        Dim ws As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet
        ' Find the end of the trace
        Dim lastIProw As Integer = ws.Range("A5").SpecialCells(Excel.Constants.xlLastCell).Row
        'Dim currentValue As String = ""
        Dim currentTime As Double
        Dim previousTime As Double
        'Dim timeDeltaD As Double
        'Dim timeDelta as Date
        Dim currentTimeAsDate As Date
        Dim previousTimeAsDate As Date
        Dim timeDelta As TimeSpan
        Dim timeDeltaSeconds As Integer
        Dim hilightColor As Integer

        ' TODO: Get rid of magic numbers
        ' Walk backwards so index works as we add rows
        For i As Integer = lastIProw To Common.cER_FirstDataRow + 1 Step -1
            'Debug.Print(String.Format(">{0}< >{1}<", ws.Cells(i, Common.cER_TimeOfDay_Column).Value, ws.Cells(i - 1, Common.cER_TimeOfDay_Column).Value))
            ' Does not work
            'currentTimeAsDate = DateAndTime.TimeValue(ws.Cells(i, Common.cER_TimeOfDay_Column).Value)
            'previousTimeAsDate = DateAndTime.TimeValue(ws.Cells(i - 1, Common.cER_TimeOfDay_Column).Value)
            ' Does not work
            'currentTimeAsDate = Date.Parse(ws.Cells(i, Common.cER_TimeOfDay_Column).Value)
            'previousTimeAsDate = Date.Parse(ws.Cells(i - 1, Common.cER_TimeOfDay_Column).Value)
            ' Does not work
            'currentTimeAsDate = ws.Cells(i, Common.cER_TimeOfDay_Column).Value
            'previousTimeAsDate = ws.Cells(i - 1, Common.cER_TimeOfDay_Column).Value
            ' 
            currentTimeAsDate = DateTime.FromOADate(ws.Cells(i, Common.cER_TimeOfDay_Column).Value)
            previousTimeAsDate = DateTime.FromOADate(ws.Cells(i - 1, Common.cER_TimeOfDay_Column).Value)

            currentTime = ws.Cells(i, Common.cER_TimeOfDay_Column).Value
            previousTime = ws.Cells(i - 1, Common.cER_TimeOfDay_Column).Value

            'timeDelta = TimeSpan.FromSeconds(currentTime - previousTime)


            'timeDelta = currentTime - previousTime

            timeDelta = currentTimeAsDate.Subtract(previousTimeAsDate)

            timeDeltaSeconds = (timeDelta.Minutes * 60) + timeDelta.Seconds

            If timeDeltaSeconds > 1 Then
                If timeDeltaSeconds > 60 Then
                    hilightColor = 192     ' Dark
                ElseIf timeDeltaSeconds > 5 Then
                    hilightColor = 255     ' Red
                ElseIf timeDeltaSeconds > 2 Then
                    hilightColor = 49407   ' Orange
                Else
                    hilightColor = 65535   ' Yellow
                End If

                With ws.Cells(i, 2).Interior
                    .Pattern = Excel.Constants.xlSolid
                    .PatternColorIndex = Excel.Constants.xlAutomatic
                    .Color = hilightColor
                End With

                ws.Rows(i).Insert(Shift:=Excel.XlDirection.xlDown, CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove)
            End If
        Next i
    End Sub

    Private Sub RemoveColumns()
        Dim hostCount As Integer = CInt(txtHostCount.Text)

        With Globals.ThisAddIn.Application
            ' Count backwards so deletes don't mess up column numbers
            For i As Integer = hostCount - 1 To 0
                .Columns((i * 7) + 6).Delete()
            Next i
        End With
    End Sub
    ' Remove the "(<Hex#>)" part of the Seq and Ack numbers

    Private Sub RemoveHex()
        Dim ws As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet

        ' TODO: Get rid of magic numbers and ranges

        Dim currentValue As String = ""
        Dim lastIProw As Integer = ws.Range("A5").SpecialCells(Excel.Constants.xlLastCell).Row

        For i As Integer = 5 To lastIProw
            'Debug.Print(String.Format(">{0}< >{1}<", ws.Cells(i, Common.cER_TCPSeqNumber_Column).Value, ws.Cells(i, Common.cER_TCPAckNumber_Column).Value))
            currentValue = ws.Cells(i, Common.cER_TCPSeqNumber_Column).Value

            If Not currentValue Is Nothing Then
                ws.Cells(i, Common.cER_TCPSeqNumber_Column).Value = RegularExpressions.Regex.Replace(currentValue, "\(.*\)", "")
            End If

            currentValue = ws.Cells(i, Common.cER_TCPAckNumber_Column).Value

            If Not currentValue Is Nothing Then
                ws.Cells(i, Common.cER_TCPAckNumber_Column).Value = RegularExpressions.Regex.Replace(currentValue, "\(.*\)", "")
            End If
        Next i
    End Sub

#End Region


End Class
