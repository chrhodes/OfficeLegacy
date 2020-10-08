Imports System.Text.RegularExpressions
Imports PacificLife.Life

Public Class ProcessFile
    '**********************************************************************
    '
    ' ProcessFile.vb
    '
    ' This class is responsible for opening the file and processing the
    ' columns according to the settings in the configuration file.
    '
    ' The processing is done on the way in for DataType and after the fact
    ' for formatting.
    '
    ' The configuration file must list all the columns in the source file
    ' and they must be in the same order as the source file as the 
    ' QueryTable.Add() method requires the .TextFileColumnDataTypes to be
    ' passed an array of column types.  The columns must match the order
    ' of the input file.
    '
    ' Changes to this file are likely.  
    '
    '**********************************************************************

    Dim isProcessing As Boolean = False

    Public Function ShouldProcessFile(ByVal fileConfigInfo As Data.DataSet, ByVal fileName As String, ByRef fileNumber As Integer) As Boolean
        PLLog.Trace3("Enter", Globals.cPLLOG_NAME)

        Dim files As Data.DataTable = Config.FileConfigInfo.Tables("File")

        fileNumber = 0

        For Each dr As Data.DataRow In files.Rows

            Try
                Dim fileNameExpression As String = dr.Item("NamePattern")
                Dim rx As Regex = New Regex(fileNameExpression)

                If rx.Match(fileName).Success Then
                    PLLog.Trace3("Exit:True", Globals.cPLLOG_NAME)
                    Return True
                End If
            Catch ex As Exception
                MessageBox.Show(String.Format("Cannot find ""NamePattern"" attribute in <File> element.  Check {0}", Globals.cCONFIG_FILE_NAME))
                MessageBox.Show(ex.ToString)
            End Try

            ' NOTE: fileNumber is returned!  See ByRef above.
            fileNumber += 1
        Next

        PLLog.Trace3("Exit:False", Globals.cPLLOG_NAME)
        Return False

    End Function

    ' TODO: Need to know which file matched in FileConfig as different section could have different numbers of columns.

    Public Sub ProcessFile(ByVal fileConfigInfo As Data.DataSet, ByVal fileName As String, ByRef ws As Excel.Worksheet, ByVal fileNumber As Integer)
        PLLog.Trace1("Enter", Globals.cPLLOG_NAME)

        Dim columns As Data.DataTable = Config.FileConfigInfo.Tables("Column")
        Dim files As Data.DataTable = Config.FileConfigInfo.Tables("File")
        Dim dataStartRow As Integer
        Dim dataOutputRow As Integer
        Dim dataOutputColumn As Integer = 1 ' Excel columns start at 1
        Dim debugOutput As Boolean = False

        ' Get the starting row of the data and the starting row of the output from the <File> element.

        For Each fileRow As Data.DataRow In files.Rows
            If fileRow.Item("File_Id") <> fileNumber Then
                Continue For
            Else
                dataStartRow = fileRow.Item("DataStartRow")
                dataOutputRow = fileRow.Item("DataOutputRow")
                debugOutput = CBool(fileRow.Item("Debug"))
            End If
        Next

        If debugOutput Then
            dataOutputRow += 10 ' Make room for debugging output
            dataOutputColumn += 1
        End If

        ' TODO: See if hard coded 128 is ok or if need to support something larger.
        'Dim fieldTypes() As Integer = {1, 1, 2, 2, 2, 2, 2, 2, 2, 2, 2}
        Dim fieldTypes(128) As Integer
        Dim fieldType As String

        ' Build the field type array based on the types listed in the config file.
        ' Note, the columns must be listed in the same order as the input file.
        Dim columnNumber As Integer = 0

        ' TODO: Must be a better way of doing this.  Perhaps a query?

        For Each dataRow As Data.DataRow In columns.Rows
            If dataRow.Item("Columns_Id") <> fileNumber Then
                Continue For
            End If

            fieldType = dataRow.Item("DataType")

            Debug.Print(fieldType)

            Select Case fieldType
                Case "Text"
                    fieldTypes(columnNumber) = Excel.XlColumnDataType.xlTextFormat

                Case "DMYDate"
                    fieldTypes(columnNumber) = Excel.XlColumnDataType.xlDMYFormat

                Case "DYMDate"
                    fieldTypes(columnNumber) = Excel.XlColumnDataType.xlDYMFormat

                Case "EMDDate"
                    fieldTypes(columnNumber) = Excel.XlColumnDataType.xlEMDFormat

                Case "MDYDate"
                    fieldTypes(columnNumber) = Excel.XlColumnDataType.xlMDYFormat

                Case "MYDDate"
                    fieldTypes(columnNumber) = Excel.XlColumnDataType.xlMYDFormat

                Case "YDMDate"
                    fieldTypes(columnNumber) = Excel.XlColumnDataType.xlYDMFormat

                Case "YMDDate"
                    fieldTypes(columnNumber) = Excel.XlColumnDataType.xlYMDFormat

                Case "General"
                    fieldTypes(columnNumber) = Excel.XlColumnDataType.xlGeneralFormat

                Case "Skip"
                    fieldTypes(columnNumber) = Excel.XlColumnDataType.xlSkipColumn

                Case Else
                    PLLog.Error("Unrecognized column type: " & fieldType, Globals.cPLLOG_NAME)
                    Throw New Exception("Unrecognized column type: " & fieldType)

            End Select

            columnNumber += 1
        Next

        ' Import the file that was opened using the field types.

        PLLog.Trace1("Before QueryTables.Add", Globals.cPLLOG_NAME)

        With ws.QueryTables.Add(Connection:=("TEXT;" & fileName), Destination:=ws.Cells(dataOutputRow, dataOutputColumn))
            .Name = "DNR"
            .FieldNames = True
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .RefreshStyle = Excel.XlCellInsertionMode.xlInsertDeleteCells
            .SavePassword = False
            .SaveData = False
            .AdjustColumnWidth = True
            .RefreshPeriod = 0
            .TextFilePromptOnRefresh = False
            .TextFilePlatform = 437
            .TextFileStartRow = dataStartRow
            .TextFileParseType = Excel.XlTextParsingType.xlDelimited
            .TextFileTextQualifier = Excel.XlTextQualifier.xlTextQualifierDoubleQuote
            .TextFileConsecutiveDelimiter = False
            .TextFileCommaDelimiter = True
            .TextFileSpaceDelimiter = False
            .TextFileTabDelimiter = False
            .TextFileColumnDataTypes = fieldTypes
            .TextFileTrailingMinusNumbers = True
            .Refresh(BackgroundQuery:=False)
        End With

        ' For reasons that aren't clear, Excel is unhappy saving the workbook in 2007 OpenXML format
        ' with the QueryTable we just created inside.  Since we don't really need it, just remove
        ' it now that the data has been imported.

        ws.QueryTables.Item("DNR").Delete()

        PLLog.Trace1("After QueryTables.Item.Delete", Globals.cPLLOG_NAME)
        ' Now make any other formatting changes specified.

        columnNumber = dataOutputColumn
        Dim column As Excel.Range

        If debugOutput Then
            ' Display some diagnostic info
            ws.Cells(1, 1).Value = "Name"
            ws.Cells(2, 1).Value = "DataType"
            ws.Cells(3, 1).Value = "Width"
            ws.Cells(4, 1).Value = "Wrap"
            ws.Cells(5, 1).Value = "FontSize"
            ws.Cells(6, 1).Value = "HeaderFontSize"
            ws.Cells(7, 1).Value = "HeaderBold"
            ws.Cells(8, 1).Value = "HeaderWrap"
        End If

        For Each dataRow As Data.DataRow In columns.Rows
            If dataRow.Item("Columns_Id") <> fileNumber Then
                Continue For
            End If

            column = ws.Columns(columnNumber)

            With column
                .ColumnWidth = dataRow.Item("Width")
                .WrapText = CBool(dataRow.Item("Wrap"))
                .Font.Size = dataRow.Item("FontSize")
                '.Font.Bold = CBool(dataRow.Item("Bold"))
            End With

            With ws.Cells(dataOutputRow, columnNumber)
                .WrapText = CBool(dataRow.Item("HeaderWrap"))
                .Font.Size = dataRow.Item("HeaderFontSize")
                .Font.Bold = CBool(dataRow.Item("HeaderBold"))
            End With

            If debugOutput Then
                ws.Cells(1, columnNumber).Value = dataRow.Item("Name")
                ws.Cells(2, columnNumber).Value = dataRow.Item("DataType")
                ws.Cells(3, columnNumber).Value = dataRow.Item("Width")
                ws.Cells(4, columnNumber).Value = dataRow.Item("Wrap")
                ws.Cells(5, columnNumber).Value = dataRow.Item("FontSize")
                ws.Cells(6, columnNumber).Value = dataRow.Item("HeaderFontSize")
                ws.Cells(7, columnNumber).Value = dataRow.Item("HeaderBold")
                ws.Cells(8, columnNumber).Value = dataRow.Item("HeaderWrap")
            End If

            columnNumber += 1
        Next

        PLLog.Trace1("Exit", Globals.cPLLOG_NAME)
    End Sub
End Class
