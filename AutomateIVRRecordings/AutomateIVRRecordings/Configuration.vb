Imports System.IO

Public Class Configuration

    Private Sub Sheet2_Startup(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Startup

    End Sub

    Private Sub Sheet2_Shutdown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shutdown

    End Sub

    Private Sub bProduceOutputWorksheet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bProduceOutputWorksheet.Click
        ' Grab all the stuff we need off the configuration sheet
        Dim TT10FilePrefix As String = Range(Globals.cCFG_TT10_FilePrefix_Cell).Value
        Dim TT10MessagePrefix As String = Range(Globals.cCFG_TT10_MessagePrefix_Cell).Value
        Dim TT10MessageSuffix As String = Range(Globals.cCFG_TT10_MessageSuffix_Cell).Value

        Dim TT11FilePrefix As String = Range(Globals.cCFG_TT11_FilePrefix_Cell).Value
        Dim TT11MessagePrefix As String = Range(Globals.cCFG_TT11_MessagePrefix_Cell).Value
        Dim TT11MessageSuffix As String = Range(Globals.cCFG_TT11_MessageSuffix_Cell).Value

        Dim TT12FilePrefix As String = Range(Globals.cCFG_TT12_FilePrefix_Cell).Value
        Dim TT12MessagePrefix As String = Range(Globals.cCFG_TT12_MessagePrefix_Cell).Value
        Dim TT12MessageSuffix As String = Range(Globals.cCFG_TT12_MessageSuffix_Cell).Value

        Dim TT13FilePrefix As String = Range(Globals.cCFG_TT13_FilePrefix_Cell).Value
        Dim TT13MessagePrefix As String = Range(Globals.cCFG_TT13_MessagePrefix_Cell).Value
        Dim TT13MessageSuffix As String = Range(Globals.cCFG_TT13_MessageSuffix_Cell).Value

        Dim TT14FilePrefix As String = Range(Globals.cCFG_TT14_FilePrefix_Cell).Value
        Dim TT14MessagePrefix As String = Range(Globals.cCFG_TT14_MessagePrefix_Cell).Value
        Dim TT14MessageSuffix As String = Range(Globals.cCFG_TT14_MessageSuffix_Cell).Value

        Dim inputWorksheetName As String = Range(Globals.cCFG_InputWorksheetName_Cell).Value
        Dim inputStartingRow As Integer = Range(Globals.cCFG_InputStartingRow_Cell).Value
        Dim inputEndingRow As Integer = Range(Globals.cCFG_InputEndingRow_Cell).Value

        Dim outputWorksheetName As String = Range(Globals.cCFG_Input_OutputWorksheetName_Cell).Value
        Dim outputStartingRow As Integer = Range(Globals.cCFG_Input_OutputStartingRow_Cell).Value
        Dim outputEndingRow As Integer = outputStartingRow

        Dim inputWorksheet As Excel.Worksheet
        Dim outputWorksheet As Excel.Worksheet
        Dim configurationWorksheet As Excel.Worksheet = Application.ActiveSheet

        inputWorksheet = Application.Worksheets(inputWorksheetName)
        outputWorksheet = Util.ExcelHelper.NewWorksheet(outputWorksheetName, , "Configuration")

        Dim fileName As String
        Dim message As String
        Dim fundCode As String
        Dim fundName As String

        For currentInputRow = inputStartingRow To inputEndingRow
            fundCode = inputWorksheet.Cells(currentInputRow, 1).Value
            fundName = inputWorksheet.Cells(currentInputRow, 2).Value
            'System.Diagnostics.Trace.WriteLine(String.Format("{0} {1}", fundCode, fundName))

            ' Note: AddRowToOutput increments outputEndingRow

            fileName = String.Format("{0}-{1}{2}", TT10FilePrefix, fundCode, ".txt")
            message = String.Format("{0}{1}{2}", TT10MessagePrefix, fundName, TT10MessageSuffix)
            AddRowToOutputWorksheet(outputWorksheetName, outputEndingRow, fileName, message)

            fileName = String.Format("{0}-{1}{2}", TT11FilePrefix, fundCode, ".txt")
            message = String.Format("{0}{1}{2}", TT11MessagePrefix, fundName, TT11MessageSuffix)
            AddRowToOutputWorksheet(outputWorksheetName, outputEndingRow, fileName, message)

            fileName = String.Format("{0}-{1}{2}", TT12FilePrefix, fundCode, ".txt")
            message = String.Format("{0}{1}{2}", TT12MessagePrefix, fundName, TT12MessageSuffix)
            AddRowToOutputWorksheet(outputWorksheetName, outputEndingRow, fileName, message)

            fileName = String.Format("{0}-{1}{2}", TT13FilePrefix, fundCode, ".txt")
            message = String.Format("{0}{1}{2}", TT13MessagePrefix, fundName, TT13MessageSuffix)
            AddRowToOutputWorksheet(outputWorksheetName, outputEndingRow, fileName, message)

            fileName = String.Format("{0}-{1}{2}", TT14FilePrefix, fundCode, ".txt")
            message = String.Format("{0}{1}{2}", TT14MessagePrefix, fundName, TT14MessageSuffix)
            AddRowToOutputWorksheet(outputWorksheetName, outputEndingRow, fileName, message)
        Next

        outputWorksheet.Columns.AutoFit()

        ' Return focus to the Configuration sheet
        ' propogate the information down to the Output section
        ' and activate the next likely input cell.

        configurationWorksheet.Activate()

        Range(Globals.cCFG_OutputStartingRow_Cell).Value = outputStartingRow
        Range(Globals.cCFG_OutputEndingRow_Cell).Value = outputEndingRow - 1
        Range(Globals.cCFG_Output_OutputWorksheetName_Cell).Value = outputWorksheetName

        Range(Globals.cCFG_Output_OutputWorksheetName_Cell).Activate()
    End Sub

    Sub AddRowToOutputWorksheet(ByVal outputSheetName As String, ByRef outputRow As Integer, ByVal fileName As String, ByVal message As String)
        Application.Worksheets(outputSheetName).Cells(outputRow, 1).Value = fileName
        Application.Worksheets(outputSheetName).Cells(outputRow, 2).Value = message
        outputRow += 1
    End Sub

    Private Sub bGenerateOutputFiles_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bGenerateOutputFiles.Click
        Dim outputFolderPath As String = Range(Globals.cCFG_OutputFolderPath_Cell).Value
        Dim startingRow As Integer = Range(Globals.cCFG_OutputStartingRow_Cell).Value
        Dim endingRow As Integer = Range(Globals.cCFG_OutputEndingRow_Cell).Value
        Dim outputWorksheetName As String = Range(Globals.cCFG_Output_OutputWorksheetName_Cell).Value
        Dim commandFileName As String = Range(Globals.cCFG_CommandFileName_Cell).Value

        Dim outputWorksheet As Excel.Worksheet = Application.Worksheets(outputWorksheetName)

        Dim fileOutput As New FileOutput
        Dim outputFileName As String
        Dim fileContents As String
        Dim commandStreamWriter As StreamWriter

        If Util.File.IsValidOutputFolder(outputFolderPath, True) Then
            ' Build the output command file.  This will be useful when the
            ' file processing utility can take command line input.

            commandStreamWriter = fileOutput.CreateCommandFile(String.Format("{0}\{1}", outputFolderPath, commandFileName))

            For currentRow = startingRow To endingRow
                outputFileName = outputWorksheet.Cells(currentRow, 1).Value
                fileContents = outputWorksheet.Cells(currentRow, 2).Value
                'System.Diagnostics.Trace.WriteLine(String.Format("{0} {1}", outputFileName, fileContents))

                ' Produce individual output files

                If IsValidOutputFileName(outputFileName) And IsValidFileContents(fileContents) Then
                    fileOutput.CreateOutputFile(String.Format("{0}\{1}", outputFolderPath, outputFileName), fileContents)
                Else
                    MsgBox(String.Format("Skipping output row {0}", currentRow))
                End If

                ' Add output to command file.

                fileOutput.AddCommandLine(commandStreamWriter, outputFileName)
            Next

            fileOutput.CloseCommandFile(commandStreamWriter)
        Else
            ' Folder didn't exist and user did not want to create it.
            Return
        End If
    End Sub


    Private Function IsValidFileContents(ByVal fileContents As String) As Boolean
        If fileContents.Length > 0 Then
            Return True
        Else
            MsgBox("Missing file contents")
            Return False
        End If
    End Function

    Private Function IsValidOutputFileName(ByVal outputFileName As String) As Boolean
        If outputFileName.Length > 0 Then
            Return True
        Else
            MsgBox("Missing output filename")
            Return False
        End If
    End Function
End Class
