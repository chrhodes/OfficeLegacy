Option Strict Off

Imports System.IO
Imports System.Reflection
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Windows.Forms

Imports AddinHelper
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop
Imports PacificLife.Life

' TODO: This could probably do with some cleanup.  Had to make everything shared.

Public Class Excel_FolderMaps
    Inherits AddinHelper.AppMethod

#Region "General Addin Constants"
    Private Const _MODULE_NAME As String = Common.PROJECT_NAME & "Excel_FolderMaps"
    Private Const _NAME As String = "Excel_FolderMaps"
    Private Const _BITMAP_NAME As String = "folder map.bmp"
    Private Const _CAPTION As String = "Folder Map"
    Private Const _TOOL_TIP_TEXT As String = "Folder Map"
    Private Const _DESCRIPTION As String = "FolderMaps does ..."
#End Region

#Region "Private Constants and Ennumerations"

    Private Const _INDENT_LEVEL As Short = 1
    Private Const _COL_WIDTH As Short = 3
    Private Const _NOTE_WIDTH As Short = 20
    Private Const _FILE_FONT_SIZE As Short = 6
    Private Const _FOLDER_FONT_SIZE As Short = 8

    Private Const _HEADING_ROW As Integer = 2
    Private Const _INITIAL_ROW As Integer = _HEADING_ROW + 1

    Private Const _ERROR_EMPTY_CELL As String = "Cell is empty.  Must select a popluated starting cell first."

    Private Const _FOLDER_INFO_COL As Short = 1 ' Folder level info starts here
    Private Const _FOLDER_INFO_LEN As Short = 10
    Private Const _FILE_INFO_COL As Short = 5   ' File Info starts here
    Private Const _FILE_INFO_LEN As Short = 4
    Private Const _NOTE_COL As Short = 10
    Private Const _INITIAL_COL As Short = 11    ' Map Info starts here

    Private Const _MAKE_BOLD As Boolean = True

    Public Enum _DateType As Integer
        LastCreate = 1
        LastWrite = 2
        LastAccess = 3
    End Enum
#End Region

#Region "Private Types and Variables"

    Private Shared _Row As Integer
    Private Shared _Column As Short

    Private Shared _StartingFolder As String

    Private Shared _LimitLevels As Boolean
    Private Shared _LimitLevel As Short
    Private Shared _GroupResults As Boolean
    Private Shared _GroupLevel As Short
    Private Shared _GroupStartRow As Short
    Private Shared _PatternMatchFolderHighlight As Boolean
    Private Shared _FolderMatchPattern As String
    Private Shared _PatternMatchFileOutput As Boolean
    Private Shared _FileMatchPattern As String
    Private Shared _ShowFiles As Boolean
    Private Shared _ShowFolders As Boolean
    Private Shared _SkipFoldersWithNoFiles As Boolean
    Private Shared _TotalFolders As Integer
    Private Shared _TotalFiles As Integer
    Private Shared _TotalSize As Long
    Private Shared _MonthsSinceCreated As Integer
    Private Shared _MonthsSinceWritten As Integer
    Private Shared _MonthsSinceAccessed As Integer
    Private Shared _MonthsCreatedColor As Integer
    Private Shared _MonthsWrittenColor As Integer
    Private Shared _MonthsAccessedColor As Integer
    Private Shared _MonthsDefaultColor As Integer
    Private Shared _ColorCodeDates As Boolean
    Private Shared _FolderMatchColor As Integer
    Private Shared _PatternMatchFileColor As Integer
    Private Shared _NoAccessColor As Integer
    Private Shared _PathTooLongColor As Integer
    Private Shared _CheckIllegalCharacters As Boolean
    Private Shared _IllegalCharactersColor As Integer
    Private Shared _IllegalFileCharacters As String
    Private Shared _IllegalFolderCharacters As String

    Private Shared _CheckSharePointFileNameLength As Boolean
    Private Shared _IllegalFileNameLengthColor As Integer
    Private Shared _MaxFileNameLength As Integer

    Private _FolderMapsMethod As AppMethod

#End Region

#Region "Public Methods"

    Public Sub New(ByRef commandBar As CommandBar, ByRef buttonStyle As MsoButtonStyle)
        MyBase.Name = _NAME
        MyBase.CommandBar = commandBar
        MyBase.EventHandler = AddressOf Action
        MyBase.ButtonStyle = buttonStyle
        MyBase.BitMapName = _BITMAP_NAME
        MyBase.Asmbly = [Assembly].GetExecutingAssembly
        MyBase.Caption = _CAPTION
        MyBase.ToolTipText = _TOOL_TIP_TEXT
        MyBase.Description = _DESCRIPTION

        MyBase.Initialize()
        'AddLocalButtons()
    End Sub


    '$-------------------------------------------------------------------
    '
    ' $Name        :CreateFolderMap()$
    ' $Type        :Public Sub()$
    '
    ' $Arguments   :$
    '
    ' $Alters      :$
    '
    ' $Description :Main routine.  Called from event handler.$
    '
    ' $ToDo        :$
    '
    '$$------------------------------------------------------------------

    Public Shared Sub CreateFolderMap()
        Using frmF As New frmExcel_FolderMaps()
            Try
                If System.Windows.Forms.DialogResult.Cancel = frmF.ShowDialog() Then
                    Return
                End If

                Globals.ThisAddIn.Application.ScreenUpdating = False

                ' Grab the information we need from the form.  I guess we could
                ' unload the form and ditch it, but that would lose any state.
                InitializeLocalsFromFormFields(frmF)

                If _StartingFolder.Length > 0 Then
                    PopulateFolderMap(_StartingFolder)
                    SaveFolderMap(_StartingFolder)
                Else
                    MessageBox.Show("Must select starting folder")
                End If

            Catch ex As Exception
                ' TODO: EntLib
                MessageBox.Show(ex.ToString)
                Throw (ex)
            Finally
                Globals.ThisAddIn.Application.ScreenUpdating = True
            End Try
        End Using
    End Sub

    Private Shared Sub InitializeLocalsFromFormFields(ByVal frmF As frmExcel_FolderMaps)
        With frmF
            _StartingFolder = .txtStartingFolder.Text
            _LimitLevels = .chkLimitLevels.CheckState
            _LimitLevel = CShort(.txtLimitLevel.Text) + _INITIAL_COL    ' Map data starts at _INITIAL_COL
            _GroupResults = .chkGroupResults.CheckState
            _GroupLevel = CShort(.txtGroupLevel.Text)
            _PatternMatchFolderHighlight = .chkPatternMatchFolderHighlight.CheckState
            _FolderMatchPattern = .txtFolderMatchPattern.Text
            _PatternMatchFileOutput = .chkPatternMatchFileOutput.CheckState
            _FileMatchPattern = .txtFileMatchPattern.Text
            _ShowFiles = .chkShowFiles.CheckState
            _ShowFolders = .chkShowFolders.CheckState
            _SkipFoldersWithNoFiles = .chkSkipFoldersWithNoFiles.CheckState

            '_PathTooLongColor = .pnlPathTooLongColor.BackColor.ToArgb

            _FolderMatchColor = RGB(.pnlFolderHighlightColor.BackColor.R, .pnlFolderHighlightColor.BackColor.G, .pnlFolderHighlightColor.BackColor.B)
            _PathTooLongColor = RGB(.pnlPathTooLongColor.BackColor.R, .pnlPathTooLongColor.BackColor.G, .pnlPathTooLongColor.BackColor.B)
            _NoAccessColor = RGB(.pnlNoAccessColor.BackColor.R, .pnlNoAccessColor.BackColor.G, .pnlNoAccessColor.BackColor.B)

            _PatternMatchFileColor = RGB(.pnlPatternMatchFileColor.BackColor.R, .pnlPatternMatchFileColor.BackColor.G, .pnlPatternMatchFileColor.BackColor.B)

            _ColorCodeDates = .chkColorCodeDates.CheckState

            _MonthsDefaultColor = RGB(.pnlDefaultColor.BackColor.R, .pnlDefaultColor.BackColor.G, .pnlDefaultColor.BackColor.B)

            _MonthsCreatedColor = RGB(.pnlMonthCreatedColor.BackColor.R, .pnlMonthCreatedColor.BackColor.G, .pnlMonthCreatedColor.BackColor.B)
            _MonthsWrittenColor = RGB(.pnlMonthWrittenColor.BackColor.R, .pnlMonthWrittenColor.BackColor.G, .pnlMonthWrittenColor.BackColor.B)
            _MonthsAccessedColor = RGB(.pnlMonthAccessedColor.BackColor.R, .pnlMonthAccessedColor.BackColor.G, .pnlMonthAccessedColor.BackColor.B)

            _MonthsSinceCreated = .txtMonthsSinceCreated.Text
            _MonthsSinceWritten = .txtMonthsSinceWritten.Text
            _MonthsSinceAccessed = .txtMonthsSinceAccessed.Text

            '_PathTooLongColor = .pnlPathTooLongColor.BackColor.ToArgb
            _PathTooLongColor = RGB(.pnlPathTooLongColor.BackColor.R, .pnlPathTooLongColor.BackColor.G, .pnlPathTooLongColor.BackColor.B)

            _CheckIllegalCharacters = .chkIllegalCharacters.CheckState
            _IllegalCharactersColor = .pnlIllegalCharactersColor.BackColor.ToArgb
            _IllegalFileCharacters = .txtIllegalFileCharacters.Text
            _IllegalFolderCharacters = .txtIllegalFolderCharacters.Text

            _CheckSharePointFileNameLength = .chkFileNameLength.CheckState
            '_IllegalFileNameLengthColor = .pnlIllegalFileNameLengthColor.BackColor.ToArgb
            _IllegalFileNameLengthColor = RGB(.pnlIllegalFileNameLengthColor.BackColor.R, .pnlIllegalFileNameLengthColor.BackColor.G, .pnlIllegalFileNameLengthColor.BackColor.B)
            _MaxFileNameLength = CInt(.txtMaxFileNameLength.Text)

        End With
    End Sub

    '$-------------------------------------------------------------------
    '
    ' $Name        :PopulateFolderMap$
    ' $Type        :Public Sub()$
    '
    ' $Arguments   :$
    '
    ' $Alters      :$
    '
    ' $Description :Open the folder to map and descend.  Called from form.$
    '
    ' $ToDo        :$
    '
    '$$------------------------------------------------------------------

    Public Shared Sub PopulateFolderMap(ByRef startingFolder As String)
        Try
            With Globals.ThisAddIn.Application

                ' Set starting point for folder output

                _Row = _INITIAL_ROW
                _Column = _INITIAL_COL
                _TotalFolders = 0
                _TotalFiles = 0
                _TotalSize = 0

                ' Always add a new sheet so can accumulate results in Workbook.
                ' Need to handle case when no Workbook exists.  Tried a variety of 
                ' ways to determine empty workbook.  PERSONAL.XLS may be open if 
                ' macros have been created so cannot rely on Workbooks.count.
                ' Worksheets.Count throws an exception.
                '
                'Dim wb as Microsoft.Office.Interop.Excel.Workbook
                '
                'For Each wb In .Workbooks
                '    Debug.Print(wb.Name)
                'Next

                'Dim ws As Microsoft.Office.Interop.Excel.Worksheet

                'For Each ws In .Worksheets
                '    Debug.Print(ws.Name)
                'Next

                ' ActiveWorkbook seems to work reliably.  Maybe Interop issue, who knows.

                If .ActiveWorkbook Is Nothing Then
                    .Workbooks.Add() ' Get a new WorkSheet (or more :)) for free.
                Else
                    ' TODO: Prompt to use existing sheet if found.
                    .ActiveWorkbook.Worksheets.Add()
                End If

                .Cells(_HEADING_ROW, _FOLDER_INFO_COL).Value = "Cummulative Folder Count"
                .Cells(_HEADING_ROW, _FOLDER_INFO_COL + 1).Value = "Cummulative File Count"
                .Cells(_HEADING_ROW, _FOLDER_INFO_COL + 2).Value = "Cummulative Size"

                .Cells(_HEADING_ROW, _FILE_INFO_COL).Value = "Count"
                .Cells(_HEADING_ROW, _FILE_INFO_COL + 1).Value = "Size"
                .Cells(_HEADING_ROW, _FILE_INFO_COL + 2).Value = "Last Create"
                .Cells(_HEADING_ROW, _FILE_INFO_COL + 3).Value = "Last Write"
                .Cells(_HEADING_ROW, _FILE_INFO_COL + 4).Value = "Last Access"

                .Cells(_HEADING_ROW, _INITIAL_COL).Value = startingFolder

                Dim startingRow As Integer = _Row
                Dim startingColumn As Integer = _Column

                Dim numberFoldersLocal As Integer = 0
                Dim numberFilesLocal As Integer = 0
                Dim sizeFilesLocal As Long = 0

                Dim dirInfo As New FileInfo(startingFolder)

                Dim maxLastCreate As Date = Date.MinValue
                Dim maxLastWrite As Date = Date.MinValue
                Dim maxLastAccess As Date = Date.MinValue

                Dim column As Short = 1
                Dim fontSize As Short = 10

                ' Note: maxLastDate is passed in, even though we don't use the updated value.
                ' it is used during the recursion process when ListFolders calls itself.  Do
                ' not be mislead and rework this logic!

                ListFolders( _
                    startingFolder, _Column, fontSize, _
                    _TotalFolders, _TotalFiles, _TotalSize, _
                    maxLastCreate, maxLastWrite, maxLastAccess, _
                    _ShowFiles, _ShowFolders)

                Dim orientation As Microsoft.Office.Interop.Excel.XlPageOrientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlPortrait

                FormatFolderMapSheet(startingFolder, orientation)
            End With
        Catch ex As Exception
            MsgBox(ex.Message)
            Throw ex
        End Try
    End Sub ' FolderMap_Start

    '$-------------------------------------------------------------------
    '
    ' $Name        :SaveFolderMap()$
    ' $Type        :Public Sub()$
    '
    ' $Arguments   :$
    '
    ' $Alters      :$
    '
    ' $Description :$
    '
    ' $ToDo        :$
    '
    '$$------------------------------------------------------------------

    Public Shared Sub SaveFolderMap(ByVal startingFolder As String)
        Using saveFileDialog As New SaveFileDialog()
            Dim strOutputFile As String
            Try
                saveFileDialog.FileName = "ExcelFolderMaps.xls"
                saveFileDialog.InitialDirectory = startingFolder
                If System.Windows.Forms.DialogResult.Cancel = saveFileDialog.ShowDialog() Then
                    Return
                Else
                    strOutputFile = saveFileDialog.FileName
                End If
                If "" = strOutputFile Then
                    Return
                End If
                Globals.ThisAddIn.Application.ActiveWorkbook.SaveAs(strOutputFile)
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Using
    End Sub

    Shared Function GetEndOfSectionDown( _
    ByVal intStartRow As Integer, _
    ByVal intStartCol As Integer, _
    ByVal intLastRow As Integer _
) As Integer
        Dim intMatchingRow As Integer

        With Globals.ThisAddIn.Application
            ' Search down for a matching cell
            intMatchingRow = .Cells(intStartRow, intStartCol).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).Row

            If intStartCol = _INITIAL_COL Then
                ' We have back'd all the way back to the first column.
                ' Return either then next matching cell down or the last
                ' populated row on the sheet.

                If intMatchingRow < intLastRow Then
                    ' Section ends on the row prior to the match.
                    GetEndOfSectionDown = intMatchingRow - 1
                Else
                    ' Return end of populated section
                    GetEndOfSectionDown = intLastRow
                End If
            Else
                If intMatchingRow <= intLastRow Then
                    ' Back up one column and search down for a populated cell.
                    ' Treat row prior to matching row as new end.
                    GetEndOfSectionDown = GetEndOfSectionDown(intStartRow, intStartCol - 1, intMatchingRow - 1)
                Else
                    ' Back up one column and search down for a populated cell.
                    ' Treat end of worksheet as end.
                    GetEndOfSectionDown = GetEndOfSectionDown(intStartRow, intStartCol - 1, intLastRow)
                End If
            End If
        End With
    End Function

    '**********************************************************************
    '   P r i v a t e    M e t h o d s
    '**********************************************************************

    '$-------------------------------------------------------------------
    '
    ' $Name        :$FolderFiles_List
    ' $Type        :Private Sub()$
    '
    ' $Arguments   :$
    '
    ' $Returns     :Count of rows added.$
    '
    ' $Alters      :$
    '
    ' $Description :
    '   Lists all the files in the specified folder.  Handles details
    '   of what appears on each line.$
    '
    ' $ToDo        :$
    '
    '$$------------------------------------------------------------------

    Private Shared Sub ListFiles( _
        ByRef folder As String, _
        ByRef column As Short, _
        ByRef numberFiles As Integer, _
        ByRef sizeFiles As Long, _
        ByRef maxLastCreateDate As Date, _
        ByRef maxLastWriteDate As Date, _
        ByRef maxLastAccessDate As Date, _
        ByVal showFiles As Boolean _
    )
        Dim file As String
        Dim defaultFileColor As Integer = System.ConsoleColor.Black

        Dim files() As String = Directory.GetFiles(folder)    ' Full path names

        numberFiles = 0
        sizeFiles = 0

        Dim fileAdded As Boolean = False

        For Each file In files
            ' Is there a better way to use FileInfo?  Initialize just one and reuse?
            Try
                fileAdded = False

                Dim fileInfo As FileInfo = New FileInfo(file)

                If _PatternMatchFileOutput Then
                    If Not Regex.Match(fileInfo.Name, _FileMatchPattern).Success Then
                        Continue For
                    Else
                        defaultFileColor = _PatternMatchFileColor
                    End If
                End If

                numberFiles += 1
                sizeFiles += fileInfo.Length

                ' If there is a local file with a more current date, alert our caller.

                If maxLastCreateDate < fileInfo.CreationTime Then
                    maxLastCreateDate = fileInfo.CreationTime
                End If

                If maxLastWriteDate < fileInfo.LastWriteTime Then
                    maxLastWriteDate = fileInfo.LastWriteTime
                End If

                If maxLastAccessDate < fileInfo.LastAccessTime Then
                    maxLastAccessDate = fileInfo.LastAccessTime
                End If

                ' Note, we add file rows if checking for illegal characters or file name length
                ' even if not listing files.

                If _CheckIllegalCharacters Then
                    If HasIllegalFileNameCharacters(fileInfo.Name) Then
                        AddFileRow(fileInfo, column, _FILE_FONT_SIZE, False, _IllegalCharactersColor)
                        fileAdded = True
                    End If
                End If

                ' TODO: This overrides an IllegalCharacters color.  Maybe bold or something else.
                If _CheckSharePointFileNameLength Then
                    If fileInfo.Name.Length > _MaxFileNameLength Then
                        AddFileRow(fileInfo, column, _FILE_FONT_SIZE, False, _IllegalFileNameLengthColor)
                        fileAdded = True
                    End If
                End If

                If showFiles And Not fileAdded Then
                    AddFileRow(fileInfo, column, _FILE_FONT_SIZE, False, defaultFileColor)
                End If
            Catch ex As System.IO.PathTooLongException
                ' TODO: We now know that the current file exists at a path+filename that exceeds
                ' the allowable lengths, however, we don't have the size and date info.
                ' Need to do some extra work to get this.  Not sure it is worth it.  For now
                ' just log that the path is too long.

                Dim baseFileName As String

                If _ShowFolders Then
                    ' There will be folder information to show where the file is located so just
                    ' display the filename.

                    baseFileName = file.Substring(file.LastIndexOf("\") + 1)
                Else
                    ' Need to display the full path to the file since no folders are being displayed.

                    baseFileName = file
                End If

                'System.IO.Directory.SetCurrentDirectory(folder)
                '' Get the base filename skipping the last "\"

                '' This still blows up!
                'Dim fileInfo As FileInfo = New FileInfo(baseFileName)
                'sizeFiles += fileInfo.Length

                Dim errorInfo As String = String.Format("{0} - path({1}) + filename({2}) too long({3})", _
                                                        baseFileName, folder.Length, baseFileName.Length, file.Length)
                AddErrorRow(errorInfo, column, _FILE_FONT_SIZE, False, _PathTooLongColor)
            Catch ex As Exception
                MessageBox.Show(ex.ToString)
                Throw ex
            End Try
        Next file
    End Sub ' ListFiles()

    '$-------------------------------------------------------------------
    '
    ' $Name        :FolderFolders_List$
    ' $Type        :Sub()$
    '
    ' $Arguments   :$
    '
    ' $Returns     :Count of rows added.$
    '
    ' $Alters      :$
    '
    ' $Description :
    '   This routine calls ListFiles() and then calls
    '   itself recursively to descend the folder hierarchy.$
    '
    ' $ToDo        :$
    '
    '$$------------------------------------------------------------------

    Private Shared Sub ListFolders( _
        ByVal startingFolder As String, _
        ByVal column As Short, _
        ByVal fontSize As Short, _
        ByRef numberFoldersCummulative As Integer, _
        ByRef numberFilesCummulative As Integer, _
        ByRef sizeFilesCummulative As Long, _
        ByRef maxLastCreateDate As Date, _
        ByRef maxLastWriteDate As Date, _
        ByRef maxLastAccessDate As Date, _
        Optional ByRef showFiles As Boolean = False, _
        Optional ByRef showFolders As Boolean = True _
    )
        'Dim objFld As Scripting.Folder
        Dim intColMax As Short
        'Dim innerDir As String

        Dim numberFoldersCummulativeLocal As Integer = 0
        Dim numberFilesCummulativeLocal As Integer = 0
        Dim sizeFilesCummulativeLocal As Long = 0

        Dim numberFoldersLocal As Integer = 0
        Dim numberFilesLocal As Integer = 0
        Dim sizeFilesLocal As Long = 0

        Dim dirInfo As New FileInfo(startingFolder)
        Dim folderAdded As Boolean = False

        ' Start off with Date.MinValue for dates.
        ' The ListFiles methods will further update this with file information.

        Dim localDirCreateDate As Date = Date.MinValue
        Dim localDirLastWriteDate As Date = Date.MinValue
        Dim localDirLastAccessDate As Date = Date.MinValue

        If _CheckIllegalCharacters Then
            If HasIllegalFolderNameCharacters(dirInfo.Name) Then
                AddFolderRow(dirInfo, column, fontSize, _MAKE_BOLD, _IllegalCharactersColor)
                folderAdded = True
            End If
        End If

        If showFolders And Not folderAdded Then
            AddFolderRow(dirInfo, column, fontSize, _MAKE_BOLD)
        End If

        ' Save the current location as we need to come back and add the totals to this row
        ' We already bumped _Row (in AddFolderRow) when we added the Folder we are on
        Dim currentRow As Integer = _Row - 1
        Dim currentColumn As Integer = column

        Try
            ' First list the files in the current folder.  
            ' Note: We call ListFiles even if not showing files (blnShowFiles is False) 
            ' so we can get information about the files to include in the totals.

            ListFiles( _
                startingFolder, _
                column + _INDENT_LEVEL, _
                numberFilesLocal, _
                sizeFilesLocal, _
                localDirCreateDate, _
                localDirLastWriteDate, _
                localDirLastAccessDate, _
                showFiles)

            ' Update the dates with the information from the files that were found.
            ' The dates will not have changed if there were no local files.  If
            ' that is the case use the directory dates.

            If localDirCreateDate > maxLastCreateDate Then
                maxLastCreateDate = localDirCreateDate
            Else
                maxLastCreateDate = dirInfo.CreationTime
            End If

            If localDirLastWriteDate > maxLastWriteDate Then
                maxLastWriteDate = localDirLastWriteDate
            Else
                maxLastWriteDate = dirInfo.LastWriteTime
            End If

            If localDirLastAccessDate > maxLastAccessDate Then
                maxLastAccessDate = localDirLastAccessDate
            Else
                maxLastAccessDate = dirInfo.LastAccessTime
            End If
        Catch ex As Exception
            PLLog.Error(ex, Common.PROJECT_NAME)
            Throw (ex)  ' So we can add code to catch later.
        End Try

        Dim dirs() As String = Directory.GetDirectories(startingFolder)

        ' Then explore each sub folder

        For Each innerDir As String In dirs
            Try
                Dim innerDirInfo As FileInfo = New FileInfo(innerDir)

                numberFoldersLocal += 1

                Dim numberFoldersI As Integer = 0
                Dim numberFilesI As Integer = 0
                Dim sizeFilesI As Long = 0

                ' Stop if limiting the depth of the exploration

                If False = _LimitLevels Or (True = _LimitLevels And _LimitLevel > column) Then
                    ' Call ourselves recursively to display sub folders.

                    localDirCreateDate = Date.MinValue
                    localDirLastWriteDate = Date.MinValue
                    localDirLastAccessDate = Date.MinValue

                    ListFolders( _
                        innerDir, column + _INDENT_LEVEL, _FOLDER_FONT_SIZE, _
                        numberFoldersI, numberFilesI, sizeFilesI, _
                        localDirCreateDate, localDirLastWriteDate, localDirLastAccessDate, _
                        showFiles, showFolders)

                    If localDirCreateDate > maxLastCreateDate Then
                        maxLastCreateDate = localDirCreateDate
                    End If

                    If localDirLastWriteDate > maxLastWriteDate Then
                        maxLastWriteDate = localDirLastWriteDate
                    End If

                    If localDirLastAccessDate > maxLastAccessDate Then
                        maxLastAccessDate = localDirLastAccessDate
                    End If

                    numberFoldersCummulativeLocal += numberFoldersI
                    numberFilesCummulativeLocal += numberFilesI
                    sizeFilesCummulativeLocal += sizeFilesI

                    intColMax = column + _INDENT_LEVEL
                End If
            Catch ex As UnauthorizedAccessException
                ' Add a message to indicate there is no access to the item previously added.
                ' Move over +1 so grouping isn't impacted.
                AddErrorRow("<No Access>", column + _INDENT_LEVEL + 1, _FOLDER_FONT_SIZE, False, _NoAccessColor)
            Catch ex As System.IO.PathTooLongException
                'Dim indexS As String = File.LastIndexOf("\")
                'Dim length As Integer = File.Length
                'Dim errorInfo As String = String.Format("{0} - Path too long ({1})", File.Substring(File.LastIndexOf("\")), File.Length)
                'AddErrorRow(errorInfo, column, _FILE_FONT_SIZE)
            Catch ex As Exception
                PLLog.Error(ex, Common.PROJECT_NAME)
                Throw (ex)
                ' TODO: PLException.PLApplicationException.Publish(ex)
            End Try
        Next innerDir

        ' Add the numberFolders, numberFiles, and sizeFiles to the folder row 
        ' now that we know stuff about the files in the current folder and below

        numberFoldersCummulativeLocal += numberFoldersLocal
        numberFilesCummulativeLocal += numberFilesLocal
        sizeFilesCummulativeLocal += sizeFilesLocal

        Try
            UpdateFolderRow( _
                currentRow, currentColumn, _
                numberFoldersCummulativeLocal, numberFilesCummulativeLocal, sizeFilesCummulativeLocal, _
                numberFoldersLocal, numberFilesLocal, sizeFilesLocal, _
                maxLastCreateDate, maxLastWriteDate, maxLastAccessDate, _
                fontSize)
        Catch ex As Exception
            PLLog.Error(ex, Common.PROJECT_NAME)
            Throw (ex)  ' So we can add code to catch later
        End Try

        numberFoldersCummulative += numberFoldersCummulativeLocal
        numberFilesCummulative += numberFilesCummulativeLocal
        sizeFilesCummulative += sizeFilesCummulativeLocal

        If _SkipFoldersWithNoFiles Then
            If 0 = numberFilesLocal And 0 = numberFilesCummulative Then
                Globals.ThisAddIn.Application.ActiveSheet.Rows(currentRow).Delete()
                _Row -= 1
            End If
        End If

        ' Save highest column used
        If intColMax > _Column Then
            _Column = intColMax
        End If
    End Sub ' ListFolders()


    '----------------------------------------------------------------------
    '
    ' FolderMapSheet_Format
    '
    '
    ' Formats the Folder Map Sheet and Page.  Can this be changed to
    ' call the common worksheet format thing?
    '
    '----------------------------------------------------------------------

    Private Shared Sub FormatFolderMapSheet(ByRef strSheetName As String, Optional ByRef enumOrientation As Microsoft.Office.Interop.Excel.XlPageOrientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlPortrait)
        Dim i As Short

        With Globals.ThisAddIn.Application
            For i = _INITIAL_COL To _Column
                .Columns(i).ColumnWidth = _COL_WIDTH
            Next i

            .Columns(_Column + 1).ColumnWidth = _NOTE_WIDTH

            .ActiveSheet.Range("A2:I2").Select()
            With .Selection
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlGeneral
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlBottom
                .WrapText = True
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = Microsoft.Office.Interop.Excel.Constants.xlContext
                .MergeCells = False
                .Font.Bold = True
            End With

            .Columns("A:A").ColumnWidth = 13.57
            .Columns("B:B").ColumnWidth = 13.57
            .Columns("C:C").ColumnWidth = 13.43

            .Columns("E:E").ColumnWidth = 6.43
            .Columns("F:F").ColumnWidth = 7.71
            .Columns("G:G").ColumnWidth = 13.57
            .Columns("H:H").ColumnWidth = 13.57
            .Columns("I:I").ColumnWidth = 13.43

            .Range("K2").Select()

            With .Selection.Font
                .Name = "Arial"
                .Size = 12
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleNone
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Bold = True
            End With

            .Columns("A:J").Select()
            .Selection.Columns.Group()
            .Columns("A:D").Group()
            .Columns("G:I").Group()

            .Range("A1").Select()

        End With ' g_Excel
    End Sub ' FormatFolderMapSheet

    Private Shared Sub AddFolderRow( _
        ByVal dirInfo As FileInfo, _
        Optional ByVal column As Short = 1, _
        Optional ByVal fontSize As Short = 10, _
        Optional ByVal makeBold As Boolean = False, _
        Optional ByVal fontColor As Integer = System.ConsoleColor.Black _
    )
        Dim strS As String

        If True = _GroupResults Then
            If _GroupLevel = column Then
                ' Mark starting point.  We may need it later for grouping.  Don't reset.
                If 0 = _GroupStartRow Then
                    _GroupStartRow = _Row
                End If
            End If
        End If

        If _PatternMatchFolderHighlight Then
            If (Regex.Match(dirInfo.Name, _FolderMatchPattern).Success) Then
                fontColor = _FolderMatchColor
            End If
        End If

        With Globals.ThisAddIn.Application
            With .ActiveSheet.Cells(_Row, column)
                ' If we are not displaying folders but for some reason
                ' are calling AddFolderRow(), then use full path.

                If _ShowFolders Then
                    .Value = String.Format("{0}\", dirInfo.Name)
                Else
                    .Value = String.Format("{0}\", dirInfo.FullName)
                End If

                With .Font
                    .Bold = makeBold
                    .Size = fontSize
                    '.ColorIndex = fontColor
                    .Color = fontColor
                End With
            End With
        End With

        ' Now check if we need to do any grouping.

        If _GroupStartRow Then
            ' We have been adding rows to a possible grouping set.
            If _GroupLevel > column Then
                ' We have transitioned up the chain above the grouping set.
                strS = _GroupStartRow & ":" & _Row - 1
                Globals.ThisAddIn.Application.ActiveSheet.Rows(strS).Group()
                ' Reset the grouping counter.
                _GroupStartRow = 0
            End If
        End If

        ' Next row to add content
        _Row += 1
    End Sub

    '$-------------------------------------------------------------------
    '
    ' $Name        :AddFileRow$
    ' $Type        :$
    '
    ' $Arguments   :$
    '
    ' $Alters      :m_lngRow$
    '
    ' $Description :
    '   Adds a line to the spreadsheet indented at the appropriate level.$
    '
    ' $ToDo        :$
    '
    '$$------------------------------------------------------------------

    Private Shared Sub AddFileRow( _
        ByVal fileInfo As FileInfo, _
        Optional ByVal column As Short = 1, _
        Optional ByVal fontSize As Short = 10, _
        Optional ByVal makeBold As Boolean = False, _
        Optional ByVal fontColor As Integer = System.ConsoleColor.Black _
    )
        Dim strS As String

        If True = _GroupResults Then
            If _GroupLevel = column Then
                ' Mark starting point.  We may need it later for grouping.  Don't reset.
                If 0 = _GroupStartRow Then
                    _GroupStartRow = _Row
                End If
            End If
        End If

        With Globals.ThisAddIn.Application
            With .ActiveSheet.Cells(_Row, column)
                ' If we are not displaying folders but for some reason 
                ' are calling AddFileRow(), then use full path.

                If _ShowFolders Then
                    .Value = String.Format(" - {0}", fileInfo.Name)
                Else
                    .Value = String.Format(" - {0}", fileInfo.FullName)
                End If

                With .Font
                    .Bold = makeBold
                    .Size = fontSize
                    '.ColorIndex = fontColor
                    .Color = fontColor
                End With
            End With

            ' Start at _FILE_INFO_COL + 1 as we don't display file count on file row.

            With .ActiveSheet.Cells(_Row, _FILE_INFO_COL + 1)
                .Value = fileInfo.Length
                .NumberFormat = "#,##0_);(#,##0)"

                With .Font
                    .Bold = makeBold
                    .Size = fontSize
                End With
            End With

            Dim rng As Microsoft.Office.Interop.Excel.Range
            Dim dateD As Date

            rng = .ActiveSheet.Cells(_Row, _FILE_INFO_COL + 2)
            dateD = fileInfo.CreationTime
            rng.Value = dateD

            ColorCodeDate(rng, dateD, False, fontSize, _DateType.LastCreate)

            rng = .ActiveSheet.Cells(_Row, _FILE_INFO_COL + 3)
            dateD = fileInfo.LastWriteTime
            rng.Value = dateD

            ColorCodeDate(rng, dateD, False, fontSize, _DateType.LastWrite)

            rng = .ActiveSheet.Cells(_Row, _FILE_INFO_COL + 4)
            dateD = fileInfo.LastAccessTime
            rng.Value = dateD

            ColorCodeDate(rng, dateD, False, fontSize, _DateType.LastAccess)

        End With

        ' Now check if we need to do any grouping.

        If _GroupStartRow Then
            ' We have been adding rows to a possible grouping set.
            If _GroupLevel > column Then
                ' We have transitioned up the chain above the grouping set.
                strS = _GroupStartRow & ":" & _Row - 1
                Globals.ThisAddIn.Application.ActiveSheet.Rows(strS).Group()
                ' Reset the grouping counter.
                _GroupStartRow = 0
            End If
        End If

        ' Next row to add content.
        _Row = _Row + 1
    End Sub ' Row_Add

    Private Shared Sub AddErrorRow( _
        ByVal errorInfo As String, _
        Optional ByVal column As Short = 1, _
        Optional ByVal fontSize As Short = 10, _
        Optional ByVal makeBold As Boolean = False, _
        Optional ByVal fontColor As Integer = System.ConsoleColor.Black _
    )
        Dim strS As String

        '    Debug.Print m_lngRow & " " & intCol & ":" & m_intGroupStartRow & " " & strText

        If True = _GroupResults Then
            If _GroupLevel = column Then
                ' Mark starting point.  We may need it later for grouping.  Don't reset.
                If 0 = _GroupStartRow Then
                    _GroupStartRow = _Row
                End If
            End If
        End If

        With Globals.ThisAddIn.Application
            With .ActiveSheet.Cells(_Row, column)
                .Value = String.Format(" - {0}", errorInfo)

                With .Font
                    .Bold = makeBold
                    .Size = fontSize
                    '.ColorIndex = fontColor
                    .Color = fontColor
                End With
            End With
        End With

        ' Now check if we need to do any grouping.

        If _GroupStartRow Then
            ' We have been adding rows to a possible grouping set.
            If _GroupLevel > column Then
                ' We have transitioned up the chain above the grouping set.
                strS = _GroupStartRow & ":" & _Row - 1
                Globals.ThisAddIn.Application.ActiveSheet.Rows(strS).Group()
                ' Reset the grouping counter.
                _GroupStartRow = 0
            End If
        End If

        ' Next row to add content.
        _Row = _Row + 1
    End Sub

    Private Shared Sub UpdateFolderRow( _
        ByVal row As Integer, _
        ByVal column As Integer, _
        ByVal numberFoldersCummulative As Integer, _
        ByVal numberFilesCummulative As Integer, _
        ByVal sizeFilesCummulative As Long, _
        ByVal numberFolders As Integer, _
        ByVal numberFiles As Integer, _
        ByVal sizeFiles As Long, _
        ByVal maxLastCreateDate As Date, _
        ByVal maxLastWriteDate As Date, _
        ByVal maxLastAccessDate As Date, _
        Optional ByVal fontSize As Short = 10 _
    )
        With Globals.ThisAddIn.Application
            With .ActiveSheet.Cells(row, _FOLDER_INFO_COL)
                .Value = numberFoldersCummulative
                .NumberFormat = "#,##0_);(#,##0)"

                With .Font
                    ' .Bold = blnBold
                    .Size = fontSize
                End With
            End With

            With .ActiveSheet.Cells(row, _FOLDER_INFO_COL + 1)
                .Value = numberFilesCummulative
                .NumberFormat = "#,##0_);(#,##0)"

                With .Font
                    ' .Bold = blnBold
                    .Size = fontSize
                End With
            End With

            With .ActiveSheet.Cells(row, _FOLDER_INFO_COL + 2)
                .Value = sizeFilesCummulative
                .NumberFormat = "#,##0_);(#,##0)"

                With .Font
                    ' .Bold = blnBold
                    .Size = fontSize
                End With
            End With

            With .ActiveSheet.Cells(row, _FILE_INFO_COL)
                .Value = numberFiles
                .NumberFormat = "#,##0_);(#,##0)"

                With .Font
                    ' .Bold = blnBold
                    .Size = fontSize
                End With
            End With

            With .ActiveSheet.Cells(row, _FILE_INFO_COL + 1)
                .Value = sizeFiles
                .NumberFormat = "#,##0_);(#,##0)"

                With .Font
                    ' .Bold = blnBold
                    .Size = fontSize
                End With
            End With

            Dim rng As Microsoft.Office.Interop.Excel.Range
            Dim dateD As Date

            rng = .ActiveSheet.Cells(row, _FILE_INFO_COL + 2)
            dateD = maxLastCreateDate
            rng.Value = dateD

            ColorCodeDate(rng, dateD, False, fontSize, _DateType.LastCreate)

            rng = .ActiveSheet.Cells(row, _FILE_INFO_COL + 3)
            dateD = maxLastWriteDate
            rng.Value = dateD

            ColorCodeDate(rng, dateD, False, fontSize, _DateType.LastWrite)

            rng = .ActiveSheet.Cells(row, _FILE_INFO_COL + 4)
            dateD = maxLastAccessDate
            rng.Value = dateD

            ColorCodeDate(rng, dateD, False, fontSize, _DateType.LastAccess)
        End With
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <param name="dt"></param>
    ''' <param name="makeBold"></param>
    ''' <param name="fontSize"></param>
    ''' <param name="dateType"></param>
    ''' <remarks></remarks>
    ''' TODO: Maybe pass in format structure
    ''' 
    Private Shared Sub ColorCodeDate( _
            ByVal rng As Microsoft.Office.Interop.Excel.Range, _
            ByVal dt As DateTime, _
            ByVal makeBold As Boolean, _
            ByVal fontSize As Short, _
            ByVal dateType As _DateType _
        )

        If Not _ColorCodeDates Then
            With rng.Font
                .Bold = makeBold
                .Size = fontSize
            End With
        Else
            With rng.Font
                .Bold = makeBold
                .Size = fontSize
            End With

            Select Case dateType
                Case _DateType.LastCreate
                    If (DateTime.Compare(DateTime.Now, dt.AddMonths(_MonthsSinceCreated)) > 0) Then
                        With rng.Font
                            '.ColorIndex = _MonthsCreatedColor
                            .Color = _MonthsCreatedColor
                        End With
                    Else
                        'rng.Font.ColorIndex = _MonthsDefaultColor
                        rng.Font.Color = _MonthsDefaultColor
                    End If

                Case _DateType.LastWrite
                    If (DateTime.Compare(DateTime.Now, dt.AddMonths(_MonthsSinceWritten)) > 0) Then
                        With rng.Font
                            '.ColorIndex = _MonthsWrittenColor
                            .Color = _MonthsWrittenColor
                        End With
                    Else
                        'rng.Font.ColorIndex = _MonthsDefaultColor
                        rng.Font.Color = _MonthsDefaultColor
                    End If

                Case _DateType.LastAccess
                    If (DateTime.Compare(DateTime.Now, dt.AddMonths(_MonthsSinceAccessed)) > 0) Then
                        With rng.Font
                            '.ColorIndex = _MonthsAccessedColor
                            .Color = _MonthsAccessedColor
                        End With
                    Else
                        'rng.Font.ColorIndex = _MonthsDefaultColor
                        rng.Font.Color = _MonthsDefaultColor
                    End If

            End Select
        End If
    End Sub

    Private Shared Function HasIllegalFileNameCharacters(ByVal name As String) As Boolean
        Dim illegalCharactersMatch As Match = Regex.Match(name, _IllegalFileCharacters)

        Return illegalCharactersMatch.Success

    End Function

    Private Shared Function HasIllegalFolderNameCharacters(ByVal name As String) As Boolean
        Dim illegalCharactersMatch As Match = Regex.Match(name, _IllegalFolderCharacters)

        Return illegalCharactersMatch.Success

    End Function

    Private Sub Action(ByVal Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean)
        Try
            CreateFolderMap()
        Catch ex As Exception
            MessageBox.Show(String.Format("Exception: {0}.{1}() - {2}",
                         System.Reflection.Assembly.GetExecutingAssembly().FullName,
                         System.Reflection.MethodInfo.GetCurrentMethod().Name,
                         ex.ToString()
                         ))
        End Try
    End Sub
#End Region
    ' Everything below is experimental

    Private Class FolderInformation
        Private _path As String

        Private _countLocalFolders As Integer
        Private _countLocalFiles As Integer
        Private _sizeLocalFiles As Long

        Private _countTotalFolders As Integer
        Private _countTotalFiles As Integer
        Private _sizeTotaFiles As Long

        Public Sub ListFiles()

        End Sub

        Public Sub ListFolders()

        End Sub
    End Class

    Private Class FileInformation

        Public Sub New()

        End Sub
    End Class

    Public Class ExcelOutputFormatter
        Implements IOutputFormatter

        Public _Row As Integer

        Public Sub AddFileRow(ByVal fileInfo As System.IO.FileInfo, Optional ByVal column As Short = 1, Optional ByVal fontSize As Short = 10, Optional ByVal makeBold As Boolean = False) Implements IOutputFormatter.AddFileRow

        End Sub

        Public Sub AddFolderRow(ByVal folderName As String, Optional ByVal column As Short = 1, Optional ByVal fontSize As Short = 10, Optional ByVal blnBold As Boolean = False) Implements IOutputFormatter.AddFolderRow

        End Sub

        Public Sub AddRow(ByVal strText As String, Optional ByVal intCol As Short = 1, Optional ByVal intSize As Short = 10, Optional ByVal blnBold As Boolean = False) Implements IOutputFormatter.AddRow

        End Sub

        Public Sub ColorCodeDate(ByVal rng As Microsoft.Office.Interop.Excel.Range, ByVal dt As Date, ByVal makeBold As Boolean, ByVal fontSize As Integer, ByVal dateType As _DateType) Implements IOutputFormatter.ColorCodeDate

        End Sub

        Public Sub UpdateFolderRow(ByVal row As Integer, ByVal column As Integer, ByVal numberFoldersCummulative As Integer, ByVal numberFilesCummulative As Integer, ByVal sizeFilesCummulative As Long, ByVal numberFolders As Integer, ByVal numberFiles As Integer, ByVal sizeFiles As Long, ByVal maxLastUpdateDate As Date, ByVal maxLastAccessDate As Date, Optional ByVal fontSize As Short = 10) Implements IOutputFormatter.UpdateFolderRow

        End Sub
    End Class

    Interface IOutputFormatter
        ' TODO: Need to figure out first parameter

        Sub AddRow( _
            ByVal strText As String, _
            Optional ByVal intCol As Short = 1, _
            Optional ByVal intSize As Short = 10, _
            Optional ByVal blnBold As Boolean = False _
        )

        Sub AddFolderRow( _
            ByVal folderName As String, _
            Optional ByVal column As Short = 1, _
            Optional ByVal fontSize As Short = 10, _
            Optional ByVal blnBold As Boolean = False _
        )

        Sub AddFileRow( _
            ByVal fileInfo As System.IO.FileInfo, _
            Optional ByVal column As Short = 1, _
            Optional ByVal fontSize As Short = 10, _
            Optional ByVal makeBold As Boolean = False _
        )

        Sub UpdateFolderRow( _
            ByVal row As Integer, _
            ByVal column As Integer, _
            ByVal numberFoldersCummulative As Integer, _
            ByVal numberFilesCummulative As Integer, _
            ByVal sizeFilesCummulative As Long, _
            ByVal numberFolders As Integer, _
            ByVal numberFiles As Integer, _
            ByVal sizeFiles As Long, _
            ByVal maxLastUpdateDate As DateTime, _
            ByVal maxLastAccessDate As DateTime, _
            Optional ByVal fontSize As Short = 10 _
        )

        Sub ColorCodeDate( _
            ByVal rng As Microsoft.Office.Interop.Excel.Range, _
            ByVal dt As DateTime, _
            ByVal makeBold As Boolean, _
            ByVal fontSize As Integer, _
            ByVal dateType As _DateType _
        )

    End Interface
End Class
