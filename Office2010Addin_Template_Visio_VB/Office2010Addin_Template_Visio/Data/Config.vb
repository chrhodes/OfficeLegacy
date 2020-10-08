Imports PacificLife.Life
Imports System.Collections.Generic
Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Data
Imports System.Diagnostics

Public Class Config
    '----------------------------------------------------------------------------------------------------
    ' Class Config
    '
    ' This class contains methods to load and maintain configuration information used by the addin.
    ' The configuration information can be hard coded e.g. come from the Globals file or can come
    ' from an XML configuration file.  If a configuration file is used an dataset local to this class
    ' contains the data once loaded.
    '
    ' Two examples are provided.  One for general config info contained in ConfigInfo (_configinfo)
    ' and one for ScoreCard Info contained in ScoreCardInfo (_scorecardinfo)
    ' Once loaded into the dataset the config info appears as tables.

    Private Shared _isInitialized As Boolean = False

    Private Shared _configInfo As Data.DataSet
    Private Shared _scorecardInfo As Data.DataSet

    Private Shared _teamNameToCells As Dictionary(Of String, String)


    Public Shared Property ConfigInfo() As Data.DataSet
        Get
            If _configInfo Is Nothing Then
                LoadTeamConfigDataFromXMLFile()
            End If

            Return _configInfo
        End Get
        Set(ByVal value As Data.DataSet)
            _configInfo = value
        End Set
    End Property

    Public Shared Property DefinedNames() As Data.DataTable
        Get
            Return ScorecardInfo.Tables("name")
        End Get

        Set(ByVal value As Data.DataTable)

        End Set
    End Property

    Public Shared Property ScorecardInfo() As Data.DataSet
        Get
            If _scorecardInfo Is Nothing Then
                LoadScorecardConfigDataFromXMLFile()
            End If

            Return _scorecardInfo
        End Get

        Set(ByVal value As Data.DataSet)
            _scorecardInfo = value
        End Set
    End Property

    'Public Shared ReadOnly Property TeamNameToCells() As Dictionary(Of String, String)
    '    Get
    '        If _teamNameToCells Is Nothing Then
    '            _teamNameToCells = New Dictionary(Of String, String)

    '            Dim wb As Excel.Workbook = Globals.ThisAddIn.Application.ActiveWorkbook
    '            Dim teamNameToCell As Excel.Range = wb.Names.Item("TeamName_To_Cell_Array").RefersToRange
    '            Common.ExcelHelper.DisplayExcelRange(teamNameToCell)

    '            For i As Integer = 1 To teamNameToCell.Rows.Count
    '                If teamNameToCell.Cells(i, 1).Value <> "" Then
    '                    _teamNameToCells.Add(teamNameToCell.Cells(i, 1).Value, teamNameToCell.Cells(i, 2).Value)
    '                End If
    '            Next
    '        End If

    '        Return _teamNameToCells
    '    End Get
    '    'Set(ByVal value As Dictionary(Of String, String))

    '    'End Set
    'End Property

    'Public Shared Property Teams() As Data.DataTable
    '    Get
    '        Return ConfigInfo.Tables("team")
    '    End Get

    '    Set(ByVal value As Data.DataTable)

    '    End Set
    'End Property



    Public Shared Sub IntializeApplication()
        PLLog.Trace1("Enter", Common.PROJECT_NAME)

        If Not _isInitialized Then
            If Globals.ThisAddIn.Application.ActiveWorkbook Is Nothing Then
                MessageBox.Show("Open workbook before loading config data")
            Else
                LoadTeamConfigDataFromXMLFile()
                'LoadScorecardConfigDataFromXMLFile()

                _isInitialized = True
            End If
        End If

        PLLog.Trace1("Exit", Common.PROJECT_NAME)
    End Sub

    Private Shared Sub LoadTeamConfigDataFromXMLFile()
        PLLog.Trace1("Enter", Common.PROJECT_NAME)

        If Globals.ThisAddIn.Application.ActiveWorkbook Is Nothing Then
            MessageBox.Show("Open workbook before loading config data")
            Return
        End If

        Dim workbookPath As String = Globals.ThisAddIn.Application.ActiveWorkbook.Path

        'Util.ApplicationInfo()

        If _configInfo Is Nothing Then
            _configInfo = New Data.DataSet
        Else
            ' TODO: Need to handle calling this again.  Probably just empty dataset and repopulate
            ' so code holding reference to Teams still works.
        End If

        _configInfo.Clear()
        _configInfo.ReadXml(String.Format("{0}\config-teams.xml", workbookPath), XmlReadMode.Auto)

        ' This shows how ReadXml parsed the file.
        'DisplayDataSet(ConfigInfo)

        PLLog.Trace1("Exit", Common.PROJECT_NAME)
    End Sub

    Private Shared Sub LoadScorecardConfigDataFromXMLFile()
        PLLog.Trace1("Enter", Common.PROJECT_NAME)

        If Globals.ThisAddIn.Application.ActiveWorkbook Is Nothing Then
            MessageBox.Show("Open ScoreCard workbook before loading config data")
            Return
        End If

        Dim workbookPath As String = Globals.ThisAddIn.Application.ActiveWorkbook.Path

        'Util.ApplicationInfo()

        If _scorecardInfo Is Nothing Then
            _scorecardInfo = New Data.DataSet
        Else
            ' TODO: Need to handle calling this again.  Probably just empty dataset and repopulate
            ' so code holding reference to Teams still works.
        End If

        _scorecardInfo.Clear()
        _scorecardInfo.ReadXml(workbookPath & "\config-scorecard.xml", XmlReadMode.Auto)

        ' This shows how ReadXml parsed the file.
        'DisplayDataSet(ScorecardInfo)

        PLLog.Trace1("Exit", Common.PROJECT_NAME)
    End Sub

    Public Shared Sub ReIntializeApplication()
        PLLog.Trace1("Enter", Common.PROJECT_NAME)

        _teamNameToCells = Nothing

        If Not _configInfo Is Nothing Then
            _configInfo.Dispose()
            _configInfo = Nothing
        End If

        _isInitialized = False

        PLLog.Trace1("Exit", Common.PROJECT_NAME)
    End Sub

    Private Shared Sub DisplayDataSet(ByVal dataSet As Data.DataSet)
        DisplayTables(dataSet.Tables)
    End Sub

    Private Shared Sub DisplayTables(ByVal tables As DataTableCollection)
        PLLog.Trace1("Enter", Common.PROJECT_NAME)

        For Each dataTable As Data.DataTable In tables
            Trace.WriteLine(String.Format("Table:   >{0}<", dataTable.TableName))
            Trace.WriteLine("Columns:")

            For Each dataColumn As Data.DataColumn In dataTable.Columns
                Trace.Write(String.Format(" >{0}<", dataColumn.ColumnName))
            Next

            Trace.WriteLine("")
            Trace.WriteLine(String.Format("Rows:{0}", vbCrLf))

            For Each dataRow As Data.DataRow In dataTable.Rows
                For Each columnName As Data.DataColumn In dataTable.Columns
                    Trace.Write(String.Format(" >{0}<", dataRow.Item(columnName.ColumnName)))
                Next
                Trace.WriteLine("")
            Next
        Next

        PLLog.Trace1("Exit", Common.PROJECT_NAME)
    End Sub
End Class
