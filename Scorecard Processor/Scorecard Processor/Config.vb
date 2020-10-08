Imports PacificLife.Life
Imports System.Collections.Generic

Public Class Config
    Private Shared _isInitialized As Boolean = False

    Private Shared _configInfo As Data.DataSet
    Private Shared _scorecardInfo As Data.DataSet

    Private Shared _teamNameToCells As Dictionary(Of String, String)

    Public Shared ReadOnly Property TeamNameToCells() As Dictionary(Of String, String)
        Get
            If _teamNameToCells Is Nothing Then
                _teamNameToCells = New Dictionary(Of String, String)

                Dim wb As Excel.Workbook = Globals.ThisAddIn.Application.ActiveWorkbook
                Dim teamNameToCell As Excel.Range = wb.Names.Item("TeamName_To_Cell_Array").RefersToRange
                Util.DisplayExcelRange(teamNameToCell)

                For i As Integer = 1 To teamNameToCell.Rows.Count
                    If teamNameToCell.Cells(i, 1).Value <> "" Then
                        _teamNameToCells.Add(teamNameToCell.Cells(i, 1).Value, teamNameToCell.Cells(i, 2).Value)
                    End If
                Next
            End If

            Return _teamNameToCells
        End Get
        'Set(ByVal value As Dictionary(Of String, String))

        'End Set
    End Property

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

    Public Shared Property Teams() As Data.DataTable
        Get
            Return ConfigInfo.Tables("team")
        End Get

        Set(ByVal value As Data.DataTable)

        End Set
    End Property

    Public Shared Property DefinedNames() As Data.DataTable
        Get
            Return ScorecardInfo.Tables("name")
        End Get

        Set(ByVal value As Data.DataTable)

        End Set
    End Property

    Public Shared Sub IntializeApplication()
        PLLog.Trace1("Enter", "Scorecard")

        If Not _isInitialized Then
            If Globals.ThisAddIn.Application.ActiveWorkbook Is Nothing Then
                MessageBox.Show("Open ScoreCard workbook before loading config data")
            Else
                LoadTeamConfigDataFromXMLFile()
                LoadScorecardConfigDataFromXMLFile()

                _isInitialized = True
            End If
        End If

        PLLog.Trace1("Exit", "Scorecard")
    End Sub

    Private Shared Sub LoadTeamConfigDataFromXMLFile()
        PLLog.Trace1("Enter", "Scorecard")

        If Globals.ThisAddIn.Application.ActiveWorkbook Is Nothing Then
            MessageBox.Show("Open ScoreCard workbook before loading config data")
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
        _configInfo.ReadXml(workbookPath & "\config-teams.xml", XmlReadMode.Auto)

        ' This shows how ReadXml parsed the file.
        DisplayDataSet(ConfigInfo)

        PLLog.Trace1("Exit", "Scorecard")
    End Sub

    Private Shared Sub LoadScorecardConfigDataFromXMLFile()
        PLLog.Trace1("Enter", "Scorecard")

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
        DisplayDataSet(ScorecardInfo)

        PLLog.Trace1("Exit", "Scorecard")
    End Sub

    Public Shared Sub ReIntializeApplication()
        PLLog.Trace1("Enter", "Scorecard")

        _teamNameToCells = Nothing

        If Not _configInfo Is Nothing Then
            _configInfo.Dispose()
            _configInfo = Nothing
        End If

        _isInitialized = False

        PLLog.Trace1("Exit", "Scorecard")
    End Sub

    Private Shared Sub DisplayDataSet(ByVal dataSet As Data.DataSet)
        DisplayTables(dataSet.Tables)
    End Sub

    Private Shared Sub DisplayTables(ByVal tables As DataTableCollection)
        PLLog.Trace1("Enter", "Scorecard")

        For Each dataTable As Data.DataTable In tables
            Trace.WriteLine("Table:   >" & dataTable.TableName & "<")
            Trace.WriteLine("Columns:")

            For Each dataColumn As Data.DataColumn In dataTable.Columns
                Trace.Write(" >" & dataColumn.ColumnName & "<")
            Next

            Trace.WriteLine("")
            Trace.WriteLine("Rows:" & vbCrLf)

            For Each dataRow As Data.DataRow In dataTable.Rows
                For Each columnName As Data.DataColumn In dataTable.Columns
                    Trace.Write(" >" & dataRow.Item(columnName.ColumnName).ToString & "<")
                Next
                Trace.WriteLine("")
            Next
        Next

        PLLog.Trace1("Exit", "Scorecard")
    End Sub
End Class
