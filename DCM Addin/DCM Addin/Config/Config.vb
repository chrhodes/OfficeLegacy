Imports PacificLife.Life
Imports System.Collections.Generic

Public Class Config
    '**********************************************************************
    '
    ' Config.vb
    '
    ' This class is responsible for opening the configuration file and
    ' loading it into a DataSet that is used in various places.
    '
    ' Changes to this file are unlikely.  
    ' The action occurs in ProcessFile.vb
    '
    ' TODO: Could get fancy and validate the config file with a schema.
    '
    '**********************************************************************

    Private Shared _fileConfigInfo As Data.DataSet

    Public Shared Property FileConfigInfo() As Data.DataSet
        Get
            If _fileConfigInfo Is Nothing Then
                LoadFileConfigDataFromXMLFile(String.Empty)
            End If

            Return _fileConfigInfo
        End Get

        Set(ByVal value As Data.DataSet)
            _fileConfigInfo = value
        End Set
    End Property

    ' TODO: Need to think through when this is going to be called.  
    ' Seems like we want to call whenever we open a file.
    ' What happens if file is on SharePoint site?

    Public Shared Sub LoadFileConfigDataFromXMLFile(ByVal filePath As String)
        PLLog.Trace3("Enter", Globals.cPLLOG_NAME)

        If Globals.ThisAddIn.Application.ActiveWorkbook Is Nothing Then
            MessageBox.Show("Open workbook before loading config data")
            Return
        End If

        Dim workbookPath As String = Globals.ThisAddIn.Application.ActiveWorkbook.Path

        If _fileConfigInfo Is Nothing Then
            _fileConfigInfo = New Data.DataSet
        Else
            ' Handle calling this again.  Reset dataset and repopulate
            ' so code holding reference to FileConfigInfo still works.
            ' Clear() did not reset internal numbers.  Need to reset.
            '_fileConfigInfo.Clear()
            _fileConfigInfo.Reset()
        End If

        Try
            _fileConfigInfo.ReadXml(workbookPath & "\" & Globals.cCONFIG_FILE_NAME, XmlReadMode.Auto)
        Catch ex As Exception
            PLLog.Error(ex, Globals.cPLLOG_NAME)
        End Try

#If DEBUG Then
        DisplayDataSet(_fileConfigInfo)
#End If

        PLLog.Trace3("Exit", Globals.cPLLOG_NAME)
    End Sub

    Private Shared Sub DisplayDataSet(ByVal dataSet As Data.DataSet)
        DisplayTables(dataSet.Tables)
    End Sub

    Private Shared Sub DisplayTables(ByVal tables As DataTableCollection)
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
    End Sub
End Class
