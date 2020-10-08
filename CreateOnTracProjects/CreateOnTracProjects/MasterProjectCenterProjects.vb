Imports System.IO

Public Class MasterProjectCenterProjects

    Private Sub Sheet4_Startup(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Startup

    End Sub

    Private Sub Sheet4_Shutdown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shutdown

    End Sub

    Private Sub btnCreateMasterProjects_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateMasterProjects.Click
        Dim ws As Excel.Worksheet = Application.ActiveSheet
        Dim startingRow As Integer = Range("$C$2").Value
        Dim endingRow As Integer = Range("$C$3").Value
        Dim fileName As String = Range("$C$4").Value
        Dim projectInfo As Excel.Range
        Dim sb As New StringBuilder

        projectInfo = ws.Cells(startingRow, 2)

        sb.Append("<UpcomingProjects>" & vbCrLf)

        For i As Integer = 0 To endingRow - startingRow
            sb.Append("  <MasterProject")

            sb.Append(" ProjectName=""" & projectInfo.Offset(i, 0).Value & """")            ' ProjectName
            sb.Append(" ProjectDueDate=""" & projectInfo.Offset(i, 7).Value & """")         ' ProjectDueDate
            sb.Append(" Comments=""" & projectInfo.Offset(i, 13).Value & """")              ' Comments

            sb.Append("/>" & vbCrLf)
            System.Diagnostics.Debug.WriteLine(sb.ToString)
        Next

        sb.Append("</UpcomingProjects>")

        Dim fileStream As FileStream
        fileStream = System.IO.File.Create(fileName)
        Dim sw As StreamWriter = New StreamWriter(fileStream)
        sw.Write(sb.ToString())
        sw.Close()
        fileStream.Close()

        MessageBox.Show("MasterProjectCenter XML file: " & fileName & " complete")

        Return
    End Sub
End Class
