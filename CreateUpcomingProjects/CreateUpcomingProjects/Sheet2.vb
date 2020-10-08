
Imports System.IO


Public Class Sheet2

    Private Sub Sheet2_Startup(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Startup

    End Sub

    Private Sub Sheet2_Shutdown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shutdown

    End Sub

    Private Sub btnCreateXMLFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateXMLFile.Click
        Dim ws As Excel.Worksheet = Application.ActiveSheet
        Dim startingRow As Integer = Range("$C$2").Value
        Dim endingRow As Integer = Range("$C$3").Value
        Dim fileName As String = Range("$C$4").Value
        Dim projectInfo As Excel.Range
        Dim sb As New StringBuilder

        projectInfo = ws.Cells(startingRow, 2)

        sb.Append("<UpcomingProjects>" & vbCrLf)

        For i As Integer = 0 To endingRow - startingRow
            sb.Append("  <Project")

            sb.Append(" ProjectName=""" & projectInfo.Offset(i, 0).Value & """")            ' ProjectName
            sb.Append(" ProjectCoordinator=""" & projectInfo.Offset(i, 1).Value & """")     ' ProjectCoordinator
            sb.Append(" Submitter=""" & projectInfo.Offset(i, 2).Value & """")              ' Submitter
            sb.Append(" AuthorOwner=""" & projectInfo.Offset(i, 3).Value & """")            ' AuthorOwner
            sb.Append(" WorkflowType=""" & projectInfo.Offset(i, 4).Value & """")           ' WorkflowType
            sb.Append(" ProductType=""" & projectInfo.Offset(i, 5).Value & """")            ' ProductType
            sb.Append(" IntendedAudience=""" & projectInfo.Offset(i, 6).Value & """")       ' IntendedAudience
            sb.Append(" ProjectDueDate=""" & projectInfo.Offset(i, 7).Value & """")         ' ProjectDueDate
            sb.Append(" PrintForPublication=""" & projectInfo.Offset(i, 8).Value & """")    ' PrintForPublication
            sb.Append(" CorporateWebReview=""" & projectInfo.Offset(i, 9).Value & """")     ' CorporateWebReview
            sb.Append(" LifelineWebReview=""" & projectInfo.Offset(i, 10).Value & """")     ' LifeLineWebReview
            sb.Append(" CorporateWebPublish=""" & projectInfo.Offset(i, 11).Value & """")   ' CorporateWebPublish
            sb.Append(" LifelineWebPublish=""" & projectInfo.Offset(i, 12).Value & """")    ' LifeLineWebPublish
            sb.Append(" Comments=""" & projectInfo.Offset(i, 13).Value & """")              ' Comments

            sb.Append("/>" & vbCrLf)
            System.Diagnostics.Debug.WriteLine(sb.ToString)
        Next

        sb.Append("</UpcomiingProjects>")

        Dim fileStream As FileStream
        fileStream = System.IO.File.Create(fileName)
        Dim sw As StreamWriter = New StreamWriter(fileStream)
        sw.Write(sb.ToString())
        sw.Close()
        fileStream.Close()

        'MessageBox.Show(sb.ToString())

        Return
    End Sub
End Class
