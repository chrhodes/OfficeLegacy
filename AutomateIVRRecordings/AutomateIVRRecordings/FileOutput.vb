Imports System.IO

Public Class FileOutput
    Public Sub AddCommandLine(ByVal streamWriter As StreamWriter, ByVal commandInput As String)
        Dim sb As StringBuilder = New StringBuilder(128)

        sb.AppendLine(String.Format("COMMAND <arg> <arg> {0}", commandInput))

        streamWriter.Write(sb.ToString())
        streamWriter.Flush()
    End Sub

    Public Sub CloseCommandFile(ByVal streamWriter As StreamWriter)
        Try
            streamWriter.Flush()
            streamWriter.Close()
        Catch ex As Exception
            MsgBox(String.Format("Cannot close stream {0}{1}", ControlChars.CrLf, ex.ToString()))
        End Try
    End Sub

    Public Function CreateCommandFile(ByVal outputFilePathAndName As String) As StreamWriter
        Dim streamWriter As StreamWriter

        Try
            streamWriter = File.CreateText(outputFilePathAndName)
        Catch ex As Exception
            MsgBox(String.Format("Cannot create {0}{1}{2}", outputFilePathAndName, ControlChars.CrLf, ex.ToString()))
            Return Nothing
        End Try

        Dim sb As New StringBuilder

        '----------------------------------------------------------
        ' Modify this section to produce the lines to output
        '----------------------------------------------------------

        sb.AppendLine("REM -----------------------------------------------------------------")
        sb.AppendLine("REM IVR File Processing Commands produced by AutomateIVRRecordings")
        sb.AppendLine("REM")
        sb.AppendLine("REM Tell story ...")
        sb.AppendLine("REM More story ...")
        sb.AppendLine("REM ")
        sb.AppendLine("REM -----------------------------------------------------------------")

        streamWriter.Write(sb.ToString())
        streamWriter.Flush()

        Return streamWriter
    End Function

    Public Sub CreateOutputFile(ByRef outputFilePathAndName As String, ByRef fileContents As String)
        ' This uses a FileStream and a StreamWriter
        Try
            Dim fs As FileStream
            fs = System.IO.File.Create(outputFilePathAndName)
            Dim sw As StreamWriter = New StreamWriter(fs)
            sw.Write(fileContents)
            sw.Flush()
            sw.Close()
            fs.Close()
        Catch ex As Exception
            MsgBox("Could not create output file: " & ControlChars.CrLf & ex.ToString())
        End Try
    End Sub
End Class
