Imports System.IO

Namespace Util
    Public Class File

        Public Shared Function IsValidOutputFolder(ByVal folderName As String, Optional ByVal promptToCreate As Boolean = False)
            Dim result As Boolean = False

            If System.IO.Directory.Exists(folderName) Then
                result = True
            Else
                If promptToCreate Then
                    Dim prompt As System.Windows.Forms.DialogResult

                    prompt = MessageBox.Show(String.Format("Output folder ({0}) does not exist.  Do you want to create it?", folderName), "Output Folder Does Not Exist", MessageBoxButtons.YesNo, MessageBoxIcon.Stop)

                    If prompt = DialogResult.Yes Then
                        Dim dirInfo As DirectoryInfo

                        dirInfo = Directory.CreateDirectory(folderName)
                        result = True
                    Else
                        result = False
                    End If
                End If
            End If

            Return result
        End Function

    End Class
End Namespace