Imports System.IO

Imports PacificLife.Life

Public Class FileOutput
    Public Sub CreateOutput( _
        ByRef stsadmRequestList As List(Of STSADMRequestData), _
        ByVal outputFolder As String, _
        ByVal outputFileName As String _
    )
        PLLog.Trace("Enter", Globals.cPLLog_Category_Name)

        Dim sb As New StringBuilder

        '----------------------------------------------------
        ' Modify this section to produce the lines to output
        '----------------------------------------------------
        sb.AppendLine("REM STSADM Commands produced by CreateSharePointSTSADMCommands.xlsx")
        sb.AppendLine()
        sb.Append("SET STSADM=""C:\Program Files\Common Files\Microsoft Shared\web server extensions\12\BIN\stsadm""")
        sb.AppendLine()
        sb.AppendLine()

        Dim shouldReturn As Boolean

        ' Add request info for first request.  Too much stuff in file if do in loop as below.

        AddRequestInfo(stsadmRequestList(0), sb)

        For Each request As STSADMRequestData In stsadmRequestList
            'AddRequestInfo(request, sb)

            Select Case request.RequestType.ToLower
                Case "path"
                    AddAddPathOutputLine(request, sb, shouldReturn)

                Case "site"
                    AddCreateSiteOutputLine(request, sb, shouldReturn)

                Case "web"
                    AddCreateWebOutputLine(request, sb, shouldReturn)

                Case Else
                    MessageBox.Show(String.Format("Unrecognized TargetType: {0}", request.RequestType))

            End Select

            If shouldReturn Then
                Return
            End If
        Next

        sb.AppendLine()

        '----------------------------------------------------
        ' Save the output
        '----------------------------------------------------

        If Not Util.File.ValidOutputFolder(outputFolder, True) Then
            MessageBox.Show(String.Format("Output Folder ({0}) does not exist", outputFolder))
            Return
        End If

        Dim fileStream As FileStream
        fileStream = System.IO.File.Create(String.Format("{0}\{1}", outputFolder, outputFileName))
        Dim sw As StreamWriter = New StreamWriter(fileStream)

        sw.Write(sb.ToString())
        sw.Flush()
        sw.Close()

        'fileStream.Flush()
        fileStream.Close()

        MessageBox.Show("Created file: " & fileStream.Name)

        PLLog.Trace("Exit", Globals.cPLLog_Category_Name)
    End Sub

    Private Sub AddRequestInfo(ByVal requestData As STSADMRequestData, ByVal sb As StringBuilder)
        sb.AppendLine(String.Format("REM Requested By {0} on {1} needed by {2} for {3}", requestData.RequestedBy, requestData.RequestDate, requestData.DateNeeded, requestData.Purpose))
    End Sub

    Private Shared Sub AddAddPathOutputLine(ByRef requestData As STSADMRequestData, ByVal sb As StringBuilder, ByRef shouldReturn As Boolean)
        shouldReturn = False

        sb.AppendFormat("%STSADM% -o {0,-10}", "addpath")

        If requestData.SiteUrl <> "" Then
            sb.AppendFormat(" -url {0,-30}", requestData.SiteUrl)
        Else
            MessageBox.Show("Must provide value for SiteURL")
            shouldReturn = True : Return
        End If

        If requestData.ManagedPathType <> "" Then
            sb.AppendFormat(" -type {0,-10}", requestData.ManagedPathType)
        Else
            MessageBox.Show("Must provide value for ManagedPathType")
            shouldReturn = True : Return
        End If

        sb.AppendLine()
    End Sub

    Private Shared Sub AddCreateSiteOutputLine(ByRef request As STSADMRequestData, ByVal sb As StringBuilder, ByRef shouldReturn As Boolean)
        shouldReturn = False

        sb.AppendFormat("%STSADM% -o {0,-10}", "createsite")

        If request.SiteUrl <> "" Then
            ' TODO: Add URL validation.  This should be done in SiteCollectionsDataInfo class.
            sb.AppendFormat(" -url {0,-30}", request.SiteUrl)
        Else
            MessageBox.Show("Must provide value for SiteURL")
            shouldReturn = True : Return
        End If

        If request.SiteTemplate <> "" Then
            sb.AppendFormat(" -sitetemplate {0,-10}", request.SiteTemplate)
        Else
            MessageBox.Show("Must provide value for SiteTemplate")
            shouldReturn = True : Return
        End If

        If request.Title <> "" Then
            sb.AppendFormat(" -title {0,-20}", request.Title)
        Else
            MessageBox.Show("Must provide value for Title")
            shouldReturn = True : Return
        End If

        If request.Description <> "" Then
            sb.Append(" -description " & request.Description)
        End If

        ' Primary Site Collection Administrator

        If request.PrimaryOwnerEmail <> "" Then
            sb.Append(" -owneremail " & request.PrimaryOwnerEmail)
        Else
            MessageBox.Show("Must provide value for PrimaryOwnerEmail")
            shouldReturn = True : Return
        End If

        If request.PrimaryOwnerLogin <> "" Then
            sb.Append(" -ownerlogin " & request.PrimaryOwnerLogin)
        Else
            MessageBox.Show("Must provide value for PrimaryOwnerLogin")
            shouldReturn = True : Return
        End If

        If request.PrimaryOwnerName <> "" Then
            sb.Append(" -ownername " & request.PrimaryOwnerName)
        End If

        ' Secondary Site Collection Administrator

        If request.SecondaryOwnerLogin <> "" Then
            sb.Append(" -secondarylogin " & request.SecondaryOwnerLogin)
            'Else
            '    MessageBox.Show("Must provide value for SecondaryOwnerLogin")
            '    shouldReturn = True : Return
        End If

        If request.SecondaryOwnerName <> "" Then
            sb.Append(" -secondaryname " & request.SecondaryOwnerName)
        End If

        If request.SecondaryOwnerEmail <> "" Then
            sb.Append(" -secondaryemail " & request.SecondaryOwnerEmail)
        End If

        'If scInfo.Offset(i, Globals.cSC_LCID_Offset).Value <> "" Then
        '    sb.Append(" -lcid " & scInfo.Offset(i, Globals.cSC_LCID_Offset).Value)
        'End If

        If request.Quota <> "" Then
            sb.Append(" -quota " & request.Quota)
        End If

        'If scRequest.WebApplication <> "" Then
        '    ' TODO: Add URL validation
        '    sb.AppendFormat(" -hostheaderwebapplicationurl {0,-30}", scRequest.WebApplicationUrl)
        'Else
        '    MessageBox.Show("Must provide value for SiteURL")
        '    shouldReturn = True : Return
        'End If

        sb.AppendLine()
    End Sub

    Private Shared Sub AddCreateWebOutputLine(ByRef request As STSADMRequestData, ByVal sb As StringBuilder, ByRef shouldReturn As Boolean)
        shouldReturn = False

        sb.AppendFormat("%STSADM% -o {0,-10}", "createweb")

        If request.SiteUrl <> "" Then
            ' TODO: Add URL validation.  This should be done in SiteCollectionsDataInfo class.
            sb.AppendFormat(" -url {0,-30}", request.SiteUrl)
        Else
            MessageBox.Show("Must provide value for SiteURL")
            shouldReturn = True : Return
        End If

        If request.SiteTemplate <> "" Then
            sb.AppendFormat(" -sitetemplate {0,-10}", request.SiteTemplate)
        Else
            MessageBox.Show("Must provide value for SiteTemplate")
            shouldReturn = True : Return
        End If

        'If scInfo.Offset(i, Globals.cSC_LCID_Offset).Value <> "" Then
        '    sb.Append(" -lcid " & scInfo.Offset(i, Globals.cSC_LCID_Offset).Value)
        'End If

        If request.Title <> "" Then
            sb.AppendFormat(" -title {0,-20}", request.Title)
        Else
            MessageBox.Show("Must provide value for Title")
            shouldReturn = True : Return
        End If

        If request.Description <> "" Then
            sb.Append(" -description " & request.Description)
        End If

        If request.UniquePermissions <> "" Then
            sb.Append(" -unique ")
        End If

        sb.AppendLine()
    End Sub


End Class
