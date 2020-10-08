Public Class Servers
    Dim _hosts() As String = _
    { _
        "ldspide01v.devlifeint.devpl01.net", _
        "ldspide02v.devlifeint.devpl01.net", _
        "ldspide03v.devlifeint.devpl01.net", _
        "ldspide04v.devlifeint.devpl01.net", _
        "ldspide05v.devlifeint.devpl01.net", _
        "lifesps601.devlifeint.devpl01.net", _
        "lifesps701.devlifeint.devpl01.net", _
        "lifesrch601.devlifeint.devpl01.net", _
        "lifesrch701.devlifeint.devpl01.net", _
        "lifesps401.tstlifeint.tstpl01.net", _
        "lifesps501.tstlifeint.tstpl01.net", _
        "lifesrch401.tstlifeint.tstpl01.net", _
        "lifesrch501.tstlifeint.tstpl01.net", _
        "lsspa01v.life.pacificlife.net", _
        "lsspa02v.life.pacificlife.net", _
        "lssps01v.life.pacificlife.net", _
        "lssps02v.life.pacificlife.net", _
        "lpspa01v.life.pacificlife.net", _
        "lpspa02v.life.pacificlife.net", _
        "lpsps01v.life.pacificlife.net", _
        "lpsps02v.life.pacificlife.net" _
    }

    Public ReadOnly Property Hosts() As String()
        Get
            Return _hosts
        End Get
    End Property

    Public Sub New(ByRef cache As System.Net.CredentialCache, ByVal webService As SystemManagementWS.WMIInfoWS)
        For Each host As String In Hosts
            Dim url As String = String.Format("http://{0}/SystemManagement/WMIInfoWS.asmx", host)

            ' Build the credential cache so we can get to all the places we need to.

            ' TODO: Perhaps regular expressions
            Select Case host
                ' Development Servers
                Case "ldspide01v.devlifeint.devpl01.net", _
                     "ldspide02v.devlifeint.devpl01.net", _
                     "ldspide03v.devlifeint.devpl01.net", _
                     "ldspide04v.devlifeint.devpl01.net", _
                     "ldspide05v.devlifeint.devpl01.net", _
                     "lifesps601.devlifeint.devpl01.net", _
                     "lifesps701.devlifeint.devpl01.net", _
                     "lifesrch601.devlifeint.devpl01.net", _
                     "lifesrch701.devlifeint.devpl01.net"

                    cache.Add(New Uri(url), "NTLM", New System.Net.NetworkCredential("dspappca", "Development2007", "DEVLIFEINT"))

                Case "lifesps401.tstlifeint.tstpl01.net", _
                     "lifesps501.tstlifeint.tstpl01.net", _
                     "lifesrch401.tstlifeint.tstpl01.net", _
                     "lifesrch501.tstlifeint.tstpl01.net"

                    cache.Add(New Uri(url), "NTLM", New System.Net.NetworkCredential("tspappca", "Testing2007", "TSTLIFEINT"))

                Case "lsspa01v.life.pacificlife.net", _
                     "lsspa02v.life.pacificlife.net", _
                     "lssps01v.life.pacificlife.net", _
                     "lssps02v.life.pacificlife.net"

                    cache.Add(New Uri(url), "NTLM", New System.Net.NetworkCredential("sspappca", "Staging2007", "PACIFICMUTUAL"))

                Case "lpspa01v.life.pacificlife.net", _
                     "lpspa02v.life.pacificlife.net", _
                     "lpsps01v.life.pacificlife.net", _
                     "lpsps02v.life.pacificlife.net"

                    cache.Add(New Uri(url), "NTLM", New System.Net.NetworkCredential("pspappca", "Production2007", "PACIFICMUTUAL"))
            End Select
        Next

        webService.Credentials = cache

    End Sub
End Class
