Imports Microsoft.Office.DocumentFormat.OpenXml
Imports System.IO.Packaging
Imports System.Xml
Imports System.IO
Imports System.Collections
Imports System.Diagnostics


' TODO: May want to add some validation to property settors.

Public Class STSADMRequestData

#Region "Properties"

    Private _requestType As String

    Public Property RequestType() As String
        Get
            Return _requestType
        End Get
        Set(ByVal value As String)
            _requestType = value
        End Set
    End Property

    Private _requestedBy As String

    Public Property RequestedBy() As String
        Get
            Return _requestedBy
        End Get
        Set(ByVal Value As String)
            _requestedBy = Value
        End Set
    End Property

    Private _requestDate As Date

    Public Property RequestDate() As Date
        Get
            Return _requestDate
        End Get
        Set(ByVal Value As Date)
            _requestDate = Value
        End Set
    End Property

    Private _dateNeeded As Date

    Public Property DateNeeded() As Date
        Get
            Return _dateNeeded
        End Get
        Set(ByVal Value As Date)
            _dateNeeded = Value
        End Set
    End Property

    Private _purpose As String

    Public Property Purpose() As String
        Get
            Return _purpose
        End Get
        Set(ByVal Value As String)
            _purpose = Value
        End Set
    End Property

    Private _sharePointFarm As String

    Public Property SharePointFarm() As String
        Get
            Return _sharePointFarm
        End Get
        Set(ByVal Value As String)
            _sharePointFarm = Value
        End Set
    End Property

    Private _webApplication As String

    Public Property WebApplication() As String
        Get
            Return _webApplication
        End Get
        Set(ByVal Value As String)
            _webApplication = Value
        End Set
    End Property

    Private _managedPathName As String

    Public Property ManagedPathName() As String
        Get
            Return _managedPathName
        End Get

        Set(ByVal Value As String)
            _managedPathName = Value
        End Set
    End Property

    Private _managedPathType As String

    Public Property ManagedPathType() As String
        Get
            Return _managedPathType
        End Get
        Set(ByVal Value As String)
            Select Case Value
                Case "explicit"
                    _managedPathType = "explicitinclusion"

                Case "wildcard"
                    _managedPathType = "wildcardinclusion"

                Case ""
                    _managedPathType = ""

                Case Else
                    MessageBox.Show("Unrecognized ManagedPathType: " & Value)

            End Select
        End Set
    End Property

    Private _urlName As String

    Public Property UrlName() As String
        Get
            Return _urlName
        End Get
        Set(ByVal Value As String)
            _urlName = Value
        End Set
    End Property

    Private _uniquePermissions As String

    Public Property UniquePermissions() As String
        Get
            Return _uniquePermissions
        End Get
        Set(ByVal Value As String)
            _uniquePermissions = Value
        End Set
    End Property

    Private _title As String

    Public Property Title() As String
        Get
            Return _title
        End Get
        Set(ByVal Value As String)
            _title = """" & Value & """"
        End Set
    End Property

    Private _description As String

    Public Property Description() As String
        Get
            Return _description
        End Get
        Set(ByVal Value As String)
            _description = """" & Value & """"
        End Set
    End Property

    Private _siteUrl As String

    Public Property SiteUrl() As String
        Get
            Return _siteUrl
        End Get
        Set(ByVal Value As String)
            _siteUrl = Value
        End Set
    End Property

    Private _webApplicationUrl As String

    Public Property WebApplicationUrl() As String
        Get
            Return _webApplicationUrl
        End Get
        Set(ByVal Value As String)
            _webApplicationUrl = Value
        End Set
    End Property

    Private _siteTemplate As String

    Public Property SiteTemplate() As String
        Get
            Return _siteTemplate
        End Get
        Set(ByVal Value As String)
            _siteTemplate = Value
        End Set
    End Property

    Private _primaryOwnerLogin As String

    Public Property PrimaryOwnerLogin() As String
        Get
            Return _primaryOwnerLogin
        End Get
        Set(ByVal Value As String)
            _primaryOwnerLogin = Value
        End Set
    End Property

    Private _primaryOwnerName As String

    Public Property PrimaryOwnerName() As String
        Get
            Return _primaryOwnerName
        End Get
        Set(ByVal Value As String)
            _primaryOwnerName = Value
        End Set
    End Property

    Private _primaryOwnerEmail As String

    Public Property PrimaryOwnerEmail() As String
        Get
            Return _primaryOwnerEmail
        End Get
        Set(ByVal Value As String)
            _primaryOwnerEmail = Value
        End Set
    End Property

    Private _secondaryOwnerLogin As String

    Public Property SecondaryOwnerLogin() As String
        Get
            Return _secondaryOwnerLogin
        End Get
        Set(ByVal Value As String)
            _secondaryOwnerLogin = Value
        End Set
    End Property

    Private _secondaryOwnerName As String

    Public Property SecondaryOwnerName() As String
        Get
            Return _secondaryOwnerName
        End Get
        Set(ByVal Value As String)
            _secondaryOwnerName = Value
        End Set
    End Property

    Private _secondaryOwnerEmail As String

    Public Property SecondaryOwnerEmail() As String
        Get
            Return _secondaryOwnerEmail
        End Get
        Set(ByVal Value As String)
            _secondaryOwnerEmail = Value
        End Set
    End Property

    Private _quota As Object

    Public Property Quota() As String
        Get
            Return _quota
        End Get
        Set(ByVal Value As String)
            _quota = Value
        End Set
    End Property

#End Region

    ' Take the data off the excel spreadsheet and populate properties on this object
    ' with values.  This allows the remainder of the code to be shielded from Excel
    ' and data validation to be centralized.

    Public Sub PopulateFromExcelRange(ByRef requestInfo As Excel.Range)
        RequestType = requestInfo.Offset(0, Globals.cSC_RequestType_Offset).Value

        RequestedBy = requestInfo.Offset(0, Globals.cSC_RequestedBy_Offset).Value
        RequestDate = requestInfo.Offset(0, Globals.cSC_RequestDate_Offset).Value
        DateNeeded = requestInfo.Offset(0, Globals.cSC_DateNeeded_Offset).Value
        Purpose = requestInfo.Offset(0, Globals.cSC_Purpose_Offset).Value

        SharePointFarm = requestInfo.Offset(0, Globals.cSC_SharePointFarm_Offset).Value

        WebApplication = requestInfo.Offset(0, Globals.cSC_WebApplication_Offset).Value

        ManagedPathName = requestInfo.Offset(0, Globals.cSC_ManagedPathName_Offset).Value
        ManagedPathType = requestInfo.Offset(0, Globals.cSC_ManagedPathType_Offset).Value

        UrlName = requestInfo.Offset(0, Globals.cSC_UrlName_Offset).Value

        UniquePermissions = requestInfo.Offset(0, Globals.cSC_UniquePermissions_Offset).Value

        Title = requestInfo.Offset(0, Globals.cSC_Title_Offset).Value
        Description = requestInfo.Offset(0, Globals.cSC_Description_Offset).Value

        SiteUrl = requestInfo.Offset(0, Globals.cSC_SiteURL_Offset).Value
        WebApplicationUrl = requestInfo.Offset(0, Globals.cSC_WebApplicationURL_Offset).Value

        SiteTemplate = requestInfo.Offset(0, Globals.cSC_SiteTemplate_Offset).Value

        PrimaryOwnerLogin = requestInfo.Offset(0, Globals.cSC_PrimaryOwnerLogin_Offset).Value
        PrimaryOwnerName = requestInfo.Offset(0, Globals.cSC_PrimaryOwnerName_Offset).Value
        PrimaryOwnerEmail = requestInfo.Offset(0, Globals.cSC_PrimaryOwnerEmail_Offset).Value

        SecondaryOwnerLogin = requestInfo.Offset(0, Globals.cSC_SecondaryOwnerLogin_Offset).Value
        SecondaryOwnerName = requestInfo.Offset(0, Globals.cSC_SecondaryOwnerName_Offset).Value
        SecondaryOwnerEmail = requestInfo.Offset(0, Globals.cSC_SecondaryOwnerEmail_Offset).Value

        Quota = requestInfo.Offset(0, Globals.cSC_Quota_Offset).Value
    End Sub

    ' Generate an Xml element from the properties on this obect.
    ' This is used by the Word Content Controls

    Public Function ToXml() As String
        Dim xmlDoc As New Xml.XmlDocument

        Dim root As Xml.XmlElement = xmlDoc.CreateElement("SiteCollectionCreation")

        xmlDoc.AppendChild(root)

        AppendChild(xmlDoc, root, "RequestedBy", RequestedBy)
        AppendChild(xmlDoc, root, "RequestDate", RequestDate)
        AppendChild(xmlDoc, root, "DateNeeded", DateNeeded)
        AppendChild(xmlDoc, root, "Purpose", Purpose)
        AppendChild(xmlDoc, root, "SharePointFarm", SharePointFarm)
        AppendChild(xmlDoc, root, "WebApplication", WebApplication)
        AppendChild(xmlDoc, root, "Title", Title)
        AppendChild(xmlDoc, root, "Description", Description)
        AppendChild(xmlDoc, root, "URL", SiteUrl)
        AppendChild(xmlDoc, root, "PrimaryAdministrator", PrimaryOwnerLogin)
        AppendChild(xmlDoc, root, "SecondaryAdministrator", SecondaryOwnerLogin)
        AppendChild(xmlDoc, root, "Quota", Quota)

        Return xmlDoc.InnerXml
    End Function

    Private Sub AppendChild(ByVal xmlDoc As XmlDocument, ByVal root As XmlElement, ByVal nodeName As String, ByVal nodeValue As String)
        Dim newElement As XmlElement = xmlDoc.CreateElement(nodeName)
        newElement.InnerText = nodeValue
        root.AppendChild(newElement)
    End Sub
End Class
