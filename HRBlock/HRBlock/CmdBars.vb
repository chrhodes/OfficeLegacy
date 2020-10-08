Option Strict Off

Imports System.Reflection

Imports AddinHelper
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Core

Imports PacificLife.Life

Public Class CmdBars
    Private Const _MODULE_NAME As String = Globals.PROJECT_NAME & ".CmdBars"

    Private _commandBarName As String   ' App specific name of top level command bar.

    ' Menu 

    Private _CommonMenu As CommandBarPopup
    Private _CommonToolbar As CommandBar

    ' Sub Menu

    'Private _SubMenu As CommandBarPopup

    Private _AppMenu As CommandBarPopup
    Private _AppToolbar As CommandBar

    ' Toolbar items.

    Public Sub CreateCommandBars()
        PLLog.Trace("Enter", Globals.PROJECT_NAME)

        Try
            ' Handle general Menu and Toolbar stuff here.  

            If Globals.HAS_COMMON_MENU Then
                CreateCommonMenu()
            End If

            If Globals.HAS_COMMON_TOOLBAR Then
                CreateCommonToolBar()
            End If

            If Globals.HAS_APPLICATION_MENU Then
                CreateApplicationMenu()
            End If

            If Globals.HAS_APPLICATION_TOOLBAR Then
                CreateApplicationToolbar()
            End If

        Catch ex As Exception
            PLLog.Error(ex, Globals.PROJECT_NAME)
        End Try

        PLLog.Trace("Exit", Globals.PROJECT_NAME)
    End Sub

    Public Sub CreateApplicationMenu()

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub CreateApplicationToolbar()
        PLLog.Trace("Enter", Globals.PROJECT_NAME)

        'Try
        '    _AppToolbar = CommandBarHelper.AddToolBar(Globals.ThisAddIn.Application.CommandBars, Globals.APPLICATION_TOOLBAR_NAME)

        '    ' Add new commands to toolbar

        '    'Dim _Excel_GetSiteInfo As Excel_GetSiteInfo = New Excel_GetSiteInfo(_AppToolbar, MsoButtonStyle.msoButtonIcon)

        '    'Dim _Excel_FolderMaps As Excel_FolderMaps = New Excel_FolderMaps(_AppToolbar, MsoButtonStyle.msoButtonIcon)
        '    'Dim _Excel_GroupDown As Excel_GroupDown = New Excel_GroupDown(_AppToolbar, MsoButtonStyle.msoButtonIcon)
        '    'Dim _Excel_SearchDown As Excel_SearchDown = New Excel_SearchDown(_AppToolbar, MsoButtonStyle.msoButtonIcon)
        '    'Dim _Excel_SearchUp As Excel_SearchUp = New Excel_SearchUp(_AppToolbar, MsoButtonStyle.msoButtonIcon)
        '    'Dim _Excel_UngroupSelection As Excel_UngroupSelection = New Excel_UngroupSelection(_AppToolbar, MsoButtonStyle.msoButtonIcon)
        '    'Dim _Excel_ProtectSheets As Excel_ProtectSheets = New Excel_ProtectSheets(_AppToolbar, MsoButtonStyle.msoButtonIcon)
        '    'Dim _Excel_UnProtectSheets As Excel_UnProtectSheets = New Excel_UnProtectSheets(_AppToolbar, MsoButtonStyle.msoButtonIcon)
        '    'Dim _Excel_AllLandscape As Excel_AllLandscape = New Excel_AllLandscape(_AppToolbar, MsoButtonStyle.msoButtonIcon)
        '    'Dim _Excel_Landscape As Excel_Landscape = New Excel_Landscape(_AppToolbar, MsoButtonStyle.msoButtonIcon)
        '    'Dim _Excel_AllPortrait As Excel_AllPortrait = New Excel_AllPortrait(_AppToolbar, MsoButtonStyle.msoButtonIcon)
        '    'Dim _Excel_Portrait As Excel_Portrait = New Excel_Portrait(_AppToolbar, MsoButtonStyle.msoButtonIcon)
        '    'Dim _Excel_MaxMargins_All As Excel_MaxMargins_All = New Excel_MaxMargins_All(_AppToolbar, MsoButtonStyle.msoButtonIcon)
        '    'Dim _Excel_MaxMargins As Excel_MaxMargins = New Excel_MaxMargins(_AppToolbar, MsoButtonStyle.msoButtonIcon)
        '    'Dim _Excel_TableOfContents As Excel_TableOfContents = New Excel_TableOfContents(_AppToolbar, MsoButtonStyle.msoButtonIcon)
        '    'Dim _Excel_WatchWindow As Excel_WatchWindow = New Excel_WatchWindow(_AppToolbar, MsoButtonStyle.msoButtonIcon)

        '    ' Add any local commands we want available.
        '    AddLocalCommands()

        '    ' Check to see if _AppToolbar has any controls on it.  If not, hide.

        '    If _AppToolbar.Controls.Count = 0 Then
        '        _AppToolbar.Visible = False
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.ToString)
        '    PLLog.Error(ex, Globals.PROJECT_NAME)
        'End Try

        PLLog.Trace("Exit", Globals.PROJECT_NAME)
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub CreateCommonMenu()
        PLLog.Trace("Enter", Globals.PROJECT_NAME)

        Dim commandBar As CommandBar

        Try
            ' Not all applications use the same name for the main CommandBar

            'AddinHelper.CommandBarHelper.DisplayCommandBars(Globals.HostApp, Globals.AppName)

            'Select Case Globals.AppName
            '    Case "Microsoft Excel"
            _commandBarName = "Worksheet Menu Bar"

            '    Case Else
            '_commandBarName = "Menu Bar"
            'End Select

            'Select Case Globals.AppName
            '    Case "Outlook"
            '        commandBar = Globals.HostApp.ActiveExplorer.CommandBars(_commandBarName)

            'Case Else
            commandBar = Globals.ThisAddIn.Application.CommandBars(_commandBarName)
            'End Select

            If Not commandBar Is Nothing Then
                ' ToDo: Handle the error if cannot find expected CommandBar.
            End If

            ' Create a new menu on main menu bar
            _CommonMenu = CommandBarHelper.AddMenu(commandBar, Globals.COMMON_MENU_CAPTION)
            ' and add some buttons
            'Dim _AddFooter As AddFooter = New AddFooter(_CommonMenu.CommandBar, MsoButtonStyle.msoButtonIconAndCaption)
            'Dim _Help As Help = New Help(_CommonMenu.CommandBar, MsoButtonStyle.msoButtonIconAndCaption)

            ' Add a submenu 
            Dim _SubMenu As CommandBarPopup = CommandBarHelper.AddSubMenu(_CommonMenu, "&SubMenu")
            ' and some buttons to it
            'Dim _AddinInfo As AddinInfo = New AddinInfo(_SubMenu.CommandBar, MsoButtonStyle.msoButtonCaption)
            'Dim _EnvironmentInfo As EnvironmentInfo = New EnvironmentInfo(_SubMenu.CommandBar, MsoButtonStyle.msoButtonCaption)

            commandBar = Nothing
        Catch ex As Exception
            MsgBox(ex.ToString)
            PLLog.Error(ex, Globals.PROJECT_NAME)
        End Try

        PLLog.Trace("Exit", Globals.PROJECT_NAME)
    End Sub ' Menu_Create

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks>Any routines here should be common across the apps.</remarks>
    Public Sub CreateCommonToolBar()
        PLLog.Trace("Enter", Globals.PROJECT_NAME)

        Try
            _CommonToolbar = CommandBarHelper.AddToolBar(Globals.ThisAddIn.Application.CommandBars, Globals.COMMON_TOOLBAR_NAME)

            'Dim _Action1 As Action1 = New Action1(_CommonToolbar, MsoButtonStyle.msoButtonIcon)
            'Dim _Action2 As Action2 = New Action2(_CommonToolbar, MsoButtonStyle.msoButtonIcon)
        Catch ex As Exception
            MsgBox(ex.ToString)
            PLLog.Error(ex, Globals.PROJECT_NAME)
        End Try

        PLLog.Trace("Exit", Globals.PROJECT_NAME)
    End Sub


    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub RemoveCommandBars()
        PLLog.Trace("Enter", Globals.PROJECT_NAME)

        If Globals.HAS_COMMON_MENU Then
            If Not _CommonMenu Is Nothing Then
                Try
                    _CommonMenu.Delete()
                Catch ex As Exception

                End Try
            End If
        End If

        If Globals.HAS_COMMON_TOOLBAR Then
            If Not _CommonToolbar Is Nothing Then
                Try
                    _CommonToolbar.Delete()
                Catch ex As Exception

                End Try
            End If

        End If

        If Globals.HAS_APPLICATION_MENU Then
            If Not _AppMenu Is Nothing Then
                Try
                    _AppMenu.Delete()
                Catch ex As Exception

                End Try
            End If

        End If

        If Globals.HAS_APPLICATION_TOOLBAR Then
            If Not _AppToolbar Is Nothing Then
                Try
                    _AppToolbar.Delete()
                Catch ex As Exception

                End Try
            End If
        End If

        PLLog.Trace("Exit", Globals.PROJECT_NAME)
    End Sub

#Region "Local Application Commands"

    Private Sub AddLocalCommands()
        PLLog.Trace("Enter", Globals.PROJECT_NAME)

        ' Add local application commands to our toolbar.  These commands may not normally
        ' be on a toolbar.  These numbers can be discovered by dragging commands
        ' onto a toolbar while recording a macro.

        On Error Resume Next

        'With _AppToolbar.Controls
        '    .Add(Type:=MsoControlType.msoControlButton, Id:=3160)   ' UnGroup
        '    .Add(Type:=MsoControlType.msoControlButton, Id:=3159)   ' Group
        '    .Add(Type:=MsoControlType.msoControlButton, Id:=292)    ' Delete Cells
        '    .Add(Type:=MsoControlType.msoControlButton, Id:=295)    ' Insert Cells
        '    .Add(Type:=MsoControlType.msoControlButton, Id:=294)    ' Delete Columns
        '    .Add(Type:=MsoControlType.msoControlButton, Id:=297)    ' Insert Columns
        '    .Add(Type:=MsoControlType.msoControlButton, Id:=293)    ' Delete Rows
        '    .Add(Type:=MsoControlType.msoControlButton, Id:=296)    ' Insert Rows

        '    ' Don't know how to get these to enable??
        '    '        .Add Type:=msoControlButton, ID:=666           ' Align Top
        '    '        .Add Type:=msoControlButton, ID:=669           ' Align Middle
        '    '        .Add Type:=msoControlButton, ID:=667           ' Align Bottom
        '    '        .Add Type:=msoControlButton, ID:=408           ' Distribute Horizontally
        '    '        .Add Type:=msoControlButton, ID:=465           ' Distribute Verically

        '    .Add(Type:=MsoControlType.msoControlButton, Id:=798)    ' Merge Cells
        '    .Add(Type:=MsoControlType.msoControlButton, Id:=800)    ' UnMerge Cells
        '    .Add(Type:=MsoControlType.msoControlButton, Id:=1742)   ' Merge Across
        '    .Add(Type:=MsoControlType.msoControlButton, Id:=405)    ' Vertical Text
        '    .Add(Type:=MsoControlType.msoControlButton, Id:=406)    ' Rotate Text Up
        '    .Add(Type:=MsoControlType.msoControlButton, Id:=441)    ' Select Visible Cells
        '    .Add(Type:=MsoControlType.msoControlButton, Id:=442)    ' Select Current Region
        '    .Add(Type:=MsoControlType.msoControlButton, Id:=855)    ' Format Cells
        '    .Add(Type:=MsoControlType.msoControlButton, Id:=755)    ' PasteSpecial
        '    .Add(Type:=MsoControlType.msoControlButton, Id:=370)    ' Paste Values
        '    .Add(Type:=MsoControlType.msoControlButton, Id:=369)    ' Paste Formatting
        '    .Add(Type:=MsoControlType.msoControlButton, Id:=47)     ' Clear Contents
        '    .Add(Type:=MsoControlType.msoControlButton, Id:=368)    ' Clear Formatting
        '    .Add(Type:=MsoControlType.msoControlButton, Id:=893)    ' Protect WorkSheet
        '    .Add(Type:=MsoControlType.msoControlButton, Id:=894)    ' Protect WorkBook
        'End With

        PLLog.Trace("Exit", Globals.PROJECT_NAME)
    End Sub
#End Region
End Class
