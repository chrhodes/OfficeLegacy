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

    Private WithEvents m_cbcbListBox As CommandBarComboBox

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

        Try
            _AppToolbar = CommandBarHelper.AddToolBar(Globals.ThisAddIn.Application.CommandBars, Globals.APPLICATION_TOOLBAR_NAME)

            ' Add new commands to toolbar 

            Dim _Word_WatchWindow As Word_WatchWindow = New Word_WatchWindow(_AppToolbar, MsoButtonStyle.msoButtonIcon)

            ' and any local commands we want available.

            Word_AddLocalCommands()

            ' TODO: Check to see if _AppToolbar has any controls on it.  IF not, hide.

            If _AppToolbar.Controls.Count = 0 Then
                _AppToolbar.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
            PLLog.Error(ex, Globals.PROJECT_NAME)
        End Try

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
            '_commandBarName = "Worksheet Menu Bar"

            '    Case Else
            _commandBarName = "Menu Bar"
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
            Dim _AddFooter As Word_AddFooter = New Word_AddFooter(_CommonMenu.CommandBar, MsoButtonStyle.msoButtonIconAndCaption)
            Dim _Help As Help = New Help(_CommonMenu.CommandBar, MsoButtonStyle.msoButtonIconAndCaption)

            ' Add a submenu 
            Dim _SubMenu As CommandBarPopup = CommandBarHelper.AddSubMenu(_CommonMenu, "&SubMenu")
            ' and some buttons to it
            Dim _AddinInfo As AddinInfo = New AddinInfo(_SubMenu.CommandBar, MsoButtonStyle.msoButtonCaption)
            Dim _EnvironmentInfo As EnvironmentInfo = New EnvironmentInfo(_SubMenu.CommandBar, MsoButtonStyle.msoButtonCaption)

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

            Dim _Action1 As Action1 = New Action1(_CommonToolbar, MsoButtonStyle.msoButtonIcon)
            Dim _Action2 As Word_GetSiteInfo = New Word_GetSiteInfo(_CommonToolbar, MsoButtonStyle.msoButtonIcon)
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
                    ' This throws an exception for some reason on Office 2007
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

    Private Sub Word_AddLocalCommands()
        ' Add commands to our toolbar.  These commands are normally not
        ' on a toolbar.  These numbers can be discovered by dragging commands
        ' onto a toolbar while recording a macro.

        Try
            With _AppToolbar.Controls
                .Add(Type:=MsoControlType.msoControlButton, Id:=60) ' Double Underline
                .Add(Type:=MsoControlType.msoControlButton, Id:=61) ' Word Underline
                .Add(Type:=MsoControlType.msoControlButton, Id:=290) ' Strikethrough
                .Add(Type:=MsoControlType.msoControlButton, Id:=289) ' Drop capital

                .Add(Type:=MsoControlType.msoControlButton, Id:=295) ' Insert cells
                .Add(Type:=MsoControlType.msoControlButton, Id:=292) ' Delete cells
                .Add(Type:=MsoControlType.msoControlButton, Id:=3681) ' Insert rows above
                .Add(Type:=MsoControlType.msoControlButton, Id:=3683) ' Insert rows below
                .Add(Type:=MsoControlType.msoControlButton, Id:=293) ' Delete rows
                'This fails on Office XP (10.0) and is silently ignored on 9.0
                '            .Add Type:=msoControlButton, ID:=297    ' Insert column
                .Add(Type:=MsoControlType.msoControlButton, Id:=3685) ' Insert columns to left
                .Add(Type:=MsoControlType.msoControlButton, Id:=3687) ' Insert columns to right
                .Add(Type:=MsoControlType.msoControlButton, Id:=294) ' Delete columns

                .Add(Type:=MsoControlType.msoControlButton, Id:=779) ' Paragraph dialog
                'TODO: The next two don't have icons!  Take them off and put above
                ' in the custom section, I guess.
                .Add(Type:=MsoControlType.msoControlButton, Id:=768) ' Insert Date/Time
                .Add(Type:=MsoControlType.msoControlButton, Id:=3221) ' Summary Info
                '
                .Add(Type:=MsoControlType.msoControlButton, Id:=3479) ' AutoText
                .Add(Type:=MsoControlType.msoControlButton, Id:=253) ' Format Font
                .Add(Type:=MsoControlType.msoControlButton, Id:=3262) ' Format Object
            End With
        Catch ex As Exception
            PLLog.Error(ex, Globals.PROJECT_NAME)
        End Try

        PLLog.Trace("Exit", Globals.PROJECT_NAME)
    End Sub

#End Region
End Class
