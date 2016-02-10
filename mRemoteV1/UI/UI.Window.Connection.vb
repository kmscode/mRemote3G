Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports Crownwood.Magic.Controls
Imports mRemote3G.App
Imports mRemote3G.Config
Imports mRemote3G.Connection
Imports mRemote3G.Connection.Protocol
Imports mRemote3G.Forms
Imports mRemote3G.Messages
Imports mRemote3G.My
Imports mRemote3G.Tools
Imports PSTaskDialog
Imports WeifenLuo.WinFormsUI.Docking

Namespace UI

    Namespace Window
        Public Class Connection
            Inherits Base

#Region "Form Init"

            Friend WithEvents cmenTab As ContextMenuStrip
            Private components As IContainer
            Friend WithEvents cmenTabFullscreen As ToolStripMenuItem
            Friend WithEvents cmenTabScreenshot As ToolStripMenuItem
            Friend WithEvents cmenTabTransferFile As ToolStripMenuItem
            Friend WithEvents cmenTabSendSpecialKeys As ToolStripMenuItem
            Friend WithEvents cmenTabSep1 As ToolStripSeparator
            Friend WithEvents cmenTabRenameTab As ToolStripMenuItem
            Friend WithEvents cmenTabDuplicateTab As ToolStripMenuItem
            Friend WithEvents cmenTabDisconnect As ToolStripMenuItem
            Friend WithEvents cmenTabSmartSize As ToolStripMenuItem
            Friend WithEvents cmenTabSendSpecialKeysCtrlAltDel As ToolStripMenuItem
            Friend WithEvents cmenTabSendSpecialKeysCtrlEsc As ToolStripMenuItem
            Friend WithEvents cmenTabViewOnly As ToolStripMenuItem
            Friend WithEvents cmenTabReconnect As ToolStripMenuItem
            Friend WithEvents cmenTabExternalApps As ToolStripMenuItem
            Friend WithEvents cmenTabStartChat As ToolStripMenuItem
            Friend WithEvents cmenTabRefreshScreen As ToolStripMenuItem
            Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
            Friend WithEvents cmenTabPuttySettings As ToolStripMenuItem

            Public WithEvents TabController As TabControl

            Private Sub InitializeComponent()
                Me.components = New System.ComponentModel.Container()
                Dim resources = New ComponentResourceManager(GetType(Connection))
                Me.TabController = New TabControl()
                Me.cmenTab = New ContextMenuStrip(Me.components)
                Me.cmenTabFullscreen = New ToolStripMenuItem()
                Me.cmenTabSmartSize = New ToolStripMenuItem()
                Me.cmenTabViewOnly = New ToolStripMenuItem()
                Me.ToolStripSeparator1 = New ToolStripSeparator()
                Me.cmenTabScreenshot = New ToolStripMenuItem()
                Me.cmenTabStartChat = New ToolStripMenuItem()
                Me.cmenTabTransferFile = New ToolStripMenuItem()
                Me.cmenTabRefreshScreen = New ToolStripMenuItem()
                Me.cmenTabSendSpecialKeys = New ToolStripMenuItem()
                Me.cmenTabSendSpecialKeysCtrlAltDel = New ToolStripMenuItem()
                Me.cmenTabSendSpecialKeysCtrlEsc = New ToolStripMenuItem()
                Me.cmenTabPuttySettings = New ToolStripMenuItem()
                Me.cmenTabExternalApps = New ToolStripMenuItem()
                Me.cmenTabSep1 = New ToolStripSeparator()
                Me.cmenTabRenameTab = New ToolStripMenuItem()
                Me.cmenTabDuplicateTab = New ToolStripMenuItem()
                Me.cmenTabReconnect = New ToolStripMenuItem()
                Me.cmenTabDisconnect = New ToolStripMenuItem()
                Me.cmenTab.SuspendLayout()
                Me.SuspendLayout()
                '
                'TabController
                '
                Me.TabController.Anchor = CType((((AnchorStyles.Top Or AnchorStyles.Bottom) _
                                                  Or AnchorStyles.Left) _
                                                 Or AnchorStyles.Right),
                                                AnchorStyles)
                Me.TabController.Appearance = TabControl.VisualAppearance.MultiDocument
                Me.TabController.Cursor = Cursors.Hand
                Me.TabController.DragFromControl = False
                Me.TabController.IDEPixelArea = False
                Me.TabController.Location = New Point(0, -1)
                Me.TabController.Name = "TabController"
                Me.TabController.Size = New Size(632, 454)
                Me.TabController.TabIndex = 0
                '
                'cmenTab
                '
                Me.cmenTab.ImageScalingSize = New Size(20, 20)
                Me.cmenTab.Items.AddRange(New ToolStripItem() _
                                             {Me.cmenTabFullscreen, Me.cmenTabSmartSize, Me.cmenTabViewOnly,
                                              Me.ToolStripSeparator1, Me.cmenTabScreenshot, Me.cmenTabStartChat,
                                              Me.cmenTabTransferFile, Me.cmenTabRefreshScreen, Me.cmenTabSendSpecialKeys,
                                              Me.cmenTabPuttySettings, Me.cmenTabExternalApps, Me.cmenTabSep1,
                                              Me.cmenTabRenameTab, Me.cmenTabDuplicateTab, Me.cmenTabReconnect,
                                              Me.cmenTabDisconnect})
                Me.cmenTab.Name = "cmenTab"
                Me.cmenTab.RenderMode = ToolStripRenderMode.Professional
                Me.cmenTab.Size = New Size(245, 380)
                '
                'cmenTabFullscreen
                '
                Me.cmenTabFullscreen.Image = My.Resources.Resources.arrow_out
                Me.cmenTabFullscreen.Name = "cmenTabFullscreen"
                Me.cmenTabFullscreen.Size = New Size(244, 26)
                Me.cmenTabFullscreen.Text = "Fullscreen (RDP)"
                '
                'cmenTabSmartSize
                '
                Me.cmenTabSmartSize.Image = My.Resources.Resources.SmartSize
                Me.cmenTabSmartSize.Name = "cmenTabSmartSize"
                Me.cmenTabSmartSize.Size = New Size(244, 26)
                Me.cmenTabSmartSize.Text = "SmartSize (RDP/VNC)"
                '
                'cmenTabViewOnly
                '
                Me.cmenTabViewOnly.Name = "cmenTabViewOnly"
                Me.cmenTabViewOnly.Size = New Size(244, 26)
                Me.cmenTabViewOnly.Text = "View Only (VNC)"
                '
                'ToolStripSeparator1
                '
                Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
                Me.ToolStripSeparator1.Size = New Size(241, 6)
                '
                'cmenTabScreenshot
                '
                Me.cmenTabScreenshot.Image = My.Resources.Resources.Screenshot_Add
                Me.cmenTabScreenshot.Name = "cmenTabScreenshot"
                Me.cmenTabScreenshot.Size = New Size(244, 26)
                Me.cmenTabScreenshot.Text = "Screenshot"
                '
                'cmenTabStartChat
                '
                Me.cmenTabStartChat.Image = My.Resources.Resources.Chat
                Me.cmenTabStartChat.Name = "cmenTabStartChat"
                Me.cmenTabStartChat.Size = New Size(244, 26)
                Me.cmenTabStartChat.Text = "Start Chat (VNC)"
                Me.cmenTabStartChat.Visible = False
                '
                'cmenTabTransferFile
                '
                Me.cmenTabTransferFile.Image = My.Resources.Resources.SSHTransfer
                Me.cmenTabTransferFile.Name = "cmenTabTransferFile"
                Me.cmenTabTransferFile.Size = New Size(244, 26)
                Me.cmenTabTransferFile.Text = "Transfer File (SSH)"
                '
                'cmenTabRefreshScreen
                '
                Me.cmenTabRefreshScreen.Image = My.Resources.Resources.Refresh
                Me.cmenTabRefreshScreen.Name = "cmenTabRefreshScreen"
                Me.cmenTabRefreshScreen.Size = New Size(244, 26)
                Me.cmenTabRefreshScreen.Text = "Refresh Screen (VNC)"
                '
                'cmenTabSendSpecialKeys
                '
                Me.cmenTabSendSpecialKeys.DropDownItems.AddRange(New ToolStripItem() _
                                                                    {Me.cmenTabSendSpecialKeysCtrlAltDel,
                                                                     Me.cmenTabSendSpecialKeysCtrlEsc})
                Me.cmenTabSendSpecialKeys.Image = My.Resources.Resources.Keyboard
                Me.cmenTabSendSpecialKeys.Name = "cmenTabSendSpecialKeys"
                Me.cmenTabSendSpecialKeys.Size = New Size(244, 26)
                Me.cmenTabSendSpecialKeys.Text = "Send special Keys (VNC)"
                '
                'cmenTabSendSpecialKeysCtrlAltDel
                '
                Me.cmenTabSendSpecialKeysCtrlAltDel.Name = "cmenTabSendSpecialKeysCtrlAltDel"
                Me.cmenTabSendSpecialKeysCtrlAltDel.Size = New Size(169, 26)
                Me.cmenTabSendSpecialKeysCtrlAltDel.Text = "Ctrl+Alt+Del"
                '
                'cmenTabSendSpecialKeysCtrlEsc
                '
                Me.cmenTabSendSpecialKeysCtrlEsc.Name = "cmenTabSendSpecialKeysCtrlEsc"
                Me.cmenTabSendSpecialKeysCtrlEsc.Size = New Size(169, 26)
                Me.cmenTabSendSpecialKeysCtrlEsc.Text = "Ctrl+Esc"
                '
                'cmenTabPuttySettings
                '
                Me.cmenTabPuttySettings.Name = "cmenTabPuttySettings"
                Me.cmenTabPuttySettings.Size = New Size(244, 26)
                Me.cmenTabPuttySettings.Text = "PuTTY Settings"
                '
                'cmenTabExternalApps
                '
                Me.cmenTabExternalApps.Image = CType(resources.GetObject("cmenTabExternalApps.Image"), Image)
                Me.cmenTabExternalApps.Name = "cmenTabExternalApps"
                Me.cmenTabExternalApps.Size = New Size(244, 26)
                Me.cmenTabExternalApps.Text = "External Applications"
                '
                'cmenTabSep1
                '
                Me.cmenTabSep1.Name = "cmenTabSep1"
                Me.cmenTabSep1.Size = New Size(241, 6)
                '
                'cmenTabRenameTab
                '
                Me.cmenTabRenameTab.Image = My.Resources.Resources.Rename
                Me.cmenTabRenameTab.Name = "cmenTabRenameTab"
                Me.cmenTabRenameTab.Size = New Size(244, 26)
                Me.cmenTabRenameTab.Text = "Rename Tab"
                '
                'cmenTabDuplicateTab
                '
                Me.cmenTabDuplicateTab.Name = "cmenTabDuplicateTab"
                Me.cmenTabDuplicateTab.Size = New Size(244, 26)
                Me.cmenTabDuplicateTab.Text = "Duplicate Tab"
                '
                'cmenTabReconnect
                '
                Me.cmenTabReconnect.Image = CType(resources.GetObject("cmenTabReconnect.Image"), Image)
                Me.cmenTabReconnect.Name = "cmenTabReconnect"
                Me.cmenTabReconnect.Size = New Size(244, 26)
                Me.cmenTabReconnect.Text = "Reconnect"
                '
                'cmenTabDisconnect
                '
                Me.cmenTabDisconnect.Image = My.Resources.Resources.Pause1
                Me.cmenTabDisconnect.Name = "cmenTabDisconnect"
                Me.cmenTabDisconnect.Size = New Size(244, 26)
                Me.cmenTabDisconnect.Text = "Disconnect"
                '
                'Connection
                '
                Me.ClientSize = New Size(632, 453)
                Me.Controls.Add(Me.TabController)
                Me.Font = New Font("Microsoft Sans Serif", 8.25!, FontStyle.Regular, GraphicsUnit.Point, CType(0, Byte))
                Me.Icon = My.Resources.Resources.mRemote_Icon
                Me.Name = "Connection"
                Me.TabText = "UI.Window.Connection"
                Me.Text = "UI.Window.Connection"
                Me.cmenTab.ResumeLayout(False)
                Me.ResumeLayout(False)
            End Sub

#End Region

#Region "Public Methods"

            Public Sub New(Panel As DockContent, Optional ByVal FormText As String = "")

                If FormText = "" Then
                    FormText = Language.Language.strNewPanel
                End If

                Me.WindowType = Type.Connection
                Me.DockPnl = Panel
                Me.InitializeComponent()
                Me.Text = FormText
                Me.TabText = FormText
            End Sub

            Public Function AddConnectionTab(conI As Info) As TabPage
                Try
                    Dim nTab As New TabPage
                    nTab.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Top

                    If MySettingsProperty.Settings.ShowProtocolOnTabs Then
                        nTab.Title = conI.Protocol.ToString & ": "
                    Else
                        nTab.Title = ""
                    End If

                    nTab.Title &= conI.Name

                    If MySettingsProperty.Settings.ShowLogonInfoOnTabs Then
                        nTab.Title &= " ("

                        If conI.Domain <> "" Then
                            nTab.Title &= conI.Domain
                        End If

                        If conI.Username <> "" Then
                            If conI.Domain <> "" Then
                                nTab.Title &= "\"
                            End If

                            nTab.Title &= conI.Username
                        End If

                        nTab.Title &= ")"
                    End If

                    nTab.Title = nTab.Title.Replace("&", "&&")

                    Dim conIcon As Drawing.Icon = Global.mRemote3G.Connection.Icon.FromString(conI.Icon)
                    If conIcon IsNot Nothing Then
                        nTab.Icon = conIcon
                    End If

                    If MySettingsProperty.Settings.OpenTabsRightOfSelected Then
                        Me.TabController.TabPages.Insert(Me.TabController.SelectedIndex + 1, nTab)
                    Else
                        Me.TabController.TabPages.Add(nTab)
                    End If

                    nTab.Selected = True
                    _ignoreChangeSelectedTabClick = False

                    Return nTab
                Catch ex As Exception
                    App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "AddConnectionTab (UI.Window.Connections) failed" & vbNewLine & ex.ToString(), True)
                End Try

                Return Nothing
            End Function

#End Region

#Region "Private Methods"

            Public Sub UpdateSelectedConnection()
                If TabController.SelectedTab Is Nothing Then
                    frmMain.SelectedConnection = Nothing
                Else
                    Dim interfaceControl = TryCast(TabController.SelectedTab.Tag, InterfaceControl)
                    If interfaceControl Is Nothing Then
                        frmMain.SelectedConnection = Nothing
                    Else
                        frmMain.SelectedConnection = interfaceControl.Info
                    End If
                End If
            End Sub

#End Region

#Region "Form"

            Private Sub Connection_Load(sender As Object, e As EventArgs) Handles Me.Load
                ApplyLanguage()
            End Sub

            Private _documentHandlersAdded As Boolean = False
            Private _floatHandlersAdded As Boolean = False

            Private Sub Connection_DockStateChanged(sender As Object, e As EventArgs) Handles Me.DockStateChanged
                If DockState = DockState.Float Then
                    If _documentHandlersAdded Then
                        RemoveHandler frmMain.ResizeBegin, AddressOf Connection_ResizeBegin
                        RemoveHandler frmMain.ResizeEnd, AddressOf Connection_ResizeEnd
                        _documentHandlersAdded = False
                    End If
                    AddHandler DockHandler.FloatPane.FloatWindow.ResizeBegin, AddressOf Connection_ResizeBegin
                    AddHandler DockHandler.FloatPane.FloatWindow.ResizeEnd, AddressOf Connection_ResizeEnd
                    _floatHandlersAdded = True
                ElseIf DockState = DockState.Document Then
                    If _floatHandlersAdded Then
                        RemoveHandler DockHandler.FloatPane.FloatWindow.ResizeBegin, AddressOf Connection_ResizeBegin
                        RemoveHandler DockHandler.FloatPane.FloatWindow.ResizeEnd, AddressOf Connection_ResizeEnd
                        _floatHandlersAdded = False
                    End If
                    AddHandler frmMain.ResizeBegin, AddressOf Connection_ResizeBegin
                    AddHandler frmMain.ResizeEnd, AddressOf Connection_ResizeEnd
                    _documentHandlersAdded = True
                End If
            End Sub

            Private Sub ApplyLanguage()
                cmenTabFullscreen.Text = Language.Language.strMenuFullScreenRDP
                cmenTabSmartSize.Text = Language.Language.strMenuSmartSize
                cmenTabViewOnly.Text = Language.Language.strMenuViewOnly
                cmenTabScreenshot.Text = Language.Language.strMenuScreenshot
                cmenTabStartChat.Text = Language.Language.strMenuStartChat
                cmenTabTransferFile.Text = Language.Language.strMenuTransferFile
                cmenTabRefreshScreen.Text = Language.Language.strMenuRefreshScreen
                cmenTabSendSpecialKeys.Text = Language.Language.strMenuSendSpecialKeys
                cmenTabSendSpecialKeysCtrlAltDel.Text = Language.Language.strMenuCtrlAltDel
                cmenTabSendSpecialKeysCtrlEsc.Text = Language.Language.strMenuCtrlEsc
                cmenTabExternalApps.Text = Language.Language.strMenuExternalTools
                cmenTabRenameTab.Text = Language.Language.strMenuRenameTab
                cmenTabDuplicateTab.Text = Language.Language.strMenuDuplicateTab
                cmenTabReconnect.Text = Language.Language.strMenuReconnect
                cmenTabDisconnect.Text = Language.Language.strMenuDisconnect
                cmenTabPuttySettings.Text = Language.Language.strPuttySettings
            End Sub

            Private Sub Connection_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
                If Not frmMain.IsClosing And (
                        (MySettingsProperty.Settings.ConfirmCloseConnection = ConfirmClose.All And TabController.TabPages.Count > 0) Or
                        (MySettingsProperty.Settings.ConfirmCloseConnection = ConfirmClose.Multiple And TabController.TabPages.Count > 1)) Then
                    Dim result As DialogResult = cTaskDialog.MessageBox(Me, My.Application.Info.ProductName, String.Format(Language.Language.strConfirmCloseConnectionPanelMainInstruction, Me.Text), "", "", "", Language.Language.strCheckboxDoNotShowThisMessageAgain, eTaskDialogButtons.YesNo, eSysIcons.Question, Nothing)
                    If cTaskDialog.VerificationChecked Then
                        MySettingsProperty.Settings.ConfirmCloseConnection =
                            MySettingsProperty.Settings.ConfirmCloseConnection - 1
                    End If
                    If result = DialogResult.No Then
                        e.Cancel = True
                        Exit Sub
                    End If
                End If

                Try
                    For Each tabP As TabPage In Me.TabController.TabPages
                        If tabP.Tag IsNot Nothing Then
                            Dim interfaceControl As InterfaceControl = tabP.Tag
                            interfaceControl.Protocol.Close()
                        End If
                    Next
                Catch ex As Exception
                    App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "UI.Window.Connection.Connection_FormClosing() failed" & vbNewLine & ex.ToString(), True)
                End Try
            End Sub

            Public Shadows Event ResizeBegin As EventHandler

            Private Sub Connection_ResizeBegin(sender As Object, e As EventArgs)
                RaiseEvent ResizeBegin(Me, e)
            End Sub

            Public Shadows Event ResizeEnd As EventHandler

            Public Sub Connection_ResizeEnd(sender As Object, e As EventArgs)
                RaiseEvent ResizeEnd(sender, e)
            End Sub

#End Region

#Region "TabController"

            Private Sub TabController_ClosePressed(sender As Object, e As EventArgs) Handles TabController.ClosePressed
                If Me.TabController.SelectedTab Is Nothing Then
                    Exit Sub
                End If

                Me.CloseConnectionTab()
            End Sub

            Private Sub CloseConnectionTab()
                Dim selectedTab As TabPage = TabController.SelectedTab
                If MySettingsProperty.Settings.ConfirmCloseConnection = ConfirmClose.All Then
                    Dim result As DialogResult = cTaskDialog.MessageBox(Me, My.Application.Info.ProductName, String.Format(Language.Language.strConfirmCloseConnectionMainInstruction, selectedTab.Title), "", "", "", Language.Language.strCheckboxDoNotShowThisMessageAgain, eTaskDialogButtons.YesNo, eSysIcons.Question, Nothing)
                    If cTaskDialog.VerificationChecked Then
                        MySettingsProperty.Settings.ConfirmCloseConnection =
                            MySettingsProperty.Settings.ConfirmCloseConnection - 1
                    End If
                    If result = DialogResult.No Then
                        Exit Sub
                    End If
                End If

                Try
                    If selectedTab.Tag IsNot Nothing Then
                        Dim interfaceControl As InterfaceControl = selectedTab.Tag
                        interfaceControl.Protocol.Close()
                    Else
                        CloseTab(selectedTab)
                    End If
                Catch ex As Exception
                    App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "UI.Window.Connection.CloseConnectionTab() failed" & vbNewLine & ex.ToString(), True)
                End Try

                UpdateSelectedConnection()
            End Sub

            Private Sub TabController_DoubleClickTab(sender As TabControl, page As TabPage) _
                Handles TabController.DoubleClickTab
                _firstClickTicks = 0
                If MySettingsProperty.Settings.DoubleClickOnTabClosesIt Then
                    Me.CloseConnectionTab()
                End If
            End Sub

#Region "Drag and Drop"

            Private Sub TabController_DragDrop(sender As Object, e As DragEventArgs) Handles TabController.DragDrop
                If e.Data.GetDataPresent("System.Windows.Forms.TreeNode", True) Then
                    Runtime.OpenConnection(e.Data.GetData("System.Windows.Forms.TreeNode", True).Tag, Me,
                                           Info.Force.DoNotJump)
                End If
            End Sub

            Private Sub TabController_DragEnter(sender As Object, e As DragEventArgs) Handles TabController.DragEnter
                If e.Data.GetDataPresent("System.Windows.Forms.TreeNode", True) Then
                    e.Effect = DragDropEffects.Move
                End If
            End Sub

            Private Sub TabController_DragOver(sender As Object, e As DragEventArgs) Handles TabController.DragOver
                e.Effect = DragDropEffects.Move
            End Sub

#End Region

#End Region

#Region "Tab Menu"

            Private Sub ShowHideMenuButtons()
                Try
                    If Me.TabController.SelectedTab Is Nothing Then
                        Exit Sub
                    End If

                    Dim IC As InterfaceControl = Me.TabController.SelectedTab.Tag

                    If IC Is Nothing Then
                        Exit Sub
                    End If

                    If IC.Info.Protocol = Protocols.RDP Then
                        Dim rdp As RDP = IC.Protocol

                        cmenTabFullscreen.Visible = True
                        cmenTabFullscreen.Checked = rdp.Fullscreen

                        cmenTabSmartSize.Visible = True
                        cmenTabSmartSize.Checked = rdp.SmartSize

                        ToolStripSeparator1.Visible = True
                    Else
                        cmenTabFullscreen.Visible = False
                        cmenTabSmartSize.Visible = False
                        ToolStripSeparator1.Visible = False
                    End If

                    If IC.Info.Protocol = Protocols.VNC Then
                        Me.cmenTabSendSpecialKeys.Visible = True
                        Me.cmenTabViewOnly.Visible = True

                        Me.cmenTabSmartSize.Visible = True
                        Me.cmenTabStartChat.Visible = True
                        Me.cmenTabRefreshScreen.Visible = True
                        Me.cmenTabTransferFile.Visible = False

                        Dim vnc As VNC = IC.Protocol
                        Me.cmenTabSmartSize.Checked = vnc.SmartSize
                        Me.cmenTabViewOnly.Checked = vnc.ViewOnly
                    Else
                        Me.cmenTabSendSpecialKeys.Visible = False
                        Me.cmenTabViewOnly.Visible = False
                        Me.cmenTabStartChat.Visible = False
                        Me.cmenTabRefreshScreen.Visible = False
                        Me.cmenTabTransferFile.Visible = False
                    End If

                    If IC.Info.Protocol = Protocols.SSH1 Or IC.Info.Protocol = Protocols.SSH2 Then
                        Me.cmenTabTransferFile.Visible = True
                    End If

                    If TypeOf IC.Protocol Is PuttyBase Then
                        Me.cmenTabPuttySettings.Visible = True
                    Else
                        Me.cmenTabPuttySettings.Visible = False
                    End If

                    AddExternalApps()
                Catch ex As Exception
                    App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "ShowHideMenuButtons (UI.Window.Connections) failed" & vbNewLine & ex.ToString(), True)
                End Try
            End Sub

            Private Sub cmenTabScreenshot_Click(sender As Object, e As EventArgs) Handles cmenTabScreenshot.Click
                cmenTab.Close()
                Application.DoEvents()
                Runtime.Windows.screenshotForm.AddScreenshot(Misc.TakeScreenshot(Me))
            End Sub

            Private Sub cmenTabSmartSize_Click(sender As Object, e As EventArgs) Handles cmenTabSmartSize.Click
                Me.ToggleSmartSize()
            End Sub

            Private Sub cmenTabReconnect_Click(sender As Object, e As EventArgs) Handles cmenTabReconnect.Click
                Me.Reconnect()
            End Sub

            Private Sub cmenTabTransferFile_Click(sender As Object, e As EventArgs) Handles cmenTabTransferFile.Click
                Me.TransferFile()
            End Sub

            Private Sub cmenTabViewOnly_Click(sender As Object, e As EventArgs) Handles cmenTabViewOnly.Click
                Me.ToggleViewOnly()
            End Sub

            Private Sub cmenTabStartChat_Click(sender As Object, e As EventArgs) Handles cmenTabStartChat.Click
                Me.StartChat()
            End Sub

            Private Sub cmenTabRefreshScreen_Click(sender As Object, e As EventArgs) Handles cmenTabRefreshScreen.Click
                Me.RefreshScreen()
            End Sub

            Private Sub cmenTabSendSpecialKeysCtrlAltDel_Click(sender As Object, e As EventArgs) _
                Handles cmenTabSendSpecialKeysCtrlAltDel.Click
                Me.SendSpecialKeys(VNC.SpecialKeys.CtrlAltDel)
            End Sub

            Private Sub cmenTabSendSpecialKeysCtrlEsc_Click(sender As Object, e As EventArgs) _
                Handles cmenTabSendSpecialKeysCtrlEsc.Click
                Me.SendSpecialKeys(VNC.SpecialKeys.CtrlEsc)
            End Sub

            Private Sub cmenTabFullscreen_Click(sender As Object, e As EventArgs) Handles cmenTabFullscreen.Click
                Me.ToggleFullscreen()
            End Sub

            Private Sub cmenTabPuttySettings_Click(sender As Object, e As EventArgs) Handles cmenTabPuttySettings.Click
                Me.ShowPuttySettingsDialog()
            End Sub

            Private Sub cmenTabExternalAppsEntry_Click(sender As Object, e As EventArgs)
                StartExternalApp(sender.Tag)
            End Sub

            Private Sub cmenTabDisconnect_Click(sender As Object, e As EventArgs) Handles cmenTabDisconnect.Click
                Me.CloseTabMenu()
            End Sub

            Private Sub cmenTabDuplicateTab_Click(sender As Object, e As EventArgs) Handles cmenTabDuplicateTab.Click
                Me.DuplicateTab()
            End Sub

            Private Sub cmenTabRenameTab_Click(sender As Object, e As EventArgs) Handles cmenTabRenameTab.Click
                Me.RenameTab()
            End Sub

#End Region

#Region "Tab Actions"

            Private Sub ToggleSmartSize()
                Try
                    If Me.TabController.SelectedTab IsNot Nothing Then
                        If TypeOf Me.TabController.SelectedTab.Tag Is InterfaceControl Then
                            Dim IC As InterfaceControl = Me.TabController.SelectedTab.Tag

                            If TypeOf IC.Protocol Is RDP Then
                                Dim rdp As RDP = IC.Protocol
                                rdp.ToggleSmartSize()
                            ElseIf TypeOf IC.Protocol Is VNC Then
                                Dim vnc As VNC = IC.Protocol
                                vnc.ToggleSmartSize()
                            End If
                        End If
                    End If
                Catch ex As Exception
                    App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "ToggleSmartSize (UI.Window.Connections) failed" & vbNewLine & ex.ToString(), True)
                End Try
            End Sub

            Private Sub TransferFile()
                Try
                    If Me.TabController.SelectedTab IsNot Nothing Then
                        If TypeOf Me.TabController.SelectedTab.Tag Is InterfaceControl Then
                            Dim IC As InterfaceControl = Me.TabController.SelectedTab.Tag

                            If IC.Info.Protocol = Protocols.SSH1 Or IC.Info.Protocol = Protocols.SSH2 Then
                                SSHTransferFile()
                            ElseIf IC.Info.Protocol = Protocols.VNC Then
                                VNCTransferFile()
                            End If
                        End If
                    End If
                Catch ex As Exception
                    App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "TransferFile (UI.Window.Connections) failed" & vbNewLine & ex.ToString(), True)
                End Try
            End Sub

            Private Sub SSHTransferFile()
                Try

                    Dim IC As InterfaceControl = Me.TabController.SelectedTab.Tag

                    Runtime.Windows.Show(Type.SSHTransfer)

                    Dim conI As Info = IC.Info

                    Runtime.Windows.sshtransferForm.Hostname = conI.Hostname
                    Runtime.Windows.sshtransferForm.Username = conI.Username
                    Runtime.Windows.sshtransferForm.Password = conI.Password
                    Runtime.Windows.sshtransferForm.Port = conI.Port
                Catch ex As Exception
                    App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "SSHTransferFile (UI.Window.Connections) failed" & vbNewLine & ex.ToString(), True)
                End Try
            End Sub

            Private Sub VNCTransferFile()
                Try
                    Dim IC As InterfaceControl = Me.TabController.SelectedTab.Tag
                    Dim vnc As VNC = IC.Protocol
                    vnc.StartFileTransfer()
                Catch ex As Exception
                    App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "VNCTransferFile (UI.Window.Connections) failed" & vbNewLine & ex.ToString(), True)
                End Try
            End Sub

            Private Sub ToggleViewOnly()
                Try
                    If Me.TabController.SelectedTab IsNot Nothing Then
                        If TypeOf Me.TabController.SelectedTab.Tag Is InterfaceControl Then
                            Dim IC As InterfaceControl = Me.TabController.SelectedTab.Tag

                            If TypeOf IC.Protocol Is VNC Then
                                cmenTabViewOnly.Checked = Not cmenTabViewOnly.Checked

                                Dim vnc As VNC = IC.Protocol
                                vnc.ToggleViewOnly()
                            End If
                        End If
                    End If
                Catch ex As Exception
                    App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "ToggleViewOnly (UI.Window.Connections) failed" & vbNewLine & ex.ToString(), True)
                End Try
            End Sub

            Private Sub StartChat()
                Try
                    If Me.TabController.SelectedTab IsNot Nothing Then
                        If TypeOf Me.TabController.SelectedTab.Tag Is InterfaceControl Then
                            Dim IC As InterfaceControl = Me.TabController.SelectedTab.Tag

                            If TypeOf IC.Protocol Is VNC Then
                                Dim vnc As VNC = IC.Protocol
                                vnc.StartChat()
                            End If
                        End If
                    End If
                Catch ex As Exception
                    App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "StartChat (UI.Window.Connections) failed" & vbNewLine & ex.ToString(), True)
                End Try
            End Sub

            Private Sub RefreshScreen()
                Try
                    If Me.TabController.SelectedTab IsNot Nothing Then
                        If TypeOf Me.TabController.SelectedTab.Tag Is InterfaceControl Then
                            Dim IC As InterfaceControl = Me.TabController.SelectedTab.Tag

                            If TypeOf IC.Protocol Is VNC Then
                                Dim vnc As VNC = IC.Protocol
                                vnc.RefreshScreen()
                            End If
                        End If
                    End If
                Catch ex As Exception
                    App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "RefreshScreen (UI.Window.Connections) failed" & vbNewLine & ex.ToString(), True)
                End Try
            End Sub

            Private Sub SendSpecialKeys(Keys As VNC.SpecialKeys)
                Try
                    If Me.TabController.SelectedTab IsNot Nothing Then
                        If TypeOf Me.TabController.SelectedTab.Tag Is InterfaceControl Then
                            Dim IC As InterfaceControl = Me.TabController.SelectedTab.Tag

                            If TypeOf IC.Protocol Is VNC Then
                                Dim vnc As VNC = IC.Protocol
                                vnc.SendSpecialKeys(Keys)
                            End If
                        End If
                    End If
                Catch ex As Exception
                    App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "SendSpecialKeys (UI.Window.Connections) failed" & vbNewLine & ex.ToString(), True)
                End Try
            End Sub

            Private Sub ToggleFullscreen()
                Try
                    If Me.TabController.SelectedTab IsNot Nothing Then
                        If TypeOf Me.TabController.SelectedTab.Tag Is InterfaceControl Then
                            Dim IC As InterfaceControl = Me.TabController.SelectedTab.Tag

                            If TypeOf IC.Protocol Is RDP Then
                                Dim rdp As RDP = IC.Protocol
                                rdp.ToggleFullscreen()
                            End If
                        End If
                    End If
                Catch ex As Exception
                    App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "ToggleFullscreen (UI.Window.Connections) failed" & vbNewLine & ex.ToString(), True)
                End Try
            End Sub

            Private Sub ShowPuttySettingsDialog()
                Try
                    If Me.TabController.SelectedTab IsNot Nothing Then
                        If TypeOf Me.TabController.SelectedTab.Tag Is InterfaceControl Then
                            Dim objInterfaceControl As InterfaceControl = Me.TabController.SelectedTab.Tag

                            If TypeOf objInterfaceControl.Protocol Is PuttyBase Then
                                Dim objPuttyBase As PuttyBase = objInterfaceControl.Protocol

                                objPuttyBase.ShowSettingsDialog()
                            End If
                        End If
                    End If
                Catch ex As Exception
                    App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "ShowPuttySettingsDialog (UI.Window.Connections) failed" & vbNewLine & ex.ToString(), True)
                End Try
            End Sub

            Private Sub AddExternalApps()
                Try
                    'clean up
                    'since new items are added below, we have to dispose of any previous items first
                    If cmenTabExternalApps.DropDownItems.Count > 0 Then
                        For Each mitem As ToolStripMenuItem In cmenTabExternalApps.DropDownItems
                            mitem.Dispose()
                        Next mitem
                        cmenTabExternalApps.DropDownItems.Clear()
                    End If


                    'add ext apps
                    For Each extA As ExternalTool In Runtime.ExternalTools
                        Dim nItem As New ToolStripMenuItem
                        nItem.Text = extA.DisplayName
                        nItem.Tag = extA

                        nItem.Image = extA.Image

                        AddHandler nItem.Click, AddressOf cmenTabExternalAppsEntry_Click

                        cmenTabExternalApps.DropDownItems.Add(nItem)
                    Next
                Catch ex As Exception
                    App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "cMenTreeTools_DropDownOpening failed (UI.Window.Tree)" & vbNewLine & ex.ToString(), True)
                End Try
            End Sub

            Private Sub StartExternalApp(ExtA As ExternalTool)
                Try
                    If Me.TabController.SelectedTab IsNot Nothing Then
                        If TypeOf Me.TabController.SelectedTab.Tag Is InterfaceControl Then
                            Dim IC As InterfaceControl = Me.TabController.SelectedTab.Tag

                            ExtA.Start(IC.Info)
                        End If
                    End If

                Catch ex As Exception
                    App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "cmenTabExternalAppsEntry_Click failed (UI.Window.Tree)" & vbNewLine & ex.ToString(), True)
                End Try
            End Sub

            Private Sub CloseTabMenu()
                Try
                    If Me.TabController.SelectedTab IsNot Nothing Then
                        If TypeOf Me.TabController.SelectedTab.Tag Is InterfaceControl Then
                            Dim IC As InterfaceControl = Me.TabController.SelectedTab.Tag

                            IC.Protocol.Close()
                        End If
                    End If
                Catch ex As Exception
                    App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "CloseTabMenu (UI.Window.Connections) failed" & vbNewLine & ex.ToString(), True)
                End Try
            End Sub

            Private Sub DuplicateTab()
                Try
                    If Me.TabController.SelectedTab IsNot Nothing Then
                        If TypeOf Me.TabController.SelectedTab.Tag Is InterfaceControl Then
                            Dim IC As InterfaceControl = Me.TabController.SelectedTab.Tag

                            Runtime.OpenConnection(IC.Info, Info.Force.DoNotJump)
                            _ignoreChangeSelectedTabClick = False
                        End If
                    End If
                Catch ex As Exception
                    App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "DuplicateTab (UI.Window.Connections) failed" & vbNewLine & ex.ToString(), True)
                End Try
            End Sub

            Private Sub Reconnect()
                Try
                    If Me.TabController.SelectedTab IsNot Nothing Then
                        If TypeOf Me.TabController.SelectedTab.Tag Is InterfaceControl Then
                            Dim IC As InterfaceControl = Me.TabController.SelectedTab.Tag
                            Dim conI As Info = IC.Info

                            IC.Protocol.Close()

                            Runtime.OpenConnection(conI, Info.Force.DoNotJump)
                        End If
                    End If
                Catch ex As Exception
                    App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "Reconnect (UI.Window.Connections) failed" & vbNewLine & ex.ToString(), True)
                End Try
            End Sub

            Private Sub RenameTab()
                Try
                    Dim nTitle As String = InputBox(Language.Language.strNewTitle & ":", ,
                                                    Me.TabController.SelectedTab.Title.Replace("&&", "&"))

                    If nTitle <> "" Then
                        Me.TabController.SelectedTab.Title = nTitle.Replace("&", "&&")
                    End If
                Catch ex As Exception
                    App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "RenameTab (UI.Window.Connections) failed" & vbNewLine & ex.ToString(), True)
                End Try
            End Sub

#End Region

#Region "Protocols"

            Public Sub Prot_Event_Closed(sender As Object)
                Dim Prot As Protocol.Base = sender
                CloseTab(Prot.InterfaceControl.Parent)
            End Sub

#End Region

#Region "Tabs"

            Private Delegate Sub CloseTabCB(TabToBeClosed As TabPage)

            Private Sub CloseTab(TabToBeClosed As TabPage)
                If Me.TabController.InvokeRequired Then
                    Dim s As New CloseTabCB(AddressOf CloseTab)

                    Try
                        Me.TabController.Invoke(s, TabToBeClosed)
                    Catch comEx As COMException
                        Me.TabController.Invoke(s, TabToBeClosed)
                    Catch ex As Exception
                        App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "Couldn't close tab" & vbNewLine & ex.ToString(), True)
                    End Try
                Else
                    Try
                        Me.TabController.TabPages.Remove(TabToBeClosed)
                        _ignoreChangeSelectedTabClick = False
                    Catch comEx As COMException
                        CloseTab(TabToBeClosed)
                    Catch ex As Exception
                        App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "Couldn't close tab" & vbNewLine & ex.ToString(), True)
                    End Try

                    If Me.TabController.TabPages.Count = 0 Then
                        Me.Close()
                    End If
                End If
            End Sub

            Private _ignoreChangeSelectedTabClick As Boolean = False

            Private Sub TabController_SelectionChanged(sender As Object, e As EventArgs) _
                Handles TabController.SelectionChanged
                _ignoreChangeSelectedTabClick = True
                UpdateSelectedConnection()
                FocusIC()
                RefreshIC()
            End Sub

            Private _firstClickTicks As Integer = 0
            Private _doubleClickRectangle As Rectangle

            Private Sub TabController_MouseUp(sender As Object, e As MouseEventArgs) Handles TabController.MouseUp
                Try
                    If Not NativeMethods.GetForegroundWindow() = frmMain.Handle And Not _ignoreChangeSelectedTabClick _
                        Then
                        Dim clickedTab As TabPage = TabController.TabPageFromPoint(e.Location)
                        If clickedTab IsNot Nothing And TabController.SelectedTab IsNot clickedTab Then
                            NativeMethods.SetForegroundWindow(Handle)
                            TabController.SelectedTab = clickedTab
                        End If
                    End If
                    _ignoreChangeSelectedTabClick = False

                    Select Case e.Button
                        Case MouseButtons.Left
                            Dim currentTicks As Integer = Environment.TickCount
                            Dim elapsedTicks As Integer = currentTicks - _firstClickTicks
                            If _
                                elapsedTicks > SystemInformation.DoubleClickTime Or
                                Not _doubleClickRectangle.Contains(MousePosition) Then
                                _firstClickTicks = currentTicks
                                _doubleClickRectangle =
                                    New Rectangle(MousePosition.X - (SystemInformation.DoubleClickSize.Width / 2),
                                                  MousePosition.Y - (SystemInformation.DoubleClickSize.Height / 2),
                                                  SystemInformation.DoubleClickSize.Width,
                                                  SystemInformation.DoubleClickSize.Height)
                                FocusIC()
                            Else
                                TabController.OnDoubleClickTab(TabController.SelectedTab)
                            End If
                        Case MouseButtons.Middle
                            CloseConnectionTab()
                        Case MouseButtons.Right
                            ShowHideMenuButtons()
                            NativeMethods.SetForegroundWindow(Handle)
                            cmenTab.Show(TabController, e.Location)
                    End Select
                Catch ex As Exception
                    App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "TabController_MouseUp (UI.Window.Connections) failed" & vbNewLine & ex.ToString(), True)
                End Try
            End Sub

            Private Sub FocusIC()
                Try
                    If Me.TabController.SelectedTab IsNot Nothing Then
                        If TypeOf Me.TabController.SelectedTab.Tag Is InterfaceControl Then
                            Dim IC As InterfaceControl = Me.TabController.SelectedTab.Tag
                            IC.Protocol.Focus()
                        End If
                    End If
                Catch ex As Exception
                    App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "FocusIC (UI.Window.Connections) failed" & vbNewLine & ex.ToString(), True)
                End Try
            End Sub

            Public Sub RefreshIC()
                Try
                    If Me.TabController.SelectedTab IsNot Nothing Then
                        If TypeOf Me.TabController.SelectedTab.Tag Is InterfaceControl Then
                            Dim IC As InterfaceControl = Me.TabController.SelectedTab.Tag

                            If IC.Info.Protocol = Protocols.VNC Then
                                TryCast(IC.Protocol, VNC).RefreshScreen()
                            End If
                        End If
                    End If
                Catch ex As Exception
                    App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "RefreshIC (UI.Window.Connection) failed" & vbNewLine & ex.ToString(), True)
                End Try
            End Sub

#End Region

#Region "Window Overrides"

            Protected Overloads Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
                Try
                    If m.Msg = NativeMethods.WM_MOUSEACTIVATE Then
                        Dim selectedTab As TabPage = TabController.SelectedTab
                        If selectedTab IsNot Nothing Then
                            Dim tabClientRectangle As Rectangle =
                                    selectedTab.RectangleToScreen(selectedTab.ClientRectangle)
                            If tabClientRectangle.Contains(MousePosition) Then
                                Dim interfaceControl = TryCast(TabController.SelectedTab.Tag, InterfaceControl)
                                If interfaceControl IsNot Nothing AndAlso interfaceControl.Info IsNot Nothing Then
                                    If interfaceControl.Info.Protocol = Protocols.RDP Then
                                        interfaceControl.Protocol.Focus()
                                        Return ' Do not pass to base class
                                    End If
                                End If
                            End If
                        End If
                    End If
                Catch ex As Exception
                    Runtime.MessageCollector.AddExceptionMessage("UI.Window.Connection.WndProc() failed.", ex, , True)
                End Try

                MyBase.WndProc(m)
            End Sub

#End Region

#Region "Tab drag and drop"

            Public Property InTabDrag As Boolean = False

            Private Sub TabController_PageDragStart(sender As Object, e As MouseEventArgs) _
                Handles TabController.PageDragEnd, TabController.PageDragStart
                Cursor = Cursors.SizeWE
            End Sub

            Private Sub TabController_PageDragMove(sender As Object, e As MouseEventArgs) _
                Handles TabController.PageDragMove
                InTabDrag = True _
                ' For some reason PageDragStart gets raised again after PageDragEnd so set this here instead

                Dim sourceTab As TabPage = TabController.SelectedTab
                Dim destinationTab As TabPage = TabController.TabPageFromPoint(e.Location)

                If Not TabController.TabPages.Contains(destinationTab) Or sourceTab Is destinationTab Then Return

                Dim targetIndex As Integer = TabController.TabPages.IndexOf(destinationTab)

                TabController.TabPages.SuspendEvents()
                TabController.TabPages.Remove(sourceTab)
                TabController.TabPages.Insert(targetIndex, sourceTab)
                TabController.SelectedTab = sourceTab
                TabController.TabPages.ResumeEvents()
            End Sub

            Private Sub TabController_PageDragEnd(sender As Object, e As MouseEventArgs) _
                Handles TabController.PageDragEnd, TabController.PageDragQuit
                Cursor = Cursors.Default
                InTabDrag = False
                Dim interfaceControl = TryCast(TabController.SelectedTab.Tag, InterfaceControl)
                If interfaceControl IsNot Nothing Then interfaceControl.Protocol.Focus()
            End Sub

#End Region
        End Class
    End Namespace

End Namespace