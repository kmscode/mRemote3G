Imports WeifenLuo.WinFormsUI.Docking
Imports mRemote3G.App

Namespace UI
    Namespace Window
        Public Class Tree
            Inherits Base
#Region "Form Stuff"
            Private Sub Tree_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
                ApplyLanguage()

                AddHandler Themes.ThemeManager.ThemeChanged, AddressOf ApplyTheme
                ApplyTheme()

                txtSearch.Multiline = True
                txtSearch.MinimumSize = New Size(0, 14)
                txtSearch.Size = New Size(txtSearch.Size.Width, 14)
                txtSearch.Multiline = False
            End Sub

            Private Sub ApplyLanguage()
                Text = Language.Language.strConnections
                TabText = Language.Language.strConnections

                mMenAddConnection.ToolTipText = Language.Language.strAddConnection
                mMenAddFolder.ToolTipText = Language.Language.strAddFolder
                mMenView.ToolTipText = Language.Language.strMenuView.Replace("&", "")
                mMenViewExpandAllFolders.Text = Language.Language.strExpandAllFolders
                mMenViewCollapseAllFolders.Text = Language.Language.strCollapseAllFolders
                mMenSortAscending.ToolTipText = Language.Language.strSortAsc

                cMenTreeConnect.Text = Language.Language.strConnect
                cMenTreeConnectWithOptions.Text = Language.Language.strConnectWithOptions
                cMenTreeConnectWithOptionsConnectToConsoleSession.Text = Language.Language.strConnectToConsoleSession
                cMenTreeConnectWithOptionsDontConnectToConsoleSession.Text = Language.Language.strDontConnectToConsoleSessionMenuItem
                cMenTreeConnectWithOptionsConnectInFullscreen.Text = Language.Language.strConnectInFullscreen
                cMenTreeConnectWithOptionsNoCredentials.Text = Language.Language.strConnectNoCredentials
                cMenTreeConnectWithOptionsChoosePanelBeforeConnecting.Text = Language.Language.strChoosePanelBeforeConnecting
                cMenTreeDisconnect.Text = Language.Language.strMenuDisconnect

                cMenTreeToolsExternalApps.Text = Language.Language.strMenuExternalTools
                cMenTreeToolsTransferFile.Text = Language.Language.strMenuTransferFile

                cMenTreeDuplicate.Text = Language.Language.strDuplicate
                cMenTreeRename.Text = Language.Language.strRename
                cMenTreeDelete.Text = Language.Language.strMenuDelete

                cMenTreeImport.Text = Language.Language.strImportMenuItem
                cMenTreeImportFile.Text = Language.Language.strImportFromFileMenuItem
                cMenTreeImportActiveDirectory.Text = Language.Language.strImportAD
                cMenTreeImportPortScan.Text = Language.Language.strImportPortScan
                cMenTreeExportFile.Text = Language.Language.strExportToFileMenuItem

                cMenTreeAddConnection.Text = Language.Language.strAddConnection
                cMenTreeAddFolder.Text = Language.Language.strAddFolder

                cMenTreeToolsSort.Text = Language.Language.strSort
                cMenTreeToolsSortAscending.Text = Language.Language.strSortAsc
                cMenTreeToolsSortDescending.Text = Language.Language.strSortDesc
                cMenTreeMoveUp.Text = Language.Language.strMoveUp
                cMenTreeMoveDown.Text = Language.Language.strMoveDown

                txtSearch.Text = Language.Language.strSearchPrompt
            End Sub

            Public Sub ApplyTheme()
                With Themes.ThemeManager.ActiveTheme
                    msMain.BackColor = .ToolbarBackgroundColor
                    msMain.ForeColor = .ToolbarTextColor
                    tvConnections.BackColor = .ConnectionsPanelBackgroundColor
                    tvConnections.ForeColor = .ConnectionsPanelTextColor
                    tvConnections.LineColor = .ConnectionsPanelTreeLineColor
                    BackColor = .ToolbarBackgroundColor
                    txtSearch.BackColor = .SearchBoxBackgroundColor
                    txtSearch.ForeColor = .SearchBoxTextPromptColor
                End With
            End Sub
#End Region

#Region "Public Methods"
            Public Sub New(ByVal panel As DockContent)
                WindowType = Type.Tree
                DockPnl = panel
                InitializeComponent()
                FillImageList()

                DescriptionTooltip = New ToolTip
                DescriptionTooltip.InitialDelay = 300
                DescriptionTooltip.ReshowDelay = 0
            End Sub

            Public Sub InitialRefresh()
                tvConnections_AfterSelect(tvConnections, New TreeViewEventArgs(tvConnections.SelectedNode, TreeViewAction.ByMouse))
            End Sub
#End Region

#Region "Public Properties"
            Public Property DescriptionTooltip() As ToolTip
#End Region

#Region "Private Methods"
            Private Sub FillImageList()
                Try
                    imgListTree.Images.Add(My.Resources.Root)
                    imgListTree.Images.Add(My.Resources.Folder)
                    imgListTree.Images.Add(My.Resources.Play)
                    imgListTree.Images.Add(My.Resources.Pause)
                    imgListTree.Images.Add(My.Resources.PuttySessions)
                Catch ex As Exception
                    Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "FillImageList (UI.Window.Tree) failed" & vbNewLine & ex.ToString(), True)
                End Try
            End Sub

            Private Sub tvConnections_BeforeLabelEdit(ByVal sender As Object, ByVal e As NodeLabelEditEventArgs) Handles tvConnections.BeforeLabelEdit
                cMenTreeDelete.ShortcutKeys = Keys.None
            End Sub

            Private Sub tvConnections_AfterLabelEdit(ByVal sender As Object, ByVal e As NodeLabelEditEventArgs) Handles tvConnections.AfterLabelEdit
                Try
                    cMenTreeDelete.ShortcutKeys = Keys.Delete

                    mRemote3G.Tree.Node.FinishRenameSelectedNode(e.Label)
                    Runtime.Windows.configForm.pGrid_SelectedObjectChanged()
                    ShowHideTreeContextMenuItems(e.Node)
                    Runtime.SaveConnectionsBG()
                Catch ex As Exception
                   App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "tvConnections_AfterLabelEdit (UI.Window.Tree) failed" & vbNewLine & ex.ToString(), True)
                End Try
            End Sub

            Private Sub tvConnections_AfterSelect(ByVal sender As Object, ByVal e As TreeViewEventArgs) Handles tvConnections.AfterSelect
                Try
                    Select Case mRemote3G.Tree.Node.GetNodeType(e.Node)
                        Case mRemote3G.Tree.Node.Type.Connection, mRemote3G.Tree.Node.Type.PuttySession
                            Runtime.Windows.configForm.SetPropertyGridObject(e.Node.Tag)
                        Case mRemote3G.Tree.Node.Type.Container
                            Runtime.Windows.configForm.SetPropertyGridObject(TryCast(e.Node.Tag, Container.Info).ConnectionInfo)
                        Case mRemote3G.Tree.Node.Type.Root, mRemote3G.Tree.Node.Type.PuttyRoot
                            Runtime.Windows.configForm.SetPropertyGridObject(e.Node.Tag)
                        Case Else
                            Exit Sub
                    End Select

                    Runtime.Windows.configForm.pGrid_SelectedObjectChanged()
                    ShowHideTreeContextMenuItems(e.Node)

                    Runtime.LastSelected = mRemote3G.Tree.Node.GetConstantID(e.Node)
                Catch ex As Exception
                   App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "tvConnections_AfterSelect (UI.Window.Tree) failed" & vbNewLine & ex.ToString(), True)
                End Try
            End Sub

            Private Sub tvConnections_NodeMouseClick(ByVal sender As Object, ByVal e As TreeNodeMouseClickEventArgs) Handles tvConnections.NodeMouseClick
                Try
                    ShowHideTreeContextMenuItems(tvConnections.SelectedNode)
                    tvConnections.SelectedNode = e.Node

                    If e.Button = MouseButtons.Left Then
                        If My.Settings.SingleClickOnConnectionOpensIt And
                            (mRemote3G.Tree.Node.GetNodeType(e.Node) = mRemote3G.Tree.Node.Type.Connection Or
                             mRemote3G.Tree.Node.GetNodeType(e.Node) = mRemote3G.Tree.Node.Type.PuttySession) Then
                            Runtime.OpenConnection()
                        End If

                        If My.Settings.SingleClickSwitchesToOpenConnection And mRemote3G.Tree.Node.GetNodeType(e.Node) = mRemote3G.Tree.Node.Type.Connection Then
                            Runtime.SwitchToOpenConnection(e.Node.Tag)
                        End If
                    End If
                Catch ex As Exception
                   App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "tvConnections_NodeMouseClick (UI.Window.Tree) failed" & vbNewLine & ex.ToString(), True)
                End Try
            End Sub

            Private Shared Sub tvConnections_NodeMouseDoubleClick(ByVal sender As Object, ByVal e As TreeNodeMouseClickEventArgs) Handles tvConnections.NodeMouseDoubleClick
                If mRemote3G.Tree.Node.GetNodeType(mRemote3G.Tree.Node.SelectedNode) = mRemote3G.Tree.Node.Type.Connection Or
                   mRemote3G.Tree.Node.GetNodeType(mRemote3G.Tree.Node.SelectedNode) = mRemote3G.Tree.Node.Type.PuttySession Then
                    Runtime.OpenConnection()
                End If
            End Sub

            Private Sub tvConnections_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles tvConnections.MouseMove
                Try
                    mRemote3G.Tree.Node.SetNodeToolTip(e, DescriptionTooltip)
                Catch ex As Exception
                   App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "tvConnections_MouseMove (UI.Window.Tree) failed" & vbNewLine & ex.ToString(), True)
                End Try
            End Sub

            Private Shared Sub EnableMenuItemsRecursive(ByVal items As ToolStripItemCollection, Optional ByVal enable As Boolean = True)
                Dim menuItem As ToolStripMenuItem
                For Each item As ToolStripItem In items
                    menuItem = TryCast(item, ToolStripMenuItem)
                    If menuItem Is Nothing Then Continue For
                    menuItem.Enabled = enable
                    If menuItem.HasDropDownItems Then
                        EnableMenuItemsRecursive(menuItem.DropDownItems, enable)
                    End If
                Next
            End Sub

            Private Sub ShowHideTreeContextMenuItems(ByVal selectedNode As TreeNode)
                If selectedNode Is Nothing Then Return

                Try
                    cMenTree.Enabled = True
                    EnableMenuItemsRecursive(cMenTree.Items)

                    Select Case mRemote3G.Tree.Node.GetNodeType(selectedNode)
                        Case mRemote3G.Tree.Node.Type.Connection
                            Dim connectionInfo As mRemote3G.Connection.Info = selectedNode.Tag

                            If connectionInfo.OpenConnections.Count = 0 Then
                                cMenTreeDisconnect.Enabled = False
                            End If

                            If Not (connectionInfo.Protocol = mRemote3G.Connection.Protocol.Protocols.SSH1 Or
                                    connectionInfo.Protocol = mRemote3G.Connection.Protocol.Protocols.SSH2) Then
                                cMenTreeToolsTransferFile.Enabled = False
                            End If

                            If Not (connectionInfo.Protocol = mRemote3G.Connection.Protocol.Protocols.RDP) Then
                                cMenTreeConnectWithOptionsConnectInFullscreen.Enabled = False
                                cMenTreeConnectWithOptionsConnectToConsoleSession.Enabled = False
                            End If

                            If (connectionInfo.Protocol = mRemote3G.Connection.Protocol.Protocols.IntApp) Then
                                cMenTreeConnectWithOptionsNoCredentials.Enabled = False
                            End If
                        Case mRemote3G.Tree.Node.Type.PuttySession
                            Dim puttySessionInfo As mRemote3G.Connection.PuttySession.PuttyInfo = selectedNode.Tag

                            cMenTreeAddConnection.Enabled = False
                            cMenTreeAddFolder.Enabled = False

                            If puttySessionInfo.OpenConnections.Count = 0 Then
                                cMenTreeDisconnect.Enabled = False
                            End If

                            If Not (puttySessionInfo.Protocol = mRemote3G.Connection.Protocol.Protocols.SSH1 Or
                                    puttySessionInfo.Protocol = mRemote3G.Connection.Protocol.Protocols.SSH2) Then
                                cMenTreeToolsTransferFile.Enabled = False
                            End If

                            cMenTreeConnectWithOptionsConnectInFullscreen.Enabled = False
                            cMenTreeConnectWithOptionsConnectToConsoleSession.Enabled = False
                            cMenTreeToolsSort.Enabled = False
                            cMenTreeDuplicate.Enabled = False
                            cMenTreeRename.Enabled = False
                            cMenTreeDelete.Enabled = False
                            cMenTreeMoveUp.Enabled = False
                            cMenTreeMoveDown.Enabled = False
                        Case mRemote3G.Tree.Node.Type.Container
                            cMenTreeConnectWithOptionsConnectInFullscreen.Enabled = False
                            cMenTreeConnectWithOptionsConnectToConsoleSession.Enabled = False
                            cMenTreeDisconnect.Enabled = False

                            Dim openConnections As Integer = 0
                            Dim connectionInfo As mRemote3G.Connection.Info
                            For Each node As TreeNode In selectedNode.Nodes
                                If TypeOf node.Tag Is mRemote3G.Connection.Info Then
                                    connectionInfo = node.Tag
                                    openConnections = openConnections + connectionInfo.OpenConnections.Count
                                End If
                            Next
                            If openConnections = 0 Then
                                cMenTreeDisconnect.Enabled = False
                            End If

                            cMenTreeToolsTransferFile.Enabled = False
                            cMenTreeToolsExternalApps.Enabled = False
                        Case mRemote3G.Tree.Node.Type.Root
                            cMenTreeConnect.Enabled = False
                            cMenTreeConnectWithOptions.Enabled = False
                            cMenTreeConnectWithOptionsConnectInFullscreen.Enabled = False
                            cMenTreeConnectWithOptionsConnectToConsoleSession.Enabled = False
                            cMenTreeConnectWithOptionsChoosePanelBeforeConnecting.Enabled = False
                            cMenTreeDisconnect.Enabled = False
                            cMenTreeToolsTransferFile.Enabled = False
                            cMenTreeToolsExternalApps.Enabled = False
                            cMenTreeDuplicate.Enabled = False
                            cMenTreeDelete.Enabled = False
                            cMenTreeMoveUp.Enabled = False
                            cMenTreeMoveDown.Enabled = False
                        Case mRemote3G.Tree.Node.Type.PuttyRoot
                            cMenTreeAddConnection.Enabled = False
                            cMenTreeAddFolder.Enabled = False
                            cMenTreeConnect.Enabled = False
                            cMenTreeConnectWithOptions.Enabled = False
                            cMenTreeDisconnect.Enabled = False
                            cMenTreeToolsTransferFile.Enabled = False
                            cMenTreeConnectWithOptions.Enabled = False
                            cMenTreeToolsSort.Enabled = False
                            cMenTreeToolsExternalApps.Enabled = False
                            cMenTreeDuplicate.Enabled = False
                            cMenTreeRename.Enabled = True
                            cMenTreeDelete.Enabled = False
                            cMenTreeMoveUp.Enabled = False
                            cMenTreeMoveDown.Enabled = False
                        Case Else
                            cMenTree.Enabled = False
                    End Select
                Catch ex As Exception
                   App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "ShowHideTreeContextMenuItems (UI.Window.Tree) failed" & vbNewLine & ex.ToString(), True)
                End Try
            End Sub
#End Region

#Region "Drag and Drop"
            Private Shared Sub tvConnections_DragDrop(ByVal sender As Object, ByVal e As DragEventArgs) Handles tvConnections.DragDrop
                Try
                    'Check that there is a TreeNode being dragged
                    If e.Data.GetDataPresent("System.Windows.Forms.TreeNode", True) = False Then Exit Sub

                    'Get the TreeView raising the event (in case multiple on form)
                    Dim selectedTreeview As TreeView = CType(sender, TreeView)

                    'Get the TreeNode being dragged
                    Dim dropNode As TreeNode = CType(e.Data.GetData("System.Windows.Forms.TreeNode"), TreeNode)

                    'The target node should be selected from the DragOver event
                    Dim targetNode As TreeNode = selectedTreeview.SelectedNode

                    If dropNode Is targetNode Then
                        Exit Sub
                    End If

                    If mRemote3G.Tree.Node.GetNodeType(dropNode) = mRemote3G.Tree.Node.Type.Root Then
                        Exit Sub
                    End If

                    If dropNode Is targetNode.Parent Then
                        Exit Sub
                    End If

                    'Remove the drop node from its current location
                    dropNode.Remove()

                    'If there is no targetNode add dropNode to the bottom of
                    'the TreeView root nodes, otherwise add it to the end of
                    'the dropNode child nodes

                    If mRemote3G.Tree.Node.GetNodeType(targetNode) = mRemote3G.Tree.Node.Type.Root Or mRemote3G.Tree.Node.GetNodeType(targetNode) = mRemote3G.Tree.Node.Type.Container Then
                        targetNode.Nodes.Insert(0, dropNode)
                    Else
                        targetNode.Parent.Nodes.Insert(targetNode.Index + 1, dropNode)
                    End If

                    If mRemote3G.Tree.Node.GetNodeType(dropNode) = mRemote3G.Tree.Node.Type.Connection Or mRemote3G.Tree.Node.GetNodeType(dropNode) = mRemote3G.Tree.Node.Type.Container Then
                        If mRemote3G.Tree.Node.GetNodeType(dropNode.Parent) = mRemote3G.Tree.Node.Type.Container Then
                            dropNode.Tag.Parent = dropNode.Parent.Tag
                        ElseIf mRemote3G.Tree.Node.GetNodeType(dropNode.Parent) = mRemote3G.Tree.Node.Type.Root Then
                            dropNode.Tag.Parent = Nothing
                            If mRemote3G.Tree.Node.GetNodeType(dropNode) = mRemote3G.Tree.Node.Type.Connection Then
                                dropNode.Tag.Inherit.TurnOffInheritanceCompletely()
                            ElseIf mRemote3G.Tree.Node.GetNodeType(dropNode) = mRemote3G.Tree.Node.Type.Container Then
                                dropNode.Tag.ConnectionInfo.Inherit.TurnOffInheritanceCompletely()
                            End If
                        End If
                    End If

                    'Ensure the newly created node is visible to
                    'the user and select it
                    dropNode.EnsureVisible()
                    selectedTreeview.SelectedNode = dropNode

                    Runtime.SaveConnectionsBG()
                Catch ex As Exception
                   App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "tvConnections_DragDrop (UI.Window.Tree) failed" & vbNewLine & ex.ToString(), True)
                End Try
            End Sub

            Private Shared Sub tvConnections_DragEnter(ByVal sender As Object, ByVal e As DragEventArgs) Handles tvConnections.DragEnter
                Try
                    'See if there is a TreeNode being dragged
                    If e.Data.GetDataPresent("System.Windows.Forms.TreeNode", True) Then
                        'TreeNode found allow move effect
                        e.Effect = DragDropEffects.Move
                    Else
                        'No TreeNode found, prevent move
                        e.Effect = DragDropEffects.None
                    End If
                Catch ex As Exception
                   App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "tvConnections_DragEnter (UI.Window.Tree) failed" & vbNewLine & ex.ToString(), True)
                End Try
            End Sub

            Private Shared Sub tvConnections_DragOver(ByVal sender As Object, ByVal e As DragEventArgs) Handles tvConnections.DragOver
                Try
                    'Check that there is a TreeNode being dragged 
                    If e.Data.GetDataPresent("System.Windows.Forms.TreeNode", True) = False Then Exit Sub

                    'Get the TreeView raising the event (in case multiple on form)
                    Dim selectedTreeview As TreeView = CType(sender, TreeView)

                    'As the mouse moves over nodes, provide feedback to 
                    'the user by highlighting the node that is the 
                    'current drop target
                    Dim pt As Point = CType(sender, TreeView).PointToClient(New Point(e.X, e.Y))
                    Dim targetNode As TreeNode = selectedTreeview.GetNodeAt(pt)

                    'Select the node currently under the cursor
                    selectedTreeview.SelectedNode = targetNode

                    'Check that the selected node is not the dropNode and
                    'also that it is not a child of the dropNode and 
                    'therefore an invalid target
                    Dim dropNode As TreeNode = CType(e.Data.GetData("System.Windows.Forms.TreeNode"), TreeNode)

                    Dim puttyRootInfo As Root.PuttySessions.PuttySessionsInfo
                    Do Until targetNode Is Nothing
                        puttyRootInfo = TryCast(targetNode.Tag, Root.PuttySessions.PuttySessionsInfo)
                        If puttyRootInfo IsNot Nothing Or targetNode Is dropNode Then
                            e.Effect = DragDropEffects.None
                            Return
                        End If
                        targetNode = targetNode.Parent
                    Loop

                    'Currently selected node is a suitable target
                    e.Effect = DragDropEffects.Move
                Catch ex As Exception
                   App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "tvConnections_DragOver (UI.Window.Tree) failed" & vbNewLine & ex.ToString(), True)
                End Try
            End Sub

            Private Sub tvConnections_ItemDrag(ByVal sender As Object, ByVal e As ItemDragEventArgs) Handles tvConnections.ItemDrag
                Try
                    Dim dragTreeNode As TreeNode = TryCast(e.Item, TreeNode)
                    If dragTreeNode Is Nothing Then Return

                    If dragTreeNode.Tag Is Nothing Then Return
                    If TypeOf dragTreeNode.Tag Is mRemote3G.Connection.PuttySession.PuttyInfo Or
                       Not (TypeOf dragTreeNode.Tag Is mRemote3G.Connection.Info Or
                            TypeOf dragTreeNode.Tag Is Container.Info) Then
                        tvConnections.SelectedNode = dragTreeNode
                        Return
                    End If

                    'Set the drag node and initiate the DragDrop 
                    DoDragDrop(e.Item, DragDropEffects.Move)
                Catch ex As Exception
                   App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "tvConnections_ItemDrag (UI.Window.Tree) failed" & vbNewLine & ex.ToString(), True)
                End Try
            End Sub
#End Region

#Region "Tree Context Menu"
            Private Sub cMenTreeAddConnection_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles cMenTreeAddConnection.Click, mMenAddConnection.Click
                AddConnection()
                Runtime.SaveConnectionsBG()
            End Sub

            Private Sub cMenTreeAddFolder_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles cMenTreeAddFolder.Click, mMenAddFolder.Click
                AddFolder()
                Runtime.SaveConnectionsBG()
            End Sub

            Private Shared Sub cMenTreeConnect_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles cMenTreeConnect.Click
                Runtime.OpenConnection(mRemote3G.Connection.Info.Force.DoNotJump)
            End Sub

            Private Shared Sub cMenTreeConnectWithOptionsConnectToConsoleSession_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles cMenTreeConnectWithOptionsConnectToConsoleSession.Click
                Runtime.OpenConnection(mRemote3G.Connection.Info.Force.UseConsoleSession Or mRemote3G.Connection.Info.Force.DoNotJump)
            End Sub

            Private Shared Sub cMenTreeConnectWithOptionsNoCredentials_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles cMenTreeConnectWithOptionsNoCredentials.Click
                Runtime.OpenConnection(mRemote3G.Connection.Info.Force.NoCredentials)
            End Sub

            Private Shared Sub cMenTreeConnectWithOptionsDontConnectToConsoleSession_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles cMenTreeConnectWithOptionsDontConnectToConsoleSession.Click
                Runtime.OpenConnection(mRemote3G.Connection.Info.Force.DontUseConsoleSession Or mRemote3G.Connection.Info.Force.DoNotJump)
            End Sub

            Private Shared Sub cMenTreeConnectWithOptionsConnectInFullscreen_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles cMenTreeConnectWithOptionsConnectInFullscreen.Click
                Runtime.OpenConnection(mRemote3G.Connection.Info.Force.Fullscreen Or mRemote3G.Connection.Info.Force.DoNotJump)
            End Sub

            Private Shared Sub cMenTreeConnectWithOptionsChoosePanelBeforeConnecting_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles cMenTreeConnectWithOptionsChoosePanelBeforeConnecting.Click
                Runtime.OpenConnection(mRemote3G.Connection.Info.Force.OverridePanel Or mRemote3G.Connection.Info.Force.DoNotJump)
            End Sub

            Private Sub cMenTreeDisconnect_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles cMenTreeDisconnect.Click
                DisconnectConnection()
            End Sub

            Private Shared Sub cMenTreeToolsTransferFile_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles cMenTreeToolsTransferFile.Click
                SshTransferFile()
            End Sub

            Private Sub mMenSortAscending_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles mMenSortAscending.Click
                tvConnections.BeginUpdate()
                mRemote3G.Tree.Node.Sort(tvConnections.Nodes.Item(0), SortOrder.Ascending)
                tvConnections.EndUpdate()
                Runtime.SaveConnectionsBG()
            End Sub

            Private Sub cMenTreeToolsSortAscending_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles cMenTreeToolsSortAscending.Click
                tvConnections.BeginUpdate()
                mRemote3G.Tree.Node.Sort(tvConnections.SelectedNode, SortOrder.Ascending)
                tvConnections.EndUpdate()
                Runtime.SaveConnectionsBG()
            End Sub

            Private Sub cMenTreeToolsSortDescending_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles cMenTreeToolsSortDescending.Click
                tvConnections.BeginUpdate()
                mRemote3G.Tree.Node.Sort(tvConnections.SelectedNode, SortOrder.Descending)
                tvConnections.EndUpdate()
                Runtime.SaveConnectionsBG()
            End Sub

            Private Sub cMenTree_DropDownOpening(ByVal sender As Object, ByVal e As EventArgs) Handles cMenTree.Opening
                AddExternalApps()
            End Sub

            Private Shared Sub cMenTreeToolsExternalAppsEntry_Click(ByVal sender As Object, ByVal e As EventArgs)
                StartExternalApp(sender.Tag)
            End Sub

            Private Sub cMenTreeDuplicate_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles cMenTreeDuplicate.Click
                mRemote3G.Tree.Node.CloneNode(tvConnections.SelectedNode)
                Runtime.SaveConnectionsBG()
            End Sub

            Private Shared Sub cMenTreeRename_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles cMenTreeRename.Click
                mRemote3G.Tree.Node.StartRenameSelectedNode()
                Runtime.SaveConnectionsBG()
            End Sub

            Private Shared Sub cMenTreeDelete_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles cMenTreeDelete.Click
                mRemote3G.Tree.Node.DeleteSelectedNode()
                Runtime.SaveConnectionsBG()
            End Sub

            Private Shared Sub cMenTreeImportFile_Click(sender As System.Object, e As EventArgs) Handles cMenTreeImportFile.Click
                Import.ImportFromFile(Runtime.Windows.treeForm.tvConnections.Nodes(0), Runtime.Windows.treeForm.tvConnections.SelectedNode, True)
            End Sub

            Private Shared Sub cMenTreeImportActiveDirectory_Click(sender As System.Object, e As EventArgs) Handles cMenTreeImportActiveDirectory.Click
                Runtime.Windows.Show(Type.ActiveDirectoryImport)
            End Sub

            Private Shared Sub cMenTreeImportPortScan_Click(sender As System.Object, e As EventArgs) Handles cMenTreeImportPortScan.Click
                Runtime.Windows.Show(UI.Window.Type.PortScan, True)
            End Sub

            Private Shared Sub cMenTreeExportFile_Click(sender As System.Object, e As EventArgs) Handles cMenTreeExportFile.Click
                Export.ExportToFile(Runtime.Windows.treeForm.tvConnections.Nodes(0), Runtime.Windows.treeForm.tvConnections.SelectedNode)
            End Sub
            Private Shared Sub cMenTreeMoveUp_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles cMenTreeMoveUp.Click
                mRemote3G.Tree.Node.MoveNodeUp()
                Runtime.SaveConnectionsBG()
            End Sub

            Private Shared Sub cMenTreeMoveDown_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles cMenTreeMoveDown.Click
                mRemote3G.Tree.Node.MoveNodeDown()
                Runtime.SaveConnectionsBG()
            End Sub
#End Region

#Region "Context Menu Actions"
            Public Sub AddConnection()
                Try
                    If tvConnections.SelectedNode Is Nothing Then tvConnections.SelectedNode = tvConnections.Nodes.Item(0)

                    Dim newTreeNode As TreeNode = mRemote3G.Tree.Node.AddNode(mRemote3G.Tree.Node.Type.Connection)
                    If newTreeNode Is Nothing Then
                       App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "UI.Window.Tree.AddConnection() failed." & vbNewLine & "mRemote3G.Tree.Node.AddNode() returned Nothing.", True)
                        Return
                    End If

                    Dim containerNode As TreeNode = tvConnections.SelectedNode
                    If mRemote3G.Tree.Node.GetNodeType(containerNode) = mRemote3G.Tree.Node.Type.Connection Then
                        containerNode = containerNode.Parent
                    End If

                    Dim newConnectionInfo As New mRemote3G.Connection.Info()
                    If mRemote3G.Tree.Node.GetNodeType(containerNode) = mRemote3G.Tree.Node.Type.Root Then
                        newConnectionInfo.Inherit.TurnOffInheritanceCompletely()
                    Else
                        newConnectionInfo.Parent = containerNode.Tag
                    End If

                    newConnectionInfo.TreeNode = newTreeNode
                    newTreeNode.Tag = newConnectionInfo
                    App.Runtime.ConnectionList.Add(newConnectionInfo)

                    containerNode.Nodes.Add(newTreeNode)

                    tvConnections.SelectedNode = newTreeNode
                    tvConnections.SelectedNode.BeginEdit()
                Catch ex As Exception
                   App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "UI.Window.Tree.AddConnection() failed." & vbNewLine & ex.ToString(), True)
                End Try
            End Sub

            Public Sub AddFolder()
                Try
                    Dim newNode As TreeNode = mRemote3G.Tree.Node.AddNode(mRemote3G.Tree.Node.Type.Container)
                    Dim newContainerInfo As New Container.Info()
                    newNode.Tag = newContainerInfo
                    newContainerInfo.TreeNode = newNode

                    Dim selectedNode As TreeNode = mRemote3G.Tree.Node.SelectedNode
                    Dim parentNode As TreeNode
                    If selectedNode Is Nothing Then
                        parentNode = tvConnections.Nodes(0)
                    Else
                        If mRemote3G.Tree.Node.GetNodeType(selectedNode) = mRemote3G.Tree.Node.Type.Connection Then
                            parentNode = selectedNode.Parent
                        Else
                            parentNode = selectedNode
                        End If
                    End If

                    newContainerInfo.ConnectionInfo = New mRemote3G.Connection.Info(newContainerInfo)
                    newContainerInfo.ConnectionInfo.Name = newNode.Text

                    ' We can only inherit from a container node, not the root node or connection nodes
                    If mRemote3G.Tree.Node.GetNodeType(parentNode) = mRemote3G.Tree.Node.Type.Container Then
                        newContainerInfo.Parent = parentNode.Tag
                    Else
                        newContainerInfo.ConnectionInfo.Inherit.TurnOffInheritanceCompletely()
                    End If

                    Runtime.ContainerList.Add(newContainerInfo)
                    parentNode.Nodes.Add(newNode)

                    tvConnections.SelectedNode = newNode
                    tvConnections.SelectedNode.BeginEdit()
                Catch ex As Exception
                   App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, String.Format(Language.Language.strErrorAddFolderFailed, ex.ToString()), True)
                End Try
            End Sub

            Private Sub DisconnectConnection()
                Try
                    If tvConnections.SelectedNode IsNot Nothing Then
                        If TypeOf tvConnections.SelectedNode.Tag Is mRemote3G.Connection.Info Then
                            Dim conI As mRemote3G.Connection.Info = tvConnections.SelectedNode.Tag
                            For i As Integer = 0 To conI.OpenConnections.Count - 1
                                conI.OpenConnections(i).Disconnect()
                            Next
                        End If

                        If TypeOf tvConnections.SelectedNode.Tag Is Container.Info Then
                            For Each n As TreeNode In tvConnections.SelectedNode.Nodes
                                If TypeOf n.Tag Is mRemote3G.Connection.Info Then
                                    Dim conI As mRemote3G.Connection.Info = n.Tag
                                    For i As Integer = 0 To conI.OpenConnections.Count - 1
                                        conI.OpenConnections(i).Disconnect()
                                    Next
                                End If
                            Next
                        End If
                    End If
                Catch ex As Exception
                   App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "DisconnectConnection (UI.Window.Tree) failed" & vbNewLine & ex.ToString(), True)
                End Try
            End Sub

            Private Shared Sub SshTransferFile()
                Try
                    Runtime.Windows.Show(Type.SSHTransfer)

                    Dim conI As mRemote3G.Connection.Info = mRemote3G.Tree.Node.SelectedNode.Tag

                    Runtime.Windows.sshtransferForm.Hostname = conI.Hostname
                    Runtime.Windows.sshtransferForm.Username = conI.Username
                    Runtime.Windows.sshtransferForm.Password = conI.Password
                    Runtime.Windows.sshtransferForm.Port = conI.Port
                Catch ex As Exception
                   App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "SSHTransferFile (UI.Window.Tree) failed" & vbNewLine & ex.ToString(), True)
                End Try
            End Sub

            Private Sub AddExternalApps()
                Try
                    'clean up
                    cMenTreeToolsExternalApps.DropDownItems.Clear()

                    'add ext apps
                    For Each extA As Tools.ExternalTool In Runtime.ExternalTools
                        Dim nItem As New ToolStripMenuItem
                        nItem.Text = extA.DisplayName
                        nItem.Tag = extA

                        nItem.Image = extA.Image

                        AddHandler nItem.Click, AddressOf cMenTreeToolsExternalAppsEntry_Click

                        cMenTreeToolsExternalApps.DropDownItems.Add(nItem)
                    Next
                Catch ex As Exception
                   App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "cMenTreeTools_DropDownOpening failed (UI.Window.Tree)" & vbNewLine & ex.ToString(), True)
                End Try
            End Sub

            Private Shared Sub StartExternalApp(ByVal externalTool As Tools.ExternalTool)
                Try
                    If mRemote3G.Tree.Node.GetNodeType(mRemote3G.Tree.Node.SelectedNode) = mRemote3G.Tree.Node.Type.Connection Or
                       mRemote3G.Tree.Node.GetNodeType(mRemote3G.Tree.Node.SelectedNode) = mRemote3G.Tree.Node.Type.PuttySession Then
                        externalTool.Start(mRemote3G.Tree.Node.SelectedNode.Tag)
                    End If
                Catch ex As Exception
                   App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "cMenTreeToolsExternalAppsEntry_Click failed (UI.Window.Tree)" & vbNewLine & ex.ToString(), True)
                End Try
            End Sub
#End Region

#Region "Menu"
            Private Shared Sub mMenViewExpandAllFolders_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles mMenViewExpandAllFolders.Click
                mRemote3G.Tree.Node.ExpandAllNodes()
            End Sub

            Private Sub mMenViewCollapseAllFolders_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles mMenViewCollapseAllFolders.Click
                If tvConnections.SelectedNode IsNot Nothing Then
                    If tvConnections.SelectedNode.IsEditing Then tvConnections.SelectedNode.EndEdit(False)
                End If
                mRemote3G.Tree.Node.CollapseAllNodes()
            End Sub
#End Region

#Region "Search"
            Private Sub txtSearch_GotFocus(ByVal sender As Object, ByVal e As EventArgs) Handles txtSearch.GotFocus
                txtSearch.ForeColor = Themes.ThemeManager.ActiveTheme.SearchBoxTextColor
                If txtSearch.Text = Language.Language.strSearchPrompt Then
                    txtSearch.Text = ""
                End If
            End Sub

            Private Sub txtSearch_LostFocus(ByVal sender As Object, ByVal e As EventArgs) Handles txtSearch.LostFocus
                If txtSearch.Text = "" Then
                    txtSearch.ForeColor = Themes.ThemeManager.ActiveTheme.SearchBoxTextPromptColor
                    txtSearch.Text = Language.Language.strSearchPrompt
                End If
            End Sub

            Private Sub txtSearch_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles txtSearch.KeyDown
                Try
                    If e.KeyCode = Keys.Escape Then
                        e.Handled = True
                        tvConnections.Focus()
                    ElseIf e.KeyCode = Keys.Up Then
                        tvConnections.SelectedNode = tvConnections.SelectedNode.PrevVisibleNode
                    ElseIf e.KeyCode = Keys.Down Then
                        tvConnections.SelectedNode = tvConnections.SelectedNode.NextVisibleNode
                    Else
                        tvConnections_KeyDown(sender, e)
                    End If
                Catch ex As Exception
                   App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "txtSearch_KeyDown (UI.Window.Tree) failed" & vbNewLine & ex.ToString(), True)
                End Try
            End Sub

            Private Sub txtSearch_TextChanged(ByVal sender As System.Object, ByVal e As EventArgs) Handles txtSearch.TextChanged
                tvConnections.SelectedNode = mRemote3G.Tree.Node.Find(tvConnections.Nodes(0), txtSearch.Text)
            End Sub

            Private Sub tvConnections_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles tvConnections.KeyPress
                Try
                    If Char.IsLetterOrDigit(e.KeyChar) Then
                        txtSearch.Text = e.KeyChar

                        txtSearch.Focus()
                        txtSearch.SelectionStart = txtSearch.TextLength
                    End If
                Catch ex As Exception
                   App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "tvConnections_KeyPress (UI.Window.Tree) failed" & vbNewLine & ex.ToString(), True)
                End Try
            End Sub

            Private Sub tvConnections_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles tvConnections.KeyDown
                Try
                    If e.KeyCode = Keys.Enter Then
                        If TypeOf tvConnections.SelectedNode.Tag Is mRemote3G.Connection.Info Then
                            e.Handled = True
                            App.Runtime.OpenConnection()
                        Else
                            If tvConnections.SelectedNode.IsExpanded Then
                                tvConnections.SelectedNode.Collapse(True)
                            Else
                                tvConnections.SelectedNode.Expand()
                            End If
                        End If
                    ElseIf e.KeyCode = Keys.Escape Xor e.KeyCode = Keys.Control Or e.KeyCode = Keys.F Then
                        txtSearch.Focus()
                        txtSearch.SelectionStart = txtSearch.TextLength
                    End If
                Catch ex As Exception
                   App.Runtime.MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "tvConnections_KeyDown (UI.Window.Tree) failed" & vbNewLine & ex.ToString(), True)
                End Try
            End Sub
#End Region
        End Class
    End Namespace
End Namespace