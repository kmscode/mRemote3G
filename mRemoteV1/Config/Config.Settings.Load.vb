Imports System.IO
Imports System.Xml
Imports mRemote3G.Forms
Imports mRemote3G.My
Imports WeifenLuo.WinFormsUI.Docking

Namespace Config
    Namespace Settings
        Public Class Load
#Region "Public Properties"
            Private _MainForm As frmMain
            Public Property MainForm() As frmMain
                Get
                    Return Me._MainForm
                End Get
                Set(ByVal value As frmMain)
                    Me._MainForm = value
                End Set
            End Property
#End Region

#Region "Public Methods"
            Public Sub New(ByVal MainForm As frmMain)
                Me._MainForm = MainForm
            End Sub

            Public Sub Load()
                Try
                    With Me._MainForm
                        ' Migrate settings from previous version
                        If MySettingsProperty.Settings.DoUpgrade Then
                            Try
                                MySettingsProperty.Settings.Upgrade()
                            Catch ex As Exception
                                App.Runtime.Log.Error("My.Settings.Upgrade() failed" & vbNewLine & ex.ToString())
                            End Try
                            MySettingsProperty.Settings.DoUpgrade = False

                            ' Clear pending update flag
                            ' This is used for automatic updates, not for settings migration, but it
                            ' needs to be cleared here because we know that we just updated.
                            MySettingsProperty.Settings.UpdatePending = False
                        End If

                        App.SupportedCultures.InstantiateSingleton()
                        If Not MySettingsProperty.Settings.OverrideUICulture = "" And App.SupportedCultures.IsNameSupported(MySettingsProperty.Settings.OverrideUICulture) Then
                            Threading.Thread.CurrentThread.CurrentUICulture = New Globalization.CultureInfo(MySettingsProperty.Settings.OverrideUICulture)
                            App.Runtime.Log.InfoFormat("Override Culture: {0}/{1}", Threading.Thread.CurrentThread.CurrentUICulture.Name, Threading.Thread.CurrentThread.CurrentUICulture.NativeName)
                        End If

                        Themes.ThemeManager.LoadTheme(MySettingsProperty.Settings.ThemeName)

                        .WindowState = FormWindowState.Normal
                        If MySettingsProperty.Settings.MainFormState = FormWindowState.Normal Then
                            If Not MySettingsProperty.Settings.MainFormLocation.IsEmpty Then .Location = MySettingsProperty.Settings.MainFormLocation
                            If Not MySettingsProperty.Settings.MainFormSize.IsEmpty Then .Size = MySettingsProperty.Settings.MainFormSize
                        Else
                            If Not MySettingsProperty.Settings.MainFormRestoreLocation.IsEmpty Then .Location = MySettingsProperty.Settings.MainFormRestoreLocation
                            If Not MySettingsProperty.Settings.MainFormRestoreSize.IsEmpty Then .Size = MySettingsProperty.Settings.MainFormRestoreSize
                        End If
                        If MySettingsProperty.Settings.MainFormState = FormWindowState.Maximized Then
                            .WindowState = FormWindowState.Maximized
                        End If

                        ' Make sure the form is visible on the screen
                        Const minHorizontal As Integer = 300
                        Const minVertical As Integer = 150
                        Dim screenBounds As Drawing.Rectangle = Screen.FromHandle(.Handle).Bounds
                        Dim newBounds As Drawing.Rectangle = .Bounds

                        If newBounds.Right < screenBounds.Left + minHorizontal Then
                            newBounds.X = screenBounds.Left + minHorizontal - newBounds.Width
                        End If
                        If newBounds.Left > screenBounds.Right - minHorizontal Then
                            newBounds.X = screenBounds.Right - minHorizontal
                        End If
                        If newBounds.Bottom < screenBounds.Top + minVertical Then
                            newBounds.Y = screenBounds.Top + minVertical - newBounds.Height
                        End If
                        If newBounds.Top > screenBounds.Bottom - minVertical Then
                            newBounds.Y = screenBounds.Bottom - minVertical
                        End If

                        .Location = newBounds.Location

                        If MySettingsProperty.Settings.MainFormKiosk = True Then
                            .Fullscreen.Value = True
                            .mMenViewFullscreen.Checked = True
                        End If

                        If MySettingsProperty.Settings.UseCustomPuttyPath Then
                            Connection.Protocol.PuttyBase.PuttyPath = MySettingsProperty.Settings.CustomPuttyPath
                        Else
                            Connection.Protocol.PuttyBase.PuttyPath = App.Info.General.PuttyPath
                        End If

                        If MySettingsProperty.Settings.ShowSystemTrayIcon Then
                            App.Runtime.NotificationAreaIcon = New Tools.Controls.NotificationAreaIcon()
                        End If

                        If MySettingsProperty.Settings.AutoSaveEveryMinutes > 0 Then
                            .tmrAutoSave.Interval = MySettingsProperty.Settings.AutoSaveEveryMinutes * 60000
                            .tmrAutoSave.Enabled = True
                        End If

                        MySettingsProperty.Settings.ConDefaultPassword = Security.Crypt.Decrypt(MySettingsProperty.Settings.ConDefaultPassword, App.Info.General.EncryptionKey)

                        Me.LoadPanelsFromXML()
                        Me.LoadExternalAppsFromXML()

                        If MySettingsProperty.Settings.AlwaysShowPanelTabs Then
                            frmMain.pnlDock.DocumentStyle = DocumentStyle.DockingWindow
                        End If

                        If MySettingsProperty.Settings.ResetToolbars = False Then
                            LoadToolbarsFromSettings()
                        Else
                            SetToolbarsDefault()
                        End If
                    End With
                Catch ex As Exception
                    App.Runtime.Log.Error("Loading settings failed" & vbNewLine & ex.ToString())
                End Try
            End Sub

            Public Sub SetToolbarsDefault()
                With MainForm
                    ToolStripPanelFromString("top").Join(.tsQuickConnect, New Point(300, 0))
                    .tsQuickConnect.Visible = True
                    ToolStripPanelFromString("bottom").Join(.tsExternalTools, New Point(3, 0))
                    .tsExternalTools.Visible = False
                End With
            End Sub

            Public Sub LoadToolbarsFromSettings()
                With Me.MainForm
                    If MySettingsProperty.Settings.QuickyTBLocation.X > MySettingsProperty.Settings.ExtAppsTBLocation.X Then
                        AddDynamicPanels()
                        AddStaticPanels()
                    Else
                        AddStaticPanels()
                        AddDynamicPanels()
                    End If
                End With
            End Sub

            Private Sub AddStaticPanels()
                With MainForm
                    ToolStripPanelFromString(MySettingsProperty.Settings.QuickyTBParentDock).Join(.tsQuickConnect, MySettingsProperty.Settings.QuickyTBLocation)
                    .tsQuickConnect.Visible = MySettingsProperty.Settings.QuickyTBVisible
                End With
            End Sub

            Private Sub AddDynamicPanels()
                With MainForm
                    ToolStripPanelFromString(MySettingsProperty.Settings.ExtAppsTBParentDock).Join(.tsExternalTools, MySettingsProperty.Settings.ExtAppsTBLocation)
                    .tsExternalTools.Visible = MySettingsProperty.Settings.ExtAppsTBVisible
                End With
            End Sub

            Private Function ToolStripPanelFromString(ByVal Panel As String) As ToolStripPanel
                Select Case LCase(Panel)
                    Case "top"
                        Return MainForm.tsContainer.TopToolStripPanel
                    Case "bottom"
                        Return MainForm.tsContainer.BottomToolStripPanel
                    Case "left"
                        Return MainForm.tsContainer.LeftToolStripPanel
                    Case "right"
                        Return MainForm.tsContainer.RightToolStripPanel
                    Case Else
                        Return MainForm.tsContainer.TopToolStripPanel
                End Select
            End Function

            Public Sub LoadPanelsFromXML()
                Try
                    With MainForm
                        App.Runtime.Windows.treePanel = Nothing
                        App.Runtime.Windows.configPanel = Nothing
                        App.Runtime.Windows.errorsPanel = Nothing

                        Do While .pnlDock.Contents.Count > 0
                            Dim dc As WeifenLuo.WinFormsUI.Docking.DockContent = .pnlDock.Contents(0)
                            dc.Close()
                        Loop

                        App.Runtime.Startup.CreatePanels()

                        Dim oldPath As String = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) & "\" & My.Application.Info.ProductName & "\" & App.Info.Settings.LayoutFileName
                        Dim newPath As String = App.Info.Settings.SettingsPath & "\" & App.Info.Settings.LayoutFileName
                        If File.Exists(newPath) Then
                            .pnlDock.LoadFromXml(newPath, AddressOf GetContentFromPersistString)
#If Not PORTABLE Then
                        ElseIf File.Exists(oldPath) Then
                            .pnlDock.LoadFromXml(oldPath, AddressOf GetContentFromPersistString)
#End If
                        Else
                            App.Runtime.Startup.SetDefaultLayout()
                        End If
                    End With
                Catch ex As Exception
                    App.Runtime.Log.Error("LoadPanelsFromXML failed" & vbNewLine & ex.ToString())
                End Try
            End Sub

            Public Sub LoadExternalAppsFromXML()
                Dim oldPath As String = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) & "\" & My.Application.Info.ProductName & "\" & App.Info.Settings.ExtAppsFilesName
                Dim newPath As String = App.Info.Settings.SettingsPath & "\" & App.Info.Settings.ExtAppsFilesName
                Dim xDom As New XmlDocument()
                If File.Exists(newPath) Then
                    xDom.Load(newPath)
#If Not PORTABLE Then
                ElseIf File.Exists(oldPath) Then
                    xDom.Load(oldPath)
#End If
                Else
                    Exit Sub
                End If

                For Each xEl As XmlElement In xDom.DocumentElement.ChildNodes
                    Dim extA As New Tools.ExternalTool
                    extA.DisplayName = xEl.Attributes("DisplayName").Value
                    extA.FileName = xEl.Attributes("FileName").Value
                    extA.Arguments = xEl.Attributes("Arguments").Value

                    If xEl.HasAttribute("WaitForExit") Then
                        extA.WaitForExit = xEl.Attributes("WaitForExit").Value
                    End If

                    If xEl.HasAttribute("TryToIntegrate") Then
                        extA.TryIntegrate = xEl.Attributes("TryToIntegrate").Value
                    End If

                    App.Runtime.ExternalTools.Add(extA)
                Next

                MainForm.SwitchToolBarText(MySettingsProperty.Settings.ExtAppsTBShowText)

                xDom = Nothing

                frmMain.AddExternalToolsToToolBar()
            End Sub
#End Region

#Region "Private Methods"
            Private Function GetContentFromPersistString(ByVal persistString As String) As IDockContent
                ' pnlLayout.xml persistence XML fix for refactoring to mRemoteNG
                If (persistString.StartsWith("mRemote.")) Then
                    persistString = persistString.Replace("mRemote.", "mRemote3G.")
                End If

                Try
                    If persistString = GetType(UI.Window.Config).ToString Then
                        Return App.Runtime.Windows.configPanel
                    End If

                    If persistString = GetType(UI.Window.Tree).ToString Then
                        Return App.Runtime.Windows.treePanel
                    End If

                    If persistString = GetType(UI.Window.ErrorsAndInfos).ToString Then
                        Return App.Runtime.Windows.errorsPanel
                    End If

                    If persistString = GetType(UI.Window.ScreenshotManager).ToString Then
                        Return App.Runtime.Windows.screenshotPanel
                    End If
                Catch ex As Exception
                    App.Runtime.Log.Error("GetContentFromPersistString failed" & vbNewLine & ex.ToString())
                End Try

                Return Nothing
            End Function
#End Region
        End Class
    End Namespace
End Namespace