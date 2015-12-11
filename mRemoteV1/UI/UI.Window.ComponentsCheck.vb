Imports WeifenLuo.WinFormsUI.Docking
Imports System.IO
Imports mRemoteNG.App.Runtime
Imports System.Threading

Namespace UI
    Namespace Window
        Public Class ComponentsCheck
            Inherits UI.Window.Base

#Region "Form Stuff"
            Friend WithEvents pbCheck1 As System.Windows.Forms.PictureBox
            Friend WithEvents lblCheck1 As System.Windows.Forms.Label
            Friend WithEvents pnlCheck2 As System.Windows.Forms.Panel
            Friend WithEvents lblCheck2 As System.Windows.Forms.Label
            Friend WithEvents pbCheck2 As System.Windows.Forms.PictureBox
            Friend WithEvents pnlCheck3 As System.Windows.Forms.Panel
            Friend WithEvents lblCheck3 As System.Windows.Forms.Label
            Friend WithEvents pbCheck3 As System.Windows.Forms.PictureBox
            Friend WithEvents pnlCheck4 As System.Windows.Forms.Panel
            Friend WithEvents lblCheck4 As System.Windows.Forms.Label
            Friend WithEvents pbCheck4 As System.Windows.Forms.PictureBox
            Friend WithEvents btnCheckAgain As System.Windows.Forms.Button
            Friend WithEvents txtCheck1 As System.Windows.Forms.TextBox
            Friend WithEvents txtCheck2 As System.Windows.Forms.TextBox
            Friend WithEvents txtCheck3 As System.Windows.Forms.TextBox
            Friend WithEvents txtCheck4 As System.Windows.Forms.TextBox
            Friend WithEvents chkAlwaysShow As System.Windows.Forms.CheckBox
            Friend WithEvents pnlChecks As System.Windows.Forms.Panel
            Friend WithEvents pnlCheck1 As System.Windows.Forms.Panel

            Private Sub InitializeComponent()
                Me.pnlCheck1 = New System.Windows.Forms.Panel()
                Me.txtCheck1 = New System.Windows.Forms.TextBox()
                Me.lblCheck1 = New System.Windows.Forms.Label()
                Me.pbCheck1 = New System.Windows.Forms.PictureBox()
                Me.pnlCheck2 = New System.Windows.Forms.Panel()
                Me.txtCheck2 = New System.Windows.Forms.TextBox()
                Me.lblCheck2 = New System.Windows.Forms.Label()
                Me.pbCheck2 = New System.Windows.Forms.PictureBox()
                Me.pnlCheck3 = New System.Windows.Forms.Panel()
                Me.txtCheck3 = New System.Windows.Forms.TextBox()
                Me.lblCheck3 = New System.Windows.Forms.Label()
                Me.pbCheck3 = New System.Windows.Forms.PictureBox()
                Me.pnlCheck4 = New System.Windows.Forms.Panel()
                Me.txtCheck4 = New System.Windows.Forms.TextBox()
                Me.lblCheck4 = New System.Windows.Forms.Label()
                Me.pbCheck4 = New System.Windows.Forms.PictureBox()
                Me.btnCheckAgain = New System.Windows.Forms.Button()
                Me.chkAlwaysShow = New System.Windows.Forms.CheckBox()
                Me.pnlChecks = New System.Windows.Forms.Panel()
                Me.pnlCheck1.SuspendLayout()
                CType(Me.pbCheck1, System.ComponentModel.ISupportInitialize).BeginInit()
                Me.pnlCheck2.SuspendLayout()
                CType(Me.pbCheck2, System.ComponentModel.ISupportInitialize).BeginInit()
                Me.pnlCheck3.SuspendLayout()
                CType(Me.pbCheck3, System.ComponentModel.ISupportInitialize).BeginInit()
                CType(Me.pbCheck4, System.ComponentModel.ISupportInitialize).BeginInit()
                Me.pnlChecks.SuspendLayout()
                Me.SuspendLayout()
                '
                'pnlCheck1
                '
                Me.pnlCheck1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
                Me.pnlCheck1.Controls.Add(Me.txtCheck1)
                Me.pnlCheck1.Controls.Add(Me.lblCheck1)
                Me.pnlCheck1.Controls.Add(Me.pbCheck1)
                Me.pnlCheck1.Location = New System.Drawing.Point(3, 3)
                Me.pnlCheck1.Name = "pnlCheck1"
                Me.pnlCheck1.Size = New System.Drawing.Size(562, 130)
                Me.pnlCheck1.TabIndex = 10
                Me.pnlCheck1.Visible = False
                '
                'txtCheck1
                '
                Me.txtCheck1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
                Me.txtCheck1.BackColor = System.Drawing.SystemColors.Control
                Me.txtCheck1.BorderStyle = System.Windows.Forms.BorderStyle.None
                Me.txtCheck1.Location = New System.Drawing.Point(129, 29)
                Me.txtCheck1.Multiline = True
                Me.txtCheck1.Name = "txtCheck1"
                Me.txtCheck1.ReadOnly = True
                Me.txtCheck1.Size = New System.Drawing.Size(430, 97)
                Me.txtCheck1.TabIndex = 2
                '
                'lblCheck1
                '
                Me.lblCheck1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
                Me.lblCheck1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                Me.lblCheck1.ForeColor = System.Drawing.SystemColors.ControlText
                Me.lblCheck1.Location = New System.Drawing.Point(108, 3)
                Me.lblCheck1.Name = "lblCheck1"
                Me.lblCheck1.Size = New System.Drawing.Size(451, 23)
                Me.lblCheck1.TabIndex = 1
                Me.lblCheck1.Text = "RDP check succeeded!"
                '
                'pbCheck1
                '
                Me.pbCheck1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
                Me.pbCheck1.Location = New System.Drawing.Point(3, 3)
                Me.pbCheck1.Name = "pbCheck1"
                Me.pbCheck1.Size = New System.Drawing.Size(72, 123)
                Me.pbCheck1.TabIndex = 0
                Me.pbCheck1.TabStop = False
                '
                'pnlCheck2
                '
                Me.pnlCheck2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
                Me.pnlCheck2.Controls.Add(Me.txtCheck2)
                Me.pnlCheck2.Controls.Add(Me.lblCheck2)
                Me.pnlCheck2.Controls.Add(Me.pbCheck2)
                Me.pnlCheck2.Location = New System.Drawing.Point(3, 139)
                Me.pnlCheck2.Name = "pnlCheck2"
                Me.pnlCheck2.Size = New System.Drawing.Size(562, 130)
                Me.pnlCheck2.TabIndex = 20
                Me.pnlCheck2.Visible = False
                '
                'txtCheck2
                '
                Me.txtCheck2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
                Me.txtCheck2.BackColor = System.Drawing.SystemColors.Control
                Me.txtCheck2.BorderStyle = System.Windows.Forms.BorderStyle.None
                Me.txtCheck2.Location = New System.Drawing.Point(129, 29)
                Me.txtCheck2.Multiline = True
                Me.txtCheck2.Name = "txtCheck2"
                Me.txtCheck2.ReadOnly = True
                Me.txtCheck2.Size = New System.Drawing.Size(430, 97)
                Me.txtCheck2.TabIndex = 2
                '
                'lblCheck2
                '
                Me.lblCheck2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
                Me.lblCheck2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                Me.lblCheck2.Location = New System.Drawing.Point(112, 3)
                Me.lblCheck2.Name = "lblCheck2"
                Me.lblCheck2.Size = New System.Drawing.Size(447, 23)
                Me.lblCheck2.TabIndex = 1
                Me.lblCheck2.Text = "RDP check succeeded!"
                '
                'pbCheck2
                '
                Me.pbCheck2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
                Me.pbCheck2.Location = New System.Drawing.Point(3, 3)
                Me.pbCheck2.Name = "pbCheck2"
                Me.pbCheck2.Size = New System.Drawing.Size(72, 123)
                Me.pbCheck2.TabIndex = 0
                Me.pbCheck2.TabStop = False
                '
                'pnlCheck3
                '
                Me.pnlCheck3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
                Me.pnlCheck3.Controls.Add(Me.txtCheck3)
                Me.pnlCheck3.Controls.Add(Me.lblCheck3)
                Me.pnlCheck3.Controls.Add(Me.pbCheck3)
                Me.pnlCheck3.Location = New System.Drawing.Point(3, 275)
                Me.pnlCheck3.Name = "pnlCheck3"
                Me.pnlCheck3.Size = New System.Drawing.Size(562, 130)
                Me.pnlCheck3.TabIndex = 30
                Me.pnlCheck3.Visible = False
                '
                'txtCheck3
                '
                Me.txtCheck3.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
                Me.txtCheck3.BackColor = System.Drawing.SystemColors.Control
                Me.txtCheck3.BorderStyle = System.Windows.Forms.BorderStyle.None
                Me.txtCheck3.Location = New System.Drawing.Point(129, 29)
                Me.txtCheck3.Multiline = True
                Me.txtCheck3.Name = "txtCheck3"
                Me.txtCheck3.ReadOnly = True
                Me.txtCheck3.Size = New System.Drawing.Size(430, 97)
                Me.txtCheck3.TabIndex = 2
                '
                'lblCheck3
                '
                Me.lblCheck3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
                Me.lblCheck3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                Me.lblCheck3.Location = New System.Drawing.Point(112, 3)
                Me.lblCheck3.Name = "lblCheck3"
                Me.lblCheck3.Size = New System.Drawing.Size(447, 23)
                Me.lblCheck3.TabIndex = 1
                Me.lblCheck3.Text = "RDP check succeeded!"
                '
                'pbCheck3
                '
                Me.pbCheck3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
                Me.pbCheck3.Location = New System.Drawing.Point(3, 3)
                Me.pbCheck3.Name = "pbCheck3"
                Me.pbCheck3.Size = New System.Drawing.Size(72, 123)
                Me.pbCheck3.TabIndex = 0
                Me.pbCheck3.TabStop = False
                '
                'pnlCheck4
                '
                Me.pnlCheck4.Location = New System.Drawing.Point(0, 0)
                Me.pnlCheck4.Name = "pnlCheck4"
                Me.pnlCheck4.Size = New System.Drawing.Size(200, 100)
                Me.pnlCheck4.TabIndex = 51
                '
                'txtCheck4
                '
                Me.txtCheck4.Location = New System.Drawing.Point(0, 0)
                Me.txtCheck4.Name = "txtCheck4"
                Me.txtCheck4.Size = New System.Drawing.Size(100, 20)
                Me.txtCheck4.TabIndex = 0
                '
                'lblCheck4
                '
                Me.lblCheck4.Location = New System.Drawing.Point(0, 0)
                Me.lblCheck4.Name = "lblCheck4"
                Me.lblCheck4.Size = New System.Drawing.Size(100, 23)
                Me.lblCheck4.TabIndex = 0
                '
                'pbCheck4
                '
                Me.pbCheck4.Location = New System.Drawing.Point(0, 0)
                Me.pbCheck4.Name = "pbCheck4"
                Me.pbCheck4.Size = New System.Drawing.Size(100, 50)
                Me.pbCheck4.TabIndex = 0
                Me.pbCheck4.TabStop = False
                '
                'btnCheckAgain
                '
                Me.btnCheckAgain.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
                Me.btnCheckAgain.FlatStyle = System.Windows.Forms.FlatStyle.Flat
                Me.btnCheckAgain.Location = New System.Drawing.Point(476, 585)
                Me.btnCheckAgain.Name = "btnCheckAgain"
                Me.btnCheckAgain.Size = New System.Drawing.Size(104, 23)
                Me.btnCheckAgain.TabIndex = 0
                Me.btnCheckAgain.Text = "Check again"
                Me.btnCheckAgain.UseVisualStyleBackColor = True
                '
                'chkAlwaysShow
                '
                Me.chkAlwaysShow.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
                Me.chkAlwaysShow.AutoSize = True
                Me.chkAlwaysShow.FlatStyle = System.Windows.Forms.FlatStyle.Flat
                Me.chkAlwaysShow.Location = New System.Drawing.Point(12, 586)
                Me.chkAlwaysShow.Name = "chkAlwaysShow"
                Me.chkAlwaysShow.Size = New System.Drawing.Size(225, 20)
                Me.chkAlwaysShow.TabIndex = 51
                Me.chkAlwaysShow.Text = "Always show this screen at startup"
                Me.chkAlwaysShow.UseVisualStyleBackColor = True
                '
                'pnlChecks
                '
                Me.pnlChecks.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
                Me.pnlChecks.AutoScroll = True
                Me.pnlChecks.Controls.Add(Me.pnlCheck1)
                Me.pnlChecks.Controls.Add(Me.pnlCheck2)
                Me.pnlChecks.Controls.Add(Me.pnlCheck3)
                Me.pnlChecks.Controls.Add(Me.pnlCheck4)
                Me.pnlChecks.Location = New System.Drawing.Point(12, 12)
                Me.pnlChecks.Name = "pnlChecks"
                Me.pnlChecks.Size = New System.Drawing.Size(568, 411)
                Me.pnlChecks.TabIndex = 52
                '
                'ComponentsCheck
                '
                Me.ClientSize = New System.Drawing.Size(592, 620)
                Me.Controls.Add(Me.pnlChecks)
                Me.Controls.Add(Me.chkAlwaysShow)
                Me.Controls.Add(Me.btnCheckAgain)
                Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                Me.Icon = Global.mRemoteNG.My.Resources.Resources.ComponentsCheck_Icon
                Me.Name = "ComponentsCheck"
                Me.TabText = "Components Check"
                Me.Text = "Components Check"
                Me.pnlCheck1.ResumeLayout(False)
                Me.pnlCheck1.PerformLayout()
                CType(Me.pbCheck1, System.ComponentModel.ISupportInitialize).EndInit()
                Me.pnlCheck2.ResumeLayout(False)
                Me.pnlCheck2.PerformLayout()
                CType(Me.pbCheck2, System.ComponentModel.ISupportInitialize).EndInit()
                Me.pnlCheck3.ResumeLayout(False)
                Me.pnlCheck3.PerformLayout()
                CType(Me.pbCheck3, System.ComponentModel.ISupportInitialize).EndInit()
                CType(Me.pbCheck4, System.ComponentModel.ISupportInitialize).EndInit()
                Me.pnlChecks.ResumeLayout(False)
                Me.ResumeLayout(False)
                Me.PerformLayout()

            End Sub
#End Region

#Region "Public Methods"
            Public Sub New(ByVal Panel As DockContent)
                Me.WindowType = Type.ComponentsCheck
                Me.DockPnl = Panel
                Me.InitializeComponent()
            End Sub
#End Region

#Region "Form Stuff"
            Private Sub ComponentsCheck_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
                ApplyLanguage()

                chkAlwaysShow.Checked = My.Settings.StartupComponentsCheck
                CheckComponents()
            End Sub

            Private Sub ApplyLanguage()
                TabText = My.Language.strComponentsCheck
                Text = My.Language.strComponentsCheck
                chkAlwaysShow.Text = My.Language.strCcAlwaysShowScreen
                btnCheckAgain.Text = My.Language.strCcCheckAgain
            End Sub

            Private Sub btnCheckAgain_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCheckAgain.Click
                CheckComponents()
            End Sub

            Private Sub chkAlwaysShow_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAlwaysShow.CheckedChanged
                My.Settings.StartupComponentsCheck = chkAlwaysShow.Checked
                My.Settings.Save()
            End Sub
#End Region

#Region "Private Methods"
            Private Sub CheckComponents()
                Dim errorMsg As String = My.Language.strCcNotInstalledProperly

                pnlCheck1.Visible = True
                pnlCheck2.Visible = True
                pnlCheck3.Visible = True
                pnlCheck4.Visible = True

                Dim rdpClient As AxMSTSCLib.AxMsRdpClient5NotSafeForScripting = Nothing

                Try
                    rdpClient = New AxMSTSCLib.AxMsRdpClient5NotSafeForScripting
                    rdpClient.CreateControl()

                    Do Until rdpClient.Created
                        Thread.Sleep(0)
                        System.Windows.Forms.Application.DoEvents()
                    Loop

                    If Not New Version(rdpClient.Version) >= mRemoteNG.Connection.Protocol.RDP.Versions.RDC60 Then
                        Throw New Exception(String.Format("Found RDC Client version {0} but version {1} or higher is required.", rdpClient.Version, mRemoteNG.Connection.Protocol.RDP.Versions.RDC60))
                    End If

                    pbCheck1.Image = My.Resources.Good_Symbol
                    lblCheck1.ForeColor = Color.DarkOliveGreen
                    lblCheck1.Text = "RDP (Remote Desktop) " & My.Language.strCcCheckSucceeded
                    txtCheck1.Text = String.Format(My.Language.strCcRDPOK, rdpClient.Version)
                Catch ex As Exception
                    pbCheck1.Image = My.Resources.Bad_Symbol
                    lblCheck1.ForeColor = Color.Firebrick
                    lblCheck1.Text = "RDP (Remote Desktop) " & My.Language.strCcCheckFailed
                    txtCheck1.Text = My.Language.strCcRDPFailed

                    MessageCollector.AddMessage(Messages.MessageClass.WarningMsg, "RDP " & errorMsg, True)
                    MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, ex.ToString(), True)
                End Try

                If rdpClient IsNot Nothing Then rdpClient.Dispose()


                Dim VNC As VncSharp.RemoteDesktop = Nothing

                Try
                    VNC = New VncSharp.RemoteDesktop
                    VNC.CreateControl()

                    Do Until VNC.Created
                        Thread.Sleep(10)
                        System.Windows.Forms.Application.DoEvents()
                    Loop

                    pbCheck2.Image = My.Resources.Good_Symbol
                    lblCheck2.ForeColor = Color.DarkOliveGreen
                    lblCheck2.Text = "VNC (Virtual Network Computing) " & My.Language.strCcCheckSucceeded
                    txtCheck2.Text = String.Format(My.Language.strCcVNCOK, VNC.ProductVersion)
                Catch ex As Exception
                    pbCheck2.Image = My.Resources.Bad_Symbol
                    lblCheck2.ForeColor = Color.Firebrick
                    lblCheck2.Text = "VNC (Virtual Network Computing) " & My.Language.strCcCheckFailed
                    txtCheck2.Text = My.Language.strCcVNCFailed

                    MessageCollector.AddMessage(Messages.MessageClass.WarningMsg, "VNC " & errorMsg, True)
                End Try

                If VNC IsNot Nothing Then VNC.Dispose()


                Dim pPath As String = ""
                If My.Settings.UseCustomPuttyPath = False Then
                    pPath = My.Application.Info.DirectoryPath & "\PuTTYNG.exe"
                Else
                    pPath = My.Settings.CustomPuttyPath
                End If

                If File.Exists(pPath) Then
                    pbCheck3.Image = My.Resources.Good_Symbol
                    lblCheck3.ForeColor = Color.DarkOliveGreen
                    lblCheck3.Text = "PuTTY (SSH/Telnet/Rlogin/RAW) " & My.Language.strCcCheckSucceeded
                    txtCheck3.Text = My.Language.strCcPuttyOK
                Else
                    pbCheck3.Image = My.Resources.Bad_Symbol
                    lblCheck3.ForeColor = Color.Firebrick
                    lblCheck3.Text = "PuTTY (SSH/Telnet/Rlogin/RAW) " & My.Language.strCcCheckFailed
                    txtCheck3.Text = My.Language.strCcPuttyFailed

                    MessageCollector.AddMessage(Messages.MessageClass.WarningMsg, "PuTTY " & errorMsg, True)
                    MessageCollector.AddMessage(Messages.MessageClass.ErrorMsg, "File " & pPath & " does not exist.", True)
                End If

            End Sub
#End Region

        End Class
    End Namespace
End Namespace