Imports mRemote3G.App

Namespace Forms

    Public Class frmChoosePanel
        Public Property Panel() As String
            Get
                Return cbPanels.SelectedItem.ToString
            End Get
            Set(ByVal value As String)
                cbPanels.SelectedItem = value
            End Set
        End Property

        Private Sub frmChoosePanel_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
            ApplyLanguage()

            AddAvailablePanels()
        End Sub

        Private Sub ApplyLanguage()
            btnOK.Text = Language.Language.strButtonOK
            lblDescription.Text = Language.Language.strLabelSelectPanel
            btnNew.Text = Language.Language.strButtonNew
            btnCancel.Text = Language.Language.strButtonCancel
            Text = Language.Language.strTitleSelectPanel
        End Sub

        Private Sub AddAvailablePanels()
            cbPanels.Items.Clear()

            For i As Integer = 0 To App.Runtime.WindowList.Count - 1
                cbPanels.Items.Add(Runtime.WindowList(i).Text.Replace("&&", "&"))
            Next

            If cbPanels.Items.Count > 0 Then
                cbPanels.SelectedItem = cbPanels.Items(0)
                cbPanels.Enabled = True
                btnOK.Enabled = True
            Else
                cbPanels.Enabled = False
                btnOK.Enabled = False
            End If
        End Sub

        Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
            Dim pnlName As String = InputBox(Language.Language.strPanelName & ":", Language.Language.strNewPanel, Language.Language.strNewPanel)

            If pnlName <> "" Then
               App.Runtime.AddPanel(pnlName)
                AddAvailablePanels()
                cbPanels.SelectedItem = pnlName
                cbPanels.Focus()
            End If
        End Sub

        Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
        End Sub

        Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
            Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        End Sub
    End Class
End NameSpace