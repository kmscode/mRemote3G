
Imports mRemote3G.App.Info
Imports mRemote3G.Tools

Namespace Root
    Namespace PuttySessions
        Public Class PuttySessionsInfo
            Inherits Root.Info

            Public Sub New()
                MyBase.New(RootType.PuttySessions)
            End Sub

            Private _name As String = Language.Language.strPuttySavedSessionsRootName
            <LocalizedAttributes.LocalizedDefaultValue("strPuttySavedSessionsRootName")>
            Public Overrides Property Name() As String
                Get
                    Return _name
                End Get
                Set(ByVal value As String)
                    If _name = value Then Return
                    _name = value
                    If TreeNode IsNot Nothing Then
                        TreeNode.Text = value
                    End If
                    My.Settings.PuttySavedSessionsName = value
                End Set
            End Property

            Private _panel As String = Language.Language.strGeneral
            <LocalizedAttributes.LocalizedCategory("strCategoryDisplay", 1),
            LocalizedAttributes.LocalizedDisplayName("strPropertyNamePanel"),
            LocalizedAttributes.LocalizedDescription("strPropertyDescriptionPanel")>
            Public Property Panel() As String
                Get
                    Return _panel
                End Get
                Set(ByVal value As String)
                    If _panel = value Then Return
                    _panel = value
                    My.Settings.PuttySavedSessionsPanel = value
                End Set
            End Property
        End Class
    End Namespace
End Namespace

