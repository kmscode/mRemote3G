Imports System.Windows.Forms
Imports System.ComponentModel
Imports mRemoteNG.Tools.LocalizedAttributes
Imports mRemoteNG.App.Runtime
Imports mRemoteNG.Config
Imports System.Reflection

Namespace Connection
    <DefaultProperty("Name")> _
    Public Class Info
#Region "Public Properties"
#Region "Display"
        <LocalizedCategory("strCategoryDisplay", 1), _
            LocalizedDisplayName("strPropertyNameName"), _
            LocalizedDescription("strPropertyDescriptionName")> _
        Public Overridable Property Name() As String = My.Language.strNewConnection

        Private _description As String = My.Settings.ConDefaultDescription
        <LocalizedCategory("strCategoryDisplay", 1), _
            LocalizedDisplayName("strPropertyNameDescription"), _
            LocalizedDescription("strPropertyDescriptionDescription")> _
        Public Overridable Property Description() As String
            Get
                Return GetInheritedPropertyValue("Description", _description)
            End Get
            Set(ByVal value As String)
                _description = value
            End Set
        End Property

        Private _icon As String = My.Settings.ConDefaultIcon
        <LocalizedCategory("strCategoryDisplay", 1), _
            TypeConverter(GetType(Icon)), _
            LocalizedDisplayName("strPropertyNameIcon"), _
            LocalizedDescription("strPropertyDescriptionIcon")> _
        Public Overridable Property Icon() As String
            Get
                Return GetInheritedPropertyValue("Icon", _icon)
            End Get
            Set(ByVal value As String)
                _icon = value
            End Set
        End Property

        Private _panel As String = My.Language.strGeneral
        <LocalizedCategory("strCategoryDisplay", 1), _
            LocalizedDisplayName("strPropertyNamePanel"), _
            LocalizedDescription("strPropertyDescriptionPanel")> _
        Public Overridable Property Panel() As String
            Get
                Return GetInheritedPropertyValue("Panel", _panel)
            End Get
            Set(ByVal value As String)
                _panel = value
            End Set
        End Property
#End Region
#Region "Connection"
        Private _hostname As String = String.Empty
        <LocalizedCategory("strCategoryConnection", 2),
            LocalizedDisplayName("strPropertyNameAddress"),
            LocalizedDescription("strPropertyDescriptionAddress")>
        Public Overridable Property Hostname() As String
            Get
                Return _hostname.Trim()
            End Get
            Set(ByVal value As String)
                If String.IsNullOrEmpty(value) Then
                    _hostname = String.Empty
                Else
                    _hostname = value.Trim()
                End If
            End Set
        End Property

        Private _username As String = My.Settings.ConDefaultUsername
        <LocalizedCategory("strCategoryConnection", 2),
            LocalizedDisplayName("strPropertyNameUsername"),
            LocalizedDescription("strPropertyDescriptionUsername")>
        Public Overridable Property Username() As String
            Get
                Return GetInheritedPropertyValue("Username", _username)
            End Get
            Set(ByVal value As String)
                _username = value.Trim()
            End Set
        End Property

        Private _password As String = My.Settings.ConDefaultPassword
        <LocalizedCategory("strCategoryConnection", 2),
            LocalizedDisplayName("strPropertyNamePassword"),
            LocalizedDescription("strPropertyDescriptionPassword"),
            PasswordPropertyText(True)>
        Public Overridable Property Password() As String
            Get
                Return GetInheritedPropertyValue("Password", _password)
            End Get
            Set(ByVal value As String)
                _password = value
            End Set
        End Property

        Private _domain As String = My.Settings.ConDefaultDomain
        <LocalizedCategory("strCategoryConnection", 2),
            LocalizedDisplayName("strPropertyNameDomain"),
            LocalizedDescription("strPropertyDescriptionDomain")>
        Public Property Domain() As String
            Get
                Return GetInheritedPropertyValue("Domain", _domain).Trim()
            End Get
            Set(ByVal value As String)
                _domain = value.Trim()
            End Set
        End Property
#End Region
#Region "Protocol"
        Private _protocol As Protocol.Protocols = Tools.Misc.StringToEnum(GetType(Protocol.Protocols), My.Settings.ConDefaultProtocol)
        <LocalizedCategory("strCategoryProtocol", 3),
            LocalizedDisplayName("strPropertyNameProtocol"),
            LocalizedDescription("strPropertyDescriptionProtocol"),
            TypeConverter(GetType(Tools.Misc.EnumTypeConverter))>
        Public Overridable Property Protocol() As Protocol.Protocols
            Get
                Return GetInheritedPropertyValue("Protocol", _protocol)
            End Get
            Set(ByVal value As Protocol.Protocols)
                _protocol = value
            End Set
        End Property

        Private _extApp As String = My.Settings.ConDefaultExtApp
        <LocalizedCategory("strCategoryProtocol", 3),
            LocalizedDisplayName("strPropertyNameExternalTool"),
            LocalizedDescription("strPropertyDescriptionExternalTool"),
            TypeConverter(GetType(Tools.ExternalToolsTypeConverter))>
        Public Property ExtApp() As String
            Get
                Return GetInheritedPropertyValue("ExtApp", _extApp)
            End Get
            Set(ByVal value As String)
                _extApp = value
            End Set
        End Property

        Private _port As Integer
        <LocalizedCategory("strCategoryProtocol", 3),
            LocalizedDisplayName("strPropertyNamePort"),
            LocalizedDescription("strPropertyDescriptionPort")>
        Public Overridable Property Port() As Integer
            Get
                Return GetInheritedPropertyValue("Port", _port)
            End Get
            Set(ByVal value As Integer)
                _port = value
            End Set
        End Property

        Private _puttySession As String = My.Settings.ConDefaultPuttySession
        <LocalizedCategory("strCategoryProtocol", 3),
            LocalizedDisplayName("strPropertyNamePuttySession"),
            LocalizedDescription("strPropertyDescriptionPuttySession"),
            TypeConverter(GetType(Putty.Sessions.SessionList))>
        Public Overridable Property PuttySession() As String
            Get
                Return GetInheritedPropertyValue("PuttySession", _puttySession)
            End Get
            Set(ByVal value As String)
                _puttySession = value
            End Set
        End Property

        Private _useConsoleSession As Boolean = My.Settings.ConDefaultUseConsoleSession
        <LocalizedCategory("strCategoryProtocol", 3),
            LocalizedDisplayName("strPropertyNameUseConsoleSession"),
            LocalizedDescription("strPropertyDescriptionUseConsoleSession"),
            TypeConverter(GetType(Tools.Misc.YesNoTypeConverter))>
        Public Property UseConsoleSession() As Boolean
            Get
                Return GetInheritedPropertyValue("UseConsoleSession", _useConsoleSession)
            End Get
            Set(ByVal value As Boolean)
                _useConsoleSession = value
            End Set
        End Property

        Private _rdpAuthenticationLevel As Protocol.RDP.AuthenticationLevel = Tools.Misc.StringToEnum(GetType(Protocol.RDP.AuthenticationLevel), My.Settings.ConDefaultRDPAuthenticationLevel)
        <LocalizedCategory("strCategoryProtocol", 3),
            LocalizedDisplayName("strPropertyNameAuthenticationLevel"),
            LocalizedDescription("strPropertyDescriptionAuthenticationLevel"),
            TypeConverter(GetType(Tools.Misc.EnumTypeConverter))>
        Public Property RDPAuthenticationLevel() As Protocol.RDP.AuthenticationLevel
            Get
                Return GetInheritedPropertyValue("RDPAuthenticationLevel", _rdpAuthenticationLevel)
            End Get
            Set(ByVal value As Protocol.RDP.AuthenticationLevel)
                _rdpAuthenticationLevel = value
            End Set
        End Property

        Private _loadBalanceInfo As String = My.Settings.ConDefaultLoadBalanceInfo
        <LocalizedCategory("strCategoryProtocol", 3),
            LocalizedDisplayName("strPropertyNameLoadBalanceInfo"),
            LocalizedDescription("strPropertyDescriptionLoadBalanceInfo")>
        Public Property LoadBalanceInfo() As String
            Get
                Return GetInheritedPropertyValue("LoadBalanceInfo", _loadBalanceInfo).Trim()
            End Get
            Set(ByVal value As String)
                _loadBalanceInfo = value.Trim()
            End Set
        End Property

        Private _renderingEngine As Protocol.HTTPBase.RenderingEngine = Tools.Misc.StringToEnum(GetType(Protocol.HTTPBase.RenderingEngine), My.Settings.ConDefaultRenderingEngine)
        <LocalizedCategory("strCategoryProtocol", 3),
            LocalizedDisplayName("strPropertyNameRenderingEngine"),
            LocalizedDescription("strPropertyDescriptionRenderingEngine"),
            TypeConverter(GetType(Tools.Misc.EnumTypeConverter))>
        Public Property RenderingEngine() As Protocol.HTTPBase.RenderingEngine
            Get
                Return GetInheritedPropertyValue("RenderingEngine", _renderingEngine)
            End Get
            Set(ByVal value As Protocol.HTTPBase.RenderingEngine)
                _renderingEngine = value
            End Set
        End Property

        Private _useCredSsp As Boolean = My.Settings.ConDefaultUseCredSsp
        <LocalizedCategory("strCategoryProtocol", 3),
            LocalizedDisplayName("strPropertyNameUseCredSsp"),
            LocalizedDescription("strPropertyDescriptionUseCredSsp"),
            TypeConverter(GetType(Tools.Misc.YesNoTypeConverter))>
        Public Property UseCredSsp() As Boolean
            Get
                Return GetInheritedPropertyValue("UseCredSsp", _useCredSsp)
            End Get
            Set(ByVal value As Boolean)
                _useCredSsp = value
            End Set
        End Property
#End Region
#Region "RD Gateway"
        Private _rdGatewayUsageMethod As Protocol.RDP.RDGatewayUsageMethod = Tools.Misc.StringToEnum(GetType(Protocol.RDP.RDGatewayUsageMethod), My.Settings.ConDefaultRDGatewayUsageMethod)
        <LocalizedCategory("strCategoryGateway", 4),
            LocalizedDisplayName("strPropertyNameRDGatewayUsageMethod"),
            LocalizedDescription("strPropertyDescriptionRDGatewayUsageMethod"),
            TypeConverter(GetType(Tools.Misc.EnumTypeConverter))>
        Public Property RDGatewayUsageMethod() As Protocol.RDP.RDGatewayUsageMethod
            Get
                Return GetInheritedPropertyValue("RDGatewayUsageMethod", _rdGatewayUsageMethod)
            End Get
            Set(ByVal value As Protocol.RDP.RDGatewayUsageMethod)
                _rdGatewayUsageMethod = value
            End Set
        End Property

        Private _rdGatewayHostname As String = My.Settings.ConDefaultRDGatewayHostname
        <LocalizedCategory("strCategoryGateway", 4),
            LocalizedDisplayName("strPropertyNameRDGatewayHostname"),
            LocalizedDescription("strPropertyDescriptionRDGatewayHostname")>
        Public Property RDGatewayHostname() As String
            Get
                Return GetInheritedPropertyValue("RDGatewayHostname", _rdGatewayHostname).Trim()
            End Get
            Set(ByVal value As String)
                _rdGatewayHostname = value.Trim()
            End Set
        End Property

        Private _rdGatewayUseConnectionCredentials As Protocol.RDP.RDGatewayUseConnectionCredentials = Tools.Misc.StringToEnum(GetType(Protocol.RDP.RDGatewayUseConnectionCredentials), My.Settings.ConDefaultRDGatewayUseConnectionCredentials)
        <LocalizedCategory("strCategoryGateway", 4),
            LocalizedDisplayName("strPropertyNameRDGatewayUseConnectionCredentials"),
            LocalizedDescription("strPropertyDescriptionRDGatewayUseConnectionCredentials"),
            TypeConverter(GetType(Tools.Misc.EnumTypeConverter))>
        Public Property RDGatewayUseConnectionCredentials() As Protocol.RDP.RDGatewayUseConnectionCredentials
            Get
                Return GetInheritedPropertyValue("RDGatewayUseConnectionCredentials", _rdGatewayUseConnectionCredentials)
            End Get
            Set(ByVal value As Protocol.RDP.RDGatewayUseConnectionCredentials)
                _rdGatewayUseConnectionCredentials = value
            End Set
        End Property

        Private _rdGatewayUsername As String = My.Settings.ConDefaultRDGatewayUsername
        <LocalizedCategory("strCategoryGateway", 4),
            LocalizedDisplayName("strPropertyNameRDGatewayUsername"),
            LocalizedDescription("strPropertyDescriptionRDGatewayUsername")>
        Public Property RDGatewayUsername() As String
            Get
                Return GetInheritedPropertyValue("RDGatewayUsername", _rdGatewayUsername).Trim()
            End Get
            Set(ByVal value As String)
                _rdGatewayUsername = value.Trim()
            End Set
        End Property

        Private _rdGatewayPassword As String = My.Settings.ConDefaultRDGatewayPassword
        <LocalizedCategory("strCategoryGateway", 4),
            LocalizedDisplayName("strPropertyNameRDGatewayPassword"),
            LocalizedDescription("strPropertyNameRDGatewayPassword"),
            PasswordPropertyText(True)>
        Public Property RDGatewayPassword() As String
            Get
                Return GetInheritedPropertyValue("RDGatewayPassword", _rdGatewayPassword)
            End Get
            Set(ByVal value As String)
                _rdGatewayPassword = value
            End Set
        End Property

        Private _rdGatewayDomain As String = My.Settings.ConDefaultRDGatewayDomain
        <LocalizedCategory("strCategoryGateway", 4),
            LocalizedDisplayName("strPropertyNameRDGatewayDomain"),
            LocalizedDescription("strPropertyDescriptionRDGatewayDomain")>
        Public Property RDGatewayDomain() As String
            Get
                Return GetInheritedPropertyValue("RDGatewayDomain", _rdGatewayDomain).Trim()
            End Get
            Set(ByVal value As String)
                _rdGatewayDomain = value.Trim()
            End Set
        End Property
#End Region
#Region "Appearance"
        Private _resolution As Protocol.RDP.RDPResolutions = Tools.Misc.StringToEnum(GetType(Protocol.RDP.RDPResolutions), My.Settings.ConDefaultResolution)
        <LocalizedCategory("strCategoryAppearance", 5),
            LocalizedDisplayName("strPropertyNameResolution"),
            LocalizedDescription("strPropertyDescriptionResolution"),
            TypeConverter(GetType(Tools.Misc.EnumTypeConverter))>
        Public Property Resolution() As Protocol.RDP.RDPResolutions
            Get
                Return GetInheritedPropertyValue("Resolution", _resolution)
            End Get
            Set(ByVal value As Protocol.RDP.RDPResolutions)
                _resolution = value
            End Set
        End Property

        Private _automaticResize As Boolean = My.Settings.ConDefaultAutomaticResize
        <LocalizedCategory("strCategoryAppearance", 5),
            LocalizedDisplayName("strPropertyNameAutomaticResize"),
            LocalizedDescription("strPropertyDescriptionAutomaticResize"),
            TypeConverter(GetType(Tools.Misc.YesNoTypeConverter))>
        Public Property AutomaticResize() As Boolean
            Get
                Return GetInheritedPropertyValue("AutomaticResize", _automaticResize)
            End Get
            Set(ByVal value As Boolean)
                _automaticResize = value
            End Set
        End Property

        Private _colors As Protocol.RDP.RDPColors = Tools.Misc.StringToEnum(GetType(Protocol.RDP.RDPColors), My.Settings.ConDefaultColors)
        <LocalizedCategory("strCategoryAppearance", 5),
            LocalizedDisplayName("strPropertyNameColors"),
            LocalizedDescription("strPropertyDescriptionColors"),
            TypeConverter(GetType(Tools.Misc.EnumTypeConverter))>
        Public Property Colors() As Protocol.RDP.RDPColors
            Get
                Return GetInheritedPropertyValue("Colors", _colors)
            End Get
            Set(ByVal value As Protocol.RDP.RDPColors)
                _colors = value
            End Set
        End Property

        Private _cacheBitmaps As Boolean = My.Settings.ConDefaultCacheBitmaps
        <LocalizedCategory("strCategoryAppearance", 5),
            LocalizedDisplayName("strPropertyNameCacheBitmaps"),
            LocalizedDescription("strPropertyDescriptionCacheBitmaps"),
            TypeConverter(GetType(Tools.Misc.YesNoTypeConverter))>
        Public Property CacheBitmaps() As Boolean
            Get
                Return GetInheritedPropertyValue("CacheBitmaps", _cacheBitmaps)
            End Get
            Set(ByVal value As Boolean)
                _cacheBitmaps = value
            End Set
        End Property

        Private _displayWallpaper As Boolean = My.Settings.ConDefaultDisplayWallpaper
        <LocalizedCategory("strCategoryAppearance", 5),
            LocalizedDisplayName("strPropertyNameDisplayWallpaper"),
            LocalizedDescription("strPropertyDescriptionDisplayWallpaper"),
            TypeConverter(GetType(Tools.Misc.YesNoTypeConverter))>
        Public Property DisplayWallpaper() As Boolean
            Get
                Return GetInheritedPropertyValue("DisplayWallpaper", _displayWallpaper)
            End Get
            Set(ByVal value As Boolean)
                _displayWallpaper = value
            End Set
        End Property

        Private _displayThemes As Boolean = My.Settings.ConDefaultDisplayThemes
        <LocalizedCategory("strCategoryAppearance", 5),
            LocalizedDisplayName("strPropertyNameDisplayThemes"),
            LocalizedDescription("strPropertyDescriptionDisplayThemes"),
            TypeConverter(GetType(Tools.Misc.YesNoTypeConverter))>
        Public Property DisplayThemes() As Boolean
            Get
                Return GetInheritedPropertyValue("DisplayThemes", _displayThemes)
            End Get
            Set(ByVal value As Boolean)
                _displayThemes = value
            End Set
        End Property

        Private _enableFontSmoothing As Boolean = My.Settings.ConDefaultEnableFontSmoothing
        <LocalizedCategory("strCategoryAppearance", 5),
            LocalizedDisplayName("strPropertyNameEnableFontSmoothing"),
            LocalizedDescription("strPropertyDescriptionEnableFontSmoothing"),
            TypeConverter(GetType(Tools.Misc.YesNoTypeConverter))>
        Public Property EnableFontSmoothing() As Boolean
            Get
                Return GetInheritedPropertyValue("EnableFontSmoothing", _enableFontSmoothing)
            End Get
            Set(ByVal value As Boolean)
                _enableFontSmoothing = value
            End Set
        End Property

        Private _enableDesktopComposition As Boolean = My.Settings.ConDefaultEnableDesktopComposition
        <LocalizedCategory("strCategoryAppearance", 5),
            LocalizedDisplayName("strPropertyNameEnableDesktopComposition"),
            LocalizedDescription("strPropertyDescriptionEnableDesktopComposition"),
            TypeConverter(GetType(Tools.Misc.YesNoTypeConverter))>
        Public Property EnableDesktopComposition() As Boolean
            Get
                Return GetInheritedPropertyValue("EnableDesktopComposition", _enableDesktopComposition)
            End Get
            Set(ByVal value As Boolean)
                _enableDesktopComposition = value
            End Set
        End Property
#End Region
#Region "Redirect"
        Private _redirectKeys As Boolean = My.Settings.ConDefaultRedirectKeys
        <LocalizedCategory("strCategoryRedirect", 6),
            LocalizedDisplayName("strPropertyNameRedirectKeys"),
            LocalizedDescription("strPropertyDescriptionRedirectKeys"),
            TypeConverter(GetType(Tools.Misc.YesNoTypeConverter))>
        Public Property RedirectKeys() As Boolean
            Get
                Return GetInheritedPropertyValue("RedirectKeys", _redirectKeys)
            End Get
            Set(ByVal value As Boolean)
                _redirectKeys = value
            End Set
        End Property

        Private _redirectDiskDrives As Boolean = My.Settings.ConDefaultRedirectDiskDrives
        <LocalizedCategory("strCategoryRedirect", 6),
            LocalizedDisplayName("strPropertyNameRedirectDrives"),
            LocalizedDescription("strPropertyDescriptionRedirectDrives"),
            TypeConverter(GetType(Tools.Misc.YesNoTypeConverter))>
        Public Property RedirectDiskDrives() As Boolean
            Get
                Return GetInheritedPropertyValue("RedirectDiskDrives", _redirectDiskDrives)
            End Get
            Set(ByVal value As Boolean)
                _redirectDiskDrives = value
            End Set
        End Property

        Private _redirectPrinters As Boolean = My.Settings.ConDefaultRedirectPrinters
        <LocalizedCategory("strCategoryRedirect", 6),
            LocalizedDisplayName("strPropertyNameRedirectPrinters"),
            LocalizedDescription("strPropertyDescriptionRedirectPrinters"),
            TypeConverter(GetType(Tools.Misc.YesNoTypeConverter))>
        Public Property RedirectPrinters() As Boolean
            Get
                Return GetInheritedPropertyValue("RedirectPrinters", _redirectPrinters)
            End Get
            Set(ByVal value As Boolean)
                _redirectPrinters = value
            End Set
        End Property

        Private _redirectPorts As Boolean = My.Settings.ConDefaultRedirectPorts
        <LocalizedCategory("strCategoryRedirect", 6),
            LocalizedDisplayName("strPropertyNameRedirectPorts"),
            LocalizedDescription("strPropertyDescriptionRedirectPorts"),
            TypeConverter(GetType(Tools.Misc.YesNoTypeConverter))>
        Public Property RedirectPorts() As Boolean
            Get
                Return GetInheritedPropertyValue("RedirectPorts", _redirectPorts)
            End Get
            Set(ByVal value As Boolean)
                _redirectPorts = value
            End Set
        End Property

        Private _redirectSmartCards As Boolean = My.Settings.ConDefaultRedirectSmartCards
        <LocalizedCategory("strCategoryRedirect", 6),
            LocalizedDisplayName("strPropertyNameRedirectSmartCards"),
            LocalizedDescription("strPropertyDescriptionRedirectSmartCards"),
            TypeConverter(GetType(Tools.Misc.YesNoTypeConverter))>
        Public Property RedirectSmartCards() As Boolean
            Get
                Return GetInheritedPropertyValue("RedirectSmartCards", _redirectSmartCards)
            End Get
            Set(ByVal value As Boolean)
                _redirectSmartCards = value
            End Set
        End Property

        Private _redirectSound As Protocol.RDP.RDPSounds = Tools.Misc.StringToEnum(GetType(Protocol.RDP.RDPSounds), My.Settings.ConDefaultRedirectSound)
        <LocalizedCategory("strCategoryRedirect", 6),
            LocalizedDisplayName("strPropertyNameRedirectSounds"),
            LocalizedDescription("strPropertyDescriptionRedirectSounds"),
            TypeConverter(GetType(Tools.Misc.EnumTypeConverter))>
        Public Property RedirectSound() As Protocol.RDP.RDPSounds
            Get
                Return GetInheritedPropertyValue("RedirectSound", _redirectSound)
            End Get
            Set(ByVal value As Protocol.RDP.RDPSounds)
                _redirectSound = value
            End Set
        End Property
#End Region
#Region "Misc"
        Private _preExtApp As String = My.Settings.ConDefaultPreExtApp
        <LocalizedCategory("strCategoryMiscellaneous", 7),
            LocalizedDisplayName("strPropertyNameExternalToolBefore"),
            LocalizedDescription("strPropertyDescriptionExternalToolBefore"),
            TypeConverter(GetType(Tools.ExternalToolsTypeConverter))>
        Public Overridable Property PreExtApp() As String
            Get
                Return GetInheritedPropertyValue("PreExtApp", _preExtApp)
            End Get
            Set(ByVal value As String)
                _preExtApp = value
            End Set
        End Property

        Private _postExtApp As String = My.Settings.ConDefaultPostExtApp
        <LocalizedCategory("strCategoryMiscellaneous", 7),
            LocalizedDisplayName("strPropertyNameExternalToolAfter"),
            LocalizedDescription("strPropertyDescriptionExternalToolAfter"),
            TypeConverter(GetType(Tools.ExternalToolsTypeConverter))>
        Public Overridable Property PostExtApp() As String
            Get
                Return GetInheritedPropertyValue("PostExtApp", _postExtApp)
            End Get
            Set(ByVal value As String)
                _postExtApp = value
            End Set
        End Property

        Private _macAddress As String = My.Settings.ConDefaultMacAddress
        <LocalizedCategory("strCategoryMiscellaneous", 7),
            LocalizedDisplayName("strPropertyNameMACAddress"),
            LocalizedDescription("strPropertyDescriptionMACAddress")>
        Public Overridable Property MacAddress() As String
            Get
                Return GetInheritedPropertyValue("MacAddress", _macAddress)
            End Get
            Set(ByVal value As String)
                _macAddress = value
            End Set
        End Property

        Private _userField As String = My.Settings.ConDefaultUserField
        <LocalizedCategory("strCategoryMiscellaneous", 7),
            LocalizedDisplayName("strPropertyNameUser1"),
            LocalizedDescription("strPropertyDescriptionUser1")>
        Public Overridable Property UserField() As String
            Get
                Return GetInheritedPropertyValue("UserField", _userField)
            End Get
            Set(ByVal value As String)
                _userField = value
            End Set
        End Property
#End Region
#Region "VNC"
        Private _vncCompression As Protocol.VNC.Compression = Tools.Misc.StringToEnum(GetType(Protocol.VNC.Compression), My.Settings.ConDefaultVNCCompression)
        <LocalizedCategory("strCategoryAppearance", 5),
           Browsable(False),
            LocalizedDisplayName("strPropertyNameCompression"),
            LocalizedDescription("strPropertyDescriptionCompression"),
           TypeConverter(GetType(Tools.Misc.EnumTypeConverter))>
        Public Property VNCCompression() As Protocol.VNC.Compression
            Get
                Return GetInheritedPropertyValue("VNCCompression", _vncCompression)
            End Get
            Set(ByVal value As Protocol.VNC.Compression)
                _vncCompression = value
            End Set
        End Property

        Private _vncEncoding As Protocol.VNC.Encoding = Tools.Misc.StringToEnum(GetType(Protocol.VNC.Encoding), My.Settings.ConDefaultVNCEncoding)
        <LocalizedCategory("strCategoryAppearance", 5),
            Browsable(False),
            LocalizedDisplayName("strPropertyNameEncoding"),
            LocalizedDescription("strPropertyDescriptionEncoding"),
            TypeConverter(GetType(Tools.Misc.EnumTypeConverter))>
        Public Property VNCEncoding() As Protocol.VNC.Encoding
            Get
                Return GetInheritedPropertyValue("VNCEncoding", _vncEncoding)
            End Get
            Set(ByVal value As Protocol.VNC.Encoding)
                _vncEncoding = value
            End Set
        End Property


        Private _vncAuthMode As Protocol.VNC.AuthMode = Tools.Misc.StringToEnum(GetType(Protocol.VNC.AuthMode), My.Settings.ConDefaultVNCAuthMode)
        <LocalizedCategory("strCategoryConnection", 2),
            Browsable(False),
            LocalizedDisplayName("strPropertyNameAuthenticationMode"),
            LocalizedDescription("strPropertyDescriptionAuthenticationMode"),
            TypeConverter(GetType(Tools.Misc.EnumTypeConverter))>
        Public Property VNCAuthMode() As Protocol.VNC.AuthMode
            Get
                Return GetInheritedPropertyValue("VNCAuthMode", _vncAuthMode)
            End Get
            Set(ByVal value As Protocol.VNC.AuthMode)
                _vncAuthMode = value
            End Set
        End Property

        Private _vncProxyType As Protocol.VNC.ProxyType = Tools.Misc.StringToEnum(GetType(Protocol.VNC.ProxyType), My.Settings.ConDefaultVNCProxyType)
        <LocalizedCategory("strCategoryMiscellaneous", 7),
           Browsable(False),
            LocalizedDisplayName("strPropertyNameVNCProxyType"),
            LocalizedDescription("strPropertyDescriptionVNCProxyType"),
           TypeConverter(GetType(Tools.Misc.EnumTypeConverter))>
        Public Property VNCProxyType() As Protocol.VNC.ProxyType
            Get
                Return GetInheritedPropertyValue("VNCProxyType", _vncProxyType)
            End Get
            Set(ByVal value As Protocol.VNC.ProxyType)
                _vncProxyType = value
            End Set
        End Property

        Private _vncProxyIP As String = My.Settings.ConDefaultVNCProxyIP
        <LocalizedCategory("strCategoryMiscellaneous", 7),
            Browsable(False),
            LocalizedDisplayName("strPropertyNameVNCProxyAddress"),
            LocalizedDescription("strPropertyDescriptionVNCProxyAddress")>
        Public Property VNCProxyIP() As String
            Get
                Return GetInheritedPropertyValue("VNCProxyIP", _vncProxyIP)
            End Get
            Set(ByVal value As String)
                _vncProxyIP = value
            End Set
        End Property

        Private _vncProxyPort As Integer = My.Settings.ConDefaultVNCProxyPort
        <LocalizedCategory("strCategoryMiscellaneous", 7),
            Browsable(False),
            LocalizedDisplayName("strPropertyNameVNCProxyPort"),
            LocalizedDescription("strPropertyDescriptionVNCProxyPort")>
        Public Property VNCProxyPort() As Integer
            Get
                Return GetInheritedPropertyValue("VNCProxyPort", _vncProxyPort)
            End Get
            Set(ByVal value As Integer)
                _vncProxyPort = value
            End Set
        End Property

        Private _vncProxyUsername As String = My.Settings.ConDefaultVNCProxyUsername
        <LocalizedCategory("strCategoryMiscellaneous", 7),
            Browsable(False),
            LocalizedDisplayName("strPropertyNameVNCProxyUsername"),
            LocalizedDescription("strPropertyDescriptionVNCProxyUsername")>
        Public Property VNCProxyUsername() As String
            Get
                Return GetInheritedPropertyValue("VNCProxyUsername", _vncProxyUsername)
            End Get
            Set(ByVal value As String)
                _vncProxyUsername = value
            End Set
        End Property

        Private _vncProxyPassword As String = My.Settings.ConDefaultVNCProxyPassword
        <LocalizedCategory("strCategoryMiscellaneous", 7), _
            Browsable(False), _
            LocalizedDisplayName("strPropertyNameVNCProxyPassword"), _
            LocalizedDescription("strPropertyDescriptionVNCProxyPassword"), _
            PasswordPropertyText(True)> _
        Public Property VNCProxyPassword() As String
            Get
                Return GetInheritedPropertyValue("VNCProxyPassword", _vncProxyPassword)
            End Get
            Set(ByVal value As String)
                _vncProxyPassword = value
            End Set
        End Property

        Private _vncColors As Protocol.VNC.Colors = Tools.Misc.StringToEnum(GetType(Protocol.VNC.Colors), My.Settings.ConDefaultVNCColors)
        <LocalizedCategory("strCategoryAppearance", 5), _
            Browsable(False), _
            LocalizedDisplayName("strPropertyNameColors"), _
            LocalizedDescription("strPropertyDescriptionColors"), _
            TypeConverter(GetType(Tools.Misc.EnumTypeConverter))> _
        Public Property VNCColors() As Protocol.VNC.Colors
            Get
                Return GetInheritedPropertyValue("VNCColors", _vncColors)
            End Get
            Set(ByVal value As Protocol.VNC.Colors)
                _vncColors = value
            End Set
        End Property

        Private _vncSmartSizeMode As Protocol.VNC.SmartSizeMode = Tools.Misc.StringToEnum(GetType(Protocol.VNC.SmartSizeMode), My.Settings.ConDefaultVNCSmartSizeMode)
        <LocalizedCategory("strCategoryAppearance", 5), _
            LocalizedDisplayName("strPropertyNameSmartSizeMode"), _
            LocalizedDescription("strPropertyDescriptionSmartSizeMode"), _
            TypeConverter(GetType(Tools.Misc.EnumTypeConverter))> _
        Public Property VNCSmartSizeMode() As Protocol.VNC.SmartSizeMode
            Get
                Return GetInheritedPropertyValue("VNCSmartSizeMode", _vncSmartSizeMode)
            End Get
            Set(ByVal value As Protocol.VNC.SmartSizeMode)
                _vncSmartSizeMode = value
            End Set
        End Property

        Private _vncViewOnly As Boolean = My.Settings.ConDefaultVNCViewOnly
        <LocalizedCategory("strCategoryAppearance", 5), _
            LocalizedDisplayName("strPropertyNameViewOnly"), _
            LocalizedDescription("strPropertyDescriptionViewOnly"), _
            TypeConverter(GetType(Tools.Misc.YesNoTypeConverter))> _
        Public Property VNCViewOnly() As Boolean
            Get
                Return GetInheritedPropertyValue("VNCViewOnly", _vncViewOnly)
            End Get
            Set(ByVal value As Boolean)
                _vncViewOnly = value
            End Set
        End Property
#End Region

        <Browsable(False)> _
        Public Property Inherit() As New Inheritance(Me)

        <Browsable(False)> _
        Public Property OpenConnections() As New Protocol.List

        <Browsable(False)> _
        Public Property IsContainer() As Boolean = False

        <Browsable(False)> _
        Public Property IsDefault() As Boolean = False

        <Browsable(False)> _
        Public Property Parent() As Container.Info

        <Browsable(False)> _
        Public Property PositionID() As Integer = 0

        <Browsable(False)> _
        Public Property ConstantID() As String = Tools.Misc.CreateConstantID

        <Browsable(False)> _
        Public Property TreeNode() As TreeNode

        <Browsable(False)> _
        Public Property IsQuickConnect() As Boolean = False

        <Browsable(False)> _
        Public Property PleaseConnect() As Boolean = False
#End Region

#Region "Constructors"
        Public Sub New()
            SetDefaults()
        End Sub

        Public Sub New(ByVal parent As Container.Info)
            SetDefaults()
            IsContainer = True
            Me.Parent = Parent
        End Sub
#End Region

#Region "Public Methods"
        Public Function Copy() As Info
            Dim newConnectionInfo As Info = MemberwiseClone()
            newConnectionInfo.ConstantID = Tools.Misc.CreateConstantID()
            newConnectionInfo._OpenConnections = New Protocol.List
            Return newConnectionInfo
        End Function

        Public Sub SetDefaults()
            If Port = 0 Then SetDefaultPort()
        End Sub

        Public Function GetDefaultPort() As Integer
            Return GetDefaultPort(Protocol)
        End Function

        Public Sub SetDefaultPort()
            Port = GetDefaultPort()
        End Sub
#End Region

#Region "Public Enumerations"
        <Flags()>
        Public Enum Force
            None = 0
            UseConsoleSession = 1
            Fullscreen = 2
            DoNotJump = 4
            OverridePanel = 8
            DontUseConsoleSession = 16
            NoCredentials = 32
        End Enum
#End Region

#Region "Private Methods"
        Private Function GetInheritedPropertyValue(Of TPropertyType)(ByVal propertyName As String, ByVal value As TPropertyType) As TPropertyType
            Dim inheritType As Type = Inherit.GetType()
            Dim inheritPropertyInfo As PropertyInfo = inheritType.GetProperty(propertyName)
            Dim inheritPropertyValue As Boolean = inheritPropertyInfo.GetValue(Inherit, BindingFlags.GetProperty, Nothing, Nothing, Nothing)

            If inheritPropertyValue And Parent IsNot Nothing Then
                Dim parentConnectionInfo As Info
                If IsContainer Then
                    parentConnectionInfo = Parent.Parent.ConnectionInfo
                Else
                    parentConnectionInfo = Parent.ConnectionInfo
                End If

                Dim connectionInfoType As Type = parentConnectionInfo.GetType()
                Dim parentPropertyInfo As PropertyInfo = connectionInfoType.GetProperty(propertyName)
                Dim parentPropertyValue As TPropertyType = parentPropertyInfo.GetValue(parentConnectionInfo, BindingFlags.GetProperty, Nothing, Nothing, Nothing)

                Return parentPropertyValue
            Else
                Return value
            End If
        End Function

        Private Shared Function GetDefaultPort(ByVal protocol As Protocol.Protocols) As Integer
            Try
                Select Case protocol
                    Case Connection.Protocol.Protocols.RDP
                        Return Connection.Protocol.RDP.Defaults.Port
                    Case Connection.Protocol.Protocols.VNC
                        Return Connection.Protocol.VNC.Defaults.Port
                    Case Connection.Protocol.Protocols.SSH1
                        Return Connection.Protocol.SSH1.Defaults.Port
                    Case Connection.Protocol.Protocols.SSH2
                        Return Connection.Protocol.SSH2.Defaults.Port
                    Case Connection.Protocol.Protocols.Telnet
                        Return Connection.Protocol.Telnet.Defaults.Port
                    Case Connection.Protocol.Protocols.Rlogin
                        Return Connection.Protocol.Rlogin.Defaults.Port
                    Case Connection.Protocol.Protocols.RAW
                        Return Connection.Protocol.RAW.Defaults.Port
                    Case Connection.Protocol.Protocols.HTTP
                        Return Connection.Protocol.HTTP.Defaults.Port
                    Case Connection.Protocol.Protocols.HTTPS
                        Return Connection.Protocol.HTTPS.Defaults.Port
                    Case Connection.Protocol.Protocols.IntApp
                        Return Connection.Protocol.IntegratedProgram.Defaults.Port
                End Select
            Catch ex As Exception
                MessageCollector.AddExceptionMessage(My.Language.strConnectionSetDefaultPortFailed, ex, Messages.MessageClass.ErrorMsg)
            End Try
        End Function
#End Region
    End Class
End Namespace