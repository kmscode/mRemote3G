﻿Imports System.ComponentModel
Imports System.IO
Imports mRemote3G.App
Imports mRemote3G.Connection
Imports mRemote3G.Connection.Protocol
Imports mRemote3G.Messages
Imports mRemote3G.My.Resources

Namespace Tools
    Public Class ExternalTool

#Region "Public Properties"

        Public Property DisplayName As String
        Public Property FileName As String
        Public Property WaitForExit As Boolean
        Public Property Arguments As String
        Public Property TryIntegrate As Boolean
        Public Property ConnectionInfo As Info

        Public ReadOnly Property Icon As Drawing.Icon
            Get
                If File.Exists(FileName) Then
                    Return Misc.GetIconFromFile(FileName)
                Else
                    Return Nothing
                End If
            End Get
        End Property

        Public ReadOnly Property Image As Image
            Get
                Try
                    Dim i As Image
                    If Icon IsNot Nothing Then
                        i = Icon.ToBitmap
                        Return i
                    Else
                        Return mRemote_Icon.ToBitmap
                    End If
                Catch ex As Exception
                    Runtime.MessageCollector.AddMessage(MessageClass.WarningMsg,
                                                        "Could not convert image to Bitmap (Tools.ExternalTool) - using default icon. Exception:" &
                                                        vbNewLine & ex.ToString(), True)
                    Return mRemote_Icon.ToBitmap
                End Try
            End Get
        End Property

#End Region

#Region "Constructors"

        Public Sub New(Optional ByVal displayName As String = "", Optional ByVal fileName As String = "",
                       Optional ByVal arguments As String = "")
            Me.DisplayName = displayName
            Me.FileName = fileName
            Me.Arguments = arguments
        End Sub

#End Region

#Region "Public Methods"
        ' Start external app
        Public Sub Start(Optional ByVal startConnectionInfo As Info = Nothing)
            Try
                If String.IsNullOrEmpty(_FileName) Then
                    ' No need to thow an exception, just log/show an error and return.
                    Runtime.MessageCollector.AddMessage(MessageClass.ErrorMsg,
                                                        "ExternalApp.Start() failed: FileName cannot be blank.", False)
                    Return
                End If

                ConnectionInfo = startConnectionInfo

                If TryIntegrate Then
                    StartIntegrated()
                    Return
                End If

                Dim process As New Process()
                With process.StartInfo
                    .UseShellExecute = True
                    .FileName = ParseArguments(FileName)
                    .Arguments = ParseArguments(Arguments)
                End With

                process.Start()

                If WaitForExit Then process.WaitForExit()
            Catch ex As Exception
                Runtime.MessageCollector.AddExceptionMessage("ExternalApp.Start() failed.", ex)
            End Try
        End Sub

        ' Start external app integrated
        Public Sub StartIntegrated()
            Try
                Dim newConnectionInfo As Info
                If ConnectionInfo Is Nothing Then
                    newConnectionInfo = New Info
                Else
                    newConnectionInfo = ConnectionInfo.Copy()
                End If

                With newConnectionInfo
                    .Protocol = Protocols.IntApp
                    .ExtApp = DisplayName
                    .Name = DisplayName
                    .Panel = Language.Language.strMenuExternalTools
                End With

                Runtime.OpenConnection(newConnectionInfo)
            Catch ex As Exception
                Runtime.MessageCollector.AddExceptionMessage("ExternalApp.StartIntegrated() failed.", ex, , True)
            End Try
        End Sub

        Private Enum EscapeType
            All
            ShellMetacharacters
            None
        End Enum

        Private Structure Replacement
            Public Sub New(start As Integer, length As Integer, value As String)
                Me.Start = start
                Me.Length = length
                Me.Value = value
            End Sub

            Public Property Start As Integer
            Public Property Length As Integer
            Public Property Value As String
        End Structure

        Public Function ParseArguments(input As String) As String
            Dim index = 0
            Dim replacements As New List(Of Replacement)

            Do
                Dim tokenStart As Integer = input.IndexOf("%", index, StringComparison.InvariantCulture)
                If tokenStart = - 1 Then Exit Do

                Dim tokenEnd As Integer = input.IndexOf("%", tokenStart + 1, StringComparison.InvariantCulture)
                If tokenEnd = - 1 Then Exit Do

                Dim tokenLength As Integer = tokenEnd - tokenStart + 1

                Dim variableNameStart As Integer = tokenStart + 1
                Dim variableNameLength As Integer = tokenLength - 2

                Dim isEnvironmentVariable = False

                Dim variableName As String

                If tokenStart > 0 Then
                    Dim tokenStartPrefix As Char = input.Substring(tokenStart - 1, 1)
                    Dim tokenEndPrefix As Char = input.Substring(tokenEnd - 1, 1)

                    If tokenStartPrefix = "\" And tokenEndPrefix = "\" Then
                        isEnvironmentVariable = True

                        ' Add the first backslash to the token
                        tokenStart = tokenStart - 1
                        tokenLength = tokenLength + 1

                        ' Remove the last backslash from the name
                        variableNameLength = variableNameLength - 1
                    ElseIf tokenStartPrefix = "^" And tokenEndPrefix = "^" Then
                        ' Add the first caret to the token
                        tokenStart = tokenStart - 1
                        tokenLength = tokenLength + 1

                        ' Remove the last caret from the name
                        variableNameLength = variableNameLength - 1

                        variableName = input.Substring(variableNameStart, variableNameLength)
                        replacements.Add(New Replacement(tokenStart, tokenLength, String.Format("%{0}%", variableName)))

                        index = tokenEnd
                        Continue Do
                    End If
                End If

                Dim token As String = input.Substring(tokenStart, tokenLength)

                Dim escape = EscapeType.All
                Dim prefix As String = input.Substring(variableNameStart, 1)
                Select Case prefix
                    Case "-"
                        escape = EscapeType.ShellMetacharacters
                    Case "!"
                        escape = EscapeType.None
                End Select

                If Not escape = EscapeType.All Then
                    ' Remove the escape character from the name
                    variableNameStart = variableNameStart + 1
                    variableNameLength = variableNameLength - 1
                End If

                If variableNameLength = 0 Then
                    index = tokenEnd
                    Continue Do
                End If

                variableName = input.Substring(variableNameStart, variableNameLength)

                Dim replacementValue As String = token
                If Not isEnvironmentVariable Then
                    replacementValue = GetVariableReplacement(variableName, token)
                End If

                Dim haveReplacement = False

                If Not replacementValue = token Then
                    haveReplacement = True
                Else
                    replacementValue = Environment.GetEnvironmentVariable(variableName)
                    If replacementValue IsNot Nothing Then haveReplacement = True
                End If

                If haveReplacement Then
                    Dim trailing As Char
                    If tokenEnd + 2 <= input.Length Then
                        trailing = input.Substring(tokenEnd + 1, 1)
                    Else
                        trailing = String.Empty
                    End If

                    If escape = EscapeType.All Then
                        replacementValue = CommandLineArguments.EscapeBackslashes(replacementValue)
                        If trailing = """" Then
                            replacementValue = CommandLineArguments.EscapeBackslashesForTrailingQuote(replacementValue)
                        End If
                    End If

                    If escape = EscapeType.All Or escape = EscapeType.ShellMetacharacters Then
                        replacementValue = CommandLineArguments.EscapeShellMetacharacters(replacementValue)
                    End If

                    replacements.Add(New Replacement(tokenStart, tokenLength, replacementValue))
                    index = tokenEnd + 1
                Else
                    index = tokenEnd
                End If
            Loop

            Dim result As String = input

            For index = result.Length To 0 Step - 1
                For Each replacement As Replacement In replacements
                    If Not replacement.Start = index Then Continue For

                    Dim before As String = result.Substring(0, replacement.Start)
                    Dim after As String = result.Substring(replacement.Start + replacement.Length)
                    result = before & replacement.Value & after
                Next
            Next

            Return result
        End Function

#End Region

#Region "Private Methods"

        Private Function GetVariableReplacement(variable As String, original As String) As String
            Dim replacement As String
            Select Case variable.ToLowerInvariant()
                Case "name"
                    If ConnectionInfo Is Nothing Then replacement = "" Else replacement = ConnectionInfo.Name
                Case "hostname"
                    If ConnectionInfo Is Nothing Then replacement = "" Else replacement = ConnectionInfo.Hostname
                Case "port"
                    If ConnectionInfo Is Nothing Then replacement = "" Else replacement = ConnectionInfo.Port
                Case "username"
                    If ConnectionInfo Is Nothing Then replacement = "" Else replacement = ConnectionInfo.Username
                Case "password"
                    If ConnectionInfo Is Nothing Then replacement = "" Else replacement = ConnectionInfo.Password
                Case "domain"
                    If ConnectionInfo Is Nothing Then replacement = "" Else replacement = ConnectionInfo.Domain
                Case "description"
                    If ConnectionInfo Is Nothing Then replacement = "" Else replacement = ConnectionInfo.Description
                    ' ReSharper disable once StringLiteralTypo
                Case "macaddress"
                    If ConnectionInfo Is Nothing Then replacement = "" Else replacement = ConnectionInfo.MacAddress
                    ' ReSharper disable once StringLiteralTypo
                Case "userfield"
                    If ConnectionInfo Is Nothing Then replacement = "" Else replacement = ConnectionInfo.UserField
                Case Else
                    Return original
            End Select
            Return replacement
        End Function

#End Region
    End Class

    Public Class ExternalToolsTypeConverter
        Inherits StringConverter

        Public Shared ReadOnly Property ExternalTools As String()
            Get
                Dim externalToolList As New List(Of String)

                ' Add a blank entry to signify that no external tool is selected
                externalToolList.Add(String.Empty)

                For Each externalTool As ExternalTool In Runtime.ExternalTools
                    externalToolList.Add(externalTool.DisplayName)
                Next

                Return externalToolList.ToArray()
            End Get
        End Property

        Public Overloads Overrides Function GetStandardValues(context As ITypeDescriptorContext) _
            As StandardValuesCollection
            Return New StandardValuesCollection(ExternalTools)
        End Function

        Public Overloads Overrides Function GetStandardValuesExclusive(context As ITypeDescriptorContext) As Boolean
            Return True
        End Function

        Public Overloads Overrides Function GetStandardValuesSupported(context As ITypeDescriptorContext) As Boolean
            Return True
        End Function
    End Class
End Namespace