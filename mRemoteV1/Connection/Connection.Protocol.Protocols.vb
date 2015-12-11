Imports mRemoteNG.Tools.LocalizedAttributes

Namespace Connection
    Namespace Protocol
        Public Class Converter
            Public Shared Function ProtocolToString(ByVal protocol As Protocols) As String
                Return protocol.ToString()
            End Function

            Public Shared Function StringToProtocol(ByVal protocol As String) As Protocols
                Try
                    Return [Enum].Parse(GetType(Protocols), protocol, True)
                Catch ex As Exception
                    Return Protocols.RDP
                End Try
            End Function

            Private Sub New()
            End Sub
        End Class

        Public Enum Protocols
            <LocalizedDescription("strRDP")>
            RDP = 0
            <LocalizedDescription("strVnc")>
            VNC = 1
            <LocalizedDescription("strSsh1")>
            SSH1 = 2
            <LocalizedDescription("strSsh2")>
            SSH2 = 3
            <LocalizedDescription("strTelnet")>
            Telnet = 4
            <LocalizedDescription("strRlogin")>
            Rlogin = 5
            <LocalizedDescription("strRAW")>
            RAW = 6
            <LocalizedDescription("strExtApp")>
            IntApp = 20
        End Enum
    End Namespace
End Namespace
