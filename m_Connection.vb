Imports System.Runtime.InteropServices
Imports Renci.SshNet
Imports Renci.SshNet.Security.Org.BouncyCastle
Imports System.IO

Namespace EasySSH
    Public Class EasySSHClient
        Friend Shared o_ssh As SftpClient
        <InterfaceType(ComInterfaceType.InterfaceIsIDispatch)> Public Interface IEasySSH
            Property ServerIP As String
            Property HostName As String
            Property ConnectionTunnel As String
            Property ServerSSHPassword As String
            Sub Connect()
            Sub Disconnect()
            Function Connect(CatchEx As Boolean)
            Function IsActive() As Boolean
            Property Applications() As ServerApps
        End Interface
        <ClassInterface(ClassInterfaceType.None)> Public Class ServerConnect
            Implements IEasySSH
            Declare Function SetProcessWorkingSetSize Lib "kernel32.dll" (ByVal process As IntPtr, ByVal minimumWorkingSetSize As Integer, ByVal maximumWorkingSetSize As Integer) As Integer
            Friend Shared o_ServerIP As String
            Friend Shared o_HostName As String
            Friend Shared o_ConnectionTunnel As String
            Friend Shared o_ServerSSHPassword As String
            Friend Shared o_SSHClient As SftpClient
            Public Property ServerIP As String Implements IEasySSH.ServerIP
                Get
                    If o_ServerIP = "" Or o_ServerIP = Nothing Then
                        Return "0.0.0.0"
                    Else
                        Return o_ServerIP
                    End If
                End Get
                Set(value As String)
                    o_ServerIP = value
                End Set
            End Property
            Public Property HostName As String Implements IEasySSH.HostName
                Get
                    If o_HostName = "" Or o_HostName = Nothing Then
                        Return ""
                    Else
                        Return o_HostName
                    End If
                End Get
                Set(value As String)
                    o_HostName = value
                End Set
            End Property
            Public Property ConnectionTunnel As String Implements IEasySSH.ConnectionTunnel
                Get
                    If o_ConnectionTunnel = "" Or o_ConnectionTunnel = Nothing Then
                        Return ""
                    Else
                        Return o_ConnectionTunnel
                    End If
                End Get
                Set(value As String)
                    o_ConnectionTunnel = value
                End Set
            End Property
            Public Property ServerSSHPassword As String Implements IEasySSH.ServerSSHPassword
                Get
                    If o_ServerSSHPassword = "" Or o_ServerSSHPassword = Nothing Then
                        Return ""
                    Else
                        Return o_ServerSSHPassword
                    End If
                End Get
                Set(value As String)
                    o_ServerSSHPassword = value
                End Set
            End Property
            Function Connect(CatchEx As Boolean) Implements IEasySSH.Connect
                Using m_SSHClient As SftpClient = New SftpClient(New PasswordConnectionInfo(ServerIP, ConnectionTunnel, HostName, ServerSSHPassword))
                    Try
                        m_SSHClient.Connect()
                        o_SSHClient = m_SSHClient
                        Return "Connection succesfull."
                    Catch ex As Exception
                        If CatchEx = True Then
                            Return ex.ToString
                        Else
                            Return "Error connection. Check SSH informations."
                        End If
                    End Try
                End Using
            End Function
            Sub Connect() Implements IEasySSH.Connect
                Using m_SSHClient As SftpClient = New SftpClient(New PasswordConnectionInfo(ServerIP, ConnectionTunnel, HostName, ServerSSHPassword))
                    Try
                        m_SSHClient.Connect()
                        o_SSHClient = m_SSHClient
                        o_ssh = m_SSHClient
                    Catch ex As Exception
                        'Return ?
                    End Try
                End Using
            End Sub
            Sub Disconnect() Implements IEasySSH.Disconnect
                o_SSHClient.Disconnect()
                o_SSHClient.Dispose()
            End Sub
            Function IsActive() As Boolean Implements IEasySSH.IsActive
                Return o_SSHClient.IsConnected
            End Function

            Property Applications() As ServerApps Implements IEasySSH.Applications
                Get
                    Return Nothing
                End Get
                Set(value As ServerApps)
                End Set
            End Property
        End Class

        <InterfaceType(ComInterfaceType.InterfaceIsIDispatch)> Public Interface IServerApps
            Function DownloadFile(ServerPath As String, DownloadPath As String) As String
        End Interface
        <ClassInterface(ClassInterfaceType.None)> Public Class ServerApps
            Implements IServerApps
            Declare Function SetProcessWorkingSetSize Lib "kernel32.dll" (ByVal process As IntPtr, ByVal minimumWorkingSetSize As Integer, ByVal maximumWorkingSetSize As Integer) As Integer

            Function DownloadFile(ServerPath As String, DownloadPath As String) As String Implements IServerApps.DownloadFile
                If o_ssh.IsConnected = True Then
                    If o_ssh.Exists(ServerPath) = True Then
                        If File.Exists(DownloadPath) = False Then
                            Try
                                Dim targetStream
                                targetStream = New IO.FileStream(DownloadPath, IO.FileMode.Append)
                                o_ssh.DownloadFile(ServerPath, targetStream)
                                targetStream.Close()
                                targetStream.Dispose()
                                targetStream = Nothing
                                Return "File downloaded."
                            Catch ex As Exception
                                Return "Check download folder, may be not exists."
                            End Try
                        Else
                            Return "Local file was exists. Please delete and try again."
                        End If
                    Else
                        Return "File not exists."
                    End If
                Else
                    Return "SSH not connected."
                End If
                Return DownloadFile
            End Function
        End Class
    End Class
End Namespace
