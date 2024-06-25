Imports EasySSH.EasySSH.EasySSHClient
Module Test
    Sub test_DownloadFile()
        Dim ServerConection As New ServerConnect
        ServerConection.ConnectionTunnel = "22"
        ServerConection.ServerIP = "0.0.0.0"
        ServerConection.ServerSSHPassword = "ABC123"
        ServerConection.HostName = "root"
        ServerConection.Connect()
        Dim result As String = ServerConection.Applications.DownloadFile("ServerFile", "DownloadFile")
        Debug.Print(result)
        ServerConection.Disconnect()
    End Sub
End Module
