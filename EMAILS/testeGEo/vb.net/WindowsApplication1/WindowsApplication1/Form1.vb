Public Class Form1

    Private Sub WebBrowser1_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted

    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        Call IEbrowserFix()

        Me.WebBrowser1.Navigate("file:///C:/Users/Marcos/Desktop/TESTE.html")
        ' Add any initialization after the InitializeComponent() call.

    End Sub


    Private Sub IEbrowserFix()
        Try
            Dim regDM As Microsoft.Win32.RegistryKey
            Dim is64 As Boolean = Environment.Is64BitOperatingSystem
            Dim KeyPath As String = ""
            If is64 Then
                KeyPath = "SOFTWARE\Wow6432Node\Microsoft\Internet Explorer\MAIN\FeatureControl\FEATURE_BROWSER_EMULATION"
            Else
                KeyPath = "SOFTWARE\Microsoft\Internet Explorer\MAIN\FeatureControl\FEATURE_BROWSER_EMULATION"
            End If

            regDM = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(KeyPath, False)
            If regDM Is Nothing Then
                regDM = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(KeyPath)
            End If

            Dim Sleutel As Object
            If Not regDM Is Nothing Then
                Dim location As String = System.Environment.GetCommandLineArgs()(0)
                Dim appName As String = System.IO.Path.GetFileName(location)

                Sleutel = regDM.GetValue(appName)
                If Sleutel Is Nothing Then
                    'Sleutel onbekend
                    regDM = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(KeyPath, True)
                    Sleutel = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(KeyPath, Microsoft.Win32.RegistryKeyPermissionCheck.ReadWriteSubTree)

                    'What OS are we using
                    Dim OsVersion As Version = System.Environment.OSVersion.Version

                    If OsVersion.Major = 6 And OsVersion.Minor = 1 Then
                        'WIN 7
                        Sleutel.SetValue(appName, 9000, Microsoft.Win32.RegistryValueKind.DWord)
                    ElseIf OsVersion.Major = 6 And OsVersion.Minor = 2 Then
                        'WIN 8
                        Sleutel.SetValue(appName, 10000, Microsoft.Win32.RegistryValueKind.DWord)
                    ElseIf OsVersion.Major = 5 And OsVersion.Minor = 1 Then
                        'WIN xp
                        Sleutel.SetValue(appName, 8000, Microsoft.Win32.RegistryValueKind.DWord)
                    End If

                    Sleutel.Close()
                End If
            End If


        Catch ex As Exception
            MsgBox("erro" + ex.Message)
        End Try
    End Sub

End Class
