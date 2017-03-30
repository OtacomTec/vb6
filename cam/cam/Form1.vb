Imports System.Runtime.InteropServices

Public Class Form1

    Const WM_CAP_START = &H400S
    Const WS_CHILD = &H40000000
    Const WS_VISIBLE = &H10000000

    Const WM_CAP_DRIVER_CONNECT = WM_CAP_START + 10
    Const WM_CAP_DRIVER_DISCONNECT = WM_CAP_START + 11
    Const WM_CAP_EDIT_COPY = WM_CAP_START + 30
    Const WM_CAP_SEQUENCE = WM_CAP_START + 62
    Const WM_CAP_FILE_SAVEAS = WM_CAP_START + 23
    Const WM_CAP_DLG_VIDEOFORMAT = WM_CAP_START + 41
    Const WM_CAP_DLG_VIDEOSOURCE = WM_CAP_START + 42
    Const WM_CAP_SET_SCALE = WM_CAP_START + 53
    Const WM_CAP_SET_PREVIEWRATE = WM_CAP_START + 52
    Const WM_CAP_SET_PREVIEW = WM_CAP_START + 50
    Const WM_CAP_FILE_SAVEDIB = WM_CAP_START + 25

    Const SWP_NOMOVE = &H2S
    Const SWP_NOSIZE = 1
    Const SWP_NOZORDER = &H4S
    Const HWND_BOTTOM = 1

    '--The capGetDriverDescription function retrieves the version 
    ' description of the capture driver--
    Declare Function capGetDriverDescriptionA Lib "avicap32.dll" _
       (ByVal wDriverIndex As Short, _
        ByVal lpszName As String, ByVal cbName As Integer, _
        ByVal lpszVer As String, _
        ByVal cbVer As Integer) As Boolean
    '--The capCreateCaptureWindow function creates a capture window--
    Declare Function capCreateCaptureWindowA Lib "avicap32.dll" _
       (ByVal lpszWindowName As String, ByVal dwStyle As Integer, _
        ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, _
        ByVal nHeight As Short, ByVal hWnd As Integer, _
        ByVal nID As Integer) As Integer

    '--This function sends the specified message to a window or windows--
    Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
       (ByVal hwnd As Integer, ByVal Msg As Integer, _
        ByVal wParam As Integer, _
       <MarshalAs(UnmanagedType.AsAny)> ByVal lParam As Object) As Integer

    '--Sets the position of the window relative to the screen buffer--
    Declare Function SetWindowPos Lib "user32" Alias "SetWindowPos" _
       (ByVal hwnd As Integer, _
        ByVal hWndInsertAfter As Integer, ByVal x As Integer, _
        ByVal y As Integer, _
        ByVal cx As Integer, ByVal cy As Integer, _
        ByVal wFlags As Integer) As Integer

    '--This function destroys the specified window--
    Declare Function DestroyWindow Lib "user32" _
       (ByVal hndw As Integer) As Boolean
    '---used to identify the video source---
    Dim VideoSource As Integer
    '---used as a window handle---
    Dim hWnd As Integer

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        btnStartRecording.Enabled = True
        btnStopRecording.Enabled = False
        '---list all the video sources---
        ListVideoSources()
    End Sub

    '---list all the various video sources---
    Private Sub ListVideoSources()
        Dim DriverName As String = Space(80)
        Dim DriverVersion As String = Space(80)
        For i As Integer = 0 To 9
            If capGetDriverDescriptionA(i, DriverName, 80, DriverVersion, 80) Then
                lstVideoSources.Items.Add(DriverName.Trim)
            End If
        Next
    End Sub

    Private Sub lstVideoSources_SelectedIndexChanged( _
       ByVal sender As System.Object, ByVal e As System.EventArgs) _
       Handles lstVideoSources.SelectedIndexChanged
        '---check which video source is selected---
        VideoSource = lstVideoSources.SelectedIndex
        '---preview the selected video source
        PreviewVideo(PictureBox1)
    End Sub

    '---preview the selected video source---
    Private Sub PreviewVideo(ByVal pbCtrl As PictureBox)
        hWnd = capCreateCaptureWindowA(VideoSource, _
            WS_VISIBLE Or WS_CHILD, 0, 0, 0, _
            0, pbCtrl.Handle.ToInt32, 0)
        If SendMessage( _
           hWnd, WM_CAP_DRIVER_CONNECT, _
           VideoSource, 0) Then
            '---set the preview scale---
            SendMessage(hWnd, WM_CAP_SET_SCALE, True, 0)
            '---set the preview rate (ms)---
            SendMessage(hWnd, WM_CAP_SET_PREVIEWRATE, 30, 0)
            '---start previewing the image---
            SendMessage(hWnd, WM_CAP_SET_PREVIEW, True, 0)
            '---resize window to fit in PictureBox control---
            SetWindowPos(hWnd, HWND_BOTTOM, 0, 0, _
               pbCtrl.Width, pbCtrl.Height, _
               SWP_NOMOVE Or SWP_NOZORDER)
        Else
            '--error connecting to video source---
            DestroyWindow(hWnd)
        End If
    End Sub
    '--disconnect from video source---
    Private Sub StopPreviewWindow()
        SendMessage(hWnd, WM_CAP_DRIVER_DISCONNECT, VideoSource, 0)
        DestroyWindow(hWnd)
    End Sub

    '---stop the preview window---
    Private Sub btnStopCamera_Click( _
       ByVal sender As System.Object, _
       ByVal e As System.EventArgs) _
       Handles btnStopCamera.Click
        StopPreviewWindow()
    End Sub

    '---Start recording the video---
    Private Sub btnStartRecording_Click( _
       ByVal sender As System.Object, _
       ByVal e As System.EventArgs) _
       Handles btnStartRecording.Click
        btnStartRecording.Enabled = False
        btnStopRecording.Enabled = True
        '---start recording---
        SendMessage(hWnd, WM_CAP_SEQUENCE, 0, 0)
    End Sub

    '---stop recording and save it on file---
    Private Sub btnStopRecording_Click( _
       ByVal sender As System.Object, _
       ByVal e As System.EventArgs) _
       Handles btnStopRecording.Click
        btnStartRecording.Enabled = True
        btnStopRecording.Enabled = False
        '---save the recording to file---
        SendMessage(hWnd, WM_CAP_FILE_SAVEAS, 0, "C:\RecordedVideo.avi")
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Dim mStatus As String
        Dim capDlgVideoSource As Boolean
        capDlgVideoSource = SendMessage(hWnd, WM_CAP_DLG_VIDEOSOURCE, 0&, 0&)
        'mStatus = "Configuração: Fonte de vídeo"
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim capDlgVideoFormat As Boolean
        capDlgVideoFormat = SendMessage(hWnd, WM_CAP_DLG_VIDEOFORMAT, 0&, 0&)
      
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        Dim capFileSaveDIB As Boolean

        Dim ArquivoFoto As String = "C:\" & Format(Now, "yyyymmdd") & " - " & Format(Now, "HHMMss") & ".png"

        capFileSaveDIB = SendMessage(hWnd, WM_CAP_FILE_SAVEDIB, 0&, ArquivoFoto)
        PictureBox2.Image = Image.FromFile(ArquivoFoto)

    End Sub
End Class
