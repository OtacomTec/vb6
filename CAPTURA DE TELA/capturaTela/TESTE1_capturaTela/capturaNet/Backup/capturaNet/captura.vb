Module captura

    Private Declare Function CreateDC _
      Lib "GDI32" Alias "CreateDCA" ( _
      ByVal lpDriverName As String, _
      ByVal lpDeviceName As String, _
      ByVal lpOutput As String, _
      ByVal lpInitData As String _
      ) As IntPtr

    Private Declare Function CreateCompatibleDC _
       Lib "GDI32" (ByVal hDC As IntPtr) As IntPtr

    Private Declare Function CreateCompatibleBitmap _
       Lib "GDI32" ( _
       ByVal hDC As IntPtr, _
       ByVal nWidth As Integer, _
       ByVal nHeight As Integer _
       ) As IntPtr

    Private Declare Function SelectObject _
       Lib "GDI32" ( _
       ByVal hDC As IntPtr, _
       ByVal hObject As IntPtr _
       ) As IntPtr

    Private Declare Function BitBlt _
       Lib "GDI32" ( _
       ByVal srchDC As IntPtr, _
       ByVal srcX As Integer, _
       ByVal srcY As Integer, _
       ByVal srcW As Integer, _
       ByVal srcH As Integer, _
       ByVal desthDC As IntPtr, _
       ByVal destX As Integer, _
       ByVal destY As Integer, _
       ByVal op As Integer _
       ) As Integer

    Private Declare Function DeleteDC _
       Lib "GDI32" (ByVal hDC As IntPtr) As Integer

    Private Declare Function DeleteObject _
       Lib "GDI32" (ByVal hObj As IntPtr) As Integer

    Const SRCCOPY As Integer = &HCC0020


    Public Function capturaTela() As Bitmap
        ' ----- pega uma imagen da tela
        Dim screenHandle As IntPtr
        Dim canvasHandle As IntPtr
        Dim screenBitmap As IntPtr
        Dim previousObject As IntPtr
        Dim resultCode As Integer
        Dim screenShot As Bitmap

        ' ----- Obtém uma referencia para o display.
        screenHandle = CreateDC("DISPLAY", "", "", "")

        ' ----- Crie um canvas que vai servir como uma exibição
        canvasHandle = CreateCompatibleDC(screenHandle)

        ' ----- Crie um bitmap que será tratado como a imagem da tela
        screenBitmap = CreateCompatibleBitmap(screenHandle, _
           Screen.PrimaryScreen.Bounds.Width, _
           Screen.PrimaryScreen.Bounds.Height)

        ' ----- Copie a imagem da tela para Canvas
        previousObject = SelectObject(canvasHandle, _
           screenBitmap)
        resultCode = BitBlt(canvasHandle, 0, 0, _
           Screen.PrimaryScreen.Bounds.Width, _
           Screen.PrimaryScreen.Bounds.Height, _
           screenHandle, 0, 0, SRCCOPY)
        screenBitmap = SelectObject(canvasHandle, _
           previousObject)

        ' ----- encerra
        resultCode = DeleteDC(screenHandle)
        resultCode = DeleteDC(canvasHandle)

        ' ----- Copia a imagem para um bitmap .NET.
        screenShot = Image.FromHbitmap(screenBitmap)
        DeleteObject(screenBitmap)

        ' ----- encerra
        Return screenShot
    End Function

End Module
