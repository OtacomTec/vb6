Attribute VB_Name = "Module1"
'-------------------------------------------------------------------------------
'GrFinger SAMPLE
'(c) 2007 - 2010 Griaule Biometrics Ltda.
'http://www.griaulebiometrics.com
'-------------------------------------------------------------------------------
'
'This sample is provided with "GrFinger Fingerprint Recognition Library" and
'can 't run without it. It's provided just as an example of using GrFinger
'Fingerprint Recognition Library and should not be used as basis for any
'commercial product.
'
'Griaule Biometrics makes no representations concerning either the merchantability
'of this software or the suitability of this sample for any particular purpose.
'
'THIS SAMPLE IS PROVIDED BY THE AUTHOR "AS IS" AND ANY EXPRESS OR
'IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES
'OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED.
'IN NO EVENT SHALL GRIAULE BE LIABLE FOR ANY DIRECT, INDIRECT,
'INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT
'NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
'DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
'THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
'(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
'THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'
'You can download the trial version of GrFinger directly from Griaule website.
'
'These notices must be retained in any copies of any part of this
'documentation and/or sample.
'
'-------------------------------------------------------------------------------
'
' -----------------------------------------------------------------------------------
' GrFinger DLL Platform Invoke Module
' -----------------------------------------------------------------------------------
' -----------------------------------------------------------------------------------
' GrFinger constants
' -----------------------------------------------------------------------------------

' Return codes
Public Const GR_OK As Integer = 0
Public Const GR_BAD_QUALITY As Integer = 0
Public Const GR_MEDIUM_QUALITY As Integer = 1
Public Const GR_HIGH_QUALITY As Integer = 2
Public Const GR_MATCH As Integer = 1
Public Const GR_NOT_MATCH As Integer = 0
Public Const GR_DEFAULT_USED As Integer = 3
' Initialization error codes
Public Const GR_ERROR_INITIALIZE_FAIL As Integer = -1
Public Const GR_ERROR_NOT_INITIALIZED As Integer = -2
Public Const GR_ERROR_FAIL_LICENSE_READ As Integer = -3
Public Const GR_ERROR_NO_VALID_LICENSE As Integer = -4
Public Const GR_ERROR_NULL_ARGUMENT As Integer = -5
Public Const GR_ERROR_FAIL As Integer = -6
Public Const GR_ERROR_ALLOC As Integer = -7
Public Const GR_ERROR_PARAMETERS As Integer = -8
' Extract and match error codes
Public Const GR_ERROR_WRONG_USE As Integer = -107
Public Const GR_ERROR_EXTRACT As Integer = -108
Public Const GR_ERROR_SIZE_OFF_RANGE As Integer = -109
Public Const GR_ERROR_RES_OFF_RANGE As Integer = -110
Public Const GR_ERROR_CONTEXT_NOT_CREATED As Integer = -111
Public Const GR_ERROR_INVALID_CONTEXT As Integer = -112
Public Const GR_ERROR_ERROR As Integer = -113
Public Const GR_ERROR_NOT_ENOUGH_SPACE As Integer = -114
' Capture error codes
Public Const GR_ERROR_CONNECT_SENSOR As Integer = -201
Public Const GR_ERROR_CAPTURING As Integer = -202
Public Const GR_ERROR_CANCEL_CAPTURING As Integer = -203
Public Const GR_ERROR_INVALID_ID_SENSOR As Integer = -204
Public Const GR_ERROR_SENSOR_NOT_CAPTURING As Integer = -205
Public Const GR_ERROR_INVALID_EXT As Integer = -206
Public Const GR_ERROR_INVALID_FILENAME As Integer = -207
Public Const GR_ERROR_INVALID_FILETYPE As Integer = -208
Public Const GR_ERROR_SENSOR As Integer = -209
' File format codes
Public Const GRCAP_IMAGE_FORMAT_BMP As Integer = 501
' Event codes
Public Const GR_PLUG As Integer = 21
Public Const GR_UNPLUG As Integer = 20
Public Const GR_FINGER_DOWN As Integer = 11
Public Const GR_FINGER_UP As Integer = 10
Public Const GR_IMAGE As Integer = 30
' Image attributes
Public Const GR_DEFAULT_RES As Integer = 500
Public Const GR_DEFAULT_DIM As Integer = 500
Public Const GR_MAX_SIZE_TEMPLATE As Integer = 10000
Public Const GR_MAX_IMAGE_WIDTH As Integer = 1280
Public Const GR_MAX_IMAGE_HEIGHT As Integer = 1280
Public Const GR_MAX_RESOLUTION As Integer = 1000 'DPI
Public Const GR_MIN_IMAGE_WIDTH As Integer = 50
Public Const GR_MIN_IMAGE_HEIGHT As Integer = 50
Public Const GR_MIN_RESOLUTION As Integer = 125 'DPI
' Matching attributes
Public Const GR_MAX_THRESHOLD As Integer = 200
Public Const GR_MIN_THRESHOLD As Integer = 10
Public Const GR_VERYLOW_FRR As Integer = 30 'FAR 1 IN 1000
Public Const GR_LOW_FRR As Integer = 45 'FAR 1 IN 10000
Public Const GR_LOW_FAR As Integer = 60 'FAR 1 IN 30000
Public Const GR_VERYLOW_FAR As Integer = 80 'FAR 1 IN 300000
Public Const GR_ROT_MIN As Integer = 0
Public Const GR_ROT_MAX As Integer = 180
' Context codes
Public Const GR_DEFAULT_CONTEXT As Integer = 0
Public Const GR_NO_CONTEXT As Integer = -1
' BiometricDisplay color codes
Public Const GR_IMAGE_NO_COLOR As Integer = &H1FFFFFFF
' Version codes
Public Const GRFINGER_FULL As Integer = 1
Public Const GRFINGER_LIGHT As Integer = 2
Public Const GRFINGER_FREE As Integer = 3

' -----------------------------------------------------------------------------------
' Callbacks
' -----------------------------------------------------------------------------------

' Callback functions
Public Sub GRCAP_STATUS_EVENT_PROC(ByVal idSensor As String, ByVal evt As Integer)

End Sub
Public Sub GRCAP_FINGER_EVENT_PROC(ByVal idSensor As String, ByVal evt As Integer)

End Sub
Public Sub GRCAP_IMAGE_EVENT_PROC(ByVal idSensor As String, ByVal width As Integer, ByVal height As Integer, ByVal rawImage As Integer, ByVal res As Integer)

End Sub
Public Sub GRCAP_IMAGE_EVENT_PROC(ByVal idSensor As String, ByVal width As Integer, ByVal height As Integer, ByVal rawImage As Byte, ByVal res As Integer)

End Sub

Dim ImageEventHandler As GRCAP_IMAGE_EVENT_PROC

    ' -----------------------------------------------------------------------------------
    ' Helper functions
    ' -----------------------------------------------------------------------------------

    Private Function PointerToByteArray(ByVal pointer As Integer, ByVal size As Integer) As Byte()
        Dim pbuffer As IntPtr
        pbuffer = IntPtr(pointer)
        Dim buffer(size - 1) As Byte
        'Marshal.Copy(pbuffer, buffer, 0, size)
        PointerToByteArray = buffer
    End Function

    Private Function ByteArrayToPointer(ByRef aa As Byte) As Integer
        Dim buffer As IntPtr
        buffer = Marshal.StringToHGlobalAuto(Space(Len(aa)))
        'Marshal.Copy(array, 0, buffer, array.Length)
        ByteArrayToPointer = buffer.ToInt32()
    End Function

    Private Declare Function CapInitialize Lib "grfinger.dll" Alias "_GrCapInitialize@4" (ByVal StatusEventHandler As GRCAP_STATUS_EVENT_PROC) As Integer

    Private Declare Function CapStartCapture Lib "grfinger.dll" Alias "_GrCapStartCapture@12" (ByVal idSensor As String, ByVal FingerEventHandler As GRCAP_FINGER_EVENT_PROC, ByVal ImageEventHandler As _GRCAP_IMAGE_EVENT_PROC) As Integer

    Private Declare Function _BiometricDisplay Lib "grfinger.dll" Alias "_GrBiometricDisplay@32" (ByVal tptQuery As Integer, ByVal rawImage As Integer, ByVal width As Integer, ByVal height As Integer, ByVal res As Integer, ByVal hdc As Integer, ByRef handle As Integer, ByVal matchContext As Integer) As Integer

    Private Declare Function _Extract Lib "grfinger.dll" Alias "_GrExtract@28" (ByVal rawimage As Integer, ByVal width As Integer, ByVal height As Integer, ByVal res As Integer, ByVal tpt As Integer, ByRef tptSize As Integer, ByVal context As Integer) As Integer

    Private Declare Function _Verify Lib "grfinger.dll" Alias "_GrVerify@16" (ByVal queryTemplate As Integer, ByVal referenceTemplate As Integer, ByRef verifyScore As Integer, ByVal context As Integer) As Integer

    Private Declare Function _IdentifyPrepare Lib "grfinger.dll" Alias "_GrIdentifyPrepare@8" (ByVal templateQuery As Integer, ByVal context As Integer) As Integer

    Private Declare Function _Identify Lib "grfinger.dll" Alias "_GrIdentify@12" (ByVal templateReference As Integer, ByRef identifyScore As Integer, ByVal context As Integer) As Integer

    Private Sub __GRCAP_IMAGE_EVENT_PROC(ByVal idSensor As String, ByVal width As Integer, ByVal height As Integer, ByVal rawImage As Integer, ByVal res As Integer)
        _ImageEventHandler(idSensor, width, height, PointerToByteArray(rawImage, width * height), res)
    End Function


    ' -----------------------------------------------------------------------------------
    ' Public capture functions
    ' -----------------------------------------------------------------------------------

    Public Function CapInitialize(ByVal StatusEventHandler As GRCAP_STATUS_EVENT_PROC) As Integer
        Dim _StatusEventHandler As GRCAP_STATUS_EVENT_PROC = StatusEventHandler
        CapInitialize = _CapInitialize(_StatusEventHandler)
    End Function

    Public Declare Function CapFinalize Lib "grfinger.dll" Alias "_GrCapFinalize@0" () As Integer

    Public Function CapStartCapture(ByVal idSensor As String, ByVal FingerEventHandler As GRCAP_FINGER_EVENT_PROC, ByVal ImageEventHandler As GRCAP_IMAGE_EVENT_PROC) As Integer
        Static Dim _FingerEventHandler As GRCAP_FINGER_EVENT_PROC = FingerEventHandler
        Static Dim __ImageEventHandler As _GRCAP_IMAGE_EVENT_PROC = AddressOf __GRCAP_IMAGE_EVENT_PROC
        _ImageEventHandler = ImageEventHandler
        CapStartCapture = _CapStartCapture(idSensor, _FingerEventHandler, __ImageEventHandler)
    End Function

    Public Declare Function CapStopCapture Lib "grfinger.dll" Alias "_GrCapStopCapture@4" (ByVal idSensor As String) As Integer

    Public Declare Function CapSaveRawImageToFile Lib "grfinger.dll" Alias "_GrCapSaveRawImageToFile@20" (ByVal rawImage As Integer, ByVal width As Integer, ByVal height As Integer, ByVal fileName As String, ByVal imageFormat As Integer) As Integer

    Public Function CapSaveRawImageToFile(ByVal rawImage As Byte(), ByVal width As Integer, ByVal height As Integer, ByVal fileName As String, ByVal imageFormat As Integer) As Integer
        Dim ri As Integer = ByteArrayToPointer(rawImage)
        CapSaveRawImageToFile = CapSaveRawImageToFile(ri, width, height, fileName, imageFormat)
        Marshal.FreeHGlobal(New IntPtr(ri))
    End Function

    Public Declare Function CapLoadImageFromFile Lib "grfinger.dll" Alias "_GrCapLoadImageFromFile@8" (ByVal fileName As String, ByVal res As Integer) As Integer

    Public Declare Function CapRawImageToHandle Lib "grfinger.dll" Alias "_GrCapRawImageToHandle@20" (ByVal rawImage As Integer, ByVal width As Integer, ByVal height As Integer, ByVal hdc As Integer, ByRef handle As Integer) As Integer

    Public Function CapRawImageToHandle(ByVal rawImage As Byte(), ByVal width As Integer, ByVal height As Integer, ByVal hdc As Integer, ByRef handle As Integer) As Integer
        Dim ri As Integer = ByteArrayToPointer(rawImage)
        CapRawImageToHandle = CapRawImageToHandle(ri, width, height, hdc, handle)
        Marshal.FreeHGlobal(New IntPtr(ri))
    End Function

    ' -----------------------------------------------------------------------------------
    ' Public extraction functions
    ' -----------------------------------------------------------------------------------

    Public Function Extract(ByRef rawimage As Byte(), ByVal width As Integer, ByVal height As Integer, ByVal res As Integer, ByRef tpt As Byte(), ByRef tptSize As Integer, ByVal context As Integer) As Integer
        Dim ri As Integer = ByteArrayToPointer(rawimage)
        Extract = Extract(ri, width, height, res, tpt, tptSize, context)
        Marshal.FreeHGlobal(New IntPtr(ri))
    End Function

    Public Function Extract(ByRef rawimage As Integer, ByVal width As Integer, ByVal height As Integer, ByVal res As Integer, ByRef tpt As Byte(), ByRef tptSize As Integer, ByVal context As Integer) As Integer
        Dim buffer As IntPtr = Marshal.StringToHGlobalAuto(Space(tpt.Length))
        Dim i As Integer = _Extract(rawimage, width, height, res, buffer.ToInt32, tptSize, context)
        If i >= 0 Then Marshal.Copy(buffer, tpt, 0, tptSize)
        Marshal.FreeHGlobal (buffer)
        Extract = i
    End Function

    ' -----------------------------------------------------------------------------------
    ' Public matching functions
    ' -----------------------------------------------------------------------------------

    Public Declare Function GrInitialize Lib "grfinger.dll" Alias "_GrInitialize@0" () As Integer

    Public Declare Function GrFinalize Lib "grfinger.dll" Alias "_GrFinalize@0" () As Integer

    Public Declare Function CreateContext Lib "grfinger.dll" Alias "_GrCreateContext@4" (ByRef contextId As Integer) As Integer

    Public Declare Function DestroyContext Lib "grfinger.dll" Alias "_GrDestroyContext@4" (ByVal contextId As Integer) As Integer

    Public Function Verify(ByRef queryTemplate As Byte, ByRef referenceTemplate As Byte, ByRef verifyScore As Integer, ByVal context As Integer) As Integer
        Verify = _Verify(ByteArrayToPointer(queryTemplate), ByteArrayToPointer(referenceTemplate), verifyScore, context)
    End Function

    Public Function IdentifyPrepare(ByRef templateQuery As Byte, ByVal context As Integer) As Integer
        IdentifyPrepare = _IdentifyPrepare(ByteArrayToPointer(templateQuery), context)
    End Function

    Public Function Identify(ByRef templateReference As Byte, ByRef identifyScore As Integer, ByVal context As Integer) As Integer
        Identify = _Identify(ByteArrayToPointer(templateReference), identifyScore, context)
    End Function

    ' -----------------------------------------------------------------------------------
    ' Public configuration functions
    ' -----------------------------------------------------------------------------------

    Public Declare Function SetIdentifyParameters Lib "grfinger.dll" Alias "_GrSetIdentifyParameters@12" (ByVal identifyThreshold As Integer, ByVal identifyRotationTolerance As Integer, ByVal context As Integer) As Integer

    Public Declare Function SetVerifyParameters Lib "grfinger.dll" Alias "_GrSetVerifyParameters@12" (ByVal verifyThreshold As Integer, ByVal verifyRotationTolerance As Integer, ByVal context As Integer) As Integer

    Public Declare Function GetIdentifyParameters Lib "grfinger.dll" Alias "_GrGetIdentifyParameters@12" (ByRef identifyThreshold As Integer, ByRef identifyRotationTolerance As Integer, ByVal context As Integer) As Integer

    Public Declare Function GetVerifyParameters Lib "grfinger.dll" Alias "_GrGetVerifyParameters@12" (ByRef verifyThreshold As Integer, ByRef verifyRotationTolerance As Integer, ByVal context As Integer) As Integer

    Public Declare Function SetBiometricDisplayColors Lib "grfinger.dll" Alias "_GrSetBiometricDisplayColors@24" (ByVal minutiaeColor As Integer, ByVal minutiaeMatchedColor As Integer, ByVal segmentColor As Integer, ByVal segmentMatchedColor As Integer, ByVal directionColor As Integer, ByVal directionMatchedColor As Integer) As Integer

    ' -----------------------------------------------------------------------------------
    ' Other public functions
    ' -----------------------------------------------------------------------------------

    Public Declare Function GetGrFingerVersion Lib "grfinger.dll" Alias "_GrGetGrFingerVersion@8" (ByRef majorRelease As Byte, ByRef minorRelease As Byte) As Integer

    Public Function BiometricDisplay(ByRef tptQuery As Byte, ByVal rawImage As Integer, ByVal width As Integer, ByVal height As Integer, ByVal res As Integer, ByVal hdc As Integer, ByRef handle As Integer, ByVal matchContext As Integer) As Integer
        Dim tptq As Integer = ByteArrayToPointer(tptQuery)
        BiometricDisplay = _BiometricDisplay(tptq, rawImage, width, height, res, hdc, handle, matchContext)
        Marshal.FreeHGlobal(New IntPtr(tptq))
    End Function

    Public Function BiometricDisplay(ByRef tptQuery As Byte(), ByRef rawImage As Byte(), ByVal width As Integer, ByVal height As Integer, ByVal res As Integer, ByVal hdc As Integer, ByRef handle As Integer, ByVal matchContext As Integer) As Integer
        Dim ri As Integer = ByteArrayToPointer(rawImage)
        BiometricDisplay = BiometricDisplay(tptQuery, ri, width, height, res, hdc, handle, matchContext)
        Marshal.FreeHGlobal(New IntPtr(ri))
    End Function

End Module

