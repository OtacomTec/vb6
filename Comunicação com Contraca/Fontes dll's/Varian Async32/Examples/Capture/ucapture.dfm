�
 TFORM1 0l  TPF0TForm1Form1Left� TopkWidth Height� CaptionOnRxChar data eventFont.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style OnCreate
FormCreatePixelsPerInch`
TextHeight TLabelLabel2LeftTopWidth� Height-AutoSizeCaption�Demo which shows how to read and capture messages from the comport component. In this example we wait for #13 delimited messages.WordWrap	  TLabelLabel3LeftTopHWidth� HeightAutoSizeCaption-1. Include ceRxChar in MonitorEvents propertyWordWrap	  TLabelLabel4LeftTop`Width� HeightAutoSizeCaption-2. Add an eventhandler for the OnRxChar eventWordWrap	  TLabelLabel1LeftTopxWidth� HeightAutoSizeCaption-3. Add the required code to the eventhandler WordWrap	  TMemoMemo1LeftTopWidth� HeightYTabOrder   TButtonButton1Left\Top� WidthKHeightCaptionExitTabOrderOnClickButton1Click  TMemoMemo2LeftTophWidth� HeightYTabOrder  TVaCommVaComm1PortNum
DeviceNameCOM%dMonitorEventsceRxChar OnRxCharVaComm1RxCharLeft� Top   