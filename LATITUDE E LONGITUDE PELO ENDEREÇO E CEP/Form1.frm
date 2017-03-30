VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12150
   LinkTopic       =   "Form1"
   ScaleHeight     =   8220
   ScaleWidth      =   12150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   525
      Left            =   4770
      TabIndex        =   2
      Top             =   180
      Width           =   1245
   End
   Begin VB.TextBox txtCep 
      Height          =   465
      Left            =   390
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   270
      Width           =   2325
   End
   Begin SHDocVwCtl.WebBrowser webLat 
      Height          =   7185
      Left            =   60
      TabIndex        =   0
      Top             =   930
      Width           =   11985
      ExtentX         =   21140
      ExtentY         =   12674
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function getGoogleMapsGeocode(sAddr As String) As String

Dim xhrRequest As XMLHTTP60
Dim sQuery As String
Dim domResponse As DOMDocument60
Dim ixnStatus As IXMLDOMNode
Dim ixnLat As IXMLDOMNode
Dim ixnLng As IXMLDOMNode


' Use the empty string to indicate failure
getGoogleMapsGeocode = ""

Set xhrRequest = New XMLHTTP60
sQuery = "https://maps.googleapis.com/maps/api/geocode/xml?address=" & sAddr & "&key=AIzaSyA-7ItUMJ8nAe7IjYA7Af7_AqzjAnKh_bI"
'sQuery = sQuery & Replace(sAddr, " ", "+")
xhrRequest.Open "GET", sQuery, False
xhrRequest.send

Set domResponse = New DOMDocument60
domResponse.loadXML xhrRequest.responseText
Set ixnStatus = domResponse.selectSingleNode("//status")

If (ixnStatus.Text <> "OK") Then
    Exit Function
End If

Set ixnLat = domResponse.selectSingleNode("/GeocodeResponse/result/geometry/location/lat")
Set ixnLng = domResponse.selectSingleNode("/GeocodeResponse/result/geometry/location/lng")

getGoogleMapsGeocode = ixnLat.Text & ", " & ixnLng.Text

End Function


Private Sub Command1_Click()
    Call getGoogleMapsGeocode(Me.txtCep.Text)
End Sub

Private Function LatLongPeloCep(CEP As String)
    
    Dim HTML1 As String
    Dim HTML2 As String
    Dim HTML3 As String
    Dim HTML4 As String
    Dim HTML5 As String
    Dim HTML6 As String
    Dim HTML7 As String
    Dim HTML8 As String
    Dim HTML9 As String
    Dim HTML10 As String
    Dim HTML11 As String
    Dim HTML12 As String
    Dim HTML13 As String
    Dim HTML14 As String
    Dim HTML15 As String
    Dim HTML16 As String
    Dim HTML17 As String
    Dim HTML18 As String
    Dim HTML19 As String
    Dim HTML20 As String
    Dim HTML21 As String
    Dim HTML22 As String
    Dim HTML23 As String
    Dim HTML24 As String
    Dim HTML25 As String
    Dim HTML26 As String
    Dim strPara As String
    
    strPara = Latitude_Cliente + "," + Longitude_Cliente
    
    HTML1 = "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN""" + vbNewLine
    HTML2 = "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"">" + vbNewLine
    HTML3 = "<html xmlns=""http://www.w3.org/1999/xhtml""  xmlns:v=""urn:schemas-microsoft-com:vml"">" + vbNewLine
    HTML4 = "<head>" + vbNewLine
    HTML5 = "<title>Google Maps</title>" + vbNewLine
    HTML6 = "<meta http-equiv=""content-type"" content=""text/html; charset=UTF-8""/>" + vbNewLine
    HTML7 = "<script src=""http://maps.google.com/maps?file=api&v=2.x&key=ABQIAAAAqEiwOrYZ0Lx-d0jtb46NVBSFvCKpyhD1OJRHiGgu5YNzbovqLxSMljIMZwxUTo1W2PGwkDpSMZ1BSw""" + vbNewLine
    HTML8 = "type=""text/javascript""></script>" + vbNewLine
    HTML9 = "<script type=""text/javascript"">" + vbNewLine
    HTML10 = "var lat = '';" + vbNewLine
    HTML11 = "var lng = '';" + vbNewLine
    HTML12 = "var address = '" + CEP + "';" + vbNewLine
    HTML13 = "geocoder.geocode( { 'address': address}, function(results, status) {" + vbNewLine
    HTML14 = "if (status == google.maps.GeocoderStatus.OK) {" + vbNewLine
    HTML15 = "lat = results[0].geometry.location.lat();" + vbNewLine
    HTML16 = "lng = results[0].geometry.location.lng();" + vbNewLine
    HTML17 = " }" + vbNewLine
    HTML18 = "});" + vbNewLine
    HTML20 = "</script>" + vbNewLine
    HTML21 = "</head>" + vbNewLine
    HTML22 = "<body onload=""initialize()"" style=""font-family: tahoma, tahoma; font-size: 8px; border: 0;margin-left: 0px;margin-right: 0px;margin-top: 0px;margin-bottom: 0px;"">" & vbNewLine
    HTML23 = "<div id=""map_canvas"" style=""width: 74%; height: 355px; float:left; border: 1px solid black;""></div>" + vbNewLine
    HTML24 = "<div id=""route"" style=""width: 24%; height:100px; float:left; border; 1px solid black;""></div>" + vbNewLine
    HTML25 = "<br/>" + vbNewLine
    HTML26 = "</body>" + vbNewLine
    HTML27 = "</html>"

    Dim fso As New FileSystemObject
    Dim strArquivo As String
   
    strArquivo = "geoLatLong" & Time
    strArquivo = Replace(strArquivo, ":", "_")
    strArquivo = strArquivo & ".html"
    
    
'    If fso.FileExists(Caminho_html + "\gmap.html") Then
'        fso.DeleteFile (Caminho_html + "\gmap.html")
'    End If
    
    Dim tx As TextStream
    Caminho_html = "C:"
    fso.CreateTextFile (Caminho_html + "\" + strArquivo)
    fso.OpenTextFile Caminho_html + "\" + strArquivo, ForWriting, True
    
    Set tx = fso.OpenTextFile(Caminho_html + "\" + strArquivo, ForWriting, True)
    
    tx.WriteLine HTML1 + HTML2 + HTML3 + HTML4 + HTML5 + HTML6 + HTML7 + HTML8 + HTML9 + HTML10 + HTML11 + HTML12 + HTML13 + HTML14 + HTML15 + HTML16 + HTML17 + HTML18 + HTML19 + HTML20 + HTML21 + HTML22 + HTML23 + HTML24 + HTML25 + HTML26 + HTML27
    tx.Close
    
    webLat.Navigate (Caminho_html + "\" + strArquivo)

End Function

    

