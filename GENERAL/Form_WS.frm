VERSION 5.00
Begin VB.Form Form_WS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CONSULTAR WEB SERVICES"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   13605
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   " CONSUMIR"
      Height          =   540
      Left            =   6615
      TabIndex        =   10
      Top             =   2415
      Width           =   1695
   End
   Begin VB.TextBox TxtClaveAcceso 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6615
      TabIndex        =   9
      Text            =   "1008201650850FA0010050000243460702164179001Dlcjvl1210"
      Top             =   1890
      Width           =   6420
   End
   Begin VB.CommandButton CmdConsultar 
      Caption         =   "CONSULTAR"
      Height          =   540
      Left            =   10920
      TabIndex        =   8
      Top             =   2415
      Width           =   1695
   End
   Begin VB.TextBox TxtResultado 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2850
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3465
      Width           =   12510
   End
   Begin VB.TextBox TxtAction 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6615
      TabIndex        =   2
      Top             =   1365
      Width           =   6420
   End
   Begin VB.TextBox TxtWsdl 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6615
      TabIndex        =   1
      Text            =   "http://mysql.diskcoversystem.com:5001/soap/service_schema?wsdl"
      Top             =   420
      Width           =   6420
   End
   Begin VB.TextBox TxtsOAP 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2640
      Left            =   105
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form_WS.frx":0000
      Top             =   420
      Width           =   6420
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " R E S U L T A D O"
      Height          =   330
      Left            =   105
      TabIndex        =   7
      Top             =   3150
      Width           =   12510
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " ACTION"
      Height          =   330
      Left            =   6615
      TabIndex        =   6
      Top             =   1050
      Width           =   6420
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " WSDL"
      Height          =   330
      Left            =   6615
      TabIndex        =   5
      Top             =   105
      Width           =   6420
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SOAP"
      Height          =   330
      Left            =   105
      TabIndex        =   4
      Top             =   105
      Width           =   6420
   End
End
Attribute VB_Name = "Form_WS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdConsultar_Click()
Dim xmlResponse As MSXML2.DOMDocument30
Dim strSoap As String
Dim strSOAPAction As String
Dim strWsdl As String

strSoap = TxtsOAP.Text
strSOAPAction = TxtAction.Text
strWsdl = TxtWsdl.Text
If InvokeWebService(strSoap, strSOAPAction, strWsdl, xmlResponse) Then
   TxtResultado.Text = xmlResponse.xml
Else
   TxtResultado.Text = "Error"
End If
Set xmlResponse = Nothing
End Sub

Private Sub Command1_Click()
    ' Start Internet Explorer and type in the url of your webservice page
    ' i.e.: http://localhost/myweb/mywebService.asmx
    ' In that page, click on the link to the method you want to call from your application
    ' Select in upper POST section the xml code from <?xml.. to </soap:Envelope>
    ' Copy this into the strXml variable, escape all quotes and replace "string" your parameter value
    ' Copy the url to your webservice page (asmx) to the strUrl variable
    ' Copy the SOAPAction value to the strSoapAction variable

    Dim strSOAPAction As String
    Dim strURL As String
    Dim strXml As String
    Dim strParam As String
    
'''    txtOutput.Text = ""
'''    strParam = "MyParameterString"
'''    strURL = "http://localhost/myweb/mywebService.asmx"
'''    strSOAPAction = "http://tempuri.org/MyMethod"
'''
'''
'''    strXml = "<?xml version=""1.0"" encoding=""utf-8""?>" & _
'''             "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
'''             "<soap:Body>" & _
'''             "<CheckActCode xmlns=""http://tempuri.org/"">" & _
'''             "<actCode>" & strParam & "</actcode>" & _
'''             "</checkactcode>" & _
'''             "</soap:body>" & _
'''             "</soap:envelope>"
'''
'''    ' Call PostWebservice and put result in text box
'''    Debug.Print PostWebservice(strURL, strSOAPAction, strXml)

End Sub

Private Function PostWebservice(ByVal AsmxUrl As String, ByVal SoapActionUrl As String, ByVal XmlBody As String) As String
    Dim objDom As Object
    Dim objXmlHttp As Object
    Dim strRet As String
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    
    On Error GoTo Err_PW
    
    ' Create objects to DOMDocument and XMLHTTP
    Set objDom = CreateObject("MSXML2.DOMDocument")
    Set objXmlHttp = CreateObject("MSXML2.XMLHTTP")
    
    ' Load XML
    objDom.async = False
    objDom.loadXML XmlBody

    ' Open the webservice
    objXmlHttp.open "POST", AsmxUrl, False
    
    ' Create headings
    objXmlHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    objXmlHttp.setRequestHeader "SOAPAction", SoapActionUrl
    
    ' Send XML command
    objXmlHttp.send objDom.xml

    ' Get all response text from webservice
    strRet = objXmlHttp.responseText
    
    ' Close object
    Set objXmlHttp = Nothing
    
    ' Extract result
    intPos1 = InStr(strRet, "Result>") + 7
    intPos2 = InStr(strRet, "<!--")
    If intPos1 >= 7 And intPos2 > 0 Then
        strRet = Mid(strRet, intPos1, intPos2 - intPos1)
    End If
    
    ' Return result
    PostWebservice = strRet
    
Exit Function
Err_PW:
    PostWebservice = "Error: " & Err.Number & " - " & Err.Description

End Function

'''Private Sub Command1_Click()
'''Dim Obj As New Cls_FirmarXML
'''Dim ObjEnviar As New WS_Recepcion
'''Dim ObjAutori As New WS_Autorizacion
'''Dim URLRecepcion As String
'''Dim URLAutorizacion As String
'''Dim Resultado As Boolean
'''Dim RutaCertificado As String
'''Dim ClaveCertificado As String
'''Dim RutaXML As String
'''Dim RutaXMLFirmado As String
'''Dim RutaXMLAutorizado As String
'''Dim RutaXMLRechazado As String
'''Dim MensajeError As String
'''Dim ArrayRecepcion() As String
'''Dim ArrayAutorizacion() As String
'''Dim Tiempo_Espera As Integer
'''Dim Tiempo_SRI As Integer
'''Dim EsperaEspera As Integer
'''Dim SRI_Aut As Tipo_Estado_SRI
'''Dim Intento_Enviar As Byte
'''Dim Intento_Autorizar As Byte
'''
'''   'Pagina de Conexion con www.diskcoversystem.com
'''    URLRecepcion = TxtWsdl
'''    URLAutorizacion = TxtWsdl
'''    RatonReloj
'''    With SRI_Aut
'''        .Clave_De_Acceso = TrimStrg(TxtClaveAcceso)
'''         FA.Estado_SRI = "CG"
'''        .Estado_SRI = FA.Estado_SRI
'''        .Documento_XML = ""
'''        .Error_SRI = ""
'''         RutaXML = RutaSysBases & "\TEMP\G\" & .Clave_De_Acceso & ".xml"
'''         RutaXMLFirmado = RutaSysBases & "\TEMP\F\" & .Clave_De_Acceso & ".xml"
'''         RutaXMLAutorizado = RutaSysBases & "\TEMP\A\" & .Clave_De_Acceso & ".xml"
'''         RutaXMLRechazado = RutaSysBases & "\TEMP\R\" & .Clave_De_Acceso & ".xml"
'''         ArrayRecepcion = ObjEnviar.FF_EnviaXML_SRI(RutaXML, URLRecepcion, RutaXMLRechazado)
'''         MsgBox ArrayRecepcion(0) & vbCrLf _
'''              & ArrayRecepcion(1) & vbCrLf _
'''              & ArrayRecepcion(2)
'''         If ArrayRecepcion(0) = "RECIBIDA" Then
'''            MsgBox "Ok"
'''         Else
'''           .Estado_SRI = "CNF"
'''           .Error_SRI = MensajeError
'''           .Documento_XML = MensajeError
'''         End If
'''    End With
'''    RatonNormal
'''End Sub

Private Sub Form_Activate()
  TxtsOAP = "<?xml version=""1.0"" encoding=""utf-8""?>" _
          & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" " _
          & "xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" " _
          & "xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/""> " _
          & "<soap:Body>" _
          & "<Indicadores xmlns=""Indicadores""> " _
          & "<Fecha>20180201</Fecha> " _
          & "</Indicadores> " _
          & "</soap:Body> " _
          & "</soap:Envelope> "
End Sub

Private Sub Form_Load()
  CentrarForm Form_WS
End Sub
