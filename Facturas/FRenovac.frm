VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FRenovacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RENOVACION DE AUTORIZACION"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtNumAutor1 
      Alignment       =   1  'Right Justify
      Height          =   336
      Left            =   2310
      MaxLength       =   10
      TabIndex        =   12
      Text            =   "0000000001"
      Top             =   2310
      Width           =   1515
   End
   Begin MSAdodcLib.Adodc AdoTipoComprobante 
      Height          =   330
      Left            =   3675
      Top             =   2730
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox TxtNumSerietres2 
      Height          =   336
      Left            =   3045
      MaxLength       =   7
      TabIndex        =   8
      Text            =   "0000001"
      Top             =   1470
      Width           =   750
   End
   Begin VB.CommandButton CmdCerrar 
      Caption         =   "&Salir"
      Height          =   765
      Left            =   3990
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Salir"
      Top             =   1650
      Width           =   990
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Generar XML"
      Height          =   750
      Left            =   3990
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Grabar"
      Top             =   825
      Width           =   960
   End
   Begin VB.TextBox TxtNumSerieUno 
      Height          =   336
      Left            =   2310
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "001"
      ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al código del establecimiento"
      Top             =   735
      Width           =   645
   End
   Begin VB.TextBox TxtNumSerieDos 
      Height          =   336
      Left            =   3045
      MaxLength       =   3
      TabIndex        =   4
      Text            =   "001"
      ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al punto dde emisión"
      Top             =   735
      Width           =   645
   End
   Begin VB.TextBox TxtNumSerietres1 
      Height          =   336
      Left            =   2310
      MaxLength       =   7
      TabIndex        =   7
      Text            =   "0000001"
      Top             =   1470
      Width           =   750
   End
   Begin VB.TextBox TxtNumAutor 
      Alignment       =   1  'Right Justify
      Height          =   336
      Left            =   2310
      MaxLength       =   10
      TabIndex        =   10
      Text            =   "0000000001"
      Top             =   1890
      Width           =   1515
   End
   Begin MSMask.MaskEdBox MBFechaRegis 
      Height          =   330
      Left            =   2310
      TabIndex        =   14
      Top             =   2730
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "0"
   End
   Begin MSDataListLib.DataCombo DCTipoComprobante 
      Bindings        =   "FRenovac.frx":0000
      DataSource      =   "AdoTipoComprobante"
      Height          =   315
      Left            =   105
      TabIndex        =   1
      Top             =   315
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSForms.Label Label5 
      Height          =   225
      Left            =   105
      TabIndex        =   11
      Top             =   2310
      Width           =   2010
      Caption         =   "No. Autorización Nueva"
      Size            =   "3545;397"
      FontName        =   "Arial"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   225
      Left            =   2310
      TabIndex        =   5
      Top             =   1155
      Width           =   1485
      Caption         =   "Desde     Hasta"
      Size            =   "2619;397"
      FontName        =   "Arial"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   225
      Left            =   105
      TabIndex        =   6
      Top             =   1470
      Width           =   2115
      Caption         =   "Secuencial Facturas"
      Size            =   "3731;397"
      FontName        =   "Arial"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label10 
      Height          =   225
      Left            =   105
      TabIndex        =   13
      Top             =   2730
      Width           =   2115
      Caption         =   "Fecha de Renovación"
      Size            =   "3731;397"
      FontName        =   "Arial"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label6 
      Height          =   225
      Left            =   105
      TabIndex        =   9
      Top             =   1890
      Width           =   2115
      Caption         =   "No. Autorización Anterior"
      Size            =   "3731;397"
      FontName        =   "Arial"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label4 
      Height          =   225
      Left            =   105
      TabIndex        =   2
      Top             =   735
      Width           =   2115
      Caption         =   "Número de Serie"
      Size            =   "3731;397"
      FontName        =   "Arial"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label3 
      Height          =   225
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   4830
      Caption         =   "Tipo Comprobante"
      Size            =   "8520;397"
      FontName        =   "Arial"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FRenovacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdCerrar_Click()
  Unload FRenovacion
End Sub

Private Sub CmdGrabar_Click()
Dim NombreArchivo As String
  Codigo = Ninguno
  Cadena = DCTipoComprobante
  If Cadena = "" Then Cadena = Ninguno
  With AdoTipoComprobante.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Descripcion = '" & Cadena & "' ")
       If Not .EOF Then Codigo = .Fields("Tipo_Comprobante_Codigo")
    End If
  End With
  NombreArchivo = RutaSysBases & "\AT\Renovacion_" & TxtNumAutor1 & ".xml"
  NumFile = FreeFile
  Open NombreArchivo For Output As #NumFile   ' Abre el archivo
 'ENCABEZADO
  Print #NumFile, "<?xml version=";
  Print #NumFile, Chr(34);
  Print #NumFile, "1.0";
  Print #NumFile, Chr(34);
  Print #NumFile, " encoding=";
  Print #NumFile, Chr(34);
  Print #NumFile, "UTF-8";
  Print #NumFile, Chr(34);
  Print #NumFile, "?>"
  Print #NumFile, AbrirXML("autorizacion")
  Print #NumFile, vbTab & CampoXML("codTipoTra", 8)
  Print #NumFile, vbTab & CampoXML("ruc", RUC)
  Print #NumFile, vbTab & CampoXML("fecha", MBFechaRegis)
  Print #NumFile, vbTab & CampoXML("autOld", Format(TxtNumAutor, "0000000000"))
  Print #NumFile, vbTab & CampoXML("autNew", Format(TxtNumAutor1, "0000000000"))
 'Detalles
  Print #NumFile, vbTab & vbTab & AbrirXML("detalles")
  Print #NumFile, vbTab & vbTab & vbTab & AbrirXML("detalle")
  Print #NumFile, vbTab & vbTab & vbTab & vbTab & CampoXML("codDoc", Codigo)
  Print #NumFile, vbTab & vbTab & vbTab & vbTab & CampoXML("estab", TxtNumSerieUno)
  Print #NumFile, vbTab & vbTab & vbTab & vbTab & CampoXML("ptoEmi", TxtNumSerieDos)
  Print #NumFile, vbTab & vbTab & vbTab & vbTab & CampoXML("finOld", Val(TxtNumSerietres1))
  Print #NumFile, vbTab & vbTab & vbTab & vbTab & CampoXML("iniNew", Val(TxtNumSerietres2))
  Print #NumFile, vbTab & vbTab & vbTab & CerrarXML("detalle")
  Print #NumFile, vbTab & vbTab & CerrarXML("detalles")
  Print #NumFile, CerrarXML("autorizacion")
  Close NumFile
  MsgBox "El Archivo " & NombreArchivo & ", fue creado con éxito"
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT Tipo_Comprobante_Codigo, Descripcion " _
       & "FROM Tipo_Comprobante " _
       & "WHERE Tipo_Comprobante_Codigo <> 0 " _
       & "ORDER BY Tipo_Comprobante_Codigo "
  SelectDBCombo DCTipoComprobante, AdoTipoComprobante, sSQL, "Descripcion"
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm FRenovacion
  ConectarAdodc AdoTipoComprobante
End Sub

Private Sub MBFechaRegis_GotFocus()
  MarcarTexto MBFechaRegis
End Sub

Private Sub MBFechaRegis_LostFocus()
  FechaValida MBFechaRegis
End Sub

Private Sub TxtNumAutor_GotFocus()
  MarcarTexto TxtNumAutor
End Sub

Private Sub TxtNumAutor1_GotFocus()
  MarcarTexto TxtNumAutor1
End Sub

Private Sub TxtNumSerieDos_GotFocus()
  MarcarTexto TxtNumSerieDos
End Sub

Private Sub TxtNumSerietres1_GotFocus()
  MarcarTexto TxtNumSerietres1
End Sub

Private Sub TxtNumSerietres2_GotFocus()
  MarcarTexto TxtNumSerietres2
End Sub

Private Sub TxtNumSerieUno_GotFocus()
  MarcarTexto TxtNumSerieUno
End Sub
