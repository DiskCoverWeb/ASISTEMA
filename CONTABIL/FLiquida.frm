VERSION 5.00
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Begin VB.Form FLiquidacionCompras 
   BackColor       =   &H00FFC0C0&
   Caption         =   "LISTAR COMPROBANTE DE LIQUIDACION DE COMPRAS"
   ClientHeight    =   7680
   ClientLeft      =   180
   ClientTop       =   465
   ClientWidth     =   14430
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7680
   ScaleWidth      =   14430
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   14430
      _ExtentX        =   25453
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   7
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir del modulo"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir Comprobante"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Cambio_Retencion"
            Object.ToolTipText     =   "Cambia Numero de Retención"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "PDF"
            Object.ToolTipText     =   "Genera la Retencion en PDF"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Autorizar"
            Object.ToolTipText     =   "Autorizar Retencion Pendiente"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Autorizar_Grupo"
            Object.ToolTipText     =   "Autoriza Retenciones en Lote"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Enviar_Mails"
            Object.ToolTipText     =   "Envia Retencion por Emails"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.CheckBox CheqClaveAcceso 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Caption         =   "Clave de Accceso:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1260
      Width           =   6105
   End
   Begin AcroPDFLibCtl.AcroPDF APDFLiquidacion 
      Height          =   3165
      Left            =   105
      TabIndex        =   17
      Top             =   2730
      Width           =   11355
      _cx             =   5080
      _cy             =   5080
   End
   Begin VB.ListBox LstResultado 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   105
      TabIndex        =   16
      Top             =   1995
      Width           =   12720
   End
   Begin VB.OptionButton OpcManuales 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Manuales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   3360
      TabIndex        =   2
      Top             =   735
      Width           =   1170
   End
   Begin VB.OptionButton OpcAut 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Autorizados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   735
      Value           =   -1  'True
      Width           =   1380
   End
   Begin VB.ComboBox CMeses 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   5145
      TabIndex        =   4
      Text            =   "Diciembre"
      Top             =   735
      Width           =   1275
   End
   Begin VB.OptionButton OpcNoAut 
      BackColor       =   &H00FFC0C0&
      Caption         =   "No Autorizados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   1575
      TabIndex        =   1
      Top             =   735
      Width           =   1695
   End
   Begin VB.TextBox TxtAutorizacion 
      Height          =   330
      Left            =   6300
      TabIndex        =   14
      ToolTipText     =   "<Ctrl+A> Autorizar manualmente desde el SRI"
      Top             =   1575
      Width           =   6525
   End
   Begin VB.TextBox TxtClave 
      Height          =   330
      Left            =   105
      TabIndex        =   12
      ToolTipText     =   "<Ctrl+A> Volver a Generar y Firmar el Documento Electronico"
      Top             =   1575
      Width           =   6105
   End
   Begin MSDataListLib.DataCombo DCComp 
      Bindings        =   "FLiquida.frx":0000
      DataSource      =   "AdoComp"
      Height          =   345
      Left            =   9030
      TabIndex        =   8
      Top             =   735
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   609
      _Version        =   393216
      Text            =   "999999999"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   105
      MaskColor       =   &H00FFC0C0&
      TabIndex        =   11
      Top             =   2100
      Width           =   330
   End
   Begin MSAdodcLib.Adodc AdoComp1 
      Height          =   330
      Left            =   315
      Top             =   3675
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      Caption         =   "Comp1"
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
   Begin MSAdodcLib.Adodc AdoComp 
      Height          =   330
      Left            =   315
      Top             =   3990
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      Caption         =   "Comp"
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
   Begin MSAdodcLib.Adodc AdoDetRet 
      Height          =   330
      Left            =   315
      Top             =   4305
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      Caption         =   "DetRet"
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
   Begin MSAdodcLib.Adodc AdoDetCom 
      Height          =   330
      Left            =   2310
      Top             =   3990
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      Caption         =   "DetCom"
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
   Begin MSDataListLib.DataCombo DCSerie 
      Bindings        =   "FLiquida.frx":0016
      DataSource      =   "AdoSerie"
      Height          =   345
      Left            =   7455
      TabIndex        =   6
      Top             =   735
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   609
      _Version        =   393216
      Text            =   "999999"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoSerie 
      Height          =   330
      Left            =   2310
      Top             =   4305
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      Caption         =   "Serie"
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
   Begin MSAdodcLib.Adodc AdoTP_Num 
      Height          =   330
      Left            =   2310
      Top             =   3675
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      Caption         =   "TP_Num"
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
   Begin MSDataListLib.DataCombo DCTP_Num 
      Bindings        =   "FLiquida.frx":002D
      DataSource      =   "AdoTP_Num"
      Height          =   345
      Left            =   11130
      TabIndex        =   10
      Top             =   735
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   609
      _Version        =   393216
      Text            =   "999999999"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &T.C."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10605
      TabIndex        =   9
      Top             =   735
      Width           =   540
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8610
      TabIndex        =   7
      Top             =   735
      Width           =   435
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   4620
      TabIndex        =   3
      Top             =   735
      Width           =   540
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Autorizacion:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6300
      TabIndex        =   13
      Top             =   1260
      Width           =   6525
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Serie No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6510
      TabIndex        =   5
      Top             =   735
      Width           =   960
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   1365
      Top             =   5460
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   11
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLiquida.frx":0045
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLiquida.frx":035F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLiquida.frx":DC71
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLiquida.frx":1B583
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLiquida.frx":28E95
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLiquida.frx":367A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLiquida.frx":36AC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLiquida.frx":64973
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLiquida.frx":64C8D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLiquida.frx":64FA7
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLiquida.frx":A6EA9
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FLiquidacionCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Retencion_No As Long
Dim Serie_R As String
 
Dim ObjAutori As New WS_Autorizacion
Dim URLRecepcion  As String
Dim URLAutorizacion As String

Dim RutaXMLAutorizado As String
Dim RutaXMLRechazado As String
Dim RutaXMLFirmado As String
Dim ArrayAutorizacion() As String
'Dim SRI_Aut As Tipo_Estado_SRI
Dim Resultado_Ret As String
 
Private Sub Command1_Click()
   Unload FLiquidacionCompras
End Sub

Private Sub DCComp_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCComp_LostFocus()
  sSQL = "SELECT (TP + ' ' + CAST(Numero AS VARCHAR)) As TC_Num " _
       & "FROM Trans_Compras " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Establecimiento+PuntoEmision = '" & DCSerie & "' " _
       & "AND Secuencial = " & Val(DCComp) & " " _
       & "AND TipoComprobante IN(3,41) " _
       & "ORDER BY TP,Numero "
  SelectDB_Combo DCTP_Num, AdoTP_Num, sSQL, "TC_Num", True
End Sub

Private Sub DCSerie_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub DCSerie_LostFocus()
   Listar_Tipo_Liquidacion DCSerie
End Sub

Private Sub DCTP_Num_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  PresionoEnter KeyCode
End Sub

Private Sub DCTP_Num_LostFocus()
  If Not IsNumeric(DCComp) Then DCComp = "0"
  Listar_Liquidacion DCSerie, Val(DCComp), SinEspaciosIzq(DCTP_Num), Val(SinEspaciosDer(DCTP_Num))
End Sub

Private Sub Form_Activate()
   sSQL = "UPDATE Trans_Compras " _
        & "SET Estado_SRI_LC = 'OK' " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND TipoComprobante IN (3,41) " _
        & "AND LEN(Clave_Acceso) = 49 " _
        & "AND LEN(AutRetencion) = 49 " _
        & "AND Estado_SRI_LC <> 'OK' "
   Ejecutar_SQL_SP sSQL
   
   LstResultado.width = FLiquidacionCompras.width - 400
   APDFLiquidacion.width = FLiquidacionCompras.width - 400
   APDFLiquidacion.Height = MDI_Y_Max - APDFLiquidacion.Top - 100
   
   APDFLiquidacion.Object.src = ""
   
   RatonNormal
End Sub

Private Sub Form_Load()
   'CentrarForm FLiquidacionCompras
''   ATSPDF.Height = MDI_Y_Max - ATSPDF.Top - 200
''   ATSPDF.width = MDI_X_Max - 250
''   LstResultado.width = MDI_X_Max - 250
   
   ConectarAdodc AdoComp
   ConectarAdodc AdoDetRet
   ConectarAdodc AdoDetCom
   ConectarAdodc AdoComp1
   ConectarAdodc AdoSerie
   ConectarAdodc AdoTP_Num
   
   CMeses.Clear
   CMeses.AddItem "Todos"
   CMeses.AddItem "Enero"
   CMeses.AddItem "Febrero"
   CMeses.AddItem "Marzo"
   CMeses.AddItem "Abril"
   CMeses.AddItem "Mayo"
   CMeses.AddItem "Junio"
   CMeses.AddItem "Julio"
   CMeses.AddItem "Agosto"
   CMeses.AddItem "Septiembre"
   CMeses.AddItem "Octubre"
   CMeses.AddItem "Noviembre"
   CMeses.AddItem "Diciembre"
   CMeses.Text = "Todos"
   
  'Pagina de Conexion con el SRI
   URLRecepcion = Leer_Campo_Empresa("Web_SRI_Recepcion")
   URLAutorizacion = Leer_Campo_Empresa("Web_SRI_Autorizado")
End Sub

Public Sub Listar_Liquidacion(Serie_R As String, Comp_No As Long, TP As String, TP_No As Long)
Dim TipoProc As String
   RatonReloj
  'Listar las FLiquidacionCompras del IVA
   TxtAutorizacion = ""
   TxtClave = ""
     
   sSQL = "SELECT C.Cliente,C.CI_RUC,C.TD,C.Direccion,C.Email,C.Email2,C.Ciudad,C.DirNumero,C.Telefono,Co.Concepto,TC.* " _
        & "FROM Trans_Compras As TC,Clientes As C,Comprobantes As Co " _
        & "WHERE TC.Item = '" & NumEmpresa & "' " _
        & "AND TC.Periodo = '" & Periodo_Contable & "' " _
        & "AND TC.Secuencial = " & Comp_No & " " _
        & "AND TC.Establecimiento+TC.PuntoEmision = '" & Serie_R & "' " _
        & "AND TC.Numero = " & Val(TP_No) & " " _
        & "AND TC.TP = '" & TP & "' " _
        & "AND TC.TipoComprobante IN(3,41) " _
        & "AND TC.IdProv = C.Codigo " _
        & "AND TC.Item = Co.Item " _
        & "AND TC.Periodo = Co.Periodo " _
        & "AND TC.TP = Co.TP " _
        & "AND TC.Numero = Co.Numero " _
        & "ORDER BY Cta_Servicio,Cta_Bienes "
   Select_Adodc AdoDetCom, sSQL
   With AdoDetCom.Recordset
    If .RecordCount > 0 Then
       'MsgBox .RecordCount
        Co.Fecha = .fields("Fecha")
        Co.Beneficiario = .fields("Cliente")
        Co.RUC_CI = .fields("CI_RUC")
        Co.Direccion = .fields("Direccion")
        Co.TD = .fields("TD")
        Co.Email = .fields("Email")
        Co.TP = .fields("TP")
        Co.Numero = .fields("Numero")
        Co.Concepto = .fields("Concepto")
        
        FA.EmailC = .fields("Email")
        FA.EmailR = .fields("Email2")
        FA.TP = Co.TP
        FA.Numero = Co.Numero
        FA.Fecha = .fields("FechaEmision")
        FA.ClaveAcceso_LC = .fields("Clave_Acceso_LC")
        FA.Estado_SRI_LC = .fields("Estado_SRI_LC")
        FA.Autorizacion_LC = .fields("Autorizacion")
        FA.Serie_LC = .fields("Establecimiento") & .fields("PuntoEmision")
        FA.Factura = .fields("Secuencial")
        FA.Sin_IVA = .fields("BaseImponible")
        FA.Con_IVA = .fields("BaseImpGrav")
        FA.Total_IVA = .fields("MontoIva")
        FA.Total_MN = FA.Sin_IVA + FA.Con_IVA + FA.Total_IVA
        FA.SubTotal = FA.Sin_IVA + FA.Con_IVA
        TxtAutorizacion = FA.Autorizacion_LC
        TxtClave = FA.ClaveAcceso_LC
        CheqClaveAcceso.Caption = " Clave de Acceso al SRI(" & FA.Estado_SRI_LC & ") "
        FechaTexto = Co.Fecha
        
        SetNombrePRN = ""
        SRI_Generar_PDF_LC FA, False
        APDFLiquidacion.Object.src = RutaDocumentoPDF
        'Presentar_PDF ATSPDF, RutaDocumentoPDF
        RatonNormal
    Else
      RatonNormal
    End If
   End With

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
  TextoImprimio = ""
  LstResultado.Clear
  Select Case Button.key
    Case "Salir"
         Unload FLiquidacionCompras
    Case "Imprimir"
         NumComp = Co.Numero
         Co.Item = NumEmpresa
         ImprimirComprobantesDe True, Co
         DCComp.SetFocus
    Case "Cambio_Retencion"
         Cambio_Retencion
    Case "PDF"
         Co.Item = NumEmpresa
         FA.TP = Co.TP
         FA.Numero = Co.Numero
         SetNombrePRN = ""
         SRI_Generar_PDF_LC FA, True
    Case "Autorizar"
         FA.TP = Co.TP
         FA.Numero = Co.Numero
         SRI_Autorizacion = SRI_Generar_XML(FA.ClaveAcceso_LC, FA.Estado_SRI_LC)
         SRI_Actualizar_XML_Liquidacion SRI_Autorizacion, FA
         RatonReloj
         If SRI_Autorizacion.Estado_SRI = "OK" Then
            FA.Numero = Co.Numero
            FA.TP = Co.TP
            SRI_Actualizar_Autorizacion_Liquidacion SRI_Autorizacion, FA
            
            SRI_Generar_PDF_LC FA, True
            SRI_Enviar_Mails FA, SRI_Autorizacion, "LC"
            LstResultado.AddItem SRI_Autorizacion.Estado_SRI
            RatonNormal
         Else
            LstResultado.AddItem SRI_Autorizacion.Estado_SRI & " - " & Replace(SRI_Autorizacion.Error_SRI, vbCrLf, " ")
            RatonNormal
         End If
    Case "Autorizar_Grupo"
'''         If Len(DCSerie) < 6 Then DCSerie = "001001"
'''         sSQL = "SELECT C.Cliente,C.CI_RUC,C.TD,C.Direccion,C.Email,C.Ciudad,C.DirNumero,C.Telefono,TC.* " _
'''              & "FROM Trans_Compras As TC,Clientes As C " _
'''              & "WHERE TC.Item = '" & NumEmpresa & "' " _
'''              & "AND TC.Periodo = '" & Periodo_Contable & "' " _
'''              & "AND TC.Serie_Retencion = '" & DCSerie & "' " _
'''              & "AND LEN(AutRetencion) BETWEEN 13 and 37 " _
'''              & "AND LEN(Clave_Acceso) >= 1 " _
'''              & "AND Estado_SRI <> 'OK' " _
'''              & "AND TC.IdProv = C.Codigo " _
'''              & "ORDER BY Serie_Retencion,SecRetencion "
'''         Select_Adodc AdoComp1, sSQL
'''         With AdoComp1.Recordset
'''          If .RecordCount > 0 Then
'''              Do While Not .EOF
'''                 RatonReloj
'''                 Co.TP = .Fields("TP")
'''                 Co.Numero = .Fields("Numero")
'''                 Co.Fecha = .Fields("Fecha")
'''
'''                 FA.TP = Co.TP
'''                 FA.Numero = Co.Numero
'''                 FA.Fecha = .Fields("Fecha")
'''
'''                 SRI_Crear_Clave_Acceso_FLiquidacionCompras FA, False
'''                 If SRI_Autorizacion.Estado_SRI <> "OK" Then
'''                    SRI_Autorizacion.Clave_De_Acceso = FA.ClaveAcceso
'''                    SRI_Autorizacion.Estado_SRI = "CF"
'''                    SRI_Autorizacion.Error_SRI = ""
'''                    RutaXMLAutorizado = RutaDocumentos & "\Comprobantes Autorizados\" & FA.ClaveAcceso & ".xml"
'''                    RutaXMLRechazado = RutaDocumentos & "\Comprobantes no Autorizados\" & FA.ClaveAcceso & ".xml"
'''                    ArrayAutorizacion = ObjAutori.FF_ObtieneNumAutorizado(URLAutorizacion, FA.ClaveAcceso, RutaXMLAutorizado, RutaXMLRechazado)
'''                    If ArrayAutorizacion(0) = "AUTORIZADO" Then
'''                       SRI_Autorizacion.Estado_SRI = "OK"
'''                       SRI_Autorizacion.Error_SRI = "OK"
'''                       SRI_Autorizacion.Autorizacion = ArrayAutorizacion(1)
'''                       SRI_Autorizacion.Fecha_Autorizacion = Format$(MidStrg(ArrayAutorizacion(2), 1, 10), "dd/MM/yyyy")
'''                       SRI_Autorizacion.Hora_Autorizacion = MidStrg(ArrayAutorizacion(2), 12, 8)
'''                       SRI_Autorizacion.Documento_XML = Leer_Archivo_Texto(RutaXMLAutorizado)
'''                       SRI_Actualizar_Documento_XML SRI_Autorizacion.Clave_De_Acceso
'''                       SRI_Actualizar_Autorizacion_Retencion SRI_Autorizacion, FA
'''                    Else
'''                       LstResultado.AddItem Co.Fecha & ": " & Co.TP & "-" & Format(Co.Numero, "000000000") & " - " _
'''                                          & SRI_Autorizacion.Estado_SRI & " - " & SRI_Autorizacion.Error_SRI
'''                    End If
'''                 End If
'''                 RatonNormal
'''                .MoveNext
'''              Loop
'''          End If
'''         End With
         
         FLiquidacionCompras.Caption = "LISTAR COMPROBANTES DE RETENCION"
    Case "Enviar_Mails"
         FA.Numero = Co.Numero
         FA.TP = Co.TP
         SRI_Autorizacion.Fecha_Autorizacion = Co.Fecha
         SRI_Autorizacion.Autorizacion = FA.Autorizacion_LC
         'FA.Retencion = 1
         'FA.Serie_R = 1
         SRI_Enviar_Mails FA, SRI_Autorizacion, "LC"
  End Select
'''  If AdoComp.Recordset.RecordCount > 0 Then
'''     DCComp = AdoComp.Recordset.Fields("SecRetencion")
'''  Else
'''     DCComp = "0"
'''  End If
End Sub

Public Sub Cambio_Retencion()
Dim Cambio_Ret As Long
Dim TPrompt As String
   TPrompt = "CAMBIAR LA RETENCION No. " & Format(Retencion_No, "00000000") & vbCrLf & vbCrLf _
           & "POR EL NUEVO NUMERO DE RETENCION:"
   Cambio_Ret = Val(InputBox(TPrompt, "CAMBIO DE NUMERO DE RETENCION", CStr(Retencion_No)))
   If Cambio_Ret > 0 And (Cambio_Ret <> Retencion_No) Then
      RatonReloj
      sSQL = "UPDATE Trans_Compras " _
           & "SET SecRetencion = " & Cambio_Ret & " " _
           & "WHERE Item = '" & NumEmpresa & "' " _
           & "AND Periodo = '" & Periodo_Contable & "' " _
           & "AND TP = '" & FA.TP & "' " _
           & "AND Numero = " & FA.Numero & " " _
           & "AND Serie_Retencion = '" & FA.Serie_R & "' " _
           & "AND SecRetencion = " & FA.Retencion & " "
      Ejecutar_SQL_SP sSQL
      RatonReloj
      sSQL = "UPDATE Trans_Air " _
           & "SET SecRetencion = " & Cambio_Ret & " " _
           & "WHERE Item = '" & NumEmpresa & "' " _
           & "AND Tipo_Trans IN ('C','I') " _
           & "AND TP = '" & FA.TP & "' " _
           & "AND Numero = " & FA.Numero & " " _
           & "AND EstabRetencion+PtoEmiRetencion = '" & FA.Serie_R & "' " _
           & "AND SecRetencion = " & FA.Retencion & " "
      Ejecutar_SQL_SP sSQL
      RatonNormal
      MsgBox "Proceso Exitoso, vuelva consultar"
   End If
End Sub

Public Sub Listar_Tipo_Liquidacion(Serie_LC As String)
Dim MesNo As Integer
  MesNo = CMeses.ListIndex
  If MesNo < 0 Then MesNo = 0
  If Len(Serie_LC) < 6 Then Serie_LC = "001001"
  sSQL = "SELECT Secuencial " _
       & "FROM Trans_Compras " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Establecimiento+PuntoEmision = '" & Serie_LC & "' " _
       & "AND TipoComprobante IN(3,41) " _
       & "AND Secuencial > 0 "
  If OpcNoAut.Value Then
      sSQL = sSQL _
          & "AND LEN(Autorizacion) = 13 " _
          & "AND LEN(Clave_Acceso_LC) >= 1 " _
          & "AND Estado_SRI_LC <> 'OK' "
  End If
  If MesNo > 0 Then sSQL = sSQL & "AND MONTH(Fecha) = " & MesNo & " "
  sSQL = sSQL _
       & "GROUP BY Secuencial " _
       & "ORDER BY Secuencial "
  SelectDB_Combo DCComp, AdoComp, sSQL, "Secuencial", True
  'If AdoComp.Recordset.RecordCount > 0 Then ListarRetencion Serie_R, DCComp
End Sub

Public Sub Listar_Tipo_Serie_Liquidacion()
Dim MesNo As Integer
  MesNo = CMeses.ListIndex
  If MesNo < 0 Then MesNo = 0
  sSQL = "SELECT (Establecimiento+PuntoEmision) As Serie_LC " _
       & "FROM Trans_Compras " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TipoComprobante IN (3,41) " _
       & "AND LEN(Establecimiento+PuntoEmision) = 6 "
  If OpcNoAut.Value Then
      sSQL = sSQL _
          & "AND LEN(Autorizacion) BETWEEN 13 and 49 " _
          & "AND LEN(Clave_Acceso_LC) >= 1 " _
          & "AND Estado_SRI_LC <> 'OK' "
  ElseIf OpcManuales.Value Then
    sSQL = sSQL _
          & "AND LEN(Autorizacion) < 13 " _
          & "AND LEN(Clave_Acceso_LC) = 1 "
  Else
    sSQL = sSQL _
          & "AND LEN(Autorizacion) > 13 " _
          & "AND LEN(Clave_Acceso_LC) > 13 " _
          & "AND Estado_SRI_LC = 'OK' "
  End If
  If MesNo > 0 Then sSQL = sSQL & "AND MONTH(Fecha) = " & MesNo & " "
  sSQL = sSQL _
       & "GROUP BY Establecimiento,PuntoEmision " _
       & "ORDER BY Establecimiento,PuntoEmision "
  SelectDB_Combo DCSerie, AdoSerie, sSQL, "Serie_LC"
End Sub

Private Sub CMeses_GotFocus()
  MarcarTexto CMeses
End Sub

Private Sub CMeses_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub CMeses_LostFocus()
  Listar_Tipo_Serie_Liquidacion
  RatonNormal
End Sub

Private Sub TxtAutorizacion_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Secuencial As String
    Keys_Especiales Shift
    If CtrlDown And KeyCode = vbKeyA Then
       If CFechaLong(FechaSistema) <= CFechaLong(Fecha_CE) Then
         Secuencial = "INGRESE LA AUTORIZACION DEL" & vbCrLf _
                    & "DOCUMENTO " & Co.Fecha & "-" & Co.TP & "-" & Format$(Co.Numero, "000000000") & vbCrLf _
                    & "Autorizacion: " & TxtAutorizacion & ":"
         Autorizacion = InputBox(Secuencial, "CAMBIO DE AUTORIZACION", TxtAutorizacion)
         If IsNumeric(Autorizacion) And Len(Autorizacion) >= 3 Then
            SQL1 = "UPDATE Trans_Air " _
                 & "SET Autorizacion = '" & Autorizacion & "' " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "AND Numero = " & Co.Numero & " " _
                 & "AND TP = '" & Co.TP & "' " _
                 & "AND Establecimiento = '" & MidStrg(FA.Serie_LC, 1, 3) & "' " _
                 & "AND PuntoEmision = '" & MidStrg(FA.Serie_LC, 4, 3) & "' " _
                 & "AND Secuencial = '" & FA.Factura & "' "
            Ejecutar_SQL_SP SQL1
        
            MsgBox "Proceso terminado"
         End If
       Else
         RatonNormal
         MsgBox MensajeNoAutorizarCE
       End If
    End If
End Sub

Private Sub TxtClave_KeyDown(KeyCode As Integer, Shift As Integer)
Dim RutaXMLFirmado As String
    Keys_Especiales Shift
    If CtrlDown And KeyCode = vbKeyA Then
       If CFechaLong(FechaSistema) <= CFechaLong(Fecha_CE) Then
          RatonReloj
          TextoImprimio = ""
          sSQL = "UPDATE Trans_Compras " _
               & "SET Estado_SRI_LC = '.', Clave_Acceso_LC = '.' " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND TP = '" & Co.TP & "' " _
               & "AND Numero = '" & Co.Numero & "' " _
               & "AND TipoComprobante IN(3,41) " _
               & "AND Estado_SRI_LC <> 'OK' "
          Ejecutar_SQL_SP sSQL
          If Len(FA.ClaveAcceso_LC) > 1 Then
             RutaXMLFirmado = RutaDocumentos & "\Comprobantes Generados\" & FA.ClaveAcceso_LC & ".xml"
             If Dir$(RutaXMLFirmado) <> "" Then Kill RutaXMLFirmado
             
             RutaXMLFirmado = RutaDocumentos & "\Comprobantes no Autorizados\" & FA.ClaveAcceso_LC & ".xml"
             If Dir$(RutaXMLFirmado) <> "" Then Kill RutaXMLFirmado
                       
             RutaXMLFirmado = RutaDocumentos & "\Comprobantes Firmados\" & FA.ClaveAcceso_LC & ".xml"
             If Dir$(RutaXMLFirmado) <> "" Then Kill RutaXMLFirmado
          End If

          If CFechaLong(FechaSistema) <= CFechaLong(Fecha_CE) Then
            'MsgBox "ClaveAcceso_LC: " & FA.ClaveAcceso_LC
             FA.Numero = Co.Numero
             FA.TP = Co.TP
             SRI_Crear_Clave_Acceso_Liquidacion FA, False, CBool(CheqClaveAcceso.Value)
             RatonNormal
          Else
             RatonNormal
             MsgBox MensajeNoAutorizarCE
          End If
          RatonNormal
         If SRI_Autorizacion.Estado_SRI <> "OK" Then
            LstResultado.AddItem SRI_Autorizacion.Estado_SRI & " - " & Replace(SRI_Autorizacion.Error_SRI, vbCrLf, ", ")
            RatonNormal
         End If
         If TextoImprimio <> "" Then
            RutaGeneraFile = RutaSysBases & "\TEMP\Informe de Errores de Liquidacion de Compras " & Replace(FechaSistema, "/", "-") & ".txt"
            NumFile = FreeFile
            Open RutaGeneraFile For Output As #NumFile ' Abre el archivo.
                 Print #NumFile, TextoImprimio;
            Close #NumFile
            MsgBox "ARCHIVO DE INFORME DE ERRORES:" & vbCrLf & vbCrLf & RutaGeneraFile
         End If
        'MsgBox "Proceso Exitoso, Vuelva a Intentar conectarce con el S.R.I."
       Else
          RatonNormal
          MsgBox MensajeNoAutorizarCE
       End If
    End If
'''    If CtrlDown And KeyCode = vbKeyW Then
'''       MsgBox Fecha_CE
'''       If CFechaLong(FechaSistema) <= CFechaLong(Fecha_CE) Then
'''          FA.Numero = Co.Numero
'''          FA.TP = Co.TP
'''
'''          SRI_Crear_Clave_Acceso_FLiquidacionCompras FA, True
'''          RatonNormal
'''       Else
'''          RatonNormal
'''          MsgBox MensajeNoAutorizarCE
'''       End If
'''    End If
End Sub
