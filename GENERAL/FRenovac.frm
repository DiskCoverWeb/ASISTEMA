VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form FRenovacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RENOVACION DE AUTORIZACION"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtXML 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   10
      Top             =   1890
      Width           =   8205
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Presentar &Petición"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   5670
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Grabar"
      Top             =   840
      Width           =   1275
   End
   Begin MSDataListLib.DataCombo DCAutorizacionOld 
      Bindings        =   "FRenovac.frx":0000
      DataSource      =   "AdoAutorizacionOld"
      Height          =   345
      Left            =   3255
      TabIndex        =   6
      Top             =   1155
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   609
      _Version        =   393216
      Text            =   "DataCombo1"
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
   Begin VB.CheckBox CheqAutOld 
      Caption         =   "Autorización Anterior"
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
      Left            =   3255
      TabIndex        =   5
      Top             =   840
      Width           =   2220
   End
   Begin MSDataListLib.DataCombo DCSerie 
      Bindings        =   "FRenovac.frx":0021
      DataSource      =   "AdoSerie"
      Height          =   345
      Left            =   1995
      TabIndex        =   4
      Top             =   1155
      Visible         =   0   'False
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   609
      _Version        =   393216
      Text            =   "001001"
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
   Begin VB.CheckBox CheqSerie 
      Caption         =   "Serie"
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
      Left            =   1995
      TabIndex        =   3
      Top             =   840
      Width           =   1170
   End
   Begin MSAdodcLib.Adodc AdoTipoComprobante 
      Height          =   330
      Left            =   210
      Top             =   5145
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
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
      Caption         =   "TipoComprobante"
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
   Begin VB.CommandButton CmdCerrar 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   7035
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Salir"
      Top             =   840
      Width           =   1275
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Generar XML"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   7035
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Grabar"
      Top             =   105
      Width           =   1275
   End
   Begin MSMask.MaskEdBox MBFechaRegis 
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   420
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "0"
   End
   Begin MSAdodcLib.Adodc AdoAutorizacion 
      Height          =   330
      Left            =   210
      Top             =   4200
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
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
      Caption         =   "Autorizacion"
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
   Begin MSDataListLib.DataCombo DCAutorizacion 
      Bindings        =   "FRenovac.frx":0038
      DataSource      =   "AdoAutorizacion"
      Height          =   345
      Left            =   105
      TabIndex        =   2
      Top             =   1155
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   609
      _Version        =   393216
      Text            =   "DataCombo1"
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
   Begin MSDataListLib.DataCombo DCTipoTramite 
      Bindings        =   "FRenovac.frx":0056
      DataSource      =   "AdoTipoTramite"
      Height          =   345
      Left            =   1470
      TabIndex        =   1
      Top             =   420
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   609
      _Version        =   393216
      Text            =   ""
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
   Begin MSAdodcLib.Adodc AdoTipoTramite 
      Height          =   330
      Left            =   210
      Top             =   3885
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
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
      Caption         =   "TipoTramite"
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
   Begin MSAdodcLib.Adodc AdoAux1 
      Height          =   330
      Left            =   210
      Top             =   2310
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
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
      Caption         =   "Aux1"
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
   Begin MSAdodcLib.Adodc AdoAux2 
      Height          =   330
      Left            =   210
      Top             =   2625
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
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
      Caption         =   "Aux2"
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
   Begin MSAdodcLib.Adodc AdoFacturas 
      Height          =   330
      Left            =   210
      Top             =   3255
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
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
      Caption         =   "Facturas"
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
   Begin MSAdodcLib.Adodc AdoFacturasNew 
      Height          =   330
      Left            =   210
      Top             =   3570
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
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
      Caption         =   "FacturasNew"
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
   Begin MSAdodcLib.Adodc AdoSerie 
      Height          =   330
      Left            =   210
      Top             =   2940
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
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
   Begin MSAdodcLib.Adodc AdoAutorizacionOld 
      Height          =   330
      Left            =   210
      Top             =   4515
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
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
      Caption         =   "AutorizacionOld"
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
   Begin MSAdodcLib.Adodc AdoAutorizaciones 
      Height          =   330
      Left            =   210
      Top             =   4830
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
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
      Caption         =   "Autorizaciones"
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
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Autorización"
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
      Left            =   105
      TabIndex        =   14
      Top             =   840
      Width           =   1800
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tipo de Tramite"
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
      Left            =   1470
      TabIndex        =   13
      Top             =   105
      Width           =   5475
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Fecha"
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
      Left            =   105
      TabIndex        =   12
      Top             =   105
      Width           =   1275
   End
   Begin VB.Label LblArchivo 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      Left            =   105
      TabIndex        =   11
      Top             =   1575
      Width           =   8205
   End
End
Attribute VB_Name = "FRenovacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Tipo_Tram As Byte
Dim Tipo_Doc As Byte

Private Sub CheqAutOld_Click()
  If CheqAutOld.value = 1 Then DCAutorizacionOld.Visible = True Else DCAutorizacionOld.Visible = False
End Sub

Private Sub CheqSerie_Click()
  If CheqSerie.value = 1 Then DCSerie.Visible = True Else DCSerie.Visible = False
End Sub

Private Sub CmdCerrar_Click()
  Unload FRenovacion
End Sub

Private Sub CmdGrabar_Click()
Dim NombreArchivo As String
Dim NombreArchivoFa As String
Dim NumFileFa As Integer
  Control_Procesos "F", "Generación de " & DCTipoTramite
  FechaValida MBFechaRegis
  NombreArchivo = RutaSysBases & "\AT\" & LblArchivo.Caption & ".xml"
  NumFile = FreeFile
 'MsgBox NombreArchivo
  Open NombreArchivo For Output As #NumFile   ' Abre el archivo
  Print #NumFile, TxtXML.Text
  Close NumFile
  MsgBox "El Archivo:" & vbCrLf & vbCrLf & NombreArchivo & vbCrLf & vbCrLf & "fue creado con éxito"
End Sub

Private Sub Command1_Click()
Dim AutorizacionNo As String
Dim AutorizacionOld As String
Dim Punto As String
Dim Estab As String
  AutorizacionNo = DCAutorizacion
  AutorizacionOld = DCAutorizacionOld
  Listar_Autorizacion AutorizacionNo
 'MsgBox sSQL & vbCrLf & vbCrLf & AdoFacturas.Recordset.RecordCount & vbCrLf & FechaIni & vbCrLf & FechaFin
  TxtXML = ""
  FechaValida MBFechaRegis
  Tipo_Tram = 0
  Tipo_Doc = 0
  LblArchivo.Caption = "Ninguno"
  With AdoTipoTramite.Recordset
   If .RecordCount Then
      .MoveFirst
      .Find ("Descripcion = '" & DCTipoTramite & "' ")
       If Not .EOF Then
          Tipo_Tram = .Fields("Codigo")
          LblArchivo.Caption = DCTipoTramite & " No " & AutorizacionNo
       End If
   End If
  End With
  Codigo = Ninguno
  If Cadena = "" Then Cadena = Ninguno
 'ENCABEZADO
  TxtXML = TxtXML & "<?xml version="
  TxtXML = TxtXML & Chr(34)
  TxtXML = TxtXML & "1.0"
  TxtXML = TxtXML & Chr(34)
  TxtXML = TxtXML & " encoding="
  TxtXML = TxtXML & Chr(34)
  TxtXML = TxtXML & "UTF-8"
  TxtXML = TxtXML & Chr(34)
  TxtXML = TxtXML & "?>" & vbCrLf
 'Tipo Tramite
  'MsgBox Tipo_Tram
  Select Case Tipo_Tram
    Case 0
         TxtXML = TxtXML & AbrirXML("autorizacion") & vbCrLf
         TxtXML = TxtXML & vbTab & CampoXML("ruc", RUC) & vbCrLf
         TxtXML = TxtXML & vbTab & CampoXML("numAut", Format$(AutorizacionNo, "0000000000")) & vbCrLf
         TxtXML = TxtXML & vbTab & CampoXML("fecha", MBFechaRegis) & vbCrLf
    Case 6, 27 'Solicitud de Autorizacion/Inclusion de punto de emision
         TxtXML = TxtXML & AbrirXML("autorizacion") & vbCrLf
         TxtXML = TxtXML & vbTab & CampoXML("codTipoTra", Tipo_Tram) & vbCrLf
         TxtXML = TxtXML & vbTab & CampoXML("ruc", RUC) & vbCrLf
         TxtXML = TxtXML & vbTab & CampoXML("numAut", Format$(AutorizacionNo, "0000000000")) & vbCrLf
         TxtXML = TxtXML & vbTab & CampoXML("fecha", MBFechaRegis) & vbCrLf
    Case 10, 11
         TxtXML = TxtXML & AbrirXML("autorizacion") & vbCrLf
         TxtXML = TxtXML & vbTab & CampoXML("codTipoTra", Tipo_Tram) & vbCrLf
         TxtXML = TxtXML & vbTab & CampoXML("ruc", RUC) & vbCrLf
         TxtXML = TxtXML & vbTab & CampoXML("fecha", MBFechaRegis) & vbCrLf
         TxtXML = TxtXML & vbTab & CampoXML("numAut", Format$(AutorizacionNo, "0000000000")) & vbCrLf
    Case 7, 8
         TxtXML = TxtXML & AbrirXML("autorizacion") & vbCrLf
         TxtXML = TxtXML & vbTab & CampoXML("codTipoTra", Tipo_Tram) & vbCrLf
         TxtXML = TxtXML & vbTab & CampoXML("ruc", RUC) & vbCrLf
         TxtXML = TxtXML & vbTab & CampoXML("fecha", MBFechaRegis) & vbCrLf
         TxtXML = TxtXML & vbTab & CampoXML("autOld", AutorizacionOld) & vbCrLf
         TxtXML = TxtXML & vbTab & CampoXML("autNew", AutorizacionNo) & vbCrLf
    Case 9
         TxtXML = TxtXML & AbrirXML("autorizacion") & vbCrLf
         TxtXML = TxtXML & vbTab & CampoXML("codTipoTra", Tipo_Tram) & vbCrLf
         TxtXML = TxtXML & vbTab & CampoXML("ruc", RUC) & vbCrLf
         TxtXML = TxtXML & vbTab & CampoXML("numAut", Format$(AutorizacionNo, "0000000000")) & vbCrLf
         TxtXML = TxtXML & vbTab & CampoXML("fecha", MBFechaRegis) & vbCrLf
  End Select
  'Autorizacion
  'Serie
  'TC
  'Vencimiento
  'FactMin
  'FactMax
  'FechaMax
  Select Case Tipo_Tram
   Case 0
       'Detalles
        TxtXML = TxtXML & vbTab & vbTab & AbrirXML("detalles") & vbCrLf
        With AdoAutorizaciones.Recordset
         If .RecordCount > 0 Then
             Do While Not .EOF
                Codigo1 = MidStrg(.Fields("Serie"), 1, 3)
                Codigo2 = MidStrg(.Fields("Serie"), 4, 3)
                Codigo3 = .Fields("Serie")
                NivelNo = .Fields("TC")
                If NivelNo = "NV" Then TipoDoc = "2" Else TipoDoc = "1"
                Contador = 0
                i = .Fields("TMin")
                J = .Fields("TMax")
                TxtXML = TxtXML & vbTab & vbTab & vbTab & AbrirXML("detalle") & vbCrLf
                TxtXML = TxtXML & vbTab & vbTab & vbTab & vbTab & CampoXML("codDoc", TipoDoc) & vbCrLf
                TxtXML = TxtXML & vbTab & vbTab & vbTab & vbTab & CampoXML("estab", Codigo1) & vbCrLf
                TxtXML = TxtXML & vbTab & vbTab & vbTab & vbTab & CampoXML("ptoEmi", Codigo2) & vbCrLf
                TxtXML = TxtXML & vbTab & vbTab & vbTab & vbTab & CampoXML("inicio", i) & vbCrLf
                TxtXML = TxtXML & vbTab & vbTab & vbTab & vbTab & CampoXML("fin", J) & vbCrLf
                TxtXML = TxtXML & vbTab & vbTab & vbTab & CerrarXML("detalle") & vbCrLf
               .MoveNext
             Loop
         End If
        End With
        TxtXML = TxtXML & vbTab & vbTab & CerrarXML("detalles") & vbCrLf
        TxtXML = TxtXML & CerrarXML("autorizacion") & vbCrLf
   Case 6 To 27
       'Detalles
        TxtXML = TxtXML & vbTab & vbTab & AbrirXML("detalles") & vbCrLf
        With AdoAutorizaciones.Recordset
         If .RecordCount > 0 Then
             Do While Not .EOF
                Codigo1 = MidStrg(.Fields("Serie"), 1, 3)
                Codigo2 = MidStrg(.Fields("Serie"), 4, 3)
                Codigo3 = .Fields("Serie")
                NivelNo = .Fields("TC")
                If NivelNo = "NV" Then TipoDoc = "2" Else TipoDoc = "1"
                Contador = 0
                   Select Case Tipo_Tram
                     Case 6, 27
                          i = .Fields("FactMin")
                     Case 7, 8
                          i = .Fields("FactMin")
                          J = .Fields("FactMax")
                     Case 9
                          J = .Fields("FactMax")
                   End Select
                'MsgBox J
                TxtXML = TxtXML & vbTab & vbTab & vbTab & AbrirXML("detalle") & vbCrLf
                TxtXML = TxtXML & vbTab & vbTab & vbTab & vbTab & CampoXML("codDoc", TipoDoc) & vbCrLf
                TxtXML = TxtXML & vbTab & vbTab & vbTab & vbTab & CampoXML("estab", Codigo1) & vbCrLf
                TxtXML = TxtXML & vbTab & vbTab & vbTab & vbTab & CampoXML("ptoEmi", Codigo2) & vbCrLf
                Select Case Tipo_Tram
                  Case 6, 27
                       TxtXML = TxtXML & vbTab & vbTab & vbTab & vbTab & CampoXML("inicio", i) & vbCrLf
                  Case 7, 8
                       TxtXML = TxtXML & vbTab & vbTab & vbTab & vbTab & CampoXML("finOld", i) & vbCrLf
                       TxtXML = TxtXML & vbTab & vbTab & vbTab & vbTab & CampoXML("iniNew", J) & vbCrLf
                  Case 9
                       TxtXML = TxtXML & vbTab & vbTab & vbTab & vbTab & CampoXML("fin", J) & vbCrLf
                End Select
                TxtXML = TxtXML & vbTab & vbTab & vbTab & CerrarXML("detalle") & vbCrLf
               .MoveNext
             Loop
         End If
        End With
        TxtXML = TxtXML & vbTab & vbTab & CerrarXML("detalles") & vbCrLf
        TxtXML = TxtXML & CerrarXML("autorizacion") & vbCrLf
   Case 10
        TxtXML = TxtXML & vbTab & vbTab & AbrirXML("detalles") & vbCrLf
        TxtXML = TxtXML & vbTab & vbTab & vbTab & AbrirXML("detalle") & vbCrLf
        TxtXML = TxtXML & vbTab & vbTab & vbTab & vbTab & CampoXML("codDoc", TipoDoc) & vbCrLf
        TxtXML = TxtXML & vbTab & vbTab & vbTab & vbTab & CampoXML("estab", "2") & vbCrLf
        TxtXML = TxtXML & vbTab & vbTab & vbTab & vbTab & CampoXML("ptoEmi", "2") & vbCrLf
        TxtXML = TxtXML & vbTab & vbTab & vbTab & vbTab & CampoXML("inicio", 1) & vbCrLf
        TxtXML = TxtXML & vbTab & vbTab & vbTab & CerrarXML("detalle") & vbCrLf
        TxtXML = TxtXML & vbTab & vbTab & CerrarXML("detalles") & vbCrLf
        TxtXML = TxtXML & CerrarXML("autorizacion") & vbCrLf
   Case 11
        TxtXML = TxtXML & vbTab & vbTab & AbrirXML("detalles") & vbCrLf
        TxtXML = TxtXML & vbTab & vbTab & vbTab & AbrirXML("detalle") & vbCrLf
        TxtXML = TxtXML & vbTab & vbTab & vbTab & vbTab & CampoXML("codDoc", TipoDoc) & vbCrLf
        TxtXML = TxtXML & vbTab & vbTab & vbTab & vbTab & CampoXML("estab", "1") & vbCrLf
        TxtXML = TxtXML & vbTab & vbTab & vbTab & vbTab & CampoXML("ptoEmi", "2") & vbCrLf
        TxtXML = TxtXML & vbTab & vbTab & vbTab & vbTab & CampoXML("fin", Factura_Hasta) & vbCrLf
        TxtXML = TxtXML & vbTab & vbTab & vbTab & CerrarXML("detalle") & vbCrLf
        TxtXML = TxtXML & vbTab & vbTab & CerrarXML("detalles") & vbCrLf
        TxtXML = TxtXML & CerrarXML("autorizacion") & vbCrLf
  End Select
End Sub

Private Sub DCAutorizacionOld_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCSerie_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT Tipo_Comprobante_Codigo, Descripcion " _
       & "FROM Tipo_Comprobante " _
       & "WHERE Tipo_Comprobante_Codigo BETWEEN 1 AND 8 " _
       & "AND TC = 'TDC' " _
       & "ORDER BY Tipo_Comprobante_Codigo "
  Select_Adodc AdoTipoComprobante, sSQL

  sSQL = "SELECT Autorizacion " _
       & "FROM Catalogo_Lineas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "UNION " _
       & "SELECT Autorizacion " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND TC NOT IN ('C','P') " _
       & "GROUP BY Autorizacion " _
       & "ORDER BY Autorizacion "
  SelectDB_Combo DCAutorizacion, AdoAutorizacion, sSQL, "Autorizacion"
  
  sSQL = "SELECT Autorizacion " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND TC NOT IN ('C','P') " _
       & "GROUP BY Autorizacion " _
       & "ORDER BY Autorizacion "
  SelectDB_Combo DCAutorizacionOld, AdoAutorizacionOld, sSQL, "Autorizacion"
  
  sSQL = "SELECT Serie " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND TC NOT IN ('C','P') " _
       & "AND Serie <> '000000' " _
       & "GROUP BY Serie " _
       & "ORDER BY Serie "
  SelectDB_Combo DCSerie, AdoSerie, sSQL, "Serie"
  
  sSQL = "SELECT Tipo_Comprobante_Codigo AS Codigo, Descripcion " _
       & "FROM Tipo_Comprobante " _
       & "WHERE TC = 'TDT' " _
       & "ORDER BY Codigo "
  SelectDB_Combo DCTipoTramite, AdoTipoTramite, sSQL, "Descripcion"
  
  Listar_Autorizacion
  RatonNormal
End Sub

Private Sub DCAutorizacion_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  PresionoEnter KeyCode
End Sub

Private Sub DCTipoTramite_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub Form_Load()
  CentrarForm FRenovacion
  ConectarAdodc AdoAux1
  ConectarAdodc AdoAux2
  ConectarAdodc AdoSerie
  ConectarAdodc AdoFacturas
  ConectarAdodc AdoFacturasNew
  ConectarAdodc AdoTipoTramite
  ConectarAdodc AdoAutorizacion
  ConectarAdodc AdoAutorizaciones
  ConectarAdodc AdoAutorizacionOld
  ConectarAdodc AdoTipoComprobante
End Sub

Private Sub MBFechaRegis_GotFocus()
  MarcarTexto MBFechaRegis
End Sub

Private Sub MBFechaRegis_LostFocus()
  FechaValida MBFechaRegis
End Sub

Public Sub Listar_Autorizacion(Optional No_Auto As String)
  If No_Auto = "" Then No_Auto = Ninguno
  sSQL = "SELECT Autorizacion,Serie,TC,Vencimiento," _
       & "MIN(Factura) As FactMin, MAX(Factura) As FactMax,MAX(Fecha) As FechaMax " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND TC NOT IN('C','P') "
  If No_Auto <> Ninguno Then sSQL = sSQL & "AND Autorizacion = '" & No_Auto & "' "
  If CheqSerie.value <> 0 Then sSQL = sSQL & "AND Serie = '" & DCSerie & "' "
  sSQL = sSQL & "GROUP BY Autorizacion,Serie,TC,Vencimiento " _
       & "ORDER BY Autorizacion,Serie,TC,Vencimiento DESC "
  Select_Adodc AdoAutorizaciones, sSQL
End Sub
