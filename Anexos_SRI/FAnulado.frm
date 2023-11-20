VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form FAnulados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ANULACION DE COMPROBANTES"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc AdoTipoComprobante 
      Height          =   330
      Left            =   3255
      Top             =   1890
      Visible         =   0   'False
      Width           =   2745
      _ExtentX        =   4842
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   2625
      MaxLength       =   7
      TabIndex        =   4
      Text            =   "0000001"
      ToolTipText     =   $"FAnulado.frx":0000
      Top             =   1155
      Width           =   1170
   End
   Begin VB.CommandButton CmdCerrar 
      Caption         =   "&Salir"
      Height          =   765
      Left            =   5040
      Picture         =   "FAnulado.frx":00A3
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Salir"
      Top             =   840
      Width           =   990
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Height          =   750
      Left            =   3990
      Picture         =   "FAnulado.frx":04E5
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Grabar"
      Top             =   840
      Width           =   960
   End
   Begin VB.TextBox TxtNumSerieUno 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   105
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "001"
      ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al código del establecimiento"
      Top             =   1155
      Width           =   645
   End
   Begin VB.TextBox TxtNumSerieDos 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   735
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "001"
      ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al punto dde emisión"
      Top             =   1155
      Width           =   645
   End
   Begin VB.TextBox TxtNumSerietres1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   1470
      MaxLength       =   7
      TabIndex        =   3
      Text            =   "0000001"
      ToolTipText     =   $"FAnulado.frx":07EF
      Top             =   1155
      Width           =   1170
   End
   Begin VB.TextBox TxtNumAutor 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   105
      MaxLength       =   10
      TabIndex        =   5
      Text            =   "0000000001"
      Top             =   1890
      Width           =   1305
   End
   Begin MSMask.MaskEdBox MBFechaRegis 
      Height          =   330
      Left            =   1470
      TabIndex        =   6
      ToolTipText     =   $"FAnulado.frx":0892
      Top             =   1890
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
   Begin MSDataListLib.DataCombo DCTipoComprobante 
      Bindings        =   "FAnulado.frx":091A
      DataSource      =   "AdoTipoComprobante"
      Height          =   360
      Left            =   105
      TabIndex        =   0
      ToolTipText     =   $"FAnulado.frx":093B
      Top             =   420
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
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
      Left            =   1470
      TabIndex        =   13
      Top             =   1575
      Width           =   1275
   End
   Begin VB.Label Label1 
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
      TabIndex        =   12
      Top             =   1575
      Width           =   1275
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Rago de Comprobantes"
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
      TabIndex        =   11
      Top             =   840
      Width           =   2325
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Serie:"
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
      TabIndex        =   10
      Top             =   840
      Width           =   1275
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tipo de Comprobantes"
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
      TabIndex        =   9
      Top             =   105
      Width           =   5895
   End
End
Attribute VB_Name = "FAnulados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdCerrar_Click()
  Unload FAnulados
End Sub

Private Sub CmdGrabar_Click()
  RatonReloj
 'Busca que sea igual a la Descripcion
  ID_Trans = Maximo_De("Trans_Anulados", "ID")
  Factura_Desde = CTNumero(TxtNumSerietres1)
  Factura_Hasta = CTNumero(TxtNumSerietres2)
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
  SetAdoAddNew "Trans_Anulados"
  SetAdoFields "T", Normal
  SetAdoFields "TipoComprobante", Codigo
  SetAdoFields "Establecimiento", TxtNumSerieUno
  SetAdoFields "PuntoEmision", TxtNumSerieDos
  SetAdoFields "Secuencial1", Factura_Desde
  SetAdoFields "Secuencial2", Factura_Hasta
  SetAdoFields "Autorizacion", TxtNumAutor
  SetAdoFields "FechaAnulacion", MBFechaRegis
  SetAdoFields "ID", ID_Trans
  SetAdoUpdate
  RatonNormal
  Titulo = "GRABAR ANULADOS"
  Mensajes = "Proceso Terminado, Desea Ingresar Otro Dato"
  If BoxMensaje = vbYes Then
     DCTipoComprobante.SetFocus
  Else
     Unload FAnulados
  End If
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT Tipo_Comprobante_Codigo, Descripcion " _
       & "FROM Tipo_Comprobante " _
       & "WHERE TC = 'TDC' " _
       & "ORDER BY Tipo_Comprobante_Codigo "
  SelectDB_Combo DCTipoComprobante, AdoTipoComprobante, sSQL, "Descripcion"
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm FAnulados
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
