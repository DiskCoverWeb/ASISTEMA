VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form IngFactCostos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modulos de Gastos"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   Begin MSDataListLib.DataList DLCtas 
      Bindings        =   "FactSubC.frx":0000
      DataSource      =   "AdoCtas"
      Height          =   2310
      Left            =   105
      TabIndex        =   12
      Top             =   1365
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   4075
      _Version        =   393216
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
   Begin VB.Frame Frame1 
      Caption         =   "Moneda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   105
      TabIndex        =   2
      Top             =   525
      Width           =   2745
      Begin VB.OptionButton OpcME 
         Caption         =   "Extranjera"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1365
         TabIndex        =   4
         Top             =   315
         Width           =   1275
      End
      Begin VB.OptionButton OpcMN 
         Caption         =   "Nacional"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   3
         Top             =   315
         Value           =   -1  'True
         Width           =   1170
      End
   End
   Begin VB.TextBox TextMonto 
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
      Left            =   4095
      MaxLength       =   25
      TabIndex        =   8
      Text            =   "0"
      Top             =   945
      Width           =   1695
   End
   Begin VB.TextBox TextFactura 
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
      Left            =   2940
      MaxLength       =   25
      TabIndex        =   7
      Text            =   "0"
      Top             =   945
      Width           =   1170
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Grabar "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6405
      Picture         =   "FactSubC.frx":0016
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   105
      Width           =   960
   End
   Begin VB.CommandButton Command1 
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
      Height          =   855
      Left            =   6405
      Picture         =   "FactSubC.frx":0458
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1995
      Width           =   960
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Eliminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6405
      Picture         =   "FactSubC.frx":0D22
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1050
      Width           =   960
   End
   Begin MSMask.MaskEdBox MBoxFechaI 
      Height          =   330
      Left            =   2310
      TabIndex        =   1
      Top             =   105
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      OLEDragMode     =   1
      OLEDropMode     =   2
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
      OLEDragMode     =   1
      OLEDropMode     =   2
   End
   Begin MSAdodcLib.Adodc AdoSubCtaDet 
      Height          =   330
      Left            =   210
      Top             =   2625
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
      Caption         =   "SubCtaDet"
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
   Begin MSAdodcLib.Adodc AdoCostosTours 
      Height          =   330
      Left            =   210
      Top             =   2310
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
      Caption         =   "CostosTours"
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
   Begin MSAdodcLib.Adodc AdoSubCtaDet1 
      Height          =   330
      Left            =   210
      Top             =   1995
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
      Caption         =   "SubCtaDet1"
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
   Begin MSAdodcLib.Adodc AdoCtas 
      Height          =   330
      Left            =   210
      Top             =   1680
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
      Caption         =   "Ctas"
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
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Desde (DD/MM/AAAA)"
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
      TabIndex        =   0
      Top             =   105
      Width           =   2220
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Monto Factura"
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
      Left            =   4095
      TabIndex        =   6
      Top             =   630
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Factura No."
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
      Left            =   2940
      TabIndex        =   5
      Top             =   630
      Width           =   1170
   End
End
Attribute VB_Name = "IngFactCostos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FilaActual As Integer
Dim SumaSubCta As Currency

Private Sub Command1_Click()
  Unload IngFactCostos
End Sub

Private Sub Command2_Click()
  If Len(DLCtas.Text) > 0 Then
  FechaValida MBoxFechaI
  FechaTexto = MBoxFechaI.Text
  Factura_No = TextFactura.Text
  Total_Factura = TextMonto.Text
  Codigo = Trim$("FA" & Format(Factura_No, "000000"))
  If OpcME.value Then
     Total_FacturaME = TextMonto.Text
     Total_Factura = Total_FacturaME * Dolar
  Else
     Total_Factura = TextMonto.Text
     Total_FacturaME = 0
  End If
  With Co
      .Fecha = FechaTexto
      .Cotizacion = Dolar
      .Usuario = CodigoUsuario
      .Concepto = "Credito de Factura No. " & Format(Factura_No, "000000")
      .CodigoB = Codigo
      .RUC_CI = "0000000000000"
      .TP = "FA"
      .Numero = Factura_No
      .Monto_Total = Total_Factura
      .Efectivo = Total_Factura
       If .T = "" Then .T = Normal
  End With
  SQL1 = "DELETE * " _
       & "FROM Comprobantes " _
       & "WHERE TP = 'FA' " _
       & "AND Numero = " & Factura_No & " " _
       & "AND Item = '" & NumEmpresa & "' "
  ConectarAdoExecute SQL1
  SQL1 = "SELECT * " _
       & "FROM Comprobantes " _
       & "WHERE Item = '" & NumEmpresa & "' "
  SelectAdodc AdoSubCtaDet, SQL1
  SetAddNew AdoSubCtaDet
  SetFields AdoSubCtaDet, "T", Co.T
  SetFields AdoSubCtaDet, "Fecha", Co.Fecha
  SetFields AdoSubCtaDet, "TP", Co.TP
  SetFields AdoSubCtaDet, "Numero", Co.Numero
  'SetFields AdoSubCtaDet, "Beneficiario", Co.CodigoB
  'SetFields AdoSubCtaDet, "RUC_CI", Co.RUC_CI
  SetFields AdoSubCtaDet, "Monto_Total", Round(Co.Monto_Total, 2)
  SetFields AdoSubCtaDet, "Concepto", Co.Concepto
  SetFields AdoSubCtaDet, "Efectivo", Co.Efectivo
  SetFields AdoSubCtaDet, "Cotizacion", Co.Cotizacion
  SetFields AdoSubCtaDet, "CodigoU", Co.Usuario
  SetUpdate AdoSubCtaDet
' Grabamos SubCtas
  SQL1 = "DELETE * " _
       & "FROM Trans_SubCtas " _
       & "WHERE TP = 'FA' " _
       & "AND Numero = " & Factura_No & " " _
       & "AND Item = '" & NumEmpresa & "' "
  ConectarAdoExecute SQL1
  SQL1 = "SELECT * " _
       & "FROM Trans_SubCtas " _
       & "WHERE Item = '" & NumEmpresa & "' "
  SelectAdodc AdoSubCtaDet, SQL1
  SetAddNew AdoSubCtaDet
  SetFields AdoSubCtaDet, "T", Normal
  SetFields AdoSubCtaDet, "TC", "G"
  SetFields AdoSubCtaDet, "Cta", SinEspaciosIzq(DLCtas.Text)
  SetFields AdoSubCtaDet, "Codigo", Co.CodigoB
  SetFields AdoSubCtaDet, "Fecha", Co.Fecha
  SetFields AdoSubCtaDet, "Fecha_V", Co.Fecha
  SetFields AdoSubCtaDet, "TP", Co.TP
  SetFields AdoSubCtaDet, "Numero", Co.Numero
  SetFields AdoSubCtaDet, "Factura", Co.Numero
  SetFields AdoSubCtaDet, "Creditos", Total_Factura
  SetUpdate AdoSubCtaDet
  SQL1 = "DELETE * " _
       & "FROM Catalogo_SubCtas " _
       & "WHERE Codigo = '" & Codigo & "' " _
       & "AND Item = '" & NumEmpresa & "' "
  ConectarAdoExecute SQL1
  sSQL = "SELECT * " _
       & "FROM Catalogo_SubCtas " _
       & "WHERE Codigo = '" & Codigo & "' " _
       & "AND TC = 'F' " _
       & "AND Item = '" & NumEmpresa & "' "
  SelectAdodc AdoSubCtaDet, sSQL
  SetAddNew AdoSubCtaDet
  SetFields AdoSubCtaDet, "TC", "F"
  SetFields AdoSubCtaDet, "Codigo", Codigo
  SetFields AdoSubCtaDet, "Beneficiario", "Factura No. " & Format(Factura_No, "000000")
  SetFields AdoSubCtaDet, "Ciudad", Ninguno
  SetFields AdoSubCtaDet, "Direccion", Ninguno
  SetFields AdoSubCtaDet, "RUC_CI", "0000000000000"
  SetFields AdoSubCtaDet, "Telefono", "00000000"
  SetFields AdoSubCtaDet, "Celular", "00000000"
  SetFields AdoSubCtaDet, "FAX", "00000000"
  SetUpdate AdoSubCtaDet
  End If
End Sub

Private Sub Command3_Click()
  FechaValida MBoxFechaI
  FechaTexto = MBoxFechaI.Text
  Factura_No = TextFactura.Text
  Total_Factura = TextMonto.Text
  Mensajes = "Seguro de Eliminar el Costo"
  Titulo = "Pregunta de Grabación"
  If BoxMensaje = 6 Then
     SQL1 = "DELETE * " _
          & "FROM Comprobantes " _
          & "WHERE TP = 'FA' " _
          & "AND Numero = " & Factura_No & " "
     ConectarAdoExecute SQL1
     SQL1 = "DELETE * " _
          & "FROM Trans_SubCtas " _
          & "WHERE TP = 'FA' " _
          & "AND Numero = " & Factura_No & " "
     ConectarAdoExecute SQL1
  End If
End Sub

Private Sub DLCtas_DblClick()
  SiguienteControl
End Sub

Private Sub DLCtas_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT Codigo & Space(10) & Cuenta As Nombre_Cta " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE TC = 'G' " _
       & "AND DG = 'D' " _
       & "ORDER BY Codigo "
  SelectDBList DLCtas, AdoCtas, sSQL, "Nombre_Cta"
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm IngFactCostos
  ConectarAdodc AdoCtas
  ConectarAdodc AdoSubCtaDet
  ConectarAdodc AdoSubCtaDet1
  ConectarAdodc AdoCostosTours
End Sub

Private Sub MBoxFechaI_LostFocus()
  FechaValida MBoxFechaI
End Sub

Private Sub TextFactura_GotFocus()
  MarcarTexto TextFactura
End Sub

Private Sub TextFactura_LostFocus()
  TextoValido TextFactura, True
End Sub

Private Sub TextMonto_GotFocus()
  MarcarTexto TextMonto
End Sub

Private Sub TextMonto_LostFocus()
  TextoValido TextMonto, True
End Sub

