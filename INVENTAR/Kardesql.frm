VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form KardexSQLs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "INVENTARIO DE KARDEX"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   11775
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11865
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Kardesql.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Kardesql.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Kardesql.frx":0D2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Kardesql.frx":1046
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Kardesql.frx":1920
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Kardesql.frx":1D72
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir del módulo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Consultar"
            Object.ToolTipText     =   "Consultar E/S"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Consultar_Totales"
            Object.ToolTipText     =   "Consultar Totales"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Tipos de Impresion"
            ImageIndex      =   4
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Imp_ES"
                  Text            =   "Entrada/Salida"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Imp_Comprobante"
                  Text            =   "Comprobante"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Detallado"
                  Text            =   "Detallado"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Salidas_Cero"
            Object.ToolTipText     =   "Salidas en cero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Anular"
            Object.ToolTipText     =   "Anular Comprobante E/S"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
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
      Height          =   330
      Left            =   11970
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1470
      Width           =   330
   End
   Begin VB.TextBox TextConcepto 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   12
      Top             =   1890
      Width           =   11565
   End
   Begin MSDataGridLib.DataGrid DGQuery 
      Bindings        =   "Kardesql.frx":25D2
      Height          =   4530
      Left            =   105
      TabIndex        =   11
      Top             =   3255
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   7990
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoProd 
      Height          =   330
      Left            =   450
      Top             =   4650
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "Prod"
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
   Begin MSAdodcLib.Adodc AdoDetKardex 
      Height          =   330
      Left            =   105
      Top             =   7875
      Width           =   8310
      _ExtentX        =   14658
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
      Caption         =   "DetKardex"
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
   Begin MSAdodcLib.Adodc AdoOrden 
      Height          =   330
      Left            =   450
      Top             =   3930
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "Orden"
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
   Begin MSAdodcLib.Adodc AdoProducto 
      Height          =   330
      Left            =   450
      Top             =   4290
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "Producto"
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
   Begin VB.Frame Frame1 
      Height          =   1065
      Left            =   105
      TabIndex        =   0
      Top             =   735
      Width           =   11565
      Begin VB.OptionButton OpcNC 
         Caption         =   "&NC No."
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
         Left            =   5355
         TabIndex        =   16
         Top             =   630
         Width           =   1065
      End
      Begin VB.OptionButton OpcBarra 
         Caption         =   "&Codigo de Barra"
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
         Left            =   6615
         TabIndex        =   13
         Top             =   630
         Width           =   1695
      End
      Begin MSDataListLib.DataCombo DCProducto 
         Bindings        =   "Kardesql.frx":25ED
         DataSource      =   "AdoProducto"
         Height          =   315
         Left            =   4305
         TabIndex        =   6
         Top             =   210
         Width           =   7110
         _ExtentX        =   12541
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DCOrden 
         Bindings        =   "Kardesql.frx":2607
         DataSource      =   "AdoOrden"
         Height          =   315
         Left            =   8505
         TabIndex        =   9
         Top             =   630
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.OptionButton OpcCD 
         Caption         =   "&Diario No."
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
         Left            =   3990
         TabIndex        =   8
         Top             =   630
         Width           =   1275
      End
      Begin VB.OptionButton OpcO 
         Caption         =   "&Guía No."
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
         Left            =   2730
         TabIndex        =   7
         Top             =   630
         Width           =   1170
      End
      Begin VB.OptionButton OpcR 
         Caption         =   "&Responsable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2730
         TabIndex        =   5
         Top             =   210
         Value           =   -1  'True
         Width           =   1485
      End
      Begin MSMask.MaskEdBox MBoxFechaI 
         Height          =   330
         Left            =   1365
         TabIndex        =   2
         Top             =   210
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
      Begin MSMask.MaskEdBox MBoxFechaF 
         Height          =   330
         Left            =   1365
         TabIndex        =   4
         Top             =   630
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
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha I&nicial"
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
         TabIndex        =   1
         Top             =   210
         Width           =   1275
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha &Final"
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
         TabIndex        =   3
         Top             =   630
         Width           =   1275
      End
   End
   Begin MSAdodcLib.Adodc AdoCodigo 
      Height          =   330
      Left            =   420
      Top             =   5040
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "Codigo"
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
   Begin VB.Label LabelTot 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.0"
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
      Left            =   9660
      TabIndex        =   14
      Top             =   7875
      Width           =   2010
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Valor Total"
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
      Left            =   8400
      TabIndex        =   15
      Top             =   7875
      Width           =   1275
   End
End
Attribute VB_Name = "KardexSQLs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ListarDetallesInv(Optional ConTotales As Boolean)
  RatonReloj
  FechaValida MBoxFechaI
  FechaValida MBoxFechaF
  If OpcR.value Then
     Codigo = SinEspaciosIzq(DCProducto.Text)
     TextoValidoVar Codigo
  ElseIf OpcCD.value Then
     Codigo = DCOrden.Text
     TextoValidoVar Codigo, True
  Else
     Codigo = DCOrden.Text
     TextoValidoVar Codigo
  End If
  FechaIni = BuscarFecha(MBoxFechaI.Text)
  FechaFin = BuscarFecha(MBoxFechaF.Text)
  sSQL = "SELECT P.Producto,K.Fecha,K.TP,K.Numero,Cliente,Concepto,Entrada,Salida,K.Orden_No " _
       & "FROM Trans_Kardex As K,Catalogo_Productos As P,Comprobantes As Co,Clientes As C " _
       & "WHERE K.Item = '" & NumEmpresa & "' " _
       & "AND K.Periodo = '" & Periodo_Contable & "' "
  If OpcR.value Then
     sSQL = sSQL & "AND K.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
          & "AND K.Codigo_P = '" & Codigo & "' "
  ElseIf OpcCD.value Then
     sSQL = sSQL & "AND K.TP = 'CD' " _
          & "AND K.Numero = " & Val(Codigo) & " "
  ElseIf OpcNC.value Then
     sSQL = sSQL & "AND K.TP = 'NC' " _
          & "AND K.Numero = " & Val(Codigo) & " "
  ElseIf OpcO.value Then
     sSQL = sSQL & "AND K.Orden_No = '" & Codigo & "' "
  Else
     sSQL = sSQL & "AND K.Codigo_Barra = '" & Codigo & "' "
  End If
  sSQL = sSQL _
       & "AND K.Codigo_Inv = P.Codigo_Inv " _
       & "AND K.TP = Co.TP " _
       & "AND K.Numero = Co.Numero " _
       & "AND K.Item = Co.Item " _
       & "AND K.Item = P.Item " _
       & "AND C.Codigo = Co.Codigo_B " _
       & "AND K.Periodo = Co.Periodo  " _
       & "AND K.Periodo = P.Periodo  " _
       & "ORDER BY K.Orden_No,P.Producto,K.Fecha,K.TP,K.Numero,K.ID "
  Select_Adodc AdoDetKardex, sSQL
  With AdoDetKardex.Recordset
  If .RecordCount > 0 Then
      TextConcepto.Text = .RecordCount & vbCrLf & .fields("Cliente") & vbCrLf & .fields("Concepto")
  Else
     TextConcepto.Text = "No existen datos"
  End If
  End With
  If ConTotales Then
     sSQL = "SELECT K.CodBodega,K.Codigo_Barra,K.Codigo_Inv,P.Producto,K.Fecha,K.TP,K.Numero,Entrada,Salida,K.Orden_No,K.Valor_Unitario,K.Valor_Total "
  Else
     sSQL = "SELECT K.CodBodega,K.Codigo_Barra,K.Codigo_Inv,P.Producto,K.Fecha,K.TP,K.Numero,Entrada,Salida,K.Orden_No "
  End If
  sSQL = sSQL & "FROM Trans_Kardex As K,Catalogo_Productos As P " _
       & "WHERE K.Item = '" & NumEmpresa & "' "
  If OpcR.value Then
     sSQL = sSQL & "AND K.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
          & "AND K.Codigo_P = '" & Codigo & "' "
  ElseIf OpcCD.value Then
     sSQL = sSQL & "AND K.TP = 'CD' " _
          & "AND K.Numero = " & Val(Codigo) & " "
  ElseIf OpcNC.value Then
     sSQL = sSQL & "AND K.TP = 'NC' " _
          & "AND K.Numero = " & Val(Codigo) & " "
  ElseIf OpcO.value Then
     sSQL = sSQL & "AND K.Orden_No = '" & Codigo & "' "
  Else
     sSQL = sSQL & "AND K.Codigo_Barra = '" & Codigo & "' "
  End If
  sSQL = sSQL & "AND K.Item = '" & NumEmpresa & "' " _
       & "AND K.Periodo = '" & Periodo_Contable & "' " _
       & "AND K.Codigo_Inv = P.Codigo_Inv " _
       & "AND K.Item = P.Item " _
       & "AND K.Periodo = P.Periodo  " _
       & "ORDER BY K.Orden_No,K.Codigo_Inv,K.Fecha,K.TP,K.Numero,K.ID "
  Select_Adodc_Grid DGQuery, AdoDetKardex, sSQL
  DGQuery.Visible = False
  Total = 0
  Debe = 0
  Haber = 0
  If ConTotales Then
     With AdoDetKardex.Recordset
      If .RecordCount > 0 Then
          Do While Not .EOF
             If .fields("Salida") > 0 Then Haber = Haber + .fields("Valor_Total")
             Total = Total + .fields("Valor_Total")
            .MoveNext
          Loop
      End If
     End With
  End If
  Total = Total - Haber
  'MsgBox Total
  LabelTot.Caption = Format(Total, "#,##0.00")
  DGQuery.Visible = True
  RatonNormal
End Sub

Private Sub AnularDiarioInv(NumComp1 As Long)
If ClaveContador Then
   NumItem = NumEmpresa
   TipoComp = "CD"
   Mensajes = "Seguro de Anular" & vbCrLf & "El Comprobante No. " & TipoComp & "-" & NumComp1
   Titulo = "Pregunta de Anulacion"
   If BoxMensaje = vbYes Then
     'Actualizamos Comprobante
      sSQL = "UPDATE Comprobantes " _
           & "SET T = '" & Anulado & "', " _
           & "Concepto = '(COMPROBANTE ANULADO) ' & Concepto " _
           & "WHERE Numero = " & NumComp1 & " " _
           & "AND TP = '" & TipoComp & "' " _
           & "AND Periodo = '" & Periodo_Contable & "' " _
           & "AND Item = '" & NumItem & "' "
      Ejecutar_SQL_SP sSQL
     'Actualizar Transacciones
      sSQL = "UPDATE Transacciones " _
           & "SET T = '" & Anulado & "', " _
           & "Debe = 0,Haber = 0,Saldo = 0 " _
           & "WHERE TP = '" & TipoComp & "' " _
           & "AND Numero = " & NumComp1 & " " _
           & "AND Periodo = '" & Periodo_Contable & "' " _
           & "AND Item = '" & NumItem & "' "
      Ejecutar_SQL_SP sSQL
     'Actualizar TransaccionesSC
      sSQL = "DELETE * " _
           & "FROM Trans_SubCtas " _
           & "WHERE TP = '" & TipoComp & "' " _
           & "AND Numero = " & NumComp1 & " " _
           & "AND Periodo = '" & Periodo_Contable & "' " _
           & "AND Item = '" & NumItem & "' "
      Ejecutar_SQL_SP sSQL
     'Actualizar Retencion
      sSQL = "DELETE * " _
           & "FROM Trans_Air " _
           & "WHERE TP = '" & TipoComp & "' " _
           & "AND Numero = " & NumComp1 & " " _
           & "AND Periodo = '" & Periodo_Contable & "' " _
           & "AND Item = '" & NumItem & "' "
      Ejecutar_SQL_SP sSQL
      sSQL = "DELETE * " _
           & "FROM Trans_Compras " _
           & "WHERE TP = '" & TipoComp & "' " _
           & "AND Numero = " & NumComp1 & " " _
           & "AND Periodo = '" & Periodo_Contable & "' " _
           & "AND Item = '" & NumItem & "' "
      Ejecutar_SQL_SP sSQL
      sSQL = "DELETE * " _
           & "FROM Trans_Ventas " _
           & "WHERE TP = '" & TipoComp & "' " _
           & "AND Numero = " & NumComp1 & " " _
           & "AND Periodo = '" & Periodo_Contable & "' " _
           & "AND Item = '" & NumItem & "' "
      Ejecutar_SQL_SP sSQL
      sSQL = "DELETE * " _
           & "FROM Trans_Exportaciones " _
           & "WHERE TP = '" & TipoComp & "' " _
           & "AND Numero = " & NumComp1 & " " _
           & "AND Periodo = '" & Periodo_Contable & "' " _
           & "AND Item = '" & NumItem & "' "
      Ejecutar_SQL_SP sSQL
      sSQL = "DELETE * " _
           & "FROM Trans_Importaciones " _
           & "WHERE TP = '" & TipoComp & "' " _
           & "AND Numero = " & NumComp1 & " " _
           & "AND Periodo = '" & Periodo_Contable & "' " _
           & "AND Item = '" & NumItem & "' "
      Ejecutar_SQL_SP sSQL
     'Actualizar Kardex
      sSQL = "SELECT Codigo_Inv " _
           & "FROM Trans_Kardex " _
           & "WHERE TP = '" & TipoComp & "' " _
           & "AND Numero = " & NumComp1 & " " _
           & "AND Periodo = '" & Periodo_Contable & "' " _
           & "AND Item = '" & NumItem & "' "
      Select_Adodc AdoDetKardex, sSQL
      With AdoDetKardex.Recordset
       If .RecordCount > 0 Then
           Do While Not .EOF
              Codigo = .fields("Codigo_Inv")
              sSQL = "UPDATE Trans_Kardex " _
                   & "SET Procesado = 0 " _
                   & "WHERE Codigo_Inv = '" & Codigo & "' " _
                   & "AND Periodo = '" & Periodo_Contable & "' " _
                   & "AND Item = '" & NumItem & "' "
              Ejecutar_SQL_SP sSQL
             .MoveNext
           Loop
       End If
      End With
      sSQL = "DELETE * " _
           & "FROM Trans_Kardex " _
           & "WHERE TP = '" & TipoComp & "' " _
           & "AND Numero = " & NumComp1 & " " _
           & "AND Periodo = '" & Periodo_Contable & "' " _
           & "AND Item = '" & NumItem & "' "
      Ejecutar_SQL_SP sSQL
   End If
End If
End Sub

Private Sub Imprimir_ES()
  DGQuery.Visible = False
  If OpcR.value Then
     Codigo = SinEspaciosIzq(DCProducto.Text)
     TextoValidoVar Codigo
     TipoDoc = "R"
  ElseIf OpcCD.value Then
     Total = 0
     Codigo = DCOrden.Text
     TextoValidoVar Codigo, True
     NumComp = Val(Codigo)
     TipoDoc = "CD"
  ElseIf OpcNC.value Then
     Total = 0
     Codigo = DCOrden.Text
     TextoValidoVar Codigo, True
     NumComp = Val(Codigo)
     TipoDoc = "NC"
  ElseIf OpcO.value Then
     Codigo = DCOrden.Text
     TextoValidoVar Codigo
     TipoDoc = "G"
     Total = 0
  Else
     Codigo = DCOrden.Text
     TextoValidoVar Codigo
     TipoDoc = "B"
     Total = 0
  End If
  'MsgBox "gsdfgsf"
  Imprimir_Nota_Inventario AdoProd, AdoDetKardex, NumComp, Codigo, TipoDoc, MBoxFechaI, MBoxFechaF, Total
  Mensajes = "Imprimir Copia de Nota de Entrada/Salida"
  Titulo = "COPIA DE NOTA"
  If BoxMensaje = vbYes Then Imprimir_Nota_Inventario AdoProd, AdoDetKardex, NumComp, Codigo, TipoDoc, MBoxFechaI, MBoxFechaF, Total
  ListarDetallesInv
  DGQuery.Visible = True
End Sub

Private Sub Command2_Click()
 Unload KardexSQLs
End Sub

Private Sub Imprimir_Detallado()
  DGQuery.Visible = False
  If OpcR.value Then
     SQLMsg1 = "Beneficiario: " & DCProducto.Text
  ElseIf OpcCD.value Then
     SQLMsg1 = "Diario No. " & DCOrden.Text
  ElseIf OpcNC.value Then
     SQLMsg1 = "Nota de Credito No. " & DCOrden.Text
  ElseIf OpcO.value Then
     SQLMsg1 = "Orden No. " & DCOrden.Text
  Else
     SQLMsg1 = "Codigo de Barra: " & DCOrden.Text
  End If
  ImprimirAdodc AdoDetKardex, 1, 8, , "TOTAL DE INVENTARIO: USD " & Format(Total, "#,##0.00")
  DGQuery.Visible = True
End Sub

Private Sub Imprimir_Comprobante()
  Co.TP = "CD"
  If OpcCD.value Then
     Co.TP = "CD"
  ElseIf OpcNC.value Then
     Co.TP = "NC"
  End If
  Co.Numero = Val(DCOrden.Text)
  Co.Item = NumEmpresa
  Control_Procesos "I", "Imprimio Comprobante de: " & Co.TP & ", No. " & Co.Numero
  ImprimirComprobantesDe False, Co
End Sub

Private Sub DGQuery_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then GenerarDataTexto KardexSQLs, AdoDetKardex
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT TK.Codigo_P & Space(5) & C.Cliente As Nombre_Cta " _
       & "FROM Trans_Kardex As TK,Clientes As C " _
       & "WHERE TK.Item = '" & NumEmpresa & "' " _
       & "AND C.Codigo = TK.Codigo_P " _
       & "GROUP BY TK.Codigo_P,C.Cliente "
  SelectDB_Combo DCProducto, AdoProducto, sSQL, "Nombre_Cta"
  RatonNormal
  MBoxFechaI.SetFocus
End Sub

Private Sub Form_Load()
  CentrarForm KardexSQLs
  ConectarAdodc AdoProd
  ConectarAdodc AdoOrden
  ConectarAdodc AdoProducto
  ConectarAdodc AdoDetKardex
End Sub

Private Sub MBoxFechaF_GotFocus()
  MarcarTexto MBoxFechaF
End Sub

Private Sub MBoxFechaF_LostFocus()
  FechaValida MBoxFechaF
End Sub

Private Sub MBoxFechaI_GotFocus()
  MarcarTexto MBoxFechaI
End Sub

Private Sub MBoxFechaI_LostFocus()
  FechaValida MBoxFechaI, False
End Sub

Private Sub OpcBarra_Click()
  sSQL = "SELECT Codigo_Barra " _
       & "FROM Trans_Kardex " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "GROUP BY Codigo_Barra " _
       & "ORDER BY Codigo_Barra "
  SelectDB_Combo DCOrden, AdoOrden, sSQL, "Codigo_Barra"
  DCOrden.SetFocus
End Sub

Private Sub OpcCD_Click()
  sSQL = "SELECT Numero " _
       & "FROM Trans_Kardex " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TP = 'CD' " _
       & "GROUP BY Numero " _
       & "ORDER BY Numero DESC "
  SelectDB_Combo DCOrden, AdoOrden, sSQL, "Numero"
  DCOrden.SetFocus
End Sub

Private Sub OpcNC_Click()
  sSQL = "SELECT Numero " _
       & "FROM Trans_Kardex " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TP = 'NC' " _
       & "GROUP BY Numero " _
       & "ORDER BY Numero DESC "
  SelectDB_Combo DCOrden, AdoOrden, sSQL, "Numero"
  DCOrden.SetFocus
End Sub

Private Sub OpcO_Click()
  sSQL = "SELECT Orden_No " _
       & "FROM Trans_Kardex " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "GROUP BY Orden_No "
  SelectDB_Combo DCOrden, AdoOrden, sSQL, "Orden_No"
  DCOrden.SetFocus
End Sub

Private Sub OpcR_Click()
  DCProducto.SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.key
      Case "Salir"
           Unload KardexSQLs
      Case "Consultar"
           ListarDetallesInv
      Case "Consultar_Totales"
           ListarDetallesInv True
      Case "Imprimir"
          'Ver submenu
      Case "Salidas_Cero"
           If OpcCD.value Then Co.TP = "CD" Else Co.TP = "NC"
           Co.Numero = Val(DCOrden.Text)
           RutaBackup = Generar_Salidas_Excel(Co)
           Abrir_Excel RutaBackup
           'If Len(RutaBackup) > 1 Then MsgBox "Proceso Terminado"
      Case "Anular"
           If OpcCD.value Then AnularDiarioInv CLng(DCOrden.Text)
    End Select
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.key
      Case "Imp_ES"
           Imprimir_ES
      Case "Imp_Comprobante"
           Imprimir_Comprobante
      Case "Detallado"
           Imprimir_Detallado
    End Select
End Sub
