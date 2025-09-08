VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form ResumenKardex 
   Caption         =   "RESUMEN DE EXISTENCIAS"
   ClientHeight    =   10620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10620
   ScaleWidth      =   20250
   WindowState     =   1  'Minimized
   Begin ComctlLib.Toolbar TBKardex 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   20250
      _ExtentX        =   35719
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
            Object.ToolTipText     =   "Salir del Modulo"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Stock"
            Object.ToolTipText     =   "Resumen de Existecia"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Stock_1"
            Object.ToolTipText     =   "Resumen de Existencia Agrupado"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Lote"
            Object.ToolTipText     =   "Resumen de Existencia por Lotes"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Barras"
            Object.ToolTipText     =   "Resumen en Codigos de Barra"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "Imprimir"
            Description     =   ""
            Object.ToolTipText     =   "Imprimir Resultado"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Excel"
            Object.ToolTipText     =   "Enviar a Excel el resultado"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
      EndProperty
      Begin VB.Frame Frame1 
         Height          =   645
         Left            =   4200
         TabIndex        =   0
         Top             =   0
         Width           =   11145
         Begin VB.TextBox TxtMonto 
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
            Height          =   330
            Left            =   6510
            MultiLine       =   -1  'True
            TabIndex        =   25
            Text            =   "ResumenK.frx":0000
            Top             =   210
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.CheckBox CheqMonto 
            Caption         =   "MONTO"
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
            TabIndex        =   24
            Top             =   210
            Width           =   1065
         End
         Begin VB.CheckBox CheqExist 
            Caption         =   "Listar Catalogo Completo"
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
            Left            =   8085
            TabIndex        =   5
            Top             =   210
            Width           =   2535
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
            Left            =   10710
            Picture         =   "ResumenK.frx":0007
            TabIndex        =   19
            Top             =   210
            Width           =   330
         End
         Begin MSMask.MaskEdBox MBoxFechaF 
            Height          =   330
            Left            =   3990
            TabIndex        =   4
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
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fecha Inicial"
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
            Caption         =   "Fecha Final"
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
            TabIndex        =   3
            Top             =   210
            Width           =   1275
         End
      End
   End
   Begin VB.Frame FrmProducto 
      Caption         =   "Por:"
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
      Left            =   2205
      TabIndex        =   30
      Top             =   1155
      Visible         =   0   'False
      Width           =   13770
      Begin VB.OptionButton OpcLote 
         Caption         =   "Lote"
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
         Left            =   4200
         TabIndex        =   35
         Top             =   210
         Width           =   750
      End
      Begin VB.OptionButton OpcMarca 
         Caption         =   "Marca"
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
         Left            =   3150
         TabIndex        =   34
         Top             =   210
         Width           =   855
      End
      Begin VB.OptionButton OpcBarra 
         Caption         =   "Codigo Barra"
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
         TabIndex        =   33
         Top             =   210
         Width           =   1485
      End
      Begin VB.OptionButton OpcProducto 
         Caption         =   "Producto"
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
         TabIndex        =   31
         Top             =   210
         Value           =   -1  'True
         Width           =   1170
      End
      Begin MSDataListLib.DataCombo DCTipoBusqueda 
         Bindings        =   "ResumenK.frx":08D1
         DataSource      =   "AdoBusqueda"
         Height          =   315
         Left            =   5145
         TabIndex        =   32
         Top             =   210
         Width           =   8520
         _ExtentX        =   15028
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   "DC"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame FrmSubModulo 
      Caption         =   "Su-Modulo de:"
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
      Left            =   2205
      TabIndex        =   26
      Top             =   2625
      Visible         =   0   'False
      Width           =   13770
      Begin VB.OptionButton OpcGasto 
         Caption         =   "Centro de Costo"
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
         TabIndex        =   28
         Top             =   210
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton OpcCxP 
         Caption         =   "CxP/Proveedores"
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
         TabIndex        =   27
         Top             =   210
         Width           =   1905
      End
      Begin MSDataListLib.DataCombo DCSubModulo 
         Bindings        =   "ResumenK.frx":08EB
         DataSource      =   "AdoSubModulo"
         Height          =   360
         Left            =   3990
         TabIndex        =   29
         Top             =   210
         Width           =   9675
         _ExtentX        =   17066
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   16777215
         Text            =   "DC"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame FrmCuenta 
      Caption         =   "Cuenta de:"
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
      Left            =   2205
      TabIndex        =   20
      Top             =   1890
      Visible         =   0   'False
      Width           =   13770
      Begin VB.OptionButton OpcCosto 
         Caption         =   "Costo"
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
         TabIndex        =   22
         Top             =   210
         Width           =   960
      End
      Begin VB.OptionButton OpcInv 
         Caption         =   "Inventario"
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
         TabIndex        =   21
         Top             =   210
         Value           =   -1  'True
         Width           =   1275
      End
      Begin MSDataListLib.DataCombo DCCtaInv 
         Bindings        =   "ResumenK.frx":0906
         DataSource      =   "AdoCtaInv"
         Height          =   315
         Left            =   2520
         TabIndex        =   23
         Top             =   210
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   "DC"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CheckBox CheqCtaInv 
      Caption         =   "TIPO DE CTA."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   105
      TabIndex        =   12
      Top             =   1995
      Width           =   2010
   End
   Begin VB.CheckBox CheqSubMod 
      Caption         =   "POR SUBMODULO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   105
      TabIndex        =   11
      Top             =   2730
      Width           =   2010
   End
   Begin MSDataListLib.DataCombo DCTInv 
      Bindings        =   "ResumenK.frx":091E
      DataSource      =   "AdoTInv"
      Height          =   315
      Left            =   10710
      TabIndex        =   7
      Top             =   735
      Visible         =   0   'False
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   "DC"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox CheqProducto 
      Caption         =   "PRODUCTO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   105
      TabIndex        =   10
      Top             =   1260
      Width           =   1905
   End
   Begin VB.CheckBox CheqBod 
      Caption         =   "BODEGA"
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
      TabIndex        =   8
      Top             =   735
      Width           =   2010
   End
   Begin VB.CheckBox CheqGrupo 
      Caption         =   "TIPO GRUPO"
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
      Left            =   9240
      TabIndex        =   6
      Top             =   735
      Width           =   1485
   End
   Begin MSDataGridLib.DataGrid DGQuery 
      Bindings        =   "ResumenK.frx":0934
      Height          =   4635
      Left            =   105
      TabIndex        =   15
      Top             =   3360
      Width           =   13560
      _ExtentX        =   23918
      _ExtentY        =   8176
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
   Begin MSAdodcLib.Adodc AdoDetKardex 
      Height          =   330
      Left            =   210
      Top             =   9765
      Width           =   6630
      _ExtentX        =   11695
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
   Begin MSAdodcLib.Adodc AdoBusqueda 
      Height          =   330
      Left            =   735
      Top             =   6300
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
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
      Caption         =   "Busqueda"
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
   Begin MSAdodcLib.Adodc AdoTInv 
      Height          =   330
      Left            =   735
      Top             =   4725
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
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
      Caption         =   "TInv"
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
   Begin MSDataListLib.DataCombo DCBodega 
      Bindings        =   "ResumenK.frx":094F
      DataSource      =   "AdoBodega"
      Height          =   315
      Left            =   2205
      TabIndex        =   9
      Top             =   735
      Visible         =   0   'False
      Width           =   6945
      _ExtentX        =   12250
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   "BODEGA"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoBodega 
      Height          =   330
      Left            =   735
      Top             =   5355
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
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
      Caption         =   "Bodega"
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   735
      Top             =   5985
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
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
      Caption         =   "Aux"
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
   Begin MSAdodcLib.Adodc AdoSubModulo 
      Height          =   330
      Left            =   735
      Top             =   5040
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
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
      Caption         =   "SubModulo"
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
   Begin MSAdodcLib.Adodc AdoCtaInv 
      Height          =   330
      Left            =   735
      Top             =   5670
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
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
      Caption         =   "CtaInv"
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
   Begin VB.Label LabelStock 
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
      Left            =   8505
      TabIndex        =   16
      Top             =   9765
      Width           =   2010
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Stock Total"
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
      Left            =   7245
      TabIndex        =   17
      Top             =   9765
      Width           =   1275
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   16380
      Top             =   945
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ResumenK.frx":0967
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ResumenK.frx":0C81
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ResumenK.frx":0F9B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ResumenK.frx":12B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ResumenK.frx":15CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ResumenK.frx":18E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ResumenK.frx":1C03
            Key             =   ""
         EndProperty
      EndProperty
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
      Left            =   11865
      TabIndex        =   14
      Top             =   9765
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
      Left            =   10605
      TabIndex        =   13
      Top             =   9765
      Width           =   1275
   End
End
Attribute VB_Name = "ResumenKardex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim QTipoInv As Boolean
Dim CNivel_1 As Boolean
Dim GrupoInv As String

'''Public Sub Grabar_Stock_Barra()
'''  'If Codigo1 <> Ninguno Then
'''     If Contador <= 0 Then Contador = 1
'''     SetAdoAddNew "Saldo_Diarios"
'''     SetAdoFields "TP", "INVE"
'''     SetAdoFields "Comprobante", CodigoL
'''     SetAdoFields "CodigoC", Ninguno
'''     SetAdoFields "Saldo_Anterior", PFil
'''     SetAdoFields "Saldo_Actual", PCol
'''     SetAdoFields "Ingresos", Entrada
'''     SetAdoFields "Egresos", Salida
'''     SetAdoFields "Recibo", Codigo
'''     SetAdoFields "Cta", Codigo1
'''     SetAdoFields "Total", Precio / Contador
'''     SetAdoFields "Numero", Numero
'''     SetAdoFields "Item", NumEmpresa
'''     SetAdoFields "CodigoU", CodigoUsuario
'''     SetAdoUpdate
'''  Numero = Numero + 1
'''  'End If
'''End Sub

Private Sub Imprimir()
 'impceros=true
 'MsgBox QTipoInv
  DGQuery.Visible = False
  MensajeEncabData = "R E S U M E N    D E    E X I S T E N C I A S"
  SQLMsg1 = " "
  If CheqBod.value = 1 Then SQLMsg1 = "POR BODEGA " & UCase(DCBodega)
  If CheqSubMod.value = 1 Then
     If SQLMsg1 <> "" Then
        SQLMsg1 = SQLMsg1 & " Y DE " & UCase(DCSubModulo)
     Else
        SQLMsg1 = "DE " & UCase(DCSubModulo)
     End If
  End If
  SQLMsg2 = "Desde: " & MBoxFechaI.Text & " hasta: " & MBoxFechaF.Text
  Cuadricula = True
  
  If Opcion = 1 Then
     Imprimir_Resumen_Kardex AdoDetKardex, 10, QTipoInv, Total
  Else
     Imprimir_Resumen_Barra AdoDetKardex, 8
  End If
  DGQuery.Visible = True
End Sub

Private Sub CheqBod_Click()
  If CheqBod.value Then
     sSQL = "SELECT (CodBod + ' - ' + Bodega) As Bodegas " _
          & "FROM Catalogo_Bodegas " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "ORDER BY CodBod "
     SelectDB_Combo DCBodega, AdoBodega, sSQL, "Bodegas"
     
     DCBodega.Visible = True
     DCBodega.SetFocus
  Else
     DCBodega.Visible = False
  End If
End Sub

Private Sub CheqCtaInv_Click()
  If CheqCtaInv.value Then
     Listar_Por_Tipo_Cta
     FrmCuenta.Visible = True
     OpcInv.SetFocus
  Else
     FrmCuenta.Visible = False
  End If
End Sub

Private Sub CheqGrupo_Click()
  If CheqGrupo.value Then
     DCTInv.Visible = True
     DCTInv.SetFocus
  Else
     DCTInv.Visible = False
  End If
End Sub

Private Sub CheqMonto_Click()
  If CheqMonto.value Then
     TxtMonto.Visible = True
     TxtMonto.SetFocus
  Else
     TxtMonto.Visible = False
  End If
End Sub

Private Sub CheqProducto_Click()
  If CheqProducto.value Then
    'Listar_Por_Producto
     FrmProducto.Visible = True
     OpcProducto.SetFocus
  Else
     FrmProducto.Visible = False
  End If
End Sub

Private Sub CheqSubMod_Click()
  If CheqSubMod.value Then
     Listar_Por_Tipo_SubModulo
     FrmSubModulo.Visible = True
     OpcGasto.SetFocus
  Else
     FrmSubModulo.Visible = False
  End If
End Sub

Private Sub Command2_Click()
  Unload ResumenKardex
End Sub

Private Sub Stock(StockSuperior As Boolean)
  DGQuery.Visible = False
  QTipoInv = False
  Control_Procesos "I", "Proceso Stock de Inventario, del " & MBoxFechaI & " al " & MBoxFechaF
''  If CheqProducto.value Then
''     Opcion = 2
''    'Cta As Codigo_No
'''''     StockInventBarra
''
''     sSQL = "SELECT Recibo As Serie_No,Comprobante As Detalle," _
''          & "Total As Promedio,Saldo_Anterior As Saldo_Ant,Ingresos As Entradas," _
''          & "Egresos As Salidas,Saldo_Actual As Stock_Act " _
''          & "FROM Saldo_Diarios " _
''          & "WHERE Item = '" & NumEmpresa & "' " _
''          & "AND CodigoU = '" & CodigoUsuario & "' " _
''          & "AND TP = 'INVE' "
''     If CheqMonto.value = 1 Then
''        sSQL = sSQL & "AND Saldo_Actual = " & Val(TxtMonto.Text) & " "
''     Else
''        sSQL = sSQL & "AND Saldo_Actual <> 0 "
''     End If
''     If (OpcProducto.value = 1) And (Codigo3 <> "Todos") Then sSQL = sSQL & "AND Recibo = '" & Codigo3 & "' "
''     If CheqExist.value = 1 Then sSQL = sSQL & "AND Saldo_Actual <> 0 "
''     sSQL = sSQL & "ORDER BY Numero "
''  Else
     Opcion = 1
     Reporte_Resumen_Existencias_SP MBoxFechaI, MBoxFechaF, Cod_Bodega
     
     Progreso_Barra.Mensaje_Box = "Procesando Resumen de Existencia"
     Progreso_Iniciar
    
    'StockInvent StockSuperior
     sSQL = "SELECT TC,Codigo_Inv,Producto,Unidad,Stock_Anterior,Entradas,Salidas,Stock_Actual, Promedio As Costo_Unit,Valor_Total, 0 As Diferencias, Ubicacion " _
          & "FROM Catalogo_Productos As CP " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & SQL_Tipo_Busqueda
     If CheqGrupo.value <> 0 Then sSQL = sSQL & "AND Codigo_Inv LIKE '" & Buscar_Grupo_Inventario & "%' "
     sSQL = sSQL & "ORDER BY Codigo_Inv "
''  End If
  SQLDec = "Costo_Unit " & CStr(Dec_Costo) & "|Total 2|."
  Select_Adodc_Grid DGQuery, AdoDetKardex, sSQL, SQLDec
  'MsgBox Opcion & vbCrLf & SQLDec & vbCrLf & Cod_Bodega
  Total = 0
  Debitos = 0
  Creditos = 0
  DGQuery.Visible = False
  With AdoDetKardex.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          If OpcProducto.value <> 1 Then
             If .fields("TC") <> "I" Then
                 Debitos = Debitos + Redondear(.fields("Entradas") * .fields("Costo_Unit"), 2)
                 Creditos = Creditos + Redondear(.fields("Salidas") * .fields("Costo_Unit"), 2)
                 Total = Total + Redondear(.fields("Total"), 2)
             End If
          End If
         .MoveNext
       Loop
   End If
  End With
  DGQuery.Visible = True
 'Total = Debitos - Creditos
  LabelTot.Caption = Format(Total, "#,##0.00")
  DGQuery.Visible = True
  RatonNormal
End Sub

Private Sub DCTInv_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCTInv_LostFocus()
   Listar_Por_Producto
End Sub

Private Sub Form_Activate()
  RatonReloj
  'Mayorizar_Inventario_SP
  QTipoInv = False
  
  sSQL = "SELECT Codigo_Inv, Producto " _
       & "FROM Catalogo_Productos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'I' " _
       & "AND INV <> 0 " _
       & "ORDER BY Codigo_Inv "
  SelectDB_Combo DCTInv, AdoTInv, sSQL, "Producto"

'''  sSQL = "UPDATE Catalogo_Productos " _
'''       & "SET INV = " & Val(adTrue) & " " _
'''       & "WHERE TC = 'I' " _
'''       & "AND Item = '" & NumEmpresa & "' "
'''  Ejecutar_SQL_SP sSQL
  
  sSQL = "SELECT TC, Codigo_Inv, Producto, Unidad, Stock_Anterior, Entradas, Salidas, Stock_Actual, Promedio, PVP, Valor_Total " _
       & "FROM Catalogo_Productos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo_Inv "
  Select_Adodc_Grid DGQuery, AdoDetKardex, sSQL
  MBoxFechaI.Text = FechaSistema
  MBoxFechaF.Text = FechaSistema
  
  DGQuery.Height = MDI_Y_Max - DGQuery.Top - 500
  DGQuery.width = MDI_X_Max - DGQuery.Left
  Label3.Top = DGQuery.Top + DGQuery.Height + 50
  Label5.Top = DGQuery.Top + DGQuery.Height + 50
  LabelTot.Top = DGQuery.Top + DGQuery.Height + 50
  LabelStock.Top = DGQuery.Top + DGQuery.Height + 50
  AdoDetKardex.Top = DGQuery.Top + DGQuery.Height + 50
  
  ResumenKardex.WindowState = vbMaximized
  MBoxFechaI.SetFocus
  RatonNormal
End Sub

Private Sub Form_Load()
  RatonReloj
  ConectarAdodc AdoAux
  ConectarAdodc AdoTInv
  ConectarAdodc AdoBodega
  ConectarAdodc AdoCtaInv
  ConectarAdodc AdoBusqueda
  ConectarAdodc AdoDetKardex
  ConectarAdodc AdoSubModulo
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
  FechaValida MBoxFechaI
End Sub

Private Sub OpcBarra_Click()
   Listar_Por_Producto
End Sub

Private Sub OpcCosto_Click()
    Listar_Por_Tipo_Cta
End Sub

Private Sub OpcCxP_Click()
    Listar_Por_Tipo_SubModulo
End Sub

Private Sub OpcGasto_Click()
    Listar_Por_Tipo_SubModulo
End Sub

Private Sub OpcInv_Click()
    Listar_Por_Tipo_Cta
End Sub

Private Sub OpcLote_Click()
    Listar_Por_Producto
End Sub

Private Sub OpcMarca_Click()
    Listar_Por_Producto
End Sub

Private Sub OpcProducto_Click()
  Listar_Por_Producto
End Sub

Private Sub TBKardex_ButtonClick(ByVal Button As ComctlLib.Button)
    TextoValido TxtMonto, True
    FechaValida MBoxFechaI
    FechaValida MBoxFechaF
    FechaInicial = MBoxFechaI
    FechaFinal = MBoxFechaF
    FechaIni = BuscarFecha(FechaInicial)
    FechaFin = BuscarFecha(FechaFinal)
    
    sSQL = "SELECT * " _
         & "FROM Fechas_Balance " _
         & "WHERE Detalle = 'Inventario' " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' "
    Select_Adodc AdoAux, sSQL
    If AdoAux.Recordset.RecordCount <= 0 Then
       AdoAux.Recordset.AddNew
       AdoAux.Recordset.fields("Detalle") = "Inventario"
       AdoAux.Recordset.fields("Item") = NumEmpresa
       AdoAux.Recordset.fields("Periodo") = Periodo_Contable
       AdoAux.Recordset.fields("Cerrado") = adFalse
    End If
    AdoAux.Recordset.fields("Fecha_Inicial") = MBoxFechaI
    AdoAux.Recordset.fields("Fecha_Final") = MBoxFechaF
    AdoAux.Recordset.Update
    
    Cod_Bodega = Ninguno
    If CheqBod.value <> 0 Then
       If Len(DCBodega) >= 1 Then Cod_Bodega = SinEspaciosIzq(DCBodega)
    End If
    
   'MsgBox Button.key
    Select Case Button.key
      Case "Salir"
           Unload ResumenKardex
      Case "Imprimir"
           Imprimir
      Case "Stock"
           Stock True
      Case "Stock_1"
           Stock False
           'Stock_1
      Case "Lote"
           Resumen_Lote
      Case "Barras"
           Resumen_Barras
      Case "Excel"
           DGQuery.Visible = False
           Exportar_AdoDB_Excel AdoDetKardex.Recordset, "Existencia " & BuscarFecha(MBoxFechaI) & " al " & BuscarFecha(MBoxFechaF)
          'GenerarDataTexto ResumenKardex, AdoDetKardex
           DGQuery.Visible = True
    End Select
End Sub

Private Sub TxtMonto_GotFocus()
  MarcarTexto TxtMonto
End Sub

Private Sub TxtMonto_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtMonto_LostFocus()
  TextoValido TxtMonto, True
End Sub

Public Sub StockInvent(StockSuperior As Boolean)
Dim CodigoAux  As String
    RatonReloj
    MiTiempo = Time
    DGQuery.Visible = False
    Progreso_Barra.Mensaje_Box = "Procesando Resumen de Existencia"
    Progreso_Iniciar
    If CheqBod.value = 0 Then Cod_Bodega = Ninguno Else Cod_Bodega = SinEspaciosIzq(DCBodega)
    Reporte_Resumen_Existencias_SP MBoxFechaI, MBoxFechaF, Cod_Bodega
   'SQLDec = "Promedio " & CStr(Dec_Costo) & "|Valor_Total 2|."
                                              
    sSQL = "SELECT TC,Codigo_Inv,Stock_Anterior,Entradas,Salidas,Stock_Actual,Promedio,Valor_Total,ID " _
         & "FROM Catalogo_Productos " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Stock_Actual <> 0 " _
         & "AND TC = 'P' " _
         & "ORDER BY Codigo_Inv "
    Select_Adodc AdoAux, sSQL
    With AdoAux.Recordset
     If .RecordCount > 0 Then
         Progreso_Barra.Valor_Maximo = .RecordCount
         Do While Not .EOF
            Total = 0
            ID_Trans = .fields("ID")
            CodigoAux = .fields("Codigo_Inv")
            Stock_Inv = .fields("Stock_Actual")
            If Stock_Inv <> 0 Then
               Progreso_Barra.Mensaje_Box = "Existencia de " & CodigoAux
               Progreso_Esperar
               Total = 0
               Valor_Prom = 0
               CodigoB = CodigoAux
               sSQL = "SELECT TOP 1 Costo, Total " _
                    & "FROM Trans_Kardex " _
                    & "WHERE T <> '" & Anulado & "' " _
                    & "AND Codigo_Inv = '" & CodigoAux & "' " _
                    & "AND Item = '" & NumEmpresa & "' " _
                    & "AND Periodo = '" & Periodo_Contable & "' " _
                    & "AND Fecha <= #" & FechaFin & "# " _
                    & "ORDER BY Fecha DESC, Entrada, Salida DESC, TP DESC, Numero DESC, ID DESC "
               Select_Adodc AdoDetKardex, sSQL
              'MsgBox sSQL
               If AdoDetKardex.Recordset.RecordCount > 0 Then
                  Total = Redondear(AdoDetKardex.Recordset.fields("Total"), 2)
                  'Valor_Prom = Redondear(AdoDetKardex.Recordset.Fields("Costo"), Dec_Costo)
                  Valor_Prom = Redondear(Total / Stock_Inv, Dec_Costo)
               End If
              'MsgBox Total & vbCrLf & vbCrLf & sSQL
               'Total = Redondear(.Fields("Stock_Actual") * Valor_Prom, 2)
              .fields("Valor_Total") = Total
              .fields("Promedio") = Valor_Prom
              .Update
'''               MsgBox .Fields("Stock_Anterior") & vbCrLf _
'''                    & .Fields("Entradas") & vbCrLf _
'''                    & .Fields("Salidas") & vbCrLf _
'''                    & .Fields("Stock_Actual") & vbCrLf _
'''                    & .Fields("Promedio") & vbCrLf _
'''                    & .Fields("Valor_Total")
              'Actualizamos Saldos superiores
               If StockSuperior Then
                  Do While (CodigoAux <> "0")
                     Progreso_Barra.Mensaje_Box = "Existencia de " & CodigoB & " -> " & CodigoAux
                     Progreso_Esperar True

                     CodigoAux = CodigoCuentaSup(CodigoAux)
                     If CodigoAux <> "0" Then
                        sSQL = "UPDATE Catalogo_Productos " _
                             & "SET Valor_Total = Valor_Total + " & Total & " " _
                             & "WHERE Codigo_Inv = '" & CodigoAux & "' " _
                             & "AND Item = '" & NumEmpresa & "' " _
                             & "AND Periodo = '" & Periodo_Contable & "' "
                        Ejecutar_SQL_SP sSQL
                     End If
                  Loop
               End If
            End If
           .MoveNext
         Loop
     End If
    End With
    DGQuery.Visible = True
    Progreso_Final
    MsgBox Format(Time - MiTiempo, "HH:MM:SS")
End Sub

Public Sub Resumen_Lote()
  DGQuery.Visible = False
  Debitos = 0
  Creditos = 0
  Stock_Inv = 0
  sSQL = "SELECT TK.Codigo_Inv, CP.Producto, TK.CodBodega, TK.Lote_No, TK.Fecha_Fab, TK.Fecha_Exp, CP.Reg_Sanitario, " _
       & "TK.Modelo, TK.Procedencia, TK.Serie_No, SUM(TK.Entrada) As Entradas, SUM(TK.Salida) As Salidas, " _
       & "SUM(TK.Entrada-TK.Salida) As Stock_Lote,  AVG(Valor_Unitario) As Valor_Unit, " _
       & "(SUM(TK.Entrada-TK.Salida) * AVG(Valor_Unitario)) As Total_Inventario " _
       & "FROM Catalogo_Productos As CP, Trans_Kardex As TK " _
       & "WHERE CP.Item = '" & NumEmpresa & "' " _
       & "AND CP.Periodo = '" & Periodo_Contable & "' " _
       & "AND TK.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & SQL_Tipo_Busqueda _
       & "AND CP.Item = TK.Item " _
       & "AND CP.Periodo = TK.Periodo " _
       & "AND CP.Codigo_Inv = TK.Codigo_Inv " _
       & "GROUP BY TK.Codigo_Inv, CP.Producto, TK.CodBodega, TK.Lote_No, TK.Fecha_Fab, TK.Fecha_Exp, CP.Reg_Sanitario, " _
       & "TK.Modelo, TK.Procedencia, TK.Serie_No " _
       & "ORDER BY TK.Codigo_Inv, TK.Lote_No "
'''       & "UNION " _
'''       & "SELECT TK.Codigo As Codigo_Inv,CP.Producto,TK.CodBodega,0 As Entradas,SUM(TK.Cantidad) As Salidas " _
'''       & "FROM Catalogo_Productos As CP, Detalle_Factura As TK " _
'''       & "WHERE CP.Item = '" & NumEmpresa & "' " _
'''       & "AND CP.Periodo = '" & Periodo_Contable & "' " _
'''       & "AND CP.TC = 'P' " _
'''       & "AND TK.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
'''       & "AND TK.C = " & Val(adFalse) & " "
'''  If CheqBusqueda.value = 1 Then sSQL = sSQL & "AND CP.Producto LIKE '%" & TxtBusqueda & "%' "
'''  If CheqBod.value = 1 Then sSQL = sSQL & "AND TK.CodBodega = '" & SinEspaciosIzq(DCBodega.Text) & "' "
'''  If CheqLote.value = 1 Then sSQL = sSQL & "AND TK.Lote_No = '" & SinEspaciosIzq(DCLote.Text) & "' "
'''  sSQL = sSQL _
'''       & "AND CP.Item = TK.Item " _
'''       & "AND CP.Periodo = TK.Periodo " _
'''       & "AND CP.Codigo_Inv = TK.Codigo " _
'''       & "GROUP BY TK.Codigo, CP.Producto,TK.CodBodega " _
'''       & "ORDER BY TK.Codigo "
 'MsgBox sSQL
  Select_Adodc AdoDetKardex, sSQL
  DGQuery.Visible = False
  With AdoDetKardex.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Debitos = Debitos + .fields("Entradas")
          Creditos = Creditos + .fields("Salidas")
          Stock_Inv = Stock_Inv + .fields("Stock_Lote")
         .MoveNext
       Loop
   End If
  End With
  LabelStock.Caption = Format(Stock_Inv, "#,##0.00")
  DGQuery.Visible = True
End Sub

Public Sub Resumen_Barras()
  DGQuery.Visible = False
  Debitos = 0
  Creditos = 0
  Stock_Inv = 0
  sSQL = "SELECT TK.Codigo_Inv, CP.Producto, TK.CodBodega, TK.Codigo_Barra, CP.Reg_Sanitario, " _
       & "SUM(TK.Entrada) As Entradas, SUM(TK.Salida) As Salidas, " _
       & "SUM(TK.Entrada-TK.Salida) As Stock_Lote, AVG(TK.Valor_Unitario) As Valor_Unit, " _
       & "((SUM(TK.Entrada)-SUM(TK.Salida)) * AVG(TK.Valor_Unitario)) As Total_Inventario " _
       & "FROM Catalogo_Productos As CP, Trans_Kardex As TK " _
       & "WHERE CP.Item = '" & NumEmpresa & "' " _
       & "AND CP.Periodo = '" & Periodo_Contable & "' " _
       & "AND TK.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & SQL_Tipo_Busqueda _
       & "AND CP.Item = TK.Item " _
       & "AND CP.Periodo = TK.Periodo " _
       & "AND CP.Codigo_Inv = TK.Codigo_Inv " _
       & "GROUP BY TK.Codigo_Inv, CP.Producto, TK.CodBodega, TK.Codigo_Barra, CP.Reg_Sanitario " _
       & "HAVING SUM(TK.Entrada-TK.Salida) <> 0 " _
       & "ORDER BY TK.Codigo_Inv, TK.Codigo_Barra "
 'MsgBox sSQL
  SQLDec = "Valor_Unit " & CStr(Dec_Costo) & "|Total_Inventario 2|."
  Select_Adodc_Grid DGQuery, AdoDetKardex, sSQL, SQLDec
  DGQuery.Visible = False
  With AdoDetKardex.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Debitos = Debitos + .fields("Entradas")
          Creditos = Creditos + .fields("Salidas")
          Stock_Inv = Stock_Inv + .fields("Stock_Lote")
         .MoveNext
       Loop
   End If
  End With
  LabelStock.Caption = Format(Stock_Inv, "#,##0.00")
  DGQuery.Visible = True
End Sub

Public Sub Listar_Por_Tipo_Cta()
  If OpcInv.value Then
     sSQL = "SELECT CC.Cuenta,TK.Cta_Inv " _
          & "FROM Catalogo_Cuentas As CC, Trans_Kardex As TK " _
          & "WHERE CC.Item = '" & NumEmpresa & "' " _
          & "AND CC.Periodo = '" & Periodo_Contable & "' " _
          & "AND LEN(TK.Cta_Inv) > 1 " _
          & "AND CC.Codigo = TK.Cta_Inv " _
          & "AND CC.Item = TK.Item " _
          & "AND CC.Periodo = TK.Periodo " _
          & "GROUP BY CC.Cuenta,TK.Cta_Inv " _
          & "ORDER BY CC.Cuenta,TK.Cta_Inv "
  Else
     sSQL = "SELECT CC.Cuenta,TK.Contra_Cta " _
          & "FROM Catalogo_Cuentas As CC, Trans_Kardex As TK " _
          & "WHERE CC.Item = '" & NumEmpresa & "' " _
          & "AND CC.Periodo = '" & Periodo_Contable & "' " _
          & "AND LEN(TK.Contra_Cta) > 1 " _
          & "AND CC.Codigo = TK.Contra_Cta " _
          & "AND CC.Item = TK.Item " _
          & "AND CC.Periodo = TK.Periodo " _
          & "GROUP BY CC.Cuenta,TK.Contra_Cta " _
          & "ORDER BY CC.Cuenta,TK.Contra_Cta "
  End If
  SelectDB_Combo DCCtaInv, AdoCtaInv, sSQL, "Cuenta"
End Sub

Public Sub Listar_Por_Tipo_SubModulo()
  If OpcGasto.value Then
     sSQL = "SELECT TC, Codigo, Detalle As SubModulo " _
          & "FROM Catalogo_SubCtas " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND Detalle <> '" & Ninguno & "' " _
          & "ORDER BY TC,Detalle "
  Else
     sSQL = "SELECT CP.TC, CP.Codigo, CP.Cta, (C.Cliente + REPLICATE(' ', 60 - LEN(C.Cliente)) + CP.Cta) As SubModulo " _
          & "FROM Catalogo_CxCxP As CP, Clientes As C " _
          & "WHERE CP.Item = '" & NumEmpresa & "' " _
          & "AND CP.Periodo = '" & Periodo_Contable & "' " _
          & "AND C.Cliente <> '" & Ninguno & "' " _
          & "AND CP.TC = 'P' " _
          & "AND CP.Codigo = C.Codigo " _
          & "ORDER BY C.Cliente,CP.Cta "
  End If
  SelectDB_Combo DCSubModulo, AdoSubModulo, sSQL, "SubModulo"
End Sub

Public Sub Listar_Por_Producto()
  If OpcMarca.value Then
     sSQL = "SELECT CodMar As Codigo, Marca As Producto " _
          & "FROM Catalogo_Marcas " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND CodMar <> '" & Ninguno & "' " _
          & "ORDER BY Marca "
  ElseIf OpcBarra.value Then
         sSQL = "SELECT Codigo_Barra As Codigo, Codigo_Barra As Producto " _
              & "FROM Trans_Kardex " _
              & "WHERE Item = '" & NumEmpresa & "' " _
              & "AND Periodo = '" & Periodo_Contable & "' " _
              & "GROUP BY Codigo_Barra " _
              & "ORDER BY Codigo_Barra "
  ElseIf OpcLote.value Then
         sSQL = "SELECT Lote_No As Codigo, Lote_No As Producto " _
              & "FROM Trans_Kardex " _
              & "WHERE Item = '" & NumEmpresa & "' " _
              & "AND Periodo = '" & Periodo_Contable & "' " _
              & "GROUP BY Lote_No " _
              & "ORDER BY Lote_No "
  Else
     sSQL = "SELECT Codigo_Inv As Codigo, Producto " _
          & "FROM Catalogo_Productos " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND LEN(Cta_Inventario) > 2 " _
          & "AND Codigo_Inv LIKE '" & Buscar_Grupo_Inventario & "%' " _
          & "AND TC = 'P' " _
          & "ORDER BY Codigo_Inv "
  End If
  SelectDB_Combo DCTipoBusqueda, AdoBusqueda, sSQL, "Producto"
End Sub

Public Function Buscar_Grupo_Inventario() As String
Dim Result As String
Dim vProducto As String
  Result = Ninguno
  vProducto = DCTInv.Text
  If vProducto = "" Then vProducto = Ninguno
  With AdoTInv.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Producto = '" & vProducto & "' ")
       If Not .EOF Then Result = .fields("Codigo_Inv")
   End If
  End With
  Buscar_Grupo_Inventario = Result
End Function

Public Function SQL_Tipo_Busqueda() As String
Dim BSQL As String

  BSQL = " "
  CodigoInv = Ninguno
  If OpcProducto.value Then
     With AdoBusqueda.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
         .Find ("Producto = '" & DCTipoBusqueda & "' ")
          If Not .EOF Then CodigoInv = .fields("Codigo")
      End If
     End With
  Else
     CodigoInv = DCTipoBusqueda
  End If
  
  If CheqBod.value <> 0 Then BSQL = BSQL & "AND TK.CodBodega = '" & Cod_Bodega & "' "

  If CheqProducto.value <> 0 Then
     If OpcBarra.value Then
        BSQL = BSQL & "AND TK.Codigo_Barra = '" & CodigoInv & "' "
     ElseIf OpcLote.value Then
        BSQL = BSQL & "AND TK.Lote_No = '" & CodigoInv & "' "
     Else
        BSQL = BSQL & "AND TK.Codigo_Inv = '" & CodigoInv & "' "
     End If
  End If
  
  If CheqMonto.value <> 0 Then BSQL = BSQL & "AND CP.Stock_Actual = " & Val(TxtMonto) & " "
  
  If CheqExist.value = 0 Then
     BSQL = BSQL _
          & "AND CP.Valor_Total <> 0 "
  End If
 'MsgBox BSQL
  SQL_Tipo_Busqueda = BSQL
End Function

