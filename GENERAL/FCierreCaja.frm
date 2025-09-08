VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FCierreCaja 
   Caption         =   "CIERRE DE CAJA"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8505
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar TbarCierre 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   12
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir del Modulo"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Diario_Caja"
            Object.ToolTipText     =   "Diario de Caja"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Grabar"
            Object.ToolTipText     =   "Grabar Diario de Caja"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Cuadre_Caja"
            Object.ToolTipText     =   "Cuadre de Caja"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Diario"
            Object.ToolTipText     =   "Diario"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Asiento"
            Object.ToolTipText     =   "Asientos"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Reactivar"
            Object.ToolTipText     =   "Reactivar"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "SRI"
            Object.ToolTipText     =   "S.R.I."
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "IESS"
            Object.ToolTipText     =   "I.E.S.S."
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Anuladas"
            Object.ToolTipText     =   "Facturas Anuladas"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Comparar"
            Object.ToolTipText     =   "Comparar Cierre con Banco"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Excel"
            Object.ToolTipText     =   "Enviar a Excel los resultados"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
      EndProperty
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   645
         Left            =   7140
         TabIndex        =   5
         Top             =   0
         Width           =   10620
         Begin MSDataListLib.DataCombo DCBenef 
            Bindings        =   "FCierreCaja.frx":0000
            DataSource      =   "AdoClientes"
            Height          =   360
            Left            =   2625
            TabIndex        =   9
            Top             =   315
            Visible         =   0   'False
            Width           =   6315
            _ExtentX        =   11139
            _ExtentY        =   635
            _Version        =   393216
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.CheckBox CheqOrdDep 
            Caption         =   "Ordenar Por Depósito"
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
            Left            =   9030
            TabIndex        =   11
            Top             =   105
            Width           =   1380
         End
         Begin VB.CheckBox CheqCajero 
            Caption         =   "Por Cajero"
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
            Left            =   2625
            TabIndex        =   10
            Top             =   0
            Width           =   1275
         End
         Begin MSMask.MaskEdBox MBFechaI 
            Height          =   330
            Left            =   0
            TabIndex        =   6
            ToolTipText     =   "Formato de Fecha: DD/MM/AA"
            Top             =   315
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
         Begin MSMask.MaskEdBox MBFechaF 
            Height          =   330
            Left            =   1260
            TabIndex        =   8
            ToolTipText     =   "Formato de Fecha: DD/MM/AA"
            Top             =   315
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
            BorderStyle     =   1  'Fixed Single
            Caption         =   " &Periodo de Cierre"
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
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Width           =   2535
         End
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&S"
      Height          =   225
      Left            =   210
      TabIndex        =   12
      Top             =   12705
      Width           =   225
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   11880
      Left            =   105
      TabIndex        =   0
      Top             =   735
      Width           =   21675
      _ExtentX        =   38232
      _ExtentY        =   20955
      _Version        =   393216
      Tabs            =   6
      Tab             =   5
      TabsPerRow      =   6
      TabHeight       =   520
      BackColor       =   12632319
      TabCaption(0)   =   "&1.- VENTAS/ABONOS/ANTICIPOS"
      TabPicture(0)   =   "FCierreCaja.frx":001A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(1)=   "LabelAbonos"
      Tab(0).Control(2)=   "LabelCheque"
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(4)=   "AdoCxC"
      Tab(0).Control(5)=   "AdoVentas"
      Tab(0).Control(6)=   "DGVentas"
      Tab(0).Control(7)=   "DGCxC"
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "&2.- INVENTARIO"
      TabPicture(1)   =   "FCierreCaja.frx":0036
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "AdoInv"
      Tab(1).Control(1)=   "DGInv"
      Tab(1).Control(2)=   "DGProductos"
      Tab(1).Control(3)=   "DGCierres"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "&3.- CONTABILIDAD"
      TabPicture(2)   =   "FCierreCaja.frx":0052
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DGAsiento1"
      Tab(2).Control(1)=   "DGAsiento"
      Tab(2).Control(2)=   "DGAnticipos"
      Tab(2).Control(3)=   "LabelHaber1"
      Tab(2).Control(4)=   "LabelDebe1"
      Tab(2).Control(5)=   "Label15"
      Tab(2).Control(6)=   "LabelHaber"
      Tab(2).Control(7)=   "LabelDebe"
      Tab(2).Control(8)=   "Label11"
      Tab(2).Control(9)=   "Label1"
      Tab(2).Control(10)=   "LblDiferencia"
      Tab(2).Control(11)=   "LblConcepto"
      Tab(2).Control(12)=   "LblConcepto1"
      Tab(2).Control(13)=   "LblDiferencia1"
      Tab(2).Control(14)=   "Label13"
      Tab(2).ControlCount=   15
      TabCaption(3)   =   "&4.- ANULADAS"
      TabPicture(3)   =   "FCierreCaja.frx":006E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "DGFactAnul"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "&5.- REPORTE DE AUDITORIA"
      TabPicture(4)   =   "FCierreCaja.frx":008A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label9"
      Tab(4).Control(1)=   "Label12"
      Tab(4).Control(2)=   "Label14"
      Tab(4).Control(3)=   "Label18"
      Tab(4).Control(4)=   "LblConIVA"
      Tab(4).Control(5)=   "LblSinIVA"
      Tab(4).Control(6)=   "LblDescuento"
      Tab(4).Control(7)=   "LblIVA"
      Tab(4).Control(8)=   "Label7"
      Tab(4).Control(9)=   "LblServicio"
      Tab(4).Control(10)=   "Label16"
      Tab(4).Control(11)=   "LblTotalFacturado"
      Tab(4).Control(12)=   "AdoSRI"
      Tab(4).Control(13)=   "DGSRI"
      Tab(4).ControlCount=   14
      TabCaption(5)   =   "&6.- REPORTE DEL BANCO"
      TabPicture(5)   =   "FCierreCaja.frx":00A6
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).Control(0)=   "DCBanco"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "DGBanco"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).ControlCount=   2
      Begin MSDataGridLib.DataGrid DGCxC 
         Bindings        =   "FCierreCaja.frx":00C2
         Height          =   1800
         Left            =   -74895
         TabIndex        =   15
         ToolTipText     =   "<Ctrl+P> Protestar Cheques"
         Top             =   5880
         Width           =   14190
         _ExtentX        =   25030
         _ExtentY        =   3175
         _Version        =   393216
         AllowUpdate     =   0   'False
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
      Begin MSDataGridLib.DataGrid DGVentas 
         Bindings        =   "FCierreCaja.frx":00D7
         Height          =   3090
         Left            =   -74895
         TabIndex        =   1
         Top             =   420
         Width           =   14160
         _ExtentX        =   24977
         _ExtentY        =   5450
         _Version        =   393216
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
      Begin MSDataGridLib.DataGrid DGCierres 
         Bindings        =   "FCierreCaja.frx":00EF
         Height          =   4320
         Left            =   -74895
         TabIndex        =   18
         Top             =   420
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   7620
         _Version        =   393216
         BackColor       =   12648447
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
      Begin MSDataGridLib.DataGrid DGProductos 
         Bindings        =   "FCierreCaja.frx":0108
         Height          =   2115
         Left            =   -72795
         TabIndex        =   17
         Top             =   3675
         Width           =   9150
         _ExtentX        =   16140
         _ExtentY        =   3731
         _Version        =   393216
         BackColor       =   12648447
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
      Begin MSDataGridLib.DataGrid DGInv 
         Bindings        =   "FCierreCaja.frx":0123
         Height          =   2325
         Left            =   -72795
         TabIndex        =   16
         Top             =   420
         Width           =   9150
         _ExtentX        =   16140
         _ExtentY        =   4101
         _Version        =   393216
         BackColor       =   12648447
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
      Begin MSDataGridLib.DataGrid DGAsiento1 
         Bindings        =   "FCierreCaja.frx":0138
         Height          =   1695
         Left            =   -74895
         TabIndex        =   20
         Top             =   3975
         Width           =   11040
         _ExtentX        =   19473
         _ExtentY        =   2990
         _Version        =   393216
         BackColor       =   12640511
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
      Begin MSDataGridLib.DataGrid DGAsiento 
         Bindings        =   "FCierreCaja.frx":0152
         Height          =   1380
         Left            =   -74895
         TabIndex        =   19
         Top             =   840
         Width           =   11040
         _ExtentX        =   19473
         _ExtentY        =   2434
         _Version        =   393216
         BackColor       =   12640511
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
      Begin MSDataGridLib.DataGrid DGFactAnul 
         Bindings        =   "FCierreCaja.frx":016B
         Height          =   4110
         Left            =   -74895
         TabIndex        =   33
         Top             =   420
         Width           =   11040
         _ExtentX        =   19473
         _ExtentY        =   7250
         _Version        =   393216
         BackColor       =   12632319
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
      Begin MSDataGridLib.DataGrid DGSRI 
         Bindings        =   "FCierreCaja.frx":0185
         Height          =   3375
         Left            =   -74895
         TabIndex        =   34
         Top             =   420
         Width           =   11040
         _ExtentX        =   19473
         _ExtentY        =   5953
         _Version        =   393216
         BackColor       =   16777152
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
      Begin MSDataGridLib.DataGrid DGBanco 
         Bindings        =   "FCierreCaja.frx":019A
         Height          =   4950
         Left            =   105
         TabIndex        =   47
         Top             =   840
         Width           =   19230
         _ExtentX        =   33920
         _ExtentY        =   8731
         _Version        =   393216
         BackColor       =   16761024
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
      Begin MSAdodcLib.Adodc AdoVentas 
         Height          =   330
         Left            =   -74895
         Top             =   3990
         Width           =   3375
         _ExtentX        =   5953
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
         Caption         =   "Ventas"
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
      Begin MSAdodcLib.Adodc AdoCxC 
         Height          =   330
         Left            =   -74895
         Top             =   7770
         Width           =   3480
         _ExtentX        =   6138
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
         Caption         =   "CxC"
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
      Begin MSAdodcLib.Adodc AdoSRI 
         Height          =   330
         Left            =   -74895
         Top             =   5250
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
         Caption         =   "SRI"
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
      Begin MSDataListLib.DataCombo DCBanco 
         Bindings        =   "FCierreCaja.frx":01B1
         DataSource      =   "AdoCtaBanco"
         Height          =   315
         Left            =   105
         TabIndex        =   48
         Top             =   420
         Width           =   6525
         _ExtentX        =   11509
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16744576
         ForeColor       =   192
         Text            =   "Banco"
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
      Begin MSDataGridLib.DataGrid DGAnticipos 
         Bindings        =   "FCierreCaja.frx":01CB
         Height          =   1800
         Left            =   -74895
         TabIndex        =   49
         Top             =   7350
         Width           =   14190
         _ExtentX        =   25030
         _ExtentY        =   3175
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   12648447
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
      Begin MSAdodcLib.Adodc AdoInv 
         Height          =   330
         Left            =   -72795
         Top             =   3045
         Width           =   3900
         _ExtentX        =   6879
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
         Caption         =   "Inv"
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
      Begin VB.Label LabelHaber1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Left            =   -62085
         TabIndex        =   25
         Top             =   5760
         Width           =   1800
      End
      Begin VB.Label LabelDebe1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Left            =   -63870
         TabIndex        =   24
         Top             =   5760
         Width           =   1800
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTALES "
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
         Left            =   -64920
         TabIndex        =   21
         Top             =   5760
         Width           =   1065
      End
      Begin VB.Label LabelHaber 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Left            =   -62085
         TabIndex        =   28
         Top             =   3045
         Width           =   1800
      End
      Begin VB.Label LabelDebe 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Left            =   -63870
         TabIndex        =   29
         Top             =   3045
         Width           =   1800
      End
      Begin VB.Label LblTotalFacturado 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Left            =   -63240
         TabIndex        =   46
         Top             =   5550
         Width           =   1800
      End
      Begin VB.Label Label16 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " T O T A L"
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
         Left            =   -63240
         TabIndex        =   45
         Top             =   5235
         Width           =   1800
      End
      Begin VB.Label LblServicio 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Left            =   -65025
         TabIndex        =   44
         Top             =   5550
         Width           =   1800
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TOTAL  SERVICIO"
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
         Left            =   -65025
         TabIndex        =   43
         Top             =   5235
         Width           =   1800
      End
      Begin VB.Label LblIVA 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Left            =   -66810
         TabIndex        =   42
         Top             =   5550
         Width           =   1800
      End
      Begin VB.Label LblDescuento 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Left            =   -68595
         TabIndex        =   41
         Top             =   5550
         Width           =   1800
      End
      Begin VB.Label LblSinIVA 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Left            =   -70380
         TabIndex        =   40
         Top             =   5550
         Width           =   1800
      End
      Begin VB.Label LblConIVA 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Left            =   -72165
         TabIndex        =   39
         Top             =   5550
         Width           =   1800
      End
      Begin VB.Label Label18 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TOTAL  I.V.A."
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
         Left            =   -66810
         TabIndex        =   38
         Top             =   5235
         Width           =   1800
      End
      Begin VB.Label Label14 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " DESCUENTO"
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
         Left            =   -68595
         TabIndex        =   37
         Top             =   5235
         Width           =   1800
      End
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " SIN I.V.A."
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
         Left            =   -70380
         TabIndex        =   36
         Top             =   5235
         Width           =   1800
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CON I.V.A."
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
         Left            =   -72165
         TabIndex        =   35
         Top             =   5235
         Width           =   1800
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTALES "
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
         Left            =   -64920
         TabIndex        =   32
         Top             =   3030
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Diferencia "
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
         Left            =   -67755
         TabIndex        =   31
         Top             =   3030
         Width           =   1065
      End
      Begin VB.Label LblDiferencia 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Left            =   -66705
         TabIndex        =   30
         Top             =   3030
         Width           =   1800
      End
      Begin VB.Label LblConcepto 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "."
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
         Left            =   -74895
         TabIndex        =   27
         Top             =   420
         Width           =   11040
      End
      Begin VB.Label LblConcepto1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "."
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
         Left            =   -74895
         TabIndex        =   26
         Top             =   3465
         Width           =   11040
      End
      Begin VB.Label LblDiferencia1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Left            =   -66705
         TabIndex        =   23
         Top             =   5760
         Width           =   1800
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Diferencia "
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
         Left            =   -67755
         TabIndex        =   22
         Top             =   5760
         Width           =   1065
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TOTAL ABONOS"
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
         Left            =   -71325
         TabIndex        =   14
         Top             =   7770
         Width           =   1590
      End
      Begin VB.Label LabelCheque 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Left            =   -69750
         TabIndex        =   13
         Top             =   7770
         Width           =   1695
      End
      Begin VB.Label LabelAbonos 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Left            =   -69855
         TabIndex        =   3
         Top             =   3990
         Width           =   1695
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TOTAL VENTAS"
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
         Left            =   -71430
         TabIndex        =   2
         Top             =   3990
         Width           =   1590
      End
   End
   Begin MSAdodcLib.Adodc AdoAsiento 
      Height          =   330
      Left            =   210
      Top             =   5565
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
      Caption         =   "Asiento"
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
   Begin MSAdodcLib.Adodc AdoSQL 
      Height          =   330
      Left            =   210
      Top             =   4305
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
      Caption         =   "SQL"
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
      Left            =   210
      Top             =   3990
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
   Begin MSAdodcLib.Adodc AdoClientes 
      Height          =   330
      Left            =   210
      Top             =   3045
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
      Caption         =   "Clientes"
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
   Begin MSAdodcLib.Adodc AdoVentaAct 
      Height          =   330
      Left            =   210
      Top             =   4620
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
      Caption         =   "VentaAct"
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
   Begin MSAdodcLib.Adodc AdoInv1 
      Height          =   330
      Left            =   210
      Top             =   4935
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
      Caption         =   "Inv1"
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
   Begin MSAdodcLib.Adodc AdoFactAnul 
      Height          =   330
      Left            =   210
      Top             =   5250
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
      Caption         =   "FactAnul"
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
   Begin MSAdodcLib.Adodc AdoProductos 
      Height          =   330
      Left            =   210
      Top             =   3360
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
      Caption         =   "Productos"
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
   Begin MSAdodcLib.Adodc AdoCxC1 
      Height          =   330
      Left            =   210
      Top             =   6195
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
      Caption         =   "CxC1"
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
   Begin MSAdodcLib.Adodc AdoAsiento1 
      Height          =   330
      Left            =   210
      Top             =   5880
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
      Caption         =   "Asiento1"
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
   Begin MSAdodcLib.Adodc AdoCierres 
      Height          =   330
      Left            =   210
      Top             =   6510
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
      Caption         =   "Cierres"
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
   Begin MSComDlg.CommonDialog CDialogDir 
      Left            =   3255
      Top             =   3150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSAdodcLib.Adodc AdoBanco 
      Height          =   330
      Left            =   210
      Top             =   6825
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
      Caption         =   "Banco"
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
   Begin MSAdodcLib.Adodc AdoCtaBanco 
      Height          =   330
      Left            =   210
      Top             =   7140
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
      Caption         =   "CtaBanco"
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
   Begin MSAdodcLib.Adodc AdoAnticipos 
      Height          =   330
      Left            =   210
      Top             =   7455
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
      Caption         =   "Anticipos"
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
   Begin ComctlLib.ImageList ImageList1 
      Left            =   525
      Top             =   8190
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   14
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FCierreCaja.frx":01E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FCierreCaja.frx":0500
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FCierreCaja.frx":081A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FCierreCaja.frx":0B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FCierreCaja.frx":0E4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FCierreCaja.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FCierreCaja.frx":1482
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FCierreCaja.frx":1774
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FCierreCaja.frx":1A8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FCierreCaja.frx":22A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FCierreCaja.frx":25C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FCierreCaja.frx":28DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FCierreCaja.frx":2BF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FCierreCaja.frx":3848
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FCierreCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FormaCierre As Boolean
Dim Por_Combos As Boolean
Dim ContSC As Integer
Dim ContCtas As Integer
Dim NumTrans As Long
Dim CtasProc() As CtasAsiento
Dim ErrorInventario As String
Dim ErrorFacturas As String
Dim CtaDeAnticipos As String
Dim Combos As String
Dim FechaDelCierre As String

'''Private Sub Leer_Excel_AdoDB(sPath As String)
'''On Error GoTo error_sub
''''Variables para acceder a la hoja excel
'''Dim obj_Excel As Object
'''Dim Obj_Hoja As Object
'''
'''Dim I As Long
'''Dim N As Long
'''Dim No_Bancos As Long
'''
'''Dim TextoCel As String
'''Dim ContItem As Integer
'''Dim DepItem As String
'''
'''    If Len(Dir(sPath)) = 0 Then
'''       MsgBox "El archivo no existe", vbCritical
'''       Exit Sub
'''    End If
'''    RatonReloj
'''    No_Bancos = 0
'''
'''   'Crear la instancia de la aplicación Excel
'''    Set obj_Excel = CreateObject("Excel.Application")
'''    'obj_Excel.Visible = True
'''    FlexGrid.Visible = False
'''   'MsgBox path_excel
'''    obj_Excel.Workbooks.open Filename:=path_excel
'''    If Val(obj_Excel.Application.version) >= 8 Then Set Obj_Hoja = obj_Excel.ActiveSheet Else Set Obj_Hoja = obj_Excel
'''   'Almacenamos en la variable Type el rango, es decir la primera fila la ultima fila, La primer Columna y la ultima
'''    With Rango
'''        .NumFila1 = Format$(Obj_Hoja.UsedRange.Row)
'''        .NumFila2 = Format$(Obj_Hoja.UsedRange.Row + Obj_Hoja.UsedRange.rows.Count - 1)
'''        .NumCol1 = Format$(Obj_Hoja.UsedRange.Column)
'''        .NumCol2 = Format$(Obj_Hoja.UsedRange.Column + Obj_Hoja.UsedRange.Columns.Count - 1)
'''        'MsgBox obj_Hoja.UsedRange.Row & vbCrLf & obj_Hoja.UsedRange.Rows.Count
'''         Progreso_Iniciar
'''         Progreso_Barra.Incremento = 0
'''         Progreso_Barra.Valor_Maximo = .NumFila2
'''    End With
'''
'''    FechaIni = BuscarFecha(MBFechaI)
'''    FechaFin = BuscarFecha(MBFechaF)
'''    FA.Fecha_Desde = MBFechaI
'''    FA.Fecha_Hasta = MBFechaF
'''
'''    Trans_No = 255
'''    sSQL = "DELETE * " _
'''         & "FROM Asiento_SC " _
'''         & "WHERE Item = '" & NumEmpresa & "' " _
'''         & "AND T_No = " & Trans_No & " " _
'''         & "AND CodigoU = '" & CodigoUsuario & "' "
'''    Ejecutar_SQL_SP sSQL
'''
'''   'MsgBox Rango.NumCol2 & vbCrLf & Rango.NumFila2
'''   'Recorremos las filas del FlexGrid para agregar los datos
'''    For I = 1 To Rango.NumFila2 - 1
'''        Progreso_Barra.Mensaje_Box = "(" & I & "/" & Rango.NumFila2 & ") Importando del Excel al Sistema "
'''        Progreso_Esperar
'''       'Recorremos las columnas y Fila del FlexGrid
'''        CodigoCli = obj_Worksheet.cells(I + 1, 3).value
'''        Abono = Val(obj_Worksheet.cells(I + 1, 5).value)
'''        Mifecha = obj_Worksheet.cells(I + 1, 6).value
'''        Beneficiario = obj_Worksheet.cells(I + 1, 4).value
'''        If Len(CodigoCli) > 1 Then
'''           If Len(CodigoCli) < 10 Then CodigoCli = "0" & CodigoCli
'''           SetAdoAddNew "Asiento_SC"
'''           SetAdoFields "FECHA_V", Mifecha
'''           SetAdoFields "Cta", CodigoCli
'''           SetAdoFields "Beneficiario", Beneficiario
'''           SetAdoFields "Valor", Abono
'''           SetAdoFields "Item", NumEmpresa
'''           SetAdoFields "T_No", Trans_No
'''           SetAdoFields "CodigoU", CodigoUsuario
'''           SetAdoUpdate
'''        End If
'''    Next I
'''    obj_Workbook.Close
'''    obj_Excel.Quit
'''    Descargar
'''
'''    Progreso_Barra.Mensaje_Box = "Actualizando Abonos"
'''    Progreso_Esperar
'''    Cta = SinEspaciosIzq(DCBanco.Text)
'''
'''    Progreso_Barra.Mensaje_Box = "Actualizando Codigos "
'''    Progreso_Esperar
'''    sSQL = "UPDATE Asiento_SC " _
'''         & "SET Codigo=C.Codigo " _
'''         & "FROM Asiento_SC As SC, Clientes AS C " _
'''         & "WHERE SC.Item = '" & NumEmpresa & "' " _
'''         & "AND SC.T_No = " & Trans_No & " " _
'''         & "AND SC.CodigoU = '" & CodigoUsuario & "' " _
'''         & "AND SC.Cta = C.CI_RUC "
'''    Ejecutar_SQL_SP sSQL
'''
'''    Progreso_Barra.Mensaje_Box = "Actualizando Abonos"
'''    Progreso_Esperar
'''    Cta = SinEspaciosIzq(DCBanco.Text)
'''    sSQL = "UPDATE Asiento_SC " _
'''         & "SET TC='Ok' " _
'''         & "FROM Asiento_SC As SC, Trans_Abonos As TA " _
'''         & "WHERE SC.Item = '" & NumEmpresa & "' " _
'''         & "AND SC.T_No = " & Trans_No & " " _
'''         & "AND SC.CodigoU = '" & CodigoUsuario & "' " _
'''         & "AND TA.Periodo = '" & Periodo_Contable & "' " _
'''         & "AND TA.Cta = '" & Cta & "' " _
'''         & "AND SC.Item = TA.Item " _
'''         & "AND SC.Valor = TA.Abono " _
'''         & "AND SC.Codigo = TA.CodigoC " _
'''         & "AND SC.FECHA_V = TA.Fecha "
'''    Ejecutar_SQL_SP sSQL
'''
'''    sSQL = "UPDATE Trans_Abonos " _
'''         & "SET X = '.' " _
'''         & "WHERE Item = '" & NumEmpresa & "' " _
'''         & "AND Periodo = '" & Periodo_Contable & "' " _
'''         & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
'''         & "AND Cta = '" & Cta & "' "
'''    Ejecutar_SQL_SP sSQL
'''
'''    sSQL = "UPDATE Trans_Abonos " _
'''         & "SET X = 'O' " _
'''         & "FROM Trans_Abonos As TA, Asiento_SC As SC " _
'''         & "WHERE TA.Item = '" & NumEmpresa & "' " _
'''         & "AND TA.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
'''         & "AND SC.T_No = " & Trans_No & " " _
'''         & "AND SC.CodigoU = '" & CodigoUsuario & "' " _
'''         & "AND TA.Periodo = '" & Periodo_Contable & "' " _
'''         & "AND TA.Cta = '" & Cta & "' " _
'''         & "AND SC.Item = TA.Item " _
'''         & "AND SC.Valor = TA.Abono " _
'''         & "AND SC.Codigo = TA.CodigoC " _
'''         & "AND SC.FECHA_V = TA.Fecha "
'''    Ejecutar_SQL_SP sSQL
'''    sSQL = "SELECT TA.*,C.Cliente " _
'''         & "FROM Trans_Abonos As TA, Clientes As C " _
'''         & "WHERE TA.Item = '" & NumEmpresa & "' " _
'''         & "AND TA.Periodo = '" & Periodo_Contable & "' " _
'''         & "AND TA.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
'''         & "AND TA.Cta = '" & Cta & "' " _
'''         & "AND TA.X = '.' " _
'''         & "AND TA.CodigoC = C.Codigo " _
'''         & "ORDER BY TA.Fecha,C.Cliente "
'''    Select_Adodc AdoAux, sSQL
'''    With AdoAux.Recordset
'''     If .RecordCount > 0 Then
'''         Do While Not .EOF
'''            SetAdoAddNew "Asiento_SC"
'''            SetAdoFields "FECHA_V", .fields("Fecha")
'''            SetAdoFields "Cta", Cta
'''            SetAdoFields "Codigo", .fields("CodigoC")
'''            SetAdoFields "Beneficiario", "- " & .fields("Cliente")
'''            SetAdoFields "Valor", .fields("Abono")
'''            SetAdoFields "Item", NumEmpresa
'''            SetAdoFields "T_No", Trans_No
'''            SetAdoFields "CodigoU", CodigoUsuario
'''            SetAdoUpdate
'''           .MoveNext
'''         Loop
'''     End If
'''    End With
'''
'''    Progreso_Barra.Mensaje_Box = "Determinando Abonos descuadrados"
'''    Progreso_Esperar
'''    sSQL = "DELETE * " _
'''         & "FROM Asiento_SC " _
'''         & "WHERE Item = '" & NumEmpresa & "' " _
'''         & "AND T_No = " & Trans_No & " " _
'''         & "AND CodigoU = '" & CodigoUsuario & "' " _
'''         & "AND TC = 'Ok' "
'''    Ejecutar_SQL_SP sSQL
'''
'''    sSQL = "SELECT FECHA_V As Fecha, Cta As Codigo, Beneficiario, Valor As Deposito " _
'''         & "FROM Asiento_SC " _
'''         & "WHERE Item = '" & NumEmpresa & "' " _
'''         & "AND T_No = " & Trans_No & " " _
'''         & "AND CodigoU = '" & CodigoUsuario & "' " _
'''         & "ORDER BY Beneficiario, Fecha_V "
'''    Select_Adodc_Grid DGBanco, AdoBanco, sSQL
'''    RatonNormal
'''    MsgBox "Revise en la pestaña " & vbCrLf _
'''           & """REPORTE DEL BANCO""" & vbCrLf _
'''           & "los depositos descuadrados "
'''    Progreso_Final
'''Exit Sub
'''error_sub:
'''MsgBox Err.Description
'''Descargar
'''Progreso_Final
'''End Sub

'''Private Sub Descargar()
'''    On Local Error Resume Next
'''    Set obj_Workbook = Nothing
'''    Set obj_Excel = Nothing
'''    Set obj_Worksheet = Nothing
'''    Me.MousePointer = vbDefault
'''End Sub

Private Sub CheqCajero_Click()
  If CheqCajero.value = 1 Then DCBenef.Visible = True Else DCBenef.Visible = False
End Sub

'Cierre diario de Caja y asientos contables
Private Sub Procesar_Diario_Caja()
Dim MesA As Integer
Dim FechaA As Long

  RatonReloj
  Progreso_Barra.Mensaje_Box = "Procesando el Cierre de Caja..."
  Progreso_Iniciar
  Progreso_Iniciar_Errores
  Progreso_Barra.Incremento = 0
  Progreso_Barra.Valor_Maximo = 100
  If Inv_Promedio Then FCierreCaja.Caption = "CIERRE DE CAJA INVENTARIO PRECIO PROMEDIO" Else FCierreCaja.Caption = "CIERRE DE CAJA INVENTARIO ULTIMO PRECIO"
  ErrorFacturas = ""
  ErrorInventario = ""
  Control_Procesos "F", "Cierre Diarios de Caja"
  Presentar_Inventario = False
  FechaIni = BuscarFecha(MBFechaI)
  FechaFin = BuscarFecha(MBFechaF)
  FA.Fecha_Corte = FechaSistema
  FA.Fecha_Desde = MBFechaI
  FA.Fecha_Hasta = MBFechaF

 '---------------------------------------------------------------------------------------
 'Enceramos para realizar la primer parte del cierre de Abonos, NC, Cruce de Cuentas, etc
 '---------------------------------------------------------------------------------------
  Progreso_Barra.Mensaje_Box = "Enceramos los asientos temporales"
  Progreso_Esperar
  ContSC = 0
  Trans_No = 97
  IniciarAsientosDe DGAsiento1, AdoAsiento1     ' Abonos
  Trans_No = 96
  IniciarAsientosDe DGAsiento, AdoAsiento       ' CxC
  Progreso_Barra.Mensaje_Box = "Actualizando Productos"
  Progreso_Esperar
  Insertar_Productos_Cierre_Caja_SP MBFechaI, MBFechaF
  Progreso_Barra.Mensaje_Box = "Mayorizando Inventarios"
  Progreso_Esperar
  Mayorizar_Inventario_SP
  Progreso_Barra.Mensaje_Box = "Actualizando Abonos"
  Progreso_Esperar
  Actualizar_Abonos_Facturas_SP FA, True, True
  Progreso_Barra.Mensaje_Box = "Actualizando Clientes"
  Progreso_Esperar
  Actualizar_Datos_Representantes_SP Mas_Grupos
 'PROCESAR ASIENTOS DE FACTURACION
 '---------------------------------
  Progreso_Barra.Mensaje_Box = "Procesando Asientos Contables..."
  Progreso_Esperar
  Grabar_Asientos_Facturacion Normal

  Progreso_Barra.Mensaje_Box = "Verificando Errores"
  Progreso_Esperar
  Presenta_Errores_Facturacion_SP MBFechaI, MBFechaF
    
  RatonNormal

''     Command1.SetFocus
''  Else
''     If Command2.Enabled Then Command2.SetFocus Else Command5.SetFocus
''  End If
   Progreso_Barra.Mensaje_Box = "OK"
   Progreso_Final

  FInfoError.Show
End Sub

'Grabacion de los comprobantes contables
Private Sub Grabar_Cierre_Diario()
   NuevoComp = True
   ModificarComp = False
   CopiarComp = False
   
   FechaValida MBFechaI
   FechaValida MBFechaF
   FechaTexto = MBFechaF.Text
   FechaComp = FechaTexto
   Nombre_Cajero = Ninguno
   If CheqCajero.value = 1 Then
      Nombre_Cajero = MidStrg(DCBenef.Text, 1, Len(DCBenef.Text) - Len(SinEspaciosDer(DCBenef.Text)) - 1)
   End If
  
  'Verificamos partida doble de los dos asientos
   Debe = 0: Haber = 0
   With AdoAsiento.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
        Do While Not .EOF
           Debe = Debe + .fields("DEBE")
           Haber = Haber + .fields("HABER")
          .MoveNext
        Loop
       .MoveFirst
    End If
   End With
   With AdoAsiento1.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
        Do While Not .EOF
           Debe = Debe + .fields("DEBE")
           Haber = Haber + .fields("HABER")
          .MoveNext
        Loop
       .MoveFirst
    End If
   End With
   LabelDebe.Caption = Format$(Debe, "#,##0.00")
   LabelHaber.Caption = Format$(Haber, "#,##0.00")
   LblDiferencia.Caption = Format$(Debe - Haber, "#,##0.00")
   If ((NuevoDiario) And Redondear(Debe - Haber, 2) = 0) Then
       FechaTexto = MBFechaF
       FechaComp = FechaTexto
       NumComp = ReadSetDataNum("Diario", True, False)
       Mensajes = "Esta seguro de Grabar el Cierre de Caja"
       Titulo = "Pregunta de grabación"
       If BoxMensaje = vbYes Then
          RatonReloj
          DiarioCaja = NumComp
          FechaTexto = MBFechaF
          FechaIni = BuscarFecha(MBFechaI)
          FechaFin = BuscarFecha(MBFechaF)

          If FormaCierre Then
             Imprimir_Diario_Caja AdoVentas, AdoCxC, AdoInv, AdoProductos, AdoAnticipos, MBFechaI, MBFechaF
          Else
             Imprimir_Diario_Caja_Resumen AdoVentas, AdoCxC, AdoInv, AdoProductos, AdoAnticipos, MBFechaI, MBFechaF
          End If
         'Grabacion del Comprobante de Abonos
          If AdoAsiento1.Recordset.RecordCount > 0 Then
             Trans_No = 97
             NumComp = ReadSetDataNum("Diario", True, True)
             Co.T = Normal
             Co.TP = CompDiario
             Co.Fecha = FechaTexto
             Co.Numero = NumComp
             Co.CodigoB = Ninguno
             Co.Beneficiario = Ninguno
             Co.Efectivo = 0
             Co.Monto_Total = Debe
             Co.T_No = Trans_No
             Co.Usuario = CodigoUsuario
             Co.Item = NumEmpresa
             Co.Concepto = "Cierre de Caja de Abonos del " & FechaDelCierre
             Grabar_Comprobante Co
            'Los Asientos de SubModulos
             sSQL = "UPDATE Trans_SubCtas " _
                  & "SET TP = '" & Co.TP & "', Numero = " & Co.Numero & " " _
                  & "WHERE Item = '" & NumEmpresa & "' " _
                  & "AND Periodo = '" & Periodo_Contable & "' " _
                  & "AND TP = '.' " _
                  & "AND Numero = 0 " _
                  & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
             Ejecutar_SQL_SP sSQL
            'Abonos NC
             sSQL = "UPDATE Trans_Kardex " _
                  & "SET TP = '" & Co.TP & "', Numero = " & Co.Numero & " " _
                  & "WHERE Item = '" & NumEmpresa & "' " _
                  & "AND Periodo = '" & Periodo_Contable & "' " _
                  & "AND LEN(TC) = 2 " _
                  & "AND LEN(Serie) = 6 " _
                  & "AND Factura <> 0 " _
                  & "AND Entrada <> 0 " _
                  & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
                  & "AND SUBSTRING(Detalle,1,3) ='NC:' "
             Ejecutar_SQL_SP sSQL
             
             Control_Procesos Normal, Co.Concepto
             ImprimirComprobantesDe False, Co
             IniciarAsientosDe DGAsiento1, AdoAsiento1
          End If
         'Grabacion del Comprobante de CxC
          If AdoAsiento.Recordset.RecordCount > 0 Then
             Trans_No = 96
             NumComp = ReadSetDataNum("Diario", True, True)
             Co.T = Normal
             Co.TP = CompDiario
             Co.Fecha = FechaTexto
             Co.Numero = NumComp
             Co.CodigoB = Ninguno
             Co.Efectivo = 0
             Co.Monto_Total = Debe
             Co.T_No = Trans_No
             Co.Usuario = CodigoUsuario
             Co.Item = NumEmpresa
             Co.Concepto = "Cierre de Caja de Cuentas por Cobrar del " & FechaDelCierre
             Grabar_Comprobante Co
            'CxC
             sSQL = "UPDATE Trans_Kardex " _
                  & "SET TP = '" & Co.TP & "', Numero = " & Co.Numero & " " _
                  & "WHERE Item = '" & NumEmpresa & "' " _
                  & "AND Periodo = '" & Periodo_Contable & "' " _
                  & "AND TC IN ('FA','NV','DO','NDU','NDO') " _
                  & "AND LEN(Serie) = 6 " _
                  & "AND Factura <> 0 " _
                  & "AND Salida <> 0 " _
                  & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
             Ejecutar_SQL_SP sSQL
             FechaFin = BuscarFecha(FechaSistema)
             Parametros = "'" & NumEmpresa & "','" & Periodo_Contable & "','" & FechaIni & "','" & FechaFin & "' "
             Ejecutar_SP "sp_Productos_Cierre_Caja", Parametros
             
             Control_Procesos Normal, Co.Concepto
             ImprimirComprobantesDe False, Co
             IniciarAsientosDe DGAsiento, AdoAsiento
          End If
          LabelDebe.Caption = Format$(0, "#,##0.00")
          LabelHaber.Caption = Format$(0, "#,##0.00")
          RatonNormal
          Mifecha = BuscarFecha(FechaSistema)
          sSQL = "UPDATE Trans_Abonos " _
               & "SET C = " & Val(adTrue) & " " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
          Ejecutar_SQL_SP sSQL
          
          sSQL = "UPDATE Facturas " _
               & "SET C = " & Val(adTrue) & " " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
          Ejecutar_SQL_SP sSQL
          CierreDelDia
        End If
   Else
       RatonNormal
       MsgBox "Ya esta cerrado este día o no hay datos que procesar"
   End If
End Sub

Private Sub IESS_Cierre_Diario()
Dim NumFile As Integer
Dim RutaGeneraFile As String
Dim CI_RUCC As String
Dim NombreC As String

   RatonReloj
   sSQL = "DELETE * " _
        & "FROM Asiento_Beneficiarios " _
        & "WHERE Codigo <> '-' "
   Ejecutar_SQL_SP sSQL
    
   sSQL = "UPDATE Clientes " _
        & "SET X = '.' " _
        & "WHERE Codigo <> '-' "
   Ejecutar_SQL_SP sSQL
   
   FechaIni = BuscarFecha(MBFechaI.Text)
   FechaFin = BuscarFecha(MBFechaF.Text)
   
   sSQL = "UPDATE Clientes " _
        & "SET X = 'I' " _
        & "FROM Clientes As C, Detalle_Factura As DF " _
        & "WHERE DF.Item = '" & NumEmpresa & "' " _
        & "AND DF.Periodo = '" & Periodo_Contable & "' " _
        & "AND DF.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "AND C.Codigo = DF.CodigoB "
   Ejecutar_SQL_SP sSQL
    
   sSQL = "INSERT INTO Asiento_Beneficiarios (Codigo, Beneficiario, TD, RUC_CI) " _
        & "SELECT Codigo, Cliente, TD, CI_RUC " _
        & "FROM Clientes " _
        & "WHERE X = 'I' "
   Ejecutar_SQL_SP sSQL
   
   RutaGeneraFile = LeftStrg(RutaSysBases, 2) & "\SYSBASES\ARCHIVO_" & Replace(MBFechaI, "/", "-") & ".txt"
   NumFile = FreeFile
   Open RutaGeneraFile For Output As #NumFile ' Abre el archivo.

   sSQL = "SELECT DF.Factura,DF.Fecha,DF.Cantidad,DF.Precio,DF.Precio2,CP.Producto," _
        & "C.Cliente,C.CI_RUC,DF.CodigoB,CP.Codigo_IESS,CP.Marca " _
        & "FROM Detalle_Factura As DF,Clientes As C,Catalogo_Productos As CP " _
        & "WHERE DF.Item = '" & NumEmpresa & "' " _
        & "AND DF.Periodo = '" & Periodo_Contable & "' " _
        & "AND DF.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "AND DF.T <> 'A' " _
        & "AND DF.CodigoC = C.Codigo " _
        & "AND DF.Codigo = CP.Codigo_Inv " _
        & "AND DF.Item = CP.Item " _
        & "AND DF.Periodo = CP.Periodo " _
        & "ORDER BY DF.Fecha,DF.Factura "
        
   sSQL = "SELECT C.Cliente, DF.Codigo As Codigo_Int,DF.Factura,DF.Fecha,DF.Cantidad,CP.PVP,CP.PVP_2,CP.Producto, " _
        & "C.CI_RUC,AB.Beneficiario,AB.RUC_CI,DF.CodigoB,CP.Codigo_IESS,CP.Marca,CP.Ayuda As Producto_IESS " _
        & "FROM Detalle_Factura As DF, Clientes As C, Asiento_Beneficiarios As AB, Catalogo_Productos As CP " _
        & "WHERE DF.Item = '" & NumEmpresa & "' " _
        & "AND DF.Periodo = '" & Periodo_Contable & "' " _
        & "AND DF.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "AND DF.T <> 'A' " _
        & "AND DF.CodigoC = C.Codigo " _
        & "AND DF.CodigoB = AB.Codigo " _
        & "AND DF.Codigo = CP.Codigo_Inv " _
        & "AND DF.Item = CP.Item " _
        & "AND DF.Periodo = CP.Periodo " _
        & "ORDER BY DF.Fecha,DF.Factura "
   Select_Adodc AdoAux, sSQL, , , "IESS"
   With AdoAux.Recordset
    If .RecordCount > 0 Then
        Do While Not .EOF
           CI_RUCC = .fields("CI_RUC")
           NombreC = .fields("Cliente")
           'Producto = .fields("Producto") & " (" & .fields("Marca") & ")"
           Producto = .fields("Producto_IESS") & " (" & .fields("Marca") & ")"
           If CI_RUCC <> .fields("RUC_CI") Then
              CI_RUCC = .fields("RUC_CI")
              NombreC = .fields("Beneficiario")
           End If
           Print #NumFile, Format$(Val(.fields("CI_RUC")), "0000000000");
           Print #NumFile, .fields("Cliente") & String(80 - Len(.fields("Cliente")), " ");
           Print #NumFile, CI_RUCC;
           Print #NumFile, NombreC & String(64 - Len(NombreC), " ");
           Print #NumFile, TrimStrg(.fields("Fecha"));
           Print #NumFile, .fields("Codigo_IESS") & String(40 - Len(.fields("Codigo_IESS")), " ");
           Print #NumFile, "      ";
           Producto = MidStrg(Producto, 1, 80)
           Producto = Replace(Producto, "/", " ")
           Producto = TrimStrg(Producto)
           Print #NumFile, Producto & String(80 - Len(Producto), " ");
           Cadena = Format$(.fields("Cantidad"), "0.00")
           Cadena = Replace(Cadena, ".", ",")
           Print #NumFile, String(13 - Len(Cadena), "0") & Cadena;
           'Cadena = Format$(.Fields("Precio"), "0.0000")
           Cadena = Format$(.fields("PVP"), "0.0000")
           Cadena = Replace(Cadena, ".", ",")
           Print #NumFile, String(18 - Len(Cadena), "0") & Cadena;
           'Cadena = Format$(.Fields("Precio2"), "0.0000")
           Cadena = Format$(.fields("PVP_2"), "0.0000")
           Cadena = Replace(Cadena, ".", ",")
           Print #NumFile, String(15 - Len(Cadena), "0") & Cadena;
           Print #NumFile, Format$(.fields("Factura"), "000000000")
          .MoveNext
        Loop
    End If
   End With
   Close #NumFile
   RatonNormal
   Titulo = "FACTURACION AL IESS"
   Mensajes = "ARCHIVO GENERADO EN:" & vbCrLf & vbCrLf & RutaGeneraFile & vbCrLf & vbCrLf & "Desea generar el reporte a EXCEL?"
   If BoxMensaje = vbYes Then
      DGCxC.Visible = False
      GenerarDataTexto FCierreCaja, AdoAux
      DGCxC.Visible = True
   End If
End Sub


Private Sub Command1_Click()
  Unload FCierreCaja
End Sub

Private Sub DCBenef_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DGCxC_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyP Then
     TextoBanco = DGCxC.Columns(4)
     TextoCheque = DGCxC.Columns(5)
     Mifecha = DGCxC.Columns(1)
     Factura_No = Val(DGCxC.Columns(3))
     Valor = Val(DGCxC.Columns(6))
     Cta = DGCxC.Columns(8)
     If TextoBanco <> "EFECTIVO MN" Then
        Mensajes = "Cheque del: " & TextoBanco & " No. " & TextoCheque & vbCrLf _
                 & "Fecha del Cheque: " & Mifecha & vbCrLf _
                 & "Factura No. " & Factura_No & vbCrLf _
                 & "Valor USD " & Format$(Valor, "#,##0.00")
        Titulo = "CHEQUES PROTESTADOS"
        If BoxMensaje = vbYes Then
           sSQL = "UPDATE Trans_Abonos " _
                & "SET Protestado = " & Val(adTrue) & " " _
                & "WHERE Fecha = #" & BuscarFecha(Mifecha) & "# " _
                & "AND Cta = '" & Cta & "' " _
                & "AND Factura = " & Factura_No & " " _
                & "AND Banco = '" & TextoBanco & "' " _
                & "AND Cheque = '" & TextoCheque & "' "
           Ejecutar_SQL_SP sSQL
        End If
     Else
        MsgBox "No se puede protestar Abonos en Efectivo"
     End If
  End If
End Sub

Private Sub Form_Activate()
Dim AltoMalla As Single

    FCierreCaja.WindowState = vbMaximized
    CtaDeAnticipos = Leer_Seteos_Ctas("Cta_Anticipos_Clientes")
    FormaCierre = Leer_Campo_Empresa("Cierre_Vertical")
'''   Ing_Combo = Leer_Campo_Empresa("Combo")
    SSTab1.width = MDI_X_Max - 100
    SSTab1.Height = MDI_Y_Max - 750
    AltoMalla = (SSTab1.Height - DGVentas.Top) / 3
    Command1.Top = SSTab1.Top + SSTab1.Height + 100
    
    SSTab1.Tab = 5
    DGBanco.width = SSTab1.width - 250
    DGBanco.Height = SSTab1.Height - DGBanco.Top - 100
    
    SSTab1.Tab = 4
    DGSRI.width = SSTab1.width - 250
    DGSRI.Height = SSTab1.Height - DGSRI.Top - 800
    AdoSRI.Top = DGSRI.Top + DGSRI.Height + 10
    Label7.Top = DGSRI.Top + DGSRI.Height + 10
    Label9.Top = DGSRI.Top + DGSRI.Height + 10
    Label12.Top = DGSRI.Top + DGSRI.Height + 10
    Label14.Top = DGSRI.Top + DGSRI.Height + 10
    Label16.Top = DGSRI.Top + DGSRI.Height + 10
    Label18.Top = DGSRI.Top + DGSRI.Height + 10
    LblConIVA.Top = DGSRI.Top + DGSRI.Height + 350
    LblSinIVA.Top = DGSRI.Top + DGSRI.Height + 350
    LblDescuento.Top = DGSRI.Top + DGSRI.Height + 350
    LblIVA.Top = DGSRI.Top + DGSRI.Height + 350
    LblServicio.Top = DGSRI.Top + DGSRI.Height + 350
    LblTotalFacturado.Top = DGSRI.Top + DGSRI.Height + 350

    SSTab1.Tab = 3
    DGFactAnul.width = SSTab1.width - 250
    DGFactAnul.Height = SSTab1.Height - DGFactAnul.Top - 100

    SSTab1.Tab = 2
    DGAsiento.width = SSTab1.width - 250
    DGAsiento1.width = SSTab1.width - 250
    DGAnticipos.width = SSTab1.width - 250
    LblConcepto.width = SSTab1.width - 250
    LblConcepto1.width = SSTab1.width - 250
    
    DGAsiento.Height = (SSTab1.Height / 3) - DGAsiento.Top
    DGAsiento1.Height = DGAsiento.Height
    DGAnticipos.Height = DGAsiento.Height - 30
    
    Label1.Top = DGAsiento.Top + DGAsiento.Height + 10
    Label11.Top = DGAsiento.Top + DGAsiento.Height + 10
    LblDiferencia.Top = DGAsiento.Top + DGAsiento.Height + 10
    LabelDebe.Top = DGAsiento.Top + DGAsiento.Height + 10
    LabelHaber.Top = DGAsiento.Top + DGAsiento.Height + 10

    LblConcepto1.Top = LblDiferencia.Top + LblDiferencia.Height + 200
    DGAsiento1.Top = LblConcepto1.Top + LblConcepto1.Height + 60
    Label13.Top = DGAsiento1.Top + DGAsiento1.Height + 10
    Label15.Top = DGAsiento1.Top + DGAsiento1.Height + 10
    LblDiferencia1.Top = DGAsiento1.Top + DGAsiento1.Height + 10
    LabelDebe1.Top = DGAsiento1.Top + DGAsiento1.Height + 10
    LabelHaber1.Top = DGAsiento1.Top + DGAsiento1.Height + 10
    
    DGAnticipos.Top = LblDiferencia1.Top + LblDiferencia1.Height + 100
    
    SSTab1.Tab = 1
    DGInv.width = SSTab1.width - DGInv.Left - 250
    AdoInv.width = SSTab1.width - DGInv.Left - 250
    DGInv.Height = (SSTab1.Height / 2) - DGInv.Top
    AdoInv.Top = DGInv.Top + DGInv.Height
    DGProductos.Top = AdoInv.Top + AdoInv.Height + 100
    DGProductos.width = SSTab1.width - DGProductos.Left - 200
    DGProductos.Height = SSTab1.Height - DGProductos.Top - 200
    DGCierres.Height = SSTab1.Height - DGCierres.Top - 200

    SSTab1.Tab = 0
    DGVentas.width = SSTab1.width - 250
    DGCxC.width = SSTab1.width - 250
    DGVentas.Height = (SSTab1.Height / 2) - DGAsiento.Top + 100
    DGCxC.Height = DGVentas.Height
    AdoVentas.Top = DGVentas.Height + DGVentas.Top + 15
    Label2.Top = DGVentas.Height + DGVentas.Top + 15
    LabelAbonos.Top = DGVentas.Height + DGVentas.Top + 15

    DGCxC.Top = Label2.Top + Label2.Height + 200
    AdoCxC.Top = DGCxC.Height + DGCxC.Top + 15
    Label4.Top = DGCxC.Height + DGCxC.Top + 15
    LabelCheque.Top = DGCxC.Height + DGCxC.Top + 15

   'Inicializamos los submodulos
    Trans_No = 97
    IniciarAsientosDe DGAsiento1, AdoAsiento1     ' CxC
    Trans_No = 96
    IniciarAsientosDe DGAsiento, AdoAsiento       ' Abonos

'''   AdoSRI.Width = SSTab1.Width - AdoSRI.Left - 200
'''   AdoCxC.Width = SSTab1.Width - AdoCxC.Left - 200
'''   AdoVentas.Width = SSTab1.Width - AdoVentas.Left - 200
   
   Co.TP = CompDiario
   Co.Numero = 0
   Co.RUC_CI = Ninguno
   Co.CodigoB = Ninguno
   Co.Cotizacion = 0
   Co.Beneficiario = Ninguno
   Co.Concepto = ""
   Co.Efectivo = 0
   Co.Total_Banco = 0
   Co.Item = NumEmpresa
   ModificarComp = False
   CopiarComp = False
   NuevoComp = True
      
'''   sSQL = "UPDATE Accesos " _
'''        & "SET Ok = " & Val(adFalse) & " "
'''   Ejecutar_SQL_SP sSQL
'''
'''   If SQL_Server Then
'''      sSQL = "UPDATE Accesos " _
'''           & "SET Ok = " & Val(adTrue) & " " _
'''           & "FROM Accesos As A,Facturas As X "
'''   Else
'''      sSQL = "UPDATE Accesos As A,Facturas As X " _
'''           & "SET Ok = " & Val(adTrue) & " "
'''   End If
'''   sSQL = sSQL & "WHERE A.Codigo = X.CodigoU "
'''   Ejecutar_SQL_SP sSQL
'''
'''   If SQL_Server Then
'''      sSQL = "UPDATE Accesos " _
'''           & "SET Ok = " & Val(adTrue) & " " _
'''           & "FROM Accesos As A,Trans_Abonos As X "
'''   Else
'''      sSQL = "UPDATE Accesos As A,Trans_Abonos As X " _
'''           & "SET Ok = " & Val(adTrue) & " "
'''   End If
'''   sSQL = sSQL & "WHERE A.Codigo = X.CodigoU "
'''   Ejecutar_SQL_SP sSQL

   sSQL = "SELECT (Nombre_Completo & ' - ' & Codigo) As Cajero " _
        & "FROM Accesos " _
        & "WHERE Ok <> " & Val(adFalse) & " " _
        & "ORDER BY Nombre_Completo "
   SelectDB_Combo DCBenef, AdoClientes, sSQL, "Cajero"
         
   Select Case Modulo
     Case "CONTABILIDAD": TbarCierre.buttons("Grabar").Enabled = False
     Case "CAJACREDITO": TbarCierre.buttons("Grabar").Enabled = False
   End Select
   
   If Inv_Promedio Then
      FCierreCaja.Caption = "CIERRE DE CAJA INVENTARIO PRECIO PROMEDIO"
   Else
      FCierreCaja.Caption = "CIERRE DE CAJA INVENTARIO ULTIMO PRECIO"
   End If
   NuevoDiario = False
   
  'IniciarAsientosDe DGAsiento, AdoAsiento
   Mifecha = BuscarFecha(FechaSistema)
   If Bloquear_Control Then
      TbarCierre.buttons("Diario_Caja").Enabled = False
      TbarCierre.buttons("Grabar").Enabled = False
   End If
   RatonNormal
   CierreDelDia
   MBFechaF.Text = FechaSistema
   MBFechaI.Text = PrimerDiaMes(FechaSistema)
   MBFechaI.SetFocus
End Sub

Private Sub Form_Deactivate()
  FCierreCaja.WindowState = vbMinimized
End Sub

Private Sub Form_Load()
   ConectarAdodc AdoAux
   ConectarAdodc AdoSRI
   ConectarAdodc AdoSQL
   ConectarAdodc AdoCxC
   ConectarAdodc AdoCxC1
   ConectarAdodc AdoInv
   ConectarAdodc AdoInv1
   ConectarAdodc AdoBanco
   ConectarAdodc AdoCtaBanco
   ConectarAdodc AdoCierres
   ConectarAdodc AdoVentas
   ConectarAdodc AdoAsiento
   ConectarAdodc AdoAsiento1
   ConectarAdodc AdoClientes
   ConectarAdodc AdoVentaAct
   ConectarAdodc AdoFactAnul
   ConectarAdodc AdoProductos
   ConectarAdodc AdoAnticipos
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaI_LostFocus()
  FechaValida MBFechaI, True
  MBFechaF = MBFechaI
  'LblFechas.Caption = "Cierre de Caja desde el " & FechaStrgDias(MBFechaI) & " al " & FechaStrgDias(MBFechaF)
End Sub

Private Sub MBFechaF_GotFocus()
  MarcarTexto MBFechaF
End Sub

Private Sub MBFechaF_KeyDown(KeyCode As Integer, Shift As Integer)
Keys_Especiales Shift
  If ShiftDown And KeyCode = vbKeyM Then
     MBFechaF = UltimoDiaMes(MBFechaI)
  End If
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaF_LostFocus()
  FechaValida MBFechaF, True
 'LblFechas.Caption = "Cierre de Caja desde el " & FechaStrgDias(MBFechaI.Text) & " al " & FechaStrgDias(MBFechaF.Text)
End Sub

Public Function Insertar_Cta_Cierre(SQLInsert As String, Cta As String, Valor As Currency) As String
Dim tSQLInsert As String
    tSQLInsert = SQLInsert
    If tSQLInsert = "" Then tSQLInsert = "INSERT INTO Asiento (CODIGO, DEBE, HABER, Item, CodigoU, T_No, A_No) VALUES " & vbCrLf
    If Valor > 0 Then tSQLInsert = tSQLInsert & "('" & Cta & "', " & Valor & ", 0, '" & NumEmpresa & "', '" & CodigoUsuario & "', " & Trans_No & ", " & ContAsiento & ")," & vbCrLf
    If Valor < 0 Then tSQLInsert = tSQLInsert & "('" & Cta & "', 0, " & -Valor & ", '" & NumEmpresa & "', '" & CodigoUsuario & "', " & Trans_No & ", " & ContAsiento & ")," & vbCrLf
    If Valor <> 0 Then ContAsiento = ContAsiento + 1
    Insertar_Cta_Cierre = tSQLInsert
End Function

Public Function Insertar_Cta_Cierre_SC(SQLInsert As String, Cta As String, Beneficiario As String, Detalle As String, Serie As String, Factura As Long, Valor As Currency) As String
Dim tSQLInsert As String

    tSQLInsert = SQLInsert
    If tSQLInsert = "" Then
       tSQLInsert = "INSERT INTO Asiento_SC (Codigo, Beneficiario, TC, Serie, Factura, Cta, Fecha_V, Detalle_SubCta, TM, DH, Valor, Item, CodigoU, T_No, SC_No "
       ContAsiento = 0
    End If

    If Valor > 0 Then tSQLInsert = tSQLInsert & "('" & Cta & "', " & Valor & ", 0, '" & NumEmpresa & "', '" & CodigoUsuario & "', " & Trans_No & ", " & ContAsiento & ")," & vbCrLf
    If Valor < 0 Then tSQLInsert = tSQLInsert & "('" & Cta & "', 0, " & -Valor & ", '" & NumEmpresa & "', '" & CodigoUsuario & "', " & Trans_No & ", " & ContAsiento & ")," & vbCrLf
    If Valor <> 0 Then ContAsiento = ContAsiento + 1
    Insertar_Cta_Cierre_SC = tSQLInsert

''                  SetAdoAddNew "Asiento_SC"
''                  SetAdoFields "Codigo", .fields("SubCta")
''                  SetAdoFields "Beneficiario", .fields("Detalle")
''                  SetAdoFields "TM", "1"
''                  SetAdoFields "DH", "2"
''                  SetAdoFields "Valor", Redondear(.fields("TSubTotal"), 2)
''                  SetAdoFields "TC", .fields("TC")
''                  SetAdoFields "Cta", .fields("Cta_Venta")
''                  SetAdoFields "T_No", Trans_No
''                  SetAdoFields "SC_No", ContSC
''
''                  SetAdoAddNew "Asiento_SC"
''                  SetAdoFields "FECHA_V", .fields("Fecha")
''                  SetAdoFields "TC", .fields("Tipo_Cta")
''                  SetAdoFields "Cta", .fields("Cta")
''                  SetAdoFields "Detalle_SubCta", "Abono de " & .fields("TP") & ": " & .fields("Serie") & "-" & Format(.fields("Factura"), "000000000")
''                  SetAdoFields "T_No", Trans_No
''                  SetAdoFields "SC_No", ContSC
''                  If NumeroFASubModulo Then
''                     SetAdoFields "Serie", .fields("Serie")
''                     SetAdoFields "Factura", .fields("Factura")
''                  Else
''                     SetAdoFields "Serie", "001001"
''                     SetAdoFields "Factura", 0
''                  End If


End Function

Public Sub Grabar_Asientos_Facturacion(TipoConsulta As String)
Dim AdoDBAux As ADODB.Recordset
Dim VentasDia As Boolean
Dim Ctas_Catalogo As String
Dim ErrorTemp As String
Dim sSQLIns As String
Dim sSQLInsSC As String
Dim sSQLInsInv As String

Dim Total_Vaucher As Currency
Dim T_No As Byte
Dim NoMes As Byte

   RatonReloj
   Progreso_Barra.Mensaje_Box = ""
   Progreso_Iniciar
   
   If MBFechaI.Text = MBFechaF.Text Then FechaDelCierre = MBFechaI.Text Else FechaDelCierre = "desde el " & MBFechaI.Text & " al " & MBFechaF.Text

   DGCxC.Visible = False
   DGInv.Visible = False
   DGVentas.Visible = False
   DGAsiento.Visible = False
   
   FechaValida MBFechaI
   FechaValida MBFechaF
   FechaIni = BuscarFecha(MBFechaI)
   FechaFin = BuscarFecha(MBFechaF)
   Fecha_Vence = MBFechaF
   
   ErrorInventario = ""
   Total_Vaucher = 0
   Total_Propinas = 0
   Total = 0
   VentasDia = False
   
  'MsgBox sSQL
   Progreso_Barra.Mensaje_Box = "Verificando Cuentas involucradas"
   Progreso_Esperar
     
  'Asientos de Abonos y CxC segun la fecha de proceso:
   RatonReloj
   Procesar_Diario_Caja_SP MBFechaI, MBFechaF, CheqCajero.value, SinEspaciosDer(DCBenef.Text)
   
  'Listado de los tipos de abonos
   sSQL = "SELECT TA.Fecha, C.Cliente, TA.TP, TA.Serie, TA.Autorizacion, TA.Factura, TA.Banco, TA.Cheque, TA.Abono, C.CI_RUC As COD_BANCO, TA.Comprobante, TA.Tipo_Cta, " _
        & "TA.Cta, TA.Cta_CxP, TA.CodigoC, C.Ciudad, C.Plan_Afiliado As Sectorizacion, A.Nombre_Completo As Ejecutivo, Recibo_No As Orden_No " _
        & "FROM Trans_Abonos As TA, Clientes C, Accesos As A " _
        & "WHERE TA.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "AND TA.TP <> 'OP' " _
        & "AND TA.T <> 'A' " _
        & "AND TA.Item = '" & NumEmpresa & "' " _
        & "AND TA.Periodo = '" & Periodo_Contable & "' " _
        & "AND TA.CodigoC = C.Codigo " _
        & "AND TA.Cod_Ejec = A.Codigo "
   If CheqCajero.value = 1 Then sSQL = sSQL & "AND TA.CodigoU = '" & SinEspaciosDer(DCBenef.Text) & "' "
   If CheqOrdDep.value = 1 Then
      sSQL = sSQL & "ORDER BY TA.Fecha,TA.TP,TA.Cta,TA.Banco,C.Cliente,TA.Factura "
   Else
      sSQL = sSQL & "ORDER BY TA.Fecha,TA.TP,TA.Cta,C.Cliente,TA.Banco,TA.Factura "
   End If
   Select_Adodc_Grid DGCxC, AdoCxC, sSQL
   DGCxC.Caption = UCase("Abonos del dia " & FechaDelCierre)
   RatonReloj
   With AdoCxC.Recordset
    If .RecordCount > 0 Then
        Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + (.RecordCount * 2)
       'sSQLIns = Insertar_Cta_Cierre(sSQLIns, Cta, Sumatoria)
        Do While Not .EOF
''           Progreso_Barra.Mensaje_Box = "Totalizando Abonos: " & .fields("TP") & " - " & .fields("Cta") & " - " & .fields("Cta_CxP") & " -> " & .fields("Cliente")
''           Progreso_Esperar
           Total = Total + Redondear(.fields("Abono"), 2)
          .MoveNext
        Loop
    End If
   End With
   LabelCheque.Caption = Format$(Total, "#,##0.00")

  'Listado de las CxC y Ventas Clientes
   Total = 0
   sSQL = "SELECT F.Fecha,C.Cliente,F.TC,F.Serie,F.Autorizacion,F.Factura,F.IVA As Total_IVA,F.Descuento," _
        & "F.Descuento2,F.Servicio,F.Propina,F.Total_MN,F.Saldo_MN,F.Cta_CxP,C.Ciudad,C.Plan_Afiliado As Sectorizacion," _
        & "A.Nombre_Completo As Ejecutivo, F.Nota, F.Observacion, CSC.Detalle As Centro_Costo " _
        & "FROM Facturas F, Clientes C, Accesos As A, Catalogo_SubCtas As CSC " _
        & "WHERE F.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "AND F.TC <> 'OP' " _
        & "AND F.T <> 'A' " _
        & "AND F.Item = '" & NumEmpresa & "' " _
        & "AND F.Periodo = '" & Periodo_Contable & "' "
   If CheqCajero.value = 1 Then sSQL = sSQL & "AND F.CodigoU = '" & SinEspaciosDer(DCBenef.Text) & "' "
   sSQL = sSQL _
        & "AND F.CodigoC = C.Codigo " _
        & "AND F.Cod_Ejec = A.Codigo " _
        & "AND F.Item = CSC.Item " _
        & "AND F.Periodo = CSC.Periodo " _
        & "AND F.SubCta = CSC.Codigo " _
        & "ORDER BY F.TC,F.Fecha,F.Cta_CxP,F.Factura,C.Cliente "
   Select_Adodc_Grid DGVentas, AdoVentas, sSQL
   DGVentas.Caption = UCase("Ventas del dia " & FechaDelCierre)
   RatonReloj
   With AdoVentas.Recordset
    If .RecordCount > 0 Then
        Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + (.RecordCount * 2)
        Do While Not .EOF
''           Progreso_Barra.Mensaje_Box = "Totalizando Ventas: " & .fields("TC") & " - " & .fields("Cta_CxP") & " -> " & .fields("Cliente")
''           Progreso_Esperar
           Select Case .fields("TC")
             Case "FA", "NV": Total = Total + Redondear(.fields("Total_MN"), 2)
           End Select
          .MoveNext
        Loop
    End If
   End With
   LabelAbonos.Caption = Format$(Total, "#,##0.00")
   
  'Resumen de abonos anticipados de Clientes
   sSQL = "SELECT CC.Cuenta, C.Cliente, TS.Fecha, TS.TP, TS.Numero, TS.Creditos, T.Cta AS Contra_Cta, TS.Cta " _
        & "FROM Trans_SubCtas AS TS, Transacciones AS T, Catalogo_Cuentas AS CC, Clientes AS C " _
        & "WHERE TS.Item = '" & NumEmpresa & "' " _
        & "AND TS.Periodo = '" & Periodo_Contable & "' " _
        & "AND TS.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
        & "AND TS.T <> 'A' " _
        & "AND TS.Cta = '" & CtaDeAnticipos & "' " _
        & "AND TS.Creditos > 0 " _
        & "AND TS.Periodo = T.Periodo " _
        & "AND TS.Periodo = CC.Periodo " _
        & "AND TS.Item = T.Item " _
        & "AND TS.Item = CC.Item " _
        & "AND TS.TP = T.TP " _
        & "AND TS.Numero = T.Numero " _
        & "AND T.Cta = CC.Codigo " _
        & "AND TS.Codigo = C.Codigo " _
        & "AND TS.Cta <> T.Cta " _
        & "ORDER BY T.Cta, C.Cliente, TS.Fecha, TS.TP, TS.Numero "
        
   sSQL = "SELECT * " _
        & "FROM Asiento_SC " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND T_No IN (96,97) " _
        & "AND CodigoU = '" & CodigoUsuario & "' " _
        & "ORDER BY T_No, Cta, SC_No, Codigo "
   Select_Adodc_Grid DGAnticipos, AdoAnticipos, sSQL
   
  'Listado de Facturas anuladas
   Total = 0
   sSQL = "SELECT F.T,F.TC,F.Fecha,C.Cliente,F.Factura,F.IVA As Total_IVA,F.Total_MN,F.Cta_CxP " _
        & "FROM Facturas F, Clientes C " _
        & "WHERE F.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "AND F.T = 'A' " _
        & "AND F.Item = '" & NumEmpresa & "' " _
        & "AND F.Periodo = '" & Periodo_Contable & "' " _
        & "AND F.TC <> 'OP' "
   If CheqCajero.value = 1 Then sSQL = sSQL & "AND F.CodigoU = '" & SinEspaciosDer(DCBenef.Text) & "' "
   sSQL = sSQL _
        & "AND F.CodigoC = C.Codigo " _
        & "ORDER BY F.TC,F.Fecha,F.Cta_CxP,C.Cliente,F.Factura "
   Select_Adodc_Grid DGFactAnul, AdoFactAnul, sSQL
    
  'REPORTES DE AUDITORIA TRANSACCIONALES (S.R.I.)
   DGSRI.Caption = "Autorización No. " & Autorizacion & ", Listado de Facturas del " & FechaDelCierre
   Codigo = CStr(Porc_IVA * 100)
   Total_Con_IVA = 0
   Total_Sin_IVA = 0
   Total_Desc = 0
   Total_Desc2 = 0
   Total_IVA = 0
   Total = 0
   sSQL = "SELECT F.TC,F.T,F.RUC_CI,F.TB,F.Razon_Social,F.Fecha,F.Hora,A.Nombre_Completo As Usuario," _
        & "F.Autorizacion,F.Serie,F.Factura As Secuencial,F.Con_IVA As Base_" & Codigo & ",F.Sin_IVA As Base_0," _
        & "F.Descuento,F.Descuento2,(F.SubTotal - F.Descuento - F.Descuento2) As Sub_Total, F.IVA As IVA_" & Codigo & ",F.Servicio,F.Total_MN As TOTAL,Serie_R," _
        & "Secuencial_R,F.Autorizacion_R,Total_Ret_Fuente,Total_Ret_IVA_B,Total_Ret_IVA_S,C.Contacto AS Referencia,C.CI_RUC As COD_BANCO " _
        & "FROM Facturas F, Clientes C, Accesos As A " _
        & "WHERE F.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "AND F.TC NOT IN ('C','P','OP','LC','DO') " _
        & "AND F.Item = '" & NumEmpresa & "' " _
        & "AND F.Periodo = '" & Periodo_Contable & "' " _
        & "AND F.CodigoC = C.Codigo " _
        & "AND F.CodigoU = A.Codigo "
   If CheqCajero.value = 1 Then sSQL = sSQL & "AND F.CodigoU = '" & SinEspaciosDer(DCBenef.Text) & "' "
   sSQL = sSQL & "ORDER BY F.Factura,F.TC,F.Fecha,F.Cta_CxP,C.Cliente "
   DGSRI.Caption = "Resumen de Comprobantes Generados " & FechaDelCierre
   Select_Adodc_Grid DGSRI, AdoSRI, sSQL
   With AdoSRI.Recordset
    If .RecordCount > 0 Then
        Do While Not .EOF
           If .fields("T") <> Anulado Then
               Total_Con_IVA = Total_Con_IVA + .fields("Base_" & Codigo)
               Total_Sin_IVA = Total_Sin_IVA + .fields("Base_0")
               Total_Desc = Total_Desc + .fields("Descuento")
               Total_Desc2 = Total_Desc2 + .fields("Descuento2")
               Total_IVA = Total_IVA + .fields("IVA_" & Codigo)
               Total_Servicio = Total_Servicio + .fields("Servicio")
               Total = Total + .fields("TOTAL")
           End If
          .MoveNext
        Loop
    End If
   End With
   LblConIVA.Caption = Format$(Total_Con_IVA, "#,##0.00")
   LblSinIVA.Caption = Format$(Total_Sin_IVA, "#,##0.00")
   LblDescuento.Caption = Format$(Total_Desc + Total_Desc2, "#,##0.00")
   LblIVA.Caption = Format$(Total_IVA, "#,##0.00")
   LblServicio.Caption = Format$(Total_Servicio, "#,##0.00")
   LblTotalFacturado.Caption = Format$(Total, "#,##0.00")
   
   sSQL = "SELECT TC,Codigo,Producto,SUM(Cantidad) As CANTIDADES,SUM(Total) As SUBTOTALES,SUM(Total_IVA) As SUBTOTAL_IVA, Cta_Venta " _
        & "FROM Detalle_Factura " _
        & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND T <> '" & Anulado & "' " _
        & "AND TC IN ('FA','NV') " _
        & "AND Salida_Inv = 0 " _
        & "GROUP BY TC,Codigo,Producto,Cta_Venta " _
        & "UNION " _
        & "SELECT TC,'-x-' As Codigo,'TOTAL DE VENTAS' As Producto,SUM(Cantidad) As CANTIDADES,SUM(Total) As SUBTOTALES,SUM(Total_IVA) As SUBTOTAL_IVA,'' As Cta_Venta " _
        & "FROM Detalle_Factura " _
        & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND T <> '" & Anulado & "' " _
        & "AND TC IN ('FA','NV') " _
        & "AND Salida_Inv = 0 " _
        & "GROUP BY TC " _
        & "ORDER BY TC,Codigo,Producto "
   Select_Adodc_Grid DGProductos, AdoProductos, sSQL
   DGProductos.Caption = "Resumen de Productos Facturados de Ventas " & FechaDelCierre
  
  'Asiento de Entrada y Salida de Inventario por NC
   sSQL = "SELECT TK.TC As Doc, TK.Codigo_Inv, CP.Producto, 0 As Entradas, SUM(TK.Salida) As Salidas, AVG(TK.Costo) As Costos, " _
        & "(SUM(TK.Salida) * AVG(TK.Costo)) As Totales, TK.Cta_Inv, TK.Contra_Cta, TK.Codigo_Barra, TK.CodBodega, CP.Unidad, TK.Cmds, COUNT(TK.TC) As Cant_Doc " _
        & "FROM Trans_Kardex As TK, Catalogo_Productos As CP " _
        & "WHERE TK.Item = '" & NumEmpresa & "' " _
        & "AND TK.Periodo = '" & Periodo_Contable & "' " _
        & "AND TK.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "AND TK.TP = '.' " _
        & "AND LEN(TK.TC) >= 2 " _
        & "AND LEN(TK.Serie) = 6 " _
        & "AND TK.Factura > 0 " _
        & "AND TK.Salida > 0 "
   If CheqCajero.value = 1 Then sSQL = sSQL & "AND TK.CodigoU = '" & SinEspaciosDer(DCBenef.Text) & "' "
   sSQL = sSQL _
        & "AND TK.Item = CP.Item " _
        & "AND TK.Periodo = CP.Periodo " _
        & "AND TK.Codigo_Inv = CP.Codigo_Inv " _
        & "GROUP BY TK.TC, TK.Codigo_Inv, CP.Producto, TK.Cta_Inv, TK.Contra_Cta, TK.Codigo_Barra, TK.CodBodega, CP.Unidad, TK.Cmds "
   sSQL = sSQL & "UNION " _
        & "SELECT 'NC' As Doc, TK.Codigo_Inv, CP.Producto, SUM(TK.Entrada) As Entradas, 0 As Salidas, AVG(TK.Costo) As Costos, " _
        & "(SUM(TK.Entrada) * AVG(TK.Costo)) As Totales, TK.Cta_Inv, TK.Contra_Cta, TK.Codigo_Barra, TK.CodBodega, CP.Unidad, TK.Cmds, COUNT(TK.TC) As Cant_Doc " _
        & "FROM Trans_Kardex As TK, Catalogo_Productos As CP " _
        & "WHERE TK.Item = '" & NumEmpresa & "' " _
        & "AND TK.Periodo = '" & Periodo_Contable & "' " _
        & "AND TK.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "AND TK.TP = '.' " _
        & "AND LEN(TK.TC) >= 2 " _
        & "AND LEN(TK.Serie) = 6 " _
        & "AND TK.Factura > 0 " _
        & "AND TK.Entrada > 0 "
   If CheqCajero.value = 1 Then sSQL = sSQL & "AND TK.CodigoU = '" & SinEspaciosDer(DCBenef.Text) & "' "
   sSQL = sSQL _
        & "AND TK.Item = CP.Item " _
        & "AND TK.Periodo = CP.Periodo " _
        & "AND TK.Codigo_Inv = CP.Codigo_Inv " _
        & "GROUP BY TK.Codigo_Inv, CP.Producto, TK.Cta_Inv, TK.Contra_Cta, TK.Codigo_Barra, TK.CodBodega, CP.Unidad, TK.Cmds " _
        & "ORDER BY Doc, TK.Codigo_Inv, CP.Producto, TK.Cta_Inv, TK.Contra_Cta, TK.Codigo_Barra, TK.CodBodega, CP.Unidad, TK.Cmds "
   SQLDec = "Costos " & CStr(Dec_Costo) & "|."
   Select_Adodc_Grid DGInv, AdoInv, sSQL, SQLDec
   DGInv.Caption = "Resumen de Inventario de Salida " & FechaDelCierre
     
   Combos = Ninguno
   FechaFinal = BuscarFecha("31/12/" & FechaAnio(MBFechaF))

   Select Case TipoConsulta
     Case Procesado: NuevoDiario = False
     Case Normal:    NuevoDiario = True
   End Select
   RatonReloj
  '==============================================================================================================================
  ' Totalizamos los dos asientos para ver descuadres
  '==============================================================================================================================
   Trans_No = 96
   Debe = 0
   Haber = 0
   SumaDebe = 0
   SumaHaber = 0
   Ln_No = 0
   SQL2 = "SELECT * " _
        & "FROM Asiento " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND T_No = " & Trans_No & " " _
        & "AND CodigoU = '" & CodigoUsuario & "' " _
        & "ORDER BY CODIGO,DEBE DESC,HABER "
   Select_Adodc_Grid DGAsiento, AdoAsiento, SQL2
  'Verificacion SubTotal
   With AdoAsiento.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
        Do While Not .EOF
           Debe = Debe + .fields("DEBE")
           Haber = Haber + .fields("HABER")
           SumaDebe = SumaDebe + .fields("DEBE")
           SumaHaber = SumaHaber + .fields("HABER")
          
          .fields("A_No") = Ln_No
           Ln_No = Ln_No + 1
          .MoveNext
        Loop
       .UpdateBatch
    End If
   End With
   LabelDebe.Caption = Format$(Debe, "#,##0.00")
   LabelHaber.Caption = Format$(Haber, "#,##0.00")
   LblDiferencia.Caption = Format$(Debe - Haber, "#,##0.00")
   LblConcepto.Caption = "Cierre Diario de Caja de CxC " & FechaDelCierre & ", Diario No. ?"
   
   Trans_No = 97
   Debe = 0: Haber = 0: Ln_No = 0
   SQL2 = "SELECT * " _
        & "FROM Asiento " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND T_No = " & Trans_No & " " _
        & "AND CodigoU = '" & CodigoUsuario & "' " _
        & "ORDER BY CODIGO,DEBE DESC,HABER "
   Select_Adodc_Grid DGAsiento1, AdoAsiento1, SQL2
   With AdoAsiento1.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
        Do While Not .EOF
           Debe = Debe + .fields("DEBE")
           Haber = Haber + .fields("HABER")
           SumaDebe = SumaDebe + .fields("DEBE")
           SumaHaber = SumaHaber + .fields("HABER")
          .fields("A_No") = Ln_No
           Ln_No = Ln_No + 1
          .MoveNext
        Loop
       .UpdateBatch
    End If
   End With
'  LabelVentas.Caption = Format$(TotalIngreso, "#,##0.00")
   LabelDebe1.Caption = Format$(Debe, "#,##0.00")
   LabelHaber1.Caption = Format$(Haber, "#,##0.00")
   LblDiferencia1.Caption = Format$(Debe - Haber, "#,##0.00")
   LblConcepto1.Caption = "Cierre Diario de Caja de Abonos " & FechaDelCierre & ", Diario No. ?"
     
   DGVentas.Visible = True
   DGCxC.Visible = True
   DGInv.Visible = True
   DGAsiento.Visible = True
   FCierreCaja.Caption = "CIERRE DEL DIARIO DE CAJA"
  'TextoImprimio
   If Redondear(SumaDebe - SumaHaber, 2) <> 0 Then MsgBox "Las Transacciones no cuadran, verifique las facturas emitidas o los abonos del día."
   Progreso_Barra.Mensaje_Box = "OK"
   Progreso_Final
  'MsgBox TextoImprimio
End Sub

Public Sub CierreDelDia()
Dim ErrorCierre As String
  ErrorCierre = ""
  sSQL = "SELECT TOP 1 TP, Serie, Factura, MIN(Fecha) As FechaCierre " _
       & "FROM Trans_Abonos " _
       & "WHERE C = " & Val(adFalse) & " " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND T <> 'A' " _
       & "AND Fecha <= #" & BuscarFecha(FechaSistema) & "# " _
       & "GROUP BY TP, Serie, Factura " _
       & "ORDER BY TP, Serie, Factura "
  Select_Adodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       ErrorCierre = ErrorCierre & "Cierre del día Abonos: " & .fields("FechaCierre") & " (" & .fields("TP") & "-" & .fields("Serie") & "-" & .fields("Factura") & ")" & vbCrLf & vbCrLf
   End If
  End With
  sSQL = "SELECT TOP 1 TC, Serie, Factura, MIN(Fecha) As FechaCierre " _
       & "FROM Facturas " _
       & "WHERE C = " & Val(adFalse) & " " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC NOT IN ('C','P') " _
       & "AND T <> 'A' " _
       & "AND Fecha <= #" & BuscarFecha(FechaSistema) & "# " _
       & "GROUP BY TC, Serie, Factura " _
       & "ORDER BY TC, Serie, Factura "
  Select_Adodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       ErrorCierre = ErrorCierre & "Cierre del día Facturas: " & .fields("FechaCierre") & " (" & .fields("TC") & "-" & .fields("Serie") & "-" & .fields("Factura") & ")" & vbCrLf & vbCrLf
   End If
  End With
  
  If ErrorCierre <> "" Then
     MarcarTexto MBFechaI
     MsgBox ErrorCierre
     MBFechaI.SetFocus
  End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Select Case SSTab1.Tab
      Case 1
            RatonReloj
            Progreso_Barra.Mensaje_Box = "Fechas de Cierre..."
            Progreso_Esperar
            sSQL = "SELECT Fecha " _
                 & "FROM Facturas " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "AND TC <> 'OP' " _
                 & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
                 & "UNION " _
                 & "SELECT Fecha " _
                 & "FROM Trans_Abonos " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "AND TP <> 'OP' " _
                 & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
                 & "GROUP BY Fecha " _
                 & "ORDER BY Fecha "
            Select_Adodc_Grid DGCierres, AdoCierres, sSQL
            DGCierres.Caption = "Dias Cierres"
            RatonNormal
      Case 2 'Asientos Contables
      
      Case 3
      
      Case 4
      
      Case 5
            sSQL = "SELECT (Codigo & Space(5) & Cuenta) As NomCuenta " _
                 & "FROM Catalogo_Cuentas " _
                 & "WHERE TC = 'BA' " _
                 & "AND DG = 'D' " _
                 & "AND Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "ORDER BY Codigo "
            SelectDB_Combo DCBanco, AdoCtaBanco, sSQL, "NomCuenta"
    End Select
End Sub

Private Sub TbarCierre_ButtonClick(ByVal Button As ComctlLib.Button)
   'MsgBox Button.key
    Select Case Button.key
      Case "Salir"
           Unload FCierreCaja
      Case "Diario_Caja"
           Procesar_Diario_Caja
      Case "Grabar"
           Grabar_Cierre_Diario
      Case "Cuadre_Caja"
           RatonReloj
           'FCuadreCaja.Show 1
      Case "Diario"
           Nombre_Cajero = Ninguno
           If CheqCajero.value = 1 Then
              Nombre_Cajero = MidStrg(DCBenef.Text, 1, Len(DCBenef.Text) - Len(SinEspaciosDer(DCBenef.Text)) - 1)
           End If
          'MsgBox FormaCierre
           If FormaCierre Then
              Imprimir_Diario_Caja AdoVentas, AdoCxC, AdoInv, AdoProductos, AdoAnticipos, MBFechaI, MBFechaF
           Else
              Imprimir_Diario_Caja_Resumen AdoVentas, AdoCxC, AdoInv, AdoProductos, AdoAnticipos, MBFechaI, MBFechaF
           End If
      Case "Asiento"
           DGAsiento.Visible = False
           MensajeEncabData = "RESUMEN DE VENTAS"
           SQLMsg1 = "Corte del " & MBFechaI.Text & " al " & MBFechaF.Text
           sSQL = "SELECT CODIGO,CUENTA,PARCIAL_ME,DEBE,HABER " _
                & "FROM Asiento " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND T_No = " & Trans_No & " " _
                & "AND CodigoU = '" & CodigoUsuario & "' "
           Select_Adodc AdoAsiento, sSQL
            
           ImprimirResumenAsientoCaja AdoAsiento
            
           sSQL = "SELECT * " _
                & "FROM Asiento " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND T_No = " & Trans_No & " " _
                & "AND CodigoU = '" & CodigoUsuario & "' "
           Select_Adodc_Grid DGAsiento, AdoAsiento, sSQL
           DGAsiento.Visible = True
      Case "Reactivar"
           If ClaveContador Then
              FechaValida MBFechaI
              FechaValida MBFechaF
              FechaIni = BuscarFecha(MBFechaI)
              FechaFin = BuscarFecha(MBFechaF)
              sSQL = "UPDATE Trans_Abonos " _
                   & "SET C = " & Val(adFalse) & " " _
                   & "WHERE Item = '" & NumEmpresa & "' " _
                   & "AND Periodo = '" & Periodo_Contable & "' " _
                   & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
              Ejecutar_SQL_SP sSQL
         
              sSQL = "UPDATE Facturas " _
                   & "SET C = " & Val(adFalse) & " " _
                   & "WHERE Item = '" & NumEmpresa & "' " _
                   & "AND Periodo = '" & Periodo_Contable & "' " _
                   & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
              Ejecutar_SQL_SP sSQL
              Trans_No = 97
              IniciarAsientosDe DGAsiento1, AdoAsiento1
              Trans_No = 96
              IniciarAsientosDe DGAsiento, AdoAsiento
              RatonNormal
              LabelDebe.Caption = Format$(0, "#,##0.00")
              LabelHaber.Caption = Format$(0, "#,##0.00")
              CierreDelDia
              MBFechaI.SetFocus
           End If
      Case "SRI"
            DGSRI.Visible = False
            If MBFechaI.Text = MBFechaF.Text Then
               SQLMsg3 = "Autorización No. " & Autorizacion & ", Listado de Facturas del " & MBFechaI.Text
            Else
               SQLMsg3 = "Autorización No. " & Autorizacion & ", Listado de Facturas del " & MBFechaI.Text & " al " & MBFechaF.Text
            End If
            MensajeEncabData = "RESUMEN DE FACTURAS EMITIDAS"
            With AdoSRI.Recordset
             If .RecordCount > 0 Then
                .MoveFirst
                 SQLMsg2 = "Facturas desde No. " & .fields("Secuencial")
                .MoveLast
                 SQLMsg2 = SQLMsg2 & " Hasta la No. " & .fields("Secuencial")
             End If
            End With
            SQLMsg1 = "TIPO DE DOCUMENTO:  NOTAS DE VENTA"
            ImprimirAdo_SRI AdoSRI, 7
            DGSRI.Visible = True
      Case "IESS"
           IESS_Cierre_Diario
      Case "Anuladas"
            MensajeEncabData = "FACTURAS ANULADAS"
            If MBFechaI.Text = MBFechaF.Text Then
               SQLMsg3 = "Diario de Caja del " & MBFechaI
            Else
               SQLMsg3 = "Diario de Caja del " & MBFechaI & " al " & MBFechaF
            End If
            ImprimirAdo AdoFactAnul, True, 1, 8
      Case "Comparar"
        '''  CDialogDir.InitDir = RutaSysBases 'LeftStrg(CurDir$, 3)
        '''  RutaOrigen = UCaseStrg(SelectZipFile(CDialogDir, SelectAll))
        '''  If RutaOrigen <> "" Then
        '''    'Le pasamos el Path del Libro y una variable de tipo T_Rango para retornar los valores
        '''    ' Call Obtener_Rango_Excel(RutaOrigen)
        '''     Call Leer_Excel_AdoDB(RutaOrigen)
        '''  End If
      Case "Excel"
           Select Case SSTab1.Tab
             Case 0
                  DGVentas.Visible = False
                  DGCxC.Visible = False
                  GenerarDataTexto FCierreCaja, AdoVentas
                  GenerarDataTexto FCierreCaja, AdoCxC
                  DGVentas.Visible = True
                  DGCxC.Visible = True
             Case 1
                 'DGCierres.Visible = False
                  DGInv.Visible = False
                  DGProductos.Visible = False
                  GenerarDataTexto FCierreCaja, AdoInv
                  GenerarDataTexto FCierreCaja, AdoProductos
                 'DGCierres.Visible = True
                  DGInv.Visible = True
                  DGProductos.Visible = True
             Case 2
                  DGAsiento.Visible = False
                  DGAsiento1.Visible = False
                  DGAnticipos.Visible = False
                  GenerarDataTexto FCierreCaja, AdoAsiento
                  GenerarDataTexto FCierreCaja, AdoAsiento1
                  GenerarDataTexto FCierreCaja, AdoAnticipos
                  DGAsiento.Visible = True
                  DGAsiento1.Visible = True
                  DGAnticipos.Visible = True
             Case 3
                  DGFactAnul.Visible = False
                  GenerarDataTexto FCierreCaja, AdoFactAnul
                  DGFactAnul.Visible = True
             Case 4
                  DGSRI.Visible = False
                  GenerarDataTexto FCierreCaja, AdoSRI
                  DGSRI.Visible = True
             Case 5
                  DGBanco.Visible = False
                  GenerarDataTexto FCierreCaja, AdoBanco
                  DGBanco.Visible = True
           End Select
'''      Case "Kardex"
'''           Progreso_Barra.Mensaje_Box = "Actualizando Productos"
'''           Progreso_Esperar
'''           Insertar_Productos_Cierre_Caja_SP MBFechaI, MBFechaF
'''           Progreso_Final
    End Select
End Sub
