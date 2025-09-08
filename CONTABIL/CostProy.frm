VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Begin VB.Form FCostosDelProyecto 
   Caption         =   "ESTADO DE CUENTAS"
   ClientHeight    =   8985
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   11280
   DrawMode        =   5  'Not Copy Pen
   DrawStyle       =   1  'Dash
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8985
   ScaleWidth      =   11280
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   5
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir del Modulo"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Costos"
            Object.ToolTipText     =   "Presenta Resumen de Costos"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "UpdateCostos"
            Object.ToolTipText     =   "Actualiza Costos modificados"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Excel"
            Object.ToolTipText     =   "Enviar a Excel"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Grabar"
            Object.ToolTipText     =   "Grabar Comprobante de Proyecto"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
      EndProperty
      Begin VB.Frame Frame3 
         Caption         =   "Fechas Desde - Hasta"
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
         Left            =   3045
         TabIndex        =   1
         Top             =   0
         Width           =   14190
         Begin MSMask.MaskEdBox MBoxFechaF 
            Height          =   330
            Left            =   1470
            TabIndex        =   3
            Top             =   210
            Width           =   1275
            _ExtentX        =   2249
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
         Begin MSMask.MaskEdBox MBoxFechaI 
            Height          =   330
            Left            =   105
            TabIndex        =   2
            Top             =   210
            Width           =   1275
            _ExtentX        =   2249
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
         Begin MSDataListLib.DataCombo DCCtaProyecto 
            Bindings        =   "CostProy.frx":0000
            DataSource      =   "AdoCtaProyecto"
            Height          =   345
            Left            =   4830
            TabIndex        =   5
            Top             =   210
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   609
            _Version        =   393216
            Text            =   "Proyecto"
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
         Begin VB.Label Label1 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Cuenta del Proyecto"
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
            Left            =   2835
            TabIndex        =   4
            Top             =   210
            Width           =   2010
         End
      End
   End
   Begin VB.ListBox LstCtaProy 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   105
      Style           =   1  'Checkbox
      TabIndex        =   6
      Top             =   1155
      Width           =   8520
   End
   Begin VB.ListBox LstCta 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   8715
      Style           =   1  'Checkbox
      TabIndex        =   23
      Top             =   1155
      Visible         =   0   'False
      Width           =   8520
   End
   Begin MSDataGridLib.DataGrid DGAsiento 
      Bindings        =   "CostProy.frx":001D
      Height          =   2430
      Left            =   105
      TabIndex        =   11
      Top             =   4725
      Width           =   16815
      _ExtentX        =   29660
      _ExtentY        =   4286
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
            LCID            =   12298
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
            LCID            =   12298
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
   Begin VB.TextBox TxtConcepto 
      Height          =   540
      Left            =   1680
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   22
      Top             =   4095
      Width           =   15555
   End
   Begin MSAdodcLib.Adodc AdoAux1 
      Height          =   330
      Left            =   630
      Top             =   6825
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
   Begin VB.CommandButton Command1 
      Caption         =   "&S"
      Height          =   300
      Left            =   17430
      TabIndex        =   18
      Top             =   2940
      Width           =   435
   End
   Begin MSDataListLib.DataCombo DCCta 
      Bindings        =   "CostProy.frx":0036
      DataSource      =   "AdoCtas"
      Height          =   345
      Left            =   17430
      TabIndex        =   8
      Top             =   1785
      Visible         =   0   'False
      Width           =   7260
      _ExtentX        =   12806
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
   Begin MSDataListLib.DataCombo DCCtas 
      Bindings        =   "CostProy.frx":004C
      DataSource      =   "AdoSubCta"
      Height          =   345
      Left            =   10920
      TabIndex        =   10
      Top             =   3465
      Visible         =   0   'False
      Width           =   6315
      _ExtentX        =   11139
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
   Begin VB.PictureBox PictFactura 
      Height          =   330
      Left            =   14280
      ScaleHeight     =   270
      ScaleWidth      =   4785
      TabIndex        =   17
      Top             =   11445
      Width           =   4845
   End
   Begin VB.CheckBox CheqCta 
      Caption         =   "Por Cta."
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
      Left            =   8715
      TabIndex        =   7
      Top             =   735
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc AdoSubCta 
      Height          =   330
      Left            =   630
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
      Caption         =   "SubCta"
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
   Begin VB.CheckBox CheqIndiv 
      Caption         =   "Por Centro de Costo"
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
      Left            =   8715
      TabIndex        =   9
      Top             =   3465
      Width           =   2115
   End
   Begin MSAdodcLib.Adodc AdoCtaProyecto 
      Height          =   330
      Left            =   630
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
      Caption         =   "CtaProyecto"
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
      Left            =   630
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
   Begin MSAdodcLib.Adodc AdoAsiento 
      Height          =   330
      Left            =   315
      Top             =   11445
      Width           =   4950
      _ExtentX        =   8731
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
      Caption         =   "AsientoSC"
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
      Left            =   630
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
   Begin MSAdodcLib.Adodc AdoDet 
      Height          =   330
      Left            =   630
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
      Caption         =   "Det"
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
   Begin MSAdodcLib.Adodc AdoAsientoSC 
      Height          =   330
      Left            =   630
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
   Begin MSDataGridLib.DataGrid DGAsientoSC 
      Bindings        =   "CostProy.frx":0064
      Height          =   3270
      Left            =   105
      TabIndex        =   12
      Top             =   7875
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   5768
      _Version        =   393216
      AllowUpdate     =   0   'False
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
            LCID            =   12298
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
            LCID            =   12298
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
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CONCEPTO DEL COMPROBANTE"
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
      TabIndex        =   24
      Top             =   735
      Width           =   8520
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CONCEPTO DEL COMPROBANTE"
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
      TabIndex        =   21
      Top             =   4095
      Width           =   1695
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " DIFERENCIA"
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
      Left            =   10920
      TabIndex        =   20
      Top             =   11445
      Width           =   1380
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
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   12285
      TabIndex        =   19
      Top             =   11445
      Width           =   1695
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   17535
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CostProy.frx":007F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CostProy.frx":0399
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CostProy.frx":06B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CostProy.frx":09CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CostProy.frx":0CE7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CostProy.frx":1001
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CostProy.frx":1C53
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CostProy.frx":1F6D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label LblHaber 
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
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   9030
      TabIndex        =   13
      Top             =   11445
      Width           =   1695
   End
   Begin VB.Label Label19 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " HABER"
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
      TabIndex        =   14
      Top             =   11445
      Width           =   960
   End
   Begin VB.Label LblDebe 
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
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   6300
      TabIndex        =   16
      Top             =   11445
      Width           =   1695
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " DEBE"
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
      TabIndex        =   15
      Top             =   11445
      Width           =   960
   End
End
Attribute VB_Name = "FCostosDelProyecto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CtasCostos As String

Private Sub CheqCta_Click()
If CheqCta.value = 1 Then
   LstCta.Visible = True
 Else
   LstCta.Visible = False
 End If
End Sub

Private Sub CheqIndiv_Click()
 If CheqIndiv.value = 1 Then
    DCCtas.Visible = True
 Else
    DCCtas.Visible = False
 End If
 
End Sub

Public Sub SumatoriaAsiento()
    RatonReloj
    DGAsiento.Visible = False
    Debe = 0
    Haber = 0
    sSQL = "SELECT * " _
         & "FROM Asiento " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND T_No = " & Trans_No & " " _
         & "AND CodigoU = '" & CodigoUsuario & "' "
    Select_Adodc_Grid DGAsiento, AdoAsiento, sSQL
    With AdoAsiento.Recordset
     If .RecordCount > 0 Then
         Do While Not .EOF
            Debe = Debe + .Fields("DEBE")
            Haber = Haber + .Fields("HABER")
           .MoveNext
         Loop
        .MoveFirst
     End If
    End With
    LblDebe.Caption = Format(Debe, "#,##0.00")
    LblHaber.Caption = Format(Haber, "#,##0.00")
    LblDiferencia.Caption = Format(Debe - Haber, "#,##0.00")
    DGAsiento.Visible = True
    RatonNormal
End Sub

Public Sub Grabar_Comprobante_Costos()
    SumatoriaAsiento
    If Debe = Haber Then
       FechaTexto = MBoxFechaF
       FechaComp = FechaTexto
       NumComp = ReadSetDataNum("Diario", True, False)
       Mensajes = "Esta seguro de Grabar el Comprobante No. " & NumComp
       Titulo = "PREGUNTA DE GRABACION"
       If BoxMensaje = vbYes Then
          DGAsiento.Visible = False
          DGAsientoSC.Visible = False
           
          NumComp = ReadSetDataNum("Diario", True, True)
          DiarioCaja = NumComp
         'Grabacion del Comprobante
          Co.T = Normal
          Co.TP = CompDiario
          Co.Fecha = FechaTexto
          Co.Numero = NumComp
          Co.Concepto = TxtConcepto.Text
          Co.CodigoB = Ninguno
          Co.Efectivo = Debe
          Co.Monto_Total = Debe
          Co.T_No = Trans_No
          Co.Usuario = CodigoUsuario
          Co.Item = NumEmpresa
            
          Grabar_Comprobante Co
          Control_Procesos Normal, Co.Concepto
          ImprimirComprobantesDe False, Co
               
          IniciarAsientosDe DGAsiento, AdoAsiento
          
          sSQL = "SELECT * " _
               & "FROM Asiento_SC " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND T_No = " & Trans_No & " " _
               & "AND CodigoU = '" & CodigoUsuario & "' "
          Select_Adodc_Grid DGAsientoSC, AdoAsientoSC, sSQL
          DGAsiento.Visible = True
          DGAsientoSC.Visible = True
        End If
    Else
        MsgBox "No se puede grabar el comprobante, las Transacciones no cuadran"
    End If
End Sub

Public Sub Enviar_Excel()
    DGAsiento.Visible = False
    DGAsientoSC.Visible = False
    GenerarDataTexto FCostosDelProyecto, AdoAsiento
    GenerarDataTexto FCostosDelProyecto, AdoAsientoSC
    DGAsiento.Visible = True
    DGAsientoSC.Visible = True
End Sub

Public Sub Imprimir()
  DGAsientoSC.Visible = False
  SQLMsg2 = "Desde:  " & MBoxFechaI.Text & "   al   " & MBoxFechaF.Text
  SQLMsg3 = "FACTURAS PENDIENTES"
  Imprimir_Saldos_SubCtas_Costos AdoAsientoSC, CheqCta
  DGAsientoSC.Visible = True
End Sub

Public Sub Costos_Proyecto()
Dim ContCtaCosto As Integer
Dim ContSC As Long
Dim SubModuloGasto As String

    RatonReloj
    DGAsiento.Visible = False
    DGAsientoSC.Visible = False
    Trans_No = 10
    Ln_No = 0
    ContSC = 0
    ContCtaCosto = 0
    Mifecha = MBoxFechaF
    Eliminar_Asientos_SP True
    Cadena = SinEspaciosIzq(DCCtaProyecto)
    TxtConcepto.Text = "Costeo del proyecto " & TrimStrg(MidStrg(DCCtaProyecto.Text, Len(Cadena) + 1, Len(DCCtaProyecto.Text))) & "; Subcuentas "
    For I = 0 To LstCtaProy.ListCount - 1
        If LstCtaProy.Selected(I) Then
           Cta = TrimStrg(MidStrg(LstCtaProy.List(I), 1, 18))
           InsertarAsientos AdoAsiento, Cta, 0, 0.001, 0
           TxtConcepto.Text = TxtConcepto.Text & TrimStrg(MidStrg(LstCtaProy.List(I), 19, Len(LstCtaProy.List(I)))) & ", "
           ContCtaCosto = ContCtaCosto + 1
        End If
    Next I
    CtasCostos = ""
    For I = 0 To LstCta.ListCount - 1
        If LstCta.Selected(I) Then CtasCostos = CtasCostos & "'" & TrimStrg(MidStrg(LstCta.List(I), 1, 18)) & "', "
    Next I
    If CtasCostos = "" Then CtasCostos = "'.'" Else CtasCostos = MidStrg(CtasCostos, 1, Len(CtasCostos) - 2)
    
    TxtConcepto.Text = TxtConcepto.Text & " del " & MBoxFechaI & " al " & MBoxFechaF
    If ContCtaCosto = 0 Then ContCtaCosto = 1
    
    FechaValida MBoxFechaI
    FechaValida MBoxFechaF
    FechaInicial = BuscarFecha(MBoxFechaI)
    FechaFinal = BuscarFecha(MBoxFechaF)
    CodigoCli = Ninguno
    With AdoSubCta.Recordset
     If .RecordCount > 0 Then
        .MoveFirst
        .Find ("Cliente = '" & DCCtas.Text & "' ")
         If Not .EOF Then CodigoCli = .Fields("Codigo")
     End If
    End With
    If Beneficiario = "" Then Beneficiario = Ninguno
    
    Total = 0
    Debe = 0
    Haber = 0
    
    If CheqCta.value = 1 Then
       sSQL = "SELECT TS.Cta,CC.Cuenta,CS.Detalle As Sub_Modulos,TS.Codigo,"
    Else
       sSQL = "SELECT CS.Detalle As Sub_Modulos,TS.Cta,CC.Cuenta,TS.Codigo,"
    End If
    sSQL = sSQL _
         & "SUM(TS.Debitos-TS.Creditos) As Total " _
         & "FROM Catalogo_SubCtas As CS, Catalogo_Cuentas As CC, Trans_SubCtas As TS " _
         & "WHERE TS.Item = '" & NumEmpresa & "' " _
         & "AND TS.Periodo = '" & Periodo_Contable & "' " _
         & "AND TS.Fecha BETWEEN #" & FechaInicial & "# AND #" & FechaFinal & "# " _
         & "AND TS.TC IN ('G','CC') " _
         & "AND TS.Cta LIKE '1%' "
    If CheqCta.value = 1 Then sSQL = sSQL & "AND CC.Codigo IN (" & CtasCostos & ") "
    If CheqIndiv.value = 1 Then sSQL = sSQL & "AND TS.Codigo = '" & CodigoCli & "' "
    sSQL = sSQL _
         & "AND TS.Item = CS.Item " _
         & "AND TS.Item = CC.Item " _
         & "AND TS.Periodo = CS.Periodo " _
         & "AND TS.Periodo = CC.Periodo " _
         & "AND TS.Cta = CC.Codigo " _
         & "AND TS.Codigo = CS.Codigo " _
         & "GROUP BY TS.Cta, CC.Cuenta, CS.Detalle, TS.Codigo " _
         & "HAVING SUM(TS.Debitos-TS.Creditos) <> 0 "
    If CheqCta.value = 1 Then sSQL = sSQL & "ORDER BY TS.Cta, CC.Cuenta, CS.Detalle " Else sSQL = sSQL & "ORDER BY CS.Detalle, TS.Cta, CC.Cuenta "
    Select_Adodc AdoAux, sSQL
    With AdoAux.Recordset
     If .RecordCount > 0 Then
         ValorDH = 0
         Cta = .Fields("Cta")
         Do While Not .EOF
            If Cta <> .Fields("Cta") Then
               InsertarAsientos AdoAsiento, Cta, 0, 0, ValorDH
               ValorDH = 0
               Cta = .Fields("Cta")
            End If
            Total = Total + .Fields("Total")
            ValorDH = ValorDH + .Fields("Total")
            SubTotal = .Fields("Total")
            SubModuloGasto = .Fields("Codigo")
            sSQL = "SELECT TC,Detalle " _
                 & "FROM Catalogo_SubCtas " _
                 & "WHERE Periodo = '" & Periodo_Contable & "' " _
                 & "AND Item = '" & NumEmpresa & "' " _
                 & "AND Codigo = '" & SubModuloGasto & "' "
            Select_Adodc AdoAux1, sSQL
           'MsgBox AdoAux.Recordset.RecordCount
            If AdoAux1.Recordset.RecordCount > 0 Then
               SetAdoAddNew "Asiento_SC"
               SetAdoFields "Codigo", SubModuloGasto
               SetAdoFields "TC", AdoAux1.Recordset.Fields("TC")
               SetAdoFields "Cta", .Fields("Cta")
               SetAdoFields "Beneficiario", AdoAux1.Recordset.Fields("Detalle")
               SetAdoFields "TM", "1"
               SetAdoFields "DH", "2"
               SetAdoFields "Valor", SubTotal
               SetAdoFields "FECHA_V", Mifecha
               SetAdoFields "T_No", Trans_No
               SetAdoFields "SC_No", ContSC
               SetAdoFields "Item", NumEmpresa
               SetAdoFields "CodigoU", CodigoUsuario
               SetAdoUpdate
               ContSC = ContSC + 1
            End If
           .MoveNext
         Loop
         InsertarAsientos AdoAsiento, Cta, 0, 0, ValorDH
     End If
    End With
    
    ValorDH = Redondear(Total / ContCtaCosto, 2)
    
    sSQL = "UPDATE Asiento " _
         & "SET DEBE = " & ValorDH & " " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND T_No = " & Trans_No & " " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND DEBE = 0 " _
         & "AND HABER = 0 "
    Ejecutar_SQL_SP sSQL
    
    SumatoriaAsiento
        
    sSQL = "SELECT * " _
         & "FROM Asiento_SC " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND T_No = " & Trans_No & " " _
         & "AND CodigoU = '" & CodigoUsuario & "' "
    Select_Adodc_Grid DGAsientoSC, AdoAsientoSC, sSQL
    DGAsiento.Visible = True
    DGAsientoSC.Visible = True
    Opcion = 10
    RatonNormal
End Sub

Public Sub UpDate_Costos_Proyecto()
Dim ContCtaCosto As Integer
    RatonReloj
    FechaValida MBoxFechaI
    FechaValida MBoxFechaF
    FechaInicial = BuscarFecha(MBoxFechaI)
    FechaFinal = BuscarFecha(MBoxFechaF)
    DGAsiento.Visible = False
    DGAsientoSC.Visible = False
    
    Trans_No = 10
    Ln_No = 0
    ContCtaCosto = 0
    Mifecha = MBoxFechaF
    
    Total = 0
    Debe = 0
    Haber = 0
    
    sSQL = "DELETE * " _
         & "FROM Asiento " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND T_No = " & Trans_No & " " _
         & "AND CodigoU = '" & CodigoUsuario & "' "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "SELECT * " _
         & "FROM Asiento " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND T_No = " & Trans_No & " " _
         & "AND CodigoU = '" & CodigoUsuario & "' "
    Select_Adodc_Grid DGAsiento, AdoAsiento, sSQL
    
    For I = 0 To LstCtaProy.ListCount - 1
     If LstCtaProy.Selected(I) Then
        Cta = TrimStrg(MidStrg(LstCtaProy.List(I), 1, 18))
        InsertarAsientos AdoAsiento, Cta, 0, 0.001, 0
        ContCtaCosto = ContCtaCosto + 1
     End If
    Next I
    
    sSQL = "SELECT * " _
         & "FROM Asiento_SC " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND T_No = " & Trans_No & " " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "ORDER BY Cta "
    Select_Adodc AdoAux, sSQL
    With AdoAux.Recordset
     If .RecordCount > 0 Then
         ValorDH = 0
         Cta = .Fields("Cta")
         Do While Not .EOF
            If Cta <> .Fields("Cta") Then
               InsertarAsientos AdoAsiento, Cta, 0, 0, ValorDH
               ValorDH = 0
               Cta = .Fields("Cta")
            End If
            Total = Total + .Fields("Valor")
            ValorDH = ValorDH + .Fields("Valor")
           .MoveNext
         Loop
         InsertarAsientos AdoAsiento, Cta, 0, 0, ValorDH
     End If
    End With
    
    ValorDH = Redondear(Total / ContCtaCosto, 2)
    
    sSQL = "UPDATE Asiento " _
         & "SET DEBE = " & ValorDH & " " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND T_No = " & Trans_No & " " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND DEBE = 0 " _
         & "AND HABER = 0 "
    Ejecutar_SQL_SP sSQL
    
    SumatoriaAsiento
        
    sSQL = "SELECT * " _
         & "FROM Asiento_SC " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND T_No = " & Trans_No & " " _
         & "AND CodigoU = '" & CodigoUsuario & "' "
    Select_Adodc_Grid DGAsientoSC, AdoAsientoSC, sSQL
    DGAsiento.Visible = True
    DGAsientoSC.Visible = True
    Opcion = 10
    RatonNormal
End Sub

Private Sub Command1_Click()
  Unload FCostosDelProyecto
End Sub

Private Sub DCCtaProyecto_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCCtaProyecto_LostFocus()
    LstCtaProy.Clear
    Contra_Cta = SinEspaciosIzq(DCCtaProyecto.Text)
    If Contra_Cta = "" Then Contra_Cta = Ninguno
    
    sSQL = "SELECT Codigo, Cuenta " _
         & "FROM Catalogo_Cuentas " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND DG = 'D' " _
         & "AND Codigo LIKE '" & Contra_Cta & "%' " _
         & "ORDER BY Codigo "
    Select_Adodc AdoAux, sSQL
    With AdoAux.Recordset
     If .RecordCount > 0 Then
         Do While Not .EOF
            LstCtaProy.AddItem .Fields("Codigo") & Space(19 - Len(.Fields("Codigo"))) & .Fields("Cuenta")
           .MoveNext
         Loop
     End If
    End With
End Sub

Private Sub DGAsiento_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then
     SumatoriaAsiento
     LstCtaProy.SetFocus
  End If
  
  If KeyCode = vbKeyReturn Then
     If AdoAsiento.Recordset.RecordCount > 0 Then
        AdoAsiento.Recordset.MoveNext
        If AdoAsiento.Recordset.EOF Then AdoAsiento.Recordset.MoveFirst
     End If
  End If
End Sub

Private Sub DGAsientoSC_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyDelete Then
     ID_Reg = -1
     With AdoAsientoSC.Recordset
      If .RecordCount > 0 Then
          ID_Reg = .Fields("SC_No")
          Titulo = "PREGUNTA DE ELIMINACION"
          Mensajes = "Esta seguro de Eliminar" & vbCrLf _
                   & .Fields("Beneficiario") & vbCrLf _
                   & "Por USD " & .Fields("Valor") & vbCrLf _
                   & "Linea No. " & .Fields("SC_No") & vbCrLf
          If BoxMensaje = vbYes Then
             sSQL = "DELETE * " _
                  & "FROM Asiento_SC " _
                  & "WHERE Item = '" & NumEmpresa & "' " _
                  & "AND T_No = " & Trans_No & " " _
                  & "AND CodigoU = '" & CodigoUsuario & "' " _
                  & "AND SC_No = " & ID_Reg & " "
             Ejecutar_SQL_SP sSQL
             
             sSQL = "SELECT * " _
                  & "FROM Asiento_SC " _
                  & "WHERE Item = '" & NumEmpresa & "' " _
                  & "AND T_No = " & Trans_No & " " _
                  & "AND CodigoU = '" & CodigoUsuario & "' "
             Select_Adodc_Grid DGAsientoSC, AdoAsientoSC, sSQL
          End If
      End If
     End With
  End If
End Sub

Private Sub Form_Activate()
  DGAsiento.Caption = "CONTABILIZCION"
  DGAsientoSC.Caption = "DETALLE DE SUBMODULOS"
  
  IniciarAsientosDe DGAsiento, AdoAsiento
  
  sSQL = "SELECT * " _
       & "FROM Asiento_SC " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  Select_Adodc_Grid DGAsientoSC, AdoAsientoSC, sSQL
  
  DGAsiento.width = MDI_X_Max - 100
  DGAsientoSC.width = MDI_X_Max - 100
  DGAsiento.Height = (MDI_Y_Max / 2) - 3500
  DGAsientoSC.Height = (MDI_Y_Max / 2)
  DGAsientoSC.Top = DGAsiento.Top + DGAsiento.Height + 30
  AdoAsiento.Top = DGAsientoSC.Top + DGAsientoSC.Height + 30
  PictFactura.width = MDI_X_Max - PictFactura.Left - 50
  PictFactura.Top = AdoAsiento.Top
  Label3.Top = AdoAsiento.Top
  Label4.Top = AdoAsiento.Top
  Label19.Top = AdoAsiento.Top
  LblDebe.Top = AdoAsiento.Top
  LblHaber.Top = AdoAsiento.Top
  LblDiferencia.Top = AdoAsiento.Top
  If Bloquear_Control Then
     Toolbar1.buttons("Costos").Enabled = False
     Toolbar1.buttons("Imprimir").Enabled = False
  End If
  Listar_SubModulos_Proyecto
  
  DGAsiento.Visible = True
  DGAsientoSC.Visible = True
  RatonNormal
End Sub

Private Sub Form_Load()
  DGAsiento.Visible = False
  DGAsientoSC.Visible = False
  ConectarAdodc AdoAux
  ConectarAdodc AdoAux1
  ConectarAdodc AdoDet
  ConectarAdodc AdoCtas
  ConectarAdodc AdoSubCta
  ConectarAdodc AdoAsiento
  ConectarAdodc AdoAsientoSC
  ConectarAdodc AdoCtaProyecto
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

Public Sub Listar_SubModulos_Proyecto()
  RatonReloj
  TipoCta = Ninguno
  SQLMsg1 = "SALDO DE COSTOS DEL PROYECTO"
  CheqIndiv.Caption = " Submodulo:"
  DGAsientoSC.Caption = SQLMsg1
  
  sSQL = "SELECT (Codigo & '  ' & Cuenta) As Cta_Proyecto " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND DG = 'G' " _
       & "AND Codigo LIKE '1%' " _
       & "AND LEN(Codigo) >= 10 " _
       & "ORDER BY Codigo "
  SelectDB_Combo DCCtaProyecto, AdoCtaProyecto, sSQL, "Cta_Proyecto"
  
  sSQL = "SELECT C.Detalle As Cliente,TS.Codigo " _
       & "FROM Trans_SubCtas As TS,Catalogo_SubCtas As C " _
       & "WHERE TS.Item = '" & NumEmpresa & "' " _
       & "AND TS.Periodo = '" & Periodo_Contable & "' " _
       & "AND TS.TC IN ('G','CC') " _
       & "AND TS.Cta LIKE '1%' " _
       & "AND TS.Codigo = C.Codigo " _
       & "AND TS.TC = C.TC " _
       & "AND TS.Item = C.Item " _
       & "AND TS.Periodo = C.Periodo " _
       & "GROUP BY C.Detalle, TS.Codigo " _
       & "ORDER BY C.Detalle, TS.Codigo "
  SelectDB_Combo DCCtas, AdoSubCta, sSQL, "Cliente"
  
  sSQL = "SELECT TS.Cta, CC.Cuenta " _
       & "FROM Trans_SubCtas As TS, Catalogo_Cuentas As CC " _
       & "WHERE CC.Item = '" & NumEmpresa & "' " _
       & "AND CC.Periodo = '" & Periodo_Contable & "' " _
       & "AND CC.TC IN ('G','CC') " _
       & "AND TS.Cta LIKE '1%' " _
       & "AND CC.Codigo = TS.Cta " _
       & "AND CC.Item = TS.Item " _
       & "AND CC.Periodo = TS.Periodo " _
       & "AND CC.TC = TS.TC " _
       & "GROUP BY TS.Cta, CC.Cuenta " _
       & "ORDER BY TS.Cta "
  Select_Adodc AdoCtas, sSQL
  LstCta.Clear
  With AdoCtas.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          LstCta.AddItem .Fields("Cta") & Space(19 - Len(.Fields("Cta"))) & .Fields("Cuenta")
         .MoveNext
       Loop
   End If
  End With
  RatonNormal
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
   'MsgBox Button.key
    Select Case Button.key
      Case "Costos":       Costos_Proyecto
      Case "UpdateCostos": UpDate_Costos_Proyecto
      Case "Excel":        Enviar_Excel
      Case "Grabar":       Grabar_Comprobante_Costos
      Case "Salir":        Unload FCostosDelProyecto
    End Select
End Sub

Private Sub TxtConcepto_GotFocus()
  MarcarTexto TxtConcepto
End Sub
