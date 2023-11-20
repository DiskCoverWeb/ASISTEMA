VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form MayorAux 
   Caption         =   "Movimientos de Sub Cuentas"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7395
   ScaleWidth      =   11280
   WindowState     =   2  'Maximized
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
            Picture         =   "MayorAux.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MayorAux.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MayorAux.frx":11B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MayorAux.frx":14CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MayorAux.frx":1DA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MayorAux.frx":455A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo DCAgencia 
      Bindings        =   "MayorAux.frx":635C
      DataSource      =   "AdoAgencias"
      Height          =   345
      Left            =   3885
      TabIndex        =   29
      Top             =   1050
      Width           =   6420
      _ExtentX        =   11324
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
   Begin MSDataListLib.DataCombo DCUsuario 
      Bindings        =   "MayorAux.frx":6376
      DataSource      =   "AdoUsuario"
      Height          =   345
      Left            =   3885
      TabIndex        =   28
      Top             =   735
      Width           =   6420
      _ExtentX        =   11324
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
   Begin VB.CheckBox CheckAgencia 
      Caption         =   "Agencia:"
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
      Left            =   2520
      TabIndex        =   31
      Top             =   1050
      Width           =   1275
   End
   Begin VB.CheckBox CheckUsuario 
      Caption         =   "Por &Usuario"
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
      Left            =   2520
      TabIndex        =   30
      Top             =   735
      Width           =   1380
   End
   Begin VB.Frame Frame2 
      Caption         =   "Estado"
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
      Left            =   10395
      TabIndex        =   24
      Top             =   735
      Width           =   1380
      Begin VB.OptionButton OpcT 
         Caption         =   "Todos"
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
         TabIndex        =   27
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton OpcA 
         Caption         =   "Anulados"
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
         TabIndex        =   26
         Top             =   525
         Width           =   1170
      End
      Begin VB.OptionButton OpcN 
         Caption         =   "Normal"
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
         TabIndex        =   25
         Top             =   210
         Value           =   -1  'True
         Width           =   1065
      End
   End
   Begin MSDataListLib.DataList DLCtas 
      Bindings        =   "MayorAux.frx":638F
      DataSource      =   "AdoSubCta"
      Height          =   1635
      Left            =   2415
      TabIndex        =   16
      Top             =   1785
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   2884
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
   Begin MSDataGridLib.DataGrid DGMayor 
      Bindings        =   "MayorAux.frx":63A7
      Height          =   2220
      Left            =   105
      TabIndex        =   17
      ToolTipText     =   "<CTRL+ F> Cambia fecha de Vencimiento"
      Top             =   3570
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   3916
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      BorderStyle     =   0
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
      Caption         =   "S U B M A Y O R"
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
   Begin MSDataListLib.DataCombo DCCtas 
      Bindings        =   "MayorAux.frx":63C0
      DataSource      =   "AdoCtas"
      Height          =   345
      Left            =   2415
      TabIndex        =   15
      Top             =   1470
      Visible         =   0   'False
      Width           =   7890
      _ExtentX        =   13917
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
   Begin VB.OptionButton OpcCta 
      Caption         =   "Una sola Cuenta"
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
      Left            =   10395
      TabIndex        =   14
      Top             =   2100
      Value           =   -1  'True
      Width           =   2010
   End
   Begin VB.OptionButton OpcTodas 
      Caption         =   "Todas las Cuentas"
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
      Left            =   10395
      TabIndex        =   13
      Top             =   2415
      Width           =   2010
   End
   Begin MSAdodcLib.Adodc AdoSubCta 
      Height          =   330
      Left            =   315
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
   Begin MSMask.MaskEdBox MBFechaF 
      Height          =   330
      Left            =   1155
      TabIndex        =   3
      Top             =   1050
      Width           =   1170
      _ExtentX        =   2064
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
   Begin MSMask.MaskEdBox MBFechaI 
      Height          =   330
      Left            =   1155
      TabIndex        =   1
      Top             =   735
      Width           =   1170
      _ExtentX        =   2064
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
   Begin VB.Frame Frame1 
      Caption         =   "SubCuenta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   105
      TabIndex        =   4
      Top             =   1470
      Width           =   2220
      Begin VB.OptionButton OpcCC 
         Caption         =   "&C.de C."
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
         Left            =   1155
         TabIndex        =   10
         Top             =   840
         Width           =   960
      End
      Begin VB.OptionButton OpcPM 
         Caption         =   "Pri&ma"
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
         TabIndex        =   7
         Top             =   840
         Width           =   960
      End
      Begin VB.OptionButton OpcI 
         Caption         =   "&Ingreso"
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
         Left            =   1155
         TabIndex        =   8
         Top             =   210
         Width           =   960
      End
      Begin VB.OptionButton OpcG 
         Caption         =   "&Gastos"
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
         Left            =   1155
         TabIndex        =   9
         Top             =   525
         Width           =   960
      End
      Begin VB.OptionButton OpcC 
         Caption         =   "Cx&C"
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
         TabIndex        =   5
         Top             =   210
         Value           =   -1  'True
         Width           =   750
      End
      Begin VB.OptionButton OpcP 
         Caption         =   "Cx&P"
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
         TabIndex        =   6
         Top             =   525
         Width           =   750
      End
   End
   Begin MSAdodcLib.Adodc AdoSubCta1 
      Height          =   330
      Left            =   105
      Top             =   6930
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
      CommandType     =   1
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
      Caption         =   "SubCta1"
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
      Left            =   315
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
   Begin MSAdodcLib.Adodc AdoAgencias 
      Height          =   330
      Left            =   315
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
      Caption         =   "Agencias"
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
   Begin MSAdodcLib.Adodc AdoUsuario 
      Height          =   330
      Left            =   315
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
      Caption         =   "Usuario"
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
      Left            =   315
      Top             =   5985
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   11280
      _ExtentX        =   19897
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
            Object.ToolTipText     =   "Salir del SubModulo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "UnSubmodulo"
            Object.ToolTipText     =   "Consulta una Submodulo"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "TodosSubmodulos"
            Object.ToolTipText     =   "Todos los submodulos de la Cuenta Contable"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Abonos_Anticipados"
            Object.ToolTipText     =   "Imprimir Comprbante de Ingreso de Abonos Anticipados"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Excel"
            Object.ToolTipText     =   "Exportar a Excel los resultados"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin VB.Label LabelTotSaldoAnt 
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
      Left            =   105
      TabIndex        =   12
      Top             =   3045
      Width           =   2220
   End
   Begin VB.Label LabelTotSaldo 
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
      Left            =   10080
      TabIndex        =   18
      Top             =   6930
      Width           =   1695
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Saldo Actual"
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
      TabIndex        =   19
      Top             =   6930
      Width           =   1380
   End
   Begin VB.Label LabelTotHaber 
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
      Left            =   7035
      TabIndex        =   20
      Top             =   6930
      Width           =   1695
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Creditos:"
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
      Left            =   6090
      TabIndex        =   21
      Top             =   6930
      Width           =   960
   End
   Begin VB.Label LabelTotDebe 
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
      Left            =   4410
      TabIndex        =   22
      Top             =   6930
      Width           =   1695
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Debitos:"
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
      Left            =   3465
      TabIndex        =   23
      Top             =   6930
      Width           =   960
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Saldo Anterior"
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
      Top             =   2730
      Width           =   2220
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Hasta"
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
      TabIndex        =   2
      Top             =   1050
      Width           =   1065
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Desde"
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
      Top             =   735
      Width           =   1065
   End
End
Attribute VB_Name = "MayorAux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Ctas_SubMod(TipoM As String)
  sSQL = "SELECT Codigo & '    ' & Cuenta As Nombre_Cta " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE TC = '" & TipoM & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo "
  SelectDB_Combo DCCtas, AdoCtas, sSQL, "Nombre_Cta"
 'MsgBox "."
End Sub

Public Sub Tipo_Modulo(TipoM As String)
  Codigo = SinEspaciosIzq(DCCtas.Text)
  If Codigo = "" Then Codigo = Ninguno
  Select Case TipoM
    Case "C", "P"
         sSQL = "SELECT C.Cliente As Nombre_Cta,C.Codigo " _
              & "FROM Catalogo_CxCxP As CP,Clientes As C " _
              & "WHERE CP.TC = '" & TipoM & "' " _
              & "AND CP.Item = '" & NumEmpresa & "' " _
              & "AND CP.Periodo = '" & Periodo_Contable & "' " _
              & "AND CP.Cta = '" & Codigo & "' " _
              & "AND CP.Codigo = C.Codigo " _
              & "GROUP BY C.Cliente,C.Codigo "
    Case "I", "G", "PM", "CC"
         sSQL = "SELECT Detalle As Nombre_Cta,Codigo " _
              & "FROM Catalogo_SubCtas " _
              & "WHERE TC = '" & TipoM & "' " _
              & "AND Item = '" & NumEmpresa & "' " _
              & "AND Periodo = '" & Periodo_Contable & "' " _
              & "ORDER BY Detalle,Codigo "
  End Select
  SelectDB_List DLCtas, AdoSubCta, sSQL, "Nombre_Cta"
End Sub

Public Sub Consultar_Submodulo(Individual As Boolean)
     If Individual Then
        If AdoSubCta.Recordset.RecordCount > 0 Then
           AdoSubCta.Recordset.MoveFirst
           AdoSubCta.Recordset.Find ("Nombre_Cta = '" & Codigo1 & "'")
           If Not AdoSubCta.Recordset.EOF Then Codigo1 = AdoSubCta.Recordset.fields("Codigo")
        End If
     End If
    'Consultamos el SubModulo
     If Codigo1 = "" Then Codigo1 = Ninguno
     Select Case TipoDoc
       Case "C", "P"
            sSQL = "SELECT TSC.Cta,TSC.Fecha,TSC.TP,TSC.Numero,C.Cliente,Concepto,Debitos,Creditos,"
       Case Else
            sSQL = "SELECT TSC.Cta,TSC.Fecha,TSC.TP,TSC.Numero,C.Detalle As Cliente,Concepto,Debitos,Creditos,"
     End Select
     If TipoDoc = "PM" Then
        sSQL = sSQL & "TSC.Saldo_MN,TSC.Prima,"
     Else
        sSQL = sSQL & "TSC.Saldo_MN,TSC.Factura,"
     End If
     Select Case TipoDoc
       Case "C", "P"
            sSQL = sSQL _
                 & "Parcial_ME,TSC.Detalle_SubCta,TSC.Fecha_V,TSC.Codigo,TSC.Item,TSC.ID " _
                 & "FROM Trans_SubCtas As TSC,Comprobantes As Co,Clientes As C "
       Case Else
            sSQL = sSQL _
                 & "Parcial_ME,TSC.Detalle_SubCta,TSC.Fecha_V,TSC.Codigo,TSC.Item,TSC.ID " _
                 & "FROM Trans_SubCtas As TSC,Comprobantes As Co,Catalogo_SubCtas As C "
     End Select
     sSQL = sSQL & "WHERE TSC.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
     If OpcCta.value Then sSQL = sSQL & "AND TSC.Cta = '" & Codigo & "' "
     If Individual Then sSQL = sSQL & "AND TSC.Codigo = '" & Codigo1 & "' "
     If OpcN.value Then
        sSQL = sSQL & "AND TSC.T = 'N' "
     ElseIf OpcA.value Then
        sSQL = sSQL & "AND TSC.T = 'A' "
     End If
     If CheckAgencia.value = 1 Then
        sSQL = sSQL & "AND TSC.Item = '" & SinEspaciosIzq(DCAgencia.Text) & "' "
     Else
        If Not ConSucursal Then sSQL = sSQL & "AND TSC.Item = '" & NumEmpresa & "' "
     End If
     If CheckUsuario.value = 1 Then sSQL = sSQL & "AND Co.CodigoU = '" & SinEspaciosDer(DCUsuario.Text) & "' "
     sSQL = sSQL _
          & "AND TSC.Periodo = '" & Periodo_Contable & "' " _
          & "AND TSC.TC = '" & TipoDoc & "' " _
          & "AND Co.TP = TSC.TP " _
          & "AND Co.Numero = TSC.Numero " _
          & "AND Co.Item = TSC.Item " _
          & "AND Co.Periodo = TSC.Periodo "
     Select Case TipoDoc
       Case "C", "P"
           'Nada
       Case Else
            sSQL = sSQL _
                 & "AND C.Item = TSC.Item " _
                 & "AND C.Periodo = TSC.Periodo "
     End Select
     sSQL = sSQL _
          & "AND TSC.Codigo = C.Codigo " _
          & "ORDER BY TSC.Codigo,TSC.Cta,TSC.Fecha,TSC.TP,TSC.Numero,Factura,Debitos DESC,Creditos,TSC.ID "
     Select_Adodc_Grid DGMayor, AdoSubCta1, sSQL
     DGMayor.Visible = False
     Debe = 0
     Haber = 0
     Saldo = 0
     SaldoAnterior = 0
     Cta = Ninguno
     
    'Calculamos Totales de SubModulos
     If Codigo1 = "" Then Codigo1 = Ninguno
     sSQL = "SELECT TSC.Cta,SUM(Debitos) As TDebitos,SUM(Creditos) As TCreditos "
     Select Case TipoDoc
       Case "C", "P"
            sSQL = sSQL & "FROM Trans_SubCtas As TSC,Comprobantes As Co,Clientes As C "
       Case Else
            sSQL = sSQL & "FROM Trans_SubCtas As TSC,Comprobantes As Co,Catalogo_SubCtas As C "
     End Select
     sSQL = sSQL & "WHERE TSC.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
     If OpcCta.value Then sSQL = sSQL & "AND TSC.Cta = '" & Codigo & "' "
     If Individual Then sSQL = sSQL & "AND TSC.Codigo = '" & Codigo1 & "' "
     If OpcN.value Then
        sSQL = sSQL & "AND TSC.T = 'N' "
     ElseIf OpcA.value Then
        sSQL = sSQL & "AND TSC.T = 'A' "
     End If
     If CheckAgencia.value = 1 Then
        sSQL = sSQL & "AND TSC.Item = '" & SinEspaciosIzq(DCAgencia.Text) & "' "
     Else
        If Not ConSucursal Then sSQL = sSQL & "AND TSC.Item = '" & NumEmpresa & "' "
     End If
     If CheckUsuario.value = 1 Then sSQL = sSQL & "AND Co.CodigoU = '" & SinEspaciosDer(DCUsuario.Text) & "' "
     sSQL = sSQL _
          & "AND TSC.Periodo = '" & Periodo_Contable & "' " _
          & "AND TSC.TC = '" & TipoDoc & "' " _
          & "AND Co.TP = TSC.TP " _
          & "AND Co.Numero = TSC.Numero " _
          & "AND Co.Item = TSC.Item " _
          & "AND Co.Periodo = TSC.Periodo "
     Select Case TipoDoc
       Case "C", "P"
           'Nada
       Case Else
            sSQL = sSQL _
                 & "AND C.Item = TSC.Item " _
                 & "AND C.Periodo = TSC.Periodo "
     End Select
     sSQL = sSQL _
          & "AND TSC.Codigo = C.Codigo " _
          & "GROUP BY TSC.Cta "
     Select_Adodc AdoAux, sSQL
     With AdoAux.Recordset
      If .RecordCount > 0 Then
          
          Do While Not .EOF
             Debe = Debe + .fields("TDebitos")
             Haber = Haber + .fields("TCreditos")
             Cta = .fields("Cta")
            .MoveNext
          Loop
          If AdoSubCta1.Recordset.RecordCount > 0 Then
             AdoSubCta1.Recordset.MoveLast
             Saldo = AdoSubCta1.Recordset.fields("Saldo_MN")
          End If
          SaldoAnterior = CalculosSaldoAnt(Cta, Debe, Haber, Saldo)
      End If
     End With
     
     LabelTotDebe.Caption = Format(Debe, "#,##0.00")
     LabelTotHaber.Caption = Format(Haber, "#,##0.00")
     LabelTotSaldo.Caption = Format(Saldo, "#,##0.00")
     LabelTotSaldoAnt.Caption = Format(SaldoAnterior, "#,##0.00")
     DGMayor.Visible = True
End Sub

Private Sub DCCtas_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCCtas_LostFocus()
  If OpcC.value Then Tipo_Modulo "C"
  If OpcP.value Then Tipo_Modulo "P"
  If OpcG.value Then Tipo_Modulo "G"
  If OpcI.value Then Tipo_Modulo "I"
  If OpcPM.value Then Tipo_Modulo "PM"
End Sub

Private Sub DGMayor_DblClick()
  If AdoSubCta1.Recordset.RecordCount > 0 Then
     Co.T = Normal
     Co.TP = DGMayor.Columns(2)
     Co.Numero = DGMayor.Columns(3)
     Co.Beneficiario = DGMayor.Columns(4)
     Co.Fecha = DGMayor.Columns(1)
     Co.CodigoB = DGMayor.Columns(13)
     Co.Item = DGMayor.Columns(14)
     Co.Concepto = DGMayor.Columns(5)
     Co.Efectivo = DGMayor.Columns(7)
  End If
End Sub

Private Sub DGMayor_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyF Then
     ID_Reg = Val(DGMayor.Columns(15))
     Mifecha = InputBox(vbCrLf & vbCrLf & vbCrLf & "INGRESE FECHA DE VENCIMIENTO:", "CAMBIO DE FECHA(" & ID_Reg & ")", DGMayor.Columns(12))
     If IsDate(Mifecha) Then
        sSQL = "UPDATE Trans_SubCtas " _
             & "SET Fecha_V = #" & BuscarFecha(Mifecha) & "# " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND ID = " & ID_Reg & " "
        Ejecutar_SQL_SP sSQL
     Else
        MsgBox "Fecha Ingresada: " & Mifecha & ", es invalida, vuelva a ingresar"
     End If
  End If
'''  If CtrlDown And KeyCode = vbKeyP Then Imprimir_Recibos_De_Pagos
End Sub

Private Sub Form_Activate()
  Ctas_SubMod "C"
  Tipo_Modulo "C"
  RatonNormal
  DCCtas.Visible = True
  DLCtas.Enabled = True
  sSQL = "SELECT (Nombre_Completo & '  ' & Codigo) As CodUsuario " _
       & "FROM Accesos " _
       & "WHERE Codigo <> '*' " _
       & "ORDER BY Nombre_Completo "
  SelectDB_Combo DCUsuario, AdoUsuario, sSQL, "CodUsuario", False
  
  If ConSucursal Then
     sSQL = "SELECT (Item & '  ' & Empresa) As NomEmpresa " _
          & "FROM Empresas " _
          & "WHERE Item IN (" & ListSucursales & ") " _
          & "ORDER BY Item,Empresa "
     SelectDB_Combo DCAgencia, AdoAgencias, sSQL, "NomEmpresa"
     CheckAgencia.value = 0
     DCAgencia.Visible = True
     CheckAgencia.Visible = True
  Else
     DCAgencia.Visible = False
     CheckAgencia.Visible = False
  End If
  
  MBFechaI.SetFocus
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoAux
  ConectarAdodc AdoCtas
  ConectarAdodc AdoSubCta
  ConectarAdodc AdoSubCta1
  ConectarAdodc AdoUsuario
  ConectarAdodc AdoAgencias
  
  DGMayor.Height = MDI_Y_Max - DGMayor.Top - 300
  DGMayor.width = MDI_X_Max - DGMayor.Left
  AdoSubCta1.Top = DGMayor.Top + DGMayor.Height + 10
  
  Label6.Top = DGMayor.Top + DGMayor.Height + 10
  Label9.Top = DGMayor.Top + DGMayor.Height + 10
  Label11.Top = DGMayor.Top + DGMayor.Height + 10
  LabelTotSaldo.Top = DGMayor.Top + DGMayor.Height + 10
  LabelTotDebe.Top = DGMayor.Top + DGMayor.Height + 10
  LabelTotHaber.Top = DGMayor.Top + DGMayor.Height + 10
End Sub

Private Sub MBFechaF_GotFocus()
  MarcarTexto MBFechaF
End Sub

Private Sub MBFechaF_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaF_LostFocus()
  FechaValida MBFechaF
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaI_LostFocus()
  FechaValida MBFechaI
End Sub

Private Sub OpcC_Click()
  Ctas_SubMod "C"
  Tipo_Modulo "C"
End Sub

Private Sub OpcCC_Click()
  Ctas_SubMod "CC"
  Tipo_Modulo "CC"
End Sub

Private Sub OpcCta_Click()
  DCCtas.Enabled = True
End Sub

Private Sub OpcG_Click()
  Ctas_SubMod "G"
  Tipo_Modulo "G"
End Sub

Private Sub OpcI_Click()
  Ctas_SubMod "I"
  Tipo_Modulo "I"
End Sub

Private Sub OpcP_Click()
  Ctas_SubMod "P"
  Tipo_Modulo "P"
End Sub

Private Sub OpcPM_Click()
  Ctas_SubMod "PM"
  Tipo_Modulo "PM"
End Sub

Private Sub OpcTodas_Click()
  DCCtas.Enabled = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'MsgBox Button.key
 If OpcC.value Then TipoDoc = "C"
 If OpcP.value Then TipoDoc = "P"
 If OpcI.value Then TipoDoc = "I"
 If OpcG.value Then TipoDoc = "G"
 If OpcPM.value Then TipoDoc = "PM"
 If OpcCC.value Then TipoDoc = "CC"
 FechaValida MBFechaI
 FechaValida MBFechaF
 FechaIni = BuscarFecha(MBFechaI.Text)
 FechaFin = BuscarFecha(MBFechaF.Text)
 Codigo1 = DLCtas.Text
 Codigo = SinEspaciosIzq(DCCtas.Text)
 
 Select Case Button.key
   Case "Salir"
        Unload MayorAux
   Case "UnSubmodulo"
        Consultar_Submodulo True
   Case "TodosSubmodulos"
        Consultar_Submodulo False
   Case "Imprimir"
        If OpcC.value Then
           Cadena = OpcC.Caption
        ElseIf OpcP.value Then
           Cadena = OpcP.Caption
        ElseIf OpcI.value Then
           Cadena = OpcI.Caption
        ElseIf OpcPM.value Then
           Cadena = OpcPM.Caption
        Else
           Cadena = OpcG.Caption
        End If
        FechaCorte = "Desde " & MBFechaI & " al " & MBFechaF
        DGMayor.Visible = False
        Imprimir_Mayor_Aux AdoSubCta1, Cadena
        DGMayor.Visible = True
   Case "Abonos_Anticipados"
        Control_Procesos "I", "Imprimio Comprobante de: " & Co.TP & ", No. " & NumComp
        ImprimirComprobantesDe False, Co
   Case "Excel"
        DGMayor.Visible = False
        Exportar_AdoDB_Excel AdoSubCta1.Recordset, "Mayores " & BuscarFecha(MBFechaI) & " al " & BuscarFecha(MBFechaF)
       'GenerarDataTexto MayorAux, AdoSubCta1
        DGMayor.Visible = True
 End Select
End Sub

'''Public Sub Imprimir_Recibos_De_Pagos()
'''Dim Ini_X As Single
'''Dim Ini_Y As Single
'''Dim tipoDeLetra As String
'''Dim PrintLinea As String
'''Dim PrintCar As String
'''   RatonReloj
'''   Contador = 1
'''
'''   With AdoSubCta1.Recordset
'''    If .RecordCount > 0 Then
'''        Numero = DGMayor.Columns(3)
'''       .MoveFirst
'''        TRecibo.Recibo_No = .Fields("Numero")
'''        tipoDeLetra = TipoTimes
'''
'''       'Generamos el documento
'''        tPrint.TipoImpresion = Es_PDF
'''        tPrint.NombreArchivo = "CD-" & Format$(TRecibo.Recibo_No, "00000000")
'''        tPrint.TituloArchivo = "Recibo de Pago"
'''        tPrint.TipoLetra = tipoDeLetra
'''        tPrint.OrientacionPagina = 1
'''        tPrint.PaginaA4 = True
'''        tPrint.EsCampoCorto = False
'''        tPrint.VerDocumento = True
'''
'''        Set cPrint = New cImpresion
'''        cPrint.iniciaImpresion
'''
'''        Do While Not .EOF
'''           If Numero = .Fields("Numero") And .Fields("Creditos") > 0 Then
'''              Ini_X = 1
'''              TRecibo.Tipo_Recibo = "I"
'''              TRecibo.Cobrado_a = .Fields("Cliente")
'''              TRecibo.Fecha = .Fields("Fecha")
'''              TRecibo.Total = .Fields("Creditos")
'''              TRecibo.SubTotal = .Fields("Creditos")
'''              TRecibo.IVA = 0
'''              TRecibo.Concepto = .Fields("Concepto") & ". " & vbCrLf _
'''                               & " " & vbCrLf _
'''                               & "Pago Libre y voluntariamente la pensión prorrateada del mes de Febrero del 2016; " _
'''                               & "y eximo de toda responsabilidad a la Institución y por ende a sus Autoridades. " & vbCrLf
'''              cPrint.porteDeLetra = 7
'''              cPrint.printCuadroLinea 2, 2, 2, 2, Negro, "B"
'''              cPrint.printImagen LogoTipo, Ini_X + 0.2, 1, 3, 1.5
'''              PosLinea = 1.1
'''              cPrint.colorDeLetra = Negro
'''              cPrint.printEncabezado 1.5, 1.1, TipoTimes
'''
'''              If UCaseStrg(Direccion) <> UCaseStrg(DireccionEstab) Then
'''                 cPrint.printTexto Ini_X + 1.5, PosLinea, "Sucursal: " & DireccionEstab
'''                 PosLinea = PosLinea + 0.3
'''              End If
'''              cPrint.printTexto Ini_X + 1.5, PosLinea, "Telefono(s): " & Telefono1 & " / " & Telefono2 & " / " & FAX
'''              PosLinea = PosLinea + 0.3
'''              cPrint.printTexto Ini_X + 1.5, PosLinea, ULCase(NombreCiudad) & " - Ecuador"
'''              PosLinea = PosLinea + 0.3
'''              cPrint.printTexto Ini_X + 1.5, PosLinea, "R.U.C.: " & RUC
'''              PosLinea = 2.8
'''              cPrint.porteDeLetra = 9
'''              If TRecibo.Tipo_Recibo = "I" Then
'''                 cPrint.printTexto Ini_X + 0.8, PosLinea, "COMPROBANTE DE INGRESO No."
'''              Else
'''                 cPrint.printTexto Ini_X + 0.8, PosLinea, "COMPROBANTE DE EGRESO No."
'''              End If
'''              cPrint.printTexto Ini_X + 5.9, PosLinea, Format$(Year(TRecibo.Fecha), "0000") & "-" & Format$(TRecibo.Recibo_No, "00000000") & "-" & Format$(Contador, "000")
'''              Contador = Contador + 1
'''              PosLinea = PosLinea + 0.5
'''              cPrint.porteDeLetra = 9
'''              cPrint.printCuadroLinea Ini_X - 0.1, 3.3, Ini_X + 9, 3.7, Blanco, "BF"
'''              cPrint.printCuadroLinea Ini_X - 0.1, 3.3, Ini_X + 5.9, 3.7, Negro, "B"
'''              cPrint.printCuadroLinea Ini_X + 5.9, 3.3, Ini_X + 9, 3.7, Negro, "B"
'''              cPrint.printCuadroLinea Ini_X - 0.1, 4.1, Ini_X + 9, 4.5, Blanco, "BF"
'''              cPrint.printCuadroLinea Ini_X - 0.1, 4.1, Ini_X + 9, 4.5, Negro, "B"
'''              cPrint.printTexto Ini_X, PosLinea, "Fecha: " & FechaStrg(TRecibo.Fecha)
'''              cPrint.printTexto Ini_X + 6, PosLinea, "Por USD"
'''              cPrint.printVariable Ini_X + 5.85, PosLinea - 0.05, TRecibo.Total
'''              PosLinea = PosLinea + 0.35
'''              'cPrint.porteDeLetra = 7
'''              cPrint.printTexto Ini_X, PosLinea, "Beneficiario:"
'''              cPrint.printTexto Ini_X + 1.7, PosLinea, TRecibo.Cobrado_a
'''              PosLinea = PosLinea + 0.4
'''              cPrint.printTexto Ini_X, PosLinea, "La suma de: " & Cambio_Letras(TRecibo.Total, 2)
'''              PosLinea = 4.5
'''              cPrint.printTexto Ini_X, PosLinea, "POR CONCEPTO DE:"
'''              PosLinea = PosLinea + 0.3
'''              PrintLinea = TRecibo.Concepto
'''             'MsgBox PrintLinea
'''              I = 1
'''              PrintCar = ""
'''              While PrintLinea <> "" And I < Len(PrintLinea)
'''                  If MidStrg(PrintLinea, I, 1) <> vbCr Then
'''                     PrintCar = PrintCar & MidStrg(PrintLinea, I, 1)
'''                  Else
'''                     'MsgBox PrintCar
'''                     cPrint.printTexto Ini_X, PosLinea, TrimStrg(PrintCar), , 8.5
'''                     PosLinea = PosLinea + 0.35
'''                     PrintLinea = TrimStrg(MidStrg(PrintLinea, I + 1, Len(PrintLinea)))
'''                    'MsgBox PrintLinea
'''                     PrintCar = ""
'''                     I = 0
'''                  End If
'''                  I = I + 1
'''              Wend
'''              'cPrint.printTexto Ini_X, PosLinea, TRecibo.Concepto, , 9
'''              cPrint.porteDeLetra = 7
'''              PosLinea = PosLinea + 0.3
'''              cPrint.printCuadroLinea Ini_X - 0.1, 3.3, Ini_X + 9, PosLinea, Negro, "B"
'''              cPrint.printCuadroLinea Ini_X, PosLinea + 0.4, Ini_X + 2.85, PosLinea + 0.4, Negro
'''              cPrint.printCuadroLinea Ini_X + 3.25, PosLinea + 0.4, Ini_X + 6, PosLinea + 0.4, Negro
'''              cPrint.printTexto Ini_X, PosLinea + 0.7, "ENTREGUE CONFORME"
'''              cPrint.printTexto Ini_X + 3.25, PosLinea + 0.7, "RECIBI CONFORME"
'''
'''              cPrint.paginaNueva
'''           End If
'''          .MoveNext
'''        Loop
'''        cPrint.finalizaImpresion
'''    End If
'''   End With
'''
'''   RatonNormal
'''End Sub
