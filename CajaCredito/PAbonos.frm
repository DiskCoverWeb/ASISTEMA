VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form Precancelacion 
   Caption         =   "CANCELACION DE CREDITOS"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12360
   DrawStyle       =   1  'Dash
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7275
   ScaleWidth      =   12360
   WindowState     =   2  'Maximized
   Begin MSDataListLib.DataCombo DCTipoPrestamo 
      Bindings        =   "PAbonos.frx":0000
      DataSource      =   "AdoTipoPrest"
      Height          =   315
      Left            =   3360
      TabIndex        =   7
      Top             =   945
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "DataCombo1"
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
   Begin MSDataGridLib.DataGrid DGAsiento 
      Bindings        =   "PAbonos.frx":001B
      Height          =   2010
      Left            =   105
      TabIndex        =   51
      Top             =   4725
      Width           =   12150
      _ExtentX        =   21431
      _ExtentY        =   3545
      _Version        =   393216
      AllowUpdate     =   0   'False
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
   Begin MSAdodcLib.Adodc AdoTipoPrest 
      Height          =   330
      Left            =   210
      Top             =   5565
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
      Caption         =   "TipoPrest"
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
   Begin VB.TextBox TextLinea 
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
      Left            =   8925
      MaxLength       =   49
      MultiLine       =   -1  'True
      TabIndex        =   48
      Top             =   3885
      Width           =   960
   End
   Begin VB.TextBox TextSaldo 
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
      Left            =   7455
      MaxLength       =   49
      MultiLine       =   -1  'True
      TabIndex        =   27
      Top             =   3885
      Width           =   1485
   End
   Begin VB.TextBox TextCapital 
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
      Left            =   5985
      MaxLength       =   49
      MultiLine       =   -1  'True
      TabIndex        =   25
      Top             =   3885
      Width           =   1485
   End
   Begin VB.TextBox TextComision 
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
      Left            =   4515
      MaxLength       =   49
      MultiLine       =   -1  'True
      TabIndex        =   50
      Top             =   3885
      Width           =   1485
   End
   Begin VB.TextBox TextInt 
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
      Left            =   3150
      MaxLength       =   49
      MultiLine       =   -1  'True
      TabIndex        =   23
      Top             =   3885
      Width           =   1380
   End
   Begin VB.TextBox TextMonto 
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
      Left            =   1680
      MaxLength       =   49
      MultiLine       =   -1  'True
      TabIndex        =   20
      Top             =   3885
      Width           =   1485
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   105
      TabIndex        =   0
      Top             =   0
      Width           =   6630
      Begin VB.OptionButton OpcTotal 
         Caption         =   "Precancelacion Total"
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
         Left            =   3255
         TabIndex        =   2
         Top             =   210
         Width           =   2430
      End
      Begin VB.OptionButton OpcParcial 
         Caption         =   "Pago Anticipado"
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
         Left            =   210
         TabIndex        =   1
         Top             =   210
         Value           =   -1  'True
         Width           =   2115
      End
   End
   Begin VB.TextBox TextNumero 
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
      Left            =   4725
      MaxLength       =   49
      MultiLine       =   -1  'True
      TabIndex        =   46
      Top             =   3150
      Width           =   1275
   End
   Begin VB.TextBox TxtApellidosS 
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
      Left            =   5775
      MaxLength       =   30
      TabIndex        =   14
      Top             =   1680
      Width           =   4110
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   6825
      TabIndex        =   43
      Top             =   0
      Width           =   4425
      Begin VB.OptionButton OpcL 
         Caption         =   "Libreta"
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
         TabIndex        =   45
         Top             =   210
         Width           =   1170
      End
      Begin VB.OptionButton OpcC 
         Caption         =   "Caja"
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
         Left            =   210
         TabIndex        =   44
         Top             =   210
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Comprobante de Pago"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   10980
      Picture         =   "PAbonos.frx":0034
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   1365
      Width           =   1275
   End
   Begin VB.TextBox TextSaldoDisp 
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
      Left            =   105
      MaxLength       =   49
      MultiLine       =   -1  'True
      TabIndex        =   37
      Top             =   3885
      Width           =   1590
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
      Height          =   330
      Left            =   1050
      MaxLength       =   100
      TabIndex        =   35
      Top             =   4305
      Width           =   8835
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Grabar Pago"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   10980
      Picture         =   "PAbonos.frx":08FE
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   2520
      Width           =   1275
   End
   Begin VB.TextBox TextTasa 
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
      Left            =   6090
      MaxLength       =   49
      MultiLine       =   -1  'True
      TabIndex        =   31
      Top             =   3150
      Width           =   960
   End
   Begin VB.TextBox TextDias 
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
      Left            =   8190
      MaxLength       =   49
      MultiLine       =   -1  'True
      TabIndex        =   29
      Top             =   3150
      Width           =   1695
   End
   Begin VB.TextBox TextMeses 
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
      Left            =   7140
      MaxLength       =   49
      MultiLine       =   -1  'True
      TabIndex        =   19
      Top             =   3150
      Width           =   960
   End
   Begin VB.TextBox TextTP 
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
      MaxLength       =   30
      TabIndex        =   22
      Top             =   3150
      Width           =   4530
   End
   Begin VB.CommandButton Command3 
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
      Height          =   960
      Left            =   10980
      Picture         =   "PAbonos.frx":0D40
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3675
      Width           =   1275
   End
   Begin VB.TextBox TxtNombresS 
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
      Left            =   1575
      MaxLength       =   30
      TabIndex        =   13
      Top             =   1680
      Width           =   4110
   End
   Begin VB.TextBox TxtRazonSocial 
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
      MaxLength       =   49
      TabIndex        =   15
      Top             =   2415
      Width           =   4530
   End
   Begin MSMask.MaskEdBox MBoxCuenta 
      Height          =   330
      Left            =   105
      TabIndex        =   4
      Top             =   945
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   192
      AllowPrompt     =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "CCCCCCCC-C"
      Mask            =   "########-#"
      PromptChar      =   "0"
   End
   Begin MSMask.MaskEdBox MBoxFecha 
      Height          =   330
      Left            =   105
      TabIndex        =   9
      ToolTipText     =   "Formato de Fecha: DD/MM/AA"
      Top             =   1680
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
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
   Begin MSAdodcLib.Adodc AdoTabla 
      Height          =   330
      Left            =   210
      Top             =   4935
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
      Caption         =   "Tabla"
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
   Begin MSAdodcLib.Adodc AdoGarantes 
      Height          =   330
      Left            =   210
      Top             =   5250
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
      Caption         =   "Garantes"
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
      Top             =   5880
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
   Begin MSAdodcLib.Adodc AdoCta 
      Height          =   330
      Left            =   2415
      Top             =   4935
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
      Caption         =   "Cta"
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
   Begin MSAdodcLib.Adodc AdoCtaNo 
      Height          =   330
      Left            =   2415
      Top             =   5250
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
      Caption         =   "CtaNo"
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
      Left            =   2415
      Top             =   5565
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
   Begin MSDataListLib.DataCombo DCCreditos 
      Bindings        =   "PAbonos.frx":1736
      DataSource      =   "AdoCreditos"
      Height          =   315
      Left            =   1575
      TabIndex        =   6
      Top             =   945
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "009000000"
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
   Begin MSAdodcLib.Adodc AdoCreditos 
      Height          =   330
      Left            =   2415
      Top             =   5880
      Width           =   2115
      _ExtentX        =   3731
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
      Caption         =   "Creditos"
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
   Begin MSAdodcLib.Adodc AdoPrestamos 
      Height          =   330
      Left            =   210
      Top             =   6195
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
      Caption         =   "Prestamos"
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
   Begin VB.Label LblPrestamo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   4725
      TabIndex        =   53
      Top             =   2415
      Width           =   2115
   End
   Begin VB.Label Label20 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " PRESTAMO ORIGINAL"
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
      Left            =   4725
      TabIndex        =   52
      Top             =   2100
      Width           =   2115
   End
   Begin VB.Label Label19 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Línea No"
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
      Left            =   8925
      TabIndex        =   49
      Top             =   3570
      Width           =   960
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Saldo"
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
      Left            =   7455
      TabIndex        =   28
      Top             =   3570
      Width           =   1485
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Capital"
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
      Left            =   5985
      TabIndex        =   26
      Top             =   3570
      Width           =   1485
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Comision"
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
      Left            =   4515
      TabIndex        =   34
      Top             =   3570
      Width           =   1485
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Interes"
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
      TabIndex        =   24
      Top             =   3570
      Width           =   1380
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Monto a pagar"
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
      Left            =   1680
      TabIndex        =   18
      Top             =   3570
      Width           =   1485
   End
   Begin VB.Label Label18 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Numero"
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
      Left            =   4725
      TabIndex        =   47
      Top             =   2835
      Width           =   1275
   End
   Begin VB.Label LabelEgresos 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   8505
      TabIndex        =   39
      Top             =   6825
      Width           =   1905
   End
   Begin VB.Label LabelIngresos 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   6615
      TabIndex        =   40
      Top             =   6825
      Width           =   1905
   End
   Begin VB.Label Label17 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TOTALES"
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
      Left            =   5565
      TabIndex        =   41
      Top             =   6825
      Width           =   1065
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Saldo Disponible"
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
      TabIndex        =   38
      Top             =   3570
      Width           =   1590
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Concepto"
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
      TabIndex        =   36
      Top             =   4305
      Width           =   1065
   End
   Begin VB.Label Label16 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tasa"
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
      TabIndex        =   32
      Top             =   2835
      Width           =   960
   End
   Begin VB.Label Label14 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Días Execedidos"
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
      Left            =   8190
      TabIndex        =   30
      Top             =   2835
      Width           =   1695
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Meses"
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
      Left            =   7140
      TabIndex        =   17
      Top             =   2835
      Width           =   960
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TIPO DE PRESTAMO"
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
      TabIndex        =   16
      Top             =   2835
      Width           =   4530
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Apellidos"
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
      Left            =   5775
      TabIndex        =   11
      Top             =   1365
      Width           =   4110
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nombres"
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
      Left            =   1575
      TabIndex        =   10
      Top             =   1365
      Width           =   4110
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &CUENTA No."
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
      Width           =   1380
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &LIQUIDACION DE PRESTAMOS"
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
      Left            =   1575
      TabIndex        =   5
      Top             =   630
      Width           =   9675
   End
   Begin VB.Label Label25 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Representante"
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
      Top             =   2100
      Width           =   4530
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FECHA"
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
      Top             =   1365
      Width           =   1380
   End
End
Attribute VB_Name = "Precancelacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub InsertarMontosPrestamo(DtaCta As Adodc, _
                                  CuentaNo As String, _
                                  TDebe As Currency, _
                                  THaber As Currency)
  If CuentaNo <> "00000000-0" Then
  SaldoDisp = 0: SaldoCont = 0: ID_Trans = 0
  TiempoTexto = Format(Time, FormatoTimes)
  If NumeroLineas <= 0 Then NumeroLineas = 1
  If Si_No Then
     If OpcC.value Then
        sSQL = "SELECT TOP 1 * " _
             & "FROM Trans_Cajas " _
             & "WHERE Cuenta_No = '" & CuentaNo & "' "
     Else
        sSQL = "SELECT TOP 1 * " _
             & "FROM Trans_Libretas " _
             & "WHERE Cuenta_No = '" & CuentaNo & "' " _
             & "ORDER BY Fecha DESC,IDT DESC,Hora DESC,ID DESC "
     End If
  Else
     sSQL = "SELECT TOP 1 * " _
          & "FROM Trans_Libretas " _
          & "WHERE Cuenta_No = '" & CuentaNo & "' " _
          & "ORDER BY Fecha DESC,IDT DESC,Hora DESC,ID DESC "
  End If
  SelectAdodc DtaCta, sSQL
  With DtaCta.Recordset
       SaldoDisp = 0: SaldoCont = 0
       ID_Trans = 0
       If Si_No Then
          If OpcL.value Then
             If .RecordCount > 0 Then
                 SaldoDisp = .Fields("Saldo_Disp")
                 SaldoCont = .Fields("Saldo_Cont")
                 ID_Trans = .Fields("IDT")
             End If
          End If
       Else
          If .RecordCount > 0 Then
              SaldoDisp = .Fields("Saldo_Disp")
              SaldoCont = .Fields("Saldo_Cont")
              ID_Trans = .Fields("IDT")
          End If
       End If
      .AddNew
      .Fields("Fecha") = FechaSistema
      .Fields("Cuenta_No") = CuentaNo
       If Si_No Then
          If OpcC.value Then
             .Fields("TP") = "BOVE"
          Else
             .Fields("TP") = TipoProc
             .Fields("Saldo_Cont") = SaldoCont + THaber - TDebe
             .Fields("Saldo_Disp") = SaldoDisp + THaber - TDebe
             .Fields("IDT") = ID_Trans + 1
          End If
       Else
         .Fields("TP") = TipoProc
         .Fields("Saldo_Cont") = SaldoCont + THaber - TDebe
         .Fields("Saldo_Disp") = SaldoDisp + THaber - TDebe
         .Fields("IDT") = ID_Trans + 1
       End If
      .Fields("Debitos") = TDebe
      .Fields("Creditos") = THaber
      .Fields("T") = Normal
      .Fields("CodigoU") = CodigoUsuario
      .Fields("Hora") = TiempoTexto
      .Fields("Item") = NumEmpresa
      .Fields("ME") = False
      .Fields("Cheque") = Ninguno
       SetUpdate DtaCta
  End With
  End If
End Sub

Public Sub ListarCuenta(Cuenta_No As String)
   TxtNombresS.Text = ""
   TxtApellidosS.Text = ""
   TxtNombresS.Text = ""
   TxtApellidosS.Text = ""
   TxtRazonSocial.Text = ""
   TextMonto.Text = "0"
   TextInt.Text = "0"
   TextCapital.Text = "0"
   TextSaldo.Text = "0"
   TextMeses.Text = "0"
   De_Vencidos = False: TotalEncaje = 0
   SaldoDisp = 0: SaldoCont = 0
   Total_Interes_Mora = 0
   sSQL = "SELECT * " _
        & "FROM Trans_Bloqueos " _
        & "WHERE Cuenta_No = '" & Cuenta_No & "' " _
        & "AND T = 'N' "
   SelectAdodc AdoAux, sSQL
   With AdoAux.Recordset
    If .RecordCount > 0 Then
        Do While Not .EOF
           TotalEncaje = TotalEncaje + .Fields("Valor")
          .MoveNext
        Loop
    End If
   End With
   sSQL = "SELECT TOP 1 * " _
        & "FROM Trans_Libretas " _
        & "WHERE Cuenta_No = '" & Cuenta_No & "' " _
        & "ORDER BY Fecha DESC,IDT DESC,Hora DESC,ID DESC "
   SelectAdodc AdoCtaNo, sSQL
   If AdoCtaNo.Recordset.RecordCount > 0 Then
      SaldoDisp = AdoCtaNo.Recordset.Fields("Saldo_Disp")
      SaldoCont = AdoCtaNo.Recordset.Fields("Saldo_Cont")
      TextLinea.Text = AdoCtaNo.Recordset.Fields("ID")
   End If
   SaldoDisp = SaldoDisp - TotalEncaje
   TextSaldoDisp.Text = Format(SaldoDisp, "#,##0.00")
   sSQL = "SELECT * " _
        & "FROM Clientes_Datos_Extras " _
        & "WHERE Cuenta_No = '" & Cuenta_No & "' " _
        & "AND Tipo_Dato = 'LIBRETAS' "
   SelectAdodc AdoCta, sSQL
   With AdoCta.Recordset
    If .RecordCount > 0 Then
        Moneda_US = False '.Fields("ME")
        CodigoCli = .Fields("Codigo")
    End If
   End With
   sSQL = "SELECT * " _
        & "FROM Clientes " _
        & "WHERE Codigo = '" & CodigoCli & "' "
   SelectAdodc AdoCta, sSQL
   With AdoCta.Recordset
    If .RecordCount > 0 Then
        CICliente = .Fields("CI_RUC")
        TxtNombresS.Text = .Fields("Cliente")
        TxtRazonSocial.Text = .Fields("Representante")
        Edad_Persona = Year(FechaSistema) - Year(.Fields("Fecha_N"))
        sSQL = "SELECT * " _
             & "FROM Clientes_Datos_Extras " _
             & "WHERE Cuenta_No = '" & Cuenta_No & "' " _
             & "AND Credito_No = '" & Contrato_No & "' " _
             & "AND Tipo_Dato = 'GARANTES' "
        SelectAdodc AdoGarantes, sSQL
        TextMonto.Text = "0"
        TextMeses.Text = "0"
        TextInt.Text = "0"
        TextComision.Text = "0"
        TextCapital.Text = "0"
        TextSaldo.Text = "0"
        TextMeses.Text = "0"
        TextNumero.Text = Contrato_No
       'Capital, Comision e Intereses Pendientes
        SaldoTotal = 0
        sSQL = "SELECT * " _
             & "FROM Trans_Prestamos " _
             & "WHERE Cuenta_No = '" & Cuenta_No & "' " _
             & "AND Credito_No = '" & Contrato_No & "' " _
             & "AND TP = '" & Codigo & "' " _
             & "AND T = 'P' " _
             & "ORDER BY T,TP,Credito_No,Fecha "
        SelectAdodc AdoTabla, sSQL
        With AdoTabla.Recordset
         If .RecordCount > 0 Then
             Do While Not .EOF
                SaldoTotal = SaldoTotal + .Fields("Interes") + .Fields("Comision")
               .MoveNext
             Loop
         End If
        End With
                
        If OpcParcial.value Then
           sSQL = "SELECT * "
        Else
           sSQL = "SELECT Credito_No,Cuenta_No,COUNT(Credito_No) As NMeses,SUM(Interes) As TInteres, " _
                & "SUM(Capital) As TCapital, SUM(Comision) As TComision "
        End If
        sSQL = sSQL & "FROM Trans_Prestamos " _
             & "WHERE Cuenta_No = '" & Cuenta_No & "' " _
             & "AND Credito_No = '" & Contrato_No & "' " _
             & "AND TP = '" & Codigo & "' " _
             & "AND T = 'P' "
        If OpcParcial.value Then
           sSQL = sSQL & "AND Fecha = #" & BuscarFecha(Mifecha) & "# " _
                & "ORDER BY T,TP,Credito_No,Fecha "
        Else
           sSQL = sSQL & "GROUP BY Credito_No,Cuenta_No "
        End If
        'MsgBox sSQL
        SelectAdodc AdoTabla, sSQL
        With AdoTabla.Recordset
         If .RecordCount > 0 Then
             If OpcParcial.value Then
                De_Vencidos = .Fields("V")
                MBoxFecha.Text = .Fields("Fecha")
                Cta_Prestamos = .Fields("Cta")
                TextMonto.Text = Format(.Fields("Pagos"), "#,##0.00")
                TextInt.Text = Format(.Fields("Interes"), "#,##0.00")
                TextComision.Text = Format(.Fields("Comision"), "#,##0.00")
                TextCapital.Text = Format(.Fields("Capital"), "#,##0.00")
                TextSaldo.Text = Format(.Fields("Saldo"), "#,##0.00")
                TextMeses.Text = .Fields("Cuota_No")
                Total_Saldos = .Fields("Saldo")
                TotalCapital = .Fields("Capital")
                TotalInteres = .Fields("Interes")
                TotalComision = .Fields("Comision")
             Else
                'Revisar para la cancelacion total
                'Cta_Prestamos = .Fields("Cta")
                MBoxFecha.Text = FechaSistema
                TextMeses.Text = .Fields("NMeses")
                TextInt.Text = Format(.Fields("TInteres"), "#,##0.00")
                TextComision.Text = Format(.Fields("TComision"), "#,##0.00")
                TextCapital.Text = Format(.Fields("TCapital"), "#,##0.00")
                TextMonto.Text = Format(.Fields("TCapital"), "#,##0.00")
                TextSaldo.Text = Format(.Fields("TCapital"), "#,##0.00")
                Total_Saldos = .Fields("TCapital")
                TotalCapital = .Fields("TCapital")
             End If
         End If
        End With
        Tasa = 0
        sSQL = "SELECT * " _
             & "FROM Prestamos " _
             & "WHERE Cuenta_No = '" & Cuenta_No & "' " _
             & "AND Credito_No = '" & Contrato_No & "' " _
             & "AND TP = '" & Codigo & "' "
        SelectAdodc AdoTabla, sSQL
        With AdoTabla.Recordset
         If .RecordCount > 0 Then
'             MsgBox "......"
             Tasa = .Fields("Tasa")
             If OpcTotal.value Then TextSaldo.Text = Format(.Fields("Saldo_Pendiente"), "#,##0.00")
             LblPrestamo.Caption = Format(AdoTabla.Recordset.Fields("Capital"), "#,##0.00")
         End If
        End With
        TextTasa.Text = Format(Tasa, "00.00")
        
        Codigo = SinEspaciosIzq(DCTipoPrestamo.Text)
        With AdoPrestamos.Recordset
         If .RecordCount > 0 Then
            .MoveFirst
            .Find ("CTP = '" & Codigo & "' ")
             If Not .EOF Then
                Si_No = .Fields("DM")
                TipoProc = .Fields("CTP")
                TextTP.Text = Codigo & "  " & .Fields("Descripcion")
                If Si_No Then Label5.Caption = " Dias" Else Label5.Caption = " Cuota"
                Una_Vez = .Fields("DM")
             End If
         End If
        End With
    End If
   End With
End Sub

Private Sub Command1_Click()
Dim CantGuion As Byte
Dim Imp_Rollo As Boolean
Dim Recibo_No As String
Dim Total_Seg_Desg As Currency
'Si la impresora esde odillo
Dim lhPrinter As Long
Dim lReturn As Long
Dim lpcWritten As Long
Dim lDoc As Long
Dim MyDocInfo As DOCINFO

On Error GoTo Errorhandler
CantGuion = CByte(Leer_Campo_Empresa("Cant_Ancho_PV"))
Imp_Rollo = CBool(Leer_Campo_Empresa("Impresora_Rodillo"))
Total_Seg_Desg = TotalComision + Total_Comision
Recibo_No = Trim("COMPROBANTE DE PAGO No. " & NumEmpresa & "-" & Format(ReadSetDataNum("Recibo_Ingreso", True, False), "00000000"))
Titulo = "IMPRESIONES"
Mensajes = "Imprimir Comprobante de Pago"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
   If Imp_Rollo Then
      lReturn = OpenPrinter(Printer.DeviceName, lhPrinter, 0)
      If lReturn = 0 Then
         MsgBox "The Printer Name you typed wasn't recognized."
         Exit Sub
      End If
      MyDocInfo.pDocName = "RECIBOS DE PAGO"
      MyDocInfo.pOutputFile = vbNullString
      MyDocInfo.pDatatype = vbNullString
      lDoc = StartDocPrinter(lhPrinter, 1, MyDocInfo)
      Call StartPagePrinter(lhPrinter)
   
     'MsgBox AnchoPapel
      RatonReloj
      InicioX = 0.5: InicioY = 0
      Pagina = 1
      HoraSistema = Format(Time, "HH:SS")
      Printer.FontName = TipoCourierNew
      Cadena = "Teléfono(s): " & Telefono1
      If Telefono1 <> Telefono2 Then Cadena = Cadena & "/" & Telefono2
      lReturn = WritePrinterText(lhPrinter, lpcWritten, UCase(Empresa) & vbCrLf, True, CantGuion)
      lReturn = WritePrinterText(lhPrinter, lpcWritten, NombreComercial & vbCrLf, True, CantGuion)
      lReturn = WritePrinterText(lhPrinter, lpcWritten, "R.U.C. " & RUC & vbCrLf, True, CantGuion)
      lReturn = WritePrinterText(lhPrinter, lpcWritten, Direccion & vbCrLf, True, CantGuion)
      lReturn = WritePrinterText(lhPrinter, lpcWritten, Cadena & vbCrLf, True, CantGuion)
      lReturn = WritePrinterText(lhPrinter, lpcWritten, "M A T R I Z  -  C O N D A D O" & vbCrLf, True, CantGuion)
      lReturn = WritePrinterText(lhPrinter, lpcWritten, "QUITO - ECUADOR" & vbCrLf, True, CantGuion)
      lReturn = WritePrinterText(lhPrinter, lpcWritten, String(CantGuion, "-") & vbCrLf, True, CantGuion)
      If OpcParcial.value Then
         lReturn = WritePrinterText(lhPrinter, lpcWritten, "A B O N O   A N T I C I P A D O" & vbCrLf, True, CantGuion)
      Else
         lReturn = WritePrinterText(lhPrinter, lpcWritten, "P R E C A N C E L A C I O N" & vbCrLf, True, CantGuion)
      End If
      lReturn = WritePrinterText(lhPrinter, lpcWritten, Recibo_No & vbCrLf, True, CantGuion)
      lReturn = WritePrinterText(lhPrinter, lpcWritten, "Fecha: " & FechaSistema & String(10, " ") & " Hora: " & HoraSistema & vbCrLf)
      lReturn = WritePrinterText(lhPrinter, lpcWritten, "CRÉDITO No. " & TextNumero & vbCrLf)
      Cadena = SinEspaciosIzq(TextTP)
      Cadena = Trim(MidStrg(TextTP, Len(Cadena) + 1, Len(TextTP)))
      lReturn = WritePrinterText(lhPrinter, lpcWritten, "TIPO DE CRÉDITO: " & UCase(Cadena) & vbCrLf)
      lReturn = WritePrinterText(lhPrinter, lpcWritten, "SOCIO: " & TxtNombresS & vbCrLf)
      lReturn = WritePrinterText(lhPrinter, lpcWritten, "CI/RUC: " & CICliente & vbCrLf)
      lReturn = WritePrinterText(lhPrinter, lpcWritten, "Cuenta No. " & Cuenta_No & vbCrLf)
      lReturn = WritePrinterText(lhPrinter, lpcWritten, String(CantGuion, "-") & vbCrLf, True, CantGuion)
      Cadena = "FORMA DE PAGO: "
      If OpcC.value Then
         Cadena = Cadena & "En Efectivo."
      Else
         Cadena = Cadena & "Debito a la cuenta."
      End If
      lReturn = WritePrinterText(lhPrinter, lpcWritten, Cadena & vbCrLf)
      lReturn = WritePrinterText(lhPrinter, lpcWritten, "Cuota cancelada No. " & TextMeses & " DE " & Val(SinEspaciosDer(Label9.Caption)) & vbCrLf)
      lReturn = WritePrinterText(lhPrinter, lpcWritten, String(CantGuion, "-") & vbCrLf, True, CantGuion)
      lReturn = WritePrinterText(lhPrinter, lpcWritten, "D E T A L L E           M O N T O" & vbCrLf)
      lReturn = WritePrinterText(lhPrinter, lpcWritten, String(CantGuion, "-") & vbCrLf)
      lReturn = WritePrinterText(lhPrinter, lpcWritten, "    Capital         " & Moneda & SetearBlancos(TotalCapital, 15, 0, True, , True) & vbCrLf)
      If OpcParcial.value Then
         lReturn = WritePrinterText(lhPrinter, lpcWritten, "    Interes         " & Moneda & SetearBlancos(TotalInteres, 15, 0, True, , True) & vbCrLf)
      Else
         lReturn = WritePrinterText(lhPrinter, lpcWritten, "    Int.Pronto Pago " & Moneda & SetearBlancos(TotalInteres, 15, 0, True, , True) & vbCrLf)
      End If
      If Total_Interes_Mora > 0 Then
         lReturn = WritePrinterText(lhPrinter, lpcWritten, "    Interes Mora    " & Moneda & SetearBlancos(Total_Interes_Mora, 15, 0, True, , True) & vbCrLf)
      End If
      If Total_Cobranza > 0 Then
         lReturn = WritePrinterText(lhPrinter, lpcWritten, "    Cobranza        " & Moneda & SetearBlancos(Total_Cobranza, 15, 0, True, , True) & vbCrLf)
      End If
      If Total_Seg_Desg > 0 Then
         lReturn = WritePrinterText(lhPrinter, lpcWritten, "    Seguro de Desg. " & Moneda & SetearBlancos(Total_Seg_Desg, 15, 0, True, , True) & vbCrLf)
      End If
      lReturn = WritePrinterText(lhPrinter, lpcWritten, String(24, " ") & String(14, "=") & vbCrLf)
      lReturn = WritePrinterText(lhPrinter, lpcWritten, "    Total Abono     " & Moneda & SetearBlancos(TotalLibreta, 15, 0, True, , True) & vbCrLf)
      lReturn = WritePrinterText(lhPrinter, lpcWritten, "Son: " & LCase(Cambio_Letras(TotalLibreta)) & vbCrLf)
      lReturn = WritePrinterText(lhPrinter, lpcWritten, String(CantGuion, "-") & vbCrLf)
      If OpcParcial.value Then
         lReturn = WritePrinterText(lhPrinter, lpcWritten, "DATOS RECORDATORIOS: " & vbCrLf)
         lReturn = WritePrinterText(lhPrinter, lpcWritten, "Saldo Capital       " & Moneda & SetearBlancos(Total_Saldos, 15, 0, True, , True) & vbCrLf)
      End If
      lReturn = WritePrinterText(lhPrinter, lpcWritten, vbCrLf)
      lReturn = WritePrinterText(lhPrinter, lpcWritten, vbCrLf)
      lReturn = WritePrinterText(lhPrinter, lpcWritten, vbCrLf)
      lReturn = WritePrinterText(lhPrinter, lpcWritten, String(16, "_") & String(6, " ") & String(16, "_") & vbCrLf)
      lReturn = WritePrinterText(lhPrinter, lpcWritten, "Cajero: " & Cambio_Usuario_Inicial(NombreUsuario) & String(12, " ") & "Conforme" & vbCrLf)
      lReturn = WritePrinterText(lhPrinter, lpcWritten, "NOTA: Este recibo es valido unicamente" & vbCrLf)
      lReturn = WritePrinterText(lhPrinter, lpcWritten, "con la firma y sello del cajero" & vbCrLf)
      lReturn = WritePrinterText(lhPrinter, lpcWritten, "GRACIAS POR SU PAGO" & vbCrLf, True, CantGuion)
      lReturn = WritePrinterText(lhPrinter, lpcWritten, vbCrLf)
      lReturn = WritePrinterText(lhPrinter, lpcWritten, vbCrLf)
      lReturn = WritePrinterText(lhPrinter, lpcWritten, vbCrLf)
      lReturn = WritePrinterText(lhPrinter, lpcWritten, vbCrLf)
      lReturn = WritePrinterText(lhPrinter, lpcWritten, vbCrLf)
      lReturn = WritePrinterText(lhPrinter, lpcWritten, vbCrLf)
      lReturn = WritePrinterText(lhPrinter, lpcWritten, vbCrLf)
      lReturn = EndPagePrinter(lhPrinter)
      lReturn = EndDocPrinter(lhPrinter)
      lReturn = ClosePrinter(lhPrinter)
      RatonNormal
   Else
      RatonReloj
      InicioX = 0.5: InicioY = 0
      DataAnchoCampos InicioX, AdoTabla, 8, TipoTimes, Orientacion_Pagina
      Pagina = 1
      EncabezadoEmpresa 0.1
      PrinterPaint LogoTipo, 2, 0.1, 3, 1.5
      Printer.FontName = TipoTimes
      Printer.FontSize = 12: Printer.FontBold = True
      PrinterTexto 2, 2, "C O M P R O B A N T E   D E   P A G O"
      Printer.FontSize = 11
      PrinterTexto 12, 2, NombreCiudad & ", " & FechaStrg(FechaSistema)
      PrinterTexto 2, 2.6, "Abono del Préstamo No.  " & TextNumero.Text
      PrinterTexto 12, 2.6, "Cta. Ahorro No. " & Cuenta_No
      PrinterTexto 2, 3.2, UCase(TextTP.Text)
      If OpcTotal.value Then
         PrinterTexto 12, 3.2, "Precancelación Total"
      Else
         PrinterTexto 12, 3.2, "Cuota No. " & TextMeses.Text
      End If
      PrinterTexto 2, 3.9, "SOCIO:"
      PrinterTexto 12, 4.4, "Capital"
      PrinterTexto 2, 4.5, "La cantidad de:"
      If OpcTotal.value Then
         PrinterTexto 12, 4.9, "Precancelación"
         Total_Saldos = 0
      Else
         PrinterTexto 12, 4.9, "Interés"
      End If
      PrinterTexto 12, 5.4, "Total Abono"
      PrinterTexto 12, 6, "Saldo Pendiente"
      PrinterTexto 14.8, 4.4, Moneda
      PrinterTexto 14.8, 4.9, Moneda
      PrinterTexto 14.8, 5.4, Moneda
      PrinterTexto 14.8, 6, Moneda
      Printer.FontBold = False
      PrinterTexto 3.5, 3.9, TxtNombresS.Text
      PrinterLineas 4.8, 4.5, Cambio_Letras(TotalLibreta), 7, 0.45
      PrinterVariables 16, 4.4, TotalCapital
      PrinterVariables 16, 4.9, TotalInteres
      PrinterVariables 16, 5.4, TotalLibreta
      PrinterVariables 16, 6, Total_Saldos
      Imprimir_Linea_H 5.4, 12, 19, Negro
      Imprimir_Linea_H 5.9, 12, 19, Negro, True
      PrinterTexto 2, 6, String(18, "_")
      PrinterTexto 8, 6, String(11, "_")
      PrinterTexto 2.1, 6.5, "Cajero: " & CodigoUsuario
      PrinterTexto 8.3, 6.5, "Conforme"
      RatonNormal
      Printer.EndDoc
   End If
MensajeEncabData = ""
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
Else
    RatonNormal
End If
End Sub

Private Sub Command2_Click()
Titulo = "Pregunta de Grabacion"
Mensajes = "Seguro de Grabar Abono"
If BoxMensaje = 6 Then
   RatonReloj
   'Numero = Val(TextNumero.Text)
   If OpcC.value Then SaldoDisp = TotalLibreta

   If Contra_Cta <> Cta_Libretas Then
      SetAdoAddNew "Trans_Libretas"
      SetAdoFields "T", Normal
      SetAdoFields "ME", False
      SetAdoFields "Fecha", FechaSistema
      SetAdoFields "Cuenta_No", Contrato_No
      SetAdoFields "TP", "BOVE"
      SetAdoFields "Debitos", 0
      SetAdoFields "Creditos", TotalLibreta
      SetAdoFields "CodigoU", CodigoUsuario
      SetAdoFields "Hora", Format(Time, FormatoTimes)
      SetAdoFields "Item", NumEmpresa
      SetAdoFields "Cheque", ""
      If OpcTotal.value Then
         SetAdoFields "Banco", "PRECANC. " & Cuenta_No & " " & Codigo
      Else
         SetAdoFields "Banco", "AB.AN " & Cuenta_No & " " & Codigo & " N." & Format(Val(TextMeses), "00")
      End If
      SetAdoFields "ACC", CBool(adFalse)
      SetAdoFields "CHT", CBool(adFalse)
      SetAdoUpdate
   End If
   If SaldoDisp >= TotalLibreta Then
   If Round(SumaDebe - SumaHaber, 2) = 0 Then
      RatonReloj
      If OpcL.value Then InsertarMontosPrestamo AdoTabla, Cuenta_No, TotalLibreta, 0
      If OpcParcial.value Then
         sSQL = "UPDATE Trans_Prestamos " _
              & "SET T = 'C'," _
              & "Fecha_C = #" & BuscarFecha(FechaSistema) & "# " _
              & "WHERE Cuenta_No = '" & Cuenta_No & "' " _
              & "AND Credito_No = '" & Contrato_No & "' " _
              & "AND TP = '" & Codigo & "' " _
              & "AND Fecha = #" & BuscarFecha(Mifecha) & "# " _
              & "AND Cuota_No = " & CInt(TextMeses.Text) & " "
         ConectarAdoExecute sSQL
         Saldo = CDbl(TextSaldo.Text)
         sSQL = "UPDATE Prestamos " _
              & "SET Saldo_Pendiente = " & Saldo & " "
         If Saldo <= 0 Then sSQL = sSQL & ", T = 'C' "
         sSQL = sSQL & "WHERE Cuenta_No = '" & Cuenta_No & "' " _
              & "AND Credito_No = '" & Contrato_No & "' " _
              & "AND TP = '" & Codigo & "' "
         ConectarAdoExecute sSQL
      Else
         sSQL = "UPDATE Trans_Prestamos " _
              & "SET T = 'C'," _
              & "Fecha_C = #" & BuscarFecha(FechaSistema) & "# " _
              & "WHERE Cuenta_No = '" & Cuenta_No & "' " _
              & "AND Credito_No = '" & Contrato_No & "' " _
              & "AND TP = '" & Codigo & "' " _
              & "AND T <> 'C' "
         ConectarAdoExecute sSQL
         Saldo = 0
         sSQL = "UPDATE Prestamos " _
              & "SET Saldo_Pendiente = " & Saldo & " "
         If Saldo <= 0 Then sSQL = sSQL & ", T = 'C' "
         sSQL = sSQL & "WHERE Cuenta_No = '" & Cuenta_No & "' " _
              & "AND Credito_No = '" & Contrato_No & "' " _
              & "AND TP = '" & Codigo & "' "
         ConectarAdoExecute sSQL
      End If
      Co.T = Normal
      Co.TP = CompDiario
      Co.Fecha = FechaSistema
      Co.CodigoB = Ninguno
      Co.Efectivo = 0
      Co.Monto_Total = 0
      Co.Numero = ReadSetDataNum("Diario", True, True)
      Co.Concepto = TextConcepto.Text
      Co.T_No = Trans_No
      Co.Item = NumEmpresa
      Co.Usuario = CodigoUsuario
      GrabarComprobante Co
      If OpcL.value Then Imprimir_Libreta Cuenta_No, AdoCtaNo, 1, 8, Val(TextLinea.Text)
   End If
   RatonReloj
   Mifecha = BuscarFecha(FechaSistema)
   TipoDoc = CompDiario
   Trans_No = 54
   IniciarAsientosDe DGAsiento, AdoAsiento
   ListarPrecancelacion
   DCTipoPrestamo.SetFocus
   RatonNormal
  Else
    MsgBox "Usted no puede Abonar por no tener fondos su libreta"
  End If
  RatonNormal
End If
End Sub

Private Sub Command3_Click()
  Unload Precancelacion
End Sub

Private Sub DCCreditos_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCCreditos_LostFocus()
  ListarPrecancelacion
  DCTipoPrestamo.SetFocus
End Sub

Private Sub DCTipoPrestamo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCTipoPrestamo_LostFocus()
Dim Seguro As Single
Dim Seguro1 As Single

  If AdoTipoPrest.Recordset.RecordCount > 0 Then
     AdoTipoPrest.Recordset.MoveFirst
     DCTipoPrestamo.Text = AdoTipoPrest.Recordset.Fields("TipoP")
  Trans_No = 54
  Codigo = SinEspaciosIzq(DCTipoPrestamo.Text)
  TipoProc = SinEspaciosIzq(DCTipoPrestamo.Text)
  Contrato_No = SinEspaciosIzqNoBlancos(DCTipoPrestamo.Text, 3)
  Mifecha = AdoTipoPrest.Recordset.Fields("Fecha")
  'SinEspaciosIzqNoBlancos(DCTipoPrestamo.Text, 3)
  Cuenta_No = SinEspaciosIzqNoBlancos(DCTipoPrestamo.Text, 2)
  If Cuenta_No = "0" Then Cuenta_No = "00000000-0"
  If Codigo = "" Then Codigo = Ninguno
  If Mifecha = "datos." Then Mifecha = FechaSistema
  ListarCuenta Cuenta_No
  Label25.Caption = " Representante, Edad Actual " & Edad_Persona & " años"
  If Edad_Persona <= 65 Then
     Seguro = Leer_Campo_Empresa("Seguro") / 100000
  Else
     Seguro = Leer_Campo_Empresa("Seguro2") / 10000
  End If
  Seguro1 = Leer_Campo_Empresa("Seguro2") / 10000

  Si_No = True
  If Si_No Then
     Titulo = "TIPO DE TRANSACCION"
     Mensajes = "Transaccion en: [Si] Caja y [No] Libreta"
     If BoxMensaje = vbYes Then
        OpcC.value = True
        Si_No = True
     Else
        OpcL.value = True
        Si_No = False
     End If
     Frame1.Visible = True
  Else
     Frame1.Visible = False
  End If
  IniciarAsientosDe DGAsiento, AdoAsiento
  If OpcParcial.value Then
     TextConcepto.Text = "(" & NumEmpresa & ") Abono No. " & TextMeses.Text & ", Cuenta No. " & Cuenta_No & " de " & TxtApellidosS.Text & " " & TxtNombresS.Text
  Else
     TextConcepto.Text = "(" & NumEmpresa & ") Precancelacion total de la Cuenta No. " & Cuenta_No & " de " & TxtApellidosS.Text & " " & TxtNombresS.Text
  End If
  Total = Round(CCur(TextMonto.Text), 2)
  NoMeses = Round(CCur(TextMeses.Text), 2)
  Total_Interes = Round(CCur(TextInt.Text), 2)
  Total_Comision = Round(CCur(TextComision.Text), 2)
  Debe = Round(CCur(TextCapital.Text), 2)
  TotalComAgente = 0
  Total_Comision = 0
  TotalComision = 0
  If Seguro > 0 Then
     TotalComision = Round(CCur(LblPrestamo.Caption) * Seguro, 2)
     Total_Comision = Round(Total * Seguro, 2)
     TotalComision = TotalComision - Total_Comision
  End If
  With AdoPrestamos.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("CTP = '" & TipoProc & "' ")
       If Not .EOF Then
          'MsgBox "...."
          If OpcParcial.value Then
           ' Asiento para Caja/Libreta
             InsertarAsientos AdoAsiento, Cta_Prestamos, 0, 0, Debe
             InsertarAsientos AdoAsiento, .Fields("Cta_Interes"), 0, 0, Total_Interes
             InsertarAsientos AdoAsiento, Cta_Seguro_I, 0, 0, TotalComision
             InsertarAsientos AdoAsiento, Cta_Seguro, 0, 0, Total_Comision
'''           ' Asiento de Efectivizacion de Prestamos
'''             InsertarAsientos AdoAsiento, .Fields("Cta_Int_Efec"), 0, Total_Interes, 0
'''             InsertarAsientos AdoAsiento, .Fields("Cta_Com_Efec"), 0, Total_Comision, 0
'''             InsertarAsientos AdoAsiento, .Fields("Cta_Int_Ganado"), 0, 0, Total_Interes
'''             InsertarAsientos AdoAsiento, .Fields("Cta_Com_Ganado"), 0, 0, Total_Comision
          Else
            'Haber = ((Total_Comision + Total_Interes) * .Fields("Interes_Reliq")) * Val(TextMeses.Text)
             If .Fields("Sobre_Interes") Then  'Valor de Liquidacion sobre los Interese
                 Haber = SaldoTotal * .Fields("Interes_Reliq") ' * Val(TextMeses.Text)
             Else  'Valor de Liquidacion sobre el saldo
                 Haber = Total * .Fields("Interes_Reliq") ' * Val(TextMeses.Text)
             End If
            'MsgBox SaldoTotal & "...."
           ' Asiento para la Libreta
             InsertarAsientos AdoAsiento, .Fields("Cta_Prestamo"), 0, 0, Debe
             InsertarAsientos AdoAsiento, .Fields("Cta_Reliquidacion"), 0, 0, Haber
           ' Asiento de Provisiones
            'cambio para no cobrar el interes por precancelacion total
            'InsertarAsientos AdoAsiento, .Fields("Cta_Interes"), 0, 0, Total_Interes
'             InsertarAsientos AdoAsiento, Cta_Seguro_I, 0, 0, TotalComision
'             InsertarAsientos AdoAsiento, Cta_Seguro, 0, 0, Total_Comision
''             InsertarAsientos AdoAsiento, .Fields("Cta_Int_Efec"), 0, Total_Interes, 0
''             InsertarAsientos AdoAsiento, .Fields("Cta_Com_Efec"), 0, Total_Comision, 0
           'MsgBox "..."
          End If
        'TextTP.Text = Codigo & "  " & .Fields("Descripcion")
        'TipoProc = .Fields("CTP")
       End If
   End If
  End With
 'Ingreso de Caja/Libreta
  Debe = 0: Haber = 0
  With AdoAsiento.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          Debe = Debe + .Fields("DEBE")
          Haber = Haber + .Fields("HABER")
         .MoveNext
       Loop
   End If
  End With
  TotalLibreta = Haber - Debe
  Contra_Cta = Ninguno
  If Si_No Then
     InsertarAsientos AdoAsiento, Cta_CajaG, 0, TotalLibreta, 0
     Contra_Cta = Cta_CajaG
  Else
     InsertarAsientos AdoAsiento, Cta_Libretas, 0, TotalLibreta, 0
     Contra_Cta = Cta_Libretas
  End If
  If OpcTotal.value Then TotalInteres = TotalLibreta - TotalCapital
  Debe = 0: Haber = 0
  With AdoAsiento.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          Debe = Debe + .Fields("DEBE")
          Haber = Haber + .Fields("HABER")
         .MoveNext
       Loop
   End If
  End With
  LabelIngresos.Caption = Format(Debe, "#,##0.00")
  LabelEgresos.Caption = Format(Haber, "#,##0.00")
'  TextDias.Text = "0"
  TxtNombresS.SetFocus
  End If
End Sub

Private Sub Form_Activate()
   Trans_No = 54
   sSQL = "SELECT * " _
        & "FROM Catalogo_Prestamo " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND TC <> " & Val(adFalse) & " " _
        & "ORDER BY CTP DESC "
   SelectAdodc AdoPrestamos, sSQL
   If Supervisor = False Then
     If CNivel(3) Then
        Command1.Enabled = False
        Command2.Enabled = False
     End If
   End If
   Mifecha = BuscarFecha(FechaSistema)
   TipoDoc = CompDiario
   IniciarAsientosDe DGAsiento, AdoAsiento
   ListarPrecancelacion
   OpcParcial.SetFocus
   RatonNormal
End Sub

Private Sub Form_Load()
  'CentrarForm Aprobacion
   ConectarAdodc AdoAux
   ConectarAdodc AdoCta
   ConectarAdodc AdoCtaNo
   ConectarAdodc AdoTabla
   ConectarAdodc AdoAsiento
   ConectarAdodc AdoGarantes
   ConectarAdodc AdoCreditos
   ConectarAdodc AdoPrestamos
   ConectarAdodc AdoTipoPrest
End Sub

Private Sub MBoxCuenta_GotFocus()
  MarcarTexto MBoxCuenta
End Sub

Private Sub MBoxCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Public Sub ListarPrecancelacion()
  Credito_No = SinEspaciosDer(DCCreditos.Text)
  If SQL_Server Then
     sSQL = "SELECT (TP & ' ' & Cuenta_No & ' ' & CONVERT(nvarchar(10),Fecha,103 ) & ' ' & Credito_No) As TipoP "
  Else
     sSQL = "SELECT (TP & ' ' & Cuenta_No & ' ' & CSTR(Fecha) & ' ' & Credito_No) As TipoP "
  End If
  sSQL = sSQL & "FROM Trans_Prestamos " _
       & "WHERE Cuenta_No = '" & MBoxCuenta.Text & "' " _
       & "AND Credito_No = '" & Credito_No & "' " _
       & "AND T <> 'A' " _
       & "ORDER BY TP,Cuenta_No,Fecha,Credito_No "
  SelectAdodc AdoTipoPrest, sSQL
  Label9.Caption = " TIPO DE PRESTAMO: " & Format(AdoTipoPrest.Recordset.RecordCount, "000")
  
  sSQL = "SELECT P.TP & " _
       & "' ' & P.Cuenta_No & " _
       & "' ' & P.Credito_No & " _
       & "' ' & C.Cliente As TipoP,P.Fecha " _
       & "FROM Trans_Prestamos As P,Clientes_Datos_Extras As Ct,Clientes As C " _
       & "WHERE P.Fecha > #" & BuscarFecha(FechaSistema) & "# " _
       & "AND Ct.Tipo_Dato = 'LIBRETAS' " _
       & "AND P.T = 'P' " _
       & "AND P.Cuenta_No = '" & MBoxCuenta.Text & "' " _
       & "AND P.Credito_No = '" & Credito_No & "' " _
       & "AND Ct.Codigo = C.Codigo " _
       & "AND Ct.Cuenta_No = P.Cuenta_No " _
       & "ORDER BY P.TP,P.Fecha,P.Credito_No,C.Cliente "
  SelectDBCombo DCTipoPrestamo, AdoTipoPrest, sSQL, "TipoP", False
End Sub

Private Sub MBoxCuenta_LostFocus()
  sSQL = "SELECT TP & ' ' & Credito_No As TipoCred " _
       & "FROM Trans_Prestamos " _
       & "WHERE Cuenta_No = '" & MBoxCuenta.Text & "' " _
       & "AND T = 'P' " _
       & "GROUP BY TP,Credito_No "
  SelectDBCombo DCCreditos, AdoCreditos, sSQL, "TipoCred"
End Sub
