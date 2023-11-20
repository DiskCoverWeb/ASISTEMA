VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form IGastosCaja 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso/Modificacion de SubCuentas"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   Begin MSDataListLib.DataCombo DCBanco 
      Bindings        =   "IGasCaja.frx":0000
      DataSource      =   "AdoBanco"
      Height          =   345
      Left            =   2625
      TabIndex        =   5
      Top             =   2310
      Width           =   7050
      _ExtentX        =   12435
      _ExtentY        =   609
      _Version        =   393216
      Text            =   "Banco"
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
   Begin MSDataListLib.DataList DLCtas 
      Bindings        =   "IGasCaja.frx":0017
      DataSource      =   "AdoSubCta"
      Height          =   1860
      Left            =   5250
      TabIndex        =   3
      Top             =   420
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   3281
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
      Caption         =   "Ingreso de Gastos Semanales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   105
      TabIndex        =   6
      Top             =   2730
      Width           =   10725
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
         Left            =   4095
         MaxLength       =   50
         TabIndex        =   28
         Top             =   1890
         Width           =   6525
      End
      Begin VB.TextBox TxtAutorizacion 
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
         TabIndex        =   26
         Top             =   1890
         Width           =   4005
      End
      Begin VB.TextBox TxtSecuencial 
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
         Left            =   9240
         MaxLength       =   9
         MultiLine       =   -1  'True
         TabIndex        =   24
         Text            =   "IGasCaja.frx":002F
         Top             =   1155
         Width           =   1380
      End
      Begin VB.TextBox TxtSerie 
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
         MaxLength       =   6
         TabIndex        =   22
         Text            =   "000000"
         Top             =   1155
         Width           =   855
      End
      Begin VB.ComboBox CCodRet 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7455
         TabIndex        =   20
         Text            =   "000"
         Top             =   1155
         Width           =   960
      End
      Begin MSDataListLib.DataCombo DCCliente 
         Bindings        =   "IGasCaja.frx":003B
         DataSource      =   "AdoCliente"
         Height          =   345
         Left            =   1365
         TabIndex        =   18
         Top             =   1155
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   609
         _Version        =   393216
         Text            =   "Cliente"
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
      Begin VB.TextBox TxtIVA 
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
         Left            =   5355
         MaxLength       =   12
         MultiLine       =   -1  'True
         TabIndex        =   32
         Text            =   "IGasCaja.frx":0054
         Top             =   2310
         Width           =   1485
      End
      Begin VB.Frame Frame2 
         Height          =   540
         Left            =   105
         TabIndex        =   7
         Top             =   210
         Width           =   2325
         Begin VB.OptionButton OpcE 
            Caption         =   "Egreso"
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
            Top             =   210
            Value           =   -1  'True
            Width           =   1065
         End
         Begin VB.OptionButton OpcI 
            Caption         =   "Ingreso"
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
            TabIndex        =   8
            Top             =   210
            Width           =   1065
         End
      End
      Begin MSDataGridLib.DataGrid DGMayor 
         Bindings        =   "IGasCaja.frx":005B
         Height          =   2220
         Left            =   105
         TabIndex        =   38
         Top             =   2730
         Width           =   10515
         _ExtentX        =   18547
         _ExtentY        =   3916
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         AllowDelete     =   -1  'True
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
      Begin VB.Frame Frame4 
         Caption         =   "Movimientos de: "
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
         Left            =   6090
         TabIndex        =   10
         Top             =   210
         Width           =   4530
         Begin VB.OptionButton Opc3m 
            Caption         =   "3 meses"
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
            Left            =   2310
            TabIndex        =   13
            Top             =   210
            Width           =   1065
         End
         Begin VB.OptionButton OpcTodos 
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
            Left            =   3465
            TabIndex        =   14
            Top             =   210
            Width           =   960
         End
         Begin VB.OptionButton Opc31 
            Caption         =   "31 días"
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
            TabIndex        =   12
            Top             =   210
            Width           =   1065
         End
         Begin VB.OptionButton Opc7 
            Caption         =   "7 días"
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
            TabIndex        =   11
            Top             =   210
            Value           =   -1  'True
            Width           =   960
         End
      End
      Begin VB.TextBox TextPresupuesto 
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
         Left            =   2625
         MaxLength       =   12
         MultiLine       =   -1  'True
         TabIndex        =   30
         Text            =   "IGasCaja.frx":0074
         Top             =   2310
         Width           =   1485
      End
      Begin MSMask.MaskEdBox MBoxFecha 
         Height          =   330
         Left            =   105
         TabIndex        =   16
         Top             =   1155
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
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " AUTORIZACION"
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
         TabIndex        =   25
         Top             =   1575
         Width           =   4005
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CONCEPTO"
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
         TabIndex        =   27
         Top             =   1575
         Width           =   6525
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " SECUENCIAL"
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
         TabIndex        =   23
         Top             =   840
         Width           =   1380
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " SERIE"
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
         TabIndex        =   21
         Top             =   840
         Width           =   855
      End
      Begin VB.Label LblTotal 
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
         Height          =   330
         Left            =   9135
         TabIndex        =   34
         Top             =   2310
         Width           =   1485
      End
      Begin VB.Label Label10 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "COD.RET."
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
         TabIndex        =   19
         Top             =   840
         Width           =   960
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL"
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
         Left            =   6825
         TabIndex        =   33
         Top             =   2310
         Width           =   2325
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " IVA"
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
         TabIndex        =   31
         Top             =   2310
         Width           =   1275
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " SUBTOTAL"
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
         TabIndex        =   29
         Top             =   2310
         Width           =   2535
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " PAGADO A:"
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
         Left            =   1365
         TabIndex        =   17
         Top             =   840
         Width           =   6105
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " FECHA:"
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
         TabIndex        =   15
         Top             =   840
         Width           =   1275
      End
   End
   Begin MSDataListLib.DataList DLCtaGasto 
      Bindings        =   "IGasCaja.frx":007B
      DataSource      =   "AdoCtas"
      Height          =   1860
      Left            =   105
      TabIndex        =   1
      Top             =   420
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   3281
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
   Begin VB.CommandButton Command4 
      Caption         =   "&Recibo"
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
      Left            =   9765
      Picture         =   "IGasCaja.frx":0091
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   945
      Width           =   1065
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Grabar"
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
      Left            =   9765
      Picture         =   "IGasCaja.frx":095B
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   105
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc AdoSubCta 
      Height          =   330
      Left            =   315
      Top             =   4200
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
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
   Begin VB.CommandButton Command2 
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
      Height          =   750
      Left            =   9765
      Picture         =   "IGasCaja.frx":0D9D
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   1785
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc AdoSubCta1 
      Height          =   330
      Left            =   315
      Top             =   4515
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
   Begin MSAdodcLib.Adodc AdoSubCta2 
      Height          =   330
      Left            =   315
      Top             =   5460
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
      Caption         =   "SubCta2"
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
      Top             =   5775
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
   Begin MSAdodcLib.Adodc AdoCliente 
      Height          =   330
      Left            =   315
      Top             =   4830
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
      Caption         =   "Cliente"
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
   Begin MSAdodcLib.Adodc AdoBanco 
      Height          =   330
      Left            =   315
      Top             =   6090
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
   Begin MSAdodcLib.Adodc AdoCodRet 
      Height          =   330
      Left            =   315
      Top             =   5145
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
      Caption         =   "CodRet"
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
      Top             =   6405
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
   Begin VB.Label Label16 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " PAGADO POR CAJA:"
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
      TabIndex        =   4
      Top             =   2310
      Width           =   2535
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CODIGOS DE GASTOS DE SUB-CUENTAS"
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
      Left            =   5250
      TabIndex        =   2
      Top             =   105
      Width           =   4425
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CUENTAS DE GASTOS"
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
      Width           =   5160
   End
End
Attribute VB_Name = "IGastosCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ListarCajaChica()
Dim TipoCodigo As String
   FechaValida MBoxFecha
  'If TipoCodigo = "" Then TipoCodigo = Ninguno
   If Opc7.value Then NoDias = 7
   If Opc31.value Then NoDias = 31
   If Opc3m.value Then NoDias = 240
   If OpcTodos.value Then NoDias = 3650
   FechaIni = BuscarFecha(CLongFecha(CFechaLong(MBoxFecha.Text) - NoDias))
   FechaFin = BuscarFecha(MBoxFecha.Text)
   With AdoSubCta.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
       .Find ("Detalle = '" & DLCtas.Text & "' ")
        If Not .EOF Then TipoCodigo = .Fields("Codigo")
    Else
        TipoCodigo = Ninguno
    End If
   End With
   sSQL = "SELECT * " _
        & "FROM Trans_Gastos_Caja " _
        & "WHERE Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "ORDER BY Fecha,Codigo,Numero,Ingreso DESC,Egreso "
   Select_Adodc_Grid DGMayor, AdoSubCta2, sSQL
   MBoxFecha.SetFocus
End Sub

Private Sub CCodRet_GotFocus()
  MarcarTexto CCodRet
End Sub

Private Sub CCodRet_LostFocus()
  If CCodRet = "" Then CCodRet = "000"
  If CCodRet = "000" Then
     TxtSerie.Enabled = False
     TxtSecuencial.Enabled = False
     TxtAutorizacion.Enabled = False
'     TextConcepto.SetFocus
  Else
     TxtSerie.Enabled = True
     TxtSecuencial.Enabled = True
     TxtAutorizacion.Enabled = True
     TxtSerie.SetFocus
  End If
End Sub

Private Sub Command2_Click()
  Unload IGastosCaja
End Sub

Private Sub Command3_Click()
   FechaValida MBoxFecha
   TextoValido TextPresupuesto, True
   TextoValido TxtIVA, True
   TextoValido TextConcepto
   TextoValido TxtSerie
   TextoValido TxtAutorizacion
   TextoValido TxtSecuencial, True, , 0
   Codigo = Ninguno
   Codigo1 = Ninguno
   CodigoCli = Ninguno
   ListarCajaChica
   With AdoSubCta.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
       .Find ("Detalle = '" & DLCtas.Text & "' ")
        If Not .EOF Then Codigo = .Fields("Codigo")
    End If
   End With
   If OpcI.value Then
      Codigo1 = Cta_General
   Else
      Codigo1 = SinEspaciosIzq(DLCtaGasto.Text)
   End If
   Codigo2 = SinEspaciosIzq(DCBanco)
   
   If Val(CCur(TextPresupuesto.Text)) > 0 Then
     'MsgBox Val(CCur(TextPresupuesto.Text))
      SetAddNew AdoSubCta2
      SetFields AdoSubCta2, "TC", "GC"
      If OpcI.value Then
         NumComp = ReadSetDataNum("Recibo_Ingreso", True, True)
         SetFields AdoSubCta2, "Ingreso", Redondear(TextPresupuesto.Text, 2)
         TRecibo.SubTotal = Redondear(TextPresupuesto.Text, 2)
         TRecibo.Tipo_Recibo = "I"
      Else
         NumComp = ReadSetDataNum("Recibo_Egreso", True, True)
         SetFields AdoSubCta2, "Egreso", Redondear(TextPresupuesto, 2)
         SetFields AdoSubCta2, "IVA", Redondear(TxtIVA, 2)
         SetFields AdoSubCta2, "CodRet", CCodRet
         TRecibo.SubTotal = Redondear(TextPresupuesto.Text, 2)
         TRecibo.Tipo_Recibo = "E"
         If CCodRet <> "000" Then
            SetFields AdoSubCta2, "Serie", TxtSerie
            SetFields AdoSubCta2, "Autorizacion", TxtAutorizacion
            SetFields AdoSubCta2, "Secuencial", TxtSecuencial
         End If
      End If
      TRecibo.IVA = Redondear(TxtIVA.Text, 2)
      TRecibo.Total = TRecibo.SubTotal + TRecibo.IVA
      SetFields AdoSubCta2, "Numero", NumComp
      SetFields AdoSubCta2, "Fecha", MBoxFecha
      SetFields AdoSubCta2, "Codigo", Codigo
      SetFields AdoSubCta2, "CodigoC", CodigoCliente
      SetFields AdoSubCta2, "Cta", Codigo1
      SetFields AdoSubCta2, "Contra_Cta", Codigo2
      SetFields AdoSubCta2, "Beneficiario", NombreCliente
      SetFields AdoSubCta2, "Concepto", TextConcepto
      SetFields AdoSubCta2, "Item", NumEmpresa
      SetFields AdoSubCta2, "Periodo", Periodo_Contable
      SetUpdate AdoSubCta2
      
      TRecibo.Cobrado_a = NombreCliente
      TRecibo.Concepto = TextConcepto
      If CCodRet <> "000" Then
         If TxtSerie <> "000000" Then TRecibo.Concepto = TRecibo.Concepto & ": Doc. " & TxtSerie
         If Len(TxtAutorizacion) >= 3 Then TRecibo.Concepto = TRecibo.Concepto & "-" & TxtAutorizacion
         If Val(TxtSecuencial) > 0 Then TRecibo.Concepto = TRecibo.Concepto & ", No. " & TxtSecuencial
      End If
      TRecibo.Fecha = MBoxFecha
      TRecibo.Recibo_No = Format(NumComp, "000000000")
      TRecibo.CI_RUC = ""
      Mensajes = "Imprimir Recibo de Caja"
      Titulo = "Pregunta de Impresion"
      If BoxMensaje = vbYes Then Imprimir_Recibo_Caja TRecibo
      
      'Imprimir_Recibo_De_Caja AdoSubCta1, NumComp, OpcI.value, NumEmpresa
   End If
   ListarCajaChica
   DCCliente.Text = "CONSUMIDOR FINAL"
   'LAC_1953.uejp
End Sub

Private Sub Command4_Click()
   Mensaje = "Introduzca Numero de Recibo: "
   Titulo = "NUMERO DE RECIBO"
   Cadena = "0"  'Establece el valor predeterminado.
   With AdoSubCta2.Recordset
    If .RecordCount > 0 Then
        NumItem = .Fields("Item")
        If .Fields("Ingreso") > 0 Then
           TRecibo.Tipo_Recibo = "I"
           TRecibo.SubTotal = .Fields("Ingreso")
        End If
        If .Fields("Egreso") > 0 Then
           TRecibo.Tipo_Recibo = "E"
           TRecibo.SubTotal = .Fields("Egreso")
        End If
        TRecibo.IVA = .Fields("IVA")
        TRecibo.Cobrado_a = .Fields("Beneficiario")
        TRecibo.Concepto = .Fields("Concepto")
        If .Fields("CodRet") <> "000" Then
            If .Fields("Serie") <> "000000" Then TRecibo.Concepto = TRecibo.Concepto & ": Doc. " & .Fields("Serie")
            If Len(.Fields("Autorizacion")) >= 3 Then TRecibo.Concepto = TRecibo.Concepto & "-" & .Fields("Autorizacion")
            If Val(.Fields("Secuencial")) > 0 Then TRecibo.Concepto = TRecibo.Concepto & ", No. " & .Fields("Secuencial")
        End If
        TRecibo.Fecha = .Fields("Fecha")
        TRecibo.Recibo_No = Format(.Fields("Numero"), "000000000")
        TRecibo.CI_RUC = ""
        TRecibo.Total = TRecibo.IVA + TRecibo.SubTotal
        Imprimir_Recibo_Caja TRecibo
    End If
   End With
   'Imprimir_Recibo_De_Caja AdoSubCta1, Numero, OpcI.value, NumItem
   'MsgBox NumItem
End Sub

Private Sub DCCliente_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCCliente_LostFocus()
Dim Continuar As Boolean
  Continuar = False
  With AdoCliente.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente = '" & DCCliente.Text & "'")
       If Not .EOF Then
          TipoContribuyente = ""
          CICliente = .Fields("CI_RUC")
          CodigoCliente = .Fields("Codigo")
          NombreCliente = .Fields("Cliente")
          Grupo_No = .Fields("Grupo")
          TipoDoc = .Fields("TD")
          TipoBenef = .Fields("TD")
          Continuar = True
       Else
          NombreCliente = DCCliente.Text
          Nuevo = True
          IGastosCaja.Visible = False
          FClientesFlash.Show 1
          IGastosCaja.Visible = True
          ListarClientes
          DCCliente.SetFocus
       End If
   Else
       IGastosCaja.Visible = False
       NombreCliente = DCCliente.Text
       Nuevo = True
       FClientesFlash.Show 1
       IGastosCaja.Visible = True
       ListarClientes
       DCCliente.SetFocus
   End If
  End With
  If Continuar Then
     sSQL = "SELECT TOP 1 CodRet,Serie,Secuencial,Autorizacion " _
          & "FROM Trans_Gastos_Caja " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND Fecha <= #" & BuscarFecha(MBoxFecha) & "# " _
          & "AND CodigoC = '" & CodigoCliente & "' " _
          & "ORDER BY Fecha DESC "
     Select_Adodc AdoAux, sSQL
     With AdoAux.Recordset
      If .RecordCount > 0 Then
          CCodRet = .Fields("CodRet")
          TxtSerie = .Fields("Serie")
          TxtAutorizacion = .Fields("Autorizacion")
          TxtSecuencial = .Fields("Secuencial") + 1
      Else
          CCodRet = "000"
          TxtSerie = "000000"
          TxtAutorizacion = Ninguno
          TxtSecuencial = "0"
      End If
     End With
  End If
End Sub

Private Sub DGMayor_BeforeDelete(Cancel As Integer)
  Mensajes = "¿Realmente desea eliminar el Registro " & vbCrLf
  Titulo = "Confirmación de eliminación"
  If BoxMensaje = vbYes Then Cancel = False Else Cancel = True
End Sub

Private Sub DLCtas_DblClick()
  SiguienteControl
End Sub

Private Sub DLCtas_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DLCtas_LostFocus()
  Cadena = SinEspaciosIzq(DLCtas.Text)
  Codigo1 = SinEspaciosIzq(DLCtas.Text)
  'LlenarCta Cadena
End Sub

Private Sub Form_Activate()
  Label3.Caption = " I.V.A. " & Porc_IVA * 100 & "%"
  FechaValida MBoxFecha, True
  CCodRet.Clear
  sSQL = "SELECT * " _
       & "FROM Tipo_Concepto_Retencion " _
       & "WHERE Porcentaje = 0 " _
       & "AND Ingresar_Porcentaje = 'N' " _
       & "AND Fecha_Inicio <= #" & BuscarFecha(MBoxFecha) & "# " _
       & "AND Fecha_Final >= #" & BuscarFecha(MBoxFecha) & "# " _
       & "ORDER BY Codigo "
  Select_Adodc AdoCodRet, sSQL
  CCodRet.AddItem "000"
  CCodRet.Text = "000"
  With AdoCodRet.Recordset
   If .RecordCount Then
       Do While Not .EOF
          CCodRet.AddItem .Fields("Codigo")
         .MoveNext
       Loop
   End If
  End With
  
  sSQL = "SELECT Detalle, Codigo " _
       & "FROM Catalogo_SubCtas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'G' " _
       & "AND Caja <> " & Val(adFalse) & " " _
       & "ORDER BY Detalle "
  SelectDB_List DLCtas, AdoSubCta, sSQL, "Detalle"
  sSQL = "SELECT Codigo & Space(10) & Cuenta As Nombre_Cta " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE TC = 'G' " _
       & "AND DG = 'D' "
  If OpcCoop Then
     sSQL = sSQL & "AND MidStrg(Codigo,1,1) = '4' "
  Else
     sSQL = sSQL & "AND MidStrg(Codigo,1,1) = '5' "
  End If
  sSQL = sSQL & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo "
  SelectDB_List DLCtaGasto, AdoCtas, sSQL, "Nombre_Cta"
  
  sSQL = "SELECT (Codigo & '  ' & Cuenta) As NomCuenta " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE TC IN ('BA','CJ','TJ') " _
       & "AND DG = 'D' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Mod_Gastos <> " & Val(adFalse) & " " _
       & "ORDER BY TC,Codigo "
  SelectDB_Combo DCBanco, AdoBanco, sSQL, "NomCuenta"
  ListarClientes
 ' MsgBox sSQL
  ListarCajaChica
  FechaValida MBoxFecha
  DLCtaGasto.SetFocus
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm IGastosCaja
  ConectarAdodc AdoAux
  ConectarAdodc AdoCtas
  ConectarAdodc AdoBanco
  ConectarAdodc AdoCodRet
  ConectarAdodc AdoSubCta
  ConectarAdodc AdoSubCta1
  ConectarAdodc AdoSubCta2
  ConectarAdodc AdoCliente
End Sub

Private Sub MBoxFecha_GotFocus()
  MarcarTexto MBoxFecha
End Sub

Private Sub MBoxFecha_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxFecha_LostFocus()
  FechaValida MBoxFecha
  Validar_Porc_IVA MBoxFecha
  Label3.Caption = " I.V.A. " & Porc_IVA * 100 & "%"
End Sub

Private Sub Opc31_Click()
  ListarCajaChica
End Sub

Private Sub Opc3m_Click()
  ListarCajaChica
End Sub

Private Sub Opc7_Click()
  ListarCajaChica
End Sub

Private Sub OpcE_Click()
  Label5.Caption = " PAGADO A:"
End Sub

Private Sub OpcI_Click()
  Label5.Caption = " RECIBI DE:"
End Sub

Private Sub OpcTodos_Click()
  ListarCajaChica
End Sub

Private Sub TextConcepto_GotFocus()
  MarcarTexto TextConcepto
End Sub

Private Sub TextConcepto_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextPresupuesto_GotFocus()
    MarcarTexto TextPresupuesto
End Sub

Private Sub TextPresupuesto_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextPresupuesto_LostFocus()
  TextoValido TextPresupuesto, True
  TextPresupuesto.Text = Format(CDbl(TextPresupuesto.Text), "#,##0.00")
End Sub

Private Sub TxtAutorizacion_GotFocus()
  MarcarTexto TxtAutorizacion
End Sub

Private Sub TxtAutorizacion_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtIVA_GotFocus()
  If OpcI.value Then
     TxtIVA = "0.00"
  Else
     TxtIVA = Format(CDbl(TextPresupuesto.Text) * Porc_IVA, "#,##0.00")
  End If
  MarcarTexto TxtIVA
End Sub

Private Sub TxtIVA_LostFocus()
  TxtIVA = Format(CDbl(TxtIVA), "#,##0.00")
  LblTotal.Caption = Format(CDbl(TextPresupuesto.Text) + CDbl(TxtIVA), "#,##0.00")
End Sub

Public Sub ListarClientes()
  'MsgBox "."
  sSQL = "SELECT Cliente,Codigo,CI_RUC,TD,Grupo,RISE,Especial " _
       & "FROM Clientes " _
       & "WHERE Codigo <> '_' " _
       & "ORDER BY Cliente "
  SelectDB_Combo DCCliente, AdoCliente, sSQL, "Cliente"
 'MsgBox "......"
End Sub

Private Sub TxtSecuencial_GotFocus()
  MarcarTexto TxtSecuencial
End Sub

Private Sub TxtSecuencial_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtSerie_GotFocus()
  MarcarTexto TxtSerie
End Sub

Private Sub TxtSerie_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub
