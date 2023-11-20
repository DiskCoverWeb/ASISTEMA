VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FExportaciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EXPORTACIONES"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10320
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FExportaciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   10320
   Begin MSDataListLib.DataCombo DCProveedor 
      Bindings        =   "FExportaciones.frx":0696
      DataSource      =   "AdoClientes"
      Height          =   315
      Left            =   105
      TabIndex        =   7
      Top             =   1575
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton CmdCerrar 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   9240
      Picture         =   "FExportaciones.frx":06B0
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Salir"
      Top             =   420
      Width           =   960
   End
   Begin VB.ComboBox CAño 
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
      ItemData        =   "FExportaciones.frx":0AF2
      Left            =   945
      List            =   "FExportaciones.frx":0AF4
      TabIndex        =   35
      Text            =   "2000"
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.ComboBox CMes 
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
      Left            =   1890
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.ComboBox CModificacion 
      DataSource      =   "AdoTransCompras"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3360
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   6945
   End
   Begin VB.Frame FrmTipoComprob 
      Caption         =   "TP Y NUMERO"
      Height          =   750
      Left            =   6090
      TabIndex        =   4
      Top             =   420
      Width           =   2010
      Begin VB.ComboBox CTP 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   105
         TabIndex        =   5
         ToolTipText     =   "En este combo se despliega una lista con lo stipos de comprobantes existentes tales como: Comprobante Diario, Ingreso o Egreso"
         Top             =   315
         Width           =   765
      End
      Begin VB.TextBox TxtNumeroC 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Left            =   945
         MaxLength       =   7
         MultiLine       =   -1  'True
         TabIndex        =   6
         Text            =   "FExportaciones.frx":0AF6
         ToolTipText     =   "En este campo se debe ingresar el número del comprobante, el cual no excedera los siete caracteres"
         Top             =   315
         Width           =   984
      End
   End
   Begin VB.Frame FrmRetencion 
      Caption         =   "Cuentas por Cobrar"
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
      Top             =   420
      Width           =   5895
      Begin MSDataListLib.DataCombo DCCXC 
         Bindings        =   "FExportaciones.frx":0AFA
         DataSource      =   "AdoCatalogo"
         Height          =   345
         Left            =   105
         TabIndex        =   3
         Top             =   315
         Width           =   5685
         _ExtentX        =   10028
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
   End
   Begin VB.Frame Frame1 
      Caption         =   " Periodo"
      Height          =   3480
      Left            =   105
      TabIndex        =   10
      Top             =   1890
      Width           =   10095
      Begin VB.OptionButton OpcSi 
         Caption         =   "SI"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   8610
         TabIndex        =   22
         ToolTipText     =   $"FExportaciones.frx":0B14
         Top             =   1260
         Width           =   645
      End
      Begin VB.OptionButton OpcNo 
         Caption         =   "NO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   9240
         TabIndex        =   23
         Top             =   1260
         Value           =   -1  'True
         Width           =   645
      End
      Begin VB.TextBox TxtNumDcto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4725
         MaxLength       =   16
         TabIndex        =   20
         ToolTipText     =   $"FExportaciones.frx":0B9E
         Top             =   1155
         Width           =   2535
      End
      Begin VB.Frame FraExpor 
         Caption         =   "Factura de Exportaciones"
         Height          =   1695
         Left            =   105
         TabIndex        =   24
         Top             =   1680
         Width           =   9885
         Begin VB.TextBox TxtNumSerieTresComp 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1470
            MaxLength       =   7
            TabIndex        =   29
            Text            =   "0000001"
            ToolTipText     =   $"FExportaciones.frx":0C65
            Top             =   1050
            Width           =   1380
         End
         Begin VB.TextBox TxtNumSerieUnoComp 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   210
            MaxLength       =   3
            TabIndex        =   27
            Text            =   "001"
            ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al código del establecimiento"
            Top             =   1050
            Width           =   645
         End
         Begin VB.TextBox TxtNumSerieDosComp 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   336
            Left            =   840
            MaxLength       =   3
            TabIndex        =   28
            Text            =   "001"
            ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al punto dde emisión"
            Top             =   1050
            Width           =   645
         End
         Begin VB.TextBox TxtNumAutComp 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   336
            Left            =   3045
            MaxLength       =   10
            TabIndex        =   30
            ToolTipText     =   $"FExportaciones.frx":0D08
            Top             =   1050
            Width           =   1590
         End
         Begin VB.TextBox TxtValorFOBC 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   6615
            TabIndex        =   26
            Text            =   "0.00"
            ToolTipText     =   "En este campo se debe ingresar el valor FOB correspondiente a la Factura de exportación"
            Top             =   420
            Width           =   1485
         End
         Begin MSDataListLib.DataCombo DCFacturaExpor 
            Bindings        =   "FExportaciones.frx":0D94
            DataSource      =   "AdoAux"
            Height          =   315
            Left            =   210
            TabIndex        =   25
            ToolTipText     =   $"FExportaciones.frx":0DA9
            Top             =   420
            Width           =   4950
            _ExtentX        =   8731
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSMask.MaskEdBox MBFechaEmiComp 
            Height          =   330
            Left            =   6615
            TabIndex        =   31
            ToolTipText     =   $"FExportaciones.frx":0E6B
            Top             =   1050
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "0"
         End
         Begin MSMask.MaskEdBox MBFechaRegistro 
            Height          =   330
            Left            =   8190
            TabIndex        =   32
            ToolTipText     =   $"FExportaciones.frx":0F17
            Top             =   1050
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "0"
         End
         Begin VB.Label Label20 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " REGISTRO"
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
            Left            =   8190
            TabIndex        =   54
            Top             =   840
            Width           =   1275
         End
         Begin VB.Label Label19 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " EMISIÓN"
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
            Left            =   6615
            TabIndex        =   53
            Top             =   840
            Width           =   1275
         End
         Begin VB.Label Label18 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " No. Autorización"
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
            Left            =   3045
            TabIndex        =   52
            Top             =   840
            Width           =   1590
         End
         Begin VB.Label Label16 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " No. COMPROBANTE MODIF."
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
            TabIndex        =   51
            Top             =   840
            Width           =   2640
         End
         Begin VB.Label Label15 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " FOB COMPRO."
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
            Left            =   6615
            TabIndex        =   50
            Top             =   210
            Width           =   1485
         End
         Begin VB.Label Label14 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Tipo de"
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
            TabIndex        =   49
            Top             =   210
            Width           =   4950
         End
      End
      Begin VB.Frame FraRefrendo 
         Caption         =   "Refrendo"
         ForeColor       =   &H00800000&
         Height          =   855
         Left            =   105
         TabIndex        =   14
         Top             =   735
         Width           =   4530
         Begin MSDataListLib.DataCombo DCRegimen 
            Bindings        =   "FExportaciones.frx":0F9F
            DataSource      =   "AdoRegimen"
            Height          =   315
            Left            =   1680
            TabIndex        =   17
            ToolTipText     =   "Regimen (2 caracteres)"
            Top             =   420
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.TextBox TxtAño 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   945
            MaxLength       =   4
            TabIndex        =   16
            ToolTipText     =   "Años (4 caracteres)"
            Top             =   420
            Width           =   750
         End
         Begin VB.TextBox TxtVerificador 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3465
            MaxLength       =   1
            TabIndex        =   19
            ToolTipText     =   "Verificador (1 caracter)"
            Top             =   420
            Width           =   855
         End
         Begin VB.TextBox TxtCorrelativo 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2625
            MaxLength       =   6
            TabIndex        =   18
            ToolTipText     =   "Correlativo (6 caracteres)"
            Top             =   420
            Width           =   855
         End
         Begin MSDataListLib.DataCombo DCDistrito 
            Bindings        =   "FExportaciones.frx":0FB8
            DataSource      =   "AdoDistrito"
            Height          =   315
            Left            =   105
            TabIndex        =   15
            ToolTipText     =   "Distrito aduanero (3 caracteres)"
            Top             =   420
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label10 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Verific."
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
            TabIndex        =   46
            Top             =   210
            Width           =   855
         End
         Begin VB.Label Label9 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Correla."
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
            Left            =   2625
            TabIndex        =   45
            Top             =   210
            Width           =   855
         End
         Begin VB.Label Label8 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Regimen"
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
            Left            =   1680
            TabIndex        =   44
            Top             =   210
            Width           =   960
         End
         Begin VB.Label Label7 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Año"
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
            Left            =   945
            TabIndex        =   43
            Top             =   210
            Width           =   750
         End
         Begin VB.Label Label6 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Distrito"
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
            TabIndex        =   42
            Top             =   210
            Width           =   855
         End
      End
      Begin VB.TextBox TxtValorFOB 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7245
         TabIndex        =   21
         Text            =   "0.00"
         ToolTipText     =   $"FExportaciones.frx":0FD2
         Top             =   1155
         Width           =   1275
      End
      Begin MSDataListLib.DataCombo DCExportacion 
         Bindings        =   "FExportaciones.frx":109E
         DataSource      =   "AdoExportacion"
         Height          =   315
         Left            =   105
         TabIndex        =   11
         ToolTipText     =   "Corresponde al tipo de transacción realizada (Bienes o Servicios), si es Exportación de Bienes o Ingresos del Exterior"
         Top             =   420
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox MBFechaEmbarque 
         Height          =   330
         Left            =   8085
         TabIndex        =   13
         ToolTipText     =   "En el caso de exportación de Bienes, en este campo se ingresará la feha efectuva en la que se realizaó el Embarque"
         Top             =   420
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "0"
      End
      Begin MSDataListLib.DataCombo DCTipoComprobante 
         Bindings        =   "FExportaciones.frx":10BB
         DataSource      =   "AdoTipoComprobante"
         Height          =   315
         Left            =   3045
         TabIndex        =   12
         ToolTipText     =   "En este combo de selección se desplegará una lista de comprobantes que sustentas la exportación"
         Top             =   420
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label13 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Devolucion IVA"
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
         TabIndex        =   48
         Top             =   840
         Width           =   1485
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " F.O.B."
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
         TabIndex        =   47
         Top             =   840
         Width           =   1275
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " NUMERO DOCUMENTO"
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
         TabIndex        =   41
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label4 
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
         Height          =   225
         Left            =   8085
         TabIndex        =   40
         Top             =   210
         Width           =   1170
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TIPO COMPROBANTE"
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
         Left            =   3045
         TabIndex        =   39
         Top             =   210
         Width           =   5055
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " EXPORTACION DE"
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
         TabIndex        =   38
         Top             =   210
         Width           =   2955
      End
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   8190
      Picture         =   "FExportaciones.frx":10DC
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Grabar"
      Top             =   420
      Width           =   960
   End
   Begin MSAdodcLib.Adodc AdoTipoComprobante 
      Height          =   330
      Left            =   315
      Top             =   2625
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "AdoTipoComp"
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
   Begin MSAdodcLib.Adodc AdoTransExportaciones 
      Height          =   330
      Left            =   315
      Top             =   3570
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "TransExportaciones"
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
   Begin MSAdodcLib.Adodc AdoAsientoExportaciones 
      Height          =   330
      Left            =   315
      Top             =   3885
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "AsientoExportaciones"
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
   Begin MSAdodcLib.Adodc AdoSustento 
      Height          =   330
      Left            =   315
      Top             =   1995
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "AdoSustento"
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
   Begin MSAdodcLib.Adodc AdoExportacion 
      Height          =   330
      Left            =   315
      Top             =   2310
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "AdoImportacion"
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
   Begin MSAdodcLib.Adodc AdoRegimen 
      Height          =   330
      Left            =   315
      Top             =   3255
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "AdoRegimen"
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
   Begin MSAdodcLib.Adodc AdoDistrito 
      Height          =   330
      Left            =   315
      Top             =   2940
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "AdoDistrito"
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
      Left            =   315
      Top             =   4200
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   315
      Top             =   4515
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "AdoAux"
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
   Begin MSAdodcLib.Adodc AdoCatalogo 
      Height          =   330
      Left            =   315
      Top             =   4830
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "AdoCatalogo"
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
   Begin VB.Label LblNumIdent 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
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
      Left            =   8295
      TabIndex        =   9
      Top             =   1575
      Width           =   1905
   End
   Begin VB.Label LblTD 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H008080FF&
      Height          =   330
      Left            =   7980
      TabIndex        =   8
      Top             =   1575
      Width           =   330
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " RAZON SOCIAL O NOMBRE DEL PROVEEDOR"
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
      TabIndex        =   37
      Top             =   1260
      Width           =   10095
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Periodo"
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
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "FExportaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MBFecha As MaskEdBox
Dim cod, CodImp As Byte
Dim ValorP, CodCat, CuentaC, Espizq, Cap1, Captc, CodProv As String
Dim Rf, Longitud As Byte

Private Sub CmdCerrar_Click()
  'Borra Asiento Exportaciones
  sSQL = "DELETE * " _
       & "FROM Asiento_Exportaciones " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  ConectarAdoExecute sSQL
  Unload Me
End Sub

Private Sub CmdGrabar_Click()
  RatonReloj
 'Valido las fechas
  FechaValida MBFechaEmiComp
  FechaValida MBFechaRegistro
  FechaValida MBFechaEmbarque
 'Borra si encuentra 2 o mas transacciones iguales
  Eliminar_Trans_AT "E", CodigoCliente, TxtNumSerieUnoComp, TxtNumSerieDosComp, TxtNumSerieTresComp, TxtNumAutComp, TxtCorrelativo, TxtVerificador, "", CMes, CAño, Ln_SRI
  Ln_No = 0
 'Pregunto antes de grabar
  Titulo = "Grabar Exportaciones"
  Mensajes = "Desea Grabar los Datos"
  If BoxMensaje = vbYes Then
    'Grabacion de los Datos
     Grabacion
     Titulo = "Grabar Exportaciones"
     Mensajes = "Los Datos fueron grabados correctamente" & vbCrLf _
              & "Desea ingresar otra transacción"""
     If BoxMensaje = vbYes Then
        Limpiar_Controles
        CTP.SetFocus
     Else
        Unload FExportaciones
     End If
  Else
     DCCXC.SetFocus
  End If
End Sub

Private Sub CMes_LostFocus()
  Modificacion_AT "E", CModificacion, CMes, CAño
End Sub

Private Sub CModificacion_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub CModificacion_LostFocus()
    CodImp = 0
    CodPorIva = 0
    CodPorIce = 0
    CodRetBien = 0
    CodRetServ = 0
       
    'Cargo los datos para la modificación
    CodigoCliente = Ninguno
    I = Val(SinEspaciosIzq(CModificacion))
    Cadena = SinEspaciosIzq(CModificacion)
    NombreCliente = Trim$(Mid$(CModificacion, Len(Cadena) + 1, Len(CModificacion)))
    With AdoClientes.Recordset
     If .RecordCount > 0 Then
        .MoveFirst
        .Find ("Cliente = '" & NombreCliente & "' ")
         If Not .EOF Then CodigoCliente = .Fields("Codigo")
     End If
    End With
    sSQL = "SELECT * " _
         & "FROM Trans_Exportaciones " _
         & "WHERE Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Linea_SRI = " & I & " " _
         & "AND IdFiscalProv = '" & CodigoCliente & "' " _
         & "ORDER BY Linea_SRI "
    SelectAdodc AdoTransExportaciones, sSQL
    With AdoTransExportaciones.Recordset
     If .RecordCount > 0 Then
         Ln_SRI = I
           'Busco el Proveedor
            CodProv = .Fields("IdFiscalProv")
            If AdoClientes.Recordset.RecordCount > 0 Then
               AdoClientes.Recordset.MoveFirst
               AdoClientes.Recordset.Find ("Codigo = '" & CodProv & "' ")
               If Not AdoClientes.Recordset.EOF Then
                  DCProveedor = AdoClientes.Recordset.Fields("Cliente")
                  LblTD = AdoClientes.Recordset.Fields("TD")
                  LblNumIdent = AdoClientes.Recordset.Fields("CI_RUC")

                 'Si existe beneficiario
                 'Cargo el Comprobante
                  sSQL = "SELECT Tipo_Comprobante_Codigo,Descripcion " _
                       & "FROM Tipo_Comprobante " _
                       & "WHERE Tipo_Comprobante_Codigo <> 0 "
                  SelectAdodc AdoTipoComprobante, sSQL
                  cod = .Fields("TipoComprobante")
                  If AdoTipoComprobante.Recordset.RecordCount > 0 Then
                     AdoTipoComprobante.Recordset.MoveFirst
                     AdoTipoComprobante.Recordset.Find ("Tipo_Comprobante_Codigo = '" & cod & "' ")
                     If Not AdoTipoComprobante.Recordset.EOF Then
                        DCTipoComprobante = AdoTipoComprobante.Recordset.Fields("Descripcion")
                     Else
                        MsgBox "El Comprobante no existe", vbInformation, "Aviso"
                     End If
                  End If
                    
                 'Cargo el Codigo de exportacion
                  CodImp = .Fields("ExportacionDe") 'Es un número
                  sSQL = "SELECT Concepto, Codigo " _
                       & "FROM Tabla_TipoImportacion " _
                       & "WHERE Concepto <> '.' " _
                       & "ORDER BY Concepto "
                  SelectAdodc AdoExportacion, sSQL
                  If AdoExportacion.Recordset.RecordCount > 0 Then
                     AdoExportacion.Recordset.MoveFirst
                     AdoExportacion.Recordset.Find ("Codigo = " & CodImp & " ")
                     If Not AdoExportacion.Recordset.EOF Then
                        DCExportacion = AdoExportacion.Recordset.Fields("Concepto")
                     Else
                        MsgBox "La Exportación no existe", vbInformation, "Aviso"
                     End If
                  End If
                 'Cargo la Fecha de embarque
                  MBFechaEmbarque = .Fields("FechaEmbarque")
                  If CodImp = 2 Then
                     FraRefrendo.Enabled = False
                     TxtNumDcto.Enabled = False
                  Else
                     TxtAño = .Fields("Anio")
                     DCDistrito.Text = .Fields("DistAduanero")
                     TxtVerificador = .Fields("Verificador")
                     DCRegimen.Text = .Fields("Regimen")
                     TxtCorrelativo = .Fields("Correlativo")
                     TxtNumDcto = .Fields("NumeroDctoTransporte")
                  End If
                  TxtNumeroC = .Fields("NumeroDctoTransporte")
                  TxtValorFOB = .Fields("ValorFOB")
                  TxtValorFOBC = .Fields("ValorFOBComprobante")
                  TxtNumSerieUnoComp = .Fields("Establecimiento")
                  TxtNumSerieDosComp = .Fields("PuntoEmision")
                  TxtNumSerieTresComp = .Fields("Secuencial")
                  TxtNumAutComp = .Fields("Autorizacion")
                  MBFechaEmiComp = .Fields("FechaEmision")
                  MBFechaRegistro = .Fields("FechaRegistro")
                  CTP.Text = .Fields("TP")
                  TxtNumeroC = .Fields("Numero")
               Else
                  MsgBox "Este beneficiario no existe", vbInformation, "Aviso"
               End If
             End If
      End If
  End With
End Sub

Private Sub CTP_GotFocus()
  MarcarTexto CTP
End Sub

Private Sub CTP_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub CTP_LostFocus()
  If IsNumeric(CTP.Text) Then
     MsgBox "No ingrese números. Vuelva a seleccionar.", vbInformation, "Aviso"
     CTP.Text = ""
     CTP.Text = "CD"
     CTP.SetFocus
  End If
End Sub

Private Sub DCCXC_GotFocus()
  MarcarTexto DCCXC
End Sub

Private Sub DCCXC_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCCXC_LostFocus()
  Captura_CXC
End Sub

Private Sub DCDistrito_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCDistrito_LostFocus()
  TxtAño.Text = Year(date)
End Sub

Private Sub DCExportacion_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCExportacion_LostFocus()
  'Desactivo dependiendo del Tipo de Exportacion
  If (DCExportacion = "Bienes") Then
     FraRefrendo.Enabled = True
     TxtNumDcto.Enabled = True
  Else
     FraRefrendo.Enabled = False
     TxtNumDcto.Enabled = False
  End If
End Sub

Private Sub DCProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCProveedor_LostFocus()
  Carga_TipoComprobanteSF
  If IsNumeric(DCProveedor.Text) Then
     MsgBox "No ingrese números. Vuelva a seleccionar.", vbInformation, "Aviso"
     DCProveedor.Text = ""
     Leer_Clientes
     DCProveedor.SetFocus
  Else
     NombreCliente = UCase$(DCProveedor)
     With AdoClientes.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
         .Find ("Cliente = '" & NombreCliente & "' ")
          If Not .EOF Then
             'Busca y captura el codigo de Porcentaje IVA
             CodigoCliente = .Fields("Codigo")
             DireccionCli = .Fields("Direccion")
             CICliente = .Fields("CI_RUC")
             TipoBenef = .Fields("TD")
             LblNumIdent = CICliente
             LblTD.Caption = TipoBenef
              
             TxtCorrelativo = "000001"
             'Aqui despliego el ultimo numero de la Transaccion
              sSQL = "SELECT TOP 1 * " _
                   & "FROM Trans_Exportaciones " _
                   & "WHERE IdFiscalProv = '" & CodigoCliente & "' " _
                   & "AND Item = '" & NumEmpresa & "' " _
                   & "AND Periodo = '" & Periodo_Contable & "' " _
                   & "ORDER BY Correlativo DESC "
              SelectAdodc AdoAux, sSQL
              With AdoAux.Recordset
               If .RecordCount > 0 Then TxtCorrelativo = .Fields("Correlativo")
              End With
           Else
              FClientesFlash.Show
              Leer_Clientes
           End If
      Else
         FClientesFlash.Show
         Leer_Clientes
      End If
     End With
  End If
End Sub

Private Sub DCRegimen_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCTipoComprobante_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCTipoComprobante_LostFocus()
  If IsNumeric(DCTipoComprobante.Text) Then
     MsgBox "No ingrese números. Vuelva a seleccionar.", vbInformation, "Aviso"
     DCTipoComprobante.Text = ""
     Carga_TipoComprobante
     DCTipoComprobante.SetFocus
     Captura_TipoComprobante
  Else
     Captura_TipoComprobante
  End If
End Sub

Private Sub Form_Activate()
  Ln_No = 1
  Ln_SRI = -1
  sSQL = Listar_Meses
  SelectAdodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          CMes.AddItem .Fields("Dia_Mes")
          CMes.Tag = .Fields("No_D_M")
         .MoveNext
       Loop
    End If
  End With
  CMes.Text = CMes.List(0)
  For I = Year(FechaSistema) To 2000 Step -1
      CAño.AddItem Format(I, "0000")
  Next I
  CAño.Text = CAño.List(0)

  'Cargo los datos en los combos
  Carga_Datos_Iniciales MBFecha, Nuevo
End Sub

Private Sub Form_Load()
  CentrarForm FExportaciones
  ConectarAdodc AdoSustento
  ConectarAdodc AdoCatalogo
  ConectarAdodc AdoTipoComprobante
  ConectarAdodc AdoAux
  ConectarAdodc AdoExportacion
  ConectarAdodc AdoDistrito
  ConectarAdodc AdoRegimen
  ConectarAdodc AdoClientes
  ConectarAdodc AdoClientes
  ConectarAdodc AdoAsientoExportaciones
  ConectarAdodc AdoTransExportaciones
End Sub

Public Sub Carga_TipoImportacion()
  'Carga la Tabla de Importaciones al DataCombo
  sSQL = "SELECT Concepto, Codigo " _
       & "FROM Tabla_TipoImportacion " _
       & "WHERE Concepto <> '.' " _
       & "ORDER BY Concepto "
  SelectDBCombo DCExportacion, AdoExportacion, sSQL, "Concepto"
End Sub

Public Sub Carga_TipoComprobanteSF()
  'Cargo solo la Factura
  sSQL = "SELECT Descripcion " _
       & "FROM Tipo_Comprobante " _
       & "WHERE Descripcion = 'Factura' " _
       & "ORDER BY Descripcion "
  SelectDBCombo DCFacturaExpor, AdoAux, sSQL, "Descripcion"
End Sub

Public Sub Carga_TipoComprobante()
  sSQL = "SELECT * " _
       & "FROM Tipo_Comprobante " _
       & "WHERE E <> " & Val(adFalse) & " " _
       & "ORDER BY Tipo_Comprobante_Codigo"
  SelectDBCombo DCTipoComprobante, AdoTipoComprobante, sSQL, "Descripcion"
End Sub

Public Sub Carga_Distrito()
  'Carga la Tabla de Distrito al DataCombo
  sSQL = "SELECT * " _
       & "FROM Tabla_Distrito " _
       & "WHERE Codigo <> '.' " _
       & "ORDER BY Codigo "
  SelectDBCombo DCDistrito, AdoDistrito, sSQL, "Codigo"
End Sub

Public Sub Carga_Regimen()
  'Carga la Tabla de Regimen al DataCombo
  sSQL = "SELECT * " _
       & "FROM Tabla_Regimen " _
       & "WHERE Codigo <> 200 " _
       & "ORDER BY Codigo "
  SelectDBCombo DCRegimen, AdoRegimen, sSQL, "Codigo"
End Sub

Public Sub Captura_TipoComprobante()
  'Captura lo que tiene el Combo de Tipo de Comprobante
  Captc = SinEspaciosIzq(DCTipoComprobante.Text)
  Cap1 = Trim$(DCTipoComprobante.Text)
     
  'Busca que sea igual a la Descripcion
  With AdoTipoComprobante.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Descripcion = '" & Cap1 & "' ")
       If Not .EOF Then
          cod = .Fields("Tipo_Comprobante_Codigo")
       Else
          MsgBox "Vuelva a seleccionar", vbInformation, "Aviso"
          DCTipoComprobante.SetFocus
       End If
   End If
  End With
End Sub

Public Sub Captura_CXC()
  Espizq = SinEspaciosIzq(DCCXC)
  CuentaC = Trim$(Mid$(DCCXC, Len(Espizq) + 4, Len(DCCXC)))
  'Busca que sea igual a la Descripcion
  With AdoCatalogo.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cuenta = '" & CuentaC & "' ")
       If Not .EOF Then CodCat = .Fields("Codigo")
   End If
  End With
End Sub

Public Sub Leer_Clientes()
  'Carga en el Data Combo los Clientes con su RUC
  sSQL = "SELECT * " _
       & "FROM Clientes " _
       & "WHERE Cliente <>  '.' " _
       & "AND TD <>  'O' " _
       & "AND TD <>  'E' " _
       & "ORDER BY  Cliente "
  SelectDBCombo DCProveedor, AdoClientes, sSQL, "Cliente"
End Sub

Public Sub Limpiar_Controles()
  TxtNumeroC.Text = ""
  DCExportacion.Text = ""
  FechaValida MBFechaEmbarque
  DCTipoComprobante.Text = ""
  DCDistrito.Text = ""
  TxtAño.Text = ""
  DCRegimen.Text = ""
  TxtCorrelativo.Text = ""
  TxtVerificador.Text = ""
  TxtNumDcto.Text = ""
  TxtValorFOB.Text = ""
  DCProveedor.Text = ""
  DCFacturaExpor.Text = ""
  TxtValorFOBC.Text = ""
  TxtNumSerieUnoComp.Text = ""
  TxtNumSerieDosComp.Text = ""
  TxtNumSerieTresComp.Text = ""
  TxtNumAutComp.Text = ""
  FechaValida MBFechaEmiComp
  FechaValida MBFechaRegistro
  LblTD.Caption = ""
End Sub

Private Sub MBFechaEmbarque_GotFocus()
  MarcarTexto MBFechaEmbarque
End Sub

Private Sub MBFechaEmbarque_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaEmbarque_LostFocus()
  FechaValida MBFechaEmbarque
End Sub

Private Sub MBFechaEmiComp_GotFocus()
  MarcarTexto MBFechaEmiComp
End Sub

Private Sub MBFechaEmiComp_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaEmiComp_LostFocus()
  FechaValida MBFechaEmiComp
  'Controla que la Fecha de Emisiòn este entre 01/01/2000 en adelante
  If CFechaLong(MBFechaEmiComp) < CFechaLong("01/01/2000") Then
     MsgBox "La Fecha de Emisión debe ser mayor que 01/01/2000", vbInformation, "Aviso"
     MBFechaEmiComp = "01/01/2000"
     MBFechaEmiComp.SetFocus
  End If
End Sub

Private Sub MBFechaRegistro_GotFocus()
  MarcarTexto MBFechaRegistro
End Sub

Private Sub MBFechaRegistro_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaRegistro_LostFocus()
  FechaValida MBFechaRegistro
End Sub

Private Sub OpcNo_LostFocus()
  If OpcNo.value = True Then ValorP = "N"
End Sub

Private Sub OpcSi_LostFocus()
  If OpcSi.value = True Then ValorP = "S"
End Sub

Private Sub TxtAño_GotFocus()
  MarcarTexto TxtAño
End Sub

Private Sub TxtAño_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtAño_LostFocus()
  TextoValido TxtAño
  TextoValido TxtAño
  If TxtAño = "" Then
     MsgBox "Ingrese el Año", vbInformation, "Aviso"
     TxtAño.SetFocus
  Else
     'Valida que sea la misma fecha de la liquidación
      Anio = Year(MBFechaEmbarque)
      If TxtAño <> Anio Then
         MsgBox "Año incorrecto. Vuelva a ingresar.", vbInformation, "Aviso"
         TxtAño.Text = ""
         TxtAño.SetFocus
      End If
  End If
End Sub

Private Sub TxtCorrelativo_GotFocus()
  MarcarTexto TxtCorrelativo
End Sub

Private Sub TxtCorrelativo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtCorrelativo_LostFocus()
  If Val(TxtCorrelativo) <= 0 Then TxtCorrelativo = "000000"
  TxtCorrelativo = Format(Val(Round(TxtCorrelativo)), String(6, "0"))
End Sub

Private Sub TxtNumAutComp_GotFocus()
  MarcarTexto TxtNumAutComp
End Sub

Private Sub TxtNumAutComp_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumAutComp_LostFocus()
  If Val(TxtNumAutComp) <= 0 Then TxtNumAutComp = "0000000001"
  TxtNumAutComp = Format(Val(Round(TxtNumAutComp)), String(10, "0"))
End Sub

Private Sub TxtNumDcto_GotFocus()
  MarcarTexto TxtNumDcto
End Sub

Private Sub TxtNumDcto_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter TxtNumDcto
End Sub

Private Sub TxtNumDcto_LostFocus()
  If Val(TxtNumDcto) <= 0 Then TxtNumDcto = "0000000000000000"
  TxtNumDcto = Format(Val(Round(TxtNumDcto)), String(16, "0"))
End Sub

Private Sub TxtNumeroC_GotFocus()
  MarcarTexto TxtNumeroC
End Sub

Private Sub TxtNumeroC_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumeroC_LostFocus()
  If Not IsNumeric(TxtNumeroC.Text) Then
     MsgBox "No ingrese caracteres alfabéticos. Vuelva a ingresar.", vbInformation, "Aviso"
     TxtNumeroC.Text = ""
     TxtNumeroC.SetFocus
  End If
  If TxtNumeroC.Text = "." Or TxtNumeroC.Text = " " Then
     MsgBox "Ingrese el Número de Comprobante", vbInformation, "Aviso"
  End If
  If Val(TxtNumeroC) <= 0 Then TxtNumeroC = "0"
  TxtNumeroC = Format(CTNumero(TxtNumeroC), String(7, "0"))
End Sub

Private Sub TxtNumSerieDosComp_GotFocus()
  MarcarTexto TxtNumSerieDosComp
End Sub

Private Sub TxtNumSerieDosComp_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumSerieDosComp_LostFocus()
  TextoValido TxtNumSerieDosComp, True, , 0
  If Val(TxtNumSerieDosComp) <= 0 Then TxtNumSerieDosComp = "001"
  TxtNumSerieDosComp = Format(Val(TxtNumSerieDosComp), "000")
End Sub

Private Sub TxtNumSerieTresComp_GotFocus()
  MarcarTexto TxtNumSerieTresComp
End Sub

Private Sub TxtNumSerieTresComp_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumSerieTresComp_LostFocus()
  TextoValido TxtNumSerieTresComp, True, , 0
  If Val(TxtNumSerieTresComp) <= 0 Then TxtNumSerieTresComp = "0000001"
  TxtNumSerieTresComp = Format(Val(Round(TxtNumSerieTresComp)), "0000000")
End Sub

Private Sub TxtNumSerieUnoComp_GotFocus()
  MarcarTexto TxtNumSerieUnoComp
End Sub

Private Sub TxtNumSerieUnoComp_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumSerieUnoComp_LostFocus()
  TextoValido TxtNumSerieUnoComp, True, , 0
  If Val(TxtNumSerieUnoComp) <= 0 Then TxtNumSerieUnoComp = "001"
  TxtNumSerieUnoComp = Format(Val(TxtNumSerieUnoComp), "000")
End Sub

Private Sub TxtValorFOB_GotFocus()
  MarcarTexto TxtValorFOB
End Sub

Private Sub TxtValorFOB_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtValorFOB_LostFocus()
  TextoValido TxtValorFOB, True, , 0
End Sub

Private Sub TxtValorFOBC_GotFocus()
  MarcarTexto TxtValorFOBC
End Sub

Private Sub TxtValorFOBC_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtValorFOBC_LostFocus()
  TextoValido TxtValorFOBC, True, , 0
End Sub

Private Sub TxtVerificador_GotFocus()
  MarcarTexto TxtVerificador
End Sub

Private Sub TxtVerificador_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Public Sub Captura_TipoExportacion()
  'Busca y captura el codigo de Tipo de Importación
  With AdoExportacion.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Concepto = '" & DCExportacion & "' ")
       If Not .EOF Then
          CodImp = .Fields("Codigo")
       Else
          MsgBox "Seleccione bien", vbInformation, "Aviso"
          DCExportacion.SetFocus
       End If
    End If
  End With
End Sub

Public Sub Carga_CXC()
  'Carga las Cuentas por Cobrar desde el Catalogo de Cuentas
  sSQL = "SELECT (Codigo & ' - ' & Cuenta) As Cuentas ,* " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'C' " _
       & "AND DG = 'D' " _
       & "ORDER BY DG "
  SelectDBCombo DCCXC, AdoCatalogo, sSQL, "Cuentas"
  With AdoCatalogo.Recordset
   If .RecordCount > 0 Then Rf = 1 Else Rf = 0
  End With
End Sub

Public Sub Carga_Datos_Iniciales(MBFecha As MaskEdBox, EsNuevo As Boolean)
  Ln_No = 0
  TxtAño.Text = Year(date)
    
  'Enceramos el espacio de cada usuario para emprezar con una nueva retencion
  sSQL = "DELETE * " _
       & "FROM Asiento_Exportaciones " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  ConectarAdoExecute sSQL
   
  'Carga la Tabla de Tipo Importacion al DataCombo
  Carga_TipoImportacion
   
  'Carga la Tabla de Distrito al DataCombo
  Carga_Distrito
   
  'Carga la Tabla de Regimen al DataCombo
  Carga_Regimen
   
  'Carga la Tabla de Clientes al DataCombo
  Leer_Clientes
    
  'Carga los Tipos de Comprobantes
  Carga_TipoComprobante
    
  'Carga solo la Factura en la sección de Factura de Exportaciones
  Carga_TipoComprobanteSF
    
  'Carga solo las cuentas por Cobrar
  Carga_CXC
   
  CTP.Clear
  CTP.AddItem "CE"
  CTP.AddItem "CI"
  CTP.AddItem "CD"
  CTP.Text = "CE"
   
  'Si es Nuevo ingresa por aqui
  FrmRetencion.Visible = True
  ' LblMensaje.Visible = False
  If EsNuevo Then
     'Si todas las variables tienen cero despliego mensaje y no cargo nada
     'No hay cuentas
     If Rf = 0 Then
        FrmRetencion.Visible = False
        ' LblMensaje.Visible = True
        CTP.SetFocus
     Else
        DCCXC.SetFocus
     End If
  Else
     'Si es Modificación viene por aca
     CMes.Visible = True
     CAño.Visible = True
     Modificacion_AT "E", CModificacion, CMes, CAño
     CMes.Text = MesesLetras(Month(FechaSistema))
     CAño.SetFocus
     CAño.Text = Year(FechaSistema)
     CModificacion.Visible = True
  End If
End Sub

Public Sub Grabacion()
    'Capturo el Tipo de Exportacion
    Captura_TipoExportacion

   'Selecciona el numero mayor para continuar la secuencia en el
   'campo T_No y A_No
   'Grabo en el Asiento_Importacioness e implicito Asiento_Air
    SetAdoAddNew "Asiento_Exportaciones"
    SetAdoFields "Codigo", CodCat
    SetAdoFields "CtasxCobrar", CuentaC
    SetAdoFields "ExportacionDe", CodImp
    SetAdoFields "TipoComprobante", cod
    SetAdoFields "FechaEmbarque", MBFechaEmbarque
    SetAdoFields "IdFiscalProv", CodigoCliente
    SetAdoFields "ValorFOB", TxtValorFOB
    SetAdoFields "DevIva", ValorP
    SetAdoFields "FacturaExportacion", 1
    SetAdoFields "ValorFOBComprobante", TxtValorFOBC
    SetAdoFields "Establecimiento", TxtNumSerieUnoComp
    SetAdoFields "PuntoEmision", TxtNumSerieDosComp
    SetAdoFields "Secuencial", CTNumero(TxtNumSerieTresComp)
    SetAdoFields "Autorizacion", TxtNumAutComp
    SetAdoFields "FechaEmision", MBFechaEmiComp
    SetAdoFields "FechaRegistro", MBFechaRegistro
    If DCExportacion = "Bienes" Then
       SetAdoFields "NumeroDctoTransporte", TxtNumDcto
       SetAdoFields "DistAduanero", DCDistrito
       SetAdoFields "Anio", TxtAño
       SetAdoFields "Regimen", DCRegimen
       SetAdoFields "Correlativo", TxtCorrelativo
       SetAdoFields "Verificador", TxtVerificador
    Else
       SetAdoFields "NumeroDctoTransporte", 0
       SetAdoFields "DistAduanero", 0
       SetAdoFields "Anio", 0
       SetAdoFields "Regimen", 0
       SetAdoFields "Correlativo", 0
       SetAdoFields "Verificador", 0
    End If
    SetAdoFields "A_No", 1
    SetAdoFields "T_No", Trans_No
    SetAdoUpdate
      
    'Grabamos los datos de la transaccion en la tabla definitiva de almacenamiento
    ID_Trans = Maximo_De("Trans_Exportaciones", "ID")  'va a tener el indice de transaccion unico para que no exista duplicados en a base
    sSQL = "SELECT * " _
         & "FROM Asiento_Exportaciones " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND T_No = " & Trans_No & " " _
         & "ORDER BY T_No DESC "
    SelectAdodc AdoAsientoExportaciones, sSQL
    With AdoAsientoExportaciones.Recordset
     If .RecordCount > 0 Then
         SetAdoAddNew "Trans_Exportaciones"
         SetAdoFields "T", Normal
         SetAdoFields "Codigo", .Fields("Codigo")
         SetAdoFields "CtasxCobrar", .Fields("CtasxCobrar")
         SetAdoFields "ExportacionDe", .Fields("ExportacionDe")
         SetAdoFields "TipoComprobante", .Fields("TipoComprobante")
         SetAdoFields "FechaEmbarque", .Fields("FechaEmbarque")
         SetAdoFields "NumeroDctoTransporte", .Fields("NumeroDctoTransporte")
         SetAdoFields "IdFiscalProv", .Fields("IdFiscalProv")
         SetAdoFields "ValorFOB", .Fields("ValorFOB")
         SetAdoFields "DevIva", .Fields("DevIva")
         SetAdoFields "FacturaExportacion", .Fields("FacturaExportacion")
         SetAdoFields "ValorFOBComprobante", .Fields("ValorFOBComprobante")
         SetAdoFields "DistAduanero", .Fields("DistAduanero")
         SetAdoFields "Anio", .Fields("Anio")
         SetAdoFields "Regimen", .Fields("Regimen")
         SetAdoFields "Correlativo", .Fields("Correlativo")
         SetAdoFields "Verificador", .Fields("Verificador")
         SetAdoFields "Establecimiento", .Fields("Establecimiento")
         SetAdoFields "PuntoEmision", .Fields("PuntoEmision")
         SetAdoFields "Secuencial", .Fields("Secuencial")
         SetAdoFields "Autorizacion", .Fields("Autorizacion")
         SetAdoFields "FechaEmision", .Fields("FechaEmision")
         SetAdoFields "FechaRegistro", .Fields("FechaRegistro")
         SetAdoFields "A_No", 1
         SetAdoFields "T_No", Trans_No
         SetAdoFields "TP", CTP
         SetAdoFields "Numero", TxtNumeroC
         SetAdoFields "Fecha", .Fields("FechaRegistro")
         SetAdoFields "ID", ID_Trans
         SetAdoFields "Linea_SRI", 0
         SetAdoUpdate
     End If
    End With
End Sub

Public Sub Habilita_Controles()
  'Habilito los controles para la modificacion
  CModificacion.Enabled = True
  Frame1.Enabled = True
  DCProveedor.Enabled = True
  CmdGrabar.Enabled = True
  FrmTipoComprob.Enabled = True
  FrmRetencion.Enabled = True
  ' Label23.Visible = True
  CMes.Visible = True
  CAño.Visible = True
  ' Label29.Visible = True
End Sub

Public Sub Deshabilita_Controles()
  'Deshabilito los controles para la modificacion
  CModificacion.Enabled = False
  Frame1.Enabled = False
  DCProveedor.Enabled = False
  CmdGrabar.Enabled = False
  FrmTipoComprob.Enabled = False
  FrmRetencion.Enabled = False
  ' Label23.Visible = False
  CMes.Visible = False
  CAño.Visible = False
  ' Label29.Visible = False
End Sub

Private Sub TxtVerificador_LostFocus()
  'Valido que sea ingresado 1 caracter
  Longitud = Len(TxtVerificador)
  If CInt(Longitud) < 1 Then
     MsgBox "El Verificador consta de 1 caracter", vbInformation, "Aviso"
     TxtVerificador = ""
     TxtVerificador.SetFocus
  End If
End Sub
