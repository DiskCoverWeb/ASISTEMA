VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FExportacionesAT 
   BackColor       =   &H00C0C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EXPORTACIONES"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10410
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FExportacionesAT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   10410
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrmRetencion 
      BackColor       =   &H00C0C000&
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
      Height          =   855
      Left            =   105
      TabIndex        =   0
      Top             =   0
      Width           =   8100
      Begin MSDataListLib.DataCombo DCCXC 
         Bindings        =   "FExportacionesAT.frx":0696
         DataSource      =   "AdoCatalogo"
         Height          =   345
         Left            =   1470
         TabIndex        =   2
         Top             =   315
         Width           =   6525
         _ExtentX        =   11509
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
      Begin MSForms.Label Label9 
         Height          =   225
         Left            =   210
         TabIndex        =   1
         Top             =   315
         Width           =   1065
         BackColor       =   12632064
         Caption         =   "Cuenta Por Cobrar"
         Size            =   "1879;397"
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3480
      Left            =   105
      TabIndex        =   7
      Top             =   1575
      Width           =   10200
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
         Height          =   330
         Left            =   8532
         TabIndex        =   30
         ToolTipText     =   $"FExportacionesAT.frx":06B0
         Top             =   1188
         Width           =   552
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
         Height          =   330
         Left            =   9180
         TabIndex        =   31
         Top             =   1188
         Value           =   -1  'True
         Width           =   660
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
         Left            =   4752
         MaxLength       =   16
         TabIndex        =   26
         ToolTipText     =   $"FExportacionesAT.frx":073A
         Top             =   1188
         Width           =   2496
      End
      Begin VB.Frame FraExpor 
         Caption         =   "Factura de Exportaciones"
         Height          =   1695
         Left            =   105
         TabIndex        =   32
         Top             =   1680
         Width           =   9990
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
            Left            =   1512
            MaxLength       =   7
            TabIndex        =   40
            Text            =   "0000001"
            ToolTipText     =   $"FExportacionesAT.frx":0801
            Top             =   1080
            Width           =   1200
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
            TabIndex        =   38
            Text            =   "001"
            ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al código del establecimiento"
            Top             =   1080
            Width           =   552
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
            Left            =   864
            MaxLength       =   3
            TabIndex        =   39
            Text            =   "001"
            ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al punto dde emisión"
            Top             =   1080
            Width           =   552
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
            Left            =   3024
            MaxLength       =   10
            TabIndex        =   42
            ToolTipText     =   $"FExportacionesAT.frx":08A4
            Top             =   1080
            Width           =   1416
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
            Left            =   6588
            TabIndex        =   36
            Text            =   "0.00"
            ToolTipText     =   "En este campo se debe ingresar el valor FOB correspondiente a la Factura de exportación"
            Top             =   432
            Width           =   1200
         End
         Begin MSDataListLib.DataCombo DCFacturaExpor 
            Bindings        =   "FExportacionesAT.frx":0930
            DataSource      =   "AdoAux"
            Height          =   288
            Left            =   216
            TabIndex        =   34
            ToolTipText     =   $"FExportacionesAT.frx":0945
            Top             =   432
            Width           =   4872
            _ExtentX        =   8599
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
            Height          =   336
            Left            =   6588
            TabIndex        =   44
            ToolTipText     =   $"FExportacionesAT.frx":0A07
            Top             =   1080
            Width           =   1200
            _ExtentX        =   2117
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
            Height          =   336
            Left            =   8208
            TabIndex        =   46
            ToolTipText     =   $"FExportacionesAT.frx":0AB3
            Top             =   1080
            Width           =   1200
            _ExtentX        =   2117
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
         Begin MSForms.Label Label30 
            Height          =   228
            Left            =   216
            TabIndex        =   33
            Top             =   216
            Width           =   1788
            Caption         =   "Tipo de Comprobante"
            Size            =   "3154;402"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label31 
            Height          =   228
            Left            =   216
            TabIndex        =   37
            Top             =   864
            Width           =   2496
            Caption         =   "No. Comprobante Modificado"
            Size            =   "4403;402"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label32 
            Height          =   228
            Left            =   6588
            TabIndex        =   43
            Top             =   864
            Width           =   1308
            Caption         =   "Fecha Emisión"
            Size            =   "2307;402"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label33 
            Height          =   228
            Left            =   3024
            TabIndex        =   41
            Top             =   864
            Width           =   1416
            Caption         =   "No. Autorización"
            Size            =   "2498;402"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label5 
            Height          =   228
            Left            =   6588
            TabIndex        =   35
            Top             =   216
            Width           =   2388
            Caption         =   "Valor FOB del Comprobante"
            Size            =   "4212;402"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label8 
            Height          =   228
            Left            =   8208
            TabIndex        =   45
            Top             =   864
            Width           =   1308
            Caption         =   "Fecha Registro"
            Size            =   "2307;402"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
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
            Left            =   3456
            MaxLength       =   1
            TabIndex        =   24
            ToolTipText     =   "Verificador (1 caracter)"
            Top             =   432
            Width           =   768
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
            Left            =   2592
            MaxLength       =   6
            TabIndex        =   22
            ToolTipText     =   "Correlativo (6 caracteres)"
            Top             =   432
            Width           =   768
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
            Left            =   972
            MaxLength       =   4
            TabIndex        =   18
            ToolTipText     =   "Años (4 caracteres)"
            Top             =   432
            Width           =   660
         End
         Begin MSDataListLib.DataCombo DCRegimen 
            Bindings        =   "FExportacionesAT.frx":0B3B
            DataSource      =   "AdoRegimen"
            Height          =   288
            Left            =   1728
            TabIndex        =   20
            ToolTipText     =   "Regimen (2 caracteres)"
            Top             =   432
            Width           =   768
            _ExtentX        =   1349
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
         Begin MSDataListLib.DataCombo DCDistrito 
            Bindings        =   "FExportacionesAT.frx":0B54
            DataSource      =   "AdoDistrito"
            Height          =   288
            Left            =   108
            TabIndex        =   16
            ToolTipText     =   "Distrito aduanero (3 caracteres)"
            Top             =   432
            Width           =   768
            _ExtentX        =   1349
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
         Begin MSForms.Label Label13 
            Height          =   228
            Left            =   3468
            TabIndex        =   23
            Top             =   216
            Width           =   852
            Caption         =   "Verificador"
            Size            =   "1503;402"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label10 
            Height          =   228
            Left            =   2520
            TabIndex        =   21
            Top             =   216
            Width           =   948
            Caption         =   "Correlativo"
            Size            =   "1672;402"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label6 
            Height          =   228
            Left            =   1680
            TabIndex        =   19
            Top             =   216
            Width           =   744
            Caption         =   "Regimen"
            Size            =   "1312;402"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label4 
            Height          =   228
            Left            =   948
            TabIndex        =   17
            Top             =   216
            Width           =   420
            Caption         =   "Año"
            Size            =   "741;402"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label7 
            Height          =   228
            Left            =   108
            TabIndex        =   15
            Top             =   216
            Width           =   636
            Caption         =   "Distrito"
            Size            =   "1122;402"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
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
         Left            =   7236
         TabIndex        =   28
         Text            =   "0.00"
         ToolTipText     =   $"FExportacionesAT.frx":0B6E
         Top             =   1188
         Width           =   1200
      End
      Begin MSDataListLib.DataCombo DCExportacion 
         Bindings        =   "FExportacionesAT.frx":0C3A
         DataSource      =   "AdoExportacion"
         Height          =   288
         Left            =   108
         TabIndex        =   9
         ToolTipText     =   "Corresponde al tipo de transacción realizada (Bienes o Servicios), si es Exportación de Bienes o Ingresos del Exterior"
         Top             =   432
         Width           =   2928
         _ExtentX        =   5159
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
         Bindings        =   "FExportacionesAT.frx":0C57
         DataSource      =   "AdoTipoComprobante"
         Height          =   288
         Left            =   3024
         TabIndex        =   11
         ToolTipText     =   "En este combo de selección se desplegará una lista de comprobantes que sustentas la exportación"
         Top             =   432
         Width           =   4980
         _ExtentX        =   8784
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
      Begin MSForms.Label Label18 
         Height          =   228
         Left            =   8532
         TabIndex        =   29
         Top             =   972
         Width           =   1416
         Caption         =   "Devolución I.V.A."
         Size            =   "2498;402"
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   228
         Left            =   4752
         TabIndex        =   25
         Top             =   972
         Width           =   2172
         Caption         =   "No. Documento Transporte"
         Size            =   "3831;402"
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label2 
         Height          =   228
         Left            =   108
         TabIndex        =   8
         Top             =   216
         Width           =   1308
         Caption         =   "Exportación de:"
         Size            =   "2307;402"
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label15 
         Height          =   228
         Left            =   8088
         TabIndex        =   12
         Top             =   216
         Width           =   1800
         Caption         =   "Fechas de Embarque:"
         Size            =   "3175;402"
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label23 
         Height          =   228
         Left            =   7236
         TabIndex        =   27
         Top             =   972
         Width           =   984
         Caption         =   "Valor FOB"
         Size            =   "1736;402"
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label3 
         Height          =   228
         Left            =   3048
         TabIndex        =   10
         Top             =   216
         Width           =   1476
         Caption         =   "Tipo Comprobante"
         Size            =   "2603;402"
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.CommandButton CmdGrabar 
      BackColor       =   &H00C0C000&
      Caption         =   "&Aceptar"
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
      Left            =   8295
      Picture         =   "FExportacionesAT.frx":0C78
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "Grabar"
      Top             =   105
      Width           =   960
   End
   Begin VB.CommandButton CmdCerrar 
      BackColor       =   &H00C0C000&
      Caption         =   "&Cancelar"
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
      Left            =   9345
      Picture         =   "FExportacionesAT.frx":0F82
      Style           =   1  'Graphical
      TabIndex        =   48
      ToolTipText     =   "Salir"
      Top             =   105
      Width           =   960
   End
   Begin MSAdodcLib.Adodc AdoTipoComprobante 
      Height          =   330
      Left            =   315
      Top             =   2520
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
      Top             =   3465
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
      Top             =   3780
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
      Top             =   1890
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
      Top             =   2205
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
      Top             =   3150
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
      Top             =   2835
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
      Top             =   4095
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
      Top             =   4410
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
      Top             =   4725
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
   Begin MSAdodcLib.Adodc AdoAsientos 
      Height          =   330
      Left            =   2940
      Top             =   1890
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
      Caption         =   "Asientos"
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
   Begin VB.Label LblProveedor 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cliente/Proveedor"
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
      TabIndex        =   49
      Top             =   1155
      Width           =   7890
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
      Left            =   8400
      TabIndex        =   6
      Top             =   1155
      Width           =   1905
   End
   Begin MSForms.Label Label17 
      Height          =   225
      Left            =   8400
      TabIndex        =   5
      Top             =   945
      Width           =   1800
      BackColor       =   12632064
      Caption         =   "No. de Identificación"
      Size            =   "3175;397"
      FontName        =   "Arial"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label41 
      Height          =   225
      Left            =   105
      TabIndex        =   3
      Top             =   945
      Width           =   960
      BackColor       =   12632064
      Caption         =   "Proveedor"
      Size            =   "1693;397"
      FontName        =   "Arial"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
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
      Left            =   8085
      TabIndex        =   4
      Top             =   1155
      Width           =   330
   End
End
Attribute VB_Name = "FExportacionesAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MBFecha As MaskEdBox
Dim Rf, cod, x, CodImp, CodTs, Longitud As Byte
Dim PC As Boolean
Dim SumAnio, Aniocad, AniocadAux, CodPorIva, CodPorIce, CodRetBien, CodRetServ, CodReg As Integer
Dim CalmIva, CalmIce, CalIbMi, CalIsMi, ac As Double
Dim Cap, Cap1, Ch, ValorP, CodProv, CodDis, Opc As String
Dim Espizq, Espder, Captc, PorIva, PorIce, TipoImp, CodCat, CuentaC As String

Private Sub CmdCerrar_Click()
  'Borra Asiento Compras
   sSQL = "DELETE * " _
        & "FROM Asiento_Exportaciones " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND CodigoU = '" & CodigoUsuario & "' " _
        & "AND T_No = " & Trans_No & " "
   ConectarAdoExecute sSQL
   Unload FExportacionesAT
End Sub

Private Sub CmdGrabar_Click()
    RatonReloj
   'Valido las fechas
    FechaValida MBFechaEmiComp
    FechaValida MBFechaRegistro
    FechaValida MBFechaEmbarque
   'Pregunto antes de grabar
   'Grabacion de los Datos
    Grabacion
    
    
    Unload FExportacionesAT
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
    TxtAño.Text = Year(Date)
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
   Carga_Datos_Iniciales MBFecha, Nuevo
   LblTD.Caption = TipoBenef                  ' Tipo de Cliente: C,R,P,O
   LblNumIdent = CICliente                    ' CI o RUC del Cliente
   LblProveedor.Caption = " " & NombreCliente ' Nombre del Cliente
   MBFechaEmbarque = FechaComp
   MBFechaEmiComp = FechaComp
   MBFechaRegistro = FechaComp
  'CodigoCliente
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
End Sub

Private Sub Form_Load()
    CentrarForm FExportacionesAT
    ConectarAdodc AdoSustento
    ConectarAdodc AdoCatalogo
    ConectarAdodc AdoTipoComprobante
    ConectarAdodc AdoAux
    ConectarAdodc AdoExportacion
    ConectarAdodc AdoDistrito
    ConectarAdodc AdoRegimen
    ConectarAdodc AdoClientes
    ConectarAdodc AdoAsientos
    ConectarAdodc AdoAsientoExportaciones
    ConectarAdodc AdoTransExportaciones
End Sub

Public Sub Carga_TipoImportacion()
' Carga la Tabla de Importaciones al DataCombo
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
' Carga la Tabla de Regimen al DataCombo
    sSQL = "SELECT * " _
         & "FROM Tabla_Regimen " _
         & "WHERE Codigo <> 200 " _
         & "ORDER BY Codigo "
    SelectDBCombo DCRegimen, AdoRegimen, sSQL, "Codigo"
End Sub

Public Sub Captura_TipoComprobante()
    'Captura lo que tiene el Combo de Tipo de Comprobante
    Captc = SinEspaciosIzq(DCTipoComprobante.Text)
    Cap1 = Trim(DCTipoComprobante.Text)
     
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
    'CuentaC = SinEspaciosDer(DCCXC)
    CuentaC = SinEspaciosIzq(DCCXC)
    'MsgBox CuentaC
    'Busca que sea igual a la Descripcion
    With AdoCatalogo.Recordset
     If .RecordCount > 0 Then
        .MoveFirst
        .Find ("Codigo = '" & CuentaC & "' ")
         If Not .EOF Then
            CodCat = .Fields("Codigo")
         Else
            MsgBox "Vuelva a seleccionar", vbInformation, "Aviso"
            DCCXC.SetFocus
         End If
     End If
    End With
End Sub

Public Sub Limpiar_Controles()
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
    If OpcNo.Value = True Then ValorP = "N"
End Sub

Private Sub OpcSi_LostFocus()
    If OpcSi.Value = True Then ValorP = "S"
End Sub

Private Sub TxtAño_GotFocus()
    MarcarTexto TxtAño
End Sub

Private Sub TxtAño_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtAño_LostFocus()
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
    sSQL = "SELECT (Codigo & ' - ' & Cuenta) As Cuentas,* " _
         & "FROM Catalogo_Cuentas " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TC = 'C' " _
         & "AND DG = 'D' " _
         & "ORDER BY DG "
    SelectDBCombo DCCXC, AdoCatalogo, sSQL, "Cuentas"
    With AdoCatalogo.Recordset
      If AdoCatalogo.Recordset.RecordCount > 0 Then Rf = 1 Else Rf = 0
    End With
End Sub

Public Sub Carga_Datos_Iniciales(MBFecha As MaskEdBox, EsNuevo As Boolean)
    CodPorIva = 0
    CodPorIce = "0"
    CodRetBien = 0
    CodRetServ = 0

    TxtAño.Text = Year(Date)
   'Enceramos el espacio de cada usuario para emprezar con una nueva retencion
    sSQL = "DELETE * " _
         & "FROM Asiento_Exportaciones " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND T_No = " & Trans_No & " "
    ConectarAdoExecute sSQL
   
   ' Carga la Tabla de Tipo Importacion al DataCombo
    Carga_TipoImportacion
   
   ' Carga la Tabla de Distrito al DataCombo
    Carga_Distrito
   
   ' Carga la Tabla de Regimen al DataCombo
    Carga_Regimen
   
    'Carga los Tipos de Comprobantes
    Carga_TipoComprobante
    
    'Carga solo la Factura en la sección de Factura de Exportaciones
    Carga_TipoComprobanteSF
    
    'Carga solo las cuentas por Cobrar
    Carga_CXC
       
   'Si es Nuevo ingresa por aqui
    FrmRetencion.Visible = True
    If EsNuevo Then
      'Si todas las variables tienen cero despliego mensaje y no cargo nada
      'No hay cuentas
       If Rf = 0 Then
          FrmRetencion.Visible = False
       Else
          DCCXC.SetFocus
       End If
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
    SetAdoFields "Secuencial", Convertir_Numero(TxtNumSerieTresComp)
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
End Sub

Public Sub Habilita_Controles()
    'Habilito los controles para la modificacion
    Frame1.Enabled = True
    CmdGrabar.Enabled = True
    FrmRetencion.Enabled = True
    Label23.Visible = True
End Sub

Public Sub Deshabilita_Controles()
    'Deshabilito los controles para la modificacion
    Frame1.Enabled = False
    CmdGrabar.Enabled = False
    FrmRetencion.Enabled = False
    Label23.Visible = False
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
