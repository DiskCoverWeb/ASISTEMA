VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FImportaciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IMPORTACIONES"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10635
   Icon            =   "FImportaciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   10635
   Begin TabDlg.SSTab SSTImportaciones 
      Height          =   4320
      Left            =   105
      TabIndex        =   10
      Top             =   1050
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   7620
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Importaciones"
      TabPicture(0)   =   "FImportaciones.frx":0696
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DCTipoComprobante"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "MBFechaLiquida"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "DCImportacion"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "DCSustento"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "CmdAir"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "TxtValorCIF"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "FraRefrendo"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "FraBases"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Conceptos AIR"
      TabPicture(1)   =   "FImportaciones.frx":06B2
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FraRetencion"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame FraBases 
         Caption         =   "BASES IMPONIBLES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   105
         TabIndex        =   33
         Top             =   2415
         Width           =   10200
         Begin VB.TextBox TxtMontoIva 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   4320
            TabIndex        =   38
            Text            =   "0.00"
            ToolTipText     =   $"FImportaciones.frx":06CE
            Top             =   432
            Width           =   1308
         End
         Begin VB.TextBox TxtMontoIce 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   336
            Left            =   8748
            TabIndex        =   37
            Top             =   432
            Width           =   1308
         End
         Begin VB.TextBox TxtBaseImpo 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   105
            MaxLength       =   14
            MultiLine       =   -1  'True
            TabIndex        =   36
            Text            =   "FImportaciones.frx":07B8
            ToolTipText     =   "Corresponde al valor de la importación gravada con tarifa 0% o exento"
            Top             =   432
            Width           =   1416
         End
         Begin VB.TextBox TxtBaseImpoGrav 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   1728
            MultiLine       =   -1  'True
            TabIndex        =   35
            Text            =   "FImportaciones.frx":07BF
            ToolTipText     =   "Corresponde al valor de la importación gravada con tarifa 12% en el momento de la desaduanización"
            Top             =   432
            Width           =   1416
         End
         Begin VB.TextBox TxtBaseImpoIce 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   6156
            MultiLine       =   -1  'True
            TabIndex        =   34
            Text            =   "FImportaciones.frx":07C6
            ToolTipText     =   $"FImportaciones.frx":07CB
            Top             =   432
            Width           =   1416
         End
         Begin MSDataListLib.DataCombo DCPorcenIva 
            Bindings        =   "FImportaciones.frx":086B
            DataSource      =   "AdoPorIva"
            Height          =   288
            Left            =   3240
            TabIndex        =   39
            ToolTipText     =   "Este campo corresponde al porentaje de IVA vigente a la fecha de desaduanización"
            Top             =   432
            Width           =   984
            _ExtentX        =   1746
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DCPorcenIce 
            Bindings        =   "FImportaciones.frx":0883
            DataSource      =   "AdoPorIce"
            Height          =   288
            Left            =   7668
            TabIndex        =   40
            ToolTipText     =   "Este campo corresponde al porcentaje de ICE vigente a lafecha de emisión de la importación"
            Top             =   432
            Width           =   984
            _ExtentX        =   1746
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
      End
      Begin VB.Frame FraRefrendo 
         Caption         =   "Refrendo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   855
         Left            =   5145
         TabIndex        =   27
         Top             =   1575
         Width           =   5160
         Begin VB.TextBox TxtAño 
            Height          =   330
            Left            =   1050
            MaxLength       =   4
            TabIndex        =   30
            ToolTipText     =   "Años (4 caracteres)"
            Top             =   420
            Width           =   645
         End
         Begin VB.TextBox TxtCorrelativo 
            Height          =   330
            Left            =   2835
            MaxLength       =   6
            TabIndex        =   29
            Text            =   "000001"
            ToolTipText     =   "Correlativo (6 caracteres)"
            Top             =   420
            Width           =   750
         End
         Begin VB.TextBox TxtVerificador 
            Height          =   330
            Left            =   3990
            MaxLength       =   1
            TabIndex        =   28
            ToolTipText     =   "Verificador (1 caracter)"
            Top             =   420
            Width           =   750
         End
         Begin MSDataListLib.DataCombo DCRegimen 
            Bindings        =   "FImportaciones.frx":089B
            DataSource      =   "AdoRegimen"
            Height          =   315
            Left            =   1890
            TabIndex        =   31
            ToolTipText     =   "Regimen (2 caracteres)"
            Top             =   420
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DCDistrito 
            Bindings        =   "FImportaciones.frx":08B4
            DataSource      =   "AdoDistrito"
            Height          =   315
            Left            =   105
            TabIndex        =   32
            ToolTipText     =   "Distrito aduanero (3 caracteres)"
            Top             =   420
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
      End
      Begin VB.TextBox TxtValorCIF 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   9135
         TabIndex        =   26
         Text            =   "0.00"
         ToolTipText     =   "En este campo obligatorio se debe ingresar el valor CIF de los bienes importados o el valor del pago efectuado al exterior"
         Top             =   1260
         Width           =   1170
      End
      Begin VB.Frame FraRetencion 
         Caption         =   "INGRESE LOS DATOS DE LA RETENCION _________________________________________ FORMULARIO 103"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   3924
         Left            =   -74895
         TabIndex        =   12
         Top             =   312
         Width           =   10155
         Begin VB.CheckBox ChRetF 
            Caption         =   "Retención en la Fuente"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   105
            TabIndex        =   22
            Top             =   315
            Width           =   2328
         End
         Begin VB.TextBox TxtValConA 
            Enabled         =   0   'False
            Height          =   336
            Left            =   8748
            TabIndex        =   21
            Top             =   1512
            Width           =   1308
         End
         Begin VB.TextBox TxtPorRetConA 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   336
            Left            =   7992
            TabIndex        =   20
            Top             =   1512
            Width           =   660
         End
         Begin VB.TextBox TxtTotalReten 
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
            Height          =   336
            Left            =   8640
            TabIndex        =   19
            Top             =   3132
            Width           =   1308
         End
         Begin VB.TextBox TxtSumatoria 
            Alignment       =   1  'Right Justify
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
            Height          =   336
            Left            =   8100
            MultiLine       =   -1  'True
            TabIndex        =   18
            Top             =   756
            Width           =   1956
         End
         Begin VB.TextBox TxtBimpConA 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   336
            Left            =   6480
            TabIndex        =   17
            Top             =   1512
            Width           =   1416
         End
         Begin VB.TextBox TxtNumUnoAutComRet 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   330
            Left            =   2268
            MaxLength       =   10
            MultiLine       =   -1  'True
            TabIndex        =   16
            Top             =   864
            Width           =   1632
         End
         Begin VB.TextBox TxtNumTresComRet 
            Enabled         =   0   'False
            Height          =   336
            Left            =   1404
            MaxLength       =   7
            TabIndex        =   15
            Text            =   "0000001"
            Top             =   864
            Width           =   876
         End
         Begin VB.TextBox TxtNumDosComRet 
            Enabled         =   0   'False
            Height          =   336
            Left            =   756
            MaxLength       =   3
            TabIndex        =   14
            Text            =   "001"
            Top             =   864
            Width           =   552
         End
         Begin VB.TextBox TxtNumUnoComRet 
            Enabled         =   0   'False
            Height          =   336
            Left            =   105
            MaxLength       =   3
            TabIndex        =   13
            Text            =   "001"
            Top             =   864
            Width           =   552
         End
         Begin MSDataListLib.DataCombo DCConceptoRet 
            Bindings        =   "FImportaciones.frx":08CE
            DataSource      =   "AdoConceptoRet"
            Height          =   288
            Left            =   108
            TabIndex        =   23
            Top             =   1512
            Width           =   6384
            _ExtentX        =   11271
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
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
         Begin MSDataGridLib.DataGrid DGConceptoAir 
            Bindings        =   "FImportaciones.frx":08EB
            Height          =   1176
            Left            =   108
            TabIndex        =   24
            Top             =   1896
            Width           =   9888
            _ExtentX        =   17436
            _ExtentY        =   2064
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   1
            RowHeight       =   19
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Datos Ingresados"
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
                  LCID            =   3082
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
                  LCID            =   3082
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
         Begin MSDataListLib.DataCombo DCRetFuente 
            Bindings        =   "FImportaciones.frx":0907
            DataSource      =   "AdoRetFuente"
            Height          =   315
            Left            =   2520
            TabIndex        =   25
            Top             =   315
            Visible         =   0   'False
            Width           =   7470
            _ExtentX        =   13176
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
      End
      Begin VB.CommandButton CmdAir 
         Caption         =   "AIR"
         Height          =   444
         Left            =   9180
         Picture         =   "FImportaciones.frx":0922
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Se ubica en la pestaña de Retenciones"
         Top             =   3456
         Width           =   768
      End
      Begin MSDataListLib.DataCombo DCSustento 
         Bindings        =   "FImportaciones.frx":0E48
         DataSource      =   "AdoSustento"
         Height          =   288
         Left            =   108
         TabIndex        =   41
         ToolTipText     =   "En este combo de selección se despliega una lista de tipos de sustentos tributarios."
         Top             =   648
         Width           =   10272
         _ExtentX        =   18124
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
      Begin MSDataListLib.DataCombo DCImportacion 
         Bindings        =   "FImportaciones.frx":0E62
         DataSource      =   "AdoImportacion"
         Height          =   288
         Left            =   108
         TabIndex        =   42
         ToolTipText     =   $"FImportaciones.frx":0E7F
         Top             =   1296
         Width           =   5088
         _ExtentX        =   8969
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
      Begin MSMask.MaskEdBox MBFechaLiquida 
         Height          =   336
         Left            =   6696
         TabIndex        =   43
         ToolTipText     =   "En este campo se ingresa la fecha de liquidación del comprobante en el Banco"
         Top             =   1296
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   609
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
         Bindings        =   "FImportaciones.frx":0F5B
         DataSource      =   "AdoTipoComp"
         Height          =   288
         Left            =   108
         TabIndex        =   44
         ToolTipText     =   "Corresponde al tipo de comprobante utilizado en la transacción"
         Top             =   1944
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
      ItemData        =   "FImportaciones.frx":0F7C
      Left            =   525
      List            =   "FImportaciones.frx":0F7E
      TabIndex        =   9
      Text            =   "2000"
      Top             =   0
      Visible         =   0   'False
      Width           =   885
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
      Width           =   1416
   End
   Begin VB.ComboBox CModificacion 
      DataSource      =   "AdoAux"
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
      Width           =   7215
   End
   Begin VB.ComboBox CTP 
      Height          =   288
      Left            =   8856
      TabIndex        =   5
      ToolTipText     =   "En este combo se despliega una lista con lo stipos de comprobantes existentes tales como: Comprobante Diario, Ingreso o Egreso"
      Top             =   648
      Width           =   660
   End
   Begin VB.TextBox TxtNumeroC 
      Alignment       =   1  'Right Justify
      Height          =   336
      Left            =   9504
      MaxLength       =   7
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "FImportaciones.frx":0F80
      ToolTipText     =   "En este campo se debe ingresar el número del comprobante, el cual no excedera los siete caracteres"
      Top             =   648
      Width           =   984
   End
   Begin VB.CommandButton CmdCerrar 
      Caption         =   "&Salir"
      Height          =   750
      Left            =   1260
      Picture         =   "FImportaciones.frx":0F84
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Salir"
      Top             =   4515
      Width           =   960
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Height          =   750
      Left            =   210
      Picture         =   "FImportaciones.frx":13C6
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Grabar"
      Top             =   4515
      Width           =   960
   End
   Begin MSAdodcLib.Adodc AdoTipoComprobante 
      Height          =   330
      Left            =   210
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
   Begin MSAdodcLib.Adodc AdoTransImportaciones 
      Height          =   330
      Left            =   210
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
      Caption         =   "TransCompras"
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
   Begin MSAdodcLib.Adodc AdoAsientoAir 
      Height          =   330
      Left            =   2940
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
      Caption         =   "AsientoAir"
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
   Begin MSAdodcLib.Adodc AdoAsientoImportaciones 
      Height          =   330
      Left            =   2940
      Top             =   1365
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
      Caption         =   "AsientoImportaciones"
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
   Begin MSAdodcLib.Adodc AdoRetFuente 
      Height          =   330
      Left            =   210
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
      Caption         =   "RetencionFuente"
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
   Begin MSAdodcLib.Adodc AdoTransAir 
      Height          =   330
      Left            =   2940
      Top             =   1680
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
      Caption         =   "TransAir"
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
      Left            =   210
      Top             =   1365
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
   Begin MSAdodcLib.Adodc AdoImportacion 
      Height          =   330
      Left            =   210
      Top             =   1680
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
      Left            =   210
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
      Left            =   210
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
   Begin MSAdodcLib.Adodc AdoConceptoret 
      Height          =   330
      Left            =   210
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
      Caption         =   "AdoConceptoRet"
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
   Begin MSAdodcLib.Adodc AdoPorIva 
      Height          =   330
      Left            =   2940
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
      Caption         =   "AdoPorIva"
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
   Begin MSAdodcLib.Adodc AdoPorIce 
      Height          =   330
      Left            =   2940
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
      Caption         =   "AdoPorIce"
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
      Left            =   2940
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
   Begin MSDataListLib.DataCombo DCProveedor 
      Bindings        =   "FImportaciones.frx":16D0
      DataSource      =   "AdoClientes"
      Height          =   288
      Left            =   108
      TabIndex        =   2
      ToolTipText     =   "Razón o denomicación Social. Este campo es obligatorio."
      Top             =   648
      Width           =   6492
      _ExtentX        =   11456
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   210
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
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   645
      Left            =   0
      TabIndex        =   45
      Top             =   0
      Width           =   8310
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
      Height          =   336
      Left            =   6912
      TabIndex        =   4
      Top             =   648
      Width           =   1848
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
      Height          =   336
      Left            =   6588
      TabIndex        =   3
      Top             =   648
      Width           =   336
   End
End
Attribute VB_Name = "FImportaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MBFecha As MaskEdBox
Dim OP As Boolean
Dim Espder, Espizq, Cap1, Captc, Cap, CargaTC, CodSus, CodProv As String
Dim Longitud, ac, Rf, CodImp, codTC As Byte
Dim SUM, cal, CalmIva, CalIbMi As Single

Private Sub ChRetF_Click()
  If ChRetF.value <> 0 Then
     DCRetFuente.Visible = True
     TxtNumUnoComRet.Enabled = True
     TxtNumDosComRet.Enabled = True
     TxtNumTresComRet.Enabled = True
     TxtNumUnoAutComRet.Enabled = True
     DCConceptoRet.Enabled = True
     TxtBimpConA.Enabled = True
  Else
     DCRetFuente.Visible = False
     TxtNumUnoComRet.Enabled = False
     TxtNumDosComRet.Enabled = False
     TxtNumTresComRet.Enabled = False
     TxtNumUnoAutComRet.Enabled = False
     DCConceptoRet.Enabled = False
     TxtBimpConA.Enabled = False
  End If
End Sub

Private Sub CmdAir_Click()
  SSTImportaciones.Tab = 1
End Sub

Private Sub CmdCerrar_Click()
  'Borra Asiento Importaciones
  sSQL = "DELETE * " _
       & "FROM Asiento_Importaciones " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  ConectarAdoExecute sSQL
  'Borra Asiento Air
  sSQL = "DELETE * " _
       & "FROM Asiento_Air " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND Tipo_Trans = 'I' "
  ConectarAdoExecute sSQL
  Unload Me
End Sub

Private Sub CmdGrabar_Click()
  RatonReloj
  'Valido por si acaso exista algun valor con 0
  TextoValido TxtBaseImpo, True, , 2
  TextoValido TxtBaseImpoGrav, True, , 2
  TextoValido TxtMontoIva, True, , 2
  TextoValido TxtMontoIce, True, , 2
  TextoValido TxtBaseImpoIce, True, , 2
  FechaValida MBFechaLiquida
  Ln_No = 0
  'Borra si encuentra 2 o mas transacciones iguales
  Eliminar_Trans_AT "I", CodigoCliente, "", "", "", "", TxtCorrelativo, TxtVerificador, "", CMes, CAño, Ln_SRI
  Eliminar_Trans_Air "I", CodigoCliente, CMes, CAño, Ln_SRI
 'Pregunto antes de grabar
  Titulo = "Grabar Importaciones"
  Mensajes = "Desea Grabar los Datos"
  If BoxMensaje = vbYes Then
     'Grabacion de los Datos
     Grabacion
     Titulo = "Grabar Importaciones"
     Mensajes = "Los Datos fueron grabados correctamente" & vbCrLf _
              & "Desea ingresar otra transacción"""
     If BoxMensaje = vbYes Then
        Limpiar_Controles
        SSTImportaciones.Tab = 0
        DCProveedor.SetFocus
     Else
        Unload FImportaciones
     End If
  Else
     DCProveedor.SetFocus
  End If
End Sub

Private Sub CMes_LostFocus()
  Modificacion_AT "I", CModificacion, CMes, CAño
End Sub

Private Sub CModificacion_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub CModificacion_LostFocus()
    Encerar_Var
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
         & "FROM Trans_Importaciones " _
         & "WHERE Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Linea_SRI = " & I & " " _
         & "AND IdFiscalProv = '" & CodigoCliente & "' " _
         & "ORDER BY Linea_SRI "
    SelectAdodc AdoTransImportaciones, sSQL
    With AdoTransImportaciones.Recordset
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
                 'Busco el Sustento Tributario para capturar la descripcion
                  CodSus = .Fields("CodSustento")
                  If AdoSustento.Recordset.RecordCount > 0 Then
                     AdoSustento.Recordset.MoveFirst
                     AdoSustento.Recordset.Find ("Credito_Tributario = '" & CodSus & "' ")
                     If Not AdoSustento.Recordset.EOF Then
                        DCSustento = AdoSustento.Recordset.Fields("Descripcion")
                     Else
                        MsgBox "Este Código no existe", vbInformation, "Aviso"
                     End If
                   End If
                  'Cargo el Comprobante
                   sSQL = "SELECT Tipo_Comprobante_Codigo,Descripcion " _
                        & "FROM Tipo_Comprobante " _
                        & "WHERE Tipo_Comprobante_Codigo <> 0 "
                   SelectAdodc AdoTipoComprobante, sSQL
                   codTC = .Fields("TipoComprobante")
                   If AdoTipoComprobante.Recordset.RecordCount > 0 Then
                      AdoTipoComprobante.Recordset.MoveFirst
                      AdoTipoComprobante.Recordset.Find ("Tipo_Comprobante_Codigo = '" & codTC & "' ")
                      If Not .EOF Then
                         DCTipoComprobante = AdoTipoComprobante.Recordset.Fields("Descripcion")
                      Else
                         MsgBox "El Comprobante no existe", vbInformation, "Aviso"
                      End If
                   End If
                  'Cargo el Codigo de importacion
                   CodImp = .Fields("ImportacionDe") 'Es un número
                   sSQL = "SELECT Concepto, Codigo " _
                       & "FROM Tabla_TipoImportacion " _
                       & "WHERE Concepto <> '.' " _
                       & "ORDER BY Concepto "
                  SelectAdodc AdoImportacion, sSQL
                  If AdoImportacion.Recordset.RecordCount > 0 Then
                     AdoImportacion.Recordset.MoveFirst
                     AdoImportacion.Recordset.Find ("Codigo = " & CodImp & " ")
                     If Not AdoImportacion.Recordset.EOF Then
                        DCImportacion = AdoImportacion.Recordset.Fields("Concepto")
                     Else
                        MsgBox "La Importación no existe", vbInformation, "Aviso"
                     End If
                   End If
                   MBFechaLiquida = .Fields("FechaLiquidacion")
                   Carga_PorcentajeIva MBFechaLiquida
                   'Es un numero el campo PorcentajeIva y lo convierto a String
                   CodPorIva = CStr(.Fields("PorcentajeIva"))
                   If AdoPorIva.Recordset.RecordCount > 0 Then
                      AdoPorIva.Recordset.MoveFirst
                      AdoPorIva.Recordset.Find ("Codigo = '" & CodPorIva & "' ")
                      If Not AdoPorIva.Recordset.EOF Then
                         DCPorcenIva = AdoPorIva.Recordset.Fields("Porc")
                      Else
                         MsgBox "Porcentaje de IVA no existe", vbInformation, "Aviso"
                      End If
                   End If
                   CodPorIce = CStr(.Fields("PorcentajeIce"))
                   If AdoPorIce.Recordset.RecordCount > 0 Then
                      AdoPorIce.Recordset.MoveFirst
                      AdoPorIce.Recordset.Find ("Codigo = '" & CodPorIce & "' ")
                      If Not AdoPorIce.Recordset.EOF Then
                         DCPorcenIce = AdoPorIce.Recordset.Fields("Porc")
                      Else
                         MsgBox "Porcentaje de ICE no existe", vbInformation, "Aviso"
                      End If
                   End If
                   If CodImp = 2 Then
                      FraRefrendo.Enabled = False
                   Else
                      TxtAño = .Fields("Anio")
                      DCDistrito.Text = .Fields("DistAduanero")
                      TxtVerificador = .Fields("Verificador")
                      DCRegimen.Text = .Fields("Regimen")
                      TxtCorrelativo = .Fields("Correlativo")
                   End If
                   TxtNumeroC = .Fields("Numero")
                   TxtValorCIF = .Fields("ValorCIF")
                   TxtBaseImpo = .Fields("BaseImponible")
                   TxtBaseImpoGrav = .Fields("BaseImpGrav")
                   TxtMontoIva = .Fields("MontoIva")
                   TxtBaseImpoIce = .Fields("BaseImpIce")
                   TxtMontoIce = .Fields("MontoIce")
                   CTP.AddItem .Fields("TP")
                   sSQL = "SELECT TA.*, TI.* " _
                        & "FROM Trans_Air As TA, Trans_Importaciones As TI " _
                        & "WHERE TA.Item = '" & NumEmpresa & "' " _
                        & "AND TA.Periodo = '" & Periodo_Contable & "' " _
                        & "AND TA.IdProv = '" & CodProv & "' " _
                        & "AND TA.Linea_SRI = " & I & " " _
                        & "AND TA.IdProv = TI.IdFiscalProv " _
                        & "AND TA.Numero = TI.Numero "
                   SelectDataGrid DGConceptoAir, AdoAsientoAir, sSQL, "Sustento"
                   'Pongo la Base Imponible
                   TxtSumatoria = Val(CCur(TxtBaseImpo)) + Val(CCur(TxtBaseImpoGrav))
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

Private Sub DCConceptoRet_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCConceptoRet_LostFocus()
  If IsNumeric(DCConceptoRet.Text) Then
     MsgBox "No ingrese números. Vuelva a seleccionar.", vbInformation, "Aviso"
     DCConceptoRet.SetFocus
  Else
     With AdoConceptoRet.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
          OP = False
         .Find ("Codigo = '" & SinEspaciosIzq(DCConceptoRet) & "' ")
          Cadena = .Fields("Ingresar_Porcentaje")
          TxtPorRetConA.Enabled = False
          If Not .EOF Then
            'Verifico el codigo para activar el text e ingrese el porcentaje manualmente
            Select Case Cadena
               Case "N": TxtPorRetConA = .Fields("Porcentaje")
                    OP = True
               Case "S": TxtPorRetConA.Enabled = True
            End Select
          Else
             MsgBox "No encontro este código vuelva a buscar"
          End If
      End If
     End With
     TxtBimpConA = TxtSumatoria
  End If
End Sub

Private Sub DCDistrito_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCDistrito_LostFocus()
  TxtAño.Text = Year(date)
End Sub

Private Sub DCImportacion_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCImportacion_LostFocus()
  CmdAir.Enabled = True
  'Desactivo dependiendo del Tipo de Importación
  If (DCImportacion = "Bienes") Then
     FraRefrendo.Enabled = True
     FraBases.Enabled = True
     FraRetencion.Enabled = False
  Else
     FraRefrendo.Enabled = False
     FraBases.Enabled = False
     FraRetencion.Enabled = True
     CmdAir.Enabled = False
  End If
End Sub

Private Sub DCPorcenIce_LostFocus()
  If Not IsNumeric(DCPorcenIce) Then
     MsgBox "No ingrese caracteres. Vuelva a seleccionar.", vbInformation, "Aviso"
     DCPorcenIce = ""
    'Carga_PorcentajeIce
     DCPorcenIce.SetFocus
  Else
    'Busca y captura el codigo de Porcentaje IVA
     CodPorIce = "0"
     With AdoPorIce.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
         .Find ("Porc = " & CSng(DCPorcenIce) & " ")
          If Not .EOF Then CodPorIce = .Fields("Codigo")
       End If
      End With
      Total_IVA = 0
      Total_IVA = CTNumero(TxtBaseImpoIce, 2)
      TxtMontoIce = 0
     'Calcula el Porcentaje de Ice
      CalIbMi = Total_IVA * CTNumero(DCPorcenIce, 2) / 100
      TxtMontoIce = CalIbMi
  End If
End Sub

Private Sub DCPorcenIva_GotFocus()
  MarcarTexto DCPorcenIva
End Sub

Private Sub DCPorcenIva_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCPorcenIva_LostFocus()
  If Not IsNumeric(DCPorcenIva) Then
     MsgBox "No ingrese caracteres. Vuelva a seleccionar.", vbInformation, "Aviso"
     DCPorcenIva = ""
    'Carga_PorcentajeIva (MBFechaRegis)
     DCPorcenIva.SetFocus
  Else
    'Busca y captura el codigo de Porcentaje IVA
     CodPorIva = "0"
     With AdoPorIva.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
         .Find ("Porc = " & CByte(DCPorcenIva) & " ")
          If Not .EOF Then CodPorIva = .Fields("Codigo")
      End If
     End With
     Total_IVA = 0
     Total_IVA = CTNumero(TxtBaseImpoGrav, 2)
    'Calcula el Porcentaje de Iva
     CalmIva = (Total_IVA * DCPorcenIva) / 100
     TxtMontoIva = CalmIva
  End If
End Sub

Private Sub DCProveedor_LostFocus()
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
                  & "FROM Trans_Importaciones " _
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

Private Sub DCRegimen_LostFocus()
  If Not IsNumeric(DCRegimen.Text) Then
     MsgBox "No ingrese caracteres alfanuméricos. Vuelva a seleccionar.", vbInformation, "Aviso"
     DCRegimen.Text = ""
     Carga_Regimen
     DCRegimen.SetFocus
     Carga_Regimen
  End If
End Sub

Private Sub DCSustento_LostFocus()
  If IsNumeric(DCSustento.Text) Then
     MsgBox "No ingrese números. Vuelva a seleccionar.", vbInformation, "Aviso"
     DCSustento.Text = ""
     Carga_CreditoTributario
     DCSustento.SetFocus
     Carga_TipoComprobante (SinEspaciosIzq(DCSustento))
  Else
     Carga_TipoComprobante (SinEspaciosIzq(DCSustento))
  End If
End Sub

Private Sub DCTipoComprobante_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCTipoComprobante_LostFocus()
  If IsNumeric(DCTipoComprobante.Text) Then
     MsgBox "No ingrese números. Vuelva a seleccionar.", vbInformation, "Aviso"
     DCTipoComprobante.Text = ""
     Carga_TipoComprobante (CargaTC)
     DCTipoComprobante.SetFocus
     Captura_TipoComprobante
  Else
     Captura_TipoComprobante
  End If
End Sub

Private Sub DGConceptoAir_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyDelete Then
     Titulo = "Aviso"
     Mensajes = "Desea Eliminar la Retención"
     If BoxMensaje = vbYes Then
        With AdoAsientoAir.Recordset
         If .RecordCount > 0 Then
             Codigo = .Fields("CodRet")
             No_Desde = .Fields("SecRetencion")
             Mifecha = .Fields("FechaEmiRet")
             Codigo1 = .Fields("AutRetencion")
             J = .Fields("A_No")
             sSQL = "DELETE * " _
                  & "FROM Asiento_Air " _
                  & "WHERE Item = '" & NumEmpresa & "' " _
                  & "AND CodigoU = '" & CodigoUsuario & "' " _
                  & "AND IdFiscalProv = '" & CodigoCliente & "' " _
                  & "AND T_No = " & Trans_No & " " _
                  & "AND Tipo_Trans = 'I' " _
                  & "AND A_No = " & J & " " _
                  & "AND CodRet = '" & Codigo & "' "
             ConectarAdoExecute sSQL
         End If
         AdoAsientoAir.Refresh
        End With
        Calculo_Sumatoria
     End If
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
  CentrarForm FImportaciones
  ConectarAdodc AdoSustento
  ConectarAdodc AdoTipoComprobante
  ConectarAdodc AdoImportacion
  ConectarAdodc AdoDistrito
  ConectarAdodc AdoRegimen
  ConectarAdodc AdoConceptoRet
  ConectarAdodc AdoRetFuente
  ConectarAdodc AdoPorIce
  ConectarAdodc AdoPorIva
  ConectarAdodc AdoClientes
  ConectarAdodc AdoAsientoAir
  ConectarAdodc AdoAsientoImportaciones
  ConectarAdodc AdoTransImportaciones
  ConectarAdodc AdoTransAir
  ConectarAdodc AdoAux
End Sub

Public Sub Carga_CreditoTributario()
  'Carga la Tabla de Catalogos Tributarios al DataCombo
   sSQL = "SELECT (Credito_Tributario & ' - ' & Descripcion) As Sustento, * " _
        & "FROM Tipo_Tributario " _
        & "WHERE Credito_Tributario <> '.' " _
        & "ORDER BY Credito_Tributario "
   SelectDBCombo DCSustento, AdoSustento, sSQL, "Sustento"
End Sub

Public Sub Carga_TipoImportacion()
  'Carga la Tabla de Importaciones al DataCombo
  sSQL = "SELECT * " _
       & "FROM Tabla_TipoImportacion " _
       & "WHERE Concepto <> '.' " _
       & "ORDER BY Concepto "
  SelectDBCombo DCImportacion, AdoImportacion, sSQL, "Concepto"
End Sub

Public Sub Carga_TipoComprobante(CargaTC As String)
  'Capturo el codigo del Tipo de Catalogo Tributario
  Cap = CargaTC
  'Busco el codigo en la tabla Tipo Comprobante///descripcion
   sSQL = "SELECT CTT.Identificacion,CTT.Tipo_Trans,TC.* " _
        & "FROM Tabla_Tributaria As CTT, Tipo_Comprobante As TC " _
        & "WHERE CTT.Identificacion = '" & CargaTC & "' " _
        & "AND CTT.Tipo_Comprobante_Codigo = TC.Tipo_Comprobante_Codigo " _
        & "AND CTT.Tipo_Trans = 'I' " _
        & "ORDER BY TC.Tipo_Comprobante_Codigo "
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

Public Sub Carga_PorcentajeIva(FechaLiquid As String)
Dim FechaCodAir As String
 'Carga la Tabla de Porcentaje Iva en el DataCombo
  FechaCodAir = BuscarFecha(FechaLiquid)
  sSQL = "SELECT * " _
       & "FROM Tabla_Por_ICE_IVA " _
       & "WHERE IVA <> " & Val(adFalse) & " " _
       & "AND Fecha_Inicio <= #" & FechaCodAir & "# " _
       & "AND Fecha_Final >= #" & FechaCodAir & "# " _
       & "ORDER BY Porc "
  SelectDBCombo DCPorcenIva, AdoPorIva, sSQL, "Porc"
End Sub

Public Sub Carga_PorcentajeIce()
  sSQL = "SELECT * " _
       & "FROM Tabla_Por_ICE_IVA " _
       & "WHERE ICE <> " & Val(adFalse) & " " _
       & "ORDER BY Porc"
  SelectDBCombo DCPorcenIce, AdoPorIce, sSQL, "Porc"
End Sub

Public Sub Carga_ConceptosRetencion()
  sSQL = "SELECT (Codigo & ' - ' & Concepto) As Detalle_Conceptos,* " _
       & "FROM Tipo_Concepto_Retencion " _
       & "WHERE Codigo <> '.' " _
       & "ORDER BY Codigo "
  SelectDBCombo DCConceptoRet, AdoConceptoRet, sSQL, "Detalle_Conceptos"
  With AdoConceptoRet.Recordset
   If .RecordCount > 0 Then Rf = 1 Else Rf = 0
  End With
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
          codTC = .Fields("Tipo_Comprobante_Codigo")
       Else
          MsgBox "Vuelva a seleccionar", vbInformation, "Aviso Importaciones"
          DCTipoComprobante.SetFocus
       End If
   End If
  End With
End Sub

Public Sub Leer_Clientes()
  'Carga en el Data Combo los Clientes con su RUC
  sSQL = "SELECT * " _
       & "FROM Clientes " _
       & "WHERE Codigo <>  '.' " _
       & "ORDER BY Cliente "
  SelectDBCombo DCProveedor, AdoClientes, sSQL, "Cliente"
End Sub

Public Sub Limpiar_Controles()
  ac = 0
  SSTImportaciones.Tab = 0
  DCSustento.Text = ""
  DCImportacion.Text = ""
  FechaValida MBFechaLiquida
  DCTipoComprobante.Text = ""
  DCDistrito.Text = ""
  TxtAño.Text = ""
  DCRegimen.Text = ""
  TxtCorrelativo.Text = ""
  TxtVerificador.Text = ""
  DCProveedor.Text = ""
  LblTD.Caption = ""
  TxtValorCIF.Text = ""
  TxtBaseImpo.Text = ""
  TxtBaseImpoGrav.Text = ""
  DCPorcenIva.Text = ""
  TxtMontoIva.Text = ""
  TxtBaseImpoIce.Text = ""
  DCPorcenIce.Text = ""
  TxtMontoIce.Text = ""
  DCRetFuente.Text = ""
  TxtNumUnoComRet.Text = ""
  TxtNumDosComRet.Text = ""
  TxtNumTresComRet.Text = ""
  TxtNumUnoAutComRet.Text = ""
  TxtSumatoria.Text = ""
  DCConceptoRet.Text = ""
  TxtBimpConA.Text = ""
  TxtPorRetConA.Text = ""
  TxtValConA.Text = ""
  TxtTotalReten.Text = ""
  LblNumIdent.Caption = ""
  TxtNumeroC.Text = ""
  TxtTotalReten.Text = ""
  CTP.AddItem "CE"
  CTP.AddItem "CI"
  CTP.AddItem "CD"
  CTP.Text = "CE"
    
  'Limpia la grilla
  'Borra Asiento Air
  sSQL = "DELETE * " _
       & "FROM Asiento_Air " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND Tipo_Trans = 'I' "
  ConectarAdoExecute sSQL
End Sub

Public Sub Calculo_Sumatoria()
Dim SumaReten As Currency
  SumaReten = 0
  With AdoAsientoAir.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          SumaReten = SumaReten + .Fields("ValRet")
         .MoveNext
       Loop
   End If
  End With
  TxtTotalReten = Format(SumaReten, "#,##0.00")
End Sub

Private Sub MBFechaLiquida_GotFocus()
  MarcarTexto MBFechaLiquida
End Sub

Private Sub MBFechaLiquida_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaLiquida_LostFocus()
  FechaValida MBFechaLiquida
  Carga_PorcentajeIva (MBFechaLiquida)
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
     Anio = Year(MBFechaLiquida)
     If TxtAño <> Anio Then
        MsgBox "Año incorrecto. Vuelva a ingresar.", vbInformation, "Aviso"
        TxtAño.Text = ""
        TxtAño.SetFocus
     End If
  End If
End Sub

Private Sub TxtBaseImpo_GotFocus()
  MarcarTexto TxtBaseImpo
End Sub

Private Sub TxtBaseImpo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtBaseImpo_LostFocus()
  TextoValido TxtBaseImpo, True, , 0
End Sub

Private Sub TxtBaseImpoGrav_GotFocus()
  MarcarTexto TxtBaseImpoGrav
End Sub

Private Sub TxtBaseImpoGrav_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtBaseImpoGrav_LostFocus()
  TextoValido TxtBaseImpoGrav, True, , 0
End Sub

Private Sub TxtBaseImpoIce_GotFocus()
  MarcarTexto TxtBaseImpoIce
End Sub

Private Sub TxtBaseImpoIce_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtBaseImpoIce_LostFocus()
  TextoValido TxtBaseImpoIce, True, , 0
End Sub

Private Sub TxtBimpConA_GotFocus()
  MarcarTexto TxtBimpConA
End Sub

Private Sub TxtBimpConA_LostFocus()
  RatonNormal
  TextoValido TxtBimpConA, True, , 0
  TextoValido TxtSumatoria, True, , 0
  'Valida que la base imponible no sea mayor que la BIG y la BIcero
  If CTNumero(TxtBimpConA, 2) > CTNumero(TxtSumatoria, 2) Then
     MsgBox "La Base Imponible debe ser menor o igual a la " & vbCrLf _
     & "Base Imponible Gravada + la Base Imponible 0%", vbInformation, "Aviso"
     TxtBimpConA.Text = 0
     TxtBimpConA.SetFocus
  Else
     If OP Then
        TxtValConA = (CTNumero(TxtBimpConA, 2) * CTNumero(TxtPorRetConA, 2)) / 100
        Insertar_AsientoAir
     Else
        TxtPorRetConA.SetFocus
     End If
  End If
End Sub

Sub Insertar_AsientoAir()
  'Selecciona el numero mayor para continuar la secuencia en el
  'campo T_No y A_No
  Ln_No = Maximo_De("Asiento_Air", "A_No")
  If CTNumero(TxtBimpConA, 2) > 0 Then
     RatonReloj
     Espizq = SinEspaciosIzq(DCConceptoRet)
     Espder = Trim$(Mid$(DCConceptoRet, Len(Espizq) + 3, Len(DCConceptoRet)))
     SetAdoAddNew "Asiento_Air"
     SetAdoFields "CodRet", Espizq
     SetAdoFields "Detalle", Espder
     SetAdoFields "BaseImp", Val(TxtBimpConA)
     SetAdoFields "Porcentaje", Val(TxtPorRetConA) / 100
     SetAdoFields "ValRet", TxtValConA
     SetAdoFields "EstabRetencion", TxtNumUnoComRet
     SetAdoFields "PtoEmiRetencion", TxtNumDosComRet
     SetAdoFields "SecRetencion", TxtNumTresComRet
     SetAdoFields "AutRetencion", TxtNumUnoAutComRet
     SetAdoFields "FechaEmiRet", MBFechaLiquida
     SetAdoFields "EstabFactura", "001"
     SetAdoFields "PuntoEmiFactura", "001"
     SetAdoFields "Factura_No", TxtCorrelativo
     SetAdoFields "Cta_Retencion", SinEspaciosIzq(DCRetFuente)
     SetAdoFields "IdProv", CodigoCliente
     SetAdoFields "A_No", Ln_No
     SetAdoFields "T_No", Trans_No
     SetAdoFields "Tipo_Trans", "I"
     SetAdoUpdate
     Ln_No = Ln_No + 1
       
     'Despliega los datos en el DataGrid
     sSQL = "SELECT * " _
        & "FROM Asiento_Air " _
        & "WHERE CodRet <> '.' " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "AND CodigoU = '" & CodigoUsuario & "' " _
        & "AND T_No = " & Trans_No & " " _
        & "AND Tipo_Trans = 'I' " _
        & "ORDER BY CodRet "
     SelectDataGrid DGConceptoAir, AdoAsientoAir, sSQL
            
     'Realiza la Sumatoria de las Retenciones
     ac = ac + TxtValConA
     TxtTotalReten = ac
  End If
  RatonNormal
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

Private Sub TxtMontoIva_GotFocus()
  MarcarTexto TxtMontoIva
End Sub

Private Sub TxtMontoIva_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtMontoIva_LostFocus()
  TextoValido TxtMontoIva, True, , 0
End Sub
Private Sub TxtNumDosComRet_GotFocus()
  MarcarTexto TxtNumDosComRet
End Sub

Private Sub TxtNumDosComRet_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumDosComRet_LostFocus()
  TextoValido TxtNumDosComRet, True, , 0
  If Val(TxtNumDosComRet) <= 0 Then TxtNumDosComRet = "001"
  TxtNumDosComRet = Format(Val(TxtNumDosComRet), "000")
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

Private Sub TxtNumTresComRet_GotFocus()
  MarcarTexto TxtNumTresComRet
End Sub

Private Sub TxtNumTresComRet_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumTresComRet_LostFocus()
  TextoValido TxtNumTresComRet, True, , 0
  If Val(TxtNumTresComRet) <= 0 Then TxtNumTresComRet = "000001"
  TxtNumTresComRet = Format(Val(CCur(TxtNumTresComRet)), "0000000")
  If DCImportacion = "Bienes" Then
    'Calcula la sumatoria de Monto Iva Bienes, Monto Iva Servicios y Base Imponible
     TxtSumatoria = Val(CCur(TxtBaseImpo)) + Val(CCur(TxtBaseImpoGrav))
  Else
     TxtSumatoria = 0
  End If
End Sub

Private Sub TxtNumUnoAutComRet_GotFocus()
  MarcarTexto TxtNumUnoAutComRet
End Sub

Private Sub TxtNumUnoAutComRet_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumUnoComRet_GotFocus()
  MarcarTexto TxtNumUnoComRet
End Sub

Private Sub TxtNumUnoComRet_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumUnoComRet_LostFocus()
  TextoValido TxtNumUnoComRet, True, , 0
  If Val(TxtNumUnoComRet) <= 0 Then TxtNumUnoComRet = "001"
  TxtNumUnoComRet = Format(Val(TxtNumUnoComRet), "000")
End Sub

Private Sub TxtPorRetConA_GotFocus()
  MarcarTexto TxtPorRetConA
End Sub

Private Sub TxtPorRetConA_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtPorRetConA_LostFocus()
  If OP = False Then
     TxtValConA = (CTNumero(TxtBimpConA, 2) * CTNumero(TxtPorRetConA, 2)) / 100
     Insertar_AsientoAir
  End If
End Sub

Private Sub TxtValorCIF_GotFocus()
  MarcarTexto TxtValorCIF
End Sub

Private Sub TxtValorCIF_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtValorCIF_LostFocus()
  TextoValido TxtValorCIF, True, , 0
End Sub

Private Sub TxtVerificador_GotFocus()
  MarcarTexto TxtVerificador
End Sub

Private Sub TxtVerificador_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Public Sub Carga_RetencionFuente()
  'Carga los Conceptos de retención en la Fuente al DataCombo
  sSQL = "SELECT (Codigo & ' - ' & Cuenta) As Cuentas  " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'RF' " _
       & "AND DG = 'D' " _
       & "ORDER BY Codigo "
  SelectDBCombo DCRetFuente, AdoRetFuente, sSQL, "Cuentas"
End Sub

Public Sub Captura_TipoImportacion()
  'Busca y captura el codigo de Tipo de Importación
  With AdoImportacion.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Concepto = '" & DCImportacion & "' ")
       If Not .EOF Then
          CodImp = .Fields("Codigo")
       Else
          MsgBox "Vuelva a seleccionar", vbInformation, "Aviso"
          DCImportacion.SetFocus
       End If
    End If
 End With
End Sub

Public Sub Carga_Datos_Iniciales(MBFecha As MaskEdBox, EsNuevo As Boolean)
  Encerar_Var
  SSTImportaciones.Tab = 0
  TxtAño.Text = Year(date)
        
  'Enceramos el espacio de cada usuario para emprezar con una nueva retencion
  sSQL = "DELETE * " _
       & "FROM Asiento_Importaciones " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  ConectarAdoExecute sSQL

  'Borra Asiento Air
  sSQL = "DELETE * " _
       & "FROM Asiento_Air  " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
   ConectarAdoExecute sSQL

   'Carga la Tabla de Catalogos Tributarios al DataCombo
   Carga_CreditoTributario
   
   'Carga la Tabla de Tipo Importacion al DataCombo
   Carga_TipoImportacion
   
   'Carga la Tabla de Distrito al DataCombo
   Carga_Distrito
   
   'Carga la Tabla de Regimen al DataCombo
   Carga_Regimen
   
   'Carga la Tabla de Clientes al DataCombo
   Leer_Clientes
   
   'Carga la Tabla de Porcentaje Ice en el DataCombo
   Carga_PorcentajeIce
   
   'Carga la Tabla de Conceptos Retencion al DataCombo
   Carga_ConceptosRetencion
    
   'Carga la Retencion en la Fuente
   Carga_RetencionFuente
    
   CTP.Clear
   CTP.AddItem "CE"
   CTP.AddItem "CI"
   CTP.AddItem "CD"
   CTP.Text = "CE"
   
   'Si es Nuevo ingresa por aqui
   If EsNuevo Then
      DCProveedor.SetFocus
   Else
      'Si es Modificación viene por aca
       CMes.Visible = True
       CAño.Visible = True
       Modificacion_AT "I", CModificacion, CMes, CAño
       CMes.Text = MesesLetras(Month(FechaSistema))
       CAño.SetFocus
       CAño.Text = Year(FechaSistema)
       CModificacion.Visible = True
   End If
End Sub

Public Sub Grabacion()
   'Capturo el Tipo de Importacion
    Captura_TipoImportacion

   'Selecciona el numero mayor para continuar la secuencia en el
   'campo T_No y A_No
   'Grabo en el Asiento_Importacioness e implicito Asiento_Air
    SetAdoAddNew "Asiento_Importaciones"
    SetAdoFields "CodSustento", Cap
    SetAdoFields "ImportacionDe", CodImp
    SetAdoFields "FechaLiquidacion", MBFechaLiquida
    SetAdoFields "TipoComprobante", codTC
    SetAdoFields "IdFiscalProv", CodigoCliente
    SetAdoFields "ValorCIF", CTNumero(TxtValorCIF, 2)
    If CodImp = 1 Then
       SetAdoFields "DistAduanero", DCDistrito
       SetAdoFields "Anio", TxtAño
       SetAdoFields "Regimen", DCRegimen
       SetAdoFields "Correlativo", TxtCorrelativo
       SetAdoFields "Verificador", TxtVerificador
       SetAdoFields "BaseImponible", CTNumero(TxtBaseImpo, 2)
       SetAdoFields "BaseImpGrav", CTNumero(TxtBaseImpoGrav, 2)
       SetAdoFields "PorcentajeIva", CodPorIva
       SetAdoFields "MontoIva", CTNumero(TxtMontoIva, 2)
       SetAdoFields "BaseImpIce", CTNumero(TxtBaseImpoIce, 2)
       SetAdoFields "PorcentajeIce", CodPorIce
       SetAdoFields "MontoIce", CTNumero(TxtMontoIce, 2)
    Else
       SetAdoFields "DistAduanero", 0
       SetAdoFields "Anio", 0
       SetAdoFields "Regimen", 0
       SetAdoFields "Correlativo", 0
       SetAdoFields "Verificador", 0
       SetAdoFields "BaseImponible", 0
       SetAdoFields "BaseImpGrav", 0
       SetAdoFields "PorcentajeIva", 0
       SetAdoFields "MontoIva", 0
       SetAdoFields "BaseImpIce", 0
       SetAdoFields "PorcentajeIce", 0
       SetAdoFields "MontoIce", 0
    End If
    SetAdoFields "A_No", 1
    SetAdoFields "T_No", Trans_No
    SetAdoUpdate
      
    'Grabamos los datos de la transaccion en la tabla definitiva de almacenamiento
    ID_Trans = Maximo_De("Trans_Importaciones", "ID")  'va a tener el indice de transaccion unico para que no exista duplicados en a base
    sSQL = "SELECT * " _
         & "FROM Asiento_Importaciones " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND T_No = " & Trans_No & " " _
         & "ORDER BY T_No DESC "
    SelectAdodc AdoAsientoImportaciones, sSQL
    With AdoAsientoImportaciones.Recordset
     If .RecordCount > 0 Then
         FechaTexto = .Fields("FechaLiquidacion")
         SetAdoAddNew "Trans_Importaciones"
         SetAdoFields "T", Normal
         SetAdoFields "CodSustento", .Fields("CodSustento")
         SetAdoFields "ImportacionDe", .Fields("ImportacionDe")
         SetAdoFields "FechaLiquidacion", .Fields("FechaLiquidacion")
         SetAdoFields "TipoComprobante", .Fields("TipoComprobante")
         SetAdoFields "DistAduanero", .Fields("DistAduanero")
         SetAdoFields "Anio", .Fields("Anio")
         SetAdoFields "Regimen", .Fields("Regimen")
         SetAdoFields "Correlativo", .Fields("Correlativo")
         SetAdoFields "Verificador", .Fields("Verificador")
         SetAdoFields "IdFiscalProv", .Fields("IdFiscalProv")
         SetAdoFields "ValorCIF", .Fields("ValorCIF")
         SetAdoFields "BaseImponible", .Fields("BaseImponible")
         SetAdoFields "BaseImpGrav", .Fields("BaseImpGrav")
         SetAdoFields "PorcentajeIva", .Fields("PorcentajeIva")
         SetAdoFields "MontoIva", .Fields("MontoIva")
         SetAdoFields "BaseImpIce", .Fields("BaseImpIce")
         SetAdoFields "PorcentajeIce", .Fields("PorcentajeIce")
         SetAdoFields "MontoIce", .Fields("MontoIce")
         SetAdoFields "A_No", 1
         SetAdoFields "T_No", Trans_No
         SetAdoFields "TP", CTP
         SetAdoFields "Numero", TxtNumeroC
         SetAdoFields "Fecha", MBFechaLiquida
         SetAdoFields "ID", ID_Trans
         SetAdoFields "Linea_SRI", 0
         SetAdoUpdate
     End If
    End With
   'Selecciona el numero mayor para continuar la secuencia en el
    ID_Trans = Maximo_De("Trans_Air", "ID")
    sSQL = "SELECT * " _
         & "FROM Asiento_Air " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND T_No = " & Trans_No & " " _
         & "ORDER BY A_No "
    SelectAdodc AdoTransAir, sSQL
    
    'Verifico si el codigo de importación es Bienes o Servicios
    'Para mandar a grabar en el Asiento Air
    If CodImp = 2 Then
        With AdoAsientoAir.Recordset
         If .RecordCount > 0 Then
             Do While Not .EOF
                SetAdoAddNew "Trans_Air"
                SetAdoFields "T", Normal
                SetAdoFields "CodRet", .Fields("CodRet")
                SetAdoFields "BaseImp", .Fields("BaseImp")
                SetAdoFields "Porcentaje", .Fields("Porcentaje")
                SetAdoFields "ValRet", .Fields("ValRet")
                SetAdoFields "EstabRetencion", .Fields("EstabRetencion")
                SetAdoFields "PtoEmiRetencion", .Fields("PtoEmiRetencion")
                SetAdoFields "SecRetencion", .Fields("SecRetencion")
                SetAdoFields "AutRetencion", .Fields("AutRetencion")
                SetAdoFields "Tipo_Trans", .Fields("Tipo_Trans")
                SetAdoFields "Fecha", FechaTexto
                SetAdoFields "EstabFactura", "001"
                SetAdoFields "PuntoEmiFactura", "001"
                SetAdoFields "Factura_No", .Fields("Factura_No")
                SetAdoFields "IdProv", .Fields("IdProv")
                SetAdoFields "TP", CTP
                SetAdoFields "Numero", TxtNumeroC
                SetAdoFields "Cta_Retencion", .Fields("Cta_Retencion")
                SetAdoFields "ID", ID_Trans
                SetAdoFields "Linea_SRI", 0
                SetAdoUpdate
                'ID_Trans = ID_Trans + 1
               .MoveNext
             Loop
         End If
        End With
    End If
End Sub

Public Sub Habilita_Controles()
  'Habilito los controles para la modificacion
  CModificacion.Enabled = True
  SSTImportaciones.Enabled = True
  DCProveedor.Enabled = True
  CmdGrabar.Enabled = True
  CTP.Enabled = True
  TxtNumeroC.Enabled = True
  'Label23.Visible = True
  CMes.Visible = True
  CAño.Visible = True
  'Label29.Visible = True
End Sub

Public Sub Deshabilita_Controles()
  'Deshabilito los controles para la modificacion
  CModificacion.Enabled = False
  SSTImportaciones.Enabled = False
  DCProveedor.Enabled = False
  CmdGrabar.Enabled = False
  CTP.Enabled = False
  TxtNumeroC.Enabled = False
  'Label23.Visible = False
  CMes.Visible = False
  CAño.Visible = False
  'Label29.Visible = False
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

Public Sub Encerar_Var()
  ac = 0
  Ln_No = 0
  CodImp = 0
  DCPorcenIce = 0
  DCPorcenIva = 0
  CodPorIce = "0"
  CodPorIva = "0"
  CodRetBien = "0"
  CodRetServ = "0"
End Sub

