VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form FImportacionesAT 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IMPORTACIONES"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10635
   Icon            =   "FImportacionesAT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   10635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdCerrar 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Cancelar"
      Height          =   645
      Left            =   9555
      Picture         =   "FImportacionesAT.frx":0696
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Salir"
      Top             =   105
      Width           =   960
   End
   Begin VB.CommandButton CmdGrabar 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Aceptar"
      Height          =   645
      Left            =   8505
      Picture         =   "FImportacionesAT.frx":0AD8
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Grabar"
      Top             =   105
      Width           =   960
   End
   Begin TabDlg.SSTab SSTImportaciones 
      Height          =   4005
      Left            =   105
      TabIndex        =   0
      Top             =   840
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   7064
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   12648384
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
      TabPicture(0)   =   "FImportacionesAT.frx":0DE2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DCTipoComprobante"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "MBFechaLiquida"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "DCImportacion"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "DCSustento"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "FraBases"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "FraRefrendo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "TxtValorCIF"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "CmdAir"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Conceptos AIR"
      TabPicture(1)   =   "FImportacionesAT.frx":0DFE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FraRetencion"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.CommandButton CmdAir 
         Caption         =   "AIR"
         Height          =   765
         Left            =   9660
         Picture         =   "FImportacionesAT.frx":0E1A
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Se ubica en la pestaña de Retenciones"
         Top             =   2940
         Width           =   660
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
         Height          =   3615
         Left            =   -74895
         TabIndex        =   23
         Top             =   312
         Width           =   10155
         Begin VB.TextBox TxtNumUnoComRet 
            Enabled         =   0   'False
            Height          =   336
            Left            =   105
            MaxLength       =   3
            TabIndex        =   26
            Text            =   "001"
            Top             =   864
            Width           =   552
         End
         Begin VB.TextBox TxtNumDosComRet 
            Enabled         =   0   'False
            Height          =   336
            Left            =   756
            MaxLength       =   3
            TabIndex        =   27
            Text            =   "001"
            Top             =   864
            Width           =   552
         End
         Begin VB.TextBox TxtNumTresComRet 
            Enabled         =   0   'False
            Height          =   336
            Left            =   1404
            MaxLength       =   7
            TabIndex        =   28
            Text            =   "0000001"
            Top             =   864
            Width           =   876
         End
         Begin VB.TextBox TxtNumUnoAutComRet 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   330
            Left            =   2268
            MaxLength       =   10
            MultiLine       =   -1  'True
            TabIndex        =   29
            Top             =   864
            Width           =   1632
         End
         Begin VB.TextBox TxtBimpConA 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   336
            Left            =   6510
            TabIndex        =   32
            Top             =   1512
            Width           =   1416
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
            TabIndex        =   30
            Top             =   756
            Width           =   1956
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
            TabIndex        =   35
            Top             =   3132
            Width           =   1308
         End
         Begin VB.TextBox TxtPorRetConA 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   336
            Left            =   7992
            TabIndex        =   38
            Top             =   1512
            Width           =   660
         End
         Begin VB.TextBox TxtValConA 
            Enabled         =   0   'False
            Height          =   336
            Left            =   8748
            TabIndex        =   33
            Top             =   1512
            Width           =   1308
         End
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
            TabIndex        =   24
            Top             =   315
            Width           =   2328
         End
         Begin MSDataListLib.DataCombo DCConceptoRet 
            Bindings        =   "FImportacionesAT.frx":1340
            DataSource      =   "AdoConceptoRet"
            Height          =   405
            Left            =   105
            TabIndex        =   31
            Top             =   1515
            Width           =   6390
            _ExtentX        =   11271
            _ExtentY        =   714
            _Version        =   393216
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataGridLib.DataGrid DGConceptoAir 
            Bindings        =   "FImportacionesAT.frx":135D
            Height          =   1176
            Left            =   108
            TabIndex        =   34
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
            Bindings        =   "FImportacionesAT.frx":1379
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
      Begin VB.TextBox TxtValorCIF 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   9135
         TabIndex        =   6
         Text            =   "0.00"
         ToolTipText     =   "En este campo obligatorio se debe ingresar el valor CIF de los bienes importados o el valor del pago efectuado al exterior"
         Top             =   1260
         Width           =   1170
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
         TabIndex        =   8
         Top             =   1785
         Width           =   5160
         Begin VB.TextBox TxtVerificador 
            Height          =   330
            Left            =   3990
            MaxLength       =   1
            TabIndex        =   13
            ToolTipText     =   "Verificador (1 caracter)"
            Top             =   420
            Width           =   750
         End
         Begin VB.TextBox TxtCorrelativo 
            Height          =   330
            Left            =   2835
            MaxLength       =   6
            TabIndex        =   12
            Text            =   "000001"
            ToolTipText     =   "Correlativo (6 caracteres)"
            Top             =   420
            Width           =   750
         End
         Begin VB.TextBox TxtAño 
            Height          =   330
            Left            =   1050
            MaxLength       =   4
            TabIndex        =   10
            ToolTipText     =   "Años (4 caracteres)"
            Top             =   420
            Width           =   645
         End
         Begin MSDataListLib.DataCombo DCRegimen 
            Bindings        =   "FImportacionesAT.frx":1394
            DataSource      =   "AdoRegimen"
            Height          =   315
            Left            =   1890
            TabIndex        =   11
            ToolTipText     =   "Regimen (2 caracteres)"
            Top             =   420
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DCDistrito 
            Bindings        =   "FImportacionesAT.frx":13AD
            DataSource      =   "AdoDistrito"
            Height          =   315
            Left            =   105
            TabIndex        =   9
            ToolTipText     =   "Distrito aduanero (3 caracteres)"
            Top             =   420
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
      End
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
         Height          =   855
         Left            =   105
         TabIndex        =   14
         Top             =   2835
         Width           =   9465
         Begin VB.TextBox TxtBaseImpoIce 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   5460
            MultiLine       =   -1  'True
            TabIndex        =   19
            Text            =   "FImportacionesAT.frx":13C7
            ToolTipText     =   $"FImportacionesAT.frx":13CC
            Top             =   420
            Width           =   1485
         End
         Begin VB.TextBox TxtBaseImpoGrav 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   1575
            MultiLine       =   -1  'True
            TabIndex        =   16
            Text            =   "FImportacionesAT.frx":146C
            ToolTipText     =   "Corresponde al valor de la importación gravada con tarifa 12% en el momento de la desaduanización"
            Top             =   420
            Width           =   1485
         End
         Begin VB.TextBox TxtBaseImpo 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   105
            MaxLength       =   14
            MultiLine       =   -1  'True
            TabIndex        =   15
            Text            =   "FImportacionesAT.frx":1473
            ToolTipText     =   "Corresponde al valor de la importación gravada con tarifa 0% o exento"
            Top             =   420
            Width           =   1485
         End
         Begin VB.TextBox TxtMontoIce 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   336
            Left            =   7980
            TabIndex        =   21
            Top             =   420
            Width           =   1380
         End
         Begin VB.TextBox TxtMontoIva 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   4095
            TabIndex        =   18
            Text            =   "0.00"
            ToolTipText     =   $"FImportacionesAT.frx":147A
            Top             =   420
            Width           =   1380
         End
         Begin MSDataListLib.DataCombo DCPorcenIva 
            Bindings        =   "FImportacionesAT.frx":1564
            DataSource      =   "AdoPorIva"
            Height          =   315
            Left            =   3045
            TabIndex        =   17
            ToolTipText     =   "Este campo corresponde al porentaje de IVA vigente a la fecha de desaduanización"
            Top             =   420
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DCPorcenIce 
            Bindings        =   "FImportacionesAT.frx":157C
            DataSource      =   "AdoPorIce"
            Height          =   315
            Left            =   6930
            TabIndex        =   20
            ToolTipText     =   "Este campo corresponde al porcentaje de ICE vigente a lafecha de emisión de la importación"
            Top             =   420
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
      End
      Begin MSDataListLib.DataCombo DCSustento 
         Bindings        =   "FImportacionesAT.frx":1594
         DataSource      =   "AdoSustento"
         Height          =   405
         Left            =   105
         TabIndex        =   3
         ToolTipText     =   "En este combo de selección se despliega una lista de tipos de sustentos tributarios."
         Top             =   645
         Width           =   10170
         _ExtentX        =   17939
         _ExtentY        =   714
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DCImportacion 
         Bindings        =   "FImportacionesAT.frx":15AE
         DataSource      =   "AdoImportacion"
         Height          =   420
         Left            =   105
         TabIndex        =   4
         ToolTipText     =   $"FImportacionesAT.frx":15CB
         Top             =   1290
         Width           =   6450
         _ExtentX        =   11377
         _ExtentY        =   741
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
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
         TabIndex        =   5
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
         Bindings        =   "FImportacionesAT.frx":16A7
         DataSource      =   "AdoTipoComp"
         Height          =   405
         Left            =   105
         TabIndex        =   7
         ToolTipText     =   "Corresponde al tipo de comprobante utilizado en la transacción"
         Top             =   2310
         Width           =   4980
         _ExtentX        =   8784
         _ExtentY        =   714
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc AdoTipoComprobante 
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
      Top             =   1050
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
      Top             =   1050
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
   Begin MSAdodcLib.Adodc AdoAsientos 
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
      Left            =   6510
      TabIndex        =   2
      Top             =   315
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
      Left            =   6195
      TabIndex        =   1
      Top             =   315
      Width           =   330
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
      TabIndex        =   39
      Top             =   315
      Width           =   6105
   End
End
Attribute VB_Name = "FImportacionesAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MBFecha As MaskEdBox
Dim FechaLiquid As Date
Dim OP As Boolean
Dim cod, X, CodImp, CodTs, Longitud, codTC, Rf As Byte
Dim SumAnio, Aniocad, AniocadAux, CodPorIva, CodRetBien, CodRetServ, CodReg As Integer
Dim CalmIva, CalmIce, CalIbMi, CalIsMi, ac, SUM, cal As Double
Dim Cap, Cap1, ch, CodDis, CodProv, CargaTC, Opc, CodPorIce As String
Dim Espizq, Espder, Captc, PorIva, PorIce, TipoImp, CodSus As String

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
   Ejecutar_SQL_SP sSQL
  'Borra Asiento Air
   sSQL = "DELETE * " _
        & "FROM Asiento_Air " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND CodigoU = '" & CodigoUsuario & "' " _
        & "AND Tipo_Trans = 'I' " _
        & "AND T_No = " & Trans_No & " "
   Ejecutar_SQL_SP sSQL
   Unload FImportacionesAT
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
   'Grabacion de los Datos
    Grabacion
    sSQL = "SELECT * " _
         & "FROM Asiento " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND T_No = " & Trans_No & " "
    Select_Adodc AdoAsientos, sSQL
    OpcTM = 1
    OpcDH = 2
    NoCheque = Ninguno
   'Grabamos el Asiento de las Retenciones
    sSQL = "SELECT * " _
         & "FROM Asiento_Air " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND T_No = " & Trans_No & " " _
         & "AND Tipo_Trans = 'I' " _
         & "ORDER BY Cta_Retencion,A_No,ValRet "
    Select_Adodc AdoAux, sSQL
    With AdoAux.Recordset
     If .RecordCount > 0 Then
         Do While Not .EOF
            Cta = .Fields("Cta_Retencion")
            DetalleComp = "Retencion No. " & .Fields("SecRetencion") & " del " & (.Fields("Porcentaje") * 100) & "%, de " & NombreCliente
            Codigo = Leer_Cta_Catalogo(Cta)
            ValorDH = .Fields("ValRet")
            If ValorDH > 0 Then InsertarAsiento AdoAsientos
           .MoveNext
         Loop
     End If
    End With
    DetalleComp = Ninguno
    Unload FImportacionesAT
End Sub

Private Sub DCConceptoRet_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCConceptoRet_LostFocus()
    OP = False
    If IsNumeric(DCConceptoRet.Text) Then
       MsgBox "No ingrese números. Vuelva a seleccionar.", vbInformation, "Aviso"
       DCConceptoRet.SetFocus
    Else
       With AdoConceptoRet.Recordset
        If .RecordCount > 0 Then
           .MoveFirst
           .Find ("Codigo = '" & SinEspaciosIzq(DCConceptoRet) & "' ")
            If Not .EOF Then
               TxtPorRetConA = .Fields("Porcentaje")
               If .Fields("Ingresar_Porcentaje") = "S" Then OP = True
            Else
               MsgBox "No encontro este código vuelva a buscar"
            End If
        End If
       End With
       TxtBimpConA = TxtSumatoria
    End If
    If OP Then
       TxtPorRetConA.Enabled = True
       TxtPorRetConA.SetFocus
    Else
       TxtPorRetConA.Enabled = False
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
   'Desactivo dependiendo del Tipo de Importación
   If (DCImportacion = "Bienes") Then
      FraRefrendo.Enabled = True
      FraBases.Enabled = True
      FraRetencion.Enabled = False
   Else
      FraRefrendo.Enabled = False
      FraBases.Enabled = False
      FraRetencion.Enabled = True
   End If
End Sub

Private Sub DCPorcenIce_LostFocus()
    If Not IsNumeric(DCPorcenIce) Then
        MsgBox "No ingrese caracteres. Vuelva a seleccionar.", vbInformation, "Aviso"
        DCPorcenIce = ""
        DCPorcenIce.SetFocus
    Else
        'Busca y captura el codigo de Porcentaje IVA
        PorIce = SinEspaciosDer(DCPorcenIce.Text)
        With AdoPorIce.Recordset
            If .RecordCount > 0 Then
               .MoveFirst
               .Find ("Porc = '" & PorIce & "' ")
                CodPorIce = .Fields("Codigo")
            End If
        End With
        
        Total_IVA = 0
        Total_IVA = CTNumero(TxtBaseImpoIce, 2)
       'Calcula el Porcentaje de Ice
        CalIbMi = (Total_IVA * DCPorcenIce) / 100
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
        DCPorcenIva.SetFocus
    Else
        'Busca y captura el codigo de Porcentaje IVA
        PorIva = SinEspaciosDer(DCPorcenIva.Text)
        With AdoPorIva.Recordset
         If .RecordCount > 0 Then
            .MoveFirst
            .Find ("Porc = '" & PorIva & "' ")
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
           
           sSQL = "DELETE * " _
                & "FROM Asiento_Air " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND CodigoU = '" & CodigoUsuario & "' " _
                & "AND T_No = " & Trans_No & " " _
                & "AND FechaEmiRet = #" & BuscarFecha(Mifecha) & "# " _
                & "AND IdProv = '" & CodigoCliente & "' " _
                & "AND CodRet = '" & Codigo & "' " _
                & "AND Tipo_Trans = 'I' " _
                & "AND SecRetencion = " & No_Desde & " " _
                & "AND AutRetencion = '" & Codigo1 & "' "
           Ejecutar_SQL_SP sSQL
        End If
       Calculo_Sumatoria
      End With
    End If
 End If
End Sub

Private Sub Form_Activate()
   MBFechaLiquida = FechaComp
   LblTD.Caption = TipoBenef                  ' Tipo de Cliente: C,R,P,O
   LblNumIdent = CICliente                    ' CI o RUC del Cliente
   LblProveedor.Caption = " " & NombreCliente ' Nombre del Cliente
   MBFechaLiquida = FechaComp
  'CodigoCliente
   Carga_Datos_Iniciales MBFechaLiquida, Nuevo
   TxtCorrelativo = "000001"
  'Aqui despliego el ultimo numero de la Transaccion
   sSQL = "SELECT TOP 1 * " _
        & "FROM Trans_Importaciones " _
        & "WHERE IdFiscalProv = '" & CodigoCliente & "' " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "ORDER BY Correlativo DESC "
   Select_Adodc AdoAux, sSQL
   With AdoAux.Recordset
    If .RecordCount > 0 Then TxtCorrelativo = .Fields("Correlativo")
   End With
End Sub

Private Sub Form_Load()
    CentrarForm FImportacionesAT
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
    ConectarAdodc AdoAsientos
End Sub

Public Sub Carga_CreditoTributario()
' Carga la Tabla de Catalogos Tributarios al DataCombo
    sSQL = "SELECT (Credito_Tributario & ' - ' & Descripcion) As Sustento, * " _
         & "FROM Tipo_Tributario " _
         & "WHERE Credito_Tributario <> '.' " _
         & "ORDER BY Credito_Tributario "
    SelectDB_Combo DCSustento, AdoSustento, sSQL, "Sustento"
End Sub

Public Sub Carga_TipoImportacion()
' Carga la Tabla de Importaciones al DataCombo
    sSQL = "SELECT * " _
         & "FROM Tabla_TipoImportacion " _
         & "WHERE Concepto <> '.' " _
         & "ORDER BY Concepto "
    SelectDB_Combo DCImportacion, AdoImportacion, sSQL, "Concepto"
End Sub

Public Sub Carga_TipoComprobante(CargaTC As String)
    'Capturo el codigo del Tipo de Catalogo Tributario
         Cap = CargaTC
            
    'Busco el codigo en la tabla Tipo Comprobante///descripcion
     sSQL = "SELECT CTT.Identificacion,CTT.Tipo_Trans,TC.* " _
          & "FROM Tabla_Tributaria As CTT, Tipo_Comprobante As TC " _
          & "WHERE CTT.Identificacion = '" & CargaTC & "' " _
          & "AND TC.TC = 'TDC' " _
          & "AND CTT.Tipo_Trans = 'I' " _
          & "AND CTT.Tipo_Comprobante_Codigo = TC.Tipo_Comprobante_Codigo " _
          & "ORDER BY TC.Tipo_Comprobante_Codigo "
     SelectDB_Combo DCTipoComprobante, AdoTipoComprobante, sSQL, "Descripcion"
End Sub

Public Sub Carga_Distrito()
' Carga la Tabla de Distrito al DataCombo
    sSQL = "SELECT * " _
         & "FROM Tabla_Distrito " _
         & "WHERE Codigo <> '.' " _
         & "ORDER BY Codigo "
    SelectDB_Combo DCDistrito, AdoDistrito, sSQL, "Codigo"
End Sub

Public Sub Carga_Regimen()
' Carga la Tabla de Regimen al DataCombo
    sSQL = "SELECT * " _
         & "FROM Tabla_Regimen " _
         & "WHERE Codigo <> 200 " _
         & "ORDER BY Codigo "
    SelectDB_Combo DCRegimen, AdoRegimen, sSQL, "Codigo"
End Sub

Public Sub Carga_ConceptosRetencion(MBFecha As String)
Dim FechaCodAir As String
  FechaCodAir = BuscarFecha(MBFecha)
 'Carga la Tabla de Porcentaje Iva
  sSQL = "SELECT * " _
       & "FROM Tabla_Por_ICE_IVA " _
       & "WHERE IVA <> " & Val(adFalse) & " " _
       & "AND Fecha_Inicio <= #" & FechaCodAir & "# " _
       & "AND Fecha_Final >= #" & FechaCodAir & "# " _
       & "ORDER BY Porc DESC "
  SelectDB_Combo DCPorcenIva, AdoPorIva, sSQL, "Porc"
 'Carga los Porcentajes de ICE
  sSQL = "SELECT * " _
       & "FROM Tabla_Por_ICE_IVA " _
       & "WHERE ICE <> " & Val(adFalse) & " " _
       & "AND Fecha_Inicio <= #" & FechaCodAir & "# " _
       & "AND Fecha_Final >= #" & FechaCodAir & "# " _
       & "ORDER BY Porc"
  SelectDB_Combo DCPorcenIce, AdoPorIce, sSQL, "Porc"
 'Carga los conceptos de Retencion segun la fecha de Registro
  sSQL = "SELECT (Codigo & ' - ' & Concepto) As Detalle_Conceptos,* " _
       & "FROM Tipo_Concepto_Retencion " _
       & "WHERE Codigo <> '.' " _
       & "AND Fecha_Inicio <= #" & FechaCodAir & "# " _
       & "AND Fecha_Final >= #" & FechaCodAir & "# " _
       & "ORDER BY Codigo "
  SelectDB_Combo DCConceptoRet, AdoConceptoRet, sSQL, "Detalle_Conceptos"
  With AdoConceptoRet.Recordset
      If AdoConceptoRet.Recordset.RecordCount > 0 Then Rf = 1 Else Rf = 0
  End With
  DCConceptoRet = "329 - Por Otros Servicios (N)"
End Sub

Public Sub Captura_TipoComprobante()
    'Captura lo que tiene el Combo de Tipo de Comprobante
    Captc = SinEspaciosIzq(DCTipoComprobante.Text)
    Cap1 = TrimStrg(DCTipoComprobante.Text)
    
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
    TxtTotalReten.Text = ""
    'Limpia la grilla
    ' Borra Asiento Air
    sSQL = "DELETE * " _
         & "FROM Asiento_Air " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND Tipo_Trans = 'I' " _
         & "AND T_No = " & Trans_No & " "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "SELECT * " _
         & "FROM Asiento_Air " _
         & "WHERE codRet <> '.' " _
         & "AND Tipo_Trans = 'I' " _
         & "AND T_No = " & Trans_No & " " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "ORDER BY codRet "
    Select_Adodc_Grid DGConceptoAir, AdoAsientoAir, sSQL, "Sustento"
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
     Validar_Porc_IVA MBFechaLiquida
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
 TextoValido TxtBimpConA, True, , 2
 TextoValido TxtSumatoria, True, , 2
 If DCImportacion = "Bienes" Then
   'Valida que la base imponible no sea mayor que la BIG y la BIcero
    If CTNumero(TxtBimpConA, 2) > CTNumero(TxtSumatoria, 2) Then
       MsgBox "La Base Imponible debe ser menor o igual a la " & vbCrLf _
            & "Base Imponible Gravada + la Base Imponible 0%", vbInformation, "Aviso"
       TxtBimpConA.Text = ""
       TxtBimpConA.SetFocus
    End If
 End If
 'Capturo el codigo de Conceptos Retencion
 If Not OP Then
   If (TxtBimpConA = "") Then
      MsgBox "Ingrese la Base Imponible que corresponda", vbInformation, "Aviso"
      TxtBimpConA.SetFocus
   Else
      TxtValConA = CTNumero(TxtBimpConA, 2) * (CTNumero(TxtPorRetConA, 2) / 100)
      Insertar_DataGrid
   End If
 End If
 RatonNormal
End Sub

Sub Insertar_DataGrid()
 'Selecciona el numero mayor para continuar la secuencia en el
 'campo T_No y A_No
 'Ac = 0
  sSQL = "SELECT TOP 1 * " _
       & "FROM Asiento_Air " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND Tipo_Trans = 'I' " _
       & "ORDER BY Cta_Retencion,A_No,ValRet "
  Select_Adodc AdoAsientoAir, sSQL
  If Val(CCur(TxtBimpConA)) > 0 Then
     RatonReloj
     Espizq = SinEspaciosIzq(DCConceptoRet)
     Espder = TrimStrg(MidStrg(DCConceptoRet, Len(Espizq) + 3, Len(DCConceptoRet)))
     SetAdoAddNew "Asiento_Air"
     SetAdoFields "CodRet", Espizq
     SetAdoFields "Detalle", Espder
     SetAdoFields "BaseImp", CTNumero(TxtBimpConA, 2)
     SetAdoFields "Porcentaje", Val(TxtPorRetConA) / 100
     SetAdoFields "ValRet", CTNumero(TxtValConA, 2)
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
     SetAdoFields "A_No", Maximo_De("Asiento_Air", "A_No")
     SetAdoFields "T_No", Trans_No
     SetAdoFields "Tipo_Trans", "I"
     SetAdoUpdate
    'Despliega los datos en el DataGrid
     sSQL = "SELECT * " _
          & "FROM Asiento_Air " _
          & "WHERE CodRet <> '.' " _
          & "AND Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' " _
          & "AND T_No = " & Trans_No & " " _
          & "AND Tipo_Trans = 'I' " _
          & "ORDER BY CodRet,Cta_Retencion,A_No,ValRet "
     Select_Adodc_Grid DGConceptoAir, AdoAsientoAir, sSQL, "Sustento"
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
     TxtSumatoria = CTNumero(TxtBaseImpo, 2) + CTNumero(TxtBaseImpoGrav, 2)
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

Private Sub TxtNumUnoAutComRet_LostFocus()
'    TextoValido TxtNumUnoAutComRet, True, , 0
'    If Val(TxtNumUnoAutComRet) <= 0 Then TxtNumUnoAutComRet = "0"
'        TxtNumUnoAutComRet = Format(Val(CCur(TxtNumUnoAutComRet)), String(10, "0"))
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
  If OP Then
     TxtValConA = CTNumero(TxtBimpConA, 2) * (CTNumero(TxtPorRetConA, 2) / 100)
     Insertar_DataGrid
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
    'Valido que sea ingresado 1 caracter
    Longitud = Len(TxtVerificador)
    If CInt(Longitud) < 1 Then
       MsgBox "El Correlativo consta de 1 caracter", vbInformation, "Aviso"
       TxtVerificador = ""
       TxtVerificador.SetFocus
    End If
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
    SelectDB_Combo DCRetFuente, AdoRetFuente, sSQL, "Cuentas"
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
    Ln_No = 0
    ac = 0
    SSTImportaciones.Tab = 0
    
    CodPorIva = 0
    CodPorIce = "0"
    CodRetBien = 0
    CodRetServ = 0

    TxtAño.Text = Year(date)
        
    'Enceramos el espacio de cada usuario para emprezar con una nueva retencion
    sSQL = "DELETE * " _
         & "FROM Asiento_Importaciones " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND T_No = " & Trans_No & " "
    Ejecutar_SQL_SP sSQL

    'Borra Asiento Air
    sSQL = "DELETE * " _
         & "FROM Asiento_Air  " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND Tipo_Trans = 'I' " _
         & "AND T_No = " & Trans_No & " "
    Ejecutar_SQL_SP sSQL

    'Carga la Tabla de Catalogos Tributarios al DataCombo
    Carga_CreditoTributario
   
    'Carga la Tabla de Tipo Importacion al DataCombo
    Carga_TipoImportacion
   
    'Carga la Tabla de Distrito al DataCombo
    Carga_Distrito
   
   ' Carga la Tabla de Regimen al DataCombo
    Carga_Regimen
         
   'Carga la Tabla de Conceptos Retencion al DataCombo
    Carga_ConceptosRetencion MBFechaLiquida
    
    'Carga la Retencion en la Fuente
    Carga_RetencionFuente
   
    sSQL = "SELECT CodRet, Detalle, BaseImp, Porcentaje, ValRet, EstabRetencion, PtoEmiRetencion, SecRetencion, AutRetencion, FechaEmiRet  " _
         & "FROM Asiento_Air " _
         & "WHERE CodRet <> '.' " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND CodigoU =  '" & CodigoUsuario & "' " _
         & "ORDER BY CodRet "
    Select_Adodc_Grid DGConceptoAir, AdoAsientoAir, sSQL
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
End Sub

Public Sub Habilita_Controles()
    'Habilito los controles para la modificacion
    SSTImportaciones.Enabled = True
    CmdGrabar.Enabled = True
    'Label23.Visible = True
End Sub

Public Sub Deshabilita_Controles()
    'Deshabilito los controles para la modificacion
    SSTImportaciones.Enabled = False
    CmdGrabar.Enabled = False
    'Label23.Visible = False
End Sub

