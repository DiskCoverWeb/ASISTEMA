VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FRecap 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RECAP"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   10920
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtNumeroC 
      Alignment       =   1  'Right Justify
      Height          =   336
      Left            =   9552
      MaxLength       =   7
      MultiLine       =   -1  'True
      TabIndex        =   13
      Text            =   "FRecap.frx":0000
      ToolTipText     =   "En este campo se debe ingresar el número del comprobante, el cual no excedera los siete caracteres"
      Top             =   744
      Width           =   975
   End
   Begin VB.ComboBox CTP 
      Height          =   288
      Left            =   8820
      TabIndex        =   11
      ToolTipText     =   "En este combo se despliega una lista con lo stipos de comprobantes existentes tales como: Comprobante Diario, Ingreso o Egreso"
      Top             =   744
      Width           =   660
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
      Height          =   300
      Left            =   3564
      TabIndex        =   4
      Top             =   108
      Visible         =   0   'False
      Width           =   6048
   End
   Begin VB.TextBox TxtAnio 
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
      Height          =   336
      Left            =   660
      MaxLength       =   4
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "FRecap.frx":0004
      Top             =   108
      Visible         =   0   'False
      Width           =   876
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
      Height          =   288
      Left            =   2052
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   120
      Visible         =   0   'False
      Width           =   1416
   End
   Begin MSDataListLib.DataCombo DCProveedor 
      Bindings        =   "FRecap.frx":0009
      DataSource      =   "AdoClientes"
      Height          =   288
      Left            =   108
      TabIndex        =   6
      ToolTipText     =   "Razón o denomicación Social. Este campo es obligatorio."
      Top             =   756
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
   Begin TabDlg.SSTab SSTRecap 
      Height          =   4212
      Left            =   108
      TabIndex        =   14
      Top             =   1188
      Width           =   10416
      _ExtentX        =   18362
      _ExtentY        =   7435
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
      TabCaption(0)   =   "Detalle RECAP"
      TabPicture(0)   =   "FRecap.frx":0023
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label7"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label15"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "MBFechaEmision"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "DCTarjetaCred"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "DCTipoRecap"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "MBFechaPago"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "FraBases"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame5"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "TxtNumRecap"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "TxtNumVoucher"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "CmdCerrar"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "CmdGrabar"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "CmdAir"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "Conceptos AIR"
      TabPicture(1)   =   "FRecap.frx":003F
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FraRetencion"
      Tab(1).ControlCount=   1
      Begin VB.CommandButton CmdAir 
         Caption         =   "AIR"
         Height          =   444
         Left            =   9612
         Picture         =   "FRecap.frx":005B
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Se ubica en la pestaña de Retenciones"
         Top             =   3456
         Width           =   552
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&Grabar"
         Height          =   750
         Left            =   210
         Picture         =   "FRecap.frx":0581
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "Grabar"
         Top             =   3360
         Width           =   960
      End
      Begin VB.CommandButton CmdCerrar 
         Caption         =   "&Salir"
         Height          =   750
         Left            =   1365
         Picture         =   "FRecap.frx":088B
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "Salir"
         Top             =   3360
         Width           =   960
      End
      Begin VB.TextBox TxtNumVoucher 
         Height          =   330
         Left            =   8190
         TabIndex        =   20
         Top             =   630
         Width           =   624
      End
      Begin VB.TextBox TxtNumRecap 
         Height          =   330
         Left            =   6090
         MaxLength       =   15
         TabIndex        =   18
         Top             =   630
         Width           =   1590
      End
      Begin VB.Frame Frame5 
         Caption         =   "RETENCION DEL IVA POR BIENES Y/O SERVICIOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1596
         Left            =   5145
         TabIndex        =   38
         Top             =   1680
         Width           =   5055
         Begin VB.TextBox TxtIvaSerValRet 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   3132
            TabIndex        =   49
            Text            =   " "
            Top             =   1188
            Width           =   1632
         End
         Begin VB.TextBox TxtIvaSerMonIva 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   3132
            MultiLine       =   -1  'True
            TabIndex        =   47
            Text            =   "FRecap.frx":0CCD
            ToolTipText     =   $"FRecap.frx":0CD2
            Top             =   432
            Width           =   1632
         End
         Begin VB.TextBox TxtIvaBienValRet 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   1404
            TabIndex        =   45
            Top             =   1188
            Width           =   1632
         End
         Begin VB.TextBox TxtIvaBienMonIva 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   1404
            MultiLine       =   -1  'True
            TabIndex        =   41
            Text            =   "FRecap.frx":0D68
            ToolTipText     =   $"FRecap.frx":0D6D
            Top             =   432
            Width           =   1632
         End
         Begin MSDataListLib.DataCombo DCPorcenRetenIvaServ 
            Bindings        =   "FRecap.frx":0E0C
            DataSource      =   "AdoRetIvaServicios"
            Height          =   288
            Left            =   3132
            TabIndex        =   48
            ToolTipText     =   "Corresponde al porcentaje retenido en el IVA generado en la prestación de servicios"
            Top             =   840
            Width           =   1632
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DCPorcenRetenIvaBien 
            Bindings        =   "FRecap.frx":0E2D
            DataSource      =   "AdoRetIvaBienes"
            Height          =   288
            Left            =   1404
            TabIndex        =   43
            ToolTipText     =   $"FRecap.frx":0E4B
            Top             =   840
            Width           =   1632
            _ExtentX        =   2858
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSForms.Label Label22 
            Height          =   228
            Left            =   108
            TabIndex        =   44
            Top             =   1236
            Width           =   1272
            Caption         =   "Valor retenido"
            Size            =   "2249;397"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label20 
            Height          =   228
            Left            =   108
            TabIndex        =   40
            Top             =   432
            Width           =   1308
            Caption         =   "Monto de IVA"
            Size            =   "2307;402"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label18 
            Height          =   228
            Left            =   108
            TabIndex        =   42
            Top             =   804
            Width           =   1272
            Caption         =   "% Retención IVA"
            Size            =   "2249;397"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label16 
            Height          =   228
            Left            =   3132
            TabIndex        =   46
            Top             =   216
            Width           =   1416
            Caption         =   "IVA-SERVICIOS"
            Size            =   "2498;402"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label19 
            Height          =   228
            Left            =   1404
            TabIndex        =   39
            Top             =   216
            Width           =   1092
            Caption         =   "IVA-BIENES"
            Size            =   "1926;402"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
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
         Height          =   1485
         Left            =   210
         TabIndex        =   27
         Top             =   1785
         Width           =   4845
         Begin VB.TextBox TxtTotalConsumo 
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
            Height          =   330
            Left            =   3240
            TabIndex        =   33
            Text            =   "0.00"
            Top             =   432
            Width           =   1416
         End
         Begin VB.TextBox TxtComision 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   108
            TabIndex        =   35
            Text            =   "0.00"
            Top             =   972
            Width           =   1416
         End
         Begin VB.TextBox TxtMontoIva 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   1620
            TabIndex        =   37
            Text            =   "0.00"
            ToolTipText     =   $"FRecap.frx":0ED7
            Top             =   972
            Width           =   1308
         End
         Begin VB.TextBox TxtConsumo 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   105
            MaxLength       =   14
            MultiLine       =   -1  'True
            TabIndex        =   29
            Text            =   "FRecap.frx":0FC1
            ToolTipText     =   "Corresponde al valor de la importación gravada con tarifa 0% o exento"
            Top             =   432
            Width           =   1416
         End
         Begin VB.TextBox TxtConsumoGrav 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   1620
            MultiLine       =   -1  'True
            TabIndex        =   31
            Text            =   "FRecap.frx":0FC8
            ToolTipText     =   "Corresponde al valor de la importación gravada con tarifa 12% en el momento de la desaduanización"
            Top             =   432
            Width           =   1416
         End
         Begin MSForms.Label Label13 
            Height          =   228
            Left            =   108
            TabIndex        =   34
            Top             =   756
            Width           =   768
            Caption         =   "Comisión"
            Size            =   "1355;402"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label10 
            Height          =   228
            Left            =   3240
            TabIndex        =   32
            Top             =   216
            Width           =   1308
            Caption         =   "Total Consumo"
            Size            =   "2307;402"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label14 
            Height          =   228
            Left            =   1620
            TabIndex        =   36
            Top             =   756
            Width           =   1200
            Caption         =   "Monto de I.V.A."
            Size            =   "2117;402"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label44 
            Height          =   228
            Left            =   108
            TabIndex        =   28
            Top             =   216
            Width           =   1416
            Caption         =   "Consumo 0%"
            Size            =   "2498;402"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label11 
            Height          =   228
            Left            =   1620
            TabIndex        =   30
            Top             =   216
            Width           =   1524
            Caption         =   "Consumo Gravado"
            Size            =   "2688;402"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
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
         Height          =   3810
         Left            =   -74895
         TabIndex        =   51
         Top             =   420
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
            TabIndex        =   52
            Top             =   315
            Width           =   2328
         End
         Begin VB.TextBox TxtValConA 
            Enabled         =   0   'False
            Height          =   336
            Left            =   8715
            TabIndex        =   69
            Top             =   1470
            Width           =   1275
         End
         Begin VB.TextBox TxtPorRetConA 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   336
            Left            =   7980
            TabIndex        =   67
            Top             =   1470
            Width           =   645
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
            Left            =   8715
            TabIndex        =   72
            Top             =   3360
            Width           =   1275
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
            Left            =   8085
            MultiLine       =   -1  'True
            TabIndex        =   61
            Top             =   735
            Width           =   1905
         End
         Begin VB.TextBox TxtBimpConA 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   336
            Left            =   6510
            TabIndex        =   65
            Top             =   1470
            Width           =   1380
         End
         Begin VB.TextBox TxtNumUnoAutComRet 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   330
            Left            =   2310
            MaxLength       =   10
            MultiLine       =   -1  'True
            TabIndex        =   59
            Top             =   840
            Width           =   1590
         End
         Begin VB.TextBox TxtNumTresComRet 
            Enabled         =   0   'False
            Height          =   336
            Left            =   1365
            MaxLength       =   7
            TabIndex        =   57
            Text            =   "0000001"
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox TxtNumDosComRet 
            Enabled         =   0   'False
            Height          =   336
            Left            =   735
            MaxLength       =   3
            TabIndex        =   56
            Text            =   "001"
            Top             =   840
            Width           =   540
         End
         Begin VB.TextBox TxtNumUnoComRet 
            Enabled         =   0   'False
            Height          =   336
            Left            =   105
            MaxLength       =   3
            TabIndex        =   55
            Text            =   "001"
            Top             =   840
            Width           =   540
         End
         Begin MSDataListLib.DataCombo DCConceptoRet 
            Bindings        =   "FRecap.frx":0FCF
            DataSource      =   "AdoConceptoRet"
            Height          =   315
            Left            =   105
            TabIndex        =   63
            Top             =   1470
            Width           =   6315
            _ExtentX        =   11139
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
            Bindings        =   "FRecap.frx":0FEC
            Height          =   1275
            Left            =   105
            TabIndex        =   70
            Top             =   1995
            Width           =   9885
            _ExtentX        =   17436
            _ExtentY        =   2249
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
            Bindings        =   "FRecap.frx":1008
            DataSource      =   "AdoRetFuente"
            Height          =   315
            Left            =   2520
            TabIndex        =   53
            Top             =   315
            Visible         =   0   'False
            Width           =   7470
            _ExtentX        =   13176
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSForms.Label Label42 
            Height          =   225
            Left            =   7035
            TabIndex        =   71
            Top             =   3360
            Width           =   1695
            Caption         =   "Total Retenciones:"
            Size            =   "2990;397"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label37 
            Height          =   225
            Left            =   5355
            TabIndex        =   60
            Top             =   810
            Width           =   2535
            Caption         =   "Base Imponible de Retención:"
            Size            =   "4471;397"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
         Begin MSForms.Label Label21 
            Height          =   225
            Left            =   105
            TabIndex        =   54
            Top             =   630
            Width           =   2220
            Caption         =   "No. de Serie y Secuencial"
            Size            =   "3916;397"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label28 
            Height          =   225
            Left            =   2310
            TabIndex        =   58
            Top             =   630
            Width           =   1695
            Caption         =   "No. de Autorización"
            Size            =   "2990;397"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label24 
            Height          =   225
            Left            =   8715
            TabIndex        =   68
            Top             =   1260
            Width           =   1275
            Caption         =   "Valor Retenido"
            Size            =   "2249;397"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label25 
            Height          =   225
            Left            =   6510
            TabIndex        =   64
            Top             =   1260
            Width           =   1380
            Caption         =   "Base Imponible"
            Size            =   "2434;397"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label26 
            Height          =   225
            Left            =   7980
            TabIndex        =   66
            Top             =   1260
            Width           =   645
            Caption         =   "% Ret."
            Size            =   "1138;397"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label27 
            Height          =   225
            Left            =   105
            TabIndex        =   62
            Top             =   1260
            Width           =   6420
            Caption         =   "RETENCION EN LA FUENTE DEL IMPUESTO A  LA RENTA "
            Size            =   "11324;397"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin MSMask.MaskEdBox MBFechaPago 
         Height          =   336
         Left            =   8208
         TabIndex        =   26
         ToolTipText     =   "En este campo se ingresa la fecha de liquidación del comprobante en el Banco"
         Top             =   1260
         Width           =   1176
         _ExtentX        =   2090
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
      Begin MSDataListLib.DataCombo DCTipoRecap 
         Bindings        =   "FRecap.frx":1023
         DataSource      =   "AdoTipoComprobante"
         Height          =   312
         Left            =   216
         TabIndex        =   16
         ToolTipText     =   "Corresponde al tipo de comprobante utilizado en la transacción"
         Top             =   648
         Width           =   4956
         _ExtentX        =   8758
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
      Begin MSDataListLib.DataCombo DCTarjetaCred 
         Bindings        =   "FRecap.frx":1044
         DataSource      =   "AdoTarjetas"
         Height          =   315
         Left            =   210
         TabIndex        =   22
         ToolTipText     =   "Corresponde al tipo de comprobante utilizado en la transacción"
         Top             =   1260
         Width           =   3060
         _ExtentX        =   5398
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
      Begin MSMask.MaskEdBox MBFechaEmision 
         Height          =   336
         Left            =   6048
         TabIndex        =   24
         ToolTipText     =   "En este campo se ingresa la fecha de liquidación del comprobante en el Banco"
         Top             =   1260
         Width           =   1176
         _ExtentX        =   2090
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
      Begin MSForms.Label Label15 
         Height          =   225
         Left            =   8190
         TabIndex        =   19
         Top             =   420
         Width           =   1485
         Caption         =   "No. de Vouchers"
         Size            =   "2619;397"
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label7 
         Height          =   228
         Left            =   6048
         TabIndex        =   23
         Top             =   1080
         Width           =   1596
         Caption         =   "Fecha de Emisión"
         Size            =   "2805;397"
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label6 
         Height          =   228
         Left            =   8208
         TabIndex        =   25
         Top             =   1056
         Width           =   1380
         Caption         =   "Fecha de Pago"
         Size            =   "2434;397"
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label4 
         Height          =   225
         Left            =   210
         TabIndex        =   21
         Top             =   1050
         Width           =   1275
         Caption         =   "Tarjeta Crédito"
         Size            =   "2249;397"
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   225
         Left            =   6090
         TabIndex        =   17
         Top             =   420
         Width           =   1170
         Caption         =   "No. de Recap"
         Size            =   "2064;397"
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label3 
         Height          =   228
         Left            =   216
         TabIndex        =   15
         Top             =   420
         Width           =   1044
         Caption         =   "Tipo Recap"
         Size            =   "1841;402"
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin MSAdodcLib.Adodc AdoSustento 
      Height          =   330
      Left            =   2730
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
   Begin MSAdodcLib.Adodc AdoTipoIdentificacion 
      Height          =   330
      Left            =   210
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
      Caption         =   "AdoTipoIden"
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
   Begin MSAdodcLib.Adodc AdoTipoComprobante 
      Height          =   330
      Left            =   210
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
   Begin MSAdodcLib.Adodc AdoRetIvaBienes 
      Height          =   330
      Left            =   210
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
      Caption         =   "AdoRetenBienes"
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
   Begin MSAdodcLib.Adodc AdoRetIvaServicios 
      Height          =   330
      Left            =   210
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
      Caption         =   "AdoRetenServicios"
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
      Left            =   2730
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
      Left            =   2730
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
   Begin MSAdodcLib.Adodc AdoConceptoRet 
      Height          =   330
      Left            =   210
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
   Begin MSAdodcLib.Adodc AdoTransRecap 
      Height          =   330
      Left            =   210
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
      Left            =   210
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
   Begin MSAdodcLib.Adodc AdoClientes 
      Height          =   330
      Left            =   2730
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
   Begin MSAdodcLib.Adodc AdoTarjetas 
      Height          =   330
      Left            =   2730
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
      Caption         =   "AdoTarjetasCredito"
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
   Begin MSAdodcLib.Adodc AdoAsientoCompras 
      Height          =   330
      Left            =   210
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
      Caption         =   "AsientoCompras"
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
      Top             =   5040
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
      Left            =   210
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   2730
      Top             =   3465
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
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
   Begin MSForms.Label Label17 
      Height          =   228
      Left            =   6936
      TabIndex        =   8
      Top             =   528
      Width           =   1692
      Caption         =   "No. de Identificación"
      Size            =   "2990;397"
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
      Height          =   336
      Left            =   6612
      TabIndex        =   7
      Top             =   744
      Width           =   336
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
      TabIndex        =   9
      Top             =   744
      Width           =   1800
   End
   Begin MSForms.Label Label40 
      Height          =   228
      Left            =   9552
      TabIndex        =   12
      Top             =   528
      Width           =   768
      Caption         =   "Número"
      Size            =   "1355;402"
      FontName        =   "Arial"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label39 
      Height          =   228
      Left            =   8820
      TabIndex        =   10
      Top             =   528
      Width           =   552
      Caption         =   "Tipo"
      Size            =   "974;402"
      FontName        =   "Arial"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label29 
      Height          =   228
      Left            =   108
      TabIndex        =   0
      Top             =   216
      Visible         =   0   'False
      Width           =   432
      Caption         =   "Año:"
      Size            =   "762;402"
      FontName        =   "Arial"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   228
      Left            =   1620
      TabIndex        =   2
      Top             =   216
      Visible         =   0   'False
      Width           =   432
      Caption         =   "Mes:"
      Size            =   "762;402"
      FontName        =   "Arial"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label41 
      Height          =   225
      Left            =   105
      TabIndex        =   5
      Top             =   525
      Width           =   2115
      Caption         =   "Proveedor/Razón Social"
      Size            =   "3731;397"
      FontName        =   "Arial"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FRecap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MBFecha As MaskEdBox
Public cod, x, CodImp, CodTs, Longitud, codTC As Byte
Public Cap, Cap1, Ch, CodDis, CodProv, Opc, Ct As String
Dim CalmIva, CalmIce, CalIbMi, CalIsMi, ac, SUM, cal As Double
Dim Espizq, Espder, Captc, PorIva, TipoImp, CodSus As String
Dim SumAnio, Aniocad, AniocadAux, CodPorIva, CodRetBien, CodRetServ, CodReg As Integer

Private Sub ChRetF_Click()
  If ChRetF.Value <> 0 Then
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

Private Sub CmdCerrar_Click()
    If CmdCerrar.Caption = "Cerrar" Then
        Unload Me
    Else
        'Borra Asiento Compras
        sSQL = "DELETE * " _
             & "FROM Asiento_Importaciones " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND CodigoU = '" & CodigoUsuario & "' "
        'MsgBox "--> " & sSQL
        ConectarAdoExecute sSQL
        
        'Borra Asiento Air
        sSQL = "DELETE * " _
             & "FROM Asiento_Air " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND CodigoU = '" & CodigoUsuario & "' "
             
        'MsgBox "=====> " & sSQL
        ConectarAdoExecute sSQL
        
        sSQL = "SELECT CodRet, Detalle, BaseImp, Porcentaje, ValRet, EstabRetencion, PtoEmiretencion, SecRetencion, AutRetencion, FechaEmiRet  " _
           & "FROM Asiento_Air " _
           & "WHERE CodRet <> '.' " _
           & "ORDER BY CodRet "
        SelectDataGrid DGConceptoAir, AdoAsientoAir, sSQL
    End If
    Unload Me

End Sub

Private Sub CmdGrabar_Click()
    RatonReloj
    
    'Valido por si acaso exista algun valor con 0
    TextoValido TxtConsumo, True, , 2
    TextoValido TxtConsumoGrav, True, , 2
    TextoValido TxtTotalConsumo, True, , 2
    TextoValido TxtComision, True, , 2
    TextoValido TxtMontoIva, True, , 2
    FechaValida MBFechaEmision
    FechaValida MBFechaPago

    Ln_No = 0
   'Borra si encuentra 2 o mas transacciones iguales
    sSQL = "DELETE * " _
         & "FROM Trans_Recap " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND IdProv = '" & CodigoCliente & "' "
    ConectarAdoExecute sSQL
    
    If ChRetF.Value <> 0 Then
         sSQL = "DELETE * " _
              & "FROM Trans_Air " _
              & "WHERE Item = '" & NumEmpresa & "' " _
              & "AND Periodo = '" & Periodo_Contable & "' " _
              & "AND IdProv = '" & CodigoCliente & "' " _
              & "AND EstabRetencion = '" & TxtNumUnoComRet & "' " _
              & "AND PtoEmiRetencion = '" & TxtNumDosComRet & "' " _
              & "AND SecRetencion = '" & TxtNumTresComRet & "' " _
              & "AND AutRetencion = '" & TxtNumUnoAutComRet & "' "
         ConectarAdoExecute sSQL
     End If
     
    'Pregunto antes de grabar
    Titulo = "Aviso"
    Mensajes = "Desea Grabar los Datos"
    If BoxMensaje = vbYes Then
      'Grabacion de los Datos
       Grabacion
       Titulo = "Aviso"
       Mensajes = "Los Datos fueron grabados correctamente" & vbCrLf _
                & "Desea ingresar otra transacción"""
       If BoxMensaje = vbYes Then
           Limpiar_Controles
           SSTRecap.Tab = 0
           DCProveedor.SetFocus
        Else
           Unload FRecap
        End If
    Else
       DCProveedor.SetFocus
    End If
End Sub

Private Sub CModificacion_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub CModificacion_LostFocus()
    CodImp = 0
    CodPorIva = 0
    CodRetBien = 0
    CodRetServ = 0
       
    'Cargo los datos para la modificación
    I = Val(SinEspaciosIzq(CModificacion))
    sSQL = "SELECT * " _
         & "FROM Trans_Recap " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "ORDER BY Item "
    SelectAdodc AdoTransRecap, sSQL

    With AdoTransRecap.Recordset
       If .RecordCount > 0 Then
          .MoveFirst
          .Find ("Linea_SRI = '" & I & "' ")
          If Not .EOF Then
             'Busco el Proveedor
              CodProv = .Fields("IdFiscalProv")
              If AdoClientes.Recordset.RecordCount > 0 Then
                 AdoClientes.Recordset.MoveFirst
                 AdoClientes.Recordset.Find ("Codigo = '" & CodProv & "' ")
                 If Not AdoClientes.Recordset.EOF Then
                    DCProveedor = AdoClientes.Recordset.Fields("Cliente")
                    LblTD = AdoClientes.Recordset.Fields("TD")
                    LblNumIdent = AdoClientes.Recordset.Fields("CI_RUC")
                    
                     
                     'MBFechaLiquida = .Fields("FechaLiquidacion")
                     PorIva = ""
                     'Cargo el IVA
                     'Carga_PorcentajeIva (MBFechaLiquida)
                     CodPorIva = .Fields("PorcentajeIva") 'Es un número
                     PorIva = CStr(CodPorIva) 'Convierto a string
                     If AdoPorIva.Recordset.RecordCount > 0 Then
                        AdoPorIva.Recordset.MoveFirst
                        AdoPorIva.Recordset.Find ("Codigo = '" & PorIva & "' ")
                        If Not .EOF Then
        '                   DCPorcenIva = AdoPorIva.Recordset.Fields("Porc")
                        Else
                           MsgBox "Porcentaje de IVA no existe", vbInformation, "Aviso"
                         End If
                      End If
                                  
                     
                     
                    TxtNumeroC = .Fields("Numero")
                    
                    
                    TxtMontoIva = .Fields("MontoIva")
                    CTP.AddItem .Fields("TP")
                     
                    sSQL = "SELECT TA.*, TI.* " _
                          & "FROM Trans_Air As TA, Trans_Importaciones As TI " _
                          & "WHERE TA.IdProv = TI.IdFiscalProv " _
                          & "AND TA.Linea_SRI = " & I & " " _
                          & "AND TA.Numero = TI.Numero "
                     SelectDataGrid DGConceptoAir, AdoAsientoAir, sSQL, "Sustento"
                    
                     'Pongo la Base Imponible
                     'TxtSumatoria = Val(CCur(TxtBaseImpo)) + Val(CCur(TxtBaseImpoGrav))
                 Else
                    MsgBox "Este beneficiario no existe", vbInformation, "Aviso"
                 End If
              End If

             
           Else
              MsgBox "Vuelva a seleccionar", vbInformation, "Aviso"
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

Private Sub DCPorcenRetenIvaBien_LostFocus()
If Not IsNumeric(DCPorcenRetenIvaBien) Then
       MsgBox "No ingrese caracteres. Vuelva a seleccionar.", vbInformation, "Aviso"
       DCPorcenRetenIvaBien = ""
       Carga_RetencionIvaBienes
       DCPorcenRetenIvaBien.SetFocus
    Else
       'Busca y captura el codigo de Porcentaje retencion Iva Bienes
       With AdoRetIvaBienes.Recordset
            If .RecordCount > 0 Then
               .MoveFirst
               .Find ("Porc = " & SinEspaciosDer(DCPorcenRetenIvaBien) & " ")
               CodRetBien = .Fields("Codigo")
            End If
        End With
            
        Total_IVA = 0
        Total_IVA = Convertir_Numero(TxtIvaBienMonIva, 2)
        TxtIvaBienValRet = 0
       'Calcula la retencion Iva Bienes
        CalIbMi = (Total_IVA * CInt(DCPorcenRetenIvaBien)) / 100
        TxtIvaBienValRet = CalIbMi
    End If
    TxtIvaSerMonIva = Format(Convertir_Numero(TxtMontoIva, 2) - Convertir_Numero(TxtIvaBienMonIva, 2), "#,##0.00")
End Sub

Private Sub DCPorcenRetenIvaServ_LostFocus()
   'Activo el casillero para que ingrese el valor si el porcentaje es 70/100
    If DCPorcenRetenIvaServ = "70/100" Then
       Ct = "Si"
       TxtIvaSerValRet.Text = ""
       TxtIvaSerValRet.Enabled = True
    Else
      If Not IsNumeric(DCPorcenRetenIvaServ) Then
         MsgBox "No ingrese caracteres. Vuelva a seleccionar.", vbInformation, "Aviso"
         DCPorcenRetenIvaServ = ""
         Carga_RetencionIvaServicios
         DCPorcenRetenIvaServ.SetFocus
      End If
    End If
    
   'Busca captura el codigo de Porcentaje retencion Iva Servicios
    With AdoRetIvaServicios.Recordset
     If .RecordCount > 0 Then
        .MoveFirst
        .Find ("Porc = " & SinEspaciosDer(DCPorcenRetenIvaServ) & " ")
        CodRetServ = .Fields("Codigo")
     Else
         MsgBox "Código erróneo", vbInformation, "Aviso"
     End If
    End With
    Ct = "No"
    Total_IVA = 0
    Total_IVA = Convertir_Numero(TxtIvaSerMonIva, 2)
    TxtIvaSerValRet = 0
    If DCPorcenRetenIvaServ = "70/100" Then
    Else
       CalIsMi = (Total_IVA) * CCur(DCPorcenRetenIvaServ) / 100
       TxtIvaSerValRet = CalIsMi
       TxtIvaSerValRet.Enabled = False
    End If
End Sub

Private Sub DCProveedor_LostFocus()
If IsNumeric(DCProveedor.Text) Then
    MsgBox "No ingrese números. Vuelva a seleccionar.", vbInformation, "Aviso"
    DCProveedor.Text = ""
    Leer_Clientes
    DCProveedor.SetFocus
 Else
    NombreCliente = UCase(DCProveedor)
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

Private Sub DCTipoRecap_LostFocus()
    If IsNumeric(DCTipoRecap.Text) Then
       MsgBox "No ingrese números. Vuelva a seleccionar.", vbInformation, "Aviso"
       DCTipoRecap.Text = ""
       DCTipoRecap.SetFocus
       Captura_TipoComprobante
    Else
       If DCTipoRecap <> "" Then
           Captura_TipoComprobante
       End If
    End If
End Sub

Private Sub DGConceptoAir_Click()
    Titulo = "Aviso"
    Mensajes = "Desea Eliminar la Retención"
    If BoxMensaje = vbYes Then
       Calculo_Sumatoria
    End If
End Sub

Private Sub Form_Activate()
    'Cargo los datos en los combos
    Carga_Datos_Iniciales MBFecha, Nuevo
End Sub

Private Sub Form_Load()
    CentrarForm FRecap
    ConectarAdodc AdoSustento
    ConectarAdodc AdoTipoComprobante
    ConectarAdodc AdoTarjetas
    ConectarAdodc AdoConceptoRet
    ConectarAdodc AdoRetFuente
    ConectarAdodc AdoPorIce
    ConectarAdodc AdoPorIva
    ConectarAdodc AdoClientes
    ConectarAdodc AdoAsientoAir
    'ConectarAdodc AdoAsientoRecap
    ConectarAdodc AdoTransRecap
    ConectarAdodc AdoTransAir
    ConectarAdodc AdoAux
End Sub
Public Sub Carga_Tarjetas()
    'Capturo el codigo del Tipo de Catalogo Tributario
    'CapTar = CArgaTC
            
    'Busco el codigo en la tabla Tipo Comprobante///descripcion
     sSQL = "SELECT * " _
          & "FROM Tabla_Tarjetas_Credito " _
          & "WHERE Tarjeta_Credito_Codigo <> 0 " _
          & "ORDER BY Descripcion "
     SelectDBCombo DCTarjetaCred, AdoTarjetas, sSQL, "Descripcion"
End Sub

Public Sub Modificacion()
   Opc = "NO"
   'Verifico si debe o no renumerar
   'Busco Linea_SRI para la modifcación
''    sSQL = "SELECT Linea_SRI " _
''         & "FROM Trans_Importaciones " _
''         & "WHERE Item = '" & NumEmpresa & "' " _
''         & "AND Periodo = '" & Periodo_Contable & "' " _
''         & "AND Linea_SRI = 0 "
''    SelectAdodc AdoTransImportaciones, sSQL
''    With AdoTransImportaciones.Recordset
''     If .RecordCount > 0 Then
''       Do While Not .EOF
''          'MsgBox .Fields("Linea_SRI")
''          If .Fields("Linea_SRI") = 0 Then
''              Opc = "SI"
''          End If
''         .MoveNext
''       Loop
''     End If
''    End With
    
    If Opc = "SI" Then
       MsgBox "Tiene que generar el Talón Resumen", vbInformation, "Aviso"
       Deshabilita_Controles
    Else
        Habilita_Controles
        'Despliego los datos para modificar
        sSQL = "SELECT TI.*,C.CI_RUC,C.Cliente,C.TD " _
             & "FROM Trans_Importaciones As TI, Clientes As C " _
             & "WHERE C.Codigo = TI.IdFiscalProv " _
             & "ORDER BY TI.Linea_SRI,C.CI_RUC "
        SelectAdodc AdoAux, sSQL
        With AdoAux.Recordset
         If .RecordCount > 0 Then
             Do While Not .EOF
                Cadena = Format(.Fields("Linea_SRI"), "000") & " "
                Cadena = Cadena & .Fields("Cliente") & Space(55 - Len(.Fields("Cliente")))
                Cadena = Cadena & .Fields("CI_RUC") & Space(14 - Len(.Fields("CI_RUC")))
                Cadena = Cadena & .Fields("FechaLiquidacion") & " "
                Cadena = Cadena & .Fields("ValorCIF") & " "
                CModificacion.AddItem Cadena
               .MoveNext
              Loop
              CModificacion.Text = CModificacion.List(0)
          End If
        End With
   End If
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
    SSTRecap.Tab = 0
''    DCSustento.Text = ""
''    DCImportacion.Text = ""
''    FechaValida MBFechaLiquida
''    DCTipoComprobante.Text = ""
''    DCDistrito.Text = ""
''    TxtAño.Text = ""
''    DCRegimen.Text = ""
''    TxtCorrelativo.Text = ""
''    TxtVerificador.Text = ""
''    DCProveedor.Text = ""
''    LblTD.Caption = ""
''    TxtValorCIF.Text = ""
''    TxtBaseImpo.Text = ""
''    TxtBaseImpoGrav.Text = ""
''    DCPorcenIva.Text = ""
''    TxtMontoIva.Text = ""
''    TxtBaseImpoIce.Text = ""
''    DCPorcenIce.Text = ""
''    TxtMontoIce.Text = ""
''    DCRetFuente.Text = ""
''    TxtNumUnoComRet.Text = ""
''    TxtNumDosComRet.Text = ""
''    TxtNumTresComRet.Text = ""
''    TxtNumUnoAutComRet.Text = ""
''    TxtSumatoria.Text = ""
''    DCConceptoRet.Text = ""
''    TxtBimpConA.Text = ""
''    TxtPorRetConA.Text = ""
''    TxtValConA.Text = ""
''    TxtTotalReten.Text = ""
''    LblNumIdent.Caption = ""
''    TxtNumeroC.Text = ""
''    TxtTotalReten.Text = ""
    CTP.AddItem "CE"
    CTP.AddItem "CI"
    CTP.AddItem "CD"
    CTP.Text = "CE"
    
    'Limpia la grilla
    ' Borra Asiento Air
    sSQL = "DELETE * " _
         & "FROM Asiento_Air " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' "
    ConectarAdoExecute sSQL
    
    sSQL = "SELECT * " _
         & "FROM Asiento_Air " _
         & "WHERE codRet <> '.' " _
         & "ORDER BY codRet "
    SelectDataGrid DGConceptoAir, AdoAsientoAir, sSQL, "Sustento"
End Sub

Public Sub Calculo_Sumatoria()
    If (DGConceptoAir.Columns(4) = "") Or (DGConceptoAir.Columns(4) = 0) Then
        MsgBox "No existen datos ", vbInformation, "Aviso"
    Else
        cal = DGConceptoAir.Columns(4)
        SUM = CDbl(ac) - CDbl(cal)
        ac = SUM
        TxtTotalReten = ac
    End If
End Sub

Private Sub MBFechaEmision_GotFocus()
    MarcarTexto MBFechaEmision
End Sub

Private Sub MBFechaEmision_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub MBFechaEmision_LostFocus()
   FechaValida MBFechaEmision
    
  'Controla que la Fecha de Registro este entre 01/01/2000 en adelante
   If CFechaLong(MBFechaEmision) < CFechaLong("01/01/2000") Then
      MsgBox "La Fecha de Emisión debe ser mayor que 01/01/2000", vbInformation, "Aviso"
      MBFechaEmision = "01/01/2000"
      MBFechaEmision.SetFocus
   End If
End Sub

Private Sub MBFechaPago_GotFocus()
    MarcarTexto MBFechaPago
End Sub

Private Sub MBFechaPago_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub MBFechaPago_LostFocus()
    FechaValida MBFechaPago
    
   'Controla que la Fecha de Pago este entre 01/01/2000 en adelante
   If CFechaLong(MBFechaPago) < CFechaLong("01/01/2000") Then
      MsgBox "La Fecha de Pago debe ser mayor que 01/01/2000", vbInformation, "Aviso"
      MBFechaPago = "01/01/2000"
      MBFechaPago.SetFocus
   Else
      If MBFechaPago < MBFechaEmision Then
         MsgBox "La Fecha de Pago debe ser mayor o igual que la Fecha de Emisión", vbInformation, "Aviso"
         MBFechaPago.SetFocus
      End If
   End If
   FechaValida MBFechaPago
End Sub

Private Sub TxtBimpConA_GotFocus()
   MarcarTexto TxtBimpConA
End Sub

Private Sub TxtBimpConA_LostFocus()
 TextoValido TxtBimpConA, True, , 2
 'Capturo el codigo de Conceptos Retencion
 Espizq = SinEspaciosIzq(DCConceptoRet)
 Espder = SinCodigoIzq(DCConceptoRet)
 With AdoConceptoRet.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
       .Find ("Codigo = '" & Espizq & "' ")
        If Not .EOF Then
           If (TxtBimpConA = "") Then
              MsgBox "Ingrese la Base Imponible que corresponda", vbInformation, "Aviso"
              TxtBimpConA.SetFocus
           Else
              TxtPorRetConA = .Fields("Porcentaje")
              TxtValConA = (Val(CCur(TxtBimpConA)) * .Fields("Porcentaje")) / 100
              Insertar_DataGrid
              DCConceptoRet.SetFocus
           End If
      Else
          MsgBox "Tiene que seleccionar un código de retención", vbInformation, "Aviso"
          DCConceptoRet.SetFocus
      End If
    End If
 End With
 RatonNormal
End Sub

Sub Insertar_DataGrid()
    'Verifico si se activo la retención a la fuente
    If ChRetF.Value <> 0 Then
        'Selecciona el numero mayor para continuar la secuencia en el
        'campo T_No y A_No
        sSQL = "SELECT TOP 1 * " _
             & "FROM Asiento_Air " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "ORDER BY A_No DESC "
        SelectAdodc AdoAsientoAir, sSQL
        If AdoAsientoAir.Recordset.RecordCount > 0 Then Trans_No = AdoAsientoAir.Recordset.Fields("T_No") + 1
           If AdoAsientoAir.Recordset.RecordCount > 0 Then Ln_No = AdoAsientoAir.Recordset.Fields("A_No") + 1
             If Val(CCur(TxtBimpConA)) > 0 Then
                RatonReloj
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
                'SetAdoFields "FechaEmiRet", MBFechaLiquida
                SetAdoFields "EstabFactura", "001"
                SetAdoFields "PuntoEmiFactura", "001"
                'SetAdoFields "Factura_No", TxtCorrelativo
                SetAdoFields "Cta_Retencion", SinEspaciosIzq(DCRetFuente)
                SetAdoFields "IdProv", CodigoCliente
                SetAdoFields "TP", CTP
                SetAdoFields "Numero", TxtNumeroC
                SetAdoFields "A_No", Ln_No
                SetAdoFields "T_No", Trans_No
                SetAdoFields "Tipo_Trans", "I"
                SetAdoUpdate
                Ln_No = Ln_No + 1
                  
                'Despliega los datos en el DataGrid
                sSQL = "SELECT * " _
                     & "FROM Asiento_Air " _
                     & "WHERE CodRet <> '.' " _
                     & "ORDER BY CodRet "
                SelectDataGrid DGConceptoAir, AdoAsientoAir, sSQL, "Sustento"
                    
               'Realiza la Sumatoria de las Retenciones
               ac = ac + TxtValConA
               TxtTotalReten = ac
            End If
            RatonNormal
    End If
End Sub

Private Sub TxtComision_GotFocus()
    MarcarTexto TxtComision
End Sub

Private Sub TxtComision_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtComision_LostFocus()
    TextoValido TxtComision
    Debe = Convertir_Numero(TxtConsumoGrav, 2)
   'Calculo el Monto IVA
    Valor = Debe * 0.12
    TxtMontoIva = Valor
End Sub

Private Sub TxtConsumo_GotFocus()
    MarcarTexto TxtConsumo
End Sub

Private Sub TxtConsumo_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtConsumo_LostFocus()
    TextoValido TxtConsumo
End Sub

Private Sub TxtConsumoGrav_GotFocus()
    MarcarTexto TxtConsumoGrav
End Sub

Private Sub TxtConsumoGrav_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter TxtConsumoGrav
End Sub

Private Sub TxtConsumoGrav_LostFocus()
    TextoValido TxtConsumoGrav
   'Calculo el Total Consumo
    TxtTotalConsumo = Convertir_Numero(TxtConsumoGrav, 2) + Convertir_Numero(TxtConsumo, 2)
End Sub

Private Sub TxtIvaBienMonIva_GotFocus()
    MarcarTexto TxtIvaBienMonIva
End Sub

Private Sub TxtIvaBienMonIva_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtIvaBienMonIva_LostFocus()
    TextoValido TxtIvaBienMonIva
End Sub

Private Sub TxtIvaSerMonIva_GotFocus()
    MarcarTexto TxtIvaSerMonIva
End Sub

Private Sub TxtIvaSerMonIva_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtIvaSerMonIva_LostFocus()
    TextoValido TxtIvaSerMonIva
End Sub

Private Sub TxtMontoIva_GotFocus()
    MarcarTexto TxtMontoIva
End Sub

Private Sub TxtMontoIva_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtMontoIva_LostFocus()
    TextoValido TxtMontoIva, True, , 0
    TxtIvaBienMonIva = TxtMontoIva
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
    TextoValido TxtNumeroC, True, , 0
    If Val(TxtNumeroC) <= 0 Then TxtNumeroC = "0"
       TxtNumeroC = Format(Val(CCur(TxtNumeroC)), String(7, "0"))
End Sub

Private Sub TxtNumRecap_GotFocus()
    MarcarTexto TxtNumRecap
End Sub

Private Sub TxtNumRecap_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtNumRecap_LostFocus()
   TextoValido TxtNumRecap, True, , 0
   If Val(TxtNumRecap) <= 0 Then TxtNumRecap = "000000000000001"
      TxtNumRecap = Format(Val(Round(TxtNumRecap)), "000000000000000")
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
End Sub

Private Sub TxtNumUnoAutComRet_GotFocus()
    MarcarTexto TxtNumUnoAutComRet
End Sub

Private Sub TxtNumUnoAutComRet_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtNumUnoAutComRet_LostFocus()
    TextoValido TxtNumUnoAutComRet, True, , 0
    If Val(TxtNumUnoAutComRet) <= 0 Then TxtNumUnoAutComRet = "0"
        TxtNumUnoAutComRet = Format(Val(CCur(TxtNumUnoAutComRet)), String(10, "0"))
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

Public Sub Carga_Datos_Iniciales(MBFecha As MaskEdBox, EsNuevo As Boolean)
    Trans_No = 100
    Ln_No = 0
    SSTRecap.Tab = 0
            
    'Enceramos el espacio de cada usuario para emprezar con una nueva retencion
'    sSQL = "DELETE * " _
'         & "FROM Asiento_Recap " _
'         & "WHERE Item = '" & NumEmpresa & "' " _
'         & "AND CodigoU = '" & CodigoUsuario & "' "
'    ConectarAdoExecute sSQL
'
'    'Borra Asiento Air
'    sSQL = "DELETE * " _
'         & "FROM Asiento_Air  " _
'         & "WHERE Item = '" & NumEmpresa & "' " _
'         & "AND CodigoU = '" & CodigoUsuario & "' "
'    ConectarAdoExecute sSQL

    'Carga la Tabla de Clientes al DataCombo
    Leer_Clientes
    
    'Carga la Tabla de Tarjetas de credito
    Carga_Tarjetas
   
    'Carga el Tipo de Comprobante
    Carga_TipoComprobante
   
    'Carga la Retencion en la Fuente
    Carga_RetencionFuente
   
    sSQL = "SELECT CodRet, Detalle, BaseImp, Porcentaje, ValRet, EstabRetencion, PtoEmiRetencion, SecRetencion, AutRetencion, FechaEmiRet  " _
         & "FROM Asiento_Air " _
         & "WHERE CodRet <> '.' " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND CodigoU =  '" & CodigoUsuario & "' " _
         & "ORDER BY CodRet "
    SelectDataGrid DGConceptoAir, AdoAsientoAir, sSQL
    
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
       Modificacion
       CModificacion.Visible = True
    End If
End Sub

Public Sub Grabacion()
   'Selecciona el numero mayor para continuar la secuencia en el
   'campo T_No y A_No
   'Grabo en el Asiento_Recap e implicito Asiento_Air
    SetAdoAddNew "Asiento_Recap"
    SetAdoFields "IdProv", CodigoCliente
    SetAdoFields "TipoRecap", cod
    SetAdoFields "NumRecap", TxtNumRecap
    SetAdoFields "NumVaucher", TxtNumVoucher
    'SetAdoFields "Tarjeta", ojo
    SetAdoFields "FechaEmision", MBFechaEmision
    SetAdoFields "FechaPago", MBFechaPago
    SetAdoFields "ConsumoCero", TxtConsumo
    SetAdoFields "ConsumoGrav", TxtConsumoGrav
    SetAdoFields "TotalConsumo", TxtTotalConsumo
    SetAdoFields "Comision", TxtComision
    SetAdoFields "MontoIva", TxtMontoIva
    SetAdoFields "MontoIvaBienes", TxtMontoIva
    SetAdoFields "PorRetBienes", CodRetBien
    SetAdoFields "ValorRetBienes", TxtIvaBienValRet
    SetAdoFields "MontoIvaServicios", TxtIvaSerMonIva
    SetAdoFields "PorRetServicios", CodRetSer
    SetAdoFields "ValorRetServicios", TxtIvaSerValRet
   'Verifico si activaron los checks
    If ChRetB = 1 Then
       SetAdoFields "Cta_Bienes", SinEspaciosIzq(DCRetIBienes)
    Else
       SetAdoFields "Cta_Bienes", "."
    End If
    If ChRetS = 1 Then
       SetAdoFields "Cta_Servicio", SinEspaciosIzq(DCRetISer)
    Else
       SetAdoFields "Cta_Servicio", "."
    End If
    
    SetAdoFields "Porc_Bienes", DCPorcenRetenIvaBien
    SetAdoFields "MontoIvaBienes", TxtIvaBienMonIva
    SetAdoFields "PorRetBienes", CodRetBien
    SetAdoFields "ValorRetBienes", TxtIvaBienValRet
    SetAdoFields "Porc_Servicios", DCPorcenRetenIvaServ
    SetAdoFields "MontoIvaServicios", TxtIvaSerMonIva
    SetAdoFields "PorRetServicios", CodRetServ
    SetAdoFields "ValorRetServicios", TxtIvaSerValRet
    SetAdoFields "A_No", 0
    SetAdoFields "T_No", Trans_No
    SetAdoUpdate
      
    'Grabamos los datos de la transaccion en la tabla definitiva de almacenamiento
    ID_Trans = 1   'va a tener el indice de transaccion unico para que no exista duplicados en a base
    sSQL = "SELECT TOP 1 * " _
         & "FROM Trans_Recap " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "ORDER BY IDT DESC "
    SelectAdodc AdoAsientorecap, sSQL
    If AdoAsientorecap.Recordset.RecordCount > 0 Then ID_Trans = AdoAsientorecap.Recordset.Fields("IDT") + 1
    sSQL = "SELECT * " _
         & "FROM Asiento_Recap " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND T_No = " & Trans_No & " " _
         & "ORDER BY T_No DESC "
    SelectAdodc AdoAsientorecap, sSQL
    With AdoAsientorecap.Recordset
     If .RecordCount > 0 Then
         FechaTexto = .Fields("FechaLiquidacion")
         SetAdoAddNew "Trans_Importaciones"
         SetAdoFields "T", Normal
         SetAdoFields "TipoRecap", cod
         SetAdoFields "NumRecap", .Fields("NumRecap")
         SetAdoFields "NumVaucher", .Fields("NumVaucher")
         SetAdoFields "Tarjeta", .Fields("Tarjeta")
         SetAdoFields "FechaEmision", .Fields("FechaEmision")
         SetAdoFields "FechaPago", .Fields("FechaPago")
         SetAdoFields "ConsumoCero", .Fields("ConsumoCero")
         SetAdoFields "ConsumoGrav", .Fields("ConsumoGrav")
         SetAdoFields "TotalConsumo", .Fields("TotalConsumo")
         SetAdoFields "Comision", .Fields("Comision")
         SetAdoFields "MontoIva", .Fields("MontoIva")
         SetAdoFields "MontoIvaBienes", .Fields("MontoIvaBienes")
         SetAdoFields "PorRetBienes", .Fields("PorRetBienes")
         SetAdoFields "ValorRetBienes", .Fields("ValorRetBienes")
         SetAdoFields "MontoIvaServicios", .Fields("MontoIvaServicios")
         SetAdoFields "PorRetServicios", CodRetSer
         SetAdoFields "ValorRetServicios", .Fields("ValorRetServicios")
         
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
         SetAdoFields "TP", .Fields("TP")
         SetAdoFields "Numero", .Fields("Numero")
         SetAdoFields "Fecha", MBFechaLiquida
         SetAdoFields "IDT", ID_Trans
         SetAdoFields "Linea_SRI", 0
         SetAdoUpdate
     End If
    End With
   'Selecciona el numero mayor para continuar la secuencia en el
    ID_Trans = 1
    sSQL = "SELECT TOP 1 * " _
         & "FROM Trans_Air " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "ORDER BY IDT DESC "
    SelectAdodc AdoTransAir, sSQL
    If AdoTransAir.Recordset.RecordCount > 0 Then ID_Trans = AdoTransAir.Recordset.Fields("IDT") + 1
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
                SetAdoFields "CodRet", .Fields("CodRet")
                SetAdoFields "BaseImp", .Fields("BaseImp")
                SetAdoFields "Porcentaje", .Fields("Porcentaje")
                SetAdoFields "ValRet", .Fields("ValRet")
                SetAdoFields "EstabRetencion", .Fields("EstabRetencion")
                SetAdoFields "PtoEmiRetencion", .Fields("PtoEmiRetencion")
                SetAdoFields "SecRetencion", .Fields("SecRetencion")
                SetAdoFields "AutRetencion", .Fields("AutRetencion")
                SetAdoFields "Tipo_Trans", .Fields("Tipo_Trans")
                SetAdoFields "Numero", .Fields("Numero")
                SetAdoFields "Fecha", FechaTexto
                SetAdoFields "EstabFactura", "001"
                SetAdoFields "PuntoEmiFactura", "001"
                SetAdoFields "Factura_No", .Fields("Factura_No")
                SetAdoFields "IdProv", .Fields("IdProv")
                SetAdoFields "TP", .Fields("TP")
                SetAdoFields "Cta_Retencion", .Fields("Cta_Retencion")
                SetAdoFields "IDT", ID_Trans
                SetAdoFields "Linea_SRI", 0
                SetAdoUpdate
                ID_Trans = ID_Trans + 1
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
End Sub

Public Sub Deshabilita_Controles()
    'Deshabilito los controles para la modificacion
    CModificacion.Enabled = False
    SSTImportaciones.Enabled = False
    DCProveedor.Enabled = False
    CmdGrabar.Enabled = False
    CTP.Enabled = False
    TxtNumeroC.Enabled = False
End Sub

Private Sub TxtNumVoucher_GotFocus()
    MarcarTexto TxtNumVoucher
End Sub

Private Sub TxtNumVoucher_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtNumVoucher_LostFocus()
   TextoValido TxtNumVoucher, True, , 0
End Sub

Public Sub Carga_TipoComprobante()
    'Busco el codigo en la tabla Tipo Comprobante///descripcion
     sSQL = "SELECT Descripcion,* " _
          & "FROM Tipo_Comprobante " _
          & "WHERE Tipo_Comprobante_Codigo = 22 " _
          & "AND Tipo_Comprobante_Codigo = 23 " _
          & "AND Tipo_Comprobante_Codigo = 24 " _
          & "ORDER BY Tipo_Comprobante_Codigo "
     SelectDBCombo DCTipoRecap, AdoTipoComprobante, sSQL, "Descripcion"
End Sub

Public Sub Captura_TipoComprobante()
   'Captura lo que tiene el Combo de Tipo de Comprobante
    Captc = Trim(DCTipoRecap.Text)
     
   'Busca que sea igual a la Descripcion
    With AdoTipoComprobante.Recordset
     If .RecordCount > 0 Then
        .MoveFirst
        .Find ("Descripcion = '" & Captc & "' ")
         If Not .EOF Then
            cod = .Fields("Tarjeta_Credito_Codigo")
         Else
            MsgBox "Vuelva a seleccionar", vbInformation, "Aviso"
         End If
     End If
    End With
End Sub

Public Sub Captura_Tarjetas()
    Captc = Trim(DCTarjetaCred.Text)
     
   'Busca que sea igual a la Descripcion
    With AdoTarjetas.Recordset
     If .RecordCount > 0 Then
        .MoveFirst
        .Find ("Descripcion = '" & Captc & "' ")
         If Not .EOF Then
            cod = .Fields("Tarjeta_Credito_Codigo")
         Else
            MsgBox "Vuelva a seleccionar", vbInformation, "Aviso"
         End If
     End If
    End With
End Sub

Public Sub Carga_RetencionIvaBienes()
   sSQL = "SELECT * " _
        & "FROM Tabla_Por_IVA " _
        & "WHERE Bienes <> " & Val(adFalse) & " " _
        & "ORDER BY Porc "
   SelectDBCombo DCPorcenRetenIvaBien, AdoRetIvaBienes, sSQL, "Porc"
End Sub

Public Sub Carga_RetencionIvaServicios()
   sSQL = "SELECT * " _
        & "FROM Tabla_Por_IVA " _
        & "WHERE Servicios <> " & Val(adFalse) & " " _
        & "ORDER BY Porc "
   SelectDBCombo DCPorcenRetenIvaServ, AdoRetIvaServicios, sSQL, "Porc"
End Sub


