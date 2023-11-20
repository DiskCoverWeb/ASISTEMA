VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form FVentasAT 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VENTAS"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10620
   Icon            =   "FVentasAT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   10620
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrmRetencion 
      BackColor       =   &H00C0FFFF&
      Caption         =   "RETENCIONES DE IVA POR:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   984
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   9240
      Begin MSDataListLib.DataCombo DCRetISer 
         Bindings        =   "FVentasAT.frx":0696
         DataSource      =   "AdoRetIvaSerCC"
         Height          =   315
         Left            =   1365
         TabIndex        =   4
         Top             =   525
         Visible         =   0   'False
         Width           =   7680
         _ExtentX        =   13547
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DCRetIBienes 
         Bindings        =   "FVentasAT.frx":06B3
         DataSource      =   "AdoRetIvaBienesCC"
         Height          =   315
         Left            =   1365
         TabIndex        =   2
         Top             =   210
         Visible         =   0   'False
         Width           =   7680
         _ExtentX        =   13547
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.CheckBox ChRetS 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Servicios"
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
         TabIndex        =   3
         Top             =   630
         Width           =   1170
      End
      Begin VB.CheckBox ChRetB 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Bienes"
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
         TabIndex        =   1
         Top             =   315
         Width           =   1275
      End
   End
   Begin VB.CommandButton CmdGrabar 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Aceptar"
      Height          =   750
      Left            =   9660
      Picture         =   "FVentasAT.frx":06D3
      Style           =   1  'Graphical
      TabIndex        =   52
      ToolTipText     =   "Grabar"
      Top             =   105
      Width           =   855
   End
   Begin VB.CommandButton CmdCerrar 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Cancelar"
      Height          =   750
      Left            =   9660
      Picture         =   "FVentasAT.frx":09DD
      Style           =   1  'Graphical
      TabIndex        =   53
      ToolTipText     =   "Salir"
      Top             =   945
      Width           =   855
   End
   Begin TabDlg.SSTab SSTVentas 
      Height          =   4770
      Left            =   105
      TabIndex        =   8
      Top             =   1785
      Width           =   10380
      _ExtentX        =   18309
      _ExtentY        =   8414
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   420
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "COMPROBANTE DE VENTA"
      TabPicture(0)   =   "FVentasAT.frx":0E1F
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FrmIva"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "INSERTAR CONCEPTO AIR"
      TabPicture(1)   =   "FVentasAT.frx":0E3B
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Caption         =   "INGRESE LOS DATOS DE LA FACTURA, NOTA DE VENTA, ETC. ______________________ FORMULARIO 104"
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
         Height          =   2115
         Left            =   210
         TabIndex        =   6
         Top             =   315
         Width           =   9990
         Begin VB.CommandButton CmdAir 
            Caption         =   "&AIR"
            Height          =   855
            Left            =   8925
            Picture         =   "FVentasAT.frx":0E57
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Se ubica en la pestaña de Retenciones"
            Top             =   315
            Width           =   960
         End
         Begin VB.TextBox TxtNumSerietres 
            Height          =   336
            Left            =   1260
            MaxLength       =   9
            TabIndex        =   13
            Text            =   "0000001"
            ToolTipText     =   $"FVentasAT.frx":137D
            Top             =   1680
            Width           =   1065
         End
         Begin VB.TextBox TxtNumSerieUno 
            Height          =   336
            Left            =   105
            MaxLength       =   3
            TabIndex        =   11
            Text            =   "001"
            ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al código del establecimiento"
            Top             =   1680
            Width           =   540
         End
         Begin VB.TextBox TxtNumSerieDos 
            Height          =   336
            Left            =   630
            MaxLength       =   3
            TabIndex        =   12
            Text            =   "001"
            ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al punto dde emisión"
            Top             =   1680
            Width           =   645
         End
         Begin VB.TextBox TxtBaseImpV 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   5145
            TabIndex        =   16
            Text            =   "0.00"
            ToolTipText     =   "En este campo se debe ingresar el valor del comprobante cuya base imponible esta gravado con la tarifa del 0% de IVA"
            Top             =   1680
            Width           =   1485
         End
         Begin VB.TextBox TxtNumComprobante 
            Height          =   336
            Left            =   7560
            MaxLength       =   7
            TabIndex        =   10
            Top             =   525
            Width           =   1275
         End
         Begin VB.TextBox TxtBaseImpGravV 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   6720
            TabIndex        =   17
            Text            =   "0.00"
            ToolTipText     =   $"FVentasAT.frx":1420
            Top             =   1680
            Width           =   1590
         End
         Begin VB.TextBox TxtBaseImpoIceV 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   8295
            MultiLine       =   -1  'True
            TabIndex        =   18
            Text            =   "FVentasAT.frx":14C8
            ToolTipText     =   $"FVentasAT.frx":14CD
            Top             =   1680
            Width           =   1590
         End
         Begin MSDataListLib.DataCombo DCTipoComprobanteV 
            Bindings        =   "FVentasAT.frx":155F
            DataSource      =   "AdoTipoComprobante"
            Height          =   315
            Left            =   105
            TabIndex        =   9
            ToolTipText     =   $"FVentasAT.frx":1580
            Top             =   525
            Width           =   7365
            _ExtentX        =   12991
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSMask.MaskEdBox MBFechaEmiV 
            Height          =   330
            Left            =   2415
            TabIndex        =   14
            ToolTipText     =   $"FVentasAT.frx":1628
            Top             =   1680
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   582
            _Version        =   393216
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "0"
         End
         Begin MSMask.MaskEdBox MBFechaRegistroV 
            Height          =   330
            Left            =   3780
            TabIndex        =   15
            ToolTipText     =   $"FVentasAT.frx":16D4
            Top             =   1680
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   582
            _Version        =   393216
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "0"
         End
         Begin MSDataListLib.DataCombo DCTipoPago 
            Bindings        =   "FVentasAT.frx":175C
            DataSource      =   "AdoTipoPago"
            Height          =   315
            Left            =   1890
            TabIndex        =   83
            Top             =   945
            Width           =   6945
            _ExtentX        =   12250
            _ExtentY        =   556
            _Version        =   393216
            Text            =   "Forma de Pago"
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
         Begin VB.Label Label41 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " FORMA DE PAGO"
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
            TabIndex        =   84
            Top             =   945
            Width           =   1800
         End
         Begin VB.Label Label11 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " CADUCIDAD"
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
            Left            =   3780
            TabIndex        =   73
            Top             =   1365
            Width           =   1275
         End
         Begin VB.Label Label10 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " EMISION"
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
            Left            =   2415
            TabIndex        =   72
            Top             =   1365
            Width           =   1275
         End
         Begin VB.Label Label9 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " SERIE Y COMPROBAN"
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
            TabIndex        =   71
            Top             =   1365
            Width           =   2220
         End
         Begin VB.Label Label14 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " VALOR ICE"
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
            TabIndex        =   70
            Top             =   1365
            Width           =   1590
         End
         Begin VB.Label Label13 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " TARIFA 12%"
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
            Left            =   6720
            TabIndex        =   69
            Top             =   1365
            Width           =   1590
         End
         Begin VB.Label Label12 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " TARIFA 0%"
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
            Left            =   5145
            TabIndex        =   68
            Top             =   1365
            Width           =   1485
         End
         Begin VB.Label Label7 
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
            Height          =   225
            Left            =   7560
            TabIndex        =   67
            Top             =   315
            Width           =   1275
         End
         Begin VB.Label Label5 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " TIPO DE COMPROBANTE"
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
            TabIndex        =   66
            Top             =   315
            Width           =   7365
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Retención Presuntiva"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   2835
         TabIndex        =   22
         Top             =   2520
         Width           =   2535
         Begin VB.OptionButton OpcRetNo 
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
            Left            =   1470
            TabIndex        =   24
            Top             =   315
            Width           =   645
         End
         Begin VB.OptionButton OpcRetSi 
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
            Left            =   315
            TabIndex        =   23
            Top             =   315
            Width           =   540
         End
      End
      Begin VB.Frame FrmIva 
         Caption         =   "I.V.A. Presuntivo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   210
         TabIndex        =   19
         Top             =   2520
         Width           =   2535
         Begin VB.OptionButton OpcIvaNo 
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
            Left            =   1470
            TabIndex        =   21
            Top             =   315
            Width           =   645
         End
         Begin VB.OptionButton OpcIvaSi 
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
            Left            =   315
            TabIndex        =   20
            Top             =   315
            Width           =   540
         End
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
         Height          =   2070
         Left            =   5460
         TabIndex        =   30
         Top             =   2520
         Width           =   4752
         Begin VB.TextBox TxtIvaBienMonIvaV 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   336
            Left            =   1680
            MultiLine       =   -1  'True
            TabIndex        =   31
            Text            =   "FVentasAT.frx":1776
            ToolTipText     =   $"FVentasAT.frx":177B
            Top             =   735
            Width           =   1380
         End
         Begin VB.TextBox TxtIvaBienValRetV 
            Enabled         =   0   'False
            Height          =   336
            Left            =   1680
            TabIndex        =   33
            Top             =   1470
            Width           =   1380
         End
         Begin VB.TextBox TxtIvaSerMonIvaV 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   336
            Left            =   3156
            MultiLine       =   -1  'True
            TabIndex        =   34
            Text            =   "FVentasAT.frx":181A
            ToolTipText     =   $"FVentasAT.frx":181F
            Top             =   735
            Width           =   1380
         End
         Begin VB.TextBox TxtIvaSerValRetV 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   336
            Left            =   3156
            TabIndex        =   36
            Text            =   " "
            Top             =   1470
            Width           =   1380
         End
         Begin MSDataListLib.DataCombo DCPorcenRetenIvaBienV 
            Bindings        =   "FVentasAT.frx":18B5
            DataSource      =   "AdoRetIvaBienes"
            Height          =   315
            Left            =   1680
            TabIndex        =   32
            ToolTipText     =   $"FVentasAT.frx":18D3
            Top             =   1155
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DCPorcenRetenIvaServV 
            Bindings        =   "FVentasAT.frx":195F
            DataSource      =   "AdoRetIvaServicios"
            Height          =   315
            Left            =   3150
            TabIndex        =   35
            ToolTipText     =   "Corresponde al porcentaje retenido en el IVA generado en la prestación de servicios"
            Top             =   1155
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
         End
         Begin VB.Label Label24 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " VALOR RET."
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
            Left            =   315
            TabIndex        =   63
            Top             =   1470
            Width           =   1380
         End
         Begin VB.Label Label22 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " PORCENTAJE"
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
            Left            =   315
            TabIndex        =   62
            Top             =   1155
            Width           =   1380
         End
         Begin VB.Label Label21 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " MONTO"
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
            Left            =   315
            TabIndex        =   61
            Top             =   735
            Width           =   1380
         End
         Begin VB.Label Label20 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " SERVICIOS"
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
            Left            =   3150
            TabIndex        =   60
            Top             =   525
            Width           =   1380
         End
         Begin VB.Label Label19 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " BIENES"
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
            TabIndex        =   59
            Top             =   525
            Width           =   1380
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "PORCENTAJE DE LAS BASES IMPONIBLES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1230
         Left            =   210
         TabIndex        =   25
         Top             =   3360
         Width           =   5145
         Begin VB.TextBox TxtMontoIvaV 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   3465
            TabIndex        =   27
            ToolTipText     =   "Este valor se calcula automaticamente, es el resultado de aplicarle un porcentaje IVA a la Base Imponible gravada"
            Top             =   315
            Width           =   1485
         End
         Begin VB.TextBox TxtMontoIceV 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   336
            Left            =   3465
            TabIndex        =   29
            Top             =   735
            Width           =   1485
         End
         Begin MSDataListLib.DataCombo DCPorcenIvaV 
            Bindings        =   "FVentasAT.frx":1980
            DataSource      =   "AdoPorIva"
            Height          =   315
            Left            =   945
            TabIndex        =   26
            ToolTipText     =   $"FVentasAT.frx":1998
            Top             =   315
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DCPorcenIceV 
            Bindings        =   "FVentasAT.frx":1A2A
            DataSource      =   "AdoPorIce"
            Height          =   315
            Left            =   945
            TabIndex        =   28
            ToolTipText     =   $"FVentasAT.frx":1A42
            Top             =   735
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label Label18 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " VALOR ICE"
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
            Left            =   1995
            TabIndex        =   58
            Top             =   735
            Width           =   1485
         End
         Begin VB.Label Label17 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " VALOR I.V.A."
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
            Left            =   1995
            TabIndex        =   57
            Top             =   315
            Width           =   1485
         End
         Begin VB.Label Label16 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ICE"
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
            TabIndex        =   56
            Top             =   735
            Width           =   855
         End
         Begin VB.Label Label8 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " I.V.A."
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
            TabIndex        =   55
            Top             =   315
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
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
         Height          =   4020
         Left            =   -74895
         TabIndex        =   38
         Top             =   315
         Width           =   10155
         Begin VB.TextBox TxtValConAV 
            Enabled         =   0   'False
            Height          =   330
            Left            =   8715
            TabIndex        =   49
            Top             =   1470
            Width           =   1275
         End
         Begin VB.TextBox TxtPorRetConAV 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   336
            Left            =   8085
            TabIndex        =   48
            Top             =   1470
            Width           =   645
         End
         Begin VB.TextBox TxtBimpConAV 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   6720
            TabIndex        =   47
            Top             =   1470
            Width           =   1380
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
            Left            =   8820
            MultiLine       =   -1  'True
            TabIndex        =   45
            Top             =   840
            Width           =   1170
         End
         Begin VB.TextBox TxtNumTresComRetV 
            Height          =   336
            Left            =   1260
            MaxLength       =   9
            TabIndex        =   43
            Text            =   "0000001"
            ToolTipText     =   $"FVentasAT.frx":1AD3
            Top             =   840
            Width           =   1065
         End
         Begin VB.TextBox TxtNumDosComRetV 
            Height          =   336
            Left            =   630
            MaxLength       =   3
            TabIndex        =   42
            Text            =   "001"
            ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al punto dde emisión"
            Top             =   840
            Width           =   540
         End
         Begin VB.TextBox TxtNumUnoComRetV 
            Height          =   336
            Left            =   108
            MaxLength       =   3
            TabIndex        =   41
            Text            =   "001"
            ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al código del establecimiento"
            Top             =   840
            Width           =   540
         End
         Begin VB.TextBox TxtTotalReten 
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
            Left            =   8715
            TabIndex        =   51
            Text            =   "0.00"
            ToolTipText     =   "Sumatoria total de las retenciones"
            Top             =   3570
            Width           =   1275
         End
         Begin VB.TextBox TxtNumUnoAutComRetV 
            Height          =   330
            Left            =   2415
            MaxLength       =   49
            TabIndex        =   44
            ToolTipText     =   $"FVentasAT.frx":1B75
            Top             =   840
            Width           =   4950
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
            TabIndex        =   39
            Top             =   315
            Width           =   2328
         End
         Begin MSDataListLib.DataCombo DCConceptoRetV 
            Bindings        =   "FVentasAT.frx":1C01
            DataSource      =   "AdoConceptoRet"
            Height          =   315
            Left            =   105
            TabIndex        =   46
            Top             =   1470
            Width           =   6630
            _ExtentX        =   11695
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
         Begin MSDataGridLib.DataGrid DGConceptoAirV 
            Bindings        =   "FVentasAT.frx":1C1E
            Height          =   1545
            Left            =   105
            TabIndex        =   50
            Top             =   1890
            Width           =   9945
            _ExtentX        =   17542
            _ExtentY        =   2725
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
            Bindings        =   "FVentasAT.frx":1C3A
            DataSource      =   "AdoRetFuente"
            Height          =   315
            Left            =   2520
            TabIndex        =   40
            Top             =   315
            Visible         =   0   'False
            Width           =   7470
            _ExtentX        =   13176
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label Label33 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " BASE IMP."
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
            Left            =   6720
            TabIndex        =   64
            Top             =   1260
            Width           =   1380
         End
         Begin VB.Label Label1 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Total Retenciones"
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
            Left            =   6720
            TabIndex        =   82
            Top             =   3570
            Width           =   2010
         End
         Begin VB.Label Label28 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Serie"
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
            TabIndex        =   81
            Top             =   630
            Width           =   1065
         End
         Begin VB.Label Label29 
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
            Height          =   225
            Left            =   1260
            TabIndex        =   80
            Top             =   630
            Width           =   1065
         End
         Begin VB.Label Label30 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Autorización"
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
            Left            =   2415
            TabIndex        =   79
            Top             =   630
            Width           =   4950
         End
         Begin VB.Label Label31 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " SUMATORIA"
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
            TabIndex        =   78
            Top             =   840
            Width           =   1380
         End
         Begin VB.Label Label32 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " CODIGO DE RETENCION"
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
            TabIndex        =   77
            Top             =   1260
            Width           =   6630
         End
         Begin VB.Label Label34 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " PORC"
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
            TabIndex        =   76
            Top             =   1260
            Width           =   645
         End
         Begin VB.Label Label35 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " VALOR RET."
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
            Left            =   8715
            TabIndex        =   75
            Top             =   1260
            Width           =   1275
         End
         Begin VB.Label Label36 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Total Retenciones"
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
            Left            =   6720
            TabIndex        =   74
            Top             =   4515
            Width           =   2010
         End
      End
   End
   Begin MSAdodcLib.Adodc AdoTransVentas 
      Height          =   330
      Left            =   210
      Top             =   5190
      Visible         =   0   'False
      Width           =   2580
      _ExtentX        =   4551
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
   Begin MSAdodcLib.Adodc AdoRetFuente 
      Height          =   330
      Left            =   210
      Top             =   4245
      Visible         =   0   'False
      Width           =   2580
      _ExtentX        =   4551
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
   Begin MSAdodcLib.Adodc AdoAsientoVentas 
      Height          =   330
      Left            =   210
      Top             =   4560
      Visible         =   0   'False
      Width           =   2580
      _ExtentX        =   4551
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
   Begin MSAdodcLib.Adodc AdoRetIvaBienesCC 
      Height          =   330
      Left            =   210
      Top             =   5820
      Visible         =   0   'False
      Width           =   2580
      _ExtentX        =   4551
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
      Caption         =   "RetencionIvaBienes"
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
   Begin MSAdodcLib.Adodc AdoRetIvaSerCC 
      Height          =   330
      Left            =   210
      Top             =   5505
      Visible         =   0   'False
      Width           =   2580
      _ExtentX        =   4551
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
      Caption         =   "RetencionIvaServicios"
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
      Left            =   3045
      Top             =   2040
      Visible         =   0   'False
      Width           =   2475
      _ExtentX        =   4366
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
   Begin MSAdodcLib.Adodc AdoAsientoAir 
      Height          =   330
      Left            =   3045
      Top             =   2670
      Visible         =   0   'False
      Width           =   2475
      _ExtentX        =   4366
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
      Caption         =   "AsientoAirVentas"
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
      Top             =   3930
      Visible         =   0   'False
      Width           =   2580
      _ExtentX        =   4551
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
      Caption         =   "AdoConceptoAir"
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
      Left            =   210
      Top             =   3300
      Visible         =   0   'False
      Width           =   2580
      _ExtentX        =   4551
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
   Begin MSAdodcLib.Adodc AdoPorIva 
      Height          =   330
      Left            =   210
      Top             =   3615
      Visible         =   0   'False
      Width           =   2580
      _ExtentX        =   4551
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
   Begin MSAdodcLib.Adodc AdoRetIvaBienes 
      Height          =   330
      Left            =   210
      Top             =   2655
      Visible         =   0   'False
      Width           =   2580
      _ExtentX        =   4551
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
   Begin MSAdodcLib.Adodc AdoTipoComprobante 
      Height          =   330
      Left            =   210
      Top             =   2325
      Visible         =   0   'False
      Width           =   2580
      _ExtentX        =   4551
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
   Begin MSAdodcLib.Adodc AdoTipoIdentificacion 
      Height          =   330
      Left            =   210
      Top             =   1995
      Visible         =   0   'False
      Width           =   2580
      _ExtentX        =   4551
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
   Begin MSAdodcLib.Adodc AdoSustento 
      Height          =   330
      Left            =   210
      Top             =   4875
      Visible         =   0   'False
      Width           =   2580
      _ExtentX        =   4551
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
   Begin MSAdodcLib.Adodc AdoTransAir 
      Height          =   330
      Left            =   3045
      Top             =   2355
      Visible         =   0   'False
      Width           =   2475
      _ExtentX        =   4366
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
      Caption         =   "TransAirventas"
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
      Top             =   2970
      Visible         =   0   'False
      Width           =   2580
      _ExtentX        =   4551
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   3045
      Top             =   3045
      Visible         =   0   'False
      Width           =   2475
      _ExtentX        =   4366
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
      Left            =   3045
      Top             =   3360
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
   Begin MSAdodcLib.Adodc AdoTipoPago 
      Height          =   330
      Left            =   3045
      Top             =   3675
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
      Caption         =   "TipoPago"
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
   Begin VB.Label Label25 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " PROVEEDOR / BENEFICIARIO"
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
      TabIndex        =   65
      Top             =   1155
      Width           =   9255
   End
   Begin VB.Label LblNumIdentV 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   330
      Left            =   7455
      TabIndex        =   7
      Top             =   1365
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
      Left            =   7140
      TabIndex        =   5
      Top             =   1365
      Width           =   330
   End
   Begin VB.Label LblClienteV 
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
      TabIndex        =   54
      Top             =   1365
      Width           =   7050
   End
End
Attribute VB_Name = "FVentasAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MBFecha As MaskEdBox
Dim cod, X, Rb, Rf, rs, ch As Byte
Dim OP As Boolean
Dim SumAnio, Aniocad, CodPorIva, CodPorIce, CodRetBien, CodRetServ, Ctag As Integer
Dim CalmIva, CalmIce, CalIbMi, CalIsMi, ac, SUM, cal As Double
Dim Cap, Cap1, Ct, Valor, AuxCodUs, CodC, RetP, ValorP, ValorR, Bien, Serv, CargaTC, Opc, Ch1 As String
Dim Espizq, Espder, Captc, PorIva, PorIce, CodProv, CodProv1, NumCed As String

Private Sub ChRetB_Click()
    If ChRetB.value <> 0 Then
       ch = 1
       Ch1 = "B"
       DCRetIBienes.Visible = True
       TxtIvaBienMonIvaV.Enabled = True
       DCPorcenRetenIvaBienV.Enabled = True
       TxtIvaBienValRetV.Enabled = True
    Else
       TxtIvaBienMonIvaV.Enabled = False
       DCPorcenRetenIvaBienV.Enabled = False
       TxtIvaBienValRetV.Enabled = False
    End If
    If ChRetB.value <> 0 And ChRetS.value <> 0 Then
       Ch1 = "X"
    End If
End Sub

Private Sub ChRetF_Click()
If ChRetF.value <> 0 Then
       DCRetFuente.Visible = True
       TxtNumUnoComRetV.Enabled = True
       TxtNumDosComRetV.Enabled = True
       TxtNumTresComRetV.Enabled = True
       TxtNumUnoAutComRetV.Enabled = True
       DCConceptoRetV.Enabled = True
       TxtBimpConAV.Enabled = True
    Else
       DCRetFuente.Visible = False
       TxtNumUnoComRetV.Enabled = False
       TxtNumDosComRetV.Enabled = False
       TxtNumTresComRetV.Enabled = False
       TxtNumUnoAutComRetV.Enabled = False
       DCConceptoRetV.Enabled = False
       TxtBimpConAV.Enabled = False
    End If
End Sub

Private Sub ChRetS_Click()
  If ChRetS.value <> 0 Then
       ch = 1
       Ch1 = "S"
       DCRetISer.Visible = True
       TxtIvaSerMonIvaV.Enabled = True
       DCPorcenRetenIvaServV.Enabled = True
       TxtIvaSerValRetV.Enabled = True
    Else
       DCRetISer.Visible = False
       TxtIvaSerMonIvaV.Enabled = False
       DCPorcenRetenIvaServV.Enabled = False
       TxtIvaSerValRetV.Enabled = False
    End If
    If ChRetB.value <> 0 And ChRetS.value <> 0 Then
       Ch1 = "X"
    End If
End Sub

Private Sub CmdAir_Click()
  SSTVentas.Tab = 1
  TxtNumUnoComRetV.SetFocus
End Sub

Private Sub CmdCerrar_Click()
 'Borra Asiento Ventas
  sSQL = "DELETE * " _
       & "FROM Asiento_Ventas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  Ejecutar_SQL_SP sSQL
 'Borra Asiento Air
  sSQL = "DELETE * " _
       & "FROM Asiento_Air " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND Tipo_Trans = 'V' " _
       & "AND T_No = " & Trans_No & " "
  Ejecutar_SQL_SP sSQL
  Unload FVentasAT
End Sub

Private Sub CmdGrabar_Click()
    RatonReloj
   'Valido por si acaso exista algun valor con 0
    TextoValido TxtIvaBienMonIvaV, True, , 2
    TextoValido TxtBaseImpV, True, , 2
    TextoValido TxtBaseImpGravV, True, , 2
    TextoValido TxtBaseImpoIceV, True, , 2
    TextoValido TxtMontoIvaV, True, , 2
    TextoValido TxtMontoIceV, True, , 2
    TextoValido TxtIvaBienMonIvaV, True, , 2
    TextoValido TxtIvaBienValRetV, True, , 2
    TextoValido TxtIvaSerMonIvaV, True, , 2
    TextoValido TxtIvaSerValRetV, True, , 2
   'Grabacion de los Datos
    Grabacion
    sSQL = "SELECT * " _
         & "FROM Asiento " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND T_No = " & Trans_No & " "
    Select_Adodc AdoAsientos, sSQL
    OpcTM = 1
    OpcDH = 1
    NoCheque = Ninguno
   'Grabamos el Asiento de la Compra
    sSQL = "SELECT * " _
         & "FROM Asiento_Ventas " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND T_No = " & Trans_No & " "
    Select_Adodc AdoAux, sSQL
    With AdoAux.Recordset
     If .RecordCount > 0 Then
        'Porcentaje por Servicio: 0,30,100
         Cta = .Fields("Cta_Servicio")
         DetalleComp = "Crédito de Ret. del " & .Fields("Porc_Bienes") & "%, Factura No. " & .Fields("Secuencial") & ", de " & NombreCliente
         Codigo = Leer_Cta_Catalogo(Cta)
         ValorDH = .Fields("ValorRetServicios")
         If ValorDH > 0 Then InsertarAsiento AdoAsientos
        'Porcentaje por Bienes: 0,70,100
         Cta = .Fields("Cta_Bienes")
         DetalleComp = "Crédito de Ret. del " & .Fields("Porc_Servicios") & "%, Factura No. " & .Fields("Secuencial") & ", de " & NombreCliente
         Codigo = Leer_Cta_Catalogo(Cta)
         ValorDH = .Fields("ValorRetBienes")
         If ValorDH > 0 Then InsertarAsiento AdoAsientos
     End If
    End With
   'Grabamos el Asiento de las Retenciones
    sSQL = "SELECT * " _
         & "FROM Asiento_Air " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND T_No = " & Trans_No & " " _
         & "AND Tipo_Trans = 'V' " _
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
    Unload FVentasAT
End Sub

Private Sub DCConceptoRetV_LostFocus()
    OP = False
    If IsNumeric(DCConceptoRetV.Text) Then
       MsgBox "No ingrese números. Vuelva a seleccionar.", vbInformation, "Aviso"
       DCConceptoRetV.SetFocus
    Else
       With AdoConceptoRet.Recordset
        If .RecordCount > 0 Then
           .MoveFirst
           .Find ("Codigo = '" & SinEspaciosIzq(DCConceptoRetV) & "' ")
            If Not .EOF Then
               TxtPorRetConAV = .Fields("Porcentaje")
               If .Fields("Ingresar_Porcentaje") = "S" Then OP = True
            Else
               MsgBox "No encontro este código vuelva a buscar"
            End If
        End If
       End With
       TxtBimpConAV = TxtSumatoria
    End If
    If OP Then
       TxtPorRetConAV.Enabled = True
       TxtPorRetConAV.SetFocus
    Else
       TxtPorRetConAV.Enabled = False
    End If
End Sub

Private Sub DCPorcenIceV_LostFocus()
  'Busca y captura el codigo de Porcentaje IVA
    PorIce = SinEspaciosDer(DCPorcenIceV.Text)
    With AdoPorIce.Recordset
        If .RecordCount > 0 Then
           .MoveFirst
           .Find ("Porc = '" & PorIce & "' ")
           If Not .EOF Then
              CodPorIce = .Fields("Codigo")
           Else
              MsgBox "Vuelva a seleccionar", vbInformation, "Aviso"
           End If
        End If
    End With
    Total_IVA = CTNumero(TxtBaseImpoIceV, 2)
    TxtMontoIceV = 0
   'Calcula el Porcentaje de Ice
    CalIbMi = (Total_IVA * DCPorcenIceV) / 100
    TxtMontoIceV = CalIbMi
    
   'Coloca el valor de Monto IVA dependiendo si se activo Bienes o Servicios
    If ChRetB + ChRetS = 0 Then
       TxtIvaBienMonIvaV = TxtMontoIvaV
    End If
    If ChRetB.value <> 0 Then
       TxtIvaBienMonIvaV = TxtMontoIvaV
       TxtIvaSerMonIvaV = 0
    Else
       If ChRetS.value <> 0 Then
          TxtIvaSerMonIvaV = TxtMontoIvaV
          TxtIvaBienMonIvaV = 0
       End If
    End If
End Sub

Private Sub DCPorcenIvaV_LostFocus()
  'Busca y captura el codigo de Porcentaje IVA
    PorIva = SinEspaciosDer(DCPorcenIvaV.Text)
    With AdoPorIva.Recordset
        If .RecordCount > 0 Then
           .MoveFirst
           .Find ("Porc = '" & PorIva & "' ")
           CodPorIva = .Fields("Codigo")
        End If
    End With
    Total_IVA = CTNumero(TxtBaseImpGravV, 2)
   'Calcula el Porcentaje de Iva
    CalmIva = (Total_IVA * DCPorcenIvaV) / 100
    TxtMontoIvaV = CalmIva
End Sub

Private Sub DCPorcenRetenIvaBienV_LostFocus()
    CodRetBien = 0
    If Not IsNumeric(DCPorcenRetenIvaBienV) Then
       MsgBox "No ingrese caracteres. Vuelva a seleccionar.", vbInformation, "Aviso"
       DCPorcenRetenIvaBienV = ""
       Carga_RetencionIvaBienes_Servicios
       DCPorcenRetenIvaBienV.SetFocus
    Else
       'Busca y captura el codigo de Porcentaje retencion Iva Bienes
       With AdoRetIvaBienes.Recordset
            If .RecordCount > 0 Then
               .MoveFirst
               .Find ("Porc = " & SinEspaciosDer(DCPorcenRetenIvaBienV) & " ")
               CodRetBien = .Fields("Codigo")
            End If
        End With
        Total_IVA = CTNumero(TxtIvaBienMonIvaV, 2)
        TxtIvaBienValRetV = 0
       'Calcula la retencion Iva Bienes
        CalIbMi = (Total_IVA * CInt(DCPorcenRetenIvaBienV)) / 100
        TxtIvaBienValRetV = CalIbMi
    End If
    TxtIvaSerMonIvaV = Format(CTNumero(TxtMontoIvaV, 2) - CTNumero(TxtIvaBienMonIvaV, 2), "#,##0.00")
End Sub

Private Sub DCPorcenRetenIvaServV_LostFocus()
   CodRetServ = 0
   'Activo el casillero para que ingrese el valor si el porcentaje es 70/100
    If DCPorcenRetenIvaServV = "70/100" Then
       Ct = "Si"
       TxtIvaSerValRetV.Text = ""
       TxtIvaSerValRetV.Enabled = True
    Else
      If Not IsNumeric(DCPorcenRetenIvaServV) Then
         MsgBox "No ingrese caracteres. Vuelva a seleccionar.", vbInformation, "Aviso"
         DCPorcenRetenIvaServV = ""
         Carga_RetencionIvaBienes_Servicios
         DCPorcenRetenIvaServV.SetFocus
      End If
    End If
    
    'Busca captura el codigo de Porcentaje retencion Iva Servicios
    With AdoRetIvaServicios.Recordset
     If .RecordCount > 0 Then
        .MoveFirst
        .Find ("Porc = " & SinEspaciosDer(DCPorcenRetenIvaServV) & " ")
        CodRetServ = .Fields("Codigo")
     Else
         MsgBox "Código erróneo", vbInformation, "Aviso"
     End If
    End With
    Ct = "No"
    Total_IVA = 0
    Total_IVA = CTNumero(TxtIvaSerMonIvaV, 2)
    TxtIvaSerValRetV = 0
    If DCPorcenRetenIvaServV = "70/100" Then
    Else
       CalIsMi = (Total_IVA) * CCur(DCPorcenRetenIvaServV) / 100
       TxtIvaSerValRetV = CalIsMi
       TxtIvaSerValRetV.Enabled = False
    End If
End Sub

Private Sub DCTipoComprobanteV_LostFocus()
    If IsNumeric(DCTipoComprobanteV.Text) Then
       MsgBox "No ingrese números. Vuelva a seleccionar.", vbInformation, "Aviso"
       DCTipoComprobanteV.Text = ""
       DCTipoComprobanteV.SetFocus
       Captura_TipoComprobanteV
    Else
       If DCTipoComprobanteV <> "" Then Captura_TipoComprobanteV
    End If
End Sub

Private Sub DGConceptoAirV_KeyDown(KeyCode As Integer, Shift As Integer)
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
                   & "AND IdProv = '" & CodigoCliente & "' " _
                   & "AND T_No = " & Trans_No & " " _
                   & "AND Tipo_Trans = 'V' " _
                   & "AND A_No = " & J & " " _
                   & "AND CodRet = '" & Codigo & "' "
              Ejecutar_SQL_SP sSQL
           End If
           AdoAsientoAir.Refresh
      End With
      Calculo_Sumatoria
   End If
 End If
End Sub

Private Sub MBFechaEmiV_GotFocus()
    MarcarTexto MBFechaEmiV
End Sub

Private Sub MBFechaEmiV_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub MBFechaEmiV_LostFocus()
    FechaValida MBFechaEmiV
    Validar_Porc_IVA MBFechaEmiV
    'Controla que la Fecha de Emisiòn este entre 31/01/2000 en adelante
    If CFechaLong(MBFechaEmiV) < CFechaLong("31/01/2000") Then
       MsgBox "La Fecha de Emisión debe ser mayor que 31/01/2000", vbInformation, "Aviso"
       MBFechaEmiV = "31/01/2000"
       MBFechaEmiV.SetFocus
    End If
    MBFechaRegistroV = MBFechaEmiV
    Carga_ConceptosRetencion MBFechaRegistroV
End Sub

Private Sub MBFechaRegistroV_GotFocus()
    MarcarTexto MBFechaRegistroV
End Sub

Private Sub MBFechaRegistroV_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub MBFechaRegistroV_LostFocus()
    FechaValida MBFechaRegistroV
   'Controla que la Fecha de Emisiòn este entre 31/01/2000 en adelante
    If CFechaLong(MBFechaRegistroV) < CFechaLong("31/01/2000") Then
       MsgBox "La Fecha de Registro debe ser mayor que 31/01/2000", vbInformation, "Aviso"
       MBFechaRegistroV = "31/01/2000"
       MBFechaRegistroV.SetFocus
    End If
   'Carga el porcentaje de IVA
    Carga_ConceptosRetencion MBFechaRegistroV
    Carga_RetencionIvaBienes_Servicios
End Sub

Private Sub OpcIvaNo_LostFocus()
    If OpcIvaNo.value = True Then ValorP = "N"
End Sub

Private Sub OpcIvaSi_LostFocus()
 If OpcIvaSi.value = True Then ValorP = "S"
End Sub

Private Sub OpcRetNo_Click()
    If OpcRetNo.value = True Then ValorR = "S"
End Sub

Private Sub OpcRetSi_LostFocus()
    If OpcRetSi.value = True Then ValorR = "S"
End Sub

Private Sub SSTVentas_Click(PreviousTab As Integer)
'    If PreviousTab = 1 Then DCClienteV.SetFocus
End Sub

Private Sub TxtBaseImpGravV_GotFocus()
    MarcarTexto TxtBaseImpGravV
End Sub

Private Sub TxtBaseImpGravV_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtBaseImpGravV_LostFocus()
    TextoValido TxtBaseImpGravV, True, , 0
End Sub

Private Sub TxtBaseImpoIceV_GotFocus()
    MarcarTexto TxtBaseImpoIceV
End Sub

Private Sub TxtBaseImpoIceV_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtBaseImpoIceV_LostFocus()
    TextoValido TxtBaseImpoIceV, True, , 0
End Sub

Private Sub TxtBaseImpV_GotFocus()
    MarcarTexto TxtBaseImpV
End Sub

Private Sub TxtBaseImpV_LostFocus()
    TextoValido TxtBaseImpV, True, , 0
    FechaValida MBFechaRegistroV
End Sub

Private Sub TxtBaseImpV_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub Form_Activate()
   MBFechaEmiV = FechaComp
   Carga_Datos_Iniciales MBFechaEmiV, Nuevo
   LblTD.Caption = TipoBenef                  ' Tipo de Cliente: C,R,P,O
   LblNumIdentV = CICliente                    ' CI o RUC del Cliente
   LblClienteV.Caption = " " & NombreCliente ' Nombre del Cliente
   MBFechaEmiV = FechaComp
   MBFechaRegistroV = FechaComp
   TxtNumSerietres = "0000001"
   sSQL = "SELECT (Codigo & ' ' & Descripcion) As CTipoPago " _
        & "FROM Tabla_Referenciales_SRI " _
        & "WHERE Tipo_Referencia = 'FORMA DE PAGO' " _
        & "AND Codigo IN ('01','16','17','18','19','20','21') " _
        & "ORDER BY Codigo "
   SelectDB_Combo DCTipoPago, AdoTipoPago, sSQL, "CTipoPago"
   
  'CodigoCliente
   Carga_TipoComprobantes (TipoBenef)
   DCTipoComprobanteV = "Documentos Autorizados en Ventas excepto ND y NC"
   TxtNumComprobante = "000001"
  'Aqui despliego el ultimo numero de la Transaccion
   sSQL = "SELECT TOP 1 * " _
        & "FROM Trans_Ventas " _
        & "WHERE IdProv = '" & CodigoCliente & "' " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "ORDER BY Secuencial DESC,NumeroComprobantes DESC "
   Select_Adodc AdoAux, sSQL
   With AdoAux.Recordset
    If .RecordCount > 0 Then
        If .Fields("Secuencial") = 0 Then
            TxtNumComprobante = .Fields("NumeroComprobantes")
        Else
            TxtNumSerietres = .Fields("Secuencial")
        End If
    End If
   End With
End Sub

Private Sub Form_Load()
    CentrarForm FVentasAT
    ConectarAdodc AdoSustento
    ConectarAdodc AdoTipoIdentificacion
    ConectarAdodc AdoTipoComprobante
    ConectarAdodc AdoRetIvaBienes
    ConectarAdodc AdoRetIvaServicios
    ConectarAdodc AdoPorIce
    ConectarAdodc AdoPorIva
    ConectarAdodc AdoConceptoRet
    ConectarAdodc AdoAsientoAir
    ConectarAdodc AdoAsientoVentas
    ConectarAdodc AdoTransAir
    ConectarAdodc AdoTransVentas
    ConectarAdodc AdoClientes
    ConectarAdodc AdoRetFuente
    ConectarAdodc AdoRetIvaSerCC
    ConectarAdodc AdoRetIvaBienesCC
    ConectarAdodc AdoAux
    ConectarAdodc AdoAsientos
    ConectarAdodc AdoTipoPago
End Sub

Private Sub TxtBimpConAV_GotFocus()
   MarcarTexto TxtBimpConAV
End Sub

Private Sub TxtBimpConAV_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TxtBimpConAV_LostFocus()
    TextoValido TxtBimpConAV, True, , 0
    TextoValido TxtSumatoria, True, , 0
   'Valida que la base imponible no sea mayor que la BIG y la BIcero
    If CTNumero(TxtBimpConAV, 2) > CTNumero(TxtSumatoria, 2) Then
       MsgBox "La Base Imponible debe ser menor o igual a la " & vbCrLf _
            & "Base Imponible Gravada + la Base Imponible 0%", vbInformation, "Aviso"
       TxtBimpConAV.Text = 0
       TxtBimpConAV.SetFocus
    Else
       If (TxtBimpConAV = "") Then
          MsgBox "Ingrese la Base Imponible que corresponda", vbInformation, "Aviso"
          TxtBimpConAV.SetFocus
       Else
          If OP = False Then
             TxtValConAV = CTNumero(TxtBimpConAV, 2) * (CTNumero(TxtPorRetConAV, 2) / 100)
             Insertar_DataGridV
          End If
       End If
    End If
    RatonNormal
End Sub

Sub Insertar_DataGridV()
  'Selecciona el numero mayor para continuar la secuencia en el
  'campo T_No y A_No
   If Val(TxtBimpConAV) > 0 Then
      RatonReloj
      Espizq = SinEspaciosIzq(DCConceptoRetV)
      Espder = TrimStrg(MidStrg(DCConceptoRetV, Len(Espizq) + 3, Len(DCConceptoRetV)))
      If Val(TxtNumComprobante) > 1 Then
         Factura_No = Val(TxtNumComprobante)
      Else
         Factura_No = Val(TxtNumSerietres)
      End If
      SetAdoAddNew "Asiento_Air"
      SetAdoFields "CodRet", Espizq
      SetAdoFields "Detalle", Espder
      SetAdoFields "BaseImp", CTNumero(TxtBimpConAV, 2)
      SetAdoFields "Porcentaje", Val(TxtPorRetConAV) / 100
      SetAdoFields "ValRet", CTNumero(TxtValConAV, 2)
      SetAdoFields "EstabRetencion", TxtNumUnoComRetV
      SetAdoFields "PtoEmiRetencion", TxtNumDosComRetV
      SetAdoFields "SecRetencion", TxtNumTresComRetV
      SetAdoFields "AutRetencion", TxtNumUnoAutComRetV
      SetAdoFields "FechaEmiRet", MBFechaRegistroV
      SetAdoFields "EstabFactura", TxtNumSerieUno
      SetAdoFields "PuntoEmiFactura", TxtNumSerieDos
      SetAdoFields "Factura_No", Factura_No
      SetAdoFields "IdProv", CodigoCliente
      SetAdoFields "Cta_Retencion", SinEspaciosIzq(DCRetFuente)
      SetAdoFields "A_No", Maximo_De("Asiento_Air", "A_No")
      SetAdoFields "T_No", Trans_No
      SetAdoFields "Tipo_Trans", "V"
      SetAdoUpdate
     'Despliega los datos en el DataGrid
      sSQL = "SELECT * " _
           & "FROM Asiento_Air " _
           & "WHERE CodRet <> '.' " _
           & "AND Item = '" & NumEmpresa & "' " _
           & "AND CodigoU = '" & CodigoUsuario & "' " _
           & "AND T_No = " & Trans_No & " " _
           & "AND Tipo_Trans = 'V' " _
           & "ORDER BY CodRet "
      Select_Adodc_Grid DGConceptoAirV, AdoAsientoAir, sSQL
     'Se situa en el combo de retención AIR
      If ChRetF.Visible Then DCRetFuente.SetFocus Else TxtNumUnoComRetV.SetFocus
     'Realiza la Sumatoria de las Retenciones
      ac = ac + TxtValConAV
      TxtTotalReten = ac
   End If
   RatonNormal
End Sub

Private Sub TxtIvaBienMonIvaV_GotFocus()
    MarcarTexto TxtIvaBienMonIvaV
End Sub

Private Sub TxtIvaBienMonIvaV_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtIvaBienMonIvaV_LostFocus()
    TextoValido TxtIvaBienMonIvaV, True, , 2
End Sub

Private Sub TxtIvaBienValRetV_GotFocus()
    MarcarTexto TxtIvaBienValRetV
End Sub

Private Sub TxtIvaBienValRetV_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtIvaBienValRetV_LostFocus()
    TextoValido TxtIvaBienValRetV, True, , 0
End Sub

Private Sub TxtIvaSerMonIvaV_GotFocus()
    TextoValido TxtIvaSerMonIvaV, True, , 2
    MarcarTexto TxtIvaSerMonIvaV
End Sub

Private Sub TxtIvaSerMonIvaV_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtIvaSerMonIvaV_LostFocus()
Dim Total_IVA_C As Currency
   TextoValido TxtIvaSerMonIvaV, True, , 0
  'Verifica el Monto Iva Servicios
   Total_IVA_C = CDbl(TxtIvaBienMonIvaV) + CDbl(TxtIvaSerMonIvaV)
   If Total_IVA_C > CDbl(TxtMontoIvaV) Then
      MsgBox "Monto IVA Servicios no puede ser > que Monto IVA", vbInformation, "Aviso de Ventas"
      TxtIvaSerMonIvaV.Text = ""
      TxtIvaSerMonIvaV.SetFocus
   End If
End Sub

Private Sub TxtMontoIvaV_GotFocus()
    MarcarTexto TxtMontoIvaV
End Sub

Private Sub TxtMontoIvaV_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtMontoIvaV_LostFocus()
    TextoValido TxtMontoIvaV, True, , 2
End Sub

Private Sub TxtNumComprobante_GotFocus()
    MarcarTexto TxtNumComprobante
End Sub

Private Sub TxtNumComprobante_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtNumComprobante_LostFocus()
    TextoValido TxtNumComprobante, True, , 0
    If Val(TxtNumComprobante) <= 0 Then TxtNumComprobante = "0000001"
       TxtNumComprobante = Format(Val(CCur(TxtNumComprobante)), "0000000")
    'Verifico si es uno o más comprobantes
    If CLng(TxtNumComprobante) <> 1 And TipoBenef <> "C" Then
       MBFechaEmiV.SetFocus
       TxtNumSerietres = "0000001"
       TxtNumSerietres.Enabled = False
    Else
       TxtNumSerietres.Enabled = True
       TxtNumSerieUno.SetFocus
    End If
End Sub

Private Sub TxtNumDosComRetV_GotFocus()
    MarcarTexto TxtNumDosComRetV
End Sub

Private Sub TxtNumDosComRetV_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtNumDosComRetV_LostFocus()
    TextoValido TxtNumDosComRetV, True, , 0
    If Val(TxtNumDosComRetV) <= 0 Then TxtNumDosComRetV = "001"
    TxtNumDosComRetV = Format(Val(TxtNumDosComRetV), "000")
End Sub

Private Sub TxtNumSerieDos_GotFocus()
    MarcarTexto TxtNumSerieDos
End Sub

Private Sub TxtNumSerieDos_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtNumSerieDos_LostFocus()
    TextoValido TxtNumSerieDos, True, , 0
    If Val(TxtNumSerieDos) <= 0 Then TxtNumSerieDos = "001"
    TxtNumSerieDos = Format(Val(TxtNumSerieDos), "000")
End Sub

Private Sub TxtNumSerietres_GotFocus()
    MarcarTexto TxtNumSerietres
End Sub

Private Sub TxtNumSerietres_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtNumSerietres_LostFocus()
    If Val(TxtNumSerietres) <= 0 Then TxtNumSerietres = "000000001"
    TxtNumSerietres = Format(Val(Round(TxtNumSerietres)), "000000000")
End Sub

Private Sub TxtNumSerieUno_GotFocus()
    MarcarTexto TxtNumSerieUno
End Sub

Private Sub TxtNumSerieUno_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtNumSerieUno_LostFocus()
   TextoValido TxtNumSerieUno, True, , 0
   If Val(TxtNumSerieUno) <= 0 Then TxtNumSerieUno = "001"
   TxtNumSerieUno = Format(Val(TxtNumSerieUno), "000")
End Sub

Private Sub TxtNumTresComRetV_GotFocus()
    MarcarTexto TxtNumTresComRetV
End Sub

Private Sub TxtNumTresComRetV_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtNumTresComRetV_LostFocus()
    If Val(TxtNumTresComRetV) <= 0 Then TxtNumTresComRetV = "00000001"
    TxtNumTresComRetV = Format(Val(Round(TxtNumTresComRetV)), "000000000")
    
    'Calcula la sumatoria de Monto Iva Bienes, Monto Iva Servicios y Base Imponible
    'TxtSumatoria = CDbl(TxtIvaBienMonIvaV) + CDbl(TxtIvaSerMonIvaV) + CDbl(TxtBaseImpV)
    TxtSumatoria = Val(CCur(TxtBaseImpV)) + Val(CCur(TxtBaseImpGravV))
End Sub

Private Sub TxtNumUnoAutComRetV_GotFocus()
    MarcarTexto TxtNumUnoAutComRetV
End Sub

Private Sub TxtNumUnoAutComRetV_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtNumUnoAutComRetV_LostFocus()
    If Val(TxtNumUnoAutComRetV) <= 0 Then TxtNumUnoAutComRetV = "0"
    Select Case Len(TxtNumUnoAutComRetV)
      Case 3: TxtNumUnoAutComRetV = Format(Val(Round(TxtNumUnoAutComRetV)), String(3, "0"))
      Case 10: TxtNumUnoAutComRetV = Format(Val(Round(TxtNumUnoAutComRetV)), String(10, "0"))
      Case Else: TxtNumUnoAutComRetV = Format(Val(Round(TxtNumUnoAutComRetV)), String(37, "0"))
    End Select
End Sub

Private Sub TxtNumUnoComRetV_GotFocus()
    MarcarTexto TxtNumUnoComRetV
End Sub

Private Sub TxtNumUnoComRetV_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtNumUnoComRetV_LostFocus()
   TextoValido TxtNumUnoComRetV, True, , 0
   If Val(TxtNumUnoComRetV) <= 0 Then TxtNumUnoComRetV = "001"
   TxtNumUnoComRetV = Format(Val(TxtNumUnoComRetV), "000")
End Sub

Public Sub Captura_TipoComprobanteV()
    'Captura lo que tiene el Combo de Tipo de Comprobante
    Captc = SinEspaciosIzq(DCTipoComprobanteV.Text)
    Cap1 = TrimStrg(DCTipoComprobanteV.Text)
    cod = Ninguno
    'Busca que sea igual a la Descripcion
    With AdoTipoComprobante.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
         .Find ("Descripcion = '" & Cap1 & "' ")
         If Not .EOF Then
            cod = .Fields("Tipo_Comprobante_Codigo")
         Else
            MsgBox "Vuelva a seleccionar", vbInformation, "Aviso"
         End If
      End If
    End With
End Sub

Public Sub Carga_RetencionIvaBienes_Servicios()
   sSQL = "SELECT * " _
        & "FROM Tabla_Por_IVA " _
        & "WHERE Bienes <> " & Val(adFalse) & " " _
        & "ORDER BY Porc "
   SelectDB_Combo DCPorcenRetenIvaBienV, AdoRetIvaBienes, sSQL, "Porc"
   
   sSQL = "SELECT * " _
        & "FROM Tabla_Por_IVA " _
        & "WHERE Servicios <> " & Val(adFalse) & " " _
        & "ORDER BY Porc "
   SelectDB_Combo DCPorcenRetenIvaServV, AdoRetIvaServicios, sSQL, "Porc"
End Sub

Public Sub Carga_TipoComprobantes(CargaTC As String)
    'Carga en el combo los tipos de comprobantes de acuerdo a la Identificacion
     sSQL = "SELECT CTT.Identificacion,CTT.Tipo_Trans,TC.* " _
          & "FROM Tabla_Tributaria As CTT, Tipo_Comprobante As TC " _
          & "WHERE CTT.Identificacion = '" & CargaTC & "' " _
          & "AND TC.TC = 'TDC' " _
          & "AND CTT.Tipo_Trans = 'V' " _
          & "AND CTT.Tipo_Comprobante_Codigo = TC.Tipo_Comprobante_Codigo " _
          & "ORDER BY TC.Tipo_Comprobante_Codigo "
     SelectDB_Combo DCTipoComprobanteV, AdoTipoComprobante, sSQL, "Descripcion"
     DCTipoComprobanteV = "Documentos Autorizados en Ventas excepto ND y NC"
End Sub

Public Sub Limpiar_Controles()
    ac = 0
    DCRetIBienes.Visible = False
    DCRetISer.Visible = False
    ChRetB.value = False
    ChRetS.value = False
    ChRetF.value = False
    LblNumIdentV.Caption = ""
    LblTD.Caption = ""
    OpcIvaNo.value = True
    OpcRetNo.value = True
    DCTipoComprobanteV.Text = ""
    TxtNumComprobante.Text = ""
    FechaValida MBFechaEmiV
    FechaValida MBFechaRegistroV
    TxtBaseImpV.Text = ""
    TxtBaseImpGravV.Text = ""
    TxtBaseImpoIceV.Text = ""
    DCPorcenIvaV.Text = ""
    TxtMontoIvaV.Text = ""
    DCPorcenIceV.Text = ""
    TxtMontoIceV.Text = ""
    TxtIvaBienMonIvaV.Text = ""
    DCPorcenRetenIvaBienV.Text = ""
    TxtIvaBienValRetV.Text = ""
    TxtIvaSerMonIvaV.Text = ""
    DCPorcenRetenIvaServV.Text = ""
    TxtIvaSerValRetV.Text = ""
    TxtNumUnoComRetV.Text = ""
    TxtNumDosComRetV.Text = ""
    TxtNumTresComRetV.Text = ""
    TxtNumUnoAutComRetV.Text = ""
    TxtSumatoria.Text = ""
    DCConceptoRetV.Text = ""
    TxtBimpConAV.Text = ""
    TxtPorRetConAV.Text = ""
    TxtValConAV.Text = ""
    TxtTotalReten.Text = ""
    'Limpia la grilla
    ' Borra Asiento Air
    sSQL = "DELETE * " _
         & "FROM Asiento_Air " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND Tipo_Trans = 'V' " _
         & "AND T_No = " & Trans_No & " "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "SELECT * " _
         & "FROM Asiento_Air " _
         & "WHERE codRet <> '.' " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND Tipo_Trans = 'V' " _
         & "AND T_No = " & Trans_No & " " _
         & "ORDER BY codRet "
    Select_Adodc_Grid DGConceptoAirV, AdoAsientoAir, sSQL
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

Public Sub Carga_Datos_Iniciales(MBFecha As MaskEdBox, EsNuevo As Boolean)
    Label16.Caption = "No.Autoriz.:" + Autorizacion
    'Encero todo
    ac = 0
    Limpiar_Controles
    DCPorcenIceV = 0
    DCPorcenRetenIvaBienV = 0
    DCPorcenRetenIvaServV = 0
    
    CodPorIva = 0
    CodPorIce = "0"
    CodRetBien = 0
    CodRetServ = 0

   'Enceramos el espacio de cada usuario para emprezar con una nueva retencion
    Listar_Air
   'Carga la Tabla de Conceptos Retencion al DataCombo
    Carga_ConceptosRetencion MBFechaRegistroV
   'Verifico si existen registros caso contrario despliego mensaje
   'Carga los Conceptos de retención en la Fuente al DataCombo
    sSQL = "SELECT (Codigo & ' - ' & Cuenta) As Cuentas  " _
         & "FROM Catalogo_Cuentas " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TC = 'CF' " _
         & "AND DG = 'D' " _
         & "ORDER BY Codigo "
    SelectDB_Combo DCRetFuente, AdoRetFuente, sSQL, "Cuentas"
    If AdoRetFuente.Recordset.RecordCount > 0 Then Rf = 1 Else Rf = 0
  
   'Carga los Conceptos de retención IVA Servicios al DataCombo
    sSQL = "SELECT (Codigo & ' - ' & Cuenta) As Cuentas  " _
         & "FROM Catalogo_Cuentas " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TC = 'CI' " _
         & "AND DG = 'D' " _
         & "ORDER BY Codigo "
    SelectDB_Combo DCRetISer, AdoRetIvaSerCC, sSQL, "Cuentas"
    If AdoRetIvaSerCC.Recordset.RecordCount > 0 Then rs = 1 Else rs = 0
    
    'Carga los Conceptos de retención IVA Bienes al DataCombo
    sSQL = "SELECT (Codigo & ' - ' & Cuenta) As Cuentas  " _
         & "FROM Catalogo_Cuentas " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TC = 'CI' " _
         & "AND DG = 'D' " _
         & "ORDER BY Codigo "
    SelectDB_Combo DCRetIBienes, AdoRetIvaBienesCC, sSQL, "Cuentas"
    If AdoRetIvaBienesCC.Recordset.RecordCount > 0 Then Rb = 1 Else Rb = 0
    
   'Si es Nuevo ingresa por aqui
    ChRetF.Visible = True
    DCRetFuente.Visible = True
    FrmRetencion.Visible = True
    If EsNuevo Then
       'Si todas las variables tienen cero despliego mensaje y no cargo nada
       'No hay cuentas
       If (Rf And rs And Rb) = 0 Then
           ChRetF.Visible = False
           DCRetFuente.Visible = False
           FrmRetencion.Visible = False
           Activar_BS
           'CTP.SetFocus
       Else
           ChRetB.SetFocus
       End If
    End If
End Sub

Public Sub Grabacion()
Dim Tipo_Pago As String
   'Selecciona el numero mayor para continuar la secuencia en el campo T_No y A_No
   'Grabo en el Asiento_Ventas e implicito Asiento_Air
    Tipo_Pago = SinEspaciosIzq(DCTipoPago)
    If Len(Tipo_Pago) <= 1 Then Tipo_Pago = "01"
    sSQL = "DELETE * " _
         & "FROM Asiento_Ventas " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND T_No = " & Trans_No & " "
    Ejecutar_SQL_SP sSQL
    SetAdoAddNew "Asiento_Ventas"
    SetAdoFields "IdProv", CodigoCliente
    SetAdoFields "TipoComprobante", cod
    SetAdoFields "FechaRegistro", MBFechaRegistroV
    SetAdoFields "FechaEmision", MBFechaEmiV
    SetAdoFields "BaseImponible", CTNumero(TxtBaseImpV, 2)
    SetAdoFields "IvaPresuntivo", ValorP
    SetAdoFields "Establecimiento", TxtNumSerieUno
    SetAdoFields "PuntoEmision", TxtNumSerieDos
    If TxtNumComprobante > 1 Then
       SetAdoFields "NumeroComprobantes", CTNumero(TxtNumComprobante)
       SetAdoFields "Secuencial", 0
    Else
       SetAdoFields "Secuencial", CTNumero(TxtNumSerietres)
       SetAdoFields "NumeroComprobantes", 1
    End If
    SetAdoFields "BaseImpGrav", CTNumero(TxtBaseImpGravV, 2)
    SetAdoFields "PorcentajeIva", CodPorIva
    SetAdoFields "MontoIva", CTNumero(TxtMontoIvaV, 2)
    SetAdoFields "BaseImpIce", CTNumero(TxtBaseImpoIceV, 2)
    SetAdoFields "PorcentajeIce", CodPorIce
    SetAdoFields "MontoIce", CTNumero(TxtMontoIceV, 2)
    SetAdoFields "MontoIvaBienes", CTNumero(TxtIvaBienMonIvaV, 2)
    SetAdoFields "PorRetBienes", CodRetBien
    SetAdoFields "ValorRetBienes", CTNumero(TxtIvaBienValRetV, 2)
    SetAdoFields "MontoIvaServicios", CTNumero(TxtIvaSerMonIvaV, 2)
    SetAdoFields "PorRetServicios", CodRetServ
    SetAdoFields "ValorRetServicios", CTNumero(TxtIvaSerValRetV, 2)
    SetAdoFields "RetPresuntiva", ValorR
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
    SetAdoFields "Porc_Bienes", DCPorcenRetenIvaBienV
    SetAdoFields "MontoIvaBienes", CTNumero(TxtIvaBienMonIvaV, 2)
    SetAdoFields "PorRetBienes", CodRetBien
    SetAdoFields "ValorRetBienes", CTNumero(TxtIvaBienValRetV, 2)
    SetAdoFields "Porc_Servicios", DCPorcenRetenIvaServV
    SetAdoFields "MontoIvaServicios", CTNumero(TxtIvaSerMonIvaV, 2)
    SetAdoFields "PorRetServicios", CodRetServ
    SetAdoFields "ValorRetServicios", CTNumero(TxtIvaSerValRetV, 2)
    SetAdoFields "Tipo_Pago", Tipo_Pago
    SetAdoFields "A_No", 1
    SetAdoFields "T_No", Trans_No
    SetAdoFields "CodigoU", CodigoUsuario
    SetAdoUpdate
End Sub

Public Sub Habilita_Controles()
    'Habilito los controles para la modificacion
    SSTVentas.Enabled = True
    CmdGrabar.Enabled = True
    FrmRetencion.Enabled = True
End Sub

Public Sub Deshabilita_Controles()
    'Deshabilito los controles para la modificacion
    SSTVentas.Enabled = False
    CmdGrabar.Enabled = False
    FrmRetencion.Enabled = False
End Sub

Public Sub Activar_BS()
    TxtIvaBienMonIvaV.Enabled = True
    DCPorcenRetenIvaBienV.Enabled = True
    TxtIvaBienValRetV.Enabled = True
    TxtIvaSerMonIvaV.Enabled = True
    DCPorcenRetenIvaServV.Enabled = True
    TxtIvaSerValRetV.Enabled = True
End Sub

Public Sub Listar_Air()
  'Enceramos el espacio de cada usuario para emprezar con una nueva retencion
   sSQL = "DELETE * " _
        & "FROM Asiento_Ventas " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND T_No = " & Trans_No & " " _
        & "AND CodigoU = '" & CodigoUsuario & "' "
   Ejecutar_SQL_SP sSQL
  'Borra Asiento Air
   sSQL = "DELETE * " _
        & "FROM Asiento_Air " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND T_No = " & Trans_No & " " _
        & "AND Tipo_Trans = 'V' " _
        & "AND CodigoU = '" & CodigoUsuario & "' "
   Ejecutar_SQL_SP sSQL
  'Presentamos la malla Asiento Air
   sSQL = "SELECT * " _
        & "FROM Asiento_Air " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND CodigoU =  '" & CodigoUsuario & "' " _
        & "AND T_No = " & Trans_No & " " _
        & "AND Tipo_Trans = 'V' " _
        & "ORDER BY CodRet "
   Select_Adodc_Grid DGConceptoAirV, AdoAsientoAir, sSQL
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
       & "ORDER BY Porc "
  SelectDB_Combo DCPorcenIvaV, AdoPorIva, sSQL, "Porc"
 'Carga los Porcentajes de ICE
  sSQL = "SELECT * " _
       & "FROM Tabla_Por_ICE_IVA " _
       & "WHERE ICE <> " & Val(adFalse) & " " _
       & "AND Fecha_Inicio <= #" & FechaCodAir & "# " _
       & "AND Fecha_Final >= #" & FechaCodAir & "# " _
       & "ORDER BY Porc"
  SelectDB_Combo DCPorcenIceV, AdoPorIce, sSQL, "Porc"
 'Carga los conceptos de Retencion segun la fecha de Registro
  sSQL = "SELECT (Codigo & ' - ' & Concepto) As Detalle_Conceptos,* " _
       & "FROM Tipo_Concepto_Retencion " _
       & "WHERE Codigo <> '.' " _
       & "AND Fecha_Inicio <= #" & FechaCodAir & "# " _
       & "AND Fecha_Final >= #" & FechaCodAir & "# " _
       & "ORDER BY Codigo "
  SelectDB_Combo DCConceptoRetV, AdoConceptoRet, sSQL, "Detalle_Conceptos"
  DCConceptoRetV = "329 - Por Otros Servicios (N)"
End Sub

Private Sub TxtPorRetConAV_GotFocus()
  MarcarTexto TxtPorRetConAV
End Sub

Private Sub TxtPorRetConAV_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtPorRetConAV_LostFocus()
 If OP Then
    TxtValConAV = CTNumero(TxtBimpConAV, 2) * (CTNumero(TxtPorRetConAV, 2) / 100)
    Insertar_DataGridV
 End If
End Sub
