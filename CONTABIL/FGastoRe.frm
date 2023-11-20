VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form FGastosRet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "COMPRAS"
   ClientHeight    =   7920
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   10560
   ForeColor       =   &H8000000F&
   Icon            =   "FGastoRe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   10560
   Begin TabDlg.SSTab SSTCompras 
      Height          =   5610
      Left            =   105
      TabIndex        =   11
      Top             =   2205
      Width           =   10380
      _ExtentX        =   18309
      _ExtentY        =   9895
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   420
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&1.- Comprobante de Compra"
      TabPicture(0)   =   "FGastoRe.frx":0696
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label15"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label23"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "DCSustento"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "OpcSi"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "OpcNo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "FraDctoModificado"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame5"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame4"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "CmdAir"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "&2.- Conceptos AIR"
      TabPicture(1)   =   "FGastoRe.frx":06B2
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      Begin VB.CommandButton CmdAir 
         Caption         =   "&AIR"
         Height          =   444
         Left            =   9660
         Picture         =   "FGastoRe.frx":06CE
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Se ubica en la pestaña de Retenciones"
         Top             =   630
         Width           =   552
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
         Height          =   1485
         Left            =   105
         TabIndex        =   29
         Top             =   3045
         Width           =   4950
         Begin VB.TextBox TxtMontoIva 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   3465
            TabIndex        =   31
            Text            =   "0.00"
            ToolTipText     =   "Este valor se calcula automaticamente, es el resultado de aplicarle un porcentaje IVA a la Base Imponible gravada"
            Top             =   420
            Width           =   1275
         End
         Begin VB.TextBox TxtMontoIce 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   336
            Left            =   3465
            TabIndex        =   33
            Top             =   945
            Width           =   1275
         End
         Begin MSDataListLib.DataCombo DCPorcenIva 
            Bindings        =   "FGastoRe.frx":0BF4
            DataSource      =   "AdoPorIva"
            Height          =   315
            Left            =   945
            TabIndex        =   30
            ToolTipText     =   $"FGastoRe.frx":0C0C
            Top             =   420
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DCPorcenIce 
            Bindings        =   "FGastoRe.frx":0C9E
            DataSource      =   "AdoPorIce"
            Height          =   315
            Left            =   945
            TabIndex        =   32
            ToolTipText     =   $"FGastoRe.frx":0CB6
            Top             =   945
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
            TabIndex        =   77
            Top             =   945
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
            TabIndex        =   76
            Top             =   420
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
            TabIndex        =   75
            Top             =   945
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
            TabIndex        =   74
            Top             =   420
            Width           =   855
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
         Height          =   1485
         Left            =   5145
         TabIndex        =   34
         Top             =   3045
         Width           =   5055
         Begin MSDataListLib.DataCombo DCPorcenRetenIvaServ 
            Bindings        =   "FGastoRe.frx":0D47
            DataSource      =   "AdoRetIvaServicios"
            Height          =   315
            Left            =   3150
            TabIndex        =   39
            ToolTipText     =   "Corresponde al porcentaje retenido en el IVA generado en la prestación de servicios"
            Top             =   735
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DCPorcenRetenIvaBien 
            Bindings        =   "FGastoRe.frx":0D68
            DataSource      =   "AdoRetIvaBienes"
            Height          =   315
            Left            =   1470
            TabIndex        =   36
            ToolTipText     =   $"FGastoRe.frx":0D86
            Top             =   735
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
         End
         Begin VB.TextBox TxtIvaBienMonIva 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   336
            Left            =   1470
            MultiLine       =   -1  'True
            TabIndex        =   35
            Text            =   "FGastoRe.frx":0E12
            ToolTipText     =   $"FGastoRe.frx":0E17
            Top             =   420
            Width           =   1590
         End
         Begin VB.TextBox TxtIvaBienValRet 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   336
            Left            =   1470
            TabIndex        =   37
            Top             =   1050
            Width           =   1590
         End
         Begin VB.TextBox TxtIvaSerMonIva 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   336
            Left            =   3150
            MultiLine       =   -1  'True
            TabIndex        =   38
            Text            =   "FGastoRe.frx":0EB6
            ToolTipText     =   $"FGastoRe.frx":0EBB
            Top             =   420
            Width           =   1590
         End
         Begin VB.TextBox TxtIvaSerValRet 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   336
            Left            =   3150
            TabIndex        =   40
            Text            =   " "
            Top             =   1080
            Width           =   1590
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
            Left            =   105
            TabIndex        =   82
            Top             =   1050
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
            Left            =   105
            TabIndex        =   81
            Top             =   735
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
            Left            =   105
            TabIndex        =   80
            Top             =   420
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
            TabIndex        =   79
            Top             =   210
            Width           =   1590
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
            Left            =   1470
            TabIndex        =   78
            Top             =   210
            Width           =   1590
         End
      End
      Begin VB.Frame FraDctoModificado 
         Caption         =   "NOTAS DE DEBITO/NOTAS DE CREDITO"
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
         TabIndex        =   41
         Top             =   4620
         Visible         =   0   'False
         Width           =   10095
         Begin VB.ComboBox CNumSerieTresComp 
            DataSource      =   "AdoAux"
            Height          =   315
            Left            =   6090
            TabIndex        =   45
            Top             =   420
            Width           =   1275
         End
         Begin VB.TextBox TxtNumSerieUnoComp 
            Height          =   330
            Left            =   5040
            MaxLength       =   3
            TabIndex        =   43
            Text            =   "001"
            ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al código del establecimiento"
            Top             =   420
            Width           =   540
         End
         Begin VB.TextBox TxtNumSerieDosComp 
            Height          =   336
            Left            =   5565
            MaxLength       =   3
            TabIndex        =   44
            Text            =   "001"
            ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al punto dde emisión"
            Top             =   420
            Width           =   540
         End
         Begin VB.TextBox TxtNumAutComp 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   8715
            MaxLength       =   10
            TabIndex        =   47
            ToolTipText     =   $"FGastoRe.frx":0F51
            Top             =   432
            Width           =   1275
         End
         Begin MSDataListLib.DataCombo DCDctoModif 
            Bindings        =   "FGastoRe.frx":0FDD
            DataSource      =   "AdoTipoComprobante"
            Height          =   288
            Left            =   108
            TabIndex        =   42
            ToolTipText     =   "Corresponde al tipo de comprobante que ha sido originalmente modificado antre la emisión de una nota de débito o crédito"
            Top             =   420
            Width           =   4848
            _ExtentX        =   8546
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
         End
         Begin MSMask.MaskEdBox MBFechaEmiComp 
            Height          =   330
            Left            =   7455
            TabIndex        =   46
            ToolTipText     =   $"FGastoRe.frx":0FFE
            Top             =   420
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   582
            _Version        =   393216
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "0"
         End
         Begin VB.Label Label29 
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
            Left            =   8715
            TabIndex        =   86
            Top             =   210
            Width           =   1275
         End
         Begin VB.Label Label28 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fecha"
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
            Left            =   7455
            TabIndex        =   85
            Top             =   210
            Width           =   1170
         End
         Begin VB.Label Label27 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Serie          Numero"
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
            Left            =   5040
            TabIndex        =   84
            Top             =   210
            Width           =   2325
         End
         Begin VB.Label Label26 
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
            TabIndex        =   83
            Top             =   210
            Width           =   4845
         End
      End
      Begin VB.Frame Frame1 
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
         Height          =   1800
         Left            =   105
         TabIndex        =   17
         Top             =   1155
         Width           =   10095
         Begin VB.TextBox TxtBaseImpoNoObjIVA 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   3780
            MaxLength       =   14
            MultiLine       =   -1  'True
            TabIndex        =   25
            Text            =   "FGastoRe.frx":10AA
            ToolTipText     =   "En este campo se debe ingresar el valor del comprobante cuya base imponible esta gravado con la tarifa del 0% de IVA"
            Top             =   1365
            Width           =   1380
         End
         Begin VB.TextBox TxtNumAutor 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   6720
            MaxLength       =   37
            TabIndex        =   22
            Text            =   "000000000000000000000000000000000001"
            Top             =   630
            Width           =   3300
         End
         Begin MSMask.MaskEdBox MBFechaCad 
            Height          =   330
            Left            =   1365
            TabIndex        =   24
            ToolTipText     =   $"FGastoRe.frx":10B1
            Top             =   1365
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
         Begin VB.TextBox TxtBaseImpo 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   5250
            MaxLength       =   14
            MultiLine       =   -1  'True
            TabIndex        =   26
            Text            =   "FGastoRe.frx":1168
            ToolTipText     =   "En este campo se debe ingresar el valor del comprobante cuya base imponible esta gravado con la tarifa del 0% de IVA"
            Top             =   1365
            Width           =   1380
         End
         Begin VB.TextBox TxtNumSerietres 
            Height          =   336
            Left            =   5880
            MaxLength       =   9
            TabIndex        =   21
            Text            =   "0000001"
            ToolTipText     =   $"FGastoRe.frx":116F
            Top             =   630
            Width           =   855
         End
         Begin VB.TextBox TxtNumSerieDos 
            Height          =   336
            Left            =   5460
            MaxLength       =   3
            TabIndex        =   20
            Text            =   "001"
            ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al punto dde emisión"
            Top             =   630
            Width           =   435
         End
         Begin VB.TextBox TxtNumSerieUno 
            Height          =   336
            Left            =   5040
            MaxLength       =   3
            TabIndex        =   19
            Text            =   "001"
            ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al código del establecimiento"
            Top             =   630
            Width           =   435
         End
         Begin VB.TextBox TxtBaseImpoGrav 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   6720
            MultiLine       =   -1  'True
            TabIndex        =   27
            Text            =   "FGastoRe.frx":1212
            ToolTipText     =   $"FGastoRe.frx":1219
            Top             =   1365
            Width           =   1380
         End
         Begin VB.TextBox TxtBaseImpoIce 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   8316
            MultiLine       =   -1  'True
            TabIndex        =   28
            Text            =   "FGastoRe.frx":12C1
            ToolTipText     =   $"FGastoRe.frx":12C6
            Top             =   1365
            Width           =   1380
         End
         Begin MSDataListLib.DataCombo DCTipoComprobante 
            Bindings        =   "FGastoRe.frx":1358
            DataSource      =   "AdoTipoComp"
            Height          =   315
            Left            =   105
            TabIndex        =   18
            ToolTipText     =   $"FGastoRe.frx":1379
            Top             =   630
            Width           =   4950
            _ExtentX        =   8731
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
         Begin MSMask.MaskEdBox MBFechaRegis 
            Height          =   330
            Left            =   105
            TabIndex        =   23
            ToolTipText     =   $"FGastoRe.frx":1421
            Top             =   1365
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
         Begin VB.Label Label10 
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
            Height          =   330
            Left            =   105
            TabIndex        =   97
            Top             =   1050
            Width           =   1170
         End
         Begin VB.Label Label42 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " NO OBJ. IVA"
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
            TabIndex        =   96
            Top             =   1050
            Width           =   1380
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
            Left            =   1365
            TabIndex        =   71
            Top             =   1050
            Width           =   1275
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
            TabIndex        =   66
            Top             =   1050
            Width           =   1380
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
            TabIndex        =   73
            Top             =   1050
            Width           =   1380
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
            Left            =   5250
            TabIndex        =   72
            Top             =   1050
            Width           =   1380
         End
         Begin VB.Label Label3 
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
            Height          =   330
            Left            =   6720
            TabIndex        =   70
            Top             =   315
            Width           =   3270
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
            Height          =   330
            Left            =   5880
            TabIndex        =   69
            Top             =   315
            Width           =   855
         End
         Begin VB.Label Label6 
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
            Height          =   330
            Left            =   5040
            TabIndex        =   68
            Top             =   315
            Width           =   855
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
            Height          =   330
            Left            =   105
            TabIndex        =   67
            Top             =   315
            Width           =   4950
         End
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
         Height          =   228
         Left            =   3675
         TabIndex        =   14
         ToolTipText     =   $"FGastoRe.frx":14A9
         Top             =   315
         Value           =   -1  'True
         Width           =   636
      End
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
         Left            =   2940
         TabIndex        =   13
         ToolTipText     =   $"FGastoRe.frx":1541
         Top             =   315
         Width           =   636
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
         Height          =   4965
         Left            =   -74892
         TabIndex        =   49
         Top             =   420
         Width           =   10155
         Begin MSDataListLib.DataCombo DCRetFuente 
            Bindings        =   "FGastoRe.frx":15D9
            DataSource      =   "AdoRetFuente"
            Height          =   315
            Left            =   2415
            TabIndex        =   51
            Top             =   315
            Visible         =   0   'False
            Width           =   7575
            _ExtentX        =   13361
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
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
            TabIndex        =   50
            Top             =   315
            Visible         =   0   'False
            Width           =   2328
         End
         Begin VB.TextBox TxtValConA 
            Enabled         =   0   'False
            Height          =   336
            Left            =   8715
            TabIndex        =   60
            Top             =   1470
            Width           =   1275
         End
         Begin VB.TextBox TxtPorRetConA 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   336
            Left            =   8085
            TabIndex        =   59
            Top             =   1470
            Width           =   645
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
            TabIndex        =   62
            Text            =   "0.00"
            ToolTipText     =   "Sumatoria total de las retenciones"
            Top             =   4515
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
            TabIndex        =   56
            Text            =   "FGastoRe.frx":15F4
            Top             =   735
            Width           =   1905
         End
         Begin VB.TextBox TxtBimpConA 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   6720
            TabIndex        =   58
            Top             =   1470
            Width           =   1380
         End
         Begin VB.TextBox TxtNumUnoAutComRet 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   2415
            MaxLength       =   10
            MultiLine       =   -1  'True
            TabIndex        =   55
            ToolTipText     =   $"FGastoRe.frx":15FB
            Top             =   840
            Width           =   4215
         End
         Begin VB.TextBox TxtNumTresComRet 
            Height          =   336
            Left            =   1260
            MaxLength       =   9
            TabIndex        =   54
            Text            =   "0000001"
            ToolTipText     =   $"FGastoRe.frx":1687
            Top             =   840
            Width           =   1065
         End
         Begin VB.TextBox TxtNumDosComRet 
            Height          =   336
            Left            =   630
            MaxLength       =   3
            TabIndex        =   53
            Text            =   "001"
            ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al punto dde emisión"
            Top             =   840
            Width           =   540
         End
         Begin VB.TextBox TxtNumUnoComRet 
            Height          =   336
            Left            =   105
            MaxLength       =   3
            TabIndex        =   52
            Text            =   "001"
            ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al código del establecimiento"
            Top             =   840
            Width           =   540
         End
         Begin MSDataListLib.DataCombo DCConceptoRet 
            Bindings        =   "FGastoRe.frx":1729
            DataSource      =   "AdoConceptoRet"
            Height          =   315
            Left            =   105
            TabIndex        =   57
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
         Begin MSDataGridLib.DataGrid DGConceptoAir 
            Bindings        =   "FGastoRe.frx":1746
            Height          =   2595
            Left            =   105
            TabIndex        =   61
            Top             =   1890
            Width           =   9945
            _ExtentX        =   17542
            _ExtentY        =   4577
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   1
            RowHeight       =   19
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
         Begin VB.Label Label41 
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
            TabIndex        =   95
            Top             =   4515
            Width           =   2010
         End
         Begin VB.Label Label40 
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
            TabIndex        =   94
            Top             =   1260
            Width           =   1275
         End
         Begin VB.Label Label36 
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
            TabIndex        =   93
            Top             =   1260
            Width           =   645
         End
         Begin VB.Label Label35 
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
            TabIndex        =   92
            Top             =   1260
            Width           =   1380
         End
         Begin VB.Label Label34 
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
            TabIndex        =   91
            Top             =   1260
            Width           =   6630
         End
         Begin VB.Label Label33 
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
            Left            =   6720
            TabIndex        =   90
            Top             =   735
            Width           =   1380
         End
         Begin VB.Label Label32 
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
            TabIndex        =   89
            Top             =   630
            Width           =   4215
         End
         Begin VB.Label Label31 
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
            TabIndex        =   88
            Top             =   630
            Width           =   1065
         End
         Begin VB.Label Label30 
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
            TabIndex        =   87
            Top             =   630
            Width           =   1065
         End
      End
      Begin MSDataListLib.DataCombo DCSustento 
         Bindings        =   "FGastoRe.frx":1762
         DataSource      =   "AdoSustento"
         Height          =   315
         Left            =   105
         TabIndex        =   16
         ToolTipText     =   "En este campo de selección se despliega un lista de tipos de sustentos tributarios correspondientes a la transacción escogida"
         Top             =   840
         Width           =   9465
         _ExtentX        =   16695
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
      Begin VB.PictureBox Label23 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   420
         ScaleHeight     =   165
         ScaleWidth      =   1845
         TabIndex        =   63
         Top             =   2070
         Width           =   1905
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TIPO DE SUSTENTO TRIBUTARIO"
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
         TabIndex        =   15
         Top             =   630
         Width           =   9465
      End
      Begin VB.Label Label15 
         Caption         =   " DEVOLUCION DE I.V.A. "
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
         TabIndex        =   12
         Top             =   315
         Width           =   2535
      End
   End
   Begin VB.Frame FrmRetencion 
      Caption         =   "RETENCIONES DE IVA POR:"
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
      Left            =   105
      TabIndex        =   6
      Top             =   1155
      Width           =   9255
      Begin MSDataListLib.DataCombo DCRetISer 
         Bindings        =   "FGastoRe.frx":177C
         DataSource      =   "AdoRetIvaSerCC"
         Height          =   315
         Left            =   1260
         TabIndex        =   10
         Top             =   525
         Visible         =   0   'False
         Width           =   7890
         _ExtentX        =   13917
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.CheckBox ChRetB 
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
         Height          =   330
         Left            =   105
         TabIndex        =   7
         Top             =   210
         Width           =   960
      End
      Begin VB.CheckBox ChRetS 
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
         Height          =   330
         Left            =   105
         TabIndex        =   9
         Top             =   525
         Width           =   1170
      End
      Begin MSDataListLib.DataCombo DCRetIBienes 
         Bindings        =   "FGastoRe.frx":1799
         DataSource      =   "AdoRetIvaBienesCC"
         Height          =   315
         Left            =   1260
         TabIndex        =   8
         Top             =   210
         Visible         =   0   'False
         Width           =   7890
         _ExtentX        =   13917
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
   End
   Begin VB.TextBox TxtEmail 
      Height          =   330
      Left            =   1785
      MaxLength       =   60
      TabIndex        =   99
      ToolTipText     =   $"FGastoRe.frx":17B9
      Top             =   840
      Width           =   7575
   End
   Begin VB.PictureBox Code39Clt1 
      Height          =   480
      Left            =   9450
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   100
      Top             =   1785
      Width           =   1200
   End
   Begin MSDataListLib.DataCombo DCProveedor 
      Bindings        =   "FGastoRe.frx":1845
      DataSource      =   "AdoClientes"
      Height          =   315
      Left            =   1365
      TabIndex        =   3
      ToolTipText     =   "En este combo de selección se despliega una lista de todos los proveedores"
      Top             =   420
      Width           =   5895
      _ExtentX        =   10398
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
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Height          =   750
      Left            =   9450
      Picture         =   "FGastoRe.frx":185F
      Style           =   1  'Graphical
      TabIndex        =   64
      ToolTipText     =   "Grabar"
      Top             =   105
      Width           =   960
   End
   Begin VB.CommandButton CmdCerrar 
      Caption         =   "&Salir"
      Height          =   765
      Left            =   9450
      Picture         =   "FGastoRe.frx":1B69
      Style           =   1  'Graphical
      TabIndex        =   65
      ToolTipText     =   "Salir"
      Top             =   945
      Width           =   990
   End
   Begin MSAdodcLib.Adodc AdoSustento 
      Height          =   330
      Left            =   2940
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
      Left            =   420
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
      Left            =   420
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
      Left            =   420
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
      Left            =   420
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
      Left            =   2940
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
   Begin MSAdodcLib.Adodc AdoCaTrTiCom 
      Height          =   330
      Left            =   420
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
      Caption         =   "Catal.Tributarios y Tipos de Comprobantes"
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
      Left            =   420
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
   Begin MSAdodcLib.Adodc AdoTransCompras 
      Height          =   330
      Left            =   420
      Top             =   5145
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
      Left            =   420
      Top             =   5460
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
      Left            =   2940
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
   Begin MSAdodcLib.Adodc AdoRetIvaSerCC 
      Height          =   330
      Left            =   2940
      Top             =   4515
      Visible         =   0   'False
      Width           =   3270
      _ExtentX        =   5768
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
      Caption         =   "RetencionFuenteServicios"
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
      Left            =   2940
      Top             =   4830
      Visible         =   0   'False
      Width           =   3270
      _ExtentX        =   5768
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
      Caption         =   "RetencionFuenteBienes"
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
      Left            =   420
      Top             =   6090
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
      Left            =   2940
      Top             =   5460
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
      Left            =   420
      Top             =   5775
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
      Left            =   2940
      Top             =   5145
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
   Begin MSMask.MaskEdBox MBFechaEmi 
      Height          =   330
      Left            =   105
      TabIndex        =   1
      ToolTipText     =   $"FGastoRe.frx":1FAB
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
   Begin MSAdodcLib.Adodc AdoAsientos 
      Height          =   330
      Left            =   2940
      Top             =   5775
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
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " EMAIL/CORREO:"
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
      TabIndex        =   98
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label9 
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
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   1170
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
      ForeColor       =   &H8000000F&
      Height          =   330
      Left            =   7560
      TabIndex        =   5
      Top             =   420
      Width           =   1800
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
      Left            =   7245
      TabIndex        =   4
      Top             =   420
      Width           =   330
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
      Height          =   330
      Left            =   1365
      TabIndex        =   2
      Top             =   105
      Width           =   7995
   End
End
Attribute VB_Name = "FGastosRet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MBFecha As MaskEdBox
Dim FechaRegis As Date
Dim CapDm, CapDcto, Cap1, Captc, Cap, Ct, Espizq, Espder, ValorP, ch, Ch1, CodProv, CodSus As String
Dim OP As Boolean
Dim Rb, Rs, Rf, cod As Byte
Dim AniocadAux, SumAnio, Aniocad As Integer
Dim CalIsMi, CalmIva, CalIbMi, ac As Double

Private Sub ChRetB_Click()
  If ChRetB.value <> 0 Then
     ch = 1
     Ch1 = "B"
     DCRetIBienes.Visible = True
     TxtIvaBienMonIva.Enabled = True
     DCPorcenRetenIvaBien.Enabled = True
     TxtIvaBienValRet.Enabled = True
  Else
     Ch1 = "S"
     DCRetIBienes.Visible = False
     TxtIvaBienMonIva.Enabled = False
     DCPorcenRetenIvaBien.Enabled = False
     TxtIvaBienValRet.Enabled = False
  End If
  If ChRetB.value <> 0 And ChRetB.value <> 0 Then
     Ch1 = "X"
  End If
End Sub

Private Sub ChRetS_Click()
  If ChRetS.value <> 0 Then
     ch = 1
     Ch1 = "S"
     'Carga_ConceptosVerif
     DCRetISer.Visible = True
     TxtIvaSerMonIva.Enabled = True
     DCPorcenRetenIvaServ.Enabled = True
     TxtIvaSerValRet.Enabled = True
  Else
     DCRetISer.Visible = False
     TxtIvaSerMonIva.Enabled = False
     DCPorcenRetenIvaServ.Enabled = False
     TxtIvaSerValRet.Enabled = False
  End If
  If ChRetB.value <> 0 And ChRetS.value <> 0 Then
     Ch1 = "X"
     'Carga_ConceptosVerif
  End If
End Sub

Private Sub CmdAir_Click()
  SSTCompras.Tab = 1
  TxtNumUnoComRet.SetFocus
End Sub

Private Sub CmdCerrar_Click()
  'Borra Asiento Compras
  sSQL = "DELETE * " _
       & "FROM Asiento_Compras " _
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
       & "AND Tipo_Trans = 'C' "
  ConectarAdoExecute sSQL
  Unload Me
End Sub

Private Sub CmdGrabar_Click()
Dim ProvAnt As String
  ProvAnt = DCProveedor
 'Valido por si acaso exista algun valor con 0
  TextoValido TxtIvaBienMonIva, True, , 2
  TextoValido TxtBaseImpo, True, , 2
  TextoValido TxtBaseImpoGrav, True, , 2
  TextoValido TxtBaseImpoIce, True, , 2
  TextoValido TxtMontoIva, True, , 2
  TextoValido TxtMontoIce, True, , 2
  TextoValido TxtIvaBienMonIva, True, , 2
  TextoValido TxtIvaBienValRet, True, , 2
  TextoValido TxtIvaSerMonIva, True, , 2
  TextoValido TxtIvaSerValRet, True, , 2
  'Pregunto antes de grabar
  Titulo = "GRABAR RETENCION"
  Mensajes = "Desea Grabar los Datos"
  If BoxMensaje = vbYes Then
    'Borrar todas las transacciones de compras que tengan la misma factura y la misma retencion
    'del mismo proveedor
     Grabacion
     If Ln_SRI < 0 Then
        Mensajes = "Los Datos fueron grabados correctamente" & vbCrLf _
                 & "Desea ingresar otra transacción"""
        If BoxMensaje = vbYes Then
           Ln_No = 1
           Limpiar_Controles
           Listar_Air
           SSTCompras.Tab = 0
           DCProveedor = ProvAnt
           ChRetB.SetFocus
        Else
           Unload FGastosRet
        End If
     Else
        MsgBox "Datos Modificados Correctamente"
        Unload FGastosRet
     End If
   Else
      ChRetB.SetFocus
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
          TxtPorRetConA.Enabled = False
          If Not .EOF Then
             Cadena = .Fields("Ingresar_Porcentaje")

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

Private Sub DCDctoModif_LostFocus()
  Captura_TipoComprobante_DctoModificado
End Sub

Private Sub DCPorcenIva_GotFocus()
  MarcarTexto DCPorcenIva
End Sub

Private Sub DCPorcenIva_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCProveedor_LostFocus()
  SSTCompras.Tab = 0
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
             Label41.Caption = "PROVEEDOR: "
            'Busca y captura el codigo de Porcentaje IVA
            CodigoCliente = .Fields("Codigo")
            DireccionCli = .Fields("Direccion")
            CICliente = .Fields("CI_RUC")
            TipoBenef = .Fields("TD")
            TxtEmail = .Fields("Email")
            If .Fields("RISE") Then Label41.Caption = Label41.Caption & " RISE"
            If .Fields("Especial") Then Label41.Caption = Label41.Caption & "Contribuyente especial "
            LblNumIdent = CICliente
            LblTD.Caption = TipoBenef
            TxtNumSerietres = "0000001"
            'Aqui despliego el ultimo numero de la Transaccion
            sSQL = "SELECT TOP 1 * " _
                 & "FROM Trans_Compras " _
                 & "WHERE IdProv = '" & CodigoCliente & "' " _
                 & "AND Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "ORDER BY Secuencial DESC "
            SelectAdodc AdoAux, sSQL
            If AdoAux.Recordset.RecordCount > 0 Then TxtNumSerietres = AdoAux.Recordset.Fields("Secuencial")
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

Private Sub DCTipoComprobante_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCTipoComprobante_LostFocus()
  If IsNumeric(DCTipoComprobante.Text) Then
     MsgBox "No ingrese números. Vuelva a seleccionar.", vbInformation, "Aviso"
     DCTipoComprobante.Text = ""
     Carga_TipoComprobante (DCSustento)
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
             Mifecha = BuscarFecha(.Fields("FechaEmiRet"))
             Codigo1 = .Fields("AutRetencion")
             J = .Fields("A_No")
             sSQL = "DELETE * " _
                  & "FROM Asiento_Air " _
                  & "WHERE Item = '" & NumEmpresa & "' " _
                  & "AND CodigoU = '" & CodigoUsuario & "' " _
                  & "AND IdProv = '" & CodigoCliente & "' " _
                  & "AND T_No = " & Trans_No & " " _
                  & "AND Tipo_Trans = 'C' " _
                  & "AND A_No = " & J & " " _
                  & "AND CodRet = '" & Codigo & "' "
             ConectarAdoExecute sSQL
         End If
         sSQL = "SELECT * " _
              & "FROM Asiento_Air " _
              & "WHERE CodRet <> '.' " _
              & "AND Item = '" & NumEmpresa & "' " _
              & "AND CodigoU = '" & CodigoUsuario & "' " _
              & "AND T_No = " & Trans_No & " " _
              & "AND Tipo_Trans = 'C' " _
              & "ORDER BY CodRet "
         SelectDataGrid DGConceptoAir, AdoAsientoAir, sSQL
         Calculo_Sumatoria
        End With
  End If
End If
End Sub

Private Sub MBFechaCad_GotFocus()
  MarcarTexto MBFechaCad
End Sub

Private Sub MBFechaCad_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaEmi_GotFocus()
  MarcarTexto MBFechaEmi
End Sub

Private Sub MBFechaEmi_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaEmiComp_GotFocus()
  MarcarTexto MBFechaEmiComp
End Sub

Private Sub MBFechaEmiComp_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaEmiComp_LostFocus()
  FechaValida MBFechaEmiComp
End Sub

Private Sub OpcNo_LostFocus()
  If OpcNo.value = True Then ValorP = "N"
End Sub

Private Sub OpcSi_LostFocus()
  If OpcSi.value = True Then ValorP = "S"
End Sub

Private Sub SSTCompras_Click(PreviousTab As Integer)
  Select Case PreviousTab
         Case 0: If ChRetF.Visible Then ChRetF.SetFocus 'Else CTP.SetFocus
         Case 1: OpcSi.SetFocus
  End Select
End Sub

Private Sub TxtBaseImpo_GotFocus()
  MarcarTexto TxtBaseImpo
End Sub

Private Sub TxtBaseImpo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtBaseImpo_LostFocus()
  TextoValido TxtBaseImpo, True, , 2
End Sub

Private Sub TxtBaseImpoGrav_GotFocus()
  MarcarTexto TxtBaseImpoGrav
End Sub

Private Sub TxtBaseImpoGrav_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtBaseImpoGrav_LostFocus()
  TextoValido TxtBaseImpoGrav, True, , 2
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

Private Sub TxtBaseImpoNoObjIVA_GotFocus()
  MarcarTexto TxtBaseImpoNoObjIVA
End Sub

Private Sub TxtBaseImpoNoObjIVA_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtBaseImpoNoObjIVA_LostFocus()
  TextoValido TxtBaseImpoNoObjIVA, True, , 2
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
        If (cod = 4) Or (cod = 5) Then
           DCDctoModif.SetFocus
        Else
           'TxtNumConParPol.SetFocus
        End If
     Else
        'TxtPorRetConA.SetFocus
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
   SetAdoFields "BaseImp", CTNumero(TxtBimpConA, 2)
   SetAdoFields "Porcentaje", CTNumero(TxtPorRetConA, 2) / 100
   SetAdoFields "ValRet", CTNumero(TxtValConA, 2)
   SetAdoFields "EstabRetencion", TxtNumUnoComRet
   SetAdoFields "PtoEmiRetencion", TxtNumDosComRet
   SetAdoFields "SecRetencion", CTNumero(TxtNumTresComRet)
   SetAdoFields "AutRetencion", TxtNumUnoAutComRet
   SetAdoFields "FechaEmiRet", MBFechaRegis
   SetAdoFields "Cta_Retencion", SinEspaciosIzq(DCRetFuente)
   SetAdoFields "EstabFactura", TxtNumSerieUno
   SetAdoFields "PuntoEmiFactura", TxtNumSerieDos
   SetAdoFields "Factura_No", CTNumero(TxtNumSerietres)
   SetAdoFields "IdProv", CodigoCliente
   SetAdoFields "A_No", Ln_No
   SetAdoFields "T_No", Trans_No
   SetAdoFields "Tipo_Trans", "C"
   SetAdoUpdate
          
  'Despliega los datos en el DataGrid
   sSQL = "SELECT * " _
        & "FROM Asiento_Air " _
        & "WHERE CodRet <> '.' " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "AND CodigoU = '" & CodigoUsuario & "' " _
        & "AND T_No = " & Trans_No & " " _
        & "AND Tipo_Trans = 'C' " _
        & "ORDER BY CodRet "
   SelectDataGrid DGConceptoAir, AdoAsientoAir, sSQL
   
  'Se situa en el combo de retención AIR
   If ChRetF.Visible Then DCRetFuente.SetFocus Else TxtNumUnoComRet.SetFocus
   
  'Realiza la Sumatoria de las Retenciones
   ac = ac + TxtValConA
   TxtTotalReten = ac
End If
RatonNormal
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

 'Coloca el valor de Monto IVA dependiendo si se activo Bienes o Servicios
  If ChRetB + ChRetS = 0 Then
     TxtIvaBienMonIva = TxtMontoIva
  End If
  If ChRetB.value <> 0 Then
     TxtIvaBienMonIva = TxtMontoIva
     TxtIvaSerMonIva = 0
  Else
     If ChRetS.value <> 0 Then
        TxtIvaSerMonIva = TxtMontoIva
        TxtIvaBienMonIva = 0
     End If
  End If
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

Private Sub DCPorcenRetenIvaBien_LostFocus()
  CodRetBien = 0
  If Not IsNumeric(DCPorcenRetenIvaBien) Then
     MsgBox "No ingrese caracteres. Vuelva a seleccionar.", vbInformation, "Aviso"
     DCPorcenRetenIvaBien = ""
     Carga_RetencionIvaBienes_Servicios
     DCPorcenRetenIvaBien.SetFocus
  Else
    'Busca y captura el codigo de Porcentaje retencion Iva Bienes
     With AdoRetIvaBienes.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
         .Find ("Porc = " & SinEspaciosDer(DCPorcenRetenIvaBien) & " ")
          If Not .EOF Then CodRetBien = .Fields("Codigo")
      Else
         MsgBox "Código incorrecto", vbInformation, "Aviso"
      End If
     End With
     Total_IVA = CTNumero(TxtIvaBienMonIva, 2)
    'Calcula la retencion Iva Bienes
     CalIbMi = (Total_IVA * CInt(DCPorcenRetenIvaBien)) / 100
     TxtIvaBienValRet = CalIbMi
  End If
  TxtIvaSerMonIva = Format(CTNumero(TxtMontoIva, 2) - CTNumero(TxtIvaBienMonIva, 2), "#,##0.00")
End Sub

Private Sub DCPorcenRetenIvaServ_LostFocus()
  CodRetServ = 0
 'Activo el casillero para que ingrese el valor si el porcentaje es 70/100
  If DCPorcenRetenIvaServ = "70/100" Then
     Ct = "Si"
     TxtIvaSerValRet.Text = ""
     TxtIvaSerValRet.Enabled = True
     TxtIvaSerValRet.SetFocus
  Else
     If Not IsNumeric(DCPorcenRetenIvaServ) Then
        MsgBox "No ingrese caracteres. Vuelva a seleccionar.", vbInformation, "Aviso"
        DCPorcenRetenIvaServ = ""
        Carga_RetencionIvaBienes_Servicios
        DCPorcenRetenIvaServ.SetFocus
     End If
  End If
    
 'Busca captura el codigo de Porcentaje retencion Iva Servicios
  With AdoRetIvaServicios.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Porc = " & SinEspaciosDer(DCPorcenRetenIvaServ) & " ")
       If Not .EOF Then CodRetServ = .Fields("Codigo")
   Else
      MsgBox "Código Incorrecto", vbInformation, "Aviso"
   End If
  End With
  Ct = "No"
  Total_IVA = 0
  Total_IVA = CTNumero(TxtIvaSerMonIva, 2)
  If DCPorcenRetenIvaServ = "70/100" Then
  Else
     CalIsMi = (Total_IVA * CInt(DCPorcenRetenIvaServ)) / 100
     TxtIvaSerValRet = CalIsMi
     TxtIvaSerValRet.Enabled = False
  End If
  SSTCompras.Tab = 0
  SSTCompras.SetFocus
End Sub

Private Sub DCSustento_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCSustento_LostFocus()
  If IsNumeric(DCSustento.Text) Then
     Carga_CreditoTributario
     DCSustento.SetFocus
     Carga_TipoComprobante DCSustento
  Else
     Carga_TipoComprobante DCSustento
  End If
End Sub

Private Sub Form_Activate()
  Ln_No = 1
  Ln_SRI = -1
  Trans_No = 250
  sSQL = Listar_Meses
  SelectAdodc AdoAux, sSQL
  Carga_Datos_Iniciales MBFecha, Nuevo
End Sub

Private Sub Form_Load()
  CentrarForm FGastosRet
  ConectarAdodc AdoAux
  ConectarAdodc AdoSustento
  ConectarAdodc AdoTipoIdentificacion
  ConectarAdodc AdoTipoComprobante
  ConectarAdodc AdoRetIvaBienes
  ConectarAdodc AdoRetIvaServicios
  ConectarAdodc AdoPorIce
  ConectarAdodc AdoPorIva
  ConectarAdodc AdoConceptoRet
  ConectarAdodc AdoAsientoCompras
  ConectarAdodc AdoTransCompras
  ConectarAdodc AdoAsientoAir
  ConectarAdodc AdoTransAir
  ConectarAdodc AdoClientes
  ConectarAdodc AdoRetFuente
  ConectarAdodc AdoRetIvaSerCC
  ConectarAdodc AdoRetIvaBienesCC
  ConectarAdodc AdoAsientos
End Sub

Private Sub MBFechaCad_LostFocus()
  'Verifico que la fecha de caducidad no sea mayor a la de emisión
   FechaValida MBFechaCad
   If MBFechaCad = "00/00/0000" Then
      MsgBox "Fecha no válida, vuelva a ingresar", vbInformation, "Aviso"
      MBFechaCad.SetFocus
   Else
       'Captura el año de la fecha de emisión
        Anio = Year(MBFechaEmi)
        SumAnio = Anio + 1  'Emisión + 1 año
        Aniocad = Year(MBFechaCad)
        AniocadAux = Aniocad + 1 'Asigno en otra variable el año de caducidad
        'Verifica si el año de caducidad es menor que el año de Emisión
        If (Aniocad < Anio) Then
           MsgBox "La Fecha de Caducidad no debe ser < a la Fecha de Emisión", vbInformation, "Aviso"
           FechaValida MBFechaCad
           MBFechaCad.SetFocus
        Else
           'Verifica si el año de caducidad es mayor con 2 años al año de Emisión
           If (Aniocad = AniocadAux) Then
              MsgBox "Hola La Fecha de Caducidad no debe sobrepasar dos años, máximo uno", vbInformation, "Aviso"
              FechaValida MBFechaCad
              MBFechaCad.SetFocus
           Else
             If (Aniocad > AniocadAux) Then
                MsgBox "La Fecha de Caducidad no debe sobrepasar dos años, máximo uno", vbInformation, "Aviso"
                FechaValida MBFechaCad
                MBFechaCad.SetFocus
             End If
           End If
        End If
   End If
  
   If CFechaLong(MBFechaRegis) < CFechaLong("01/01/2000") Then
      MsgBox "La Fecha de Registro debe ser mayor que 01/01/2000", vbInformation, "Aviso"
      MBFechaRegis = "01/01/2000"
      MBFechaRegis.SetFocus
   Else
      If CFechaLong(MBFechaRegis) < CFechaLong(MBFechaEmi) Then
         MsgBox "La Fecha de Registro debe ser mayor o igual que la Fecha de Emisión", vbInformation, "Aviso"
         MBFechaRegis.SetFocus
      End If
   End If
End Sub

Private Sub MBFechaEmi_LostFocus()
  FechaValida MBFechaEmi
 'Controla que la Fecha de Emisiòn este entre 01/01/2000 en adelante
  If CFechaLong(MBFechaEmi) < CFechaLong("01/01/2000") Then
     MsgBox "La Fecha de Emisión debe ser mayor que 01/01/2000", vbInformation, "Aviso"
     MBFechaEmi = "01/01/2000"
     MBFechaEmi.SetFocus
  End If
  MBFechaRegis = MBFechaEmi
End Sub

Private Sub MBFechaRegis_GotFocus()
  MarcarTexto MBFechaRegis
End Sub

Private Sub MBFechaRegis_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaRegis_LostFocus()
  FechaValida MBFechaRegis
  Carga_ConceptosRetencion MBFechaRegis
  Label13.Caption = " TARIFA " & Porc_IVA * 100 & "%"
  'Controla que la Fecha de Registro este entre 01/01/2000 en adelante
  FechaValida MBFechaRegis
 'Carga la Tabla de Porcentaje Iva en el DataCombo
 'Carga_PorcentajeIva (MBFechaRegis)
  Carga_ConceptosRetencion MBFechaRegis
End Sub

Private Sub TxtIvaBienMonIva_GotFocus()
  MarcarTexto TxtIvaBienMonIva
End Sub

Private Sub TxtIvaBienMonIva_LostFocus()
  'MsgBox CTNumero(TxtIvaBienMonIva, 2)
  TextoValido TxtIvaBienMonIva, True, , 0
End Sub

Private Sub TxtIvaSerMonIva_GotFocus()
  MarcarTexto TxtIvaSerMonIva
End Sub

Private Sub TxtIvaSerMonIva_LostFocus()
  TextoValido TxtIvaSerMonIva, True, , 0
  'Verifica el Monto Iva Servicios
  If CDbl(TxtIvaBienMonIva) + CDbl(TxtIvaSerMonIva) > CDbl(TxtMontoIva) Then
     MsgBox "Monto IVA Servicios no puede ser > que Monto IVA", vbInformation, "Aviso de Compras"
     TxtIvaSerMonIva.Text = ""
     TxtIvaSerMonIva.SetFocus
  End If
End Sub

Private Sub TxtIvaSerValRet_GotFocus()
  MarcarTexto TxtIvaSerValRet
End Sub

Private Sub TxtIvaSerValRet_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
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

Private Sub TxtNumAutComp_GotFocus()
  MarcarTexto TxtNumAutComp
End Sub

Private Sub TxtNumAutComp_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumAutComp_LostFocus()
  If Val(TxtNumAutComp) <= 0 Then TxtNumAutComp = "0000000000"
  TxtNumAutComp = Format(Val(Round(TxtNumAutComp)), String(10, "0"))
End Sub

Private Sub TxtNumAutor_GotFocus()
   MarcarTexto TxtNumAutor
End Sub

Private Sub TxtNumAutor_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumAutor_LostFocus()
  If Val(TxtNumAutor) <= 0 Then TxtNumAutor = "0000000001"
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

Private Sub TxtNumTresComRet_GotFocus()
  MarcarTexto TxtNumTresComRet
End Sub

Private Sub TxtNumTresComRet_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumTresComRet_LostFocus()
  If Val(TxtNumTresComRet) <= 0 Then TxtNumTresComRet = "000000000"
  TxtNumTresComRet = Format(Val(Round(TxtNumTresComRet)), "000000000")
  'Calcula la sumatoria de Monto Iva Bienes, Monto Iva Servicios y Base Imponible
  TxtSumatoria = CTNumero(TxtBaseImpoNoObjIVA, 2) + CTNumero(TxtBaseImpo, 2) + CTNumero(TxtBaseImpoGrav, 2)
End Sub

Private Sub TxtNumUnoAutComRet_GotFocus()
  MarcarTexto TxtNumUnoAutComRet
End Sub

Private Sub TxtNumUnoAutComRet_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumUnoAutComRet_LostFocus()
  If Val(TxtNumUnoAutComRet) <= 0 Then TxtNumUnoAutComRet = "0000000000"
  TxtNumUnoAutComRet = Format(Val(Round(TxtNumUnoAutComRet)), String(10, "0"))
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

Public Sub Leer_Clientes()
Dim Cta_Prov As String
 'Carga en el Data Combo los Clientes con su RUC
  Cta_Prov = Leer_Seteos_Ctas("Cta_Proveedores")
  sSQL = "SELECT C.Codigo,C.Cliente,C.Direccion,C.TD,C.CI_RUC,C.DirNumero,C.Telefono,C.Grupo,C.RISE," _
       & "C.Especial,C.Email,COUNT(CP.TC) As CxP " _
       & "FROM Clientes As C,Catalogo_CxCxP As CP " _
       & "WHERE CP.Item = '" & NumEmpresa & "' " _
       & "AND CP.Periodo = '" & Periodo_Contable & "' " _
       & "AND CP.Cta = '" & Cta_Prov & "' " _
       & "AND C.Codigo = CP.Codigo " _
       & "GROUP BY C.Codigo,C.Cliente,C.Direccion,C.TD,C.CI_RUC,C.DirNumero,C.Telefono,C.Grupo,C.RISE," _
       & "C.Especial,C.Email " _
       & "ORDER BY Cliente "
  SelectDBCombo DCProveedor, AdoClientes, sSQL, "Cliente"
End Sub

Public Sub Carga_CreditoTributario()
Dim Fecha_TT As String
  'Carga la Tabla de Catalogos Tributarios al DataCombo
  Fecha_TT = BuscarFecha(MBFechaEmi)
  sSQL = "SELECT (Credito_Tributario & ' - ' & Descripcion) As Sustento,* " _
       & "FROM Tipo_Tributario " _
       & "WHERE Fecha_Inicio <= #" & Fecha_TT & "# " _
       & "AND #" & Fecha_TT & "# <= Fecha_Final " _
       & "ORDER BY Credito_Tributario "
  SelectDBCombo DCSustento, AdoSustento, sSQL, "Sustento"
End Sub

Public Sub Carga_TipoComprobante(TipoSustento As String)
Dim CargaTC As String
    CargaTC = SinEspaciosIzq(TipoSustento)
    
     sSQL = "SELECT Tipo_Comprobante_Codigo, Descripcion " _
          & "FROM Tipo_Comprobante " _
          & "WHERE Tipo_Comprobante_Codigo <> 100 " _
          & "ORDER BY Descripcion "
     SelectDBCombo DCDctoModif, AdoTipoComprobante, sSQL, "Descripcion"
    
    'Capturo el codigo del Tipo de Catalogo Tributario
     Cap = CargaTC
            
    'Busco el codigo en la tabla Tipo Comprobante///descripcion
    Cadena = Ninguno
    
    Carga_CreditoTributario
    With AdoSustento.Recordset
     If .RecordCount > 0 Then
        .MoveFirst
        .Find ("Credito_Tributario = '" & CargaTC & "' ")
         If Not .EOF Then Cadena = .Fields("Codigo_Tipo_Comprobante")
     End If
    End With
    
    DCSustento = TipoSustento
    Cadena = Replace(Cadena, " ", ",")
    sSQL = "SELECT * " _
         & "FROM Tipo_Comprobante " _
         & "WHERE Tipo_Comprobante_Codigo IN (" & Cadena & ") "
    If TipoBenef = "R" Then
       sSQL = sSQL & "AND R <> " & Val(adFalse) & " "
    Else
       sSQL = sSQL & "AND C <> " & Val(adFalse) & " "
    End If
    sSQL = sSQL & "ORDER BY Tipo_Comprobante_Codigo "
    SelectDBCombo DCTipoComprobante, AdoTipoComprobante, sSQL, "Descripcion"
End Sub

Public Sub Captura_TipoComprobante()
  'Captura lo que tiene el Combo de Tipo de Comprobante
  Label15.Caption = "Fechas de " & DCTipoComprobante
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
       End If
    End If
  End With
  If (cod = 4) Or (cod = 5) Then
     FraDctoModificado.Visible = True
     Documento_Modificado
     'Carga en el combo de Documentos Modificados los
     'Tipos de Comprobantes
     sSQL = "SELECT Tipo_Comprobante_Codigo, Descripcion " _
          & "FROM Tipo_Comprobante " _
          & "WHERE Tipo_Comprobante_Codigo <> 100 " _
          & "ORDER BY Descripcion "
     SelectDBCombo DCDctoModif, AdoTipoComprobante, sSQL, "Descripcion"
  Else
     FraDctoModificado.Visible = False
  End If
End Sub

Public Sub Captura_TipoComprobante_DctoModificado()
  CapDcto = Trim$(DCDctoModif.Text)
     
  'Busca que sea igual a la Descripcion
  With AdoTipoComprobante.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Descripcion = '" & CapDcto & "' ")
       If Not .EOF Then
          CapDm = .Fields("Tipo_Comprobante_Codigo")
       Else
          MsgBox "Vuelva a seleccionar", vbInformation, "Aviso"
       End If
    End If
  End With
  If (cod = 4) Or (cod = 5) Then
     FraDctoModificado.Visible = True
     'Verifico si hay documentos modificados de ese Proveedor
     Documento_Modificado
  Else
     FraDctoModificado.Visible = False
  End If
End Sub

Sub Documento_Modificado()
    'Facturas Emitidas del proveedor
     sSQL = "SELECT * " _
          & "FROM Trans_Compras " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND IdProv = '" & CodigoCliente & "' " _
          & "ORDER BY Secuencial "
     SelectAdodc AdoAux, sSQL
     With AdoAux.Recordset
      If .RecordCount > 0 Then
          Do While Not .EOF
             CNumSerieTresComp.AddItem .Fields("Secuencial")
            .MoveNext
          Loop
      End If
     End With
End Sub

Public Sub Carga_RetencionIvaBienes_Servicios()
  sSQL = "SELECT * " _
       & "FROM Tabla_Por_IVA " _
       & "WHERE Bienes <> " & Val(adFalse) & " " _
       & "ORDER BY Porc "
  SelectDBCombo DCPorcenRetenIvaBien, AdoRetIvaBienes, sSQL, "Porc"
  
  sSQL = "SELECT * " _
       & "FROM Tabla_Por_IVA " _
       & "WHERE Servicios <> " & Val(adFalse) & " " _
       & "ORDER BY Porc "
  SelectDBCombo DCPorcenRetenIvaServ, AdoRetIvaServicios, sSQL, "Porc"
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
  SelectDBCombo DCPorcenIva, AdoPorIva, sSQL, "Porc"
 'Carga los Porcentajes de ICE
  sSQL = "SELECT * " _
       & "FROM Tabla_Por_ICE_IVA " _
       & "WHERE ICE <> " & Val(adFalse) & " " _
       & "AND Fecha_Inicio <= #" & FechaCodAir & "# " _
       & "AND Fecha_Final >= #" & FechaCodAir & "# " _
       & "ORDER BY Porc"
  SelectDBCombo DCPorcenIce, AdoPorIce, sSQL, "Porc"
 'Carga los conceptos de Retencion segun la fecha de Registro
  sSQL = "SELECT (Codigo & ' - ' & Concepto) As Detalle_Conceptos,* " _
       & "FROM Tipo_Concepto_Retencion " _
       & "WHERE Codigo <> '.' " _
       & "AND Fecha_Inicio <= #" & FechaCodAir & "# " _
       & "AND Fecha_Final >= #" & FechaCodAir & "# " _
       & "ORDER BY Codigo "
  SelectDBCombo DCConceptoRet, AdoConceptoRet, sSQL, "Detalle_Conceptos"
  'DCConceptoRet = "329 - Por Otros Servicios (N)"
End Sub

Public Sub Limpiar_Controles()
  ac = 0
  DCProveedor.Text = ""
  DCRetIBienes.Visible = False
  DCRetISer.Visible = False
  ChRetB.value = False
  ChRetF.value = False
  ChRetS.value = False
  LblNumIdent.Caption = ""
  LblTD.Caption = ""
  OpcNo.value = True
  DCSustento.Text = ""
  DCTipoComprobante.Text = ""
  TxtNumSerieUno.Text = ""
  TxtNumSerieDos.Text = ""
  TxtNumSerietres.Text = ""
  TxtNumAutor.Text = ""
  FechaValida MBFechaEmi
  FechaValida MBFechaRegis
  FechaValida MBFechaCad
  TxtBaseImpo.Text = ""
  TxtBaseImpoGrav.Text = ""
  TxtBaseImpoIce.Text = ""
  DCPorcenIva.Text = ""
  TxtMontoIva.Text = ""
  DCPorcenIce.Text = ""
  TxtMontoIce.Text = ""
  TxtIvaBienMonIva.Text = ""
  DCPorcenRetenIvaBien.Text = ""
  TxtIvaBienValRet.Text = ""
  TxtIvaSerMonIva.Text = ""
  DCPorcenRetenIvaServ.Text = ""
  TxtIvaSerValRet.Text = ""
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
  
  'Limpia la grilla
  'Borra Asiento Air
  sSQL = "DELETE * " _
       & "FROM Asiento_Air " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND Tipo_Trans = 'C' "
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

Public Sub Carga_Datos_Iniciales(MBFecha As MaskEdBox, EsNuevo As Boolean)
    Encerar_Var
    Limpiar_Controles
    Listar_Air
   'Cargo el No.Autorización de las retenciones
    TxtNumUnoAutComRet = AutorizaRet
   'Carga el Sustento Tributario
    Carga_CreditoTributario
   'Carga en el Data Combo los Clientes con su RUC
    Leer_Clientes
    DCTipoComprobante.Text = "Factura"
   'Carga la Tabla de Retencion Iva Bienes y Servicios al DataCombo
    Carga_RetencionIvaBienes_Servicios
    DCPorcenIce.Text = ""
   'Carga la Tabla de Conceptos Retencion al DataCombo
    MBFechaRegis = MBFechaEmi
    Carga_ConceptosRetencion MBFechaEmi
    
   'Verifico si existen registros caso contrario despliego mensaje
   'Carga los Conceptos de retención en la Fuente al DataCombo
    sSQL = "SELECT (Codigo & ' - ' & Cuenta) As Cuentas  " _
         & "FROM Catalogo_Cuentas " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TC = 'RF' " _
         & "AND DG = 'D' " _
         & "ORDER BY Codigo "
    SelectDBCombo DCRetFuente, AdoRetFuente, sSQL, "Cuentas"
    If AdoRetFuente.Recordset.RecordCount > 0 Then Rf = 1 Else Rf = 0

   'Carga los Conceptos de retención IVA Servicios al DataCombo
    sSQL = "SELECT (Codigo & ' - ' & Cuenta) As Cuentas  " _
         & "FROM Catalogo_Cuentas " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TC = 'RI' " _
         & "AND DG = 'D' " _
         & "ORDER BY Codigo "
    SelectDBCombo DCRetISer, AdoRetIvaSerCC, sSQL, "Cuentas"
    If AdoRetIvaSerCC.Recordset.RecordCount > 0 Then Rs = 1 Else Rs = 0
    
    'Carga los Conceptos de retención IVA Bienes al DataCombo
    sSQL = "SELECT (Codigo & ' - ' & Cuenta) As Cuentas  " _
         & "FROM Catalogo_Cuentas " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TC = 'RI' " _
         & "AND DG = 'D' " _
         & "ORDER BY Codigo "
    SelectDBCombo DCRetIBienes, AdoRetIvaBienesCC, sSQL, "Cuentas"
    If AdoRetIvaBienesCC.Recordset.RecordCount > 0 Then Rb = 1 Else Rb = 0
     
    'Si es Nuevo ingresa por aqui
    ChRetF.Visible = True
    DCRetFuente.Visible = True
    FrmRetencion.Visible = True
    If EsNuevo Then
       'Si todas las variables tienen cero despliego mensaje y no cargo nada
       'No hay cuentas
       If (Rf And Rs And Rb) = 0 Then
           ChRetF.Visible = False
           DCRetFuente.Visible = False
           FrmRetencion.Visible = False
           Activar_BS
       Else
           ChRetB.SetFocus
       End If
    End If
End Sub

Public Sub Grabacion()
Dim Cta_Prov As String
     Cta_Prov = Leer_Seteos_Ctas("Cta_Proveedores")
     RatonReloj
     Trans_No = 250
     ID_Trans = Trans_No
     NumComp = ReadSetDataNum("Diario", True, True)
     FechaTexto = MBFechaEmi
   
   'Grabo en el Asiento_Compras e implicito Asiento_Air
    If OpcSi.value = True Then ValorP = "S" Else ValorP = "N"
   'MsgBox CodigoCliente
    FechaTexto = MBFechaRegis
    Factura_No = CTNumero(TxtNumSerietres)
    SetAdoAddNew "Asiento_Compras"
    SetAdoFields "IdProv", CodigoCliente
    SetAdoFields "DevIva", ValorP
    SetAdoFields "CodSustento", Cap
    SetAdoFields "TipoComprobante", cod
    SetAdoFields "Establecimiento", TxtNumSerieUno
    SetAdoFields "PuntoEmision", TxtNumSerieDos
    SetAdoFields "Secuencial", Factura_No
    SetAdoFields "Autorizacion", TxtNumAutor
    SetAdoFields "FechaEmision", MBFechaEmi
    SetAdoFields "FechaRegistro", MBFechaRegis
    SetAdoFields "FechaCaducidad", MBFechaCad
    SetAdoFields "BaseImponible", CTNumero(TxtBaseImpo, 2)
    SetAdoFields "BaseImpGrav", CTNumero(TxtBaseImpoGrav, 2)
    SetAdoFields "PorcentajeIva", CodPorIva
    SetAdoFields "MontoIva", CTNumero(TxtMontoIva, 2)
    SetAdoFields "BaseImpIce", CTNumero(TxtBaseImpoIce, 2)
    SetAdoFields "PorcentajeIce", CodPorIce
    SetAdoFields "MontoIce", CTNumero(TxtMontoIce, 2)
    SetAdoFields "Porc_Bienes", DCPorcenRetenIvaBien
    SetAdoFields "MontoIvaBienes", CTNumero(TxtIvaBienMonIva, 2)
    SetAdoFields "PorRetBienes", CodRetBien
    SetAdoFields "ValorRetBienes", CTNumero(TxtIvaBienValRet, 2)
    SetAdoFields "Porc_Servicios", DCPorcenRetenIvaServ
    SetAdoFields "MontoIvaServicios", CTNumero(TxtIvaSerMonIva, 2)
    SetAdoFields "PorRetServicios", CodRetServ
    SetAdoFields "ValorRetServicios", CTNumero(TxtIvaSerValRet, 2)
    If (cod = 4) Or (cod = 5) Then
       SetAdoFields "DocModificado", CapDm
       SetAdoFields "FechaEmiModificado", MBFechaEmiComp
       SetAdoFields "EstabModificado", TxtNumSerieUnoComp
       SetAdoFields "PtoEmiModificado", TxtNumSerieDosComp
       SetAdoFields "SecModificado", CNumSerieTresComp
       SetAdoFields "AutModificado", TxtNumAutComp
    Else
       SetAdoFields "DocModificado", "0"
       SetAdoFields "FechaEmiModificado", date
       SetAdoFields "EstabModificado", "000"
       SetAdoFields "PtoEmiModificado", "000"
       SetAdoFields "SecModificado", "0000000"
       SetAdoFields "AutModificado", "0000000000"
    End If
   'Verifico si activaron los checks de retenciones
    If ChRetB = 1 Then SetAdoFields "Cta_Bienes", SinEspaciosIzq(DCRetIBienes)
    If ChRetS = 1 Then SetAdoFields "Cta_Servicio", SinEspaciosIzq(DCRetISer)
    SetAdoFields "A_No", Ln_No
    SetAdoFields "T_No", Trans_No
    SetAdoUpdate
    
   'Grabamos los datos de la transaccion en la tabla definitiva de almacenamiento
''    ID_Trans = Maximo_De("Trans_Compras", "ID")  'va a tener el indice de transaccion unico para que no exista duplicados en a base
''    sSQL = "SELECT * " _
''         & "FROM Asiento_Compras " _
''         & "WHERE Item = '" & NumEmpresa & "' " _
''         & "AND CodigoU = '" & CodigoUsuario & "' " _
''         & "AND T_No = " & Trans_No & " " _
''         & "ORDER BY T_No "
''    SelectAdodc AdoAsientoCompras, sSQL
    
''   'Selecciona el numero mayor para continuar la secuencia en el
''    ID_Trans = Maximo_De("Trans_Air", "ID")
''    sSQL = "SELECT * " _
''         & "FROM Asiento_Air " _
''         & "WHERE Item = '" & NumEmpresa & "' " _
''         & "AND CodigoU = '" & CodigoUsuario & "' " _
''         & "AND T_No = " & Trans_No & " " _
''         & "AND Tipo_Trans = 'C' " _
''         & "ORDER BY A_No "
''    SelectAdodc AdoTransAir, sSQL
    RatonNormal
   'Realizamos el asiento de la retencion
    
    Total_Ret = 0
    Total_RetIVA = 0
    sSQL = "SELECT * " _
         & "FROM Asiento " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND T_No = " & Trans_No & " "
    SelectAdodc AdoAsientos, sSQL
    OpcTM = 1
    OpcDH = 2
    NoCheque = Ninguno
   'Grabamos el Asiento de la Compra
    sSQL = "SELECT * " _
         & "FROM Asiento_Compras " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND T_No = " & Trans_No & " "
    SelectAdodc AdoAux, sSQL
    With AdoAux.Recordset
     If .RecordCount > 0 Then
        'Porcentaje por Servicio: 0,30,100
         Cta = .Fields("Cta_Servicio")
         DetalleComp = "Retencion del " & .Fields("Porc_Servicios") & "%, Factura No. " & .Fields("Secuencial") & ", de " & NombreCliente
         LeerCta Cta
         ValorDH = .Fields("ValorRetServicios")
         Total_RetIVA = Total_RetIVA + .Fields("ValorRetServicios")
         If ValorDH > 0 Then InsertarAsiento AdoAsientos
        'Porcentaje por Bienes: 0,70,100
         Cta = .Fields("Cta_Bienes")
         DetalleComp = "Retencion del " & .Fields("Porc_Bienes") & "%, Factura No. " & .Fields("Secuencial") & ", de " & NombreCliente
         LeerCta Cta
         ValorDH = .Fields("ValorRetBienes")
         Total_RetIVA = Total_RetIVA + .Fields("ValorRetBienes")
         If ValorDH > 0 Then InsertarAsiento AdoAsientos
     End If
    End With
   'Grabamos el Asiento de las Retenciones
    sSQL = "SELECT * " _
         & "FROM Asiento_Air " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND T_No = " & Trans_No & " " _
         & "AND Tipo_Trans = 'C' " _
         & "ORDER BY Cta_Retencion,A_No,ValRet "
    SelectAdodc AdoAux, sSQL
    With AdoAux.Recordset
     If .RecordCount > 0 Then
         Do While Not .EOF
            Cta = .Fields("Cta_Retencion")
            DetalleComp = "Retencion (" & .Fields("CodRet") & ") No. " & .Fields("SecRetencion") & " del " & (.Fields("Porcentaje") * 100) & "%, de " & NombreCliente
            LeerCta Cta
            ValorDH = .Fields("ValRet")
            Total_Ret = Total_Ret + .Fields("ValRet")
            If ValorDH > 0 Then InsertarAsiento AdoAsientos
           .MoveNext
         Loop
     End If
    End With
    DetalleComp = Ninguno
    OpcDH = 1
    NoCheque = Ninguno
    LeerCta Cta_Prov
    ValorDH = Total_RetIVA + Total_Ret
    InsertarAsiento AdoAsientos
    
    'Grabacion del Comp
     Co.TP = CompDiario
     Co.T = Normal
     Co.Fecha = FechaTexto
     Co.Numero = NumComp
     Co.Monto_Total = 0
     Co.Efectivo = 0
     Co.Concepto = "Registro de Retenciones a: " & DCProveedor _
                 & ", Autorizacion No. " & TxtNumAutor _
                 & ", Documento No. " & TxtNumSerieUno & TxtNumSerieDos & "-" & Factura_No
     Co.CodigoB = CodigoCliente
     Co.Cotizacion = 0
     Co.Item = NumEmpresa
     Co.Usuario = CodigoUsuario
     Co.T_No = Trans_No
    GrabarComprobante Co
    Control_Procesos Normal, "Grabar Comprobante de: " & Co.TP & "No. " & NumComp
    ImprimirComprobantesDe False, Co
End Sub

Public Sub Habilita_Controles()
   'Habilito los controles para la modificacion
'    CModificacion.Enabled = True
    SSTCompras.Enabled = True
    DCProveedor.Enabled = True
    CmdGrabar.Enabled = True
'    FrmTipoComprob.Enabled = True
    FrmRetencion.Enabled = True
    Label23.Visible = True
''    CMes.Visible = True
''    CAño.Visible = True
End Sub

Public Sub Deshabilita_Controles()
   'Deshabilito los controles para la modificacion
''    CModificacion.Enabled = False
    SSTCompras.Enabled = False
    DCProveedor.Enabled = False
    CmdGrabar.Enabled = False
'    FrmTipoComprob.Enabled = False
    FrmRetencion.Enabled = False
    Label23.Visible = False
''    CMes.Visible = False
''    CAño.Visible = False
End Sub

Public Sub Activar_BS()
    TxtIvaBienMonIva.Enabled = True
    DCPorcenRetenIvaBien.Enabled = True
    TxtIvaBienValRet.Enabled = True
    TxtIvaSerMonIva.Enabled = True
    DCPorcenRetenIvaServ.Enabled = True
    TxtIvaSerValRet.Enabled = True
End Sub

Public Sub Listar_Air()
  'Enceramos el espacio de cada usuario para emprezar con una nueva retencion
   sSQL = "DELETE * " _
        & "FROM Asiento_Compras " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND T_No = " & Trans_No & " " _
        & "AND CodigoU = '" & CodigoUsuario & "' "
   ConectarAdoExecute sSQL
  
  'Borra Asiento Air
   sSQL = "DELETE * " _
        & "FROM Asiento_Air " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND T_No = " & Trans_No & " " _
        & "AND Tipo_Trans = 'C' " _
        & "AND CodigoU = '" & CodigoUsuario & "' "
   ConectarAdoExecute sSQL
   
  'Borramos el asiento para empezar un comprobante nuevo
   sSQL = "DELETE * " _
        & "FROM Asiento " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND CodigoU = '" & CodigoUsuario & "' " _
        & "AND T_No = " & Trans_No & " "
   ConectarAdoExecute sSQL
   
  'Presentamos la malla Asiento Air
   sSQL = "SELECT * " _
        & "FROM Asiento_Air " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND CodigoU =  '" & CodigoUsuario & "' " _
        & "AND T_No = " & Trans_No & " " _
        & "AND Tipo_Trans = 'C' " _
        & "ORDER BY CodRet "
   SelectDataGrid DGConceptoAir, AdoAsientoAir, sSQL
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

Public Sub Encerar_Var()
  ac = 0
  Ln_No = 0
  DCPorcenIce = 0
  DCPorcenRetenIvaBien = 0
  DCPorcenRetenIvaServ = 0
  CodPorIce = "0"
  CodPorIva = "0"
  CodRetBien = "0"
  CodRetServ = "0"
End Sub
