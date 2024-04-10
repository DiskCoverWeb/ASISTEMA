VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form FComprasAT 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "COMPRAS"
   ClientHeight    =   8580
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   12060
   ForeColor       =   &H8000000F&
   Icon            =   "FComprasAT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   12060
   Begin VB.CommandButton CmdCerrar 
      BackColor       =   &H00FF8080&
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   10920
      Picture         =   "FComprasAT.frx":0696
      Style           =   1  'Graphical
      TabIndex        =   110
      ToolTipText     =   "Salir"
      Top             =   105
      Width           =   1065
   End
   Begin VB.Frame FrmRetencion 
      BackColor       =   &H00FFC0C0&
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
      TabIndex        =   0
      Top             =   105
      Width           =   9570
      Begin MSDataListLib.DataCombo DCRetISer 
         Bindings        =   "FComprasAT.frx":0AD8
         DataSource      =   "AdoRetIvaSerCC"
         Height          =   315
         Left            =   1260
         TabIndex        =   4
         Top             =   525
         Visible         =   0   'False
         Width           =   8205
         _ExtentX        =   14473
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DCRetIBienes 
         Bindings        =   "FComprasAT.frx":0AF5
         DataSource      =   "AdoRetIvaBienesCC"
         Height          =   315
         Left            =   1260
         TabIndex        =   2
         Top             =   210
         Visible         =   0   'False
         Width           =   8205
         _ExtentX        =   14473
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.CheckBox ChRetB 
         BackColor       =   &H00FFC0C0&
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
         TabIndex        =   1
         Top             =   210
         Width           =   1170
      End
      Begin VB.CheckBox ChRetS 
         BackColor       =   &H00FFC0C0&
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
         TabIndex        =   3
         Top             =   525
         Width           =   1170
      End
   End
   Begin VB.CommandButton CmdGrabar 
      BackColor       =   &H00FF8080&
      Caption         =   "&Aceptar"
      Height          =   855
      Left            =   9765
      Picture         =   "FComprasAT.frx":0B15
      Style           =   1  'Graphical
      TabIndex        =   109
      ToolTipText     =   "Grabar"
      Top             =   105
      Width           =   1065
   End
   Begin TabDlg.SSTab SSTCompras 
      Height          =   6555
      Left            =   105
      TabIndex        =   8
      Top             =   1890
      Width           =   11850
      _ExtentX        =   20902
      _ExtentY        =   11562
      _Version        =   393216
      TabHeight       =   420
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Comprobante de Compra"
      TabPicture(0)   =   "FComprasAT.frx":0E1F
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label15"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "DCSustento"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "OpcSi"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "OpcNo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "FraDctoModificado"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame5"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame4"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "CmdAir"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Conceptos AIR"
      TabPicture(1)   =   "FComprasAT.frx":0E3B
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label44"
      Tab(1).Control(1)=   "Label41"
      Tab(1).Control(2)=   "DCTipoPago"
      Tab(1).Control(3)=   "Frame2"
      Tab(1).Control(4)=   "CFormaPago"
      Tab(1).Control(5)=   "FrmPagoExterior"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Partidos Políticos"
      TabPicture(2)   =   "FComprasAT.frx":0E57
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame8"
      Tab(2).ControlCount=   1
      Begin VB.Frame FrmPagoExterior 
         Height          =   960
         Left            =   -74895
         TabIndex        =   75
         Top             =   735
         Visible         =   0   'False
         Width           =   11565
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'None
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
            TabIndex        =   79
            Top             =   210
            Width           =   1380
            Begin VB.OptionButton OpcSiAplicaDoble 
               Caption         =   "SI"
               Height          =   225
               Left            =   105
               Style           =   1  'Graphical
               TabIndex        =   80
               Top             =   105
               Value           =   -1  'True
               Width           =   540
            End
            Begin VB.OptionButton OpcNoAplicaDoble 
               Caption         =   "NO"
               Height          =   225
               Left            =   735
               Style           =   1  'Graphical
               TabIndex        =   81
               Top             =   105
               Width           =   540
            End
         End
         Begin VB.Frame Frame6 
            BorderStyle     =   0  'None
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
            TabIndex        =   83
            Top             =   525
            Width           =   1380
            Begin VB.OptionButton OpcNoFormaLegal 
               Caption         =   "NO"
               Height          =   225
               Left            =   735
               Style           =   1  'Graphical
               TabIndex        =   85
               Top             =   105
               Width           =   540
            End
            Begin VB.OptionButton OpcSiFormaLegal 
               Caption         =   "SI"
               Height          =   225
               Left            =   105
               Style           =   1  'Graphical
               TabIndex        =   84
               Top             =   105
               Value           =   -1  'True
               Width           =   540
            End
         End
         Begin MSDataListLib.DataCombo DCPais 
            Bindings        =   "FComprasAT.frx":0E73
            DataSource      =   "AdoPais"
            Height          =   315
            Left            =   105
            TabIndex        =   77
            Top             =   525
            Width           =   5370
            _ExtentX        =   9472
            _ExtentY        =   556
            _Version        =   393216
            Text            =   "DataCombo1"
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
         Begin VB.Label Label45 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Pais al que se efectua el pago"
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
            TabIndex        =   76
            Top             =   210
            Width           =   5370
         End
         Begin VB.Label Label42 
            Caption         =   "Aplica convenio de doble tributación?"
            Height          =   225
            Left            =   5880
            TabIndex        =   78
            Top             =   315
            Width           =   4005
         End
         Begin VB.Label Label43 
            Caption         =   "Pago Sujeto a retención en aplicación de la forma legal?"
            Height          =   225
            Left            =   5880
            TabIndex        =   82
            Top             =   630
            Width           =   4005
         End
      End
      Begin VB.ComboBox CFormaPago 
         Height          =   315
         Left            =   -73425
         TabIndex        =   72
         Text            =   "Combo1"
         Top             =   420
         Width           =   1800
      End
      Begin VB.CommandButton CmdAir 
         Caption         =   "&AIR"
         Height          =   765
         Left            =   10920
         Picture         =   "FComprasAT.frx":0E89
         Style           =   1  'Graphical
         TabIndex        =   70
         ToolTipText     =   "Se ubica en la pestaña de Retenciones"
         Top             =   420
         Width           =   765
      End
      Begin VB.Frame Frame8 
         Caption         =   "SOLO PARTIDOS POLITICOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4815
         Left            =   -74895
         TabIndex        =   111
         Top             =   420
         Width           =   10050
         Begin VB.TextBox TxtMonTitGrat 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   336
            Left            =   7035
            TabIndex        =   114
            Text            =   "0.00"
            ToolTipText     =   "Se debe ingresar el valor de la transacción que corresponde al titulo oneroso, es decir, no oneroso para el informante"
            Top             =   2730
            Width           =   1905
         End
         Begin VB.TextBox TxtMonTitOner 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   336
            Left            =   7035
            TabIndex        =   113
            Text            =   "0.00"
            ToolTipText     =   "Se debe ingresar el valor de la transacción que corresponde al titulo oneroso, es decir, no gratuito para el informante"
            Top             =   1995
            Width           =   1905
         End
         Begin VB.TextBox TxtNumConParPol 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   336
            Left            =   7035
            MaxLength       =   10
            TabIndex        =   112
            Text            =   "0000000000"
            ToolTipText     =   $"FComprasAT.frx":13AF
            Top             =   1260
            Width           =   1905
         End
         Begin VB.Label Label39 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " MONTO DEL CONTRATO"
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
            TabIndex        =   117
            Top             =   2730
            Width           =   4635
         End
         Begin VB.Label Label38 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " MONTO TITULO ONEROSO"
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
            TabIndex        =   116
            Top             =   1995
            Width           =   4635
         End
         Begin VB.Label Label37 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " NUMERO DE CONTRATO DEL PARTIDO POLITICO"
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
            TabIndex        =   115
            Top             =   1260
            Width           =   4635
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
         Height          =   1485
         Left            =   105
         TabIndex        =   38
         Top             =   3150
         Width           =   4950
         Begin VB.TextBox TxtMontoIva 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   3465
            TabIndex        =   42
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
            TabIndex        =   46
            Top             =   945
            Width           =   1275
         End
         Begin MSDataListLib.DataCombo DCPorcenIva 
            Bindings        =   "FComprasAT.frx":146A
            DataSource      =   "AdoPorIva"
            Height          =   315
            Left            =   945
            TabIndex        =   40
            ToolTipText     =   $"FComprasAT.frx":1482
            Top             =   420
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DCPorcenIce 
            Bindings        =   "FComprasAT.frx":1514
            DataSource      =   "AdoPorIce"
            Height          =   315
            Left            =   945
            TabIndex        =   44
            ToolTipText     =   $"FComprasAT.frx":152C
            Top             =   945
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
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
            TabIndex        =   39
            Top             =   420
            Width           =   855
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
            TabIndex        =   43
            Top             =   945
            Width           =   855
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
            TabIndex        =   41
            Top             =   420
            Width           =   1485
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
            TabIndex        =   45
            Top             =   945
            Width           =   1485
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
         Left            =   6615
         TabIndex        =   47
         Top             =   3150
         Width           =   5055
         Begin MSDataListLib.DataCombo DCPorcenRetenIvaServ 
            Bindings        =   "FComprasAT.frx":15BD
            DataSource      =   "AdoRetIvaServicios"
            Height          =   315
            Left            =   3360
            TabIndex        =   57
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
            Bindings        =   "FComprasAT.frx":15DE
            DataSource      =   "AdoRetIvaBienes"
            Height          =   315
            Left            =   1680
            TabIndex        =   52
            ToolTipText     =   $"FComprasAT.frx":15FC
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
            Left            =   1680
            MultiLine       =   -1  'True
            TabIndex        =   50
            Text            =   "FComprasAT.frx":1688
            ToolTipText     =   $"FComprasAT.frx":168D
            Top             =   420
            Width           =   1590
         End
         Begin VB.TextBox TxtIvaBienValRet 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   336
            Left            =   1680
            TabIndex        =   54
            Top             =   1050
            Width           =   1590
         End
         Begin VB.TextBox TxtIvaSerMonIva 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   336
            Left            =   3360
            MultiLine       =   -1  'True
            TabIndex        =   56
            Text            =   "FComprasAT.frx":172C
            ToolTipText     =   $"FComprasAT.frx":1731
            Top             =   420
            Width           =   1590
         End
         Begin VB.TextBox TxtIvaSerValRet 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   336
            Left            =   3360
            TabIndex        =   58
            Text            =   " "
            Top             =   1080
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
            Left            =   1680
            TabIndex        =   48
            Top             =   210
            Width           =   1590
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
            Left            =   3360
            TabIndex        =   55
            Top             =   210
            Width           =   1590
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
            TabIndex        =   49
            Top             =   420
            Width           =   1590
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
            TabIndex        =   51
            Top             =   735
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
            TabIndex        =   53
            Top             =   1050
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
         TabIndex        =   59
         Top             =   4725
         Visible         =   0   'False
         Width           =   11565
         Begin VB.ComboBox CNumSerieTresComp 
            DataSource      =   "AdoAux"
            Height          =   315
            Left            =   5775
            TabIndex        =   65
            Top             =   420
            Width           =   1275
         End
         Begin VB.TextBox TxtNumSerieUnoComp 
            Height          =   330
            Left            =   4935
            MaxLength       =   3
            TabIndex        =   63
            Text            =   "001"
            ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al código del establecimiento"
            Top             =   420
            Width           =   435
         End
         Begin VB.TextBox TxtNumSerieDosComp 
            Height          =   336
            Left            =   5355
            MaxLength       =   3
            TabIndex        =   64
            Text            =   "001"
            ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al punto dde emisión"
            Top             =   420
            Width           =   435
         End
         Begin VB.TextBox TxtNumAutComp 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   8190
            MaxLength       =   49
            TabIndex        =   69
            ToolTipText     =   $"FComprasAT.frx":17C7
            Top             =   420
            Width           =   3270
         End
         Begin MSDataListLib.DataCombo DCDctoModif 
            Bindings        =   "FComprasAT.frx":1853
            DataSource      =   "AdoTipoComprobante"
            Height          =   315
            Left            =   105
            TabIndex        =   61
            ToolTipText     =   "Corresponde al tipo de comprobante que ha sido originalmente modificado antre la emisión de una nota de débito o crédito"
            Top             =   420
            Width           =   4845
            _ExtentX        =   8546
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
         End
         Begin MSMask.MaskEdBox MBFechaEmiComp 
            Height          =   330
            Left            =   7035
            TabIndex        =   67
            ToolTipText     =   $"FComprasAT.frx":1874
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
         Begin VB.Label Label27 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Autorización SRI"
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
            TabIndex        =   68
            Top             =   210
            Width           =   3270
         End
         Begin VB.Label Label26 
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
            Left            =   7035
            TabIndex        =   66
            Top             =   210
            Width           =   1170
         End
         Begin VB.Label Label2 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Serie      Numero"
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
            Left            =   4935
            TabIndex        =   62
            Top             =   210
            Width           =   2115
         End
         Begin VB.Label Label1 
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
            TabIndex        =   60
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
         Height          =   1590
         Left            =   105
         TabIndex        =   14
         Top             =   1470
         Width           =   11565
         Begin VB.TextBox TxtBaseImpoNoObjIVA 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   3990
            MaxLength       =   14
            MultiLine       =   -1  'True
            TabIndex        =   31
            Text            =   "FComprasAT.frx":1920
            ToolTipText     =   "En este campo se debe ingresar el valor del comprobante cuya base imponible esta gravado con la tarifa del 0% de IVA"
            Top             =   1155
            Width           =   1380
         End
         Begin VB.TextBox TxtNumAutor 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   6930
            MaxLength       =   49
            TabIndex        =   23
            Text            =   "0000000000000000000000000000000000000000000000001"
            Top             =   540
            Width           =   4560
         End
         Begin MSMask.MaskEdBox MBFechaCad 
            Height          =   330
            Left            =   2625
            TabIndex        =   29
            ToolTipText     =   $"FComprasAT.frx":1927
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "0"
         End
         Begin MSMask.MaskEdBox MBFechaRegis 
            Height          =   330
            Left            =   1365
            TabIndex        =   27
            ToolTipText     =   $"FComprasAT.frx":19DE
            Top             =   1155
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
         Begin VB.TextBox TxtBaseImpo 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   5460
            MaxLength       =   14
            MultiLine       =   -1  'True
            TabIndex        =   33
            Text            =   "FComprasAT.frx":1A66
            ToolTipText     =   "En este campo se debe ingresar el valor del comprobante cuya base imponible esta gravado con la tarifa del 0% de IVA"
            Top             =   1155
            Width           =   1485
         End
         Begin VB.TextBox TxtNumSerietres 
            Height          =   336
            Left            =   5985
            MaxLength       =   9
            TabIndex        =   21
            Text            =   "0000001"
            ToolTipText     =   $"FComprasAT.frx":1A6D
            Top             =   525
            Width           =   960
         End
         Begin VB.TextBox TxtNumSerieDos 
            Height          =   336
            Left            =   5565
            MaxLength       =   3
            TabIndex        =   19
            Text            =   "001"
            ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al punto dde emisión"
            Top             =   525
            Width           =   435
         End
         Begin VB.TextBox TxtNumSerieUno 
            Height          =   336
            Left            =   5145
            MaxLength       =   3
            TabIndex        =   18
            Text            =   "001"
            ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al código del establecimiento"
            Top             =   525
            Width           =   435
         End
         Begin VB.TextBox TxtBaseImpoGrav 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   7035
            MultiLine       =   -1  'True
            TabIndex        =   35
            Text            =   "FComprasAT.frx":1B10
            ToolTipText     =   $"FComprasAT.frx":1B17
            Top             =   1155
            Width           =   1485
         End
         Begin VB.TextBox TxtBaseImpoIce 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   8610
            MultiLine       =   -1  'True
            TabIndex        =   37
            Text            =   "FComprasAT.frx":1BDD
            ToolTipText     =   $"FComprasAT.frx":1BE2
            Top             =   1155
            Width           =   1380
         End
         Begin MSDataListLib.DataCombo DCTipoComprobante 
            Bindings        =   "FComprasAT.frx":1C74
            DataSource      =   "AdoTipoComp"
            Height          =   315
            Left            =   105
            TabIndex        =   16
            ToolTipText     =   $"FComprasAT.frx":1C95
            Top             =   525
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
         Begin MSMask.MaskEdBox MBFechaEmi 
            Height          =   330
            Left            =   105
            TabIndex        =   25
            ToolTipText     =   $"FComprasAT.frx":1D3D
            Top             =   1155
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
         Begin VB.Label Label40 
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
            Height          =   225
            Left            =   3990
            TabIndex        =   30
            Top             =   945
            Width           =   1380
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
            TabIndex        =   15
            Top             =   315
            Width           =   5055
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
            Height          =   225
            Left            =   5145
            TabIndex        =   17
            Top             =   315
            Width           =   855
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
            Left            =   5985
            TabIndex        =   20
            Top             =   315
            Width           =   960
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
            Height          =   225
            Left            =   6930
            TabIndex        =   22
            Top             =   315
            Width           =   4530
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
            Height          =   225
            Left            =   5460
            TabIndex        =   32
            Top             =   945
            Width           =   1485
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
            Height          =   225
            Left            =   7035
            TabIndex        =   34
            Top             =   945
            Width           =   1485
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
            Height          =   225
            Left            =   8610
            TabIndex        =   36
            Top             =   945
            Width           =   1380
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
            Height          =   225
            Left            =   105
            TabIndex        =   24
            Top             =   945
            Width           =   1170
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
            Height          =   225
            Left            =   1365
            TabIndex        =   26
            Top             =   945
            Width           =   1170
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
            Height          =   225
            Left            =   2625
            TabIndex        =   28
            Top             =   945
            Width           =   1275
         End
      End
      Begin VB.OptionButton OpcNo 
         BackColor       =   &H00E0E0E0&
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
         Left            =   3465
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   $"FComprasAT.frx":1DE9
         Top             =   420
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
         Height          =   330
         Left            =   2730
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   $"FComprasAT.frx":1E81
         Top             =   420
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
         Height          =   4320
         Left            =   -74892
         TabIndex        =   86
         Top             =   1785
         Width           =   11565
         Begin MSDataListLib.DataCombo DCRetFuente 
            Bindings        =   "FComprasAT.frx":1F19
            DataSource      =   "AdoRetFuente"
            Height          =   315
            Left            =   2415
            TabIndex        =   88
            Top             =   315
            Visible         =   0   'False
            Width           =   9045
            _ExtentX        =   15954
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
            TabIndex        =   87
            Top             =   315
            Visible         =   0   'False
            Width           =   2328
         End
         Begin VB.TextBox TxtValConA 
            Enabled         =   0   'False
            Height          =   336
            Left            =   10185
            TabIndex        =   105
            Top             =   1995
            Width           =   1275
         End
         Begin VB.TextBox TxtPorRetConA 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   336
            Left            =   9555
            TabIndex        =   103
            Top             =   1995
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
            Left            =   9870
            TabIndex        =   108
            Text            =   "0.00"
            ToolTipText     =   "Sumatoria total de las retenciones"
            Top             =   3885
            Width           =   1590
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
            Left            =   9555
            MultiLine       =   -1  'True
            TabIndex        =   97
            Text            =   "FComprasAT.frx":1F34
            Top             =   735
            Width           =   1905
         End
         Begin VB.TextBox TxtBimpConA 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   8190
            TabIndex        =   101
            Top             =   1995
            Width           =   1380
         End
         Begin VB.TextBox TxtNumUnoAutComRet 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   2415
            MaxLength       =   49
            TabIndex        =   95
            ToolTipText     =   $"FComprasAT.frx":1F3B
            Top             =   840
            Width           =   5580
         End
         Begin VB.TextBox TxtNumTresComRet 
            Height          =   336
            Left            =   1260
            MaxLength       =   9
            TabIndex        =   93
            Text            =   "0000001"
            ToolTipText     =   $"FComprasAT.frx":1FC7
            Top             =   840
            Width           =   1065
         End
         Begin VB.TextBox TxtNumDosComRet 
            Height          =   336
            Left            =   630
            MaxLength       =   3
            TabIndex        =   91
            Text            =   "001"
            ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al punto dde emisión"
            Top             =   840
            Width           =   540
         End
         Begin VB.TextBox TxtNumUnoComRet 
            Height          =   336
            Left            =   105
            MaxLength       =   3
            TabIndex        =   90
            Text            =   "001"
            ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al código del establecimiento"
            Top             =   840
            Width           =   540
         End
         Begin MSDataListLib.DataCombo DCConceptoRet 
            Bindings        =   "FComprasAT.frx":2069
            DataSource      =   "AdoConceptoRet"
            Height          =   345
            Left            =   1470
            TabIndex        =   99
            Top             =   1260
            Width           =   9990
            _ExtentX        =   17621
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
         Begin MSDataGridLib.DataGrid DGConceptoAir 
            Bindings        =   "FComprasAT.frx":2086
            Height          =   1380
            Left            =   105
            TabIndex        =   106
            Top             =   2415
            Width           =   11355
            _ExtentX        =   20029
            _ExtentY        =   2434
            _Version        =   393216
            AllowUpdate     =   0   'False
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
         Begin VB.Label LblCodRet 
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Codigo Retencion"
            ForeColor       =   &H00FFFFFF&
            Height          =   645
            Left            =   105
            TabIndex        =   120
            Top             =   1680
            Width           =   7995
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
            Left            =   8085
            TabIndex        =   107
            Top             =   3885
            Width           =   1800
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
            Height          =   330
            Left            =   10185
            TabIndex        =   104
            Top             =   1680
            Width           =   1275
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
            Height          =   330
            Left            =   9555
            TabIndex        =   102
            Top             =   1680
            Width           =   645
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
            Height          =   330
            Left            =   8190
            TabIndex        =   100
            Top             =   1680
            Width           =   1380
         End
         Begin VB.Label Label32 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " CODIGO RET."
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
            Top             =   1260
            Width           =   1380
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
            Left            =   8190
            TabIndex        =   96
            Top             =   735
            Width           =   1380
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
            TabIndex        =   94
            Top             =   630
            Width           =   5580
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
            TabIndex        =   92
            Top             =   630
            Width           =   1065
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
            TabIndex        =   89
            Top             =   630
            Width           =   1065
         End
      End
      Begin MSDataListLib.DataCombo DCSustento 
         Bindings        =   "FComprasAT.frx":20A2
         DataSource      =   "AdoSustento"
         Height          =   360
         Left            =   105
         TabIndex        =   13
         ToolTipText     =   "En este campo de selección se despliega un lista de tipos de sustentos tributarios correspondientes a la transacción escogida"
         Top             =   1050
         Width           =   10725
         _ExtentX        =   18918
         _ExtentY        =   635
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DCTipoPago 
         Bindings        =   "FComprasAT.frx":20BC
         DataSource      =   "AdoTipoPago"
         Height          =   315
         Left            =   -69960
         TabIndex        =   74
         Top             =   420
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
      Begin VB.Label Label41 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TIPO DE PAGO"
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
         Left            =   -71535
         TabIndex        =   73
         Top             =   420
         Width           =   1590
      End
      Begin VB.Label Label44 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Forma de Pago"
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
         Left            =   -74895
         TabIndex        =   71
         Top             =   420
         Width           =   1485
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
         TabIndex        =   9
         Top             =   420
         Width           =   2220
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
         TabIndex        =   12
         Top             =   840
         Width           =   10725
      End
   End
   Begin MSAdodcLib.Adodc AdoSustento 
      Height          =   330
      Left            =   2730
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
   Begin MSAdodcLib.Adodc AdoTipoComprobante 
      Height          =   330
      Left            =   210
      Top             =   3045
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
      Top             =   3045
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
      Left            =   210
      Top             =   3990
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
      Left            =   210
      Top             =   4305
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
   Begin MSAdodcLib.Adodc AdoAsientoAir 
      Height          =   330
      Left            =   210
      Top             =   4935
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
   Begin MSAdodcLib.Adodc AdoRetIvaSerCC 
      Height          =   330
      Left            =   2730
      Top             =   4305
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
      Left            =   2730
      Top             =   4620
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
      Left            =   210
      Top             =   5565
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
      Top             =   5880
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   2730
      Top             =   3990
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
   Begin MSAdodcLib.Adodc AdoAsientos 
      Height          =   330
      Left            =   210
      Top             =   6195
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
   Begin MSAdodcLib.Adodc AdoPais 
      Height          =   330
      Left            =   210
      Top             =   6510
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
      Caption         =   "Pais"
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
      Left            =   210
      Top             =   6825
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
   Begin MSDataListLib.DataCombo DCProveedor 
      Bindings        =   "FComprasAT.frx":20D6
      DataSource      =   "AdoClientes"
      Height          =   315
      Left            =   105
      TabIndex        =   118
      ToolTipText     =   "En este combo de selección se despliega una lista de todos los proveedores"
      Top             =   1470
      Width           =   9780
      _ExtentX        =   17251
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
   Begin VB.Label Label23 
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " ESTADO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   330
      Left            =   9660
      TabIndex        =   119
      Top             =   1155
      Width           =   2325
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
      Left            =   10185
      TabIndex        =   7
      Top             =   1470
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
      Left            =   9870
      TabIndex        =   6
      Top             =   1470
      Width           =   330
   End
   Begin VB.Label Label25 
      BackColor       =   &H00000080&
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
      ForeColor       =   &H00C0FFC0&
      Height          =   330
      Left            =   105
      TabIndex        =   5
      Top             =   1155
      Width           =   9570
   End
End
Attribute VB_Name = "FComprasAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MBFecha As MaskEdBox
Dim FechaRegis As Date
Dim OP As Boolean
Dim cod, X, Rb, Rf, rs, CapDm As Byte
Dim SumAnio, Aniocad, AniocadAux, CodPorIva, CodRetBien, CodRetServ As Integer
Dim CalmIva, CalmIce, CalIbMi, CalIsMi, ac, SUM, cal As Double
Dim CapDcto, Cap, Cap1, Ct, ValorP, AuxCodUs, Opc, conta, ch, Ch1, CodSus, Bien, Serv, CargaTC, CodPorIce As String
Dim Espizq, Espder, Captc, PorIce, PorIva, CodProv, CodProv1, NumCed, Mifecha  As String
Dim FechaPorcIVA As String

Private Sub CFormaPago_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub CFormaPago_LostFocus()
 If CFormaPago.Text = "Exterior" Then
    sSQL = "SELECT * " _
         & "FROM Tabla_Naciones " _
         & "WHERE TR = 'N' " _
         & "ORDER BY Descripcion_Rubro "
    SelectDB_Combo DCPais, AdoPais, sSQL, "Descripcion_Rubro"
    FrmPagoExterior.Visible = True
 Else
    FrmPagoExterior.Visible = False
 End If
End Sub

Private Sub ChRetB_Click()
    If ChRetB.value <> 0 Then
       ch = 1
       Ch1 = "B"
       DCRetIBienes.Visible = True
       TxtIvaBienMonIva.Enabled = True
       DCPorcenRetenIvaBien.Enabled = True
       TxtIvaBienValRet.Enabled = True
    Else
       DCRetIBienes.Visible = False
       TxtIvaBienMonIva.Enabled = False
       DCPorcenRetenIvaBien.Enabled = False
       TxtIvaBienValRet.Enabled = False
       ch = 1
       Ch1 = "S"
    End If
    If ChRetB.value <> 0 And ChRetS.value <> 0 Then
       Ch1 = "X"
    End If
End Sub

Private Sub ChRetF_Click()
  If ChRetF.value = 0 Then DCRetFuente.Visible = False Else DCRetFuente.Visible = True
End Sub

Private Sub ChRetS_Click()
    If ChRetS.value <> 0 Then
       ch = 1
       Ch1 = "S"
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
    End If
End Sub

Private Sub CmdAir_Click()
  'Carga los conceptos de Retencion segun la fecha de Registro
   If Len(MBFechaEmi) < 10 Then MBFechaEmi = FechaSistema
   sSQL = "SELECT (Codigo & ' - ' & Concepto) As Detalle_Conceptos,* " _
       & "FROM Tipo_Concepto_Retencion " _
       & "WHERE Codigo <> '.' " _
       & "AND Fecha_Inicio <= #" & BuscarFecha(MBFechaEmi) & "# " _
       & "AND Fecha_Final >= #" & BuscarFecha(MBFechaEmi) & "# " _
       & "ORDER BY Codigo "
   SelectDB_Combo DCConceptoRet, AdoConceptoRet, sSQL, "Detalle_Conceptos"
  'DCConceptoRet = "329 - Por Otros Servicios (N)"
    
    sSQL = "SELECT (Codigo & ' ' & Descripcion) As CTipoPago " _
         & "FROM Tabla_Referenciales_SRI " _
         & "WHERE Tipo_Referencia = 'FORMA DE PAGO' " _
         & "ORDER BY Codigo "
    SelectDB_Combo DCTipoPago, AdoTipoPago, sSQL, "CTipoPago"
    SSTCompras.Tab = 1
    CFormaPago.SetFocus   ' TxtNumUnoComRet
End Sub

Private Sub CmdCerrar_Click()
    Total_Ret = 0
    Total_RetIVA = 0
   'Borra Asiento Compras
    sSQL = "DELETE * " _
         & "FROM Asiento_Compras " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND T_No = " & Trans_No & " "
    Ejecutar_SQL_SP sSQL
   'Borra Asiento Air
    sSQL = "DELETE * " _
         & "FROM Asiento_Air " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND Tipo_Trans = 'C' " _
         & "AND T_No = " & Trans_No & " "
    Ejecutar_SQL_SP sSQL
'''    DesConectar_Adodc AdoAux
'''    DesConectar_Adodc AdoPais
'''    DesConectar_Adodc AdoTipoPago
'''    DesConectar_Adodc AdoSustento
'''    DesConectar_Adodc AdoRetIvaBienes
'''    DesConectar_Adodc AdoRetIvaServicios
'''    DesConectar_Adodc AdoPorIce
'''    DesConectar_Adodc AdoPorIva
'''    DesConectar_Adodc AdoConceptoRet
'''    DesConectar_Adodc AdoAsientoAir
'''    DesConectar_Adodc AdoRetFuente
'''    DesConectar_Adodc AdoRetIvaSerCC
'''    DesConectar_Adodc AdoRetIvaBienesCC
    Unload FComprasAT
End Sub

Private Sub CmdGrabar_Click()
   'Valido por si acaso exista algun valor con 0
    TextoValido TxtIvaBienMonIva, True, , 2
    TextoValido TxtBaseImpo, True, , 2
    TextoValido TxtBaseImpoGrav, True, , 2
    TextoValido TxtBaseImpoNoObjIVA, True, , 2
    TextoValido TxtBaseImpoIce, True, , 2
    TextoValido TxtMontoIva, True, , 2
    TextoValido TxtMontoIce, True, , 2
    TextoValido TxtIvaBienMonIva, True, , 2
    TextoValido TxtIvaBienValRet, True, , 2
    TextoValido TxtIvaSerMonIva, True, , 2
    TextoValido TxtIvaSerValRet, True, , 2
    Grabacion
    Total_Ret = 0
    Total_RetIVA = 0
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
    If Leer_Campo_Empresa("Registrar_IVA") Then
       Cta = Cta_IVA_Inventario
       DetalleComp = "Registro del IVA en compras Doc. No. " & TxtNumSerieUno & TxtNumSerieDos & "-" & TxtNumSerietres & ", " & NombreCliente
       Codigo = Leer_Cta_Catalogo(Cta)
       ValorDH = Redondear(CCur(TxtMontoIva), 2)
       If ValorDH > 0 Then InsertarAsiento AdoAsientos
    End If
    OpcDH = 2
    sSQL = "SELECT * " _
         & "FROM Asiento_Compras " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND T_No = " & Trans_No & " "
    Select_Adodc AdoAux, sSQL
    With AdoAux.Recordset
     If .RecordCount > 0 Then
        'Porcentaje por Servicio: 0,30,100
         Cta = .fields("Cta_Servicio")
         DetalleComp = "Retencion del " & .fields("Porc_Servicios") & "%, Factura No. " & .fields("Secuencial") & ", de " & NombreCliente
         Codigo = Leer_Cta_Catalogo(Cta)
         ValorDH = .fields("ValorRetServicios")
         Total_RetIVA = Total_RetIVA + .fields("ValorRetServicios")
         If ValorDH > 0 Then InsertarAsiento AdoAsientos
        'Porcentaje por Bienes: 0,70,100
         Cta = .fields("Cta_Bienes")
         DetalleComp = "Retencion del " & .fields("Porc_Bienes") & "%, Factura No. " & .fields("Secuencial") & ", de " & NombreCliente
         Codigo = Leer_Cta_Catalogo(Cta)
         ValorDH = .fields("ValorRetBienes")
         Total_RetIVA = Total_RetIVA + .fields("ValorRetBienes")
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
    Select_Adodc AdoAux, sSQL
    With AdoAux.Recordset
     If .RecordCount > 0 Then
         Do While Not .EOF
            Cta = .fields("Cta_Retencion")
            DetalleComp = "Retencion (" & .fields("CodRet") & ") No. " & .fields("SecRetencion") & " del " & (.fields("Porcentaje") * 100) & "%, de " & NombreCliente
            Codigo = Leer_Cta_Catalogo(Cta)
            ValorDH = .fields("ValRet")
            Total_Ret = Total_Ret + .fields("ValRet")
            If ValorDH > 0 Then InsertarAsiento AdoAsientos
           .MoveNext
         Loop
     End If
    End With
    DetalleComp = Ninguno
    
'''    DesConectar_Adodc AdoAux
'''    DesConectar_Adodc AdoPais
'''    DesConectar_Adodc AdoTipoPago
'''    DesConectar_Adodc AdoSustento
'''    DesConectar_Adodc AdoTipoComprobante
'''    DesConectar_Adodc AdoRetIvaBienes
'''    DesConectar_Adodc AdoRetIvaServicios
'''    DesConectar_Adodc AdoPorIce
'''    DesConectar_Adodc AdoPorIva
'''    DesConectar_Adodc AdoConceptoRet
'''    ''DesConectar_Adodc AdoAsientoCompras
'''    DesConectar_Adodc AdoAsientoAir
'''    'DesConectar_Adodc AdoTransAir
'''    DesConectar_Adodc AdoAsientos
'''    DesConectar_Adodc AdoRetFuente
'''    DesConectar_Adodc AdoRetIvaSerCC
'''    DesConectar_Adodc AdoRetIvaBienesCC
    Unload FComprasAT
End Sub

Private Sub CNumSerieTresComp_LostFocus()
  Factura_CxP = Val(CNumSerieTresComp)
End Sub

Private Sub DCConceptoRet_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCConceptoRet_LostFocus()
    LblCodRet.Caption = ""
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
               LblCodRet.Caption = .fields("Codigo") & ": " & .fields("Concepto")
               TxtPorRetConA = .fields("Porcentaje")
               If .fields("Ingresar_Porcentaje") = "S" Then OP = True
              'MsgBox .Fields("Porcentaje")
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

Private Sub DCDctoModif_LostFocus()
    Captura_TipoComprobante_DctoModificado
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
       PorIva = DCPorcenIva.Text
       Label13.Caption = " TARIFA " & DCPorcenIva.Text & "%"
       CodPorIva = "0"
       With AdoPorIva.Recordset
        If .RecordCount > 0 Then
           .MoveFirst
           .Find ("Porc = '" & PorIva & "' ")
            If Not .EOF Then CodPorIva = .fields("Codigo")
        End If
       End With

       Total_IVA = 0
       Total_IVA = CTNumero(TxtBaseImpoGrav, 2)
      'Calcula el Porcentaje de Iva
       CalmIva = (Total_IVA * DCPorcenIva) / 100
       TxtMontoIva = CalmIva
    End If
End Sub

Private Sub DCRetFuente_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
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
      'Ultima Factura del Proveedor: CodigoCliente
       sSQL = "SELECT TOP 1 Secuencial, FechaCaducidad, Establecimiento, PuntoEmision, Autorizacion " _
            & "FROM Trans_Compras " _
            & "WHERE IdProv = '" & CodigoCliente & "' " _
            & "AND Item = '" & NumEmpresa & "' " _
            & "ORDER BY Fecha DESC,Secuencial DESC "
       Select_Adodc AdoAux, sSQL
       If AdoAux.Recordset.RecordCount > 0 Then
          TxtNumSerieUno = AdoAux.Recordset.fields("Establecimiento")
          TxtNumSerieDos = AdoAux.Recordset.fields("PuntoEmision")
          If Val(TxtNumSerietres) <= 0 Then TxtNumSerietres = AdoAux.Recordset.fields("Secuencial") + 1
          If Len(AdoAux.Recordset.fields("Autorizacion")) < 13 Then
             TxtNumAutor = AdoAux.Recordset.fields("Autorizacion")
             MBFechaCad = AdoAux.Recordset.fields("FechaCaducidad")
          End If
       Else
          TxtNumAutor = Autorizacion
       End If
    End If
End Sub

Private Sub DCTipoPago_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCTipoPago_LostFocus()
  If CFormaPago.Text = "Exterior" Then DCPais.SetFocus
End Sub

Private Sub DGConceptoAir_KeyDown(KeyCode As Integer, Shift As Integer)
 Keys_Especiales Shift
 If CtrlDown And KeyCode = vbKeyDelete Then
   Titulo = "Aviso"
   Mensajes = "Desea Eliminar la Retención"
   If BoxMensaje = vbYes Then
      With AdoAsientoAir.Recordset
       If .RecordCount > 0 Then
           Codigo = .fields("CodRet")
           No_Desde = .fields("SecRetencion")
           Mifecha = BuscarFecha(.fields("FechaEmiRet"))
           Codigo1 = .fields("AutRetencion")
           J = .fields("A_No")
           sSQL = "DELETE * " _
                & "FROM Asiento_Air " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND CodigoU = '" & CodigoUsuario & "' " _
                & "AND IdProv = '" & CodigoCliente & "' " _
                & "AND T_No = " & Trans_No & " " _
                & "AND Tipo_Trans = 'C' " _
                & "AND A_No = " & J & " " _
                & "AND CodRet = '" & Codigo & "' "
           Ejecutar_SQL_SP sSQL
       End If
       AdoAsientoAir.Refresh
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
    FechaPorcIVA = BuscarFecha(MBFechaEmiComp)
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
        Case 0: If ChRetF.Visible Then ChRetF.SetFocus
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
Dim PorcIva As Byte
 Keys_Especiales Shift
 If CtrlDown And KeyCode = vbKeyF12 Then
    PorcIva = InputBox("Ingrese el porcentaje a Proccesar:", "PORCENTAJE DE IVA", Porc_IVA * 100)
    Select Case PorcIva
      Case 8
           Porc_IVA = PorcIva / 100
           FechaPorcIVA = BuscarFecha("01/06/2017")
      Case 12
           Porc_IVA = PorcIva / 100
           FechaPorcIVA = BuscarFecha("01/06/2017")
      Case 14
           Porc_IVA = PorcIva / 100
           FechaPorcIVA = BuscarFecha("01/06/2016")
      Case Else
           Porc_IVA = 0.12
           FechaPorcIVA = BuscarFecha("01/06/2017")
    End Select
    Label13.Caption = " TARIFA " & (Porc_IVA * 100) & "%"
 End If
 PresionoEnter KeyCode
End Sub

Private Sub TxtBaseImpoGrav_LostFocus()
    TextoValido TxtBaseImpoGrav, True, , 2
   'Carga la Tabla de Porcentaje Iva
''    sSQL = "SELECT * " _
''         & "FROM Tabla_Por_ICE_IVA " _
''         & "WHERE IVA <> " & Val(adFalse) & " " _
''         & "AND Fecha_Inicio <= #" & FechaPorcIVA & "# " _
''         & "AND Fecha_Final >= #" & FechaPorcIVA & "# " _
''         & "ORDER BY Porc DESC "
''    SelectDB_Combo DCPorcenIva, AdoPorIva, sSQL, "Porc"
    DCPorcenIva = Porc_IVA * 100
End Sub

Private Sub TxtBaseImpoIce_GotFocus()
    MarcarTexto TxtBaseImpoIce
End Sub

Private Sub TxtBaseImpoIce_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtBaseImpoIce_LostFocus()
    TextoValido TxtBaseImpoIce, True, , 2
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
    TextoValido TxtBimpConA, True, , 2
    TextoValido TxtSumatoria, True, , 2
    RatonNormal
   'Valida que la base imponible no sea mayor que la BIG y la BIcero
    If CTNumero(TxtBimpConA, 2) > CTNumero(TxtSumatoria, 2) Then
       MsgBox "La Base Imponible debe ser menor o igual a la " & vbCrLf _
       & "Base Imponible Gravada + la Base Imponible 0%", vbInformation, "Aviso"
       TxtBimpConA.Text = 0
       TxtBimpConA.SetFocus
    Else
       If Not OP Then
          TxtValConA = CTNumero(TxtBimpConA, 2) * (CTNumero(TxtPorRetConA, 2) / 100)
          Insertar_DataGrid
          If (cod = 4) Or (cod = 5) Then
             DCDctoModif.SetFocus
          Else
             TxtNumConParPol.SetFocus
          End If
       End If
    End If
End Sub

Sub Insertar_DataGrid()
    'Selecciona el numero mayor para continuar la secuencia en el
    'campo T_No y A_No
    If TxtBimpConA = "" Then TxtBimpConA = "0"
    If Val(CCur(TxtBimpConA)) > 0 Then
       RatonReloj
       Espizq = SinEspaciosIzq(DCConceptoRet)
       Espder = TrimStrg(MidStrg(DCConceptoRet, Len(Espizq) + 3, Len(DCConceptoRet)))
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
       SetAdoFields "A_No", Maximo_De("Asiento_Air", "A_No")
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
       Select_Adodc_Grid DGConceptoAir, AdoAsientoAir, sSQL, "Sustento"
         
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
       PorIce = (DCPorcenIce.Text)
       With AdoPorIce.Recordset
            If .RecordCount > 0 Then
               .MoveFirst
               .Find ("Porc = '" & PorIce & "' ")
               If Not .EOF Then
                  CodPorIce = .fields("Codigo")
               Else
                  'MsgBox "Vuelva a seleccionar", vbInformation, "Aviso"
               End If
             End If
       End With
        
       Total_IVA = 0
       Total_IVA = CTNumero(TxtBaseImpoIce, 2)
       TxtMontoIce = 0
      'Calcula el Porcentaje de Ice
       CalIbMi = (Total_IVA * DCPorcenIce) / 100
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
            If Not .EOF Then CodRetBien = .fields("Codigo")
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
         If Not .EOF Then CodRetServ = .fields("Codigo")
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
       MsgBox "No ingrese números. Vuelva a seleccionar.", vbInformation, "Aviso"
       DCSustento.Text = ""
       Carga_CreditoTributario
       DCSustento.SetFocus
       Carga_TipoComprobante (SinEspaciosIzq(DCSustento))
    Else
       Carga_TipoComprobante (SinEspaciosIzq(DCSustento))
    End If
End Sub

Private Sub Form_Activate()
   Select Case Co.TD
     Case "C", "R", "P"
          Label23.Caption = " Estado: ACTIVO"
          Label25.Caption = " TIPO DE PROVEEDOR: "
     Case Else
          MsgBox "Este Beneficiario no es valido para esta operacion"
          Unload FComprasAT
   End Select
   If Len(TipoContribuyente) > 1 Then Label25.Caption = Label25.Caption & TipoContribuyente
   If Len(Co.MicroEmpresa) > 1 Then Label25.Caption = Label25.Caption & " " & Co.MicroEmpresa
   If Len(Co.AgenteRetencion) > 1 Then Label25.Caption = Label25.Caption & " " & Co.AgenteRetencion
   If Len(Co.Estado) > 1 Then Label23.Caption = " Estado: " & Co.Estado
   MBFechaEmi = FechaComp
   Co.Serie_R = Ninguno
   Co.Retencion = 0
   Co.Autorizacion_R = Ninguno
   Co.RetNueva = True
   Co.Autorizacion_R = Ninguno
   Factura_CxP = 0
   TxtBaseImpo = "0.00"
   TxtBaseImpoIce = "0.00"
   TxtBaseImpoGrav = "0.00"
   TxtBaseImpoNoObjIVA = "0.00"
   Carga_Datos_Iniciales MBFechaEmi, Nuevo
   
   LblTD.Caption = TipoBenef                  ' Tipo de Cliente: C,R,P,O
   LblNumIdent = CICliente                    ' CI o RUC del Cliente
   DCProveedor.Text = NombreCliente           ' Nombre del Cliente
   MBFechaEmi = FechaComp
   MBFechaRegis = FechaComp
   MBFechaCad = FechaComp
   TxtNumSerietres = "0000001"
   TxtNumSerieUno = "001"
   TxtNumSerieDos = "001"
   TxtNumAutor = String(10, "0")
   TxtNumUnoComRet = "001"
   TxtNumDosComRet = "001"
   TxtNumTresComRet = "0000001"
   TxtNumUnoAutComRet = String(10, "0")
   
  'Ultima Retencion Emitida
   TxtNumUnoComRet = "001"
   TxtNumDosComRet = "001"
   TxtNumTresComRet = 1
   TxtNumUnoAutComRet = "1234567890"
   'MsgBox sSQL & vbCrLf & AdoAux.Recordset.Fields("SecRetencion")
End Sub

Private Sub Form_Load()
   CentrarForm FComprasAT
   ConectarAdodc AdoAux
   ConectarAdodc AdoPais
   ConectarAdodc AdoTipoPago
   ConectarAdodc AdoSustento
   ConectarAdodc AdoTipoComprobante
   ConectarAdodc AdoRetIvaBienes
   ConectarAdodc AdoRetIvaServicios
   ConectarAdodc AdoPorIce
   ConectarAdodc AdoPorIva
   ConectarAdodc AdoConceptoRet
   ConectarAdodc AdoAsientoCompras
   ConectarAdodc AdoAsientoAir
   ConectarAdodc AdoAsientos
   ConectarAdodc AdoRetFuente
   ConectarAdodc AdoRetIvaSerCC
   ConectarAdodc AdoRetIvaBienesCC
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
End Sub

Private Sub MBFechaEmi_LostFocus()
    FechaValida MBFechaEmi
    Validar_Porc_IVA MBFechaEmi
    Label13.Caption = " TARIFA " & (Porc_IVA * 100) & "%"
   'Controla que la Fecha de Emisiòn este entre 01/01/2000 en adelante
    If CFechaLong(MBFechaEmi) < CFechaLong("01/01/2000") Then
       MsgBox "La Fecha de Emisión debe ser mayor que 01/01/2000", vbInformation, "Aviso"
       MBFechaEmi = "01/01/2000"
       MBFechaEmi.SetFocus
    End If
    MBFechaRegis = MBFechaEmi
    MBFechaEmiComp = MBFechaEmi
End Sub

Private Sub MBFechaRegis_GotFocus()
    MarcarTexto MBFechaRegis
End Sub

Private Sub MBFechaRegis_KeyDown(KeyCode As Integer, Shift As Integer)
 PresionoEnter KeyCode
End Sub

Private Sub MBFechaRegis_LostFocus()
   FechaValida MBFechaRegis
   'Controla que la Fecha de Registro este entre 01/01/2000 en adelante
   If CFechaLong(MBFechaRegis) < CFechaLong("01/01/2000") Then
      MsgBox "La Fecha de Registro debe ser mayor que 01/01/2000", vbInformation, "Aviso"
      MBFechaRegis = "01/01/2000"
      MBFechaRegis.SetFocus
   Else
      If MBFechaRegis < MBFechaEmi Then
         MsgBox "La Fecha de Registro debe ser mayor o igual que la Fecha de Emisión", vbInformation, "Aviso"
         MBFechaRegis.SetFocus
      End If
   End If
   FechaValida MBFechaRegis
 ' Carga la Tabla de Porcentaje Iva en el DataCombo
 ' Carga_PorcentajeIva (MBFechaRegis)
'   Carga_ConceptosRetencion MBFechaRegis
End Sub

Private Sub TxtIvaBienMonIva_GotFocus()
    MarcarTexto TxtIvaBienMonIva
End Sub

Private Sub TxtIvaBienMonIva_LostFocus()
    ' MsgBox CTNumero(TxtIvaBienMonIva, 2)
    TextoValido TxtIvaBienMonIva, True, , 0
End Sub

Private Sub TxtIvaSerMonIva_GotFocus()
    MarcarTexto TxtIvaSerMonIva
End Sub

Private Sub TxtIvaSerMonIva_LostFocus()
Dim Total_IVA_S As Currency
    TextoValido TxtIvaBienMonIva, True, , 2
    TextoValido TxtIvaSerMonIva, True, , 2
    TextoValido TxtMontoIva, True, , 2
    
    'Verifica el Monto Iva Servicios
    Total_IVA_S = CDbl(TxtIvaBienMonIva) + CDbl(TxtIvaSerMonIva)
    If Total_IVA_S > CDbl(TxtMontoIva) Then
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

Private Sub TxtIvaSerValRet_LostFocus()
 CmdAir.SetFocus
End Sub

Private Sub TxtMonTitGrat_GotFocus()
   MarcarTexto TxtMonTitGrat
End Sub

Private Sub TxtMonTitGrat_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TxtMonTitGrat_LostFocus()
   TextoValido TxtMonTitGrat, True, , 2
End Sub

Private Sub TxtMonTitOner_GotFocus()
   MarcarTexto TxtMonTitOner
End Sub

Private Sub TxtMonTitOner_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtMonTitOner_LostFocus()
   TextoValido TxtMonTitOner, True, , 0
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
    'TxtNumAutComp = Format(Val(Round(TxtNumAutComp)), String(10, "0"))
     
   'Verifico si escogio dcto modificado
   If (cod = 4) Or (cod = 5) Then
      'Selecciona el numero mayor para continuar la secuencia en el
      'campo T_No y A_No
      sSQL = "SELECT TOP 1 * " _
           & "FROM Asiento_Compras " _
           & "WHERE Item = '" & NumEmpresa & "' " _
           & "ORDER BY A_No DESC "
      Select_Adodc AdoAsientoCompras, sSQL
      If AdoAsientoCompras.Recordset.RecordCount > 0 Then Ln_No = AdoAsientoCompras.Recordset.fields("A_No") + 1
         RatonReloj
         SetAdoAddNew "Asiento_Compras", True
         SetAdoFields "DocModificado", cod
         SetAdoFields "FechaEmiModificado", MBFechaEmiComp
         SetAdoFields "EstabModificado", TxtNumSerieUnoComp
         SetAdoFields "PtoEmiModificado", TxtNumSerieDosComp
         SetAdoFields "SecModificado", CNumSerieTresComp
         SetAdoFields "AutModificado", TxtNumAutComp
         SetAdoFields "MontoTituloOneroso", TxtMonTitOner
         SetAdoFields "MontoTituloGratuito", TxtMonTitGrat
         SetAdoFields "A_No", Maximo_De("Asiento_Compras", "A_No")
         SetAdoFields "T_No", Trans_No
         SetAdoUpdate
   End If
   CmdAir.SetFocus
End Sub

Private Sub TxtNumAutor_GotFocus()
   MarcarTexto TxtNumAutor
End Sub

Private Sub TxtNumAutor_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtNumAutor_LostFocus()
    If Val(TxtNumAutor) <= 0 Then TxtNumAutor = "0000000001"
    sSQL = "SELECT TOP 1 * " _
         & "FROM Trans_Compras " _
         & "WHERE IdProv = '" & CodigoCliente & "' " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND Establecimiento = '" & TxtNumSerieUno & "' " _
         & "AND PuntoEmision = '" & TxtNumSerieDos & "' " _
         & "AND Secuencial = " & CLng(TxtNumSerietres) & " " _
         & "AND Autorizacion = '" & TxtNumAutor & "' " _
         & "ORDER BY Fecha DESC, Secuencial DESC "
    Select_Adodc AdoAux, sSQL
    If AdoAux.Recordset.RecordCount > 0 Then MsgBox "USTED ESTA TRATANDO DE INGRESAR UNA FACTURA EXISTENTE"
    If cod = 3 Then Co.Autorizacion_LC = TxtNumAutor Else Co.Autorizacion_LC = Ninguno
End Sub

Private Sub TxtNumConParPol_GotFocus()
    MarcarTexto TxtNumConParPol
End Sub

Private Sub TxtNumConParPol_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtNumConParPol_LostFocus()
    TextoValido TxtNumConParPol, True, , 0
    TxtNumConParPol = Format(Val(TxtNumConParPol), String(10, "0"))
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
   TxtNumUnoComRet = Format(Val(TxtNumUnoComRet), "000")
   TxtNumDosComRet = Format(Val(TxtNumDosComRet), "000")
   Co.Serie_R = TxtNumUnoComRet & TxtNumDosComRet
   TxtNumTresComRet = ReadSetDataNum("RE_SERIE_" & Co.Serie_R, True, False)
   Co.Retencion = 0
   Co.RetSecuencial = True
   sSQL = "SELECT TOP 1 AutRetencion " _
        & "FROM Trans_Air " _
        & "WHERE Tipo_Trans = 'C' " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND Fecha <= #" & BuscarFecha(MBFechaRegis) & "# " _
        & "AND EstabRetencion = '" & TxtNumUnoComRet & "' " _
        & "AND PtoEmiRetencion = '" & TxtNumDosComRet & "' " _
        & "ORDER BY SecRetencion DESC, Fecha DESC, AutRetencion DESC "
   Select_Adodc AdoAux, sSQL
   If AdoAux.Recordset.RecordCount > 0 Then
      If Len(AdoAux.Recordset.fields("AutRetencion")) >= 13 Then
         TxtNumUnoAutComRet = RUC
      Else
         TxtNumUnoAutComRet = AdoAux.Recordset.fields("AutRetencion")
      End If
   Else
      TxtNumUnoAutComRet = "1"
      TxtNumTresComRet = "1"
   End If
   Co.Autorizacion_R = TxtNumUnoAutComRet
End Sub

Private Sub TxtNumSerieDos_GotFocus()
    MarcarTexto TxtNumSerieDos
End Sub

Private Sub TxtNumSerieDos_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtNumSerieDos_LostFocus()
    If cod = 3 Then
       Co.Serie_LC = TxtNumSerieUno & TxtNumSerieDos
       TxtNumSerietres = ReadSetDataNum("LC_SERIE_" & Co.Serie_LC, True, False)
    Else
       Co.Serie_LC = Ninguno
    End If
    TextoValido TxtNumSerieDos, True, , 0
    If Val(TxtNumSerieDos) <= 0 Then TxtNumSerieDos = "001"
    TxtNumSerieDos = Format(Val(TxtNumSerieDos), "000")
    If cod = 4 Then
       sSQL = "SELECT TOP 1 * " _
            & "FROM Trans_Compras " _
            & "WHERE TipoComprobante = " & cod & " " _
            & "AND Establecimiento = '" & TxtNumSerieUno & "' " _
            & "AND PuntoEmision = '" & TxtNumSerieDos & "' " _
            & "AND Item = '" & NumEmpresa & "' " _
            & "ORDER BY Fecha DESC,Secuencial DESC "
       Select_Adodc AdoAux, sSQL
       If AdoAux.Recordset.RecordCount > 0 Then
          TxtNumSerietres = AdoAux.Recordset.fields("Secuencial") + 1
          MBFechaCad = AdoAux.Recordset.fields("FechaCaducidad")
          TxtNumSerieUno = AdoAux.Recordset.fields("Establecimiento")
          TxtNumSerieDos = AdoAux.Recordset.fields("PuntoEmision")
          TxtNumAutor = AdoAux.Recordset.fields("Autorizacion")
       Else
          TxtNumAutor = Autorizacion
          TxtNumSerietres = "000000001"
       End If
    End If
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
    TxtNumSerietres = Format(Val(Round(TxtNumSerietres)), "000000000")
    MarcarTexto TxtNumSerietres
End Sub

Private Sub TxtNumSerietres_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtNumSerietres_LostFocus()
    If Val(TxtNumSerietres) <= 0 Then TxtNumSerietres = "000000001"
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
   Co.Retencion = 0
   If Val(TxtNumTresComRet) <= 0 Then TxtNumTresComRet = "000000001"
   TxtNumTresComRet = Format(Round(Val(TxtNumTresComRet)), "000000000")
  'Calcula la sumatoria de Monto Iva Bienes, Monto Iva Servicios y Base Imponible
   TxtSumatoria = Format(CTNumero(TxtBaseImpoNoObjIVA, 2) + CTNumero(TxtBaseImpo, 2) + CTNumero(TxtBaseImpoGrav, 2), "#,##0.00")
  'TxtSumatoria = TxtBaseImpoGrav
   Co.Serie_R = TxtNumUnoComRet & TxtNumDosComRet
   Co.Retencion = Val(TxtNumTresComRet)
   Co.RetNueva = True
   sSQL = "SELECT * " _
        & "FROM Trans_Air " _
        & "WHERE Tipo_Trans = 'C' " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND EstabRetencion = '" & TxtNumUnoComRet & "' " _
        & "AND PtoEmiRetencion = '" & TxtNumDosComRet & "' " _
        & "AND SecRetencion = " & Co.Retencion & " "
   'If Len(Co.Autorizacion_R) < 13 Then sSQL = sSQL & "AND Porcentaje > 0 "
   Select_Adodc AdoAux, sSQL
   If AdoAux.Recordset.RecordCount > 0 Then
      Titulo = "RETENCION REPETIDA"
      Mensajes = "Número de Retención ya existe," & vbCrLf _
               & "si continua se borrará los datos" & vbCrLf _
               & "de este número de retención." & vbCrLf & vbCrLf _
               & "QUIERE REPROCESARLA"
      If BoxMensaje = vbYes Then
         Co.Retencion = Val(TxtNumTresComRet)
         Co.RetNueva = False
      End If
   End If
End Sub

Private Sub TxtNumUnoAutComRet_GotFocus()
    MarcarTexto TxtNumUnoAutComRet
End Sub

Private Sub TxtNumUnoAutComRet_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtNumUnoAutComRet_LostFocus()
    If Val(TxtNumUnoAutComRet) <= 0 Then TxtNumUnoAutComRet = "0000000000"
   'TxtNumUnoAutComRet = Format(Val(Round(TxtNumUnoAutComRet)), String(13, "0"))
    Co.Autorizacion_R = TxtNumUnoAutComRet
    If Val(TxtNumTresComRet) <> ReadSetDataNum("RE_SERIE_" & Co.Serie_R, True, False) Then
       Titulo = "SECUENCIAL DE RETENCION"
       Mensajes = "Número de Retención: " & Co.Serie_R & "-" & Format(Co.Retencion, "000000000") & vbCrLf _
                & "no esta en orden secuencial." & vbCrLf & vbCrLf _
                & "QUIERE PROCESARLA?"
       If BoxMensaje = vbYes Then Co.RetSecuencial = False
    End If
End Sub

Private Sub TxtNumUnoComRet_GotFocus()
   MarcarTexto TxtNumUnoComRet
End Sub

Private Sub TxtNumUnoComRet_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TxtNumUnoComRet_LostFocus()
   TextoValido TxtNumUnoComRet
   If Val(TxtNumUnoComRet) <= 0 Then TxtNumUnoComRet = "001"
End Sub

Public Sub Carga_CreditoTributario()
  'Carga la Tabla de Catalogos Tributarios al DataCombo
   sSQL = "SELECT (Credito_Tributario & ' - ' & Descripcion) As Sustento,* " _
        & "FROM Tipo_Tributario " _
        & "WHERE Credito_Tributario <> '.' " _
        & "AND Fecha_Inicio <= #" & BuscarFecha(FechaComp) & "# " _
        & "AND Fecha_Final >= #" & BuscarFecha(FechaComp) & "# " _
        & "ORDER BY Credito_Tributario "
   SelectDB_Combo DCSustento, AdoSustento, sSQL, "Sustento"
End Sub

Public Sub Carga_TipoComprobante(CargaTC As String)
     sSQL = "SELECT Tipo_Comprobante_Codigo, Descripcion " _
          & "FROM Tipo_Comprobante " _
          & "WHERE TC = 'TDC' " _
          & "ORDER BY Descripcion "
     SelectDB_Combo DCDctoModif, AdoTipoComprobante, sSQL, "Descripcion"
    
    'Capturo el codigo del Tipo de Catalogo Tributario
     Cap = CargaTC
            
    'Busco el codigo en la tabla Tipo Comprobante///descripcion
    Cadena = Ninguno
    With AdoSustento.Recordset
     If .RecordCount > 0 Then
        .MoveFirst
        .Find ("Credito_Tributario = '" & CargaTC & "' ")
         If Not .EOF Then Cadena = Replace(.fields("Codigo_Tipo_Comprobante"), " ", ",")
     End If
    End With
    sSQL = "SELECT * " _
         & "FROM Tipo_Comprobante " _
         & "WHERE Tipo_Comprobante_Codigo IN (" & Cadena & ") " _
         & "AND TC = 'TDC' "
    If TipoBenef = "R" Then
       sSQL = sSQL & "AND R <> " & Val(adFalse) & " "
    Else
       sSQL = sSQL & "AND C <> " & Val(adFalse) & " "
    End If
    sSQL = sSQL & "ORDER BY Tipo_Comprobante_Codigo "
    SelectDB_Combo DCTipoComprobante, AdoTipoComprobante, sSQL, "Descripcion"
End Sub

Public Sub Captura_TipoComprobante()
   'Captura lo que tiene el Combo de Tipo de Comprobante
    Label15.Caption = "Fechas de " & DCTipoComprobante
    Captc = SinEspaciosIzq(DCTipoComprobante.Text)
    Cap1 = TrimStrg(DCTipoComprobante.Text)
     
   'Busca que sea igual a la Descripcion
    With AdoTipoComprobante.Recordset
        If .RecordCount > 0 Then
           .MoveFirst
           .Find ("Descripcion = '" & Cap1 & "' ")
           If Not .EOF Then
              cod = .fields("Tipo_Comprobante_Codigo")
           Else
              MsgBox "Vuelva a seleccionar", vbInformation, "Aviso"
           End If
        End If
    End With
   'MsgBox Cod
    If (cod = 4) Or (cod = 5) Then
       FraDctoModificado.Visible = True
       Documento_Modificado
       'Carga en el combo de Documentos Modificados los
       'Tipos de Comprobantes
        sSQL = "SELECT Tipo_Comprobante_Codigo, Descripcion " _
             & "FROM Tipo_Comprobante " _
             & "WHERE TC = 'TDC' " _
             & "ORDER BY Descripcion "
        SelectDB_Combo DCDctoModif, AdoTipoComprobante, sSQL, "Descripcion"
    Else
        FraDctoModificado.Visible = False
    End If
End Sub

Public Sub Captura_TipoComprobante_DctoModificado()
    CapDcto = TrimStrg(DCDctoModif.Text)
     
    'Busca que sea igual a la Descripcion
    With AdoTipoComprobante.Recordset
        If .RecordCount > 0 Then
           .MoveFirst
           .Find ("Descripcion = '" & CapDcto & "' ")
           If Not .EOF Then
              CapDm = .fields("Tipo_Comprobante_Codigo")
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
     Select_Adodc AdoAux, sSQL
     With AdoAux.Recordset
      If .RecordCount > 0 Then
          Do While Not .EOF
             CNumSerieTresComp.AddItem .fields("Secuencial")
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
  SelectDB_Combo DCPorcenRetenIvaBien, AdoRetIvaBienes, sSQL, "Porc"
  
  sSQL = "SELECT * " _
       & "FROM Tabla_Por_IVA " _
       & "WHERE Servicios <> " & Val(adFalse) & " " _
       & "ORDER BY Porc "
  SelectDB_Combo DCPorcenRetenIvaServ, AdoRetIvaServicios, sSQL, "Porc"
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
End Sub

Public Sub Limpiar_Controles()
    ac = 0
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
    TxtNumSerieUno.Text = "001"
    TxtNumSerieDos.Text = "001"
    TxtNumSerietres.Text = "0"
    TxtNumAutor.Text = ""
    FechaValida MBFechaEmi
    FechaValida MBFechaRegis
    FechaValida MBFechaCad
    TxtBaseImpo.Text = "0.00"
    TxtBaseImpoGrav.Text = "0.00"
    TxtBaseImpoIce.Text = "0.00"
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
    TxtNumConParPol.Text = ""
    TxtMonTitOner.Text = ""
    TxtMonTitGrat.Text = ""
    'Limpia la grilla
    'Borra Asiento Air
    sSQL = "DELETE * " _
         & "FROM Asiento_Air " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND T_No = " & Trans_No & " " _
         & "AND Tipo_Trans = 'C' "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "SELECT * " _
         & "FROM Asiento_Air " _
         & "WHERE codRet <> '.' " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND T_No = " & Trans_No & " " _
         & "AND Tipo_Trans = 'C' " _
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
          SumaReten = SumaReten + .fields("ValRet")
         .MoveNext
       Loop
   End If
  End With
  TxtTotalReten = Format(SumaReten, "#,##0.00")
End Sub

Public Sub Carga_Datos_Iniciales(MBFecha As MaskEdBox, EsNuevo As Boolean)
   'Encero todo
    CFormaPago.Clear
    CFormaPago.AddItem "Local"
    CFormaPago.AddItem "Exterior"
    CFormaPago.Text = "Local"
    
    ac = 0
    DCPorcenIce = 0
    DCPorcenRetenIvaBien = 0
    DCPorcenRetenIvaServ = 0
    
    CodPorIva = 0
    CodPorIce = "0"
    CodRetBien = 0
    CodRetServ = 0
    
    Limpiar_Controles
    Listar_Air
   'Cargo el No.Autorización de las retenciones
    TxtNumUnoAutComRet = AutorizaRet
    
   'Carga el Sustento Tributario
    Carga_CreditoTributario
   'Carga en el Data Combo los Clientes con su RUC
    DCTipoComprobante.Text = "Factura"
   'Carga la Tabla de Retencion Iva Bienes y Servicios al DataCombo
    Carga_RetencionIvaBienes_Servicios
    DCPorcenIce.Text = ""
   'Carga la Tabla de Conceptos Retencion al DataCombo
    MBFechaRegis = MBFechaEmi
   'Verifico si existen registros caso contrario despliego mensaje
   'Carga los Conceptos de retención en la Fuente al DataCombo
    Carga_ConceptosRetencion MBFechaEmi
    sSQL = "SELECT (Codigo & ' - ' & Cuenta) As Cuentas  " _
         & "FROM Catalogo_Cuentas " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TC = 'RF' " _
         & "AND DG = 'D' " _
         & "ORDER BY Codigo "
    SelectDB_Combo DCRetFuente, AdoRetFuente, sSQL, "Cuentas"
    If AdoRetFuente.Recordset.RecordCount > 0 Then Rf = 1 Else Rf = 0
   'Carga los Conceptos de retención IVA Servicios al DataCombo
    sSQL = "SELECT (Codigo & ' - ' & Cuenta) As Cuentas  " _
         & "FROM Catalogo_Cuentas " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TC = 'RI' " _
         & "AND DG = 'D' " _
         & "ORDER BY Codigo "
    SelectDB_Combo DCRetISer, AdoRetIvaSerCC, sSQL, "Cuentas"
    If AdoRetIvaSerCC.Recordset.RecordCount > 0 Then rs = 1 Else rs = 0
    'Carga los Conceptos de retención IVA Bienes al DataCombo
    sSQL = "SELECT (Codigo & ' - ' & Cuenta) As Cuentas  " _
         & "FROM Catalogo_Cuentas " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TC = 'RB' " _
         & "AND DG = 'D' " _
         & "ORDER BY Codigo "
    SelectDB_Combo DCRetIBienes, AdoRetIvaBienesCC, sSQL, "Cuentas"
    If AdoRetIvaBienesCC.Recordset.RecordCount > 0 Then Rb = 1 Else Rb = 0
   'Si es Nuevo ingresa por aqui
    ChRetF.Visible = True
    ChRetF.value = 1
    DCRetFuente.Visible = True
    FrmRetencion.Visible = True
'    LblMensaje.Visible = False
    If EsNuevo Then
       'Si todas las variables tienen cero despliego mensaje y no cargo nada
       'No hay cuentas
       If (Rf And rs And Rb) = 0 Then
           ChRetF.Visible = False
           ChRetF.value = 0
           DCRetFuente.Visible = False
           FrmRetencion.Visible = False
           'LblMensaje.Visible = True
           Activar_BS
       Else
           ChRetB.SetFocus
       End If
    End If
End Sub

Public Sub Grabacion()
Dim PagoLocExt As String
Dim PaisEfecPago As String
Dim AplicConvDobTrib As String
Dim PagExtSujRetNorLeg As String
Dim FormaPago As String

   'Grabo en el Asiento_Compras e implicito Asiento_Air
    If OpcSi.value = True Then ValorP = "S" Else ValorP = "N"
    sSQL = "DELETE * " _
         & "FROM Asiento_Compras " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND T_No = " & Trans_No & " "
    Ejecutar_SQL_SP sSQL
   'MsgBox Cod & vbCrLf & Cap
    SetAdoAddNew "Asiento_Compras"
    SetAdoFields "IdProv", CodigoCliente
    SetAdoFields "DevIva", ValorP
    SetAdoFields "CodSustento", Cap
    SetAdoFields "TipoComprobante", cod
    SetAdoFields "Establecimiento", TxtNumSerieUno
    SetAdoFields "PuntoEmision", TxtNumSerieDos
    SetAdoFields "Secuencial", CTNumero(TxtNumSerietres)
    SetAdoFields "Autorizacion", TxtNumAutor
    SetAdoFields "FechaEmision", MBFechaEmi
    SetAdoFields "FechaRegistro", MBFechaRegis
    SetAdoFields "FechaCaducidad", MBFechaCad
    SetAdoFields "BaseNoObjIVA", CTNumero(TxtBaseImpoNoObjIVA, 2)
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
    If TxtNumConParPol = "" Or TxtNumConParPol = "0000000000" Then
       SetAdoFields "ContratoPartidoPolitico", "0000000000"
    Else
       SetAdoFields "ContratoPartidoPolitico", TxtNumConParPol
    End If
    SetAdoFields "MontoTituloOneroso", CTNumero(TxtMonTitOner, 2)
    SetAdoFields "MontoTituloGratuito", CTNumero(TxtMonTitGrat, 2)
   'Verifico si activaron los checks de retenciones
    If ChRetB = 1 Then SetAdoFields "Cta_Bienes", SinEspaciosIzq(DCRetIBienes)
    If ChRetS = 1 Then SetAdoFields "Cta_Servicio", SinEspaciosIzq(DCRetISer)
   'Forma de Pago
    FormaPago = SinEspaciosIzq(DCTipoPago.Text)
    PagoLocExt = "01"
    PaisEfecPago = "NA"
    AplicConvDobTrib = "NA"
    PagExtSujRetNorLeg = "NA"
    If CFormaPago.Text = "Exterior" Then
       PagoLocExt = "02"
       If AdoPais.Recordset.RecordCount > 0 Then
          AdoPais.Recordset.MoveFirst
          AdoPais.Recordset.Find ("Descripcion_Rubro = '" & DCPais.Text & "' ")
          If Not AdoPais.Recordset.EOF Then PaisEfecPago = AdoPais.Recordset.fields("CPais")
       End If
       If OpcSiAplicaDoble.value Then AplicConvDobTrib = "SI" Else AplicConvDobTrib = "NO"
       If OpcSiFormaLegal.value Then PagExtSujRetNorLeg = "SI" Else PagExtSujRetNorLeg = "NO"
    End If
    SetAdoFields "PagoLocExt", PagoLocExt
    SetAdoFields "PaisEfecPago", PaisEfecPago
    SetAdoFields "AplicConvDobTrib", AplicConvDobTrib
    SetAdoFields "PagExtSujRetNorLeg", PagExtSujRetNorLeg
    SetAdoFields "FormaPago", FormaPago
    SetAdoFields "A_No", 1
    SetAdoFields "T_No", Trans_No
    SetAdoFields "CodigoU", CodigoUsuario
    SetAdoUpdate
    'MsgBox "* ======> " & Trans_No
End Sub

Public Sub Habilita_Controles()
   'Habilito los controles para la modificacion
    SSTCompras.Enabled = True
    CmdGrabar.Enabled = True
    FrmRetencion.Enabled = True
    'Label23.Visible = True
End Sub

Public Sub Deshabilita_Controles()
   'Deshabilito los controles para la modificacion
    SSTCompras.Enabled = False
    CmdGrabar.Enabled = False
    FrmRetencion.Enabled = False
    'Label23.Visible = False
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
   Ejecutar_SQL_SP sSQL
  'Borra Asiento Air
   sSQL = "DELETE * " _
        & "FROM Asiento_Air " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND T_No = " & Trans_No & " " _
        & "AND Tipo_Trans = 'C' " _
        & "AND CodigoU = '" & CodigoUsuario & "' "
   Ejecutar_SQL_SP sSQL
  'Presentamos la malla Asiento Air
  'CodRet,Detalle,BaseImp,Porcentaje,ValRet,EstabRetencion,PtoEmiRetencion,SecRetencion,AutRetencion,FechaEmiRet,Item,CodigoU
   sSQL = "SELECT * " _
        & "FROM Asiento_Air " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND CodigoU =  '" & CodigoUsuario & "' " _
        & "AND T_No = " & Trans_No & " " _
        & "AND Tipo_Trans = 'C' " _
        & "ORDER BY CodRet "
   Select_Adodc_Grid DGConceptoAir, AdoAsientoAir, sSQL
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

